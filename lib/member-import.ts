import { isPlacementTeam } from '@/lib/members';

export type MemberInput = { name?: unknown; team?: unknown; email?: unknown };
export type StoredMember = {
  memberId: string;
  name: string;
  nameNormalized?: string;
  team: string;
  teamNormalized?: string;
  email?: string;
  emailNormalized?: string;
  active?: boolean;
};
export type PlannedMember = {
  row: number;
  action: 'add' | 'update' | 'skip';
  reason: string;
  member?: { name: string; nameNormalized: string; team: string; teamNormalized: string; email?: string; emailNormalized?: string };
  memberId?: string;
};

const normalizeName = (value: string) => value.trim().replace(/\s+/g, ' ').toLowerCase();
const normalizeTeam = (value: string) => value.trim().toLowerCase().replace(/[\s-]+/g, '_');
const validEmail = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;

export function planMemberImport(rows: MemberInput[], existing: StoredMember[]): PlannedMember[] {
  const plans: PlannedMember[] = [];
  const seenEmails = new Set<string>();
  const seenNames = new Set<string>();
  const emailIndex = new Map<string, StoredMember>();
  const nameIndex = new Map<string, StoredMember>();

  for (const member of existing) {
    if (member.emailNormalized) emailIndex.set(member.emailNormalized.toLowerCase(), member);
    nameIndex.set(`${normalizeName(member.name)}\u0000${normalizeTeam(member.team)}`, member);
  }

  rows.forEach((row, index) => {
    const rowNumber = index + 2;
    if (!row || typeof row !== 'object') {
      plans.push({ row: rowNumber, action: 'skip', reason: 'Row is not a valid record.' });
      return;
    }

    const name = String(row.name ?? '').trim().replace(/\s+/g, ' ');
    const teamInput = String(row.team ?? '').trim();
    const team = normalizeTeam(teamInput);
    const email = String(row.email ?? '').trim().toLowerCase();
    const nameNormalized = normalizeName(name);
    const nameKey = `${nameNormalized}\u0000${team}`;
    if (!name) {
      plans.push({ row: rowNumber, action: 'skip', reason: 'Name is required.' });
      return;
    }
    if (!isPlacementTeam(team)) {
      plans.push({ row: rowNumber, action: 'skip', reason: 'Team must be placement_cell or summer_internship.' });
      return;
    }
    if (email && !validEmail.test(email)) {
      plans.push({ row: rowNumber, action: 'skip', reason: 'Email format is invalid.' });
      return;
    }
    if ((email && seenEmails.has(email)) || seenNames.has(nameKey)) {
      plans.push({ row: rowNumber, action: 'skip', reason: 'Duplicate email or name/team in this file.' });
      return;
    }
    if (email) seenEmails.add(email);
    seenNames.add(nameKey);

    const match = (email ? emailIndex.get(email) : undefined) ?? nameIndex.get(nameKey);
    const conflictingName = nameIndex.get(nameKey);
    if (match && conflictingName && conflictingName.memberId !== match.memberId) {
      plans.push({ row: rowNumber, action: 'skip', reason: 'That name/team already belongs to a different member.' });
      return;
    }
    const duplicateEmail = email ? emailIndex.get(email) : undefined;
    if (!match && duplicateEmail) {
      plans.push({ row: rowNumber, action: 'skip', reason: 'Email already belongs to another member.' });
      return;
    }

    const planned = {
      name,
      nameNormalized,
      team,
      teamNormalized: team,
      ...(email || match?.email ? { email: email || match?.email, emailNormalized: email || match?.emailNormalized } : {}),
    };
    if (!match) {
      plans.push({ row: rowNumber, action: 'add', reason: 'New member.', member: planned });
      return;
    }

    const unchanged = match.name === planned.name &&
      normalizeTeam(match.team) === planned.team &&
      (match.emailNormalized || '') === (planned.emailNormalized || '');
    if (unchanged) {
      plans.push({ row: rowNumber, action: 'skip', reason: 'Already up to date.', memberId: match.memberId });
      return;
    }

    plans.push({ row: rowNumber, action: 'update', reason: 'Matched by email or name/team.', member: planned, memberId: match.memberId });
  });

  return plans;
}
