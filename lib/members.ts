import { Db } from 'mongodb';

export function isPlacementTeam(team: unknown): team is 'placement_cell' | 'summer_internship' {
  return team === 'placement_cell' || team === 'summer_internship';
}

export async function findActiveMember(db: Db, memberId: string) {
  const member = await db.collection('members').findOne({ memberId, active: { $ne: false } });
  if (!member || !isPlacementTeam(member.team)) return null;
  return member;
}

export async function findActorMember(db: Db, email: string) {
  const member = await db.collection('members').findOne({ emailNormalized: email, active: { $ne: false } });
  return member;
}
