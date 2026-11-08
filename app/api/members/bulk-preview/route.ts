import { NextResponse } from 'next/server';
import { getAuthorizedEmail } from '@/lib/authorization';
import { getApiErrorMessage, requestExceedsBodyLimit } from '@/lib/database-errors';
import { MemberInput, planMemberImport, StoredMember } from '@/lib/member-import';
import { getDb } from '@/lib/mongodb';

export const runtime = 'nodejs';
export const dynamic = 'force-dynamic';
export const maxDuration = 30;

export async function POST(request: Request) {
  try {
    const enteredByEmail = await getAuthorizedEmail();
    if (!enteredByEmail) return NextResponse.json({ message: 'Sign in with an active member account.' }, { status: 401 });
    if (requestExceedsBodyLimit(request)) {
      return NextResponse.json({ message: 'The member upload exceeds 4 MB. Split the file and try again.' }, { status: 413 });
    }
    const payload = await request.json();
    const rows: MemberInput[] = Array.isArray(payload?.members) ? payload.members : [];
    if (!rows.length || rows.length > 2000) {
      return NextResponse.json({ message: 'Preview between 1 and 2,000 member rows.' }, { status: 400 });
    }

    const db = await getDb();
    const existing = await db.collection('members').find({}).project({
      _id: 0,
      memberId: 1,
      name: 1,
      nameNormalized: 1,
      team: 1,
      teamNormalized: 1,
      email: 1,
      emailNormalized: 1,
      active: 1,
    }).toArray() as StoredMember[];
    const plan = planMemberImport(rows, existing);
    return NextResponse.json({
      counts: {
        add: plan.filter((item) => item.action === 'add').length,
        update: plan.filter((item) => item.action === 'update').length,
        skip: plan.filter((item) => item.action === 'skip').length,
      },
      rows: plan.map(({ row, action, reason, member }) => ({
        row,
        action,
        reason,
        name: member?.name ?? String(rows[row - 2]?.name ?? ''),
        team: member?.team ?? String(rows[row - 2]?.team ?? ''),
        email: member?.email ?? String(rows[row - 2]?.email ?? ''),
      })),
    });
  } catch (error) {
    console.error('Unable to preview member upload.', error);
    return NextResponse.json({ message: getApiErrorMessage(error, 'Unable to preview the member upload.') }, { status: 500 });
  }
}
