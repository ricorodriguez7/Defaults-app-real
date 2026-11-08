import { NextResponse } from 'next/server';
import { getAuthorizedEmail } from '@/lib/authorization';
import { isPlacementTeam } from '@/lib/members';
import { databaseFailureMessage, getDb, withMongoTransaction } from '@/lib/mongodb';

export const runtime = 'nodejs';
export const dynamic = 'force-dynamic';
export const maxDuration = 30;

type RouteContext = { params: { memberId: string } };

function normalizeText(value: string) {
  return value.trim().replace(/\s+/g, ' ');
}

function normalizeTeam(value: string) {
  return value.trim().toLowerCase().replace(/[\s-]+/g, '_');
}

function validEmail(value: string) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(value);
}

function isDuplicateKey(error: unknown) {
  return error && typeof error === 'object' && 'code' in error && error.code === 11000;
}

export async function PATCH(request: Request, { params }: RouteContext) {
  try {
    const enteredByEmail = await getAuthorizedEmail();
    if (!enteredByEmail) return NextResponse.json({ message: 'Sign in with an active member account.' }, { status: 401 });
    const payload = await request.json();
    const action = String(payload?.action ?? 'edit');
    const reason = String(payload?.reason ?? '').trim();
    if (!['edit', 'deactivate', 'reactivate'].includes(action)) {
      return NextResponse.json({ message: 'Choose edit, deactivate, or reactivate.' }, { status: 400 });
    }
    if (!reason) return NextResponse.json({ message: 'A reason is required.' }, { status: 400 });

    const now = new Date();
    await getDb();
    const result = await withMongoTransaction(async (db, session) => {
      const members = db.collection('members');
      const before = await members.findOne({ memberId: params.memberId }, { session });
      if (!before) throw new Error('MEMBER_NOT_FOUND');

      if (action === 'edit') {
        const name = normalizeText(String(payload?.name ?? ''));
        const team = normalizeTeam(String(payload?.team ?? ''));
        const email = String(payload?.email ?? '').trim().toLowerCase();
        if (!name) throw new Error('Name is required.');
        if (!isPlacementTeam(team)) throw new Error('Choose placement_cell or summer_internship.');
        if (email && !validEmail(email)) throw new Error('Enter a valid email address or leave it blank.');

        const update = {
          name,
          nameNormalized: name.toLowerCase(),
          team,
          teamNormalized: team,
          updatedAt: now,
          updatedByEmail: enteredByEmail,
        };
        if (email) {
          await members.updateOne(
            { memberId: params.memberId },
            { $set: { ...update, email, emailNormalized: email } },
            { session },
          );
        } else {
          await members.updateOne(
            { memberId: params.memberId },
            { $set: update, $unset: { email: '', emailNormalized: '' } },
            { session },
          );
        }
        const after = await members.findOne({ memberId: params.memberId }, { session });
        await db.collection('audit_log').insertOne({
          action: 'Member edited',
          collection: 'members',
          by: enteredByEmail,
          at: now,
          before: { name: before.name, team: before.team, email: before.email || null },
          after: { name: after?.name, team: after?.team, email: after?.email || null, reason },
        }, { session });
        return { member: after };
      }

      if (before.emailNormalized === enteredByEmail && action === 'deactivate') {
        throw new Error('You cannot deactivate your own account.');
      }
      const nextActive = action === 'reactivate';
      if (Boolean(before.active !== false) === nextActive) {
        throw new Error(nextActive ? 'Member is already active.' : 'Member is already inactive.');
      }
      await members.updateOne(
        { memberId: params.memberId },
        { $set: { active: nextActive, updatedAt: now, updatedByEmail: enteredByEmail } },
        { session },
      );
      await db.collection('audit_log').insertOne({
        action: nextActive ? 'Member reactivated' : 'Member deactivated',
        collection: 'members',
        by: enteredByEmail,
        at: now,
        before: { memberId: before.memberId, name: before.name, team: before.team, active: before.active !== false },
        after: { memberId: before.memberId, name: before.name, team: before.team, active: nextActive, reason },
      }, { session });
      return { active: nextActive };
    });
    return NextResponse.json(result);
  } catch (error) {
    if (error instanceof Error && error.message === 'MEMBER_NOT_FOUND') {
      return NextResponse.json({ message: 'Member not found.' }, { status: 404 });
    }
    if (isDuplicateKey(error)) {
      return NextResponse.json({ message: 'Another member already uses that email or name/team.' }, { status: 409 });
    }
    if (error instanceof Error && (
      error.message.startsWith('A reason') ||
      error.message.startsWith('You cannot') ||
      error.message.startsWith('Member is already') ||
      error.message.startsWith('Name is required') ||
      error.message.startsWith('Choose placement') ||
      error.message.startsWith('Enter a valid')
    )) {
      return NextResponse.json({ message: error.message }, { status: 400 });
    }
    console.error('Unable to update member.', error);
    return NextResponse.json({ message: databaseFailureMessage(error, 'Unable to update the member.') }, { status: 500 });
  }
}
