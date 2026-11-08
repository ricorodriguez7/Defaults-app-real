import { randomUUID } from 'node:crypto';
import { NextResponse } from 'next/server';
import { getAuthorizedEmail } from '@/lib/authorization';
import { MemberInput, planMemberImport, StoredMember } from '@/lib/member-import';
import { databaseFailureMessage, getDb, withMongoTransaction } from '@/lib/mongodb';
import { readJsonBody, RequestBodyError } from '@/lib/request-body';

export const runtime = 'nodejs';
export const dynamic = 'force-dynamic';
export const maxDuration = 30;

export async function POST(request: Request) {
  try {
    const enteredByEmail = await getAuthorizedEmail();
    if (!enteredByEmail) return NextResponse.json({ message: 'Sign in with an active member account.' }, { status: 401 });
    const payload = await readJsonBody<{ members?: MemberInput[] }>(request);
    const input = Array.isArray(payload?.members) ? payload.members : [];
    if (!input.length || input.length > 2000) {
      return NextResponse.json({ message: 'Import between 1 and 2,000 member rows.' }, { status: 400 });
    }

    await getDb();
    const result = await withMongoTransaction(async (db, session) => {
      const existing = await db.collection<StoredMember>('members').find({}, { session }).toArray();
      const plan = planMemberImport(input, existing);
      const now = new Date();
      let added = 0;
      let updated = 0;
      const auditEntries = [];

      for (const item of plan) {
        if (!item.member || item.action === 'skip') continue;
        if (item.action === 'add') {
          const member = {
            memberId: randomUUID(),
            ...item.member,
            active: true,
            createdAt: now,
            createdByEmail: enteredByEmail,
            updatedAt: now,
            updatedByEmail: enteredByEmail,
          };
          await db.collection('members').insertOne(member, { session });
          added += 1;
          auditEntries.push({
            action: 'Member added by bulk upload',
            collection: 'members',
            by: enteredByEmail,
            at: now,
            after: { memberId: member.memberId, name: member.name, team: member.team, email: member.email || null },
          });
        } else {
          await db.collection('members').updateOne(
            { memberId: item.memberId },
            { $set: { ...item.member, updatedAt: now, updatedByEmail: enteredByEmail } },
            { session },
          );
          updated += 1;
          auditEntries.push({
            action: 'Member updated by bulk upload',
            collection: 'members',
            by: enteredByEmail,
            at: now,
            before: { memberId: item.memberId },
            after: { memberId: item.memberId, ...item.member, reason: 'Bulk upload' },
          });
        }
      }

      auditEntries.push({
        action: 'Members bulk upload',
        collection: 'members',
        by: enteredByEmail,
        at: now,
        after: { rows: input.length, added, updated, skipped: plan.filter((item) => item.action === 'skip').length },
      });
      if (auditEntries.length) await db.collection('audit_log').insertMany(auditEntries, { session });
      return {
        counts: { add: added, update: updated, skip: plan.filter((item) => item.action === 'skip').length },
      };
    });
    return NextResponse.json(result);
  } catch (error) {
    if (error instanceof RequestBodyError) {
      return NextResponse.json({ message: error.message }, { status: error.status });
    }
    if (error && typeof error === 'object' && 'code' in error && error.code === 11000) {
      return NextResponse.json({ message: 'Import conflicts with an existing email or name/team member. Preview the file again.' }, { status: 409 });
    }
    console.error('Unable to import members.', error);
    return NextResponse.json({ message: databaseFailureMessage(error, 'Unable to import members.') }, { status: 500 });
  }
}
