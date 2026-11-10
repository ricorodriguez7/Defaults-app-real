import { randomUUID } from 'node:crypto';
import { NextResponse } from 'next/server';
import { getAuthorizedEmail } from '@/lib/authorization';
import { getApiErrorMessage } from '@/lib/database-errors';
import { findActorMember } from '@/lib/members';
import { withMongoTransaction } from '@/lib/mongodb';

export const runtime = 'nodejs';
export const dynamic = 'force-dynamic';
export const maxDuration = 30;

export async function POST(request: Request) {
  try {
    const email = await getAuthorizedEmail();
    if (!email) return NextResponse.json({ message: 'Sign in with an active member account.' }, { status: 401 });
    const payload = await request.json();
    const batchInput = String(payload?.batch ?? '').trim().toLowerCase();
    const batch = batchInput === 'junior' || batchInput === 'junior batch'
      ? 'junior'
      : batchInput === 'senior' || batchInput === 'senior batch'
        ? 'senior'
        : batchInput === 'unset' ? '' : null;
    const rollNos = Array.isArray(payload?.rollNos)
      ? Array.from(new Set(payload.rollNos.map((value: unknown) => String(value).trim().toUpperCase()).filter(Boolean)))
      : [];
    if (batch === null) return NextResponse.json({ message: 'Choose Junior batch, Senior batch, or Not set.' }, { status: 400 });
    if (!rollNos.length || rollNos.length > 10000) {
      return NextResponse.json({ message: 'Select between 1 and 10,000 students.' }, { status: 400 });
    }

    const updated = await withMongoTransaction(async (db, session) => {
      const actor = await findActorMember(db, email);
      if (!actor) throw new Error('Your active member record could not be found.');
      const students = await db.collection('students').find(
        { rollNoNormalized: { $in: rollNos } },
        { session, projection: { _id: 1, rollNo: 1, batch: 1 } },
      ).toArray();
      const now = new Date();
      await db.collection('students').bulkWrite(students.map((student) => ({
        updateOne: {
          filter: { _id: student._id },
          update: { $set: { batch, batchUpdatedAt: now, batchUpdatedByMemberId: actor.memberId } },
        },
      })), { session });
      const changedStudents = students.filter((student) => (student.batch ?? '') !== batch);
      if (changedStudents.length) {
        await db.collection('default_events').insertMany(changedStudents.map((student) => ({
          eventId: randomUUID(),
          defaultId: null,
          rollNo: student.rollNo,
          type: 'student_batch_updated',
          memberId: actor.memberId,
          enteredByEmail: email,
          at: now,
          reason: '',
          note: '',
          dcReference: '',
          changes: { batch: { from: student.batch || 'Not set', to: batch || 'Not set' } },
        })), { session });
      }
      await db.collection('audit_log').insertOne({
        action: 'Student batches updated', collection: 'students', by: email, at: now,
        category: 'members', actor: actor.memberId,
        after: { requested: rollNos.length, updated: students.length, batch: batch || 'Not set' },
      }, { session });
      return { updated: students.length, notFound: rollNos.length - students.length };
    });
    return NextResponse.json({ ...updated, batch });
  } catch (error) {
    console.error('Unable to update student batches.', error);
    return NextResponse.json({ message: getApiErrorMessage(error, error instanceof Error ? error.message : 'Unable to update student batches.') }, { status: 500 });
  }
}
