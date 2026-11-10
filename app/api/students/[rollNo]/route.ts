import { randomUUID } from 'node:crypto';
import { NextResponse } from 'next/server';
import { getAuthorizedEmail } from '@/lib/authorization';
import { getApiErrorMessage } from '@/lib/database-errors';
import { findActorMember } from '@/lib/members';
import { getDb, withMongoTransaction } from '@/lib/mongodb';

export const runtime = 'nodejs';
export const dynamic = 'force-dynamic';

type RouteContext = { params: { rollNo: string } };

function normalizeRollNo(value: string) {
  return value.trim().toUpperCase();
}

export async function GET(_request: Request, { params }: RouteContext) {
  try {
    const email = await getAuthorizedEmail();
    if (!email) return NextResponse.json({ message: 'Sign in with an authorized account.' }, { status: 401 });
    const db = await getDb();
    const rollNoNormalized = normalizeRollNo(params.rollNo);
    const [student, defaults, members, types] = await Promise.all([
      db.collection('students').findOne({ rollNoNormalized }, { projection: { _id: 0 } }),
      db.collection('defaults').find({ $or: [{ rollNoNormalized }, { rollNo: params.rollNo }] }).sort({ date: -1 }).toArray(),
      db.collection('members').find({}).project({ _id: 0, memberId: 1, name: 1, team: 1 }).toArray(),
      db.collection('default_types').find({}).project({ _id: 0, id: 1, name: 1 }).toArray(),
    ]);
    if (!student) return NextResponse.json({ message: 'Student not found.' }, { status: 404 });

    const membersById = new Map(members.map((member) => [member.memberId, member]));
    const typeNames = new Map(types.map((type) => [type.id, type.name]));
    const dcEntries = await db.collection('dc_entries').find({
      removed: { $ne: true },
      $or: [{ rollNoNormalized }, { rollNo: student.rollNo }],
    }).toArray();
    const dcIds = Array.from(new Set(dcEntries.map((entry) => entry.dcId).filter(Boolean)));
    const dcs = dcIds.length
      ? await db.collection('dcs').find({ id: { $in: dcIds }, status: { $ne: 'Cancelled' } })
        .project({ _id: 0, id: 1, date: 1, time: 1, status: 1 }).toArray()
      : [];
    const dcById = new Map(dcs.map((dc) => [dc.id, dc]));
    const decisionsByDefaultId = new Map<string, Array<{
      dcId: string;
      dcDate: string;
      dcTime: string;
      decision: string;
      until: string;
      remarks: string;
    }>>();
    for (const entry of dcEntries) {
      if (entry.cancelled) continue;
      const dc = dcById.get(entry.dcId);
      if (!dc) continue;
      for (const item of entry.decisions ?? []) {
        const history = decisionsByDefaultId.get(item.defaultId) ?? [];
        history.push({
          dcId: dc.id,
          dcDate: dc.date,
          dcTime: dc.time ?? '',
          decision: item.decision ?? 'Undecided',
          until: item.until ?? '',
          remarks: entry.remarks ?? '',
        });
        decisionsByDefaultId.set(item.defaultId, history);
      }
    }
    decisionsByDefaultId.forEach((history) => {
      history.sort((a, b) => a.dcDate.localeCompare(b.dcDate) || a.dcTime.localeCompare(b.dcTime) || a.dcId.localeCompare(b.dcId));
    });

    return NextResponse.json({
      student: {
        ...student,
        numberOfDefaults: defaults.filter((record) => record.status !== 'removed').length,
        studentLocked: student.studentLocked === true,
      },
      defaults: defaults.map((record) => ({
        ...record,
        typeName: typeNames.get(record.typeId) ?? record.typeId,
        loggedByName: membersById.get(record.loggedByMemberId)?.name ?? 'Imported',
        team: membersById.get(record.loggedByMemberId)?.team ?? 'Imported',
        decisionHistory: decisionsByDefaultId.get(record.id) ?? [],
      })),
    });
  } catch (error) {
    console.error('Unable to load student history from MongoDB.', error);
    return NextResponse.json({ message: getApiErrorMessage(error, 'Unable to load student history.') }, { status: 500 });
  }
}

export async function PATCH(request: Request, { params }: RouteContext) {
  try {
    const enteredByEmail = await getAuthorizedEmail();
    if (!enteredByEmail) return NextResponse.json({ message: 'Sign in with an authorized account.' }, { status: 401 });
    const payload = await request.json();
    const action = String(payload?.action ?? '');
    if (action === 'set_batch') {
      const batchInput = String(payload?.batch ?? '').trim().toLowerCase();
      const batch = batchInput === 'junior' || batchInput === 'junior batch'
        ? 'junior'
        : batchInput === 'senior' || batchInput === 'senior batch'
          ? 'senior'
          : batchInput === '' ? '' : null;
      if (batch === null) return NextResponse.json({ message: 'Choose Junior batch, Senior batch, or Not set.' }, { status: 400 });
      await withMongoTransaction(async (db, session) => {
        const rollNoNormalized = normalizeRollNo(params.rollNo);
        const student = await db.collection('students').findOne({ rollNoNormalized }, { session });
        if (!student) throw new Error('NOT_FOUND');
        const actor = await findActorMember(db, enteredByEmail);
        if (!actor) throw new Error('Your account must be in the active member directory.');
        const now = new Date();
        await db.collection('students').updateOne(
          { _id: student._id },
          { $set: { batch, batchUpdatedAt: now, batchUpdatedByMemberId: actor.memberId } },
          { session },
        );
        if ((student.batch ?? '') !== batch) {
          await db.collection('default_events').insertOne({
            eventId: randomUUID(),
            defaultId: null,
            rollNo: student.rollNo,
            type: 'student_batch_updated',
            memberId: actor.memberId,
            enteredByEmail,
            at: now,
            reason: '',
            note: '',
            dcReference: '',
            changes: { batch: { from: student.batch || 'Not set', to: batch || 'Not set' } },
          }, { session });
        }
        await db.collection('audit_log').insertOne({
          action: 'Student batch updated', collection: 'students', by: enteredByEmail, at: now,
          category: 'members', actor: actor.memberId,
          before: { rollNo: student.rollNo, batch: student.batch ?? '' },
          after: { rollNo: student.rollNo, batch: batch || 'Not set' },
        }, { session });
      });
      return NextResponse.json({ updated: true, batch });
    }
    const reason = String(payload?.reason ?? '').trim();
    if (action !== 'lock' && action !== 'unlock') {
      return NextResponse.json({ message: 'Choose lock or unlock.' }, { status: 400 });
    }
    if (!reason) return NextResponse.json({ message: 'A reason is required.' }, { status: 400 });

    await withMongoTransaction(async (db, session) => {
      const rollNoNormalized = normalizeRollNo(params.rollNo);
      const student = await db.collection('students').findOne({ rollNoNormalized }, { session });
      if (!student) throw new Error('NOT_FOUND');
      const actor = await findActorMember(db, enteredByEmail);
      if (!actor) throw new Error('Your account must be in the active member directory to change student lock status.');
      const wasLocked = student.studentLocked === true;
      if ((action === 'lock' && wasLocked) || (action === 'unlock' && !wasLocked)) {
        throw new Error(action === 'lock' ? 'Student is already locked.' : 'Student is not locked.');
      }
      const now = new Date();
      const locked = action === 'lock';
      await db.collection('students').updateOne(
        { _id: student._id },
        { $set: { studentLocked: locked, studentLockUpdatedAt: now, studentLockUpdatedByMemberId: actor.memberId } },
        { session },
      );
      await db.collection('default_events').insertOne({
        eventId: randomUUID(),
        defaultId: null,
        rollNo: student.rollNo,
        type: locked ? 'student_locked' : 'student_unlocked',
        memberId: actor.memberId,
        enteredByEmail,
        at: now,
        reason,
        note: '',
        dcReference: '',
        changes: { studentLocked: { from: wasLocked, to: locked } },
      }, { session });
    });
    return NextResponse.json({ updated: true, studentLocked: action === 'lock' });
  } catch (error) {
    if (error instanceof Error && error.message === 'NOT_FOUND') {
      return NextResponse.json({ message: 'Student not found.' }, { status: 404 });
    }
    if (error instanceof Error && (error.message.startsWith('Student is') || error.message.startsWith('Your account'))) {
      return NextResponse.json({ message: error.message }, { status: 400 });
    }
    console.error('Unable to update student lock history.', error);
    return NextResponse.json({ message: getApiErrorMessage(error, 'Unable to update student lock status.') }, { status: 500 });
  }
}
