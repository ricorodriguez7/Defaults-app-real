import { randomUUID } from 'node:crypto';
import { NextResponse } from 'next/server';
import { getAuthorizedEmail } from '@/lib/authorization';
import { getApiErrorMessage } from '@/lib/database-errors';
import { getCurrentDc, getIndiaDate, syncCurrentDcAutoAdd } from '@/lib/dc-auto-add';
import { findActorMember } from '@/lib/members';
import { getDb, withMongoTransaction } from '@/lib/mongodb';

export const runtime = 'nodejs';
export const dynamic = 'force-dynamic';
export const maxDuration = 30;

function validDate(value: unknown): value is string {
  if (typeof value !== 'string' || !/^\d{4}-\d{2}-\d{2}$/.test(value)) return false;
  const parsed = new Date(`${value}T00:00:00.000Z`);
  return !Number.isNaN(parsed.getTime()) && parsed.toISOString().slice(0, 10) === value;
}

function validTime(value: unknown): value is string {
  return typeof value === 'string' && (!value || /^([01]\d|2[0-3]):[0-5]\d$/.test(value));
}

export async function GET() {
  try {
    const email = await getAuthorizedEmail();
    if (!email) return NextResponse.json({ message: 'Sign in with an authorized account.' }, { status: 401 });

    const db = await getDb();
    const now = new Date();
    const today = getIndiaDate(now);
    const requestActor = await findActorMember(db, email);
    if (!requestActor) throw new Error('Your account must be in the active member directory.');
    const dueDateDefaults = await db.collection('defaults').find({
      status: { $ne: 'removed' },
      dcStatus: 'Deferred',
      'deferral.until': { $type: 'string', $ne: 'next_dc', $lte: today },
    }).toArray();
    const plannedDcs = await db.collection('dcs').find({ status: 'Planned' }, { projection: { _id: 0, id: 1, date: 1 } }).toArray();
    const nextDcDeferredDefaults = await db.collection('defaults').find({
      status: { $ne: 'removed' },
      dcStatus: 'Deferred',
      'deferral.until': 'next_dc',
    }).toArray();
    const nextDueDefaults = nextDcDeferredDefaults.filter((record) =>
      plannedDcs.some((dc) => {
        const deferredAt = record.deferral?.at instanceof Date
          ? record.deferral.at
          : new Date(String(record.deferral?.at ?? ''));
        return !Number.isNaN(deferredAt.getTime()) && dc.date >= getIndiaDate(deferredAt);
      }),
    );
    const dueDefaults = [...dueDateDefaults, ...nextDueDefaults];
    if (dueDefaults.length) {
      for (const record of dueDefaults) {
        const eventId = `deferral-due:${record.id}:${record.deferral?.at instanceof Date ? record.deferral.at.toISOString() : String(record.deferral?.at ?? '')}`;
        const alreadyRecorded = await db.collection('default_events').findOne({ eventId }, { projection: { _id: 1 } });
        if (alreadyRecorded) continue;
        await withMongoTransaction(async (transactionDb, session) => {
          const current = await transactionDb.collection('defaults').findOne({ id: record.id, dcStatus: 'Deferred' }, { session });
          if (!current) return { unchanged: true };
          const deferralAt = current.deferral?.at instanceof Date ? current.deferral.at.toISOString() : String(current.deferral?.at ?? '');
          const stableEventId = `deferral-due:${current.id}:${deferralAt}`;
          if (await transactionDb.collection('default_events').findOne({ eventId: stableEventId }, { session, projection: { _id: 1 } })) {
            return { unchanged: true };
          }
          const actor = await transactionDb.collection('members').findOne({ memberId: requestActor.memberId, active: { $ne: false } }, { session });
          if (!actor) throw new Error('Your account must be in the active member directory.');
          await transactionDb.collection('defaults').updateOne(
            { id: current.id, dcStatus: 'Deferred' },
            { $set: { dcStatus: 'Not scheduled', deferralDueAt: now, updatedAt: now, updatedByMemberId: actor.memberId }, $unset: { dcId: '' } },
            { session },
          );
          await transactionDb.collection('default_events').insertOne({
            eventId: stableEventId, defaultId: current.id, rollNo: current.rollNo, type: 'deferral_due',
            memberId: actor.memberId, enteredByEmail: email, at: now,
            reason: current.deferral?.reason ?? '', note: 'Deferred default returned to Ready for DC.',
            dcReference: '', changes: { dcStatus: { from: 'Deferred', to: 'Not scheduled' } },
          }, { session });
          await transactionDb.collection('audit_log').insertOne({
            action: 'Deferred default due', collection: 'defaults', by: email, at: now,
            before: { defaultId: current.id, dcStatus: 'Deferred', until: current.deferral?.until },
            after: { defaultId: current.id, dcStatus: 'Not scheduled', dueAt: now, updatedBy: actor.memberId },
          }, { session });
          return { due: true };
        });
      }
    }

    const autoAdd = await withMongoTransaction(async (transactionDb, session) =>
      syncCurrentDcAutoAdd(transactionDb, session, requestActor.memberId, email, now),
    );
    const currentDcId = (await getCurrentDc(db))?.id ?? '';

    const [dcs, entries, defaults, students, members, types] = await Promise.all([
      db.collection('dcs').find({}, { projection: { _id: 0 } }).sort({ date: -1, createdAt: -1 }).toArray(),
      db.collection('dc_entries').find({ removed: { $ne: true } }, { projection: { _id: 0 } }).toArray(),
      db.collection('defaults').find({ status: { $ne: 'removed' } }, { projection: { _id: 0 } }).toArray(),
      db.collection('students').find({}, { projection: { _id: 0, rollNo: 1, rollNoNormalized: 1, name: 1 } }).toArray(),
      db.collection('members').find({}, { projection: { _id: 0, memberId: 1, name: 1 } }).toArray(),
      db.collection('default_types').find({}, { projection: { _id: 0, id: 1, name: 1 } }).toArray(),
    ]);

    const studentByRollNo = new Map(students.map((student) => [student.rollNoNormalized, student]));
    const memberById = new Map(members.map((member) => [member.memberId, member.name]));
    const typeNameById = new Map(types.map((type) => [type.id, type.name]));
    const entriesByDc = new Map<string, typeof entries>();
    const assignedStudents = new Set<string>();
    for (const entry of entries) {
      const current = entriesByDc.get(entry.dcId) ?? [];
      current.push(entry);
      entriesByDc.set(entry.dcId, current);
      if (entry.openStudentRollNo) assignedStudents.add(entry.openStudentRollNo);
    }

    const pendingByStudent = new Map<string, { rollNo: string; name: string; defaults: typeof defaults; oldestDate: string }>();
    const discussionDefaults: typeof defaults = [];
    const deferredDefaults: typeof defaults = [];
    const dueDefaultsForResponse: typeof defaults = [];
    for (const record of defaults) {
      if (record.dcStatus === 'To discuss') {
        discussionDefaults.push({ ...record, discussion: {
          ...record.discussion,
          raisedByName: memberById.get(record.discussion?.raisedBy) ?? 'Unknown member',
          notes: (record.discussion?.notes ?? []).map((note: { by: string; at: Date; text: string }) => ({
            ...note,
            byName: memberById.get(note.by) ?? 'Unknown member',
          })),
        } });
        continue;
      }
      if (record.dcStatus === 'Deferred') {
        const decorated = { ...record, deferral: {
          ...record.deferral,
          byName: memberById.get(record.deferral?.by) ?? 'Unknown member',
          history: (record.deferral?.history ?? []).map((item: { by: string; until: string; reason: string; at: Date }) => ({
            ...item,
            byName: memberById.get(item.by) ?? 'Unknown member',
          })),
        } };
        deferredDefaults.push(decorated);
        const until = record.deferral?.until;
        const dueNow = until === 'next_dc'
          ? plannedDcs.some((dc) => new Date(dc.createdAt).getTime() >= new Date(record.deferral?.at).getTime())
          : typeof until === 'string' && until <= today;
        if (dueNow) dueDefaultsForResponse.push(decorated);
        continue;
      }
      if (record.dcStatus === 'Not scheduled' && record.deferralDueAt instanceof Date &&
        record.deferralDueAt.toISOString().slice(0, 10) === today) {
        dueDefaultsForResponse.push({ ...record, deferral: { ...record.deferral, byName: memberById.get(record.deferral?.by) ?? 'Unknown member' } });
      }
      if (record.status !== 'logged' || record.dcStatus === 'Decided' || record.dcStatus === 'Scheduled') continue;
      const rollNoNormalized = String(record.rollNoNormalized || record.rollNo).toUpperCase();
      if (assignedStudents.has(rollNoNormalized)) continue;
      const student = studentByRollNo.get(rollNoNormalized);
      const current = pendingByStudent.get(rollNoNormalized) ?? {
        rollNo: record.rollNo,
        name: student?.name ?? record.rollNo,
        defaults: [] as typeof defaults,
        oldestDate: record.date || '',
      };
      current.defaults.push(record);
      if (record.date && (!current.oldestDate || record.date < current.oldestDate)) current.oldestDate = record.date;
      pendingByStudent.set(rollNoNormalized, current);
    }

    const todayStart = new Date();
    todayStart.setHours(0, 0, 0, 0);
    const pendingStudents = Array.from(pendingByStudent.entries())
      .map(([rollNoNormalized, item]) => {
        const oldest = item.oldestDate ? new Date(`${item.oldestDate}T00:00:00`) : todayStart;
        const daysWaiting = Number.isNaN(oldest.getTime())
          ? 0
          : Math.max(0, Math.floor((todayStart.getTime() - oldest.getTime()) / 86_400_000));
        return {
          rollNo: item.rollNo,
          rollNoNormalized,
          name: item.name,
          defaultCount: item.defaults.length,
          daysWaiting,
        };
      })
      .sort((a, b) => b.daysWaiting - a.daysWaiting || a.rollNo.localeCompare(b.rollNo));

    const dcSummaries = dcs.map((dc) => {
      const dcEntries = entriesByDc.get(dc.id) ?? [];
      const dcStudentDefaults = defaults.filter((record) => record.dcId === dc.id);
      return {
        ...dc,
        autoAdd: dc.autoAdd !== false,
        isCurrent: dc.id === currentDcId,
        counts: {
          studentsCalled: dcEntries.length,
          yetToAttend: dcEntries.filter((entry) => entry.attendance === 'Pending').length,
          attended: dcEntries.filter((entry) => entry.attendance === 'Present').length,
          absent: dcEntries.filter((entry) => entry.attendance === 'Absent').length,
          defaultsToBeRemoved: dc.status === 'Cancelled' ? 0 : dcEntries.reduce(
            (total, entry) => total + (entry.decisions ?? []).filter((decision: { decision: string }) => decision.decision === 'Remove').length,
            0,
          ),
          toDiscuss: dcStudentDefaults.filter((record) => record.dcStatus === 'To discuss').length,
          deferred: dcStudentDefaults.filter((record) => record.dcStatus === 'Deferred').length,
        },
      };
    });

    const defaultSummary = (record: (typeof defaults)[number]) => ({
      id: record.id,
      rollNo: record.rollNo,
      name: studentByRollNo.get(String(record.rollNoNormalized || record.rollNo).toUpperCase())?.name ?? record.rollNo,
      company: record.company,
      typeName: typeNameById.get(record.typeId) ?? record.typeId,
      date: record.date,
      dcStatus: record.dcStatus ?? 'Not scheduled',
      deferral: record.deferral ? {
        ...record.deferral,
        byName: memberById.get(record.deferral.by) ?? 'Unknown member',
        history: (record.deferral.history ?? []).map((item: { by: string; until: string; reason: string; at: Date }) => ({
          ...item,
          byName: memberById.get(item.by) ?? 'Unknown member',
        })),
      } : undefined,
      discussion: record.discussion ? {
        ...record.discussion,
        raisedByName: memberById.get(record.discussion.raisedBy) ?? 'Unknown member',
        notes: (record.discussion.notes ?? []).map((note: { by: string; at: Date; text: string }) => ({
          ...note,
          byName: memberById.get(note.by) ?? 'Unknown member',
        })),
      } : undefined,
    });
    const waitingForStudents = defaults.filter((record) =>
      record.status !== 'removed' && record.dcStatus === 'Scheduled',
    );
    const heldBackStudentRollNos = new Set([
      ...discussionDefaults.map((record) => String(record.rollNoNormalized || record.rollNo).toUpperCase()),
      ...deferredDefaults.map((record) => String(record.rollNoNormalized || record.rollNo).toUpperCase()),
    ]);
    const readyByStudent = new Map(pendingStudents.map((student) => [student.rollNo.toUpperCase(), student]));
    const readyStudentDetails = Array.from(readyByStudent.entries()).map(([rollNo, student]) => ({
      ...student,
      defaults: defaults.filter((record) =>
        String(record.rollNoNormalized || record.rollNo).toUpperCase() === rollNo &&
        record.status !== 'removed' &&
        record.status === 'logged' &&
        (!record.dcStatus || record.dcStatus === 'Not scheduled'),
      ).map(defaultSummary),
    }));
    const waitingStudents = waitingForStudents.length
      ? new Set(waitingForStudents.map((record) => String(record.rollNoNormalized || record.rollNo).toUpperCase()))
      : new Set<string>();

    return NextResponse.json({
      dcs: dcSummaries,
      autoAdd,
      pendingStudents: readyStudentDetails,
      pendingCount: readyStudentDetails.length,
      readyDefaultCount: defaults.filter((record) => record.status === 'logged' && (!record.dcStatus || record.dcStatus === 'Not scheduled')).length,
      toDiscussDefaults: discussionDefaults.map(defaultSummary),
      toDiscussCount: discussionDefaults.length,
      deferredDefaults: deferredDefaults.map(defaultSummary),
      deferredCount: deferredDefaults.length,
      dueDefaults: dueDefaultsForResponse.map(defaultSummary),
      dueCount: dueDefaultsForResponse.length,
      heldBackStudentCount: heldBackStudentRollNos.size,
      scheduledStudentCount: waitingStudents.size,
    });
  } catch (error) {
    console.error('Unable to load DC planner data.', error);
    return NextResponse.json({ message: getApiErrorMessage(error, 'Unable to load DC data or sync automatic student assignments. Try again, and contact an administrator if it continues.') }, { status: 500 });
  }
}

export async function POST(request: Request) {
  let requestAction = '';
  try {
    const email = await getAuthorizedEmail();
    if (!email) return NextResponse.json({ message: 'Sign in with an authorized account.' }, { status: 401 });

    const payload = await request.json();
    requestAction = String(payload?.action ?? '');
    if (payload?.action === 'sync') {
      const db = await getDb();
      const result = await withMongoTransaction(async (transactionDb, session) => {
        const actor = await findActorMember(transactionDb, email);
        if (!actor) throw new Error('Your account must be in the active member directory.');
        return syncCurrentDcAutoAdd(transactionDb, session, actor.memberId, email, new Date());
      });
      return NextResponse.json({ autoAdd: result });
    }

    const date = payload?.date;
    const time = String(payload?.time ?? '').trim();
    const venue = String(payload?.venue ?? '').trim();
    const notes = String(payload?.notes ?? '').trim();
    if (!validDate(date)) return NextResponse.json({ message: 'Enter a valid DC date.' }, { status: 400 });
    if (!validTime(time)) return NextResponse.json({ message: 'Enter a valid time.' }, { status: 400 });
    if (venue.length > 200 || notes.length > 2000) {
      return NextResponse.json({ message: 'Venue must be under 200 characters and notes under 2,000.' }, { status: 400 });
    }

    const now = new Date();
    const autoAdd = payload?.autoAdd !== false;
    const dc = {
      id: randomUUID(),
      date,
      time,
      venue,
      status: 'Planned',
      notes,
      autoAdd,
      isCurrent: true,
      cancelReason: '',
      createdBy: '',
      createdAt: now,
      completedBy: '',
      completedAt: null,
    };

    await getDb();
    const result = await withMongoTransaction(async (db, session) => {
      const actor = await findActorMember(db, email);
      if (!actor) throw new Error('Your account must be in the active member directory to plan a DC.');
      dc.createdBy = actor.memberId;
      await db.collection('dcs').insertOne(dc, { session });
      await db.collection('audit_log').insertOne({
        action: 'DC planned',
        collection: 'dcs',
        by: email,
        at: now,
        after: { dcId: dc.id, date, time, venue, notes },
      }, { session });
      const autoAdd = await syncCurrentDcAutoAdd(db, session, actor.memberId, email, now);
      return { planned: true, autoAdd };
    });
    return NextResponse.json({ dc, autoAdd: result.autoAdd }, { status: 201 });
  } catch (error) {
    if (error instanceof Error && error.message.startsWith('Your account')) {
      return NextResponse.json({ message: error.message }, { status: 403 });
    }
    console.error('Unable to plan DC.', error);
    return NextResponse.json({
      message: requestAction === 'sync'
        ? getApiErrorMessage(error, 'Unable to sync students automatically. Try again, and contact an administrator if it continues.')
        : getApiErrorMessage(error, 'Unable to plan or sync the DC.'),
    }, { status: 500 });
  }
}
