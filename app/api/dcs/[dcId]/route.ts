import { randomUUID } from 'node:crypto';
import { ClientSession, Db } from 'mongodb';
import { NextResponse } from 'next/server';
import { getAuthorizedEmail } from '@/lib/authorization';
import { getApiErrorMessage } from '@/lib/database-errors';
import { syncCurrentDcAutoAdd } from '@/lib/dc-auto-add';
import { findActorMember } from '@/lib/members';
import { getDb, withMongoTransaction } from '@/lib/mongodb';

export const runtime = 'nodejs';
export const dynamic = 'force-dynamic';
export const maxDuration = 30;

type RouteContext = { params: { dcId: string } };
type DCStatus = 'Planned' | 'In progress' | 'Completed' | 'Cancelled';
type Decision = {
  defaultId: string;
  decision: 'Undecided' | 'Remove' | 'Keep' | 'Discuss with team' | 'Defer';
  reason?: string;
  until?: string;
  discussionDate?: string;
};

function validDate(value: unknown): value is string {
  if (typeof value !== 'string' || !/^\d{4}-\d{2}-\d{2}$/.test(value)) return false;
  const parsed = new Date(`${value}T00:00:00.000Z`);
  return !Number.isNaN(parsed.getTime()) && parsed.toISOString().slice(0, 10) === value;
}

function validTime(value: unknown): value is string {
  return typeof value === 'string' && (!value || /^([01]\d|2[0-3]):[0-5]\d$/.test(value));
}

function addEvent(
  db: Db,
  session: ClientSession,
  values: Record<string, unknown>,
) {
  return db.collection('default_events').insertOne({
    eventId: randomUUID(),
    reason: '',
    note: '',
    changes: {},
    ...values,
  }, { session });
}

export async function GET(_request: Request, { params }: RouteContext) {
  try {
    const email = await getAuthorizedEmail();
    if (!email) return NextResponse.json({ message: 'Sign in with an authorized account.' }, { status: 401 });
    const db = await getDb();
    const requestActor = await findActorMember(db, email);
    if (!requestActor) return NextResponse.json({ message: 'Your account must be in the active member directory.' }, { status: 403 });
    const autoAdd = await withMongoTransaction(async (transactionDb, session) =>
      syncCurrentDcAutoAdd(transactionDb, session, requestActor.memberId, email, new Date()),
    );
    const dc = await db.collection('dcs').findOne({ id: params.dcId }, { projection: { _id: 0 } });
    if (!dc) return NextResponse.json({ message: 'DC not found.' }, { status: 404 });

    const [entries, types] = await Promise.all([
      db.collection('dc_entries').find({ dcId: params.dcId, removed: { $ne: true } }, { projection: { _id: 0 } }).sort({ rollNo: 1 }).toArray(),
      db.collection('default_types').find({}, { projection: { _id: 0, id: 1, name: 1 } }).toArray(),
    ]);
    const rollNos = entries.map((entry) => entry.rollNo);
    const students = await db.collection('students').find({ rollNo: { $in: rollNos } })
      .project({ _id: 0, rollNo: 1, name: 1 }).toArray();
    const defaultIds = Array.from(new Set(entries.flatMap((entry) => (entry.decisions ?? []).map((item: Decision) => item.defaultId))));
    const defaults = await db.collection('defaults').find({ id: { $in: defaultIds } }, { projection: { _id: 0, rollNoNormalized: 0 } }).toArray();
    const studentNames = new Map(students.map((student) => [student.rollNo.toUpperCase(), student.name]));
    const defaultById = new Map(defaults.map((record) => [record.id, record]));
    const typeNames = new Map(types.map((type) => [type.id, type.name]));
    const decoratedEntries = entries.map((entry) => ({
      ...entry,
      studentName: studentNames.get(String(entry.rollNo).toUpperCase()) ?? entry.rollNo,
      defaults: (entry.decisions ?? []).map((item: Decision) => {
        const record = defaultById.get(item.defaultId);
        return record ? {
          ...record,
          typeName: typeNames.get(record.typeId) ?? record.typeId,
          decision: item.decision,
          reason: item.reason ?? '',
          until: item.until ?? '',
          discussionDate: item.discussionDate ?? '',
        } : null;
      }).filter(Boolean),
    }));

    return NextResponse.json({ dc, entries: decoratedEntries, autoAdd });
  } catch (error) {
    console.error('Unable to load DC details.', error);
    return NextResponse.json({ message: getApiErrorMessage(error, 'Unable to load the DC planner or sync automatic student assignments. Try again, and contact an administrator if it continues.') }, { status: 500 });
  }
}

export async function PATCH(request: Request, { params }: RouteContext) {
  try {
    const email = await getAuthorizedEmail();
    if (!email) return NextResponse.json({ message: 'Sign in with an authorized account.' }, { status: 401 });
    const payload = await request.json();
    const action = String(payload?.action ?? '');
    const now = new Date();
    const dcId = params.dcId;
    await getDb();

    const result = await withMongoTransaction(async (db, session) => {
      const dcs = db.collection('dcs');
      const dc = await dcs.findOne({ id: dcId }, { session });
      if (!dc) throw new Error('DC_NOT_FOUND');
      const actor = await findActorMember(db, email);
      if (!actor) throw new Error('ACTOR_REQUIRED');
      const entries = db.collection('dc_entries');
      const defaults = db.collection('defaults');
      const events = db.collection('default_events');

      if (action === 'reschedule') {
        if (dc.status !== 'Planned') throw new Error('Only a planned DC can be rescheduled.');
        if (!validDate(payload?.date)) throw new Error('Enter a valid DC date.');
        const time = String(payload?.time ?? '').trim();
        const venue = String(payload?.venue ?? '').trim();
        const notes = String(payload?.notes ?? '').trim();
        const reason = String(payload?.reason ?? '').trim();
        if (!validTime(time)) throw new Error('Enter a valid time.');
        if (!reason) throw new Error('A reason is required to reschedule a DC.');
        if (venue.length > 200 || notes.length > 2000) throw new Error('Venue or notes exceed the allowed length.');
        const before = { date: dc.date, time: dc.time, venue: dc.venue, notes: dc.notes };
        const after = { date: payload.date, time, venue, notes };
        await dcs.updateOne({ id: dcId }, { $set: { ...after, updatedAt: now, updatedBy: actor.memberId } }, { session });
        const entryList = await entries.find({ dcId, removed: { $ne: true } }, { session }).toArray();
        const ids = entryList.flatMap((entry) => (entry.decisions ?? []).map((item: Decision) => item.defaultId));
        const relatedDefaults = await defaults.find({ id: { $in: ids } }, { session }).toArray();
        for (const record of relatedDefaults) {
          await addEvent(db, session, {
            defaultId: record.id, rollNo: record.rollNo, type: 'dc_rescheduled', memberId: actor.memberId,
            enteredByEmail: email, at: now, reason, dcReference: payload.date, dcId,
            changes: { date: { from: dc.date, to: payload.date } },
          });
        }
        await db.collection('audit_log').insertOne({
          action: 'DC rescheduled', collection: 'dcs', by: email, at: now,
          before: { dcId, ...before }, after: { dcId, ...after, reason },
        }, { session });
        const autoAdd = await syncCurrentDcAutoAdd(db, session, actor.memberId, email, now);
        return { updated: true, autoAdd };
      }

      if (action === 'cancel') {
        if (dc.status !== 'Planned' && dc.status !== 'In progress') throw new Error('Only an active DC can be cancelled.');
        const reason = String(payload?.reason ?? '').trim();
        if (!reason) throw new Error('A cancellation reason is required.');
        const dcEntries = await entries.find({ dcId, removed: { $ne: true } }, { session }).toArray();
        const defaultIds = dcEntries.flatMap((entry) => (entry.decisions ?? []).map((item: Decision) => item.defaultId));
        const relatedDefaults = await defaults.find({ id: { $in: defaultIds }, dcId }, { session }).toArray();
        await dcs.updateOne({ id: dcId }, {
          $set: { status: 'Cancelled', cancelReason: reason, cancelledBy: actor.memberId, cancelledAt: now },
        }, { session });
        for (const record of relatedDefaults) {
          await defaults.updateOne({ id: record.id }, { $set: { dcStatus: 'Not scheduled' }, $unset: { dcId: '' } }, { session });
          await addEvent(db, session, {
            defaultId: record.id, rollNo: record.rollNo, type: 'dc_cancelled', memberId: actor.memberId,
            enteredByEmail: email, at: now, reason, dcReference: dc.date, dcId,
            changes: { dcStatus: { from: 'Scheduled', to: 'Not scheduled' } },
          });
        }
        for (const entry of dcEntries) {
          await entries.updateOne({ _id: entry._id }, {
            $set: {
              cancelled: true,
              attendance: 'Pending',
              absentReason: '',
              decisions: (entry.decisions ?? []).map((item: Decision) => ({ ...item, decision: 'Undecided' })),
            },
            $unset: { openStudentRollNo: '' },
          }, { session });
        }
        await db.collection('audit_log').insertOne({
          action: 'DC cancelled', collection: 'dcs', by: email, at: now,
          before: { dcId, status: dc.status }, after: { dcId, status: 'Cancelled', reason },
        }, { session });
        const autoAdd = await syncCurrentDcAutoAdd(db, session, actor.memberId, email, now);
        return { updated: true, autoAdd };
      }

      if (action === 'add_students') {
        if (dc.status !== 'Planned') throw new Error('Students can only be added to a planned DC.');
        const rolls = Array.isArray(payload?.rollNos)
          ? Array.from(new Set(payload.rollNos.map((rollNo: unknown) => String(rollNo ?? '').trim().toUpperCase()).filter(Boolean)))
          : [];
        if (!rolls.length || rolls.length > 200) throw new Error('Select between 1 and 200 students.');
        const createdEntries = [];
        for (const rollNoNormalized of rolls) {
          const student = await db.collection('students').findOne({ rollNoNormalized }, { session });
          if (!student) throw new Error(`Unknown student roll number: ${rollNoNormalized}.`);
          const existingEntry = await entries.findOne({ openStudentRollNo: rollNoNormalized }, { session });
          if (existingEntry) throw new Error(`${student.name} is already assigned to an open DC.`);
          const studentDefaults = await defaults.find({
            rollNoNormalized,
            status: 'logged',
            $or: [{ dcStatus: 'Not scheduled' }, { dcStatus: { $exists: false } }],
          }, { session }).toArray();
          if (!studentDefaults.length) throw new Error(`${student.name} has no undecided defaults to schedule.`);
          const decisions: Decision[] = studentDefaults.map((record) => ({ defaultId: record.id, decision: 'Undecided' }));
          const entry = {
            dcId,
            rollNo: student.rollNo,
            rollNoNormalized,
            attendance: 'Pending',
            absentReason: '',
            remarks: '',
            decisions,
            source: 'manual',
            excluded: false,
            openStudentRollNo: rollNoNormalized,
            createdAt: now,
            createdBy: actor.memberId,
          };
          const previousEntry = await entries.findOne({ dcId, rollNoNormalized }, { session });
          if (previousEntry) {
            await entries.updateOne(
              { _id: previousEntry._id },
              { $set: entry, $unset: { removed: '', removalReason: '', removedAt: '', cancelled: '' } },
              { session },
            );
          } else {
            await entries.insertOne(entry, { session });
          }
          for (const record of studentDefaults) {
            await defaults.updateOne({ id: record.id }, { $set: { dcId, dcStatus: 'Scheduled' } }, { session });
            await addEvent(db, session, {
              defaultId: record.id, rollNo: record.rollNo, type: 'dc_scheduled', memberId: actor.memberId,
              enteredByEmail: email, at: now, reason: '', dcReference: dc.date, dcId,
              changes: { dcStatus: { from: record.dcStatus ?? 'Not scheduled', to: 'Scheduled' } },
            });
          }
          createdEntries.push({ rollNo: student.rollNo, defaults: studentDefaults.length });
        }
        await db.collection('audit_log').insertOne({
          action: 'Students added to DC', collection: 'dc_entries', by: email, at: now,
          after: { dcId, students: createdEntries },
        }, { session });
        const autoAdd = await syncCurrentDcAutoAdd(db, session, actor.memberId, email, now);
        return { added: createdEntries.length, autoAdd };
      }

      if (action === 'remove_student') {
        if (dc.status !== 'Planned') throw new Error('Students can only be removed from a planned DC.');
        const rollNoNormalized = String(payload?.rollNo ?? '').trim().toUpperCase();
        const reason = String(payload?.reason ?? '').trim();
        if (!rollNoNormalized || !reason) throw new Error('A student and removal reason are required.');
        const entry = await entries.findOne({ dcId, rollNoNormalized, removed: { $ne: true } }, { session });
        if (!entry) throw new Error('Student is not assigned to this DC.');
        const ids = (entry.decisions ?? []).map((item: Decision) => item.defaultId);
        await defaults.updateMany(
          { id: { $in: ids }, dcId },
          { $set: { dcStatus: 'Not scheduled' }, $unset: { dcId: '' } },
          { session },
        );
        const records = await defaults.find({ id: { $in: ids } }, { session }).toArray();
        for (const record of records) {
          await addEvent(db, session, {
            defaultId: record.id, rollNo: record.rollNo, type: 'dc_unscheduled', memberId: actor.memberId,
            enteredByEmail: email, at: now, reason: `Removed from planned DC: ${reason}`,
            dcReference: dc.date, dcId,
            changes: { dcStatus: { from: 'Scheduled', to: 'Not scheduled' } },
          });
        }
        await entries.updateOne({ _id: entry._id }, {
          $set: { removed: true, excluded: true, removalReason: reason, removedAt: now },
          $unset: { openStudentRollNo: '' },
        }, { session });
        await db.collection('audit_log').insertOne({
          action: 'Student removed from DC', collection: 'dc_entries', by: email, at: now,
          before: { dcId, rollNo: entry.rollNo }, after: { dcId, rollNo: entry.rollNo, reason },
        }, { session });
        const autoAdd = await syncCurrentDcAutoAdd(db, session, actor.memberId, email, now);
        return { removed: true, autoAdd };
      }

      if (action === 'attendance') {
        if (dc.status !== 'Planned' && dc.status !== 'In progress') throw new Error('Attendance can only be changed for an active DC.');
        const rollNoNormalized = String(payload?.rollNo ?? '').trim().toUpperCase();
        const attendance = String(payload?.attendance ?? '');
        const absentReason = String(payload?.absentReason ?? '').trim();
        if (!['Pending', 'Present', 'Absent'].includes(attendance)) throw new Error('Choose Pending, Present, or Absent.');
        if (absentReason.length > 500) throw new Error('The absent reason must be under 500 characters.');
        const entry = await entries.findOne({ dcId, rollNoNormalized, removed: { $ne: true } }, { session });
        if (!entry) throw new Error('Student is not assigned to this DC.');
        if (entry.attendance === attendance && entry.absentReason === absentReason) return { updated: true };
        await entries.updateOne({ _id: entry._id }, { $set: { attendance, absentReason, attendanceUpdatedAt: now, attendanceUpdatedBy: actor.memberId } }, { session });
        if (dc.status === 'Planned') {
          await dcs.updateOne({ id: dcId, status: 'Planned' }, { $set: { status: 'In progress', updatedAt: now } }, { session });
        }
        const ids = (entry.decisions ?? []).map((item: Decision) => item.defaultId);
        const records = await defaults.find({ id: { $in: ids } }, { session }).toArray();
        for (const record of records) {
          await addEvent(db, session, {
            defaultId: record.id, rollNo: record.rollNo, type: 'dc_attendance', memberId: actor.memberId,
            enteredByEmail: email, at: now, reason: attendance === 'Absent' ? absentReason : '',
            note: attendance, dcReference: dc.date, dcId,
            changes: { attendance: { from: entry.attendance, to: attendance } },
          });
        }
        await db.collection('audit_log').insertOne({
          action: 'DC attendance updated', collection: 'dc_entries', by: email, at: now,
          before: { dcId, rollNo: entry.rollNo, attendance: entry.attendance, absentReason: entry.absentReason },
          after: { dcId, rollNo: entry.rollNo, attendance, absentReason, updatedBy: actor.memberId },
        }, { session });
        return { updated: true };
      }

      if (action === 'mark_pending_absent') {
        if (dc.status !== 'Planned' && dc.status !== 'In progress') throw new Error('Attendance can only be changed for an active DC.');
        const pendingEntries = await entries.find({ dcId, removed: { $ne: true }, attendance: 'Pending' }, { session }).toArray();
        if (!pendingEntries.length) return { updated: true, count: 0 };
        const rollNos = pendingEntries.map((entry) => entry.rollNo);
        const ids = pendingEntries.flatMap((entry) => (entry.decisions ?? []).map((item: Decision) => item.defaultId));
        const records = await defaults.find({ id: { $in: ids } }, { session }).toArray();
        const defaultsByRollNo = new Map<string, typeof records>();
        for (const record of records) {
          const rollNo = String(record.rollNo).trim().toUpperCase();
          const current = defaultsByRollNo.get(rollNo) ?? [];
          current.push(record);
          defaultsByRollNo.set(rollNo, current);
        }
        await entries.updateMany(
          { dcId, removed: { $ne: true }, attendance: 'Pending' },
          { $set: { attendance: 'Absent', absentReason: '', attendanceUpdatedAt: now, attendanceUpdatedBy: actor.memberId } },
          { session },
        );
        if (dc.status === 'Planned') {
          await dcs.updateOne({ id: dcId, status: 'Planned' }, { $set: { status: 'In progress', updatedAt: now } }, { session });
        }
        for (const entry of pendingEntries) {
          for (const record of defaultsByRollNo.get(String(entry.rollNo).trim().toUpperCase()) ?? []) {
            await addEvent(db, session, {
              defaultId: record.id, rollNo: record.rollNo, type: 'dc_attendance', memberId: actor.memberId,
              enteredByEmail: email, at: now, reason: '', note: 'Absent',
              dcReference: dc.date, dcId,
              changes: { attendance: { from: 'Pending', to: 'Absent' } },
            });
          }
        }
        await db.collection('audit_log').insertOne({
          action: 'Pending DC students marked absent', collection: 'dc_entries', by: email, at: now,
          after: { dcId, count: pendingEntries.length, rollNos, updatedBy: actor.memberId },
        }, { session });
        return { updated: true, count: pendingEntries.length };
      }

      if (action === 'decision') {
        if (dc.status !== 'Planned' && dc.status !== 'In progress') throw new Error('Decisions can only be changed for an active DC.');
        const rollNoNormalized = String(payload?.rollNo ?? '').trim().toUpperCase();
        const defaultId = String(payload?.defaultId ?? '').trim();
        const decision = String(payload?.decision ?? '');
        if (!['Undecided', 'Remove', 'Keep', 'Discuss with team', 'Defer'].includes(decision)) {
          throw new Error('Choose Remove, Keep, Undecided, Discuss with team, or Defer.');
        }
        const reason = String(payload?.reason ?? '').trim();
        const until = String(payload?.until ?? '').trim();
        const discussionDate = String(payload?.discussionDate ?? '').trim();
        if (reason.length > 2000) throw new Error('Decision reason must be under 2,000 characters.');
        if (decision === 'Discuss with team' && !reason) throw new Error('A reason is required to discuss a default with the team.');
        if (decision === 'Discuss with team' && discussionDate && !validDate(discussionDate)) throw new Error('Enter a valid discussion date.');
        if (decision === 'Defer' && !reason) throw new Error('A deferral reason is required.');
        if (decision === 'Defer' && until !== 'next_dc' && !validDate(until)) throw new Error('Choose a valid deferral date or next DC.');
        const entry = await entries.findOne({ dcId, rollNoNormalized, removed: { $ne: true } }, { session });
        if (!entry) throw new Error('Student is not assigned to this DC.');
        const existing = (entry.decisions ?? []).find((item: Decision) => item.defaultId === defaultId);
        if (!existing) throw new Error('Default is not part of this DC entry.');
        if (existing.decision === decision) return { updated: true };
        await entries.updateOne(
          { _id: entry._id, 'decisions.defaultId': defaultId },
          {
            $set: {
              'decisions.$.decision': decision,
              'decisions.$.reason': reason,
              ...(decision === 'Defer' ? { 'decisions.$.until': until } : {}),
              ...(decision === 'Discuss with team' ? { 'decisions.$.discussionDate': discussionDate } : {}),
              decisionsUpdatedAt: now,
              decisionsUpdatedBy: actor.memberId,
            },
            $unset: {
              ...(decision !== 'Defer' ? { 'decisions.$.until': '' } : {}),
              ...(decision !== 'Discuss with team' ? { 'decisions.$.discussionDate': '' } : {}),
            },
          },
          { session },
        );
        const record = await defaults.findOne({ id: defaultId }, { session });
        if (!record) throw new Error('Default is not part of this DC entry.');
        await addEvent(db, session, {
          defaultId, rollNo: record.rollNo, type: 'dc_decision', memberId: actor.memberId,
          enteredByEmail: email, at: now, reason, note: `Proposed decision: ${decision}`,
          dcReference: dc.date, dcId,
          changes: { decision: { from: existing.decision, to: decision } },
        });
        await db.collection('audit_log').insertOne({
          action: 'DC default decision updated', collection: 'dc_entries', by: email, at: now,
          before: { dcId, rollNo: entry.rollNo, defaultId, decision: existing.decision },
          after: { dcId, rollNo: entry.rollNo, defaultId, decision, until, reason, discussionDate, updatedBy: actor.memberId },
        }, { session });
        return { updated: true };
      }

      if (action === 'remarks') {
        if (dc.status !== 'Planned' && dc.status !== 'In progress') throw new Error('Remarks can only be changed for an active DC.');
        const rollNoNormalized = String(payload?.rollNo ?? '').trim().toUpperCase();
        const remarks = String(payload?.remarks ?? '').trim();
        if (remarks.length > 2000) throw new Error('Remarks must be under 2,000 characters.');
        const entry = await entries.findOne({ dcId, rollNoNormalized, removed: { $ne: true } }, { session });
        if (!entry) throw new Error('Student is not assigned to this DC.');
        if (entry.remarks === remarks) return { updated: true };
        await entries.updateOne({ _id: entry._id }, { $set: { remarks, remarksUpdatedAt: now, remarksUpdatedBy: actor.memberId } }, { session });
        await db.collection('audit_log').insertOne({
          action: 'DC student remarks updated', collection: 'dc_entries', by: email, at: now,
          before: { dcId, rollNo: entry.rollNo, remarks: entry.remarks },
          after: { dcId, rollNo: entry.rollNo, remarks },
        }, { session });
        return { updated: true };
      }

      if (action === 'complete') {
        if (dc.status !== 'Planned' && dc.status !== 'In progress') throw new Error('Only an active DC can be completed.');
        const dcEntries = await entries.find({ dcId, removed: { $ne: true } }, { session }).toArray();
        if (dcEntries.some((entry) => entry.attendance === 'Pending')) throw new Error('Mark every student Present or Absent before completing the DC.');
        const summary = { present: 0, absent: 0, remove: 0, keep: 0, undecided: 0, discuss: 0, defer: 0 };
        for (const entry of dcEntries) {
          if (entry.attendance === 'Present') summary.present += 1;
          else summary.absent += 1;
          if (entry.attendance !== 'Present') continue;
          for (const decision of (entry.decisions ?? []) as Decision[]) {
            if (decision.decision === 'Remove') summary.remove += 1;
            else if (decision.decision === 'Keep') summary.keep += 1;
            else if (decision.decision === 'Discuss with team') summary.discuss += 1;
            else if (decision.decision === 'Defer') summary.defer += 1;
            else summary.undecided += 1;
          }
        }
        const allDefaultIds = dcEntries.flatMap((entry) => (entry.decisions ?? []).map((item: Decision) => item.defaultId));
        const currentDefaults = await defaults.find({ id: { $in: allDefaultIds }, dcId }, { session }).toArray();
        const defaultById = new Map(currentDefaults.map((record) => [record.id, record]));
        for (const entry of dcEntries) {
          const isPresent = entry.attendance === 'Present';
          for (const item of (entry.decisions ?? []) as Decision[]) {
            const record = defaultById.get(item.defaultId);
              if (!record) throw new Error(`A default for ${entry.rollNo} is missing; the DC cannot be completed.`);
              if (!isPresent) {
              await defaults.updateOne({ id: record.id }, {
                $set: { dcStatus: 'Not scheduled', updatedAt: now, updatedByMemberId: actor.memberId },
                $unset: { dcId: '' },
              }, { session });
              continue;
            }
            const set: Record<string, unknown> = {
              dcStatus: item.decision === 'Undecided' ? 'Not scheduled'
                : item.decision === 'Discuss with team' ? 'To discuss'
                  : item.decision === 'Defer' ? 'Deferred' : 'Decided',
              updatedAt: now,
              updatedByMemberId: actor.memberId,
            };
            const unset: Record<string, ''> = {};
            if (item.decision === 'Undecided') {
              unset.dcId = '';
            } else if (item.decision === 'Remove' || item.decision === 'Keep') {
              const nextStatus = item.decision === 'Remove' ? 'removed' : 'locked';
              set.dcId = dcId;
              set.dcReference = dc.date;
              set.status = nextStatus;
              if (item.decision === 'Remove') set.removedAt = now;
              if (item.decision === 'Keep') set.lockedAt = now;
            } else if (item.decision === 'Discuss with team') {
              set.discussion = {
                raisedBy: actor.memberId,
                raisedAt: now,
                reason: item.reason || '',
                discussionDate: item.discussionDate || '',
                notes: [],
              };
              set.dcReference = dc.date;
            } else if (item.decision === 'Defer') {
              const history = [...(record.deferral?.history ?? []), {
                until: item.until || '', reason: item.reason || '', by: actor.memberId, at: now,
              }];
              set.deferral = {
                until: item.until || '',
                reason: item.reason || '',
                by: actor.memberId,
                at: now,
                history,
              };
              set.dcReference = dc.date;
            }
            await defaults.updateOne({ id: record.id }, {
              $set: set,
              ...(Object.keys(unset).length ? { $unset: unset } : {}),
            }, { session });
            await addEvent(db, session, {
              defaultId: record.id, rollNo: record.rollNo, type: 'dc_decision', memberId: actor.memberId,
              enteredByEmail: email, at: now,
              reason: item.decision === 'Remove'
                ? entry.remarks || 'Removed by Disciplinary Committee.'
                : item.reason || entry.remarks || '',
              note: item.decision,
              dcReference: dc.date, dcId,
              changes: {
                decision: { from: 'Undecided', to: item.decision },
                ...(item.decision === 'Remove' ? { status: { from: record.status, to: 'removed' } } : {}),
                ...(item.decision === 'Keep' ? { status: { from: record.status, to: 'locked' } } : {}),
                dcStatus: { from: 'Scheduled', to: set.dcStatus },
              },
            });
            if (item.decision === 'Discuss with team' || item.decision === 'Defer') {
              await addEvent(db, session, {
                defaultId: record.id,
                rollNo: record.rollNo,
                type: item.decision === 'Discuss with team' ? 'discuss_flagged' : 'deferred',
                memberId: actor.memberId,
                enteredByEmail: email,
                at: now,
                reason: item.reason || '',
                note: item.decision === 'Defer' ? `Until ${item.until || ''}` : '',
                dcReference: item.decision === 'Defer' ? item.until || dc.date : item.discussionDate || dc.date,
                dcId,
                changes: { dcStatus: { from: 'Scheduled', to: set.dcStatus } },
              });
            }
          }
        }
        await dcs.updateOne({ id: dcId }, {
          $set: { status: 'Completed', completedBy: actor.memberId, completedAt: now, completionSummary: summary },
        }, { session });
        await entries.updateMany({ dcId, removed: { $ne: true } }, {
          $set: { completedAt: now },
          $unset: { openStudentRollNo: '' },
        }, { session });
        await db.collection('audit_log').insertOne({
          action: 'DC completed', collection: 'dcs', by: email, at: now,
          before: { dcId, status: dc.status },
          after: { dcId, status: 'Completed', summary },
        }, { session });
        const autoAdd = await syncCurrentDcAutoAdd(db, session, actor.memberId, email, now);
        return { completed: true, summary, autoAdd };
      }

      throw new Error('Choose a valid DC action.');
    });
    return NextResponse.json(result);
  } catch (error) {
    const message = error instanceof Error ? error.message : '';
    if (message === 'DC_NOT_FOUND') return NextResponse.json({ message: 'DC not found.' }, { status: 404 });
    if (message === 'ACTOR_REQUIRED') return NextResponse.json({ message: 'Your account must be in the active member directory to manage DCs.' }, { status: 403 });
    if (message.startsWith('Only ') || message.startsWith('Students can only') || message.startsWith('A reason') ||
      message.startsWith('A deferral') || message.startsWith('Decision reason') ||
      message.startsWith('Attendance can only') || message.startsWith('Decisions can only') ||
      message.startsWith('A cancellation') || message.startsWith('Enter a valid') || message.startsWith('Choose ') ||
      message.startsWith('Venue ') || message.startsWith('Select between') || message.startsWith('Unknown student') ||
      message.includes('already assigned') || message.includes('no undecided defaults') ||
      message.startsWith('A student') || message.startsWith('Student is not assigned') ||
      message.startsWith('The absent reason') || message.startsWith('Default is not part') ||
      message.startsWith('Remarks can only') || message.startsWith('Remarks must') ||
      message.startsWith('Mark every') || message.startsWith('A default for') ||
      message.startsWith('Attendance can only') || message.startsWith('Only an active') ||
      message.startsWith('Decisions can only') || message.startsWith('Remarks can only')) {
      return NextResponse.json({ message }, { status: 400 });
    }
    if (error && typeof error === 'object' && 'code' in error && error.code === 11000) {
      console.error('Duplicate DC assignment during update.', error);
      return NextResponse.json({ message: 'This student is already assigned to an open DC.' }, { status: 409 });
    }
    console.error('Unable to update DC.', error);
    return NextResponse.json({ message: getApiErrorMessage(error, 'Unable to update the DC or sync automatic student assignments. Try again, and contact an administrator if it continues.') }, { status: 500 });
  }
}
