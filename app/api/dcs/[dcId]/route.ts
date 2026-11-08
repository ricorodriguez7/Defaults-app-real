import { randomUUID } from 'node:crypto';
import { ClientSession, Db } from 'mongodb';
import { NextResponse } from 'next/server';
import { getAuthorizedEmail } from '@/lib/authorization';
import { syncCurrentDcAutoAdd } from '@/lib/dc-auto-add';
import { findActorMember } from '@/lib/members';
import { databaseFailureMessage, getDb, withMongoTransaction } from '@/lib/mongodb';

export const runtime = 'nodejs';
export const dynamic = 'force-dynamic';
export const maxDuration = 30;

type RouteContext = { params: { dcId: string } };
type DCStatus = 'Planned' | 'Held' | 'Finalised' | 'Cancelled';
type Decision = {
  defaultId: string;
  decision: 'Undecided' | 'Remove' | 'Keep' | 'Discuss with team' | 'Defer';
  reason?: string;
  until?: string;
  discussionDate?: string;
  recommendation?: string;
  recommendationNotes?: string;
  recommendationBy?: string;
  recommendationAt?: Date;
  recommendationHistory?: Array<{ recommendation: string; notes: string; by: string; at: Date }>;
  finalDecision?: string;
  finalRemarks?: string;
  finalUntil?: string;
  finalBy?: string;
  finalAt?: Date;
  discussionNotes?: Array<{ text: string; by: string; at: Date }>;
};

function recommendationOf(item: Decision) {
  if (item.recommendation) return item.recommendation;
  return item.decision === 'Remove' ? 'remove'
    : item.decision === 'Keep' ? 'keep'
      : item.decision === 'Defer' ? 'defer'
        : item.decision === 'Discuss with team' ? 'discuss' : 'undecided';
}

function recommendationLabel(value: string) {
  return value === 'remove' ? 'Remove'
    : value === 'keep' ? 'Keep'
      : value === 'defer' ? 'Defer'
        : value === 'discuss' ? 'Discuss with team' : 'Undecided';
}

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

    const [entries, types, members] = await Promise.all([
      db.collection('dc_entries').find({ dcId: params.dcId, removed: { $ne: true } }, { projection: { _id: 0 } }).sort({ rollNo: 1 }).toArray(),
      db.collection('default_types').find({}, { projection: { _id: 0, id: 1, name: 1 } }).toArray(),
      db.collection('members').find({}, { projection: { _id: 0, memberId: 1, name: 1 } }).toArray(),
    ]);
    const rollNos = entries.map((entry) => entry.rollNo);
    const students = await db.collection('students').find({ rollNo: { $in: rollNos } })
      .project({ _id: 0, rollNo: 1, name: 1 }).toArray();
    const defaultIds = Array.from(new Set(entries.flatMap((entry) => (entry.decisions ?? []).map((item: Decision) => item.defaultId))));
    const defaults = await db.collection('defaults').find({ id: { $in: defaultIds } }, { projection: { _id: 0, rollNoNormalized: 0 } }).toArray();
    const studentNames = new Map(students.map((student) => [student.rollNo.toUpperCase(), student.name]));
    const defaultById = new Map(defaults.map((record) => [record.id, record]));
    const typeNames = new Map(types.map((type) => [type.id, type.name]));
    const memberNames = new Map(members.map((member) => [member.memberId, member.name]));
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
          recommendation: recommendationOf(item),
          recommendationNotes: item.recommendationNotes ?? item.reason ?? '',
          recommendationBy: item.recommendationBy ?? '',
          recommendationAt: item.recommendationAt ?? null,
          recommendationByName: memberNames.get(item.recommendationBy) ?? '',
          recommendationHistory: (item.recommendationHistory ?? []).map((history: { by?: string }) => ({
            ...history,
            byName: memberNames.get(history.by ?? '') ?? '',
          })),
          discussionNotes: (item.discussionNotes ?? []).map((note: { by?: string }) => ({
            ...note,
            byName: memberNames.get(note.by ?? '') ?? '',
          })),
          finalDecision: item.finalDecision ?? '',
          finalRemarks: item.finalRemarks ?? '',
          finalUntil: item.finalUntil ?? '',
          finalBy: item.finalBy ?? '',
          finalAt: item.finalAt ?? null,
          finalByName: memberNames.get(item.finalBy ?? '') ?? '',
        } : null;
      }).filter(Boolean),
    }));

    return NextResponse.json({
      dc: {
        ...dc,
        finalisedByName: dc.finalisedBy ? memberNames.get(dc.finalisedBy) ?? '' : '',
        heldByName: dc.heldBy ? memberNames.get(dc.heldBy) ?? '' : '',
      },
      entries: decoratedEntries,
      autoAdd,
    });
  } catch (error) {
    console.error('Unable to load DC details.', error);
    return NextResponse.json({ message: databaseFailureMessage(error, 'Unable to load the DC planner or sync automatic student assignments. Try again, and contact an administrator if it continues.') }, { status: 500 });
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
        if (dc.status !== 'Planned') throw new Error('Only a planned DC can be cancelled.');
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
        if (dc.status !== 'Planned') throw new Error('Attendance can only be changed for a planned DC.');
        const rollNoNormalized = String(payload?.rollNo ?? '').trim().toUpperCase();
        const attendance = String(payload?.attendance ?? '');
        const absentReason = String(payload?.absentReason ?? '').trim();
        if (!['Pending', 'Present', 'Absent'].includes(attendance)) throw new Error('Choose Pending, Present, or Absent.');
        if (absentReason.length > 500) throw new Error('The absent reason must be under 500 characters.');
        const entry = await entries.findOne({ dcId, rollNoNormalized, removed: { $ne: true } }, { session });
        if (!entry) throw new Error('Student is not assigned to this DC.');
        if (entry.attendance === attendance && entry.absentReason === absentReason) return { updated: true };
        await entries.updateOne({ _id: entry._id }, { $set: { attendance, absentReason, attendanceUpdatedAt: now, attendanceUpdatedBy: actor.memberId } }, { session });
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
        if (dc.status !== 'Planned') throw new Error('Attendance can only be changed for a planned DC.');
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
        if (dc.status !== 'Planned') throw new Error('Recommendations can only be changed for a planned DC.');
        const rollNoNormalized = String(payload?.rollNo ?? '').trim().toUpperCase();
        const defaultId = String(payload?.defaultId ?? '').trim();
        const recommendation = String(payload?.recommendation ?? payload?.decision ?? '').toLowerCase();
        const recommendationValue = recommendation === 'remove' ? 'remove'
          : recommendation === 'keep' ? 'keep'
            : recommendation === 'defer' ? 'defer'
              : recommendation === 'discuss with team' || recommendation === 'discuss' ? 'discuss'
                : recommendation === 'undecided' ? 'undecided' : '';
        if (!recommendationValue) {
          throw new Error('Choose Undecided, To be removed, To be kept, Defer, or Discuss with team.');
        }
        const decision = recommendationLabel(recommendationValue);
        const reason = String(payload?.recommendationNotes ?? payload?.reason ?? '').trim();
        const until = String(payload?.until ?? '').trim();
        const discussionDate = String(payload?.discussionDate ?? '').trim();
        if (reason.length > 2000) throw new Error('Recommendation notes must be under 2,000 characters.');
        if ((recommendationValue === 'remove' || recommendationValue === 'defer') && !reason) throw new Error('Recommendation notes are required for removal and deferral.');
        if (recommendationValue === 'defer' && until !== 'next_dc' && !validDate(until)) throw new Error('Choose a valid deferral date or next DC.');
        const entry = await entries.findOne({ dcId, rollNoNormalized, removed: { $ne: true } }, { session });
        if (!entry) throw new Error('Student is not assigned to this DC.');
        const existing = (entry.decisions ?? []).find((item: Decision) => item.defaultId === defaultId);
        if (!existing) throw new Error('Default is not part of this DC entry.');
        const previousRecommendation = recommendationOf(existing);
        if (previousRecommendation === recommendationValue && (existing.recommendationNotes ?? existing.reason ?? '') === reason && existing.until === until) return { updated: true };
        const history = [...(existing.recommendationHistory ?? [])];
        if ((previousRecommendation !== recommendationValue || (existing.recommendationNotes ?? existing.reason ?? '') !== reason) &&
          (existing.recommendation || existing.recommendationAt)) {
          history.push({
            recommendation: previousRecommendation,
            notes: existing.recommendationNotes ?? existing.reason ?? '',
            by: String(existing.recommendationBy ?? ''),
            at: existing.recommendationAt ?? now,
          });
        }
        const nextDecision = {
          ...existing,
          decision,
          reason,
          recommendation: recommendationValue,
          recommendationNotes: reason,
          recommendationBy: actor.memberId,
          recommendationAt: now,
          recommendationHistory: history,
          ...(recommendationValue === 'defer' ? { until } : {}),
          ...(recommendationValue !== 'defer' ? { until: '' } : {}),
          ...(recommendationValue === 'discuss' ? { discussionDate } : {}),
          ...(recommendationValue !== 'discuss' ? { discussionDate: '' } : {}),
        };
        const nextDecisions = (entry.decisions as Decision[]).map((item) => item.defaultId === defaultId ? nextDecision : item);
        await entries.updateOne(
          { _id: entry._id },
          { $set: { decisions: nextDecisions, decisionsUpdatedAt: now, decisionsUpdatedBy: actor.memberId } },
          { session },
        );
        const record = await defaults.findOne({ id: defaultId }, { session });
        if (!record) throw new Error('Default is not part of this DC entry.');
        await addEvent(db, session, {
          defaultId, rollNo: record.rollNo, type: 'dc_recommendation', memberId: actor.memberId,
          enteredByEmail: email, at: now, reason, note: `Recommendation: ${decision}`,
          dcReference: dc.date, dcId,
          changes: { recommendation: { from: previousRecommendation, to: recommendationValue } },
        });
        await db.collection('audit_log').insertOne({
          action: 'DC recommendation updated', collection: 'dc_entries', by: email, at: now,
          before: { dcId, rollNo: entry.rollNo, defaultId, recommendation: previousRecommendation },
          after: { dcId, rollNo: entry.rollNo, defaultId, recommendation: recommendationValue, until, notes: reason, updatedBy: actor.memberId },
        }, { session });
        return { updated: true };
      }

      if (action === 'discussion_note') {
        if (dc.status !== 'Held') throw new Error('Discussion notes can only be added to a held DC.');
        const defaultId = String(payload?.defaultId ?? '').trim();
        const rollNoNormalized = String(payload?.rollNo ?? '').trim().toUpperCase();
        const text = String(payload?.note ?? '').trim();
        if (!text || text.length > 2000) throw new Error('Enter a discussion note under 2,000 characters.');
        const entry = await entries.findOne({ dcId, rollNoNormalized, attendance: 'Present', removed: { $ne: true } }, { session });
        const existing = (entry?.decisions ?? []).find((item: Decision) => item.defaultId === defaultId);
        if (!entry || !existing || recommendationOf(existing) === 'defer') throw new Error('This default is not in the held discussion scope.');
        const notes = [...(existing.discussionNotes ?? []), { text, by: actor.memberId, at: now }];
        await entries.updateOne({ _id: entry._id }, { $set: { 'decisions.$[item].discussionNotes': notes } }, {
          session,
          arrayFilters: [{ 'item.defaultId': defaultId }],
        });
        await db.collection('audit_log').insertOne({
          action: 'DC discussion note added', collection: 'dc_entries', by: email, at: now,
          after: { dcId, rollNo: entry.rollNo, defaultId, note: text },
        }, { session });
        return { updated: true };
      }

      if (action === 'final_decision') {
        if (dc.status !== 'Held') throw new Error('Final decisions can only be added to a held DC.');
        const defaultId = String(payload?.defaultId ?? '').trim();
        const decision = String(payload?.finalDecision ?? '');
        const remarks = String(payload?.finalRemarks ?? '').trim();
        const until = String(payload?.until ?? '').trim();
        if (!['remove', 'keep', 'defer', 'send_back'].includes(decision)) throw new Error('Choose Remove, Keep/lock, Defer, or Send back to DC.');
        if (remarks.length > 2000) throw new Error('Final remarks must be under 2,000 characters.');
        if (decision === 'defer' && !remarks) throw new Error('A reason is required for a final deferral.');
        if (decision === 'defer' && until !== 'next_dc' && !validDate(until)) throw new Error('Choose a valid deferral date or next DC.');
        const entry = await entries.findOne({ dcId, attendance: 'Present', removed: { $ne: true }, 'decisions.defaultId': defaultId }, { session });
        const item = (entry?.decisions ?? []).find((candidate: Decision) => candidate.defaultId === defaultId);
        if (!entry || !item || recommendationOf(item) === 'defer') throw new Error('This default is not in the final discussion scope.');
        const pendingDefault = await defaults.findOne({ id: defaultId, dcId, dcStatus: 'Awaiting team decision', status: 'logged' }, { session });
        if (!pendingDefault) throw new Error('This default is not awaiting a team decision.');
        const nextDecisions = (entry.decisions as Decision[]).map((candidate) => candidate.defaultId === defaultId
          ? { ...candidate, finalDecision: decision, finalRemarks: remarks, finalUntil: decision === 'defer' ? until : '', finalBy: actor.memberId, finalAt: now }
          : candidate);
        await entries.updateOne({ _id: entry._id }, { $set: { decisions: nextDecisions } }, { session });
        await db.collection('audit_log').insertOne({
          action: 'DC final decision recorded', collection: 'dc_entries', by: email, at: now,
          after: { dcId, rollNo: entry.rollNo, defaultId, finalDecision: decision, finalRemarks: remarks, until },
        }, { session });
        return { updated: true };
      }

      if (action === 'remarks') {
        if (dc.status !== 'Planned') throw new Error('Remarks can only be changed for a planned DC.');
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

      if (action === 'complete' || action === 'hold') {
        if (dc.status !== 'Planned') throw new Error('Only a planned DC can be held.');
        const dcEntries = await entries.find({ dcId, removed: { $ne: true } }, { session }).toArray();
        if (!dcEntries.length) throw new Error('Add students before closing the DC meeting.');
        if (dcEntries.some((entry) => entry.attendance === 'Pending')) throw new Error('Record attendance for every student before closing the DC meeting.');
        const allDefaultIds = dcEntries.flatMap((entry) => (entry.decisions ?? []).map((item: Decision) => item.defaultId));
        const currentDefaults = await defaults.find({ id: { $in: allDefaultIds }, dcId }, { session }).toArray();
        const defaultById = new Map(currentDefaults.map((record) => [record.id, record]));
        let deferredCount = 0;
        let awaitingCount = 0;
        for (const entry of dcEntries) {
          for (const item of (entry.decisions ?? []) as Decision[]) {
            const record = defaultById.get(item.defaultId);
            if (!record) throw new Error(`A default for ${entry.rollNo} is missing; the DC cannot be held.`);
            if (entry.attendance !== 'Present') {
              await defaults.updateOne({ id: record.id, dcId }, {
                $set: { dcStatus: 'Not scheduled', updatedAt: now, updatedByMemberId: actor.memberId },
                $unset: { dcId: '' },
              }, { session });
              continue;
            }
            const recommendation = recommendationOf(item);
            if (recommendation === 'defer') {
              const reason = item.recommendationNotes ?? item.reason ?? '';
              const until = item.until ?? '';
              if (!reason || (until !== 'next_dc' && !validDate(until))) throw new Error('Complete every deferral recommendation before holding the DC.');
              const history = [...(record.deferral?.history ?? []), { until, reason, by: actor.memberId, at: now }];
              await defaults.updateOne({ id: record.id, dcId }, {
                $set: {
                  status: 'logged',
                  dcStatus: 'Deferred',
                  dcReference: dc.date,
                  deferral: { until, reason, by: actor.memberId, at: now, history },
                  updatedAt: now,
                  updatedByMemberId: actor.memberId,
                },
              }, { session });
              deferredCount += 1;
            } else {
              await defaults.updateOne({ id: record.id, dcId }, {
                $set: { dcStatus: 'Awaiting team decision', updatedAt: now, updatedByMemberId: actor.memberId },
              }, { session });
              awaitingCount += 1;
            }
            await addEvent(db, session, {
              defaultId: record.id, rollNo: record.rollNo, type: 'dc_recommendation_held', memberId: actor.memberId,
              enteredByEmail: email, at: now, reason: item.recommendationNotes ?? item.reason ?? '',
              note: `Recommendation: ${recommendation}`,
              dcReference: dc.date, dcId,
              changes: { dcStatus: { from: record.dcStatus ?? 'Scheduled', to: recommendation === 'defer' ? 'Deferred' : 'Awaiting team decision' } },
            });
          }
        }
        await dcs.updateOne({ id: dcId }, {
          $set: { status: 'Held', heldBy: actor.memberId, heldAt: now, updatedAt: now },
        }, { session });
        await entries.updateMany({ dcId, removed: { $ne: true } }, {
          $set: { heldAt: now },
          $unset: { openStudentRollNo: '' },
        }, { session });
        await db.collection('audit_log').insertOne({
          action: 'DC meeting held', collection: 'dcs', by: email, at: now,
          before: { dcId, status: dc.status },
          after: { dcId, status: 'Held', awaitingCount, deferredCount },
        }, { session });
        const autoAdd = await syncCurrentDcAutoAdd(db, session, actor.memberId, email, now);
        return { held: true, awaitingCount, deferredCount, autoAdd };
      }

      if (action === 'finalise') {
        if (dc.status !== 'Held') throw new Error('Only a held DC can be finalised.');
        const dcEntries = await entries.find({ dcId, removed: { $ne: true } }, { session }).toArray();
        const scopedDefaults = dcEntries.flatMap((entry) => entry.attendance === 'Present'
          ? ((entry.decisions ?? []) as Decision[])
            .filter((item) => recommendationOf(item) !== 'defer')
            .map((item) => ({ entry, item }))
          : []);
        const missing = scopedDefaults.filter(({ item }) => !['remove', 'keep', 'defer', 'send_back'].includes(item.finalDecision ?? ''));
        if (missing.length) throw new Error(`Record a final decision for all ${missing.length} remaining defaults before finalising.`);
        const allDefaultIds = scopedDefaults.map(({ item }) => item.defaultId);
        const currentDefaults = await defaults.find({ id: { $in: allDefaultIds }, dcId, dcStatus: 'Awaiting team decision', status: 'logged' }, { session }).toArray();
        const defaultById = new Map(currentDefaults.map((record) => [record.id, record]));
        for (const { entry, item } of scopedDefaults) {
          const record = defaultById.get(item.defaultId);
          if (!record) throw new Error(`A default for ${entry.rollNo} is missing; the DC cannot be finalised.`);
          const decision = item.finalDecision;
          const remarks = item.finalRemarks || (item.discussionNotes ?? []).map((note) => note.text).join('\n') ||
            item.recommendationNotes || entry.remarks || 'Final decision recorded by the Disciplinary Committee.';
          const set: Record<string, unknown> = {
            dcReference: dc.date,
            updatedAt: now,
            updatedByMemberId: actor.memberId,
            dcStatus: decision === 'send_back' ? 'Not scheduled' : decision === 'defer' ? 'Deferred' : 'Decided',
          };
          const unset: Record<string, ''> = {};
          if (decision === 'remove') {
            set.status = 'removed';
            set.removedAt = now;
            set.removalReason = remarks;
          } else if (decision === 'keep') {
            set.status = 'locked';
            set.lockedAt = now;
          } else if (decision === 'defer') {
            const until = item.finalUntil ?? '';
            if (!remarks || (until !== 'next_dc' && !validDate(until))) throw new Error('Complete every final deferral with a reason and valid date before finalising.');
            const history = [...(record.deferral?.history ?? []), { until, reason: remarks, by: actor.memberId, at: now }];
            set.status = 'logged';
            set.deferral = { until, reason: remarks, by: actor.memberId, at: now, history };
          } else {
            set.status = 'logged';
            unset.dcId = '';
            unset.deferral = '';
          }
          await defaults.updateOne({ id: record.id }, {
            $set: set,
            ...(Object.keys(unset).length ? { $unset: unset } : {}),
          }, { session });
          await addEvent(db, session, {
            defaultId: record.id, rollNo: record.rollNo, type: 'dc_final_decision', memberId: actor.memberId,
            enteredByEmail: email, at: now, reason: remarks, note: decision,
            dcReference: dc.date, dcId,
            changes: {
              finalDecision: { from: '', to: decision },
              dcStatus: { from: 'Awaiting team decision', to: set.dcStatus },
              ...(decision === 'remove' ? { status: { from: record.status, to: 'removed' } } : {}),
              ...(decision === 'keep' ? { status: { from: record.status, to: 'locked' } } : {}),
            },
          });
        }
        await dcs.updateOne({ id: dcId, status: 'Held' }, {
          $set: { status: 'Finalised', finalisedBy: actor.memberId, finalisedAt: now, updatedAt: now },
        }, { session });
        await db.collection('audit_log').insertOne({
          action: 'DC finalised', collection: 'dcs', by: email, at: now,
          before: { dcId, status: 'Held' },
          after: { dcId, status: 'Finalised', decisions: scopedDefaults.length },
        }, { session });
        const autoAdd = await syncCurrentDcAutoAdd(db, session, actor.memberId, email, now);
        return { finalised: true, decisions: scopedDefaults.length, autoAdd };
      }

      throw new Error('Choose a valid DC action.');
    });
    return NextResponse.json(result);
  } catch (error) {
    const message = error instanceof Error ? error.message : '';
    if (message === 'DC_NOT_FOUND') return NextResponse.json({ message: 'DC not found.' }, { status: 404 });
    if (message === 'ACTOR_REQUIRED') return NextResponse.json({ message: 'Your account must be in the active member directory to manage DCs.' }, { status: 403 });
    if (message.startsWith('Only ') || message.startsWith('Students can only') || message.startsWith('A reason') ||
      message.startsWith('A deferral') || message.startsWith('Decision reason') || message.startsWith('Recommendation notes') ||
      message.startsWith('Attendance can only') || message.startsWith('Decisions can only') ||
      message.startsWith('A cancellation') || message.startsWith('Enter a valid') || message.startsWith('Choose ') ||
      message.startsWith('Venue ') || message.startsWith('Select between') || message.startsWith('Unknown student') ||
      message.includes('already assigned') || message.includes('no undecided defaults') ||
      message.startsWith('A student') || message.startsWith('Student is not assigned') ||
      message.startsWith('The absent reason') || message.startsWith('Default is not part') ||
      message.startsWith('Remarks can only') || message.startsWith('Remarks must') ||
      message.startsWith('Mark every') || message.startsWith('A default for') ||
      message.startsWith('Attendance can only') || message.startsWith('Only an active') ||
      message.startsWith('Decisions can only') || message.startsWith('Recommendations can only') || message.startsWith('Remarks can only') ||
      message.startsWith('Add students') || message.startsWith('Record attendance') ||
      message.startsWith('Complete every') || message.startsWith('This default') ||
      message.startsWith('Enter a discussion') || message.startsWith('Discussion notes') ||
      message.startsWith('Final remarks') || message.startsWith('Record a final') ||
      message.startsWith('A final deferral') || message.startsWith('A reason is required for a final')) {
      return NextResponse.json({ message }, { status: 400 });
    }
    if (error && typeof error === 'object' && 'code' in error && error.code === 11000) {
      console.error('Duplicate DC assignment during update.', error);
      return NextResponse.json({ message: 'This student is already assigned to an open DC.' }, { status: 409 });
    }
    console.error('Unable to update DC.', error);
    return NextResponse.json({ message: databaseFailureMessage(error, 'Unable to update the DC or sync automatic student assignments. Try again, and contact an administrator if it continues.') }, { status: 500 });
  }
}
