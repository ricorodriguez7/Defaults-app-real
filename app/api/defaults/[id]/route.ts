import { randomUUID } from 'node:crypto';
import { NextResponse } from 'next/server';
import { getAuthorizedEmail } from '@/lib/authorization';
import { getApiErrorMessage } from '@/lib/database-errors';
import { syncCurrentDcAutoAdd } from '@/lib/dc-auto-add';
import { findActiveMember, findActorMember } from '@/lib/members';
import { getDb, withMongoTransaction } from '@/lib/mongodb';

export const runtime = 'nodejs';
export const dynamic = 'force-dynamic';

type RouteContext = { params: { id: string } };
type ChangeValue = { from: unknown; to: unknown };

function validDate(value: unknown) {
  if (typeof value !== 'string' || !/^\d{4}-\d{2}-\d{2}$/.test(value)) return false;
  const date = new Date(`${value}T00:00:00.000Z`);
  return !Number.isNaN(date.getTime()) && date.toISOString().slice(0, 10) === value;
}

async function decorateEvents(db: Awaited<ReturnType<typeof getDb>>, events: Record<string, any>[]) {
  const ids = Array.from(new Set(events.map((event) => event.memberId).filter(Boolean)));
  const members = await db.collection('members').find({ memberId: { $in: ids } }).project({
    _id: 0, memberId: 1, name: 1, team: 1,
  }).toArray();
  const byId = new Map(members.map((member) => [member.memberId, member]));
  return events.map((event) => {
    const changes = { ...(event.changes ?? {}) };
    const loggedByChange = changes.loggedByMemberId;
    if (loggedByChange) {
      const label = (id: unknown) => {
        const member = byId.get(String(id));
        return member ? `${member.name} (${member.team})` : 'Imported';
      };
      changes.loggedByMemberId = {
        from: label(loggedByChange.from),
        to: label(loggedByChange.to),
      };
    }
    return {
      ...event,
      changes,
      memberName: byId.get(event.memberId)?.name ?? 'Imported',
      team: byId.get(event.memberId)?.team ?? 'Imported',
    };
  });
}

export async function GET(_request: Request, { params }: RouteContext) {
  try {
    const email = await getAuthorizedEmail();
    if (!email) return NextResponse.json({ message: 'Sign in with an authorized account.' }, { status: 401 });
    const db = await getDb();
    const [record, events, members, defaultType] = await Promise.all([
      db.collection('defaults').findOne({ id: params.id }, { projection: { _id: 0, rollNoNormalized: 0 } }),
      db.collection('default_events').find({ defaultId: params.id }).sort({ at: -1 }).toArray(),
      db.collection('members').find({}).project({ _id: 0, memberId: 1, name: 1, team: 1 }).toArray(),
      db.collection('default_types').findOne({ id: (await db.collection('defaults').findOne({ id: params.id }))?.typeId }),
    ]);
    if (!record) return NextResponse.json({ message: 'Default not found.' }, { status: 404 });

    const memberById = new Map(members.map((member) => [member.memberId, member]));
    const loggedBy = memberById.get(record.loggedByMemberId);
    const discussion = record.discussion ? {
      ...record.discussion,
      raisedByName: memberById.get(record.discussion.raisedBy)?.name ?? 'Unknown member',
      resolvedByName: memberById.get(record.discussion.resolvedBy)?.name ?? '',
      notes: (record.discussion.notes ?? []).map((note: { by: string; at: Date; text: string }) => ({
        ...note,
        byName: memberById.get(note.by)?.name ?? 'Unknown member',
      })),
    } : undefined;
    const deferral = record.deferral ? {
      ...record.deferral,
      byName: memberById.get(record.deferral.by)?.name ?? 'Unknown member',
      history: (record.deferral.history ?? []).map((item: { by: string; until: string; reason: string; at: Date }) => ({
        ...item,
        byName: memberById.get(item.by)?.name ?? 'Unknown member',
      })),
    } : undefined;
    return NextResponse.json({
      record: {
        ...record,
        ...(discussion ? { discussion } : {}),
        ...(deferral ? { deferral } : {}),
        typeName: defaultType?.name ?? record.typeId,
        loggedByName: loggedBy?.name ?? 'Imported',
        team: loggedBy?.team ?? 'Imported',
      },
      events: await decorateEvents(db, events),
      members,
    });
  } catch (error) {
    console.error('Unable to load default history from MongoDB.', error);
    return NextResponse.json({ message: getApiErrorMessage(error, 'Unable to load default history.') }, { status: 500 });
  }
}

export async function PATCH(request: Request, { params }: RouteContext) {
  try {
    const enteredByEmail = await getAuthorizedEmail();
    if (!enteredByEmail) return NextResponse.json({ message: 'Sign in with an authorized account.' }, { status: 401 });
    const payload = await request.json();
    const action = String(payload?.action ?? '');
    const reason = String(payload?.reason ?? '').trim();
    const dcActions = ['flag_discussion', 'add_discussion_note', 'resolve_discussion', 'defer', 'extend_deferral'];
    if (!['edit', 'lock', 'unlock', 'restore', ...dcActions].includes(action)) {
      return NextResponse.json({ message: 'Choose a valid default action.' }, { status: 400 });
    }
    const requiredReasonActions = ['edit', 'lock', 'unlock', 'restore', 'flag_discussion', 'resolve_discussion', 'defer', 'extend_deferral'];
    if (requiredReasonActions.includes(action) && !reason) {
      return NextResponse.json({ message: 'A reason or note is required for this action.' }, { status: 400 });
    }
    if (reason.length > 2000) return NextResponse.json({ message: 'The reason must be under 2,000 characters.' }, { status: 400 });

    const db = await getDb();
    const result = await withMongoTransaction(async (transactionDb, session) => {
      const defaults = transactionDb.collection('defaults');
      const record = await defaults.findOne({ id: params.id }, { session });
      if (!record) throw new Error('NOT_FOUND');
      const actor = await findActorMember(transactionDb, enteredByEmail);
      if (!actor) throw new Error('Your account must be in the active member directory to change defaults.');

      const now = new Date();
      const eventCollection = transactionDb.collection('default_events');
      const audit = transactionDb.collection('audit_log');
      const syncAutoAdd = () => syncCurrentDcAutoAdd(transactionDb, session, actor.memberId, enteredByEmail, now);
      const discussionNote = String(payload?.note ?? '').trim();
      const until = String(payload?.until ?? '').trim();
      const discussionDate = String(payload?.discussionDate ?? '').trim();
      const validUntil = (value: string) => value === 'next_dc' || validDate(value);

      const unscheduleFromPlannedDc = async (why: string) => {
        if (record.dcStatus !== 'Scheduled' || !record.dcId) return;
        const dc = await transactionDb.collection('dcs').findOne({ id: record.dcId, status: 'Planned' }, { session });
        if (!dc) return;
        const entry = await transactionDb.collection('dc_entries').findOne({
          dcId: dc.id, rollNoNormalized: record.rollNoNormalized, removed: { $ne: true },
        }, { session });
        if (entry) {
          const decisions = (entry.decisions ?? []).filter((item: { defaultId: string }) => item.defaultId !== record.id);
          if (decisions.length) {
            await transactionDb.collection('dc_entries').updateOne({ _id: entry._id }, { $set: { decisions } }, { session });
          } else {
            await transactionDb.collection('dc_entries').updateOne(
              { _id: entry._id },
              { $set: { removed: true, removalReason: why, removedAt: now }, $unset: { openStudentRollNo: '' } },
              { session },
            );
          }
        }
        await eventCollection.insertOne({
          eventId: randomUUID(), defaultId: record.id, rollNo: record.rollNo, type: 'dc_unscheduled',
          memberId: actor.memberId, enteredByEmail, at: now, reason: why, note: '',
          dcReference: dc.date, dcId: dc.id,
          changes: { dcStatus: { from: 'Scheduled', to: action === 'flag_discussion' ? 'To discuss' : 'Deferred' } },
        }, { session });
      };

      if (dcActions.includes(action)) {
        if (record.status === 'locked' || record.status === 'removed') {
          throw new Error('Locked or removed defaults cannot be flagged or deferred.');
        }
        if (action === 'flag_discussion') {
          if (record.dcStatus === 'To discuss') throw new Error('This default is already flagged for team discussion.');
          if (record.dcStatus === 'Deferred') throw new Error('Resolve the deferral before flagging this default for discussion.');
          if (discussionDate && !validDate(discussionDate)) throw new Error('Enter a valid discussion date.');
          await unscheduleFromPlannedDc(reason);
          const priorDiscussion = record.discussion;
          await defaults.updateOne({ id: record.id }, {
            $set: {
              dcStatus: 'To discuss',
              discussion: {
                raisedBy: actor.memberId,
                raisedAt: now,
                reason,
                discussionDate: discussionDate || '',
                notes: [],
              },
              ...(priorDiscussion ? { discussionHistory: [...(record.discussionHistory ?? []), priorDiscussion] } : {}),
              updatedAt: now,
              updatedByMemberId: actor.memberId,
            },
            $unset: { dcId: '' },
          }, { session });
          await eventCollection.insertOne({
            eventId: randomUUID(), defaultId: record.id, rollNo: record.rollNo, type: 'discuss_flagged',
            memberId: actor.memberId, enteredByEmail, at: now, reason, note: '',
            dcReference: discussionDate, changes: { dcStatus: { from: record.dcStatus ?? 'Not scheduled', to: 'To discuss' } },
          }, { session });
          await audit.insertOne({
            action: 'Default flagged for team discussion', collection: 'defaults', by: enteredByEmail, at: now,
            before: { defaultId: record.id, dcStatus: record.dcStatus ?? 'Not scheduled', dcId: record.dcId },
            after: { defaultId: record.id, dcStatus: 'To discuss', reason, discussionDate, memberId: actor.memberId },
          }, { session });
          return { updated: true, autoAdd: await syncAutoAdd() };
        }

        if (action === 'add_discussion_note') {
          if (record.dcStatus !== 'To discuss') throw new Error('Only defaults flagged for discussion can receive discussion notes.');
          if (!discussionNote) throw new Error('Enter a discussion note.');
          if (discussionNote.length > 2000) throw new Error('Discussion notes must be under 2,000 characters.');
          const notes = [...(record.discussion?.notes ?? []), { by: actor.memberId, at: now, text: discussionNote }];
          await defaults.updateOne({ id: record.id }, { $set: { 'discussion.notes': notes, updatedAt: now, updatedByMemberId: actor.memberId } }, { session });
          await eventCollection.insertOne({
            eventId: randomUUID(), defaultId: record.id, rollNo: record.rollNo, type: 'discuss_note',
            memberId: actor.memberId, enteredByEmail, at: now, reason: '', note: discussionNote,
            dcReference: record.discussion?.discussionDate ?? '', changes: {},
          }, { session });
          await audit.insertOne({
            action: 'Team discussion note added', collection: 'defaults', by: enteredByEmail, at: now,
            after: { defaultId: record.id, note: discussionNote, memberId: actor.memberId },
          }, { session });
          return { updated: true, autoAdd: await syncAutoAdd() };
        }

        if (action === 'resolve_discussion') {
          if (record.dcStatus !== 'To discuss') throw new Error('This default is not waiting for team discussion.');
          const outcome = String(payload?.outcome ?? '');
          if (!['Take to DC', 'Remove', 'Defer'].includes(outcome)) throw new Error('Choose Take to DC, Remove, or Defer.');
          if (outcome === 'Defer' && !validUntil(until)) throw new Error('Choose a valid deferral date or next DC.');
          const update: Record<string, unknown> = {
            'discussion.resolvedBy': actor.memberId,
            'discussion.resolvedAt': now,
            'discussion.outcome': outcome,
            updatedAt: now,
            updatedByMemberId: actor.memberId,
          };
          if (outcome === 'Take to DC') {
            update.dcStatus = 'Not scheduled';
            await defaults.updateOne({ id: record.id }, { $set: update, $unset: { dcId: '' } }, { session });
          } else if (outcome === 'Remove') {
            update.dcStatus = 'Decided';
            update.status = 'removed';
            update.removedAt = now;
            update.dcReference = record.discussion?.discussionDate || '';
            await defaults.updateOne({ id: record.id }, { $set: update, $unset: { dcId: '' } }, { session });
          } else {
            const history = [...(record.deferral?.history ?? []), {
              until, reason, by: actor.memberId, at: now,
            }];
            update.dcStatus = 'Deferred';
            update.deferral = { until, reason, by: actor.memberId, at: now, history };
            await defaults.updateOne({ id: record.id }, { $set: update, $unset: { dcId: '' } }, { session });
          }
          await eventCollection.insertOne({
            eventId: randomUUID(), defaultId: record.id, rollNo: record.rollNo, type: outcome === 'Defer' ? 'deferred' : 'discuss_resolved',
            memberId: actor.memberId, enteredByEmail, at: now, reason, note: outcome,
            dcReference: until || record.discussion?.discussionDate || '',
            changes: { dcStatus: { from: 'To discuss', to: outcome === 'Take to DC' ? 'Not scheduled' : outcome === 'Remove' ? 'Decided' : 'Deferred' } },
          }, { session });
          await audit.insertOne({
            action: outcome === 'Defer' ? 'Default deferred after team discussion' : 'Team discussion resolved',
            collection: 'defaults', by: enteredByEmail, at: now,
            before: { defaultId: record.id, dcStatus: 'To discuss' },
            after: { defaultId: record.id, outcome, until: outcome === 'Defer' ? until : '', reason, memberId: actor.memberId },
          }, { session });
          return { updated: true, autoAdd: await syncAutoAdd() };
        }

        if (action === 'defer' || action === 'extend_deferral') {
          if (action === 'defer' && record.dcStatus === 'Deferred') throw new Error('Use extend deferral for a default that is already deferred.');
          if (action === 'extend_deferral' && record.dcStatus !== 'Deferred') throw new Error('Only deferred defaults can have their deferral extended.');
          if (record.dcStatus === 'To discuss') throw new Error('Resolve the team discussion before deferring this default.');
          if (!validUntil(until)) throw new Error('Choose a valid deferral date or next DC.');
          await unscheduleFromPlannedDc(reason);
          const history = [...(record.deferral?.history ?? []), { until, reason, by: actor.memberId, at: now }];
          await defaults.updateOne({ id: record.id }, {
            $set: {
              dcStatus: 'Deferred',
              deferral: { until, reason, by: actor.memberId, at: now, history },
              updatedAt: now,
              updatedByMemberId: actor.memberId,
            },
            $unset: { dcId: '' },
          }, { session });
          await eventCollection.insertOne({
            eventId: randomUUID(), defaultId: record.id, rollNo: record.rollNo,
            type: action === 'extend_deferral' ? 'deferral_extended' : 'deferred',
            memberId: actor.memberId, enteredByEmail, at: now, reason,
            note: until === 'next_dc' ? 'Until next DC' : `Until ${until}`,
            dcReference: until, changes: { dcStatus: { from: record.dcStatus ?? 'Not scheduled', to: 'Deferred' } },
          }, { session });
          await audit.insertOne({
            action: action === 'extend_deferral' ? 'Default deferral extended' : 'Default deferred',
            collection: 'defaults', by: enteredByEmail, at: now,
            before: { defaultId: record.id, dcStatus: record.dcStatus, deferral: record.deferral ?? null },
            after: { defaultId: record.id, dcStatus: 'Deferred', until, reason, memberId: actor.memberId },
          }, { session });
          return { updated: true, autoAdd: await syncAutoAdd() };
        }
      }

      if (record.dcStatus === 'Scheduled') throw new Error('Remove this default from its planned DC before changing it.');
      const changes: Record<string, ChangeValue> = {};
      let nextStatus = record.status;
      const update: Record<string, unknown> = { updatedAt: now, updatedByMemberId: actor.memberId };

      if (action === 'lock') {
        if (record.status !== 'logged') throw new Error('Only a logged default can be locked.');
        nextStatus = 'locked';
        update.status = nextStatus;
        update.lockedAt = now;
      } else if (action === 'unlock') {
        if (record.status !== 'locked') throw new Error('Only a locked default can be unlocked.');
        nextStatus = 'logged';
        update.status = nextStatus;
        update.unlockedAt = now;
      } else if (action === 'restore') {
        if (record.status !== 'removed') throw new Error('Only a removed default can be restored.');
        nextStatus = 'logged';
        update.status = nextStatus;
        update.restoredAt = now;
      } else {
        if (record.status === 'locked') throw new Error('Unlock this default before editing it.');
        if (record.status === 'removed') throw new Error('Restore this default before editing it.');

        const editable: Record<string, unknown> = {};
        for (const field of ['company', 'notes', 'dcReference'] as const) {
          if (payload?.[field] !== undefined) editable[field] = String(payload[field]).trim();
        }
        if (payload?.date !== undefined) {
          if (!validDate(payload.date)) throw new Error('Enter a valid date.');
          editable.date = payload.date;
        }
        if (payload?.typeId !== undefined) {
          const typeId = String(payload.typeId).trim();
          const type = await transactionDb.collection('default_types').findOne({ id: typeId, active: true }, { session });
          if (!type) throw new Error('Select an active default type.');
          editable.typeId = typeId;
        }
        if (payload?.loggedByMemberId !== undefined) {
          const memberId = String(payload.loggedByMemberId).trim();
          if (!await findActiveMember(transactionDb, memberId)) throw new Error('Choose an active placement or summer internship member.');
          editable.loggedByMemberId = memberId;
          editable.memberId = memberId;
        }
        for (const [field, value] of Object.entries(editable)) {
          if (field === 'memberId') {
            if (record.memberId !== value) update.memberId = value;
            continue;
          }
          if (record[field] !== value) {
            changes[field] = { from: record[field] ?? '', to: value };
            update[field] = value;
          }
        }
        if (!Object.keys(changes).length) throw new Error('No default fields changed.');
      }

      await defaults.updateOne({ id: params.id }, { $set: update }, { session });
      await transactionDb.collection('default_events').insertOne({
        eventId: randomUUID(),
        defaultId: record.id,
        rollNo: record.rollNo,
        type: action === 'edit' ? 'edited' : action === 'restore' ? 'restored' : action === 'lock' ? 'locked' : 'unlocked',
        memberId: actor.memberId,
        enteredByEmail,
        at: now,
        reason,
        note: '',
        dcReference: record.dcReference || '',
        changes,
      }, { session });
      return { updated: true, autoAdd: await syncAutoAdd() };
    });
    const updated = await db.collection('defaults').findOne({ id: params.id }, { projection: { _id: 0, rollNoNormalized: 0 } });
    return NextResponse.json({ updated, autoAdd: result.autoAdd });
  } catch (error) {
    if (error instanceof Error && error.message === 'NOT_FOUND') {
      return NextResponse.json({ message: 'Default not found.' }, { status: 404 });
    }
    if (
      error instanceof Error &&
      (error.message.startsWith('Only ') || error.message.startsWith('This default') ||
        error.message.startsWith('Locked or removed') || error.message.startsWith('Resolve the') ||
        error.message.startsWith('Use extend') || error.message.startsWith('Enter a discussion') ||
        error.message.startsWith('Choose Take') || error.message.startsWith('Choose a valid') ||
        error.message.startsWith('Discussion notes') || error.message.startsWith('A deferral') ||
        error.message.startsWith('Unlock ') || error.message.startsWith('Restore ') ||
        error.message.startsWith('A reason') || error.message.startsWith('Enter a valid') ||
        error.message.startsWith('Select an active') || error.message.startsWith('Choose an active') ||
          error.message.startsWith('Your account') || error.message.startsWith('No default') ||
          error.message.startsWith('Remove this default'))
    ) {
      return NextResponse.json({ message: error.message }, { status: 400 });
    }
    console.error('Unable to update default and append history.', error);
    return NextResponse.json({ message: getApiErrorMessage(error, 'Unable to update the default or sync automatic DC assignments. Try again, and contact an administrator if it continues.') }, { status: 500 });
  }
}

export async function DELETE(request: Request, { params }: RouteContext) {
  try {
    const enteredByEmail = await getAuthorizedEmail();
    if (!enteredByEmail) return NextResponse.json({ message: 'Sign in with an authorized account.' }, { status: 401 });
    const payload = await request.json();
    const reason = String(payload?.reason ?? '').trim();
    const dcReference = String(payload?.dcReference ?? '').trim();
    if (!reason || !dcReference) {
      return NextResponse.json({ message: 'A removal reason and DC reference are required.' }, { status: 400 });
    }

    const result = await withMongoTransaction(async (db, session) => {
      const record = await db.collection('defaults').findOne({ id: params.id }, { session });
      if (!record) throw new Error('NOT_FOUND');
      if (record.dcStatus === 'Scheduled') throw new Error('Remove this default from its planned DC before removing it.');
      if (record.status === 'locked') throw new Error('Unlock this default before removing it.');
      if (record.status === 'removed') throw new Error('This default is already removed.');
      const actor = await findActorMember(db, enteredByEmail);
      if (!actor) throw new Error('Your account must be in the active member directory to remove defaults.');
      const now = new Date();
      await db.collection('defaults').updateOne(
        { id: params.id },
        { $set: { status: 'removed', dcReference, removedAt: now, updatedAt: now, updatedByMemberId: actor.memberId } },
        { session },
      );
      await db.collection('default_events').insertOne({
        eventId: randomUUID(),
        defaultId: record.id,
        rollNo: record.rollNo,
        type: 'removed',
        memberId: actor.memberId,
        enteredByEmail,
        at: now,
        reason,
        note: '',
        dcReference,
        changes: { status: { from: record.status, to: 'removed' } },
      }, { session });
      return {
        autoAdd: await syncCurrentDcAutoAdd(db, session, actor.memberId, enteredByEmail, now),
      };
    });
    return NextResponse.json({ removed: true, softDeleted: true, autoAdd: result.autoAdd });
  } catch (error) {
    if (error instanceof Error && error.message === 'NOT_FOUND') {
      return NextResponse.json({ message: 'Default not found.' }, { status: 404 });
    }
    if (error instanceof Error && (error.message.startsWith('Unlock ') || error.message.startsWith('This default') ||
      error.message.startsWith('Your account') || error.message.startsWith('Remove this default'))) {
      return NextResponse.json({ message: error.message }, { status: 400 });
    }
    console.error('Unable to remove default and append history.', error);
    return NextResponse.json({ message: getApiErrorMessage(error, 'Unable to remove the default or sync automatic DC assignments. Try again, and contact an administrator if it continues.') }, { status: 500 });
  }
}
