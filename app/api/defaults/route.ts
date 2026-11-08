import { randomUUID } from 'node:crypto';
import { NextResponse } from 'next/server';
import { getAuthorizedEmail } from '@/lib/authorization';
import { syncCurrentDcAutoAdd } from '@/lib/dc-auto-add';
import { findActiveMember, findActorMember } from '@/lib/members';
import { databaseFailureMessage, getDb, withMongoTransaction } from '@/lib/mongodb';

export const runtime = 'nodejs';
export const dynamic = 'force-dynamic';
export const maxDuration = 30;

function normalizeRollNo(value: unknown) {
  return String(value ?? '').trim().toUpperCase();
}

async function loadDefaults(db: Awaited<ReturnType<typeof getDb>>, includeRemoved: boolean) {
  const query = includeRemoved ? {} : { status: { $ne: 'removed' } };
  const [records, types, members, dcs] = await Promise.all([
    db.collection('defaults').find(query, { projection: { _id: 0, rollNoNormalized: 0 } }).sort({ date: -1, createdAt: -1 }).toArray(),
    db.collection('default_types').find({}, { projection: { _id: 0, id: 1, name: 1 } }).toArray(),
    db.collection('members').find({}, { projection: { _id: 0, memberId: 1, name: 1, team: 1 } }).toArray(),
    db.collection('dcs').find({}, { projection: { _id: 0, id: 1, date: 1 } }).toArray(),
  ]);
  const typeNames = new Map(types.map((type) => [type.id, type.name]));
  const memberById = new Map(members.map((member) => [member.memberId, member]));
  const dcDateById = new Map(dcs.map((dc) => [dc.id, dc.date]));
  return records.map((record) => {
    const loggedBy = memberById.get(record.loggedByMemberId);
    const lastUpdatedBy = memberById.get(record.updatedByMemberId);
    const discussion = record.discussion ? {
      ...record.discussion,
      raisedByName: memberById.get(record.discussion.raisedBy)?.name ?? 'Unknown member',
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
    return {
      ...record,
      ...(discussion ? { discussion } : {}),
      ...(deferral ? { deferral } : {}),
      typeName: typeNames.get(record.typeId) ?? record.typeId,
      dcDate: dcDateById.get(record.dcId) ?? '',
      loggedByName: loggedBy?.name ?? 'Imported',
      team: loggedBy?.team ?? 'Imported',
      updatedByName: lastUpdatedBy?.name ?? 'Imported',
    };
  });
}

function parseDate(value: unknown) {
  if (typeof value !== 'string' || !/^\d{4}-\d{2}-\d{2}$/.test(value)) return null;
  const date = new Date(`${value}T00:00:00.000Z`);
  return Number.isNaN(date.getTime()) || date.toISOString().slice(0, 10) !== value ? null : date;
}

export async function GET(request: Request) {
  try {
    const email = await getAuthorizedEmail();
    if (!email) return NextResponse.json({ message: 'Sign in with an authorized account.' }, { status: 401 });
    const includeRemoved = new URL(request.url).searchParams.get('includeRemoved') === 'true';
    return NextResponse.json(await loadDefaults(await getDb(), includeRemoved));
  } catch (error) {
    console.error('Unable to load defaults from MongoDB.', error);
    return NextResponse.json({ message: databaseFailureMessage(error, 'Unable to load defaults.') }, { status: 500 });
  }
}

export async function POST(request: Request) {
  try {
    const enteredByEmail = await getAuthorizedEmail();
    if (!enteredByEmail) return NextResponse.json({ message: 'Sign in with an authorized account.' }, { status: 401 });

    const payload = await request.json();
    const rollNo = String(payload?.rollNo ?? '').trim();
    const rollNoNormalized = normalizeRollNo(rollNo);
    const company = String(payload?.company ?? '').trim();
    const typeId = String(payload?.typeId ?? '').trim();
    const date = String(payload?.date ?? '').trim();
    const notes = String(payload?.notes ?? '').trim();
    const loggedByMemberId = String(payload?.loggedByMemberId ?? '').trim();
    const eventReason = String(payload?.reason ?? '').trim();

    if (!rollNoNormalized) return NextResponse.json({ message: 'Roll number is required.' }, { status: 400 });
    if (!company || !typeId || !date || !loggedByMemberId) {
      return NextResponse.json({ message: 'Company, type, date and Logged by member are required.' }, { status: 400 });
    }
    const parsedDate = parseDate(date);
    if (!parsedDate) return NextResponse.json({ message: 'Enter a valid date.' }, { status: 400 });

    const result = await withMongoTransaction(async (db, session) => {
      const [student, defaultType, loggedBy, actor] = await Promise.all([
        db.collection('students').findOne({ rollNoNormalized }, { session }),
        db.collection('default_types').findOne({ id: typeId, active: true }, { session }),
        findActiveMember(db, loggedByMemberId),
        findActorMember(db, enteredByEmail),
      ]);

      if (!student) throw new Error('Unknown roll number. Please use a valid student roll number.');
      if (!defaultType) throw new Error('Select an active default type.');
      if (!loggedBy) throw new Error('Choose an active placement or summer internship member.');
      if (!actor) throw new Error('Your account must be in the active member directory to log defaults.');

      const now = new Date();
      const record = {
        id: randomUUID(),
        rollNo: student.rollNo,
        rollNoNormalized,
        memberId: loggedBy.memberId,
        loggedByMemberId: loggedBy.memberId,
        enteredByEmail,
        typeId,
        company,
        date,
        status: 'logged',
        dcReference: '',
        dcId: null,
        dcStatus: 'Not scheduled',
        notes,
        createdAt: now,
        updatedAt: now,
        updatedByMemberId: actor.memberId,
      };
      const event = {
        eventId: randomUUID(),
        defaultId: record.id,
        rollNo: record.rollNo,
        type: 'logged',
        memberId: loggedBy.memberId,
        enteredByEmail,
        at: now,
        reason: eventReason,
        note: '',
        dcReference: '',
        changes: {},
      };

      await db.collection('defaults').insertOne(record, { session });
      await db.collection('default_events').insertOne(event, { session });
      const autoAdd = await syncCurrentDcAutoAdd(db, session, actor.memberId, enteredByEmail, now);
      return { created: record, autoAdd };
    });

    return NextResponse.json({
      created: result.created,
      autoAdd: result.autoAdd,
      defaults: await loadDefaults(await getDb(), false),
    });
  } catch (error) {
    const message = error instanceof Error ? error.message : '';
    if (
      message.startsWith('Unknown roll number') ||
      message.startsWith('Select an active') ||
      message.startsWith('Choose an active') ||
      message.startsWith('Your account')
    ) {
      return NextResponse.json({ message }, { status: 400 });
    }
    console.error('Unable to save default and event to MongoDB.', error);
    return NextResponse.json({ message: databaseFailureMessage(error, 'Unable to save the default or sync it to the current DC. Try again, and contact an administrator if it continues.') }, { status: 500 });
  }
}
