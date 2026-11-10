import { NextResponse } from 'next/server';
import { getAuthorizedEmail } from '@/lib/authorization';
import { getApiErrorMessage } from '@/lib/database-errors';
import { getDb } from '@/lib/mongodb';

export const runtime = 'nodejs';
export const dynamic = 'force-dynamic';
export const maxDuration = 30;

export async function GET(request: Request) {
  try {
    const email = await getAuthorizedEmail();
    if (!email) return NextResponse.json({ message: 'Sign in with an authorized account.' }, { status: 401 });

    const params = new URL(request.url).searchParams;
    const query: Record<string, unknown> = {};
    const memberId = params.get('memberId');
    const eventType = params.get('type');
    const rollNo = params.get('rollNo')?.trim();
    const team = params.get('team');
    const from = params.get('from');
    const to = params.get('to');
    const category = params.get('category') || 'important';
    const batch = params.get('batch') || '';
    const skip = Math.max(0, Number.parseInt(params.get('skip') || '0', 10) || 0);
    const requestedLimit = Number.parseInt(params.get('limit') || '25', 10);
    const limit = Math.min(10000, Math.max(1, Number.isFinite(requestedLimit) ? requestedLimit : 25));

    if ((from && !/^\d{4}-\d{2}-\d{2}$/.test(from)) || (to && !/^\d{4}-\d{2}-\d{2}$/.test(to))) {
      return NextResponse.json({ message: 'Date filters must use YYYY-MM-DD.' }, { status: 400 });
    }
    if (memberId) query.memberId = memberId;
    if (eventType) query.type = eventType;
    if (rollNo) query.rollNo = { $regex: `^${rollNo.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}$`, $options: 'i' };
    if (from || to) {
      const at: Record<string, Date> = {};
      if (from) {
        const start = new Date(`${from}T00:00:00.000Z`);
        if (Number.isNaN(start.getTime()) || start.toISOString().slice(0, 10) !== from) {
          return NextResponse.json({ message: 'Enter a valid start date.' }, { status: 400 });
        }
        at.$gte = start;
      }
      if (to) {
        const inclusiveEnd = new Date(`${to}T00:00:00.000Z`);
        if (Number.isNaN(inclusiveEnd.getTime()) || inclusiveEnd.toISOString().slice(0, 10) !== to) {
          return NextResponse.json({ message: 'Enter a valid end date.' }, { status: 400 });
        }
        inclusiveEnd.setUTCDate(inclusiveEnd.getUTCDate() + 1);
        at.$lt = inclusiveEnd;
      }
      query.at = at;
    }

    const db = await getDb();
    const additionalFilters: Record<string, unknown>[] = [];
    if (category !== 'all') {
      const categories: Record<string, string[]> = {
        important: ['logged', 'edited', 'locked', 'unlocked', 'removed', 'restored', 'discuss_flagged', 'discuss_note', 'deferral_due'],
        dc: ['dc_note', 'dc_scheduled', 'dc_unscheduled', 'dc_attendance', 'dc_decision', 'dc_rescheduled', 'dc_cancelled', 'dc_auto_added'],
        members: ['student_locked', 'student_unlocked', 'student_batch_updated'],
        imports: [],
      };
      if (!categories[category]) return NextResponse.json({ message: 'Choose a valid activity category.' }, { status: 400 });
      if (category === 'imports') {
        const imported = await db.collection('defaults').find({ source: 'excel_import' }).project({ _id: 0, id: 1 }).toArray();
        query.defaultId = { $in: imported.map((record) => record.id) };
      } else additionalFilters.push({ type: { $in: categories[category] } });
      if (category === 'important') {
        const imported = await db.collection('defaults').find({ source: 'excel_import' }).project({ _id: 0, id: 1 }).toArray();
        additionalFilters.push({ defaultId: { $nin: imported.map((record) => record.id) } });
      }
    }
    if (batch && !['junior', 'senior', 'not_set'].includes(batch)) {
      return NextResponse.json({ message: 'Choose a valid student batch.' }, { status: 400 });
    }
    const studentsForBatch = batch
      ? await db.collection('students').find(
        batch === 'not_set' ? { $or: [{ batch: { $exists: false } }, { batch: '' }] } : { batch },
        { projection: { _id: 0, rollNo: 1 } },
      ).toArray()
      : [];
    if (batch) {
      const batchRollNos = studentsForBatch.map((student) => new RegExp(`^${String(student.rollNo).replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}$`, 'i'));
      if (query.rollNo) additionalFilters.push({ rollNo: query.rollNo }, { rollNo: { $in: batchRollNos } });
      else query.rollNo = { $in: batchRollNos };
    }
    const members = await db.collection('members').find(team ? { team } : {})
      .project({ _id: 0, memberId: 1, name: 1, team: 1 }).toArray();
    const memberById = new Map(members.map((member) => [member.memberId, member]));
    if (team) {
      const memberIds = members.map((member) => member.memberId);
      if (query.memberId) additionalFilters.push({ memberId: query.memberId }, { memberId: { $in: memberIds } });
      else query.memberId = { $in: memberIds };
    }
    if (additionalFilters.length) query.$and = [...(Array.isArray(query.$and) ? query.$and as Record<string, unknown>[] : []), ...additionalFilters];

    const collection = db.collection('default_events');
    const total = await collection.countDocuments(query);
    const events = await collection.find(query).sort({ at: -1, eventId: -1 }).skip(skip).limit(limit).toArray();
    const rollNos = Array.from(new Set(events.map((event) => event.rollNo)));
    const defaultIds = Array.from(new Set(events.map((event) => event.defaultId).filter(Boolean)));
    const importedDefaults = defaultIds.length
      ? await db.collection('defaults').find({ id: { $in: defaultIds }, source: 'excel_import' })
        .project({ _id: 0, id: 1 }).toArray()
      : [];
    const importedDefaultIds = new Set(importedDefaults.map((record) => record.id));
    const students = await db.collection('students').find({ rollNo: { $in: rollNos } })
      .project({ _id: 0, rollNo: 1, name: 1, batch: 1 }).toArray();
    const studentByRollNo = new Map(students.map((student) => [student.rollNo.toUpperCase(), student]));
    const enrichedEvents = events.map((event) => {
      const changes = { ...(event.changes ?? {}) };
      const loggedByChange = changes.loggedByMemberId;
      if (loggedByChange) {
        const label = (id: unknown) => {
          const member = memberById.get(String(id));
          return member ? `${member.name} (${member.team})` : 'Imported';
        };
        changes.loggedByMemberId = { from: label(loggedByChange.from), to: label(loggedByChange.to) };
      }
      const eventCategory = event.importBatchId || importedDefaultIds.has(event.defaultId) ? 'imports' : String(event.type).startsWith('dc_') ? 'dc'
        : String(event.type).startsWith('student_') ? 'members' : 'important';
      return {
        ...event,
        changes,
        memberName: memberById.get(event.memberId)?.name ?? 'Imported',
        team: memberById.get(event.memberId)?.team ?? 'Imported',
        studentName: studentByRollNo.get(String(event.rollNo).toUpperCase())?.name ?? event.rollNo,
        studentBatch: studentByRollNo.get(String(event.rollNo).toUpperCase())?.batch,
        category: eventCategory,
        actor: event.type === 'dc_auto_added' || event.type === 'deferral_due' ? 'system' : 'member',
        actorLabel: event.type === 'dc_auto_added' || event.type === 'deferral_due' ? 'System' : memberById.get(event.memberId)?.name ?? 'Imported',
      };
    });
    return NextResponse.json({ events: enrichedEvents, hasMore: skip + events.length < total, nextSkip: skip + events.length, total });
  } catch (error) {
    console.error('Unable to load activity feed from MongoDB.', error);
    return NextResponse.json({ message: getApiErrorMessage(error, 'Unable to load activity.') }, { status: 500 });
  }
}
