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
    const members = await db.collection('members').find(team ? { team } : {})
      .project({ _id: 0, memberId: 1, name: 1, team: 1 }).toArray();
    const memberById = new Map(members.map((member) => [member.memberId, member]));
    if (team) {
      const memberIds = members.map((member) => member.memberId);
      query.memberId = { $in: memberIds };
    }

    const events = await db.collection('default_events').find(query).sort({ at: -1 }).toArray();
    const rollNos = Array.from(new Set(events.map((event) => event.rollNo)));
    const students = await db.collection('students').find({ rollNo: { $in: rollNos } })
      .project({ _id: 0, rollNo: 1, name: 1 }).toArray();
    const nameByRollNo = new Map(students.map((student) => [student.rollNo.toUpperCase(), student.name]));
    return NextResponse.json(events.map((event) => {
      const changes = { ...(event.changes ?? {}) };
      const loggedByChange = changes.loggedByMemberId;
      if (loggedByChange) {
        const label = (id: unknown) => {
          const member = memberById.get(String(id));
          return member ? `${member.name} (${member.team})` : 'Imported';
        };
        changes.loggedByMemberId = { from: label(loggedByChange.from), to: label(loggedByChange.to) };
      }
      return {
        ...event,
        changes,
        memberName: memberById.get(event.memberId)?.name ?? 'Imported',
        team: memberById.get(event.memberId)?.team ?? 'Imported',
        studentName: nameByRollNo.get(String(event.rollNo).toUpperCase()) ?? event.rollNo,
      };
    }));
  } catch (error) {
    console.error('Unable to load activity feed from MongoDB.', error);
    return NextResponse.json({ message: getApiErrorMessage(error, 'Unable to load activity.') }, { status: 500 });
  }
}
