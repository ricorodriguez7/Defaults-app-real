import { randomUUID } from 'node:crypto';
import { NextResponse } from 'next/server';
import { getAuthorizedEmail } from '@/lib/authorization';
import { findActorMember } from '@/lib/members';
import { databaseFailureMessage, getDb, withMongoTransaction } from '@/lib/mongodb';

export const runtime = 'nodejs';
export const dynamic = 'force-dynamic';
export const maxDuration = 30;

type RouteContext = { params: { id: string } };

export async function GET(_request: Request, { params }: RouteContext) {
  try {
    const email = await getAuthorizedEmail();
    if (!email) return NextResponse.json({ message: 'Sign in with an authorized account.' }, { status: 401 });
    const db = await getDb();
    const events = await db.collection('default_events').find({ defaultId: params.id }).sort({ at: -1 }).toArray();
    const memberIds = Array.from(new Set(events.map((event) => event.memberId).filter(Boolean)));
    const members = await db.collection('members').find({ memberId: { $in: memberIds } })
      .project({ _id: 0, memberId: 1, name: 1, team: 1 }).toArray();
    const byId = new Map(members.map((member) => [member.memberId, member]));
    return NextResponse.json(events.map((event) => ({
      ...event,
      memberName: byId.get(event.memberId)?.name ?? 'Imported',
      team: byId.get(event.memberId)?.team ?? 'Imported',
    })));
  } catch (error) {
    console.error('Unable to load default events from MongoDB.', error);
    return NextResponse.json({ message: databaseFailureMessage(error, 'Unable to load the default timeline.') }, { status: 500 });
  }
}

export async function POST(request: Request, { params }: RouteContext) {
  try {
    const enteredByEmail = await getAuthorizedEmail();
    if (!enteredByEmail) return NextResponse.json({ message: 'Sign in with an authorized account.' }, { status: 401 });
    const payload = await request.json();
    const note = String(payload?.note ?? '').trim();
    if (!note) return NextResponse.json({ message: 'Enter a DC note before saving.' }, { status: 400 });

    await withMongoTransaction(async (db, session) => {
      const record = await db.collection('defaults').findOne({ id: params.id }, { session });
      if (!record) throw new Error('NOT_FOUND');
      const member = await findActorMember(db, enteredByEmail);
      if (!member) throw new Error('Only active placement or summer internship members can add DC notes.');
      const now = new Date();
      await db.collection('default_events').insertOne({
        eventId: randomUUID(),
        defaultId: record.id,
        rollNo: record.rollNo,
        type: 'dc_note',
        memberId: member.memberId,
        enteredByEmail,
        at: now,
        reason: '',
        note,
        dcReference: record.dcReference || '',
        changes: {},
      }, { session });
    });
    return NextResponse.json({ saved: true });
  } catch (error) {
    if (error instanceof Error && error.message === 'NOT_FOUND') {
      return NextResponse.json({ message: 'Default not found.' }, { status: 404 });
    }
    if (error instanceof Error && error.message.startsWith('Only active')) {
      return NextResponse.json({ message: error.message }, { status: 403 });
    }
    console.error('Unable to append DC note.', error);
    return NextResponse.json({ message: databaseFailureMessage(error, 'Unable to save the DC note.') }, { status: 500 });
  }
}
