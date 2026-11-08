import { randomUUID } from 'node:crypto';
import { NextResponse } from 'next/server';
import { getAuthorizedEmail } from '@/lib/authorization';
import { getApiErrorMessage } from '@/lib/database-errors';
import { isPlacementTeam } from '@/lib/members';
import { getDb, withMongoTransaction } from '@/lib/mongodb';

export const runtime = 'nodejs';
export const dynamic = 'force-dynamic';

function normalizeText(value: string) {
  return value.trim().replace(/\s+/g, ' ');
}

function normalizeTeam(value: string) {
  return value.trim().toLowerCase().replace(/[\s-]+/g, '_');
}

function validEmail(value: string) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(value);
}

function safeError(error: unknown) {
  return error && typeof error === 'object' && 'code' in error && error.code === 11000;
}

export async function GET(request: Request) {
  try {
    const enteredByEmail = await getAuthorizedEmail();
    if (!enteredByEmail) return NextResponse.json({ message: 'Sign in with an active member account.' }, { status: 401 });
    const params = new URL(request.url).searchParams;
    const team = params.get('team') || '';
    const status = params.get('status') || 'all';
    const search = params.get('q')?.trim() || '';
    if (team && !isPlacementTeam(team)) return NextResponse.json({ message: 'Invalid team filter.' }, { status: 400 });
    if (!['all', 'active', 'inactive'].includes(status)) return NextResponse.json({ message: 'Invalid status filter.' }, { status: 400 });

    const query: Record<string, unknown> = {};
    if (team) query.team = team;
    if (status === 'active') query.active = { $ne: false };
    if (status === 'inactive') query.active = false;
    if (search) {
      const escaped = search.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
      query.$or = [
        { name: { $regex: escaped, $options: 'i' } },
        { email: { $regex: escaped, $options: 'i' } },
      ];
    }

    const db = await getDb();
    const members = await db.collection('members').find(query).project({
      _id: 0,
      memberId: 1,
      name: 1,
      team: 1,
      email: 1,
      active: 1,
      createdAt: 1,
      createdByEmail: 1,
    }).sort({ nameNormalized: 1, teamNormalized: 1 }).toArray();
    return NextResponse.json(members);
  } catch (error) {
    console.error('Unable to load members from MongoDB.', error);
    return NextResponse.json({ message: getApiErrorMessage(error, 'Unable to load the member directory.') }, { status: 500 });
  }
}

export async function POST(request: Request) {
  try {
    const enteredByEmail = await getAuthorizedEmail();
    if (!enteredByEmail) return NextResponse.json({ message: 'Sign in with an active member account.' }, { status: 401 });
    const payload = await request.json();
    const name = normalizeText(String(payload?.name ?? ''));
    const team = normalizeTeam(String(payload?.team ?? ''));
    const email = String(payload?.email ?? '').trim().toLowerCase();
    if (!name) return NextResponse.json({ message: 'Name is required.' }, { status: 400 });
    if (!isPlacementTeam(team)) return NextResponse.json({ message: 'Choose placement_cell or summer_internship.' }, { status: 400 });
    if (email && !validEmail(email)) return NextResponse.json({ message: 'Enter a valid email address or leave it blank.' }, { status: 400 });

    await getDb();
    const now = new Date();
    const member = {
      memberId: randomUUID(),
      name,
      nameNormalized: name.toLowerCase(),
      team,
      teamNormalized: team,
      ...(email ? { email, emailNormalized: email } : {}),
      active: true,
      createdAt: now,
      createdByEmail: enteredByEmail,
      updatedAt: now,
      updatedByEmail: enteredByEmail,
    };
    try {
      await withMongoTransaction(async (db, session) => {
        await db.collection('members').insertOne(member, { session });
        await db.collection('audit_log').insertOne({
          action: 'Member added',
          collection: 'members',
          by: enteredByEmail,
          at: now,
          after: { memberId: member.memberId, name, team, email: email || null, active: true },
        }, { session });
      });
    } catch (error) {
      if (safeError(error)) {
        return NextResponse.json({ message: 'A member with this email or name/team already exists.' }, { status: 409 });
      }
      throw error;
    }
    return NextResponse.json({ member }, { status: 201 });
  } catch (error) {
    console.error('Unable to add member.', error);
    return NextResponse.json({ message: getApiErrorMessage(error, 'Unable to add the member.') }, { status: 500 });
  }
}
