import { NextResponse } from 'next/server';
import { getAuthorizedEmail } from '@/lib/authorization';
import { databaseFailureMessage, getDb } from '@/lib/mongodb';

export const runtime = 'nodejs';
export const dynamic = 'force-dynamic';
export const maxDuration = 30;

export async function GET() {
  try {
    const email = await getAuthorizedEmail();
    if (!email) return NextResponse.json({ message: 'Sign in with an authorized account.' }, { status: 401 });

    const db = await getDb();
    const types = await db.collection('default_types')
      .find({ active: true }, { projection: { _id: 0, id: 1, name: 1, description: 1 } })
      .sort({ name: 1 })
      .toArray();
    return NextResponse.json(types);
  } catch (error) {
    console.error('Unable to load default types from MongoDB.', error);
    return NextResponse.json({ message: databaseFailureMessage(error, 'Unable to load default types.') }, { status: 500 });
  }
}
