import { NextResponse } from 'next/server';
import { getAuthorizedEmail } from '@/lib/authorization';
import { databaseFailureMessage, getDb, writeAuditEntry } from '@/lib/mongodb';
import { readJsonBody, RequestBodyError } from '@/lib/request-body';

export const runtime = 'nodejs';
export const dynamic = 'force-dynamic';
export const maxDuration = 30;

function normalizeRollNo(value: unknown) {
  return String(value ?? '').trim().toUpperCase();
}

export async function GET(request: Request) {
  try {
    const email = await getAuthorizedEmail();
    if (!email) return NextResponse.json({ message: 'Sign in with an authorized account.' }, { status: 401 });

    const db = await getDb();
    const rollNo = new URL(request.url).searchParams.get('rollNo');
    const students = db.collection('students');

    if (rollNo) {
      const student = await students.findOne(
        { rollNoNormalized: normalizeRollNo(rollNo) },
        { projection: { _id: 0, rollNo: 1, name: 1, numberOfDefaults: 1, lockPenaltyStatus: 1 } },
      );
      return student
        ? NextResponse.json({ found: true, student })
        : NextResponse.json({ found: false, message: 'No student found for that roll number.' });
    }

    const roster = await students.find(
      {},
      { projection: { _id: 0, rollNo: 1, name: 1, numberOfDefaults: 1, lockPenaltyStatus: 1 } },
    ).sort({ rollNo: 1 }).toArray();
    const counts = await db.collection('defaults').aggregate<{ _id: string; count: number }>([
      { $match: { status: { $ne: 'removed' } } },
      { $group: { _id: '$rollNoNormalized', count: { $sum: 1 } } },
    ]).toArray();
    const defaultsByRollNo = new Map(counts.map((entry) => [entry._id, entry.count]));
    return NextResponse.json(roster.map((student) => ({
      ...student,
      numberOfDefaults: defaultsByRollNo.get(normalizeRollNo(student.rollNo)) ?? student.numberOfDefaults ?? 0,
    })));
  } catch (error) {
    console.error('Unable to load students from MongoDB.', error);
    return NextResponse.json({ message: databaseFailureMessage(error, 'Unable to load the student roster.') }, { status: 500 });
  }
}

export async function POST(request: Request) {
  try {
    const email = await getAuthorizedEmail();
    if (!email) return NextResponse.json({ message: 'Sign in with an authorized account.' }, { status: 401 });

    const payload = await readJsonBody<{ roster?: Array<Record<string, unknown>> }>(request);
    const roster = Array.isArray(payload?.roster) ? payload.roster : [];
    if (!roster.length || roster.length > 10000) {
      return NextResponse.json({ message: 'Upload between 1 and 10,000 student records.' }, { status: 400 });
    }

    const normalizedRoster: Array<{
      rollNo: string;
      rollNoNormalized: string;
      name: string;
      numberOfDefaults?: number;
      lockPenaltyStatus?: string;
    }> = [];
    const seen = new Set<string>();
    const duplicates = new Set<string>();

    for (const entry of roster) {
      const rollNo = String(entry?.rollNo ?? '').trim();
      const rollNoNormalized = normalizeRollNo(rollNo);
      const name = String(entry?.name ?? '').trim();

      if (!rollNoNormalized || !name) {
        return NextResponse.json({ message: 'Each roster row requires a roll number and a name.' }, { status: 400 });
      }

      if (seen.has(rollNoNormalized)) duplicates.add(rollNoNormalized);
      seen.add(rollNoNormalized);
      normalizedRoster.push({
        rollNo,
        rollNoNormalized,
        name,
        ...(typeof entry.numberOfDefaults === 'number' && Number.isInteger(entry.numberOfDefaults) && entry.numberOfDefaults >= 0
          ? { numberOfDefaults: entry.numberOfDefaults }
          : {}),
        ...(typeof entry?.lockPenaltyStatus === 'string' && entry.lockPenaltyStatus.trim()
          ? { lockPenaltyStatus: entry.lockPenaltyStatus.trim() }
          : {}),
      });
    }

    if (duplicates.size) {
      return NextResponse.json({ message: `Duplicate roll numbers found: ${Array.from(duplicates).join(', ')}` }, { status: 400 });
    }

    const db = await getDb();
    const students = db.collection('students');
    const now = new Date();
    await students.bulkWrite(normalizedRoster.map((student) => ({
      updateOne: {
        filter: { rollNoNormalized: student.rollNoNormalized },
        update: {
          $set: {
            rollNo: student.rollNo,
            rollNoNormalized: student.rollNoNormalized,
            name: student.name,
            ...(student.numberOfDefaults !== undefined ? { numberOfDefaults: student.numberOfDefaults } : {}),
            ...(student.lockPenaltyStatus ? { lockPenaltyStatus: student.lockPenaltyStatus } : {}),
            updatedAt: now,
          },
          $setOnInsert: { createdAt: now },
        },
        upsert: true,
      },
    })));

    await writeAuditEntry({
      action: 'Students uploaded',
      collection: 'students',
      by: email,
      after: { count: normalizedRoster.length },
    });

    const updated = await students.find(
      {},
      { projection: { _id: 0, rollNo: 1, name: 1, numberOfDefaults: 1, lockPenaltyStatus: 1 } },
    ).sort({ rollNo: 1 }).toArray();
    return NextResponse.json({ updated });
  } catch (error) {
    if (error instanceof RequestBodyError) {
      return NextResponse.json({ message: error.message }, { status: error.status });
    }
    console.error('Unable to update students in MongoDB.', error);
    return NextResponse.json({ message: databaseFailureMessage(error, 'Unable to update the student roster.') }, { status: 500 });
  }
}
