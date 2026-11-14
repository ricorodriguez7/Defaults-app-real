import { NextResponse } from 'next/server';
import { getAuthorizedEmail } from '@/lib/authorization';
import { getApiErrorMessage, requestExceedsBodyLimit } from '@/lib/database-errors';
import { getDb, writeAuditEntry } from '@/lib/mongodb';

export const runtime = 'nodejs';
export const dynamic = 'force-dynamic';
export const maxDuration = 30;

function normalizeRollNo(value: unknown) {
  return String(value ?? '').trim().toUpperCase();
}

function parseBatch(value: unknown) {
  const normalized = String(value ?? '').trim().toLowerCase();
  if (!normalized) return { valid: true as const, batch: undefined };
  if (normalized === 'junior' || normalized === 'junior batch') return { valid: true as const, batch: 'junior' as const };
  if (normalized === 'senior' || normalized === 'senior batch') return { valid: true as const, batch: 'senior' as const };
  return { valid: false as const, batch: undefined };
}

function batchLabel(batch: string | undefined) {
  return batch === 'junior' ? 'Junior batch' : batch === 'senior' ? 'Senior batch' : 'Not set';
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
        { projection: { _id: 0, rollNo: 1, name: 1, batch: 1, numberOfDefaults: 1, lockPenaltyStatus: 1 } },
      );
      return student
        ? NextResponse.json({ found: true, student: { ...student, batchLabel: batchLabel(student.batch) } })
        : NextResponse.json({ found: false, message: 'No student found for that roll number.' });
    }

    const roster = await students.find(
      {},
      { projection: { _id: 0, rollNo: 1, name: 1, batch: 1, numberOfDefaults: 1, lockPenaltyStatus: 1 } },
    ).sort({ rollNo: 1 }).toArray();
    const counts = await db.collection('defaults').aggregate<{ _id: string; active: number; removed: number }>([
      { $group: {
        _id: '$rollNoNormalized',
        active: { $sum: { $cond: [{ $ne: ['$status', 'removed'] }, 1, 0] } },
        removed: { $sum: { $cond: [{ $eq: ['$status', 'removed'] }, 1, 0] } },
      } },
    ]).toArray();
    const defaultsByRollNo = new Map(counts.map((entry) => [entry._id, entry]));
    return NextResponse.json(roster.map((student) => ({
      ...student,
      batchLabel: batchLabel(student.batch),
      numberOfDefaults: defaultsByRollNo.get(normalizeRollNo(student.rollNo))?.active ?? student.numberOfDefaults ?? 0,
      numberOfRemovedDefaults: defaultsByRollNo.get(normalizeRollNo(student.rollNo))?.removed ?? 0,
    })));
  } catch (error) {
    console.error('Unable to load students from MongoDB.', error);
    return NextResponse.json({ message: getApiErrorMessage(error, 'Unable to load the student roster.') }, { status: 500 });
  }
}

export async function POST(request: Request) {
  try {
    const email = await getAuthorizedEmail();
    if (!email) return NextResponse.json({ message: 'Sign in with an authorized account.' }, { status: 401 });
    if (requestExceedsBodyLimit(request)) {
      return NextResponse.json({ message: 'The roster upload exceeds 4 MB. Split the file and try again.' }, { status: 413 });
    }

    const payload = await request.json();
    const roster = Array.isArray(payload?.roster) ? payload.roster : [];
    if (!roster.length || roster.length > 10000) {
      return NextResponse.json({ message: 'Upload between 1 and 10,000 student records.' }, { status: 400 });
    }

    const normalizedRoster: Array<{
      rollNo: string;
      rollNoNormalized: string;
      name: string;
      batch?: 'junior' | 'senior';
      numberOfDefaults?: number;
      lockPenaltyStatus?: string;
    }> = [];
    const seen = new Set<string>();
    const duplicates = new Set<string>();

    const batchErrors: string[] = [];
    for (const [index, entry] of roster.entries()) {
      const rollNo = String(entry?.rollNo ?? entry?.['Roll No'] ?? entry?.['Roll Number'] ?? '').trim();
      const rollNoNormalized = normalizeRollNo(rollNo);
      const name = String(entry?.name ?? '').trim();
      const parsedBatch = parseBatch(entry?.batch ?? entry?.Batch);
      const sourceRow = Number.isInteger(entry?.rowNumber) && entry.rowNumber >= 2 ? entry.rowNumber : index + 2;

      if (!rollNoNormalized || !name) {
        return NextResponse.json({ message: `Row ${sourceRow}: each roster row requires a roll number and a name.` }, { status: 400 });
      }
      if (!parsedBatch.valid) batchErrors.push(`Row ${sourceRow}: Batch must be Junior batch or Senior batch.`);

      if (seen.has(rollNoNormalized)) duplicates.add(rollNoNormalized);
      seen.add(rollNoNormalized);
      normalizedRoster.push({
        rollNo,
        rollNoNormalized,
        name,
        ...(parsedBatch.batch ? { batch: parsedBatch.batch } : {}),
        ...(Number.isInteger(entry?.numberOfDefaults) && entry.numberOfDefaults >= 0
          ? { numberOfDefaults: entry.numberOfDefaults }
          : {}),
        ...(typeof entry?.lockPenaltyStatus === 'string' && entry.lockPenaltyStatus.trim()
          ? { lockPenaltyStatus: entry.lockPenaltyStatus.trim() }
          : {}),
      });
    }

    if (batchErrors.length) return NextResponse.json({ message: batchErrors.join('\n'), errors: batchErrors }, { status: 400 });
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
            ...(student.batch ? { batch: student.batch } : {}),
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
      category: 'members',
      after: { count: normalizedRoster.length, batchAssignments: normalizedRoster.filter((student) => student.batch).length },
    });

    const updated = await students.find(
      {},
      { projection: { _id: 0, rollNo: 1, name: 1, batch: 1, numberOfDefaults: 1, lockPenaltyStatus: 1 } },
    ).sort({ rollNo: 1 }).toArray();
    return NextResponse.json({ updated });
  } catch (error) {
    console.error('Unable to update students in MongoDB.', error);
    return NextResponse.json({ message: getApiErrorMessage(error, 'Unable to update the student roster.') }, { status: 500 });
  }
}
