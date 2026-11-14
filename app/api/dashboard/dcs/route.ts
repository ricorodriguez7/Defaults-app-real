import { NextResponse } from 'next/server';
import { getAuthorizedEmail } from '@/lib/authorization';
import { getApiErrorMessage } from '@/lib/database-errors';
import { getDb } from '@/lib/mongodb';

export const runtime = 'nodejs';
export const dynamic = 'force-dynamic';

export async function GET() {
  try {
    const email = await getAuthorizedEmail();
    if (!email) return NextResponse.json({ message: 'Sign in with an authorized account.' }, { status: 401 });
    const db = await getDb();
    const batchConditions = {
      junior: { $eq: ['$student.batch', 'junior'] },
      senior: { $eq: ['$student.batch', 'senior'] },
      notSet: { $eq: [{ $ifNull: ['$student.batch', ''] }, ''] },
    };
    const defaultBatchCounts = Object.fromEntries(Object.entries(batchConditions).flatMap(([batch, condition]) => [
      [`${batch}Removed`, { $sum: { $cond: [{ $and: [{ $eq: ['$status', 'removed'] }, condition] }, 1, 0] } }],
      [`${batch}Kept`, { $sum: { $cond: [{ $and: [{ $eq: ['$status', 'locked'] }, { $eq: ['$dcStatus', 'Decided'] }, condition] }, 1, 0] } }],
      [`${batch}ToDiscuss`, { $sum: { $cond: [{ $and: [{ $ne: ['$status', 'removed'] }, { $eq: ['$dcStatus', 'To discuss'] }, condition] }, 1, 0] } }],
      [`${batch}Deferred`, { $sum: { $cond: [{ $and: [{ $ne: ['$status', 'removed'] }, { $eq: ['$dcStatus', 'Deferred'] }, condition] }, 1, 0] } }],
    ]));
    const [dcs, defaultCounts, entryCounts] = await Promise.all([
      db.collection('dcs').find({}).project({ _id: 0 }).sort({ date: -1, createdAt: -1 }).toArray(),
      db.collection('defaults').aggregate([
        { $match: { $or: [{ removedInDcId: { $type: 'string' } }, { dcId: { $type: 'string' } }] } },
        { $lookup: { from: 'students', localField: 'rollNoNormalized', foreignField: 'rollNoNormalized', as: 'student' } },
        { $unwind: { path: '$student', preserveNullAndEmptyArrays: true } },
        {
          $group: {
            _id: { $ifNull: ['$removedInDcId', '$dcId'] },
            removed: { $sum: { $cond: [{ $eq: ['$status', 'removed'] }, 1, 0] } },
            kept: { $sum: { $cond: [{ $and: [{ $eq: ['$status', 'locked'] }, { $eq: ['$dcStatus', 'Decided'] }] }, 1, 0] } },
            toDiscuss: { $sum: { $cond: [{ $and: [{ $ne: ['$status', 'removed'] }, { $eq: ['$dcStatus', 'To discuss'] }] }, 1, 0] } },
            deferred: { $sum: { $cond: [{ $and: [{ $ne: ['$status', 'removed'] }, { $eq: ['$dcStatus', 'Deferred'] }] }, 1, 0] } },
            ...defaultBatchCounts,
          },
        },
      ]).toArray(),
      db.collection('dc_entries').aggregate([
        { $match: { removed: { $ne: true } } },
        { $lookup: { from: 'students', localField: 'rollNoNormalized', foreignField: 'rollNoNormalized', as: 'student' } },
        { $unwind: { path: '$student', preserveNullAndEmptyArrays: true } },
        {
          $group: {
            _id: '$dcId',
            called: { $sum: 1 },
            present: { $sum: { $cond: [{ $eq: ['$attendance', 'Present'] }, 1, 0] } },
            absent: { $sum: { $cond: [{ $eq: ['$attendance', 'Absent'] }, 1, 0] } },
            juniorCalled: { $sum: { $cond: [{ $eq: ['$student.batch', 'junior'] }, 1, 0] } },
            juniorPresent: { $sum: { $cond: [{ $and: [{ $eq: ['$student.batch', 'junior'] }, { $eq: ['$attendance', 'Present'] }] }, 1, 0] } },
            juniorAbsent: { $sum: { $cond: [{ $and: [{ $eq: ['$student.batch', 'junior'] }, { $eq: ['$attendance', 'Absent'] }] }, 1, 0] } },
            seniorCalled: { $sum: { $cond: [{ $eq: ['$student.batch', 'senior'] }, 1, 0] } },
            seniorPresent: { $sum: { $cond: [{ $and: [{ $eq: ['$student.batch', 'senior'] }, { $eq: ['$attendance', 'Present'] }] }, 1, 0] } },
            seniorAbsent: { $sum: { $cond: [{ $and: [{ $eq: ['$student.batch', 'senior'] }, { $eq: ['$attendance', 'Absent'] }] }, 1, 0] } },
            notSetCalled: { $sum: { $cond: [{ $eq: [{ $ifNull: ['$student.batch', ''] }, ''] }, 1, 0] } },
            notSetPresent: { $sum: { $cond: [{ $and: [{ $eq: [{ $ifNull: ['$student.batch', ''] }, ''] }, { $eq: ['$attendance', 'Present'] }] }, 1, 0] } },
            notSetAbsent: { $sum: { $cond: [{ $and: [{ $eq: [{ $ifNull: ['$student.batch', ''] }, ''] }, { $eq: ['$attendance', 'Absent'] }] }, 1, 0] } },
          },
        },
      ]).toArray(),
    ]);
    const defaultsByDc = new Map(defaultCounts.map((row) => [row._id, row]));
    const entriesByDc = new Map(entryCounts.map((row) => [row._id, row]));
    const rows: Array<Record<string, unknown>> = dcs.map((dc) => {
      const defaults = defaultsByDc.get(dc.id) ?? {};
      const entries = entriesByDc.get(dc.id) ?? {};
      return {
        ...dc,
        removedCount: defaults.removed ?? 0,
        keptCount: defaults.kept ?? 0,
        called: entries.called ?? 0,
        present: entries.present ?? 0,
        absent: entries.absent ?? 0,
        toDiscussCount: defaults.toDiscuss ?? 0,
        deferredCount: defaults.deferred ?? 0,
        batchCounts: {
          junior: { called: entries.juniorCalled ?? 0, present: entries.juniorPresent ?? 0, absent: entries.juniorAbsent ?? 0, removed: defaults.juniorRemoved ?? 0, kept: defaults.juniorKept ?? 0, toDiscuss: defaults.juniorToDiscuss ?? 0, deferred: defaults.juniorDeferred ?? 0 },
          senior: { called: entries.seniorCalled ?? 0, present: entries.seniorPresent ?? 0, absent: entries.seniorAbsent ?? 0, removed: defaults.seniorRemoved ?? 0, kept: defaults.seniorKept ?? 0, toDiscuss: defaults.seniorToDiscuss ?? 0, deferred: defaults.seniorDeferred ?? 0 },
          not_set: { called: entries.notSetCalled ?? 0, present: entries.notSetPresent ?? 0, absent: entries.notSetAbsent ?? 0, removed: defaults.notSetRemoved ?? 0, kept: defaults.notSetKept ?? 0, toDiscuss: defaults.notSetToDiscuss ?? 0, deferred: defaults.notSetDeferred ?? 0 },
        },
      };
    });
    const noDateCount = defaultsByDc.get('imported:no-date')?.removed ?? 0;
    if (noDateCount) rows.push({
      id: 'imported:no-date',
      date: '',
      status: 'Completed',
      source: 'import',
      imported: true,
      removedCount: noDateCount,
      keptCount: 0,
      called: 0,
      present: 0,
      absent: 0,
      toDiscussCount: 0,
      deferredCount: 0,
      createdAt: new Date(0),
    });
    return NextResponse.json(rows);
  } catch (error) {
    console.error('Unable to load dashboard DC removals.', error);
    return NextResponse.json({ message: getApiErrorMessage(error, 'Unable to load dashboard DC data.') }, { status: 500 });
  }
}
