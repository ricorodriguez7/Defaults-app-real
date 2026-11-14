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

    const [result] = await (await getDb()).collection('defaults').aggregate([
      {
        $facet: {
          statuses: [
            {
              $group: {
                _id: null,
                total: { $sum: 1 },
                logged: { $sum: { $cond: [{ $and: [{ $eq: ['$status', 'logged'] }, { $not: [{ $in: [{ $ifNull: ['$dcStatus', ''] }, ['To discuss', 'Deferred']] }] }] }, 1, 0] } },
                locked: { $sum: { $cond: [{ $eq: ['$status', 'locked'] }, 1, 0] } },
                removed: { $sum: { $cond: [{ $eq: ['$status', 'removed'] }, 1, 0] } },
                toDiscuss: { $sum: { $cond: [{ $and: [{ $ne: ['$status', 'removed'] }, { $ne: ['$status', 'locked'] }, { $eq: ['$dcStatus', 'To discuss'] }] }, 1, 0] } },
                deferred: { $sum: { $cond: [{ $and: [{ $ne: ['$status', 'removed'] }, { $ne: ['$status', 'locked'] }, { $eq: ['$dcStatus', 'Deferred'] }] }, 1, 0] } },
              },
            },
          ],
          byDc: [
            { $match: { $or: [{ dcId: { $type: 'string' } }, { removedInDcId: { $type: 'string' } }] } },
            {
              $group: {
                _id: { $ifNull: ['$removedInDcId', '$dcId'] },
                removed: { $sum: { $cond: [{ $eq: ['$status', 'removed'] }, 1, 0] } },
                kept: { $sum: { $cond: [{ $and: [{ $eq: ['$status', 'locked'] }, { $eq: ['$dcStatus', 'Decided'] }] }, 1, 0] } },
                toDiscuss: { $sum: { $cond: [{ $and: [{ $ne: ['$status', 'removed'] }, { $eq: ['$dcStatus', 'To discuss'] }] }, 1, 0] } },
                deferred: { $sum: { $cond: [{ $and: [{ $ne: ['$status', 'removed'] }, { $eq: ['$dcStatus', 'Deferred'] }] }, 1, 0] } },
              },
            },
          ],
        },
      },
    ]).toArray();
    return NextResponse.json({
      counts: result?.statuses?.[0] ?? { total: 0, logged: 0, locked: 0, removed: 0, toDiscuss: 0, deferred: 0 },
      byDc: result?.byDc ?? [],
    });
  } catch (error) {
    console.error('Unable to load dashboard summary.', error);
    return NextResponse.json({ message: getApiErrorMessage(error, 'Unable to load dashboard summary.') }, { status: 500 });
  }
}
