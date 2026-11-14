import { NextResponse } from 'next/server';
import { getAuthorizedEmail } from '@/lib/authorization';
import { getApiErrorMessage } from '@/lib/database-errors';
import { getDb } from '@/lib/mongodb';

export const runtime = 'nodejs';
export const dynamic = 'force-dynamic';

type RouteContext = { params: { dcId: string } };

export async function GET(_request: Request, { params }: RouteContext) {
  try {
    const email = await getAuthorizedEmail();
    if (!email) return NextResponse.json({ message: 'Sign in with an authorized account.' }, { status: 401 });
    const db = await getDb();
    const [dc, records] = await Promise.all([
      db.collection('dcs').findOne({ id: params.dcId }, { projection: { _id: 0 } }),
      db.collection('defaults').find({
        status: 'removed',
        $or: [{ removedInDcId: params.dcId }, { dcId: params.dcId }],
      }, { projection: { _id: 0, rollNoNormalized: 0 } }).sort({ date: 1, rollNo: 1 }).toArray(),
    ]);
    const isNoDateImport = params.dcId === 'imported:no-date';
    if (!dc && !isNoDateImport) return NextResponse.json({ message: 'DC not found.' }, { status: 404 });
    const rollNos = Array.from(new Set(records.map((record) => record.rollNo)));
    const typeIds = Array.from(new Set(records.map((record) => record.typeId)));
    const [students, types] = await Promise.all([
      db.collection('students').find({ rollNo: { $in: rollNos } }).project({ _id: 0, rollNo: 1, name: 1, batch: 1 }).toArray(),
      db.collection('default_types').find({ id: { $in: typeIds } }).project({ _id: 0, id: 1, name: 1 }).toArray(),
    ]);
    const studentByRollNo = new Map(students.map((student) => [student.rollNo.toUpperCase(), student]));
    const typeById = new Map(types.map((type) => [type.id, type.name]));
    return NextResponse.json({
      dc: dc ?? { id: params.dcId, date: '', status: 'Completed', source: 'import', imported: true },
      records: records.map((record) => {
        const student = studentByRollNo.get(String(record.rollNo).toUpperCase());
        return {
          ...record,
          studentName: student?.name ?? record.rollNo,
          batch: student?.batch,
          typeName: typeById.get(record.typeId) ?? record.typeId,
          removeReason: record.removeReason ?? record.removalReason ?? record.notes ?? '',
        };
      }),
    });
  } catch (error) {
    console.error('Unable to load removed defaults for dashboard DC.', error);
    return NextResponse.json({ message: getApiErrorMessage(error, 'Unable to load removed defaults for this DC.') }, { status: 500 });
  }
}
