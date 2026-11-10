import * as XLSX from 'xlsx';
import { NextResponse } from 'next/server';
import { getAuthorizedEmail } from '@/lib/authorization';
import { getApiErrorMessage } from '@/lib/database-errors';
import { createDataBackupToken } from '@/lib/data-backup-token';
import { getDb } from '@/lib/mongodb';

export const runtime = 'nodejs';
export const dynamic = 'force-dynamic';
export const maxDuration = 30;

function workbookRows(records: Record<string, unknown>[]) {
  return records.map((record) => Object.fromEntries(
    Object.entries(record).map(([key, value]) => [
      key,
      value && typeof value === 'object' && !(value instanceof Date)
        ? JSON.stringify(value)
        : value,
    ]),
  ));
}

export async function GET() {
  try {
    const email = await getAuthorizedEmail();
    if (!email) return NextResponse.json({ message: 'Sign in with an active member account.' }, { status: 401 });
    const db = await getDb();
    const names = [
      'students',
      'members',
      'default_types',
      'rules',
      'defaults',
      'default_events',
      'dcs',
      'dc_entries',
      'import_batches',
      'audit_log',
    ];
    const workbook = XLSX.utils.book_new();
    for (const name of names) {
      const records = await db.collection(name).find({}).toArray();
      const rows = workbookRows(records as unknown as Record<string, unknown>[]);
      const sheet = rows.length ? XLSX.utils.json_to_sheet(rows) : XLSX.utils.aoa_to_sheet([['No records']]);
      XLSX.utils.book_append_sheet(workbook, sheet, name.slice(0, 31));
    }
    const bytes = XLSX.write(workbook, { bookType: 'xlsx', type: 'buffer' }) as Buffer;
    const body = new Uint8Array(bytes.byteLength);
    body.set(bytes);
    const token = createDataBackupToken(email, Date.now());
    return new NextResponse(body, {
      headers: {
        'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'Content-Disposition': 'attachment; filename="default-tracker-backup.xlsx"',
        'Cache-Control': 'private, no-store',
        'X-Data-Backup-Token': token,
      },
    });
  } catch (error) {
    console.error('Unable to export all application data.', error);
    return NextResponse.json({ message: getApiErrorMessage(error, 'Unable to export all application data.') }, { status: 500 });
  }
}
