import * as XLSX from 'xlsx';
import { NextResponse } from 'next/server';
import { getAuthorizedEmail } from '@/lib/authorization';
import { getApiErrorMessage } from '@/lib/database-errors';
import { getDb } from '@/lib/mongodb';

export const runtime = 'nodejs';
export const dynamic = 'force-dynamic';
export const maxDuration = 30;

const headers = ['Roll No', 'Company', 'Default Type', 'Date', 'Logged By', 'Team', 'Notes', 'Status', 'DC Date'];

export async function GET() {
  try {
    const email = await getAuthorizedEmail();
    if (!email) return NextResponse.json({ message: 'Sign in with an active member account.' }, { status: 401 });
    const db = await getDb();
    const [types, members] = await Promise.all([
      db.collection('default_types').find({ active: true }).sort({ name: 1 }).project({ _id: 0, name: 1 }).toArray(),
      db.collection('members').find({ active: { $ne: false }, team: { $in: ['placement_cell', 'summer_internship'] } })
        .sort({ nameNormalized: 1 }).project({ _id: 0, name: 1, team: 1 }).toArray(),
    ]);
    const workbook = XLSX.utils.book_new();
    const defaultsSheet = XLSX.utils.aoa_to_sheet([
      headers,
      ['EXAMPLE001', 'Example Company - delete this row', types[0]?.name ?? 'Example default type', '01-01-2026', members[0]?.name ?? 'Example Member', members[0]?.team === 'summer_internship' ? 'Summer Internship' : 'Placement Cell', 'Delete this sample row before uploading.', 'Logged', ''],
      ['EXAMPLE002', 'Example Company - delete this row', types[1]?.name ?? types[0]?.name ?? 'Example default type', '02-01-2026', members[1]?.name ?? members[0]?.name ?? 'Example Member', members[1]?.team === 'summer_internship' ? 'Summer Internship' : 'Placement Cell', 'Delete this sample row before uploading.', 'Locked', '02-01-2026'],
      ['EXAMPLE003', 'Example Company - delete this row', types[0]?.name ?? 'Example default type', '03-01-2026', members[0]?.name ?? 'Example Member', members[0]?.team === 'summer_internship' ? 'Summer Internship' : 'Placement Cell', 'Delete this sample row before uploading.', 'Removed', '03-01-2026'],
    ]);
    XLSX.utils.book_append_sheet(workbook, defaultsSheet, 'Defaults');
    const instructions = [
      ['Column', 'Required', 'Accepted values / format'],
      ['Roll No', 'Yes', 'Must exactly match a student roll number (case-insensitive).'],
      ['Company', 'Yes', 'Free text; leading and trailing spaces are trimmed.'],
      ['Default Type', 'Yes', 'Must match an active default type name, ignoring case.'],
      ['Date', 'Yes', 'A real Excel date or text in DD-MM-YYYY format; cannot be in the future.'],
      ['Logged By', 'No', 'When supplied, must match an active member name. If blank, the signed-in importing member is recorded. Add Team if multiple members share the name.'],
      ['Team', 'No', 'Placement Cell or Summer Internship; disambiguates identical member names.'],
      ['Notes', 'No', 'Free text.'],
      ['Status', 'No', 'Logged (default), Locked, or Removed.'],
      ['DC Date', 'No', 'Date of the DC decision for Locked or Removed records.'],
      ['Upload limits', '', 'Up to 3,000 non-blank rows and under 4 MB; .xlsx or .csv.'],
      ['Sample rows', '', 'Delete the three EXAMPLE rows before uploading.'],
      ['Dropdowns', '', 'The current SheetJS Community Edition workbook writer does not create Excel dropdown validations; use the exact values on the Lists sheet.'],
    ];
    XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet(instructions), 'Instructions');
    const listRows: unknown[][] = [['Default Types', 'Active Members', 'Team']];
    const rowCount = Math.max(types.length, members.length, 1);
    for (let index = 0; index < rowCount; index += 1) {
      const member = members[index];
      listRows.push([
        types[index]?.name ?? '',
        member?.name ?? '',
        member?.team === 'placement_cell' ? 'Placement Cell' : member?.team === 'summer_internship' ? 'Summer Internship' : '',
      ]);
    }
    XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet(listRows), 'Lists');
    const bytes = XLSX.write(workbook, { bookType: 'xlsx', type: 'buffer' }) as Buffer;
    const body = new Uint8Array(bytes.byteLength);
    body.set(bytes);
    return new NextResponse(body, {
      headers: {
        'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'Content-Disposition': 'attachment; filename="defaults_import_sample.xlsx"',
        'Cache-Control': 'private, no-store',
      },
    });
  } catch (error) {
    console.error('Unable to generate defaults import sample.', error);
    return NextResponse.json({ message: getApiErrorMessage(error, 'Unable to generate the sample workbook.') }, { status: 500 });
  }
}
