import { randomUUID } from 'node:crypto';
import { ObjectId } from 'mongodb';
import * as XLSX from 'xlsx';
import { NextResponse } from 'next/server';
import { getAuthorizedEmail } from '@/lib/authorization';
import { getApiErrorMessage } from '@/lib/database-errors';
import { syncCurrentDcAutoAdd } from '@/lib/dc-auto-add';
import { getDb, withMongoTransaction } from '@/lib/mongodb';

export const runtime = 'nodejs';
export const dynamic = 'force-dynamic';
export const maxDuration = 30;

const MAX_FILE_BYTES = 4 * 1024 * 1024;
const MAX_ROWS = 3000;
const chunkSize = 300;
const expectedHeaders = [
  'Roll No', 'Company', 'Default Type', 'Date', 'Logged By', 'Team', 'Notes', 'Status', 'DC Date',
];
const requiredHeaders = ['Roll No', 'Company', 'Default Type', 'Date', 'Logged By'];
const undoBlockingEvents = ['locked', 'removed', 'dc_scheduled', 'dc_rescheduled', 'dc_decision', 'dc_auto_added'];

type ImportIssue = { row: number; column: string; reason: string; severity: 'warning' | 'error' };
type ImportPlanRow = {
  row: number;
  original: Record<string, unknown>;
  rollNo: string;
  rollNoNormalized: string;
  company: string;
  companyNormalized: string;
  typeName: string;
  typeId?: string;
  loggedByMemberId: string;
  date: string;
  notes: string;
  status: 'logged' | 'locked' | 'removed';
  dcDate: string;
};
type ImportPlan = {
  rows: ImportPlanRow[];
  issues: ImportIssue[];
  rejectedRows: Array<Record<string, unknown>>;
  counts: { rowsRead: number; ready: number; warnings: number; errors: number };
};

class ImportValidationError extends Error {}

const normalize = (value: unknown) => String(value ?? '').trim();
const normalizeKey = (value: unknown) => normalize(value).toLowerCase();
const normalizeRollNo = (value: unknown) => normalize(value).toUpperCase();
const normalizeTeam = (value: unknown) => normalize(value).toLowerCase().replace(/[\s-]+/g, '_');
const isEmpty = (value: unknown) => value === null || value === undefined || normalize(value) === '';

function getIndiaDate() {
  const parts = new Intl.DateTimeFormat('en-GB', {
    timeZone: 'Asia/Kolkata',
    year: 'numeric',
    month: '2-digit',
    day: '2-digit',
  }).formatToParts(new Date());
  const values = Object.fromEntries(parts.map((part) => [part.type, part.value]));
  return `${values.year}-${values.month}-${values.day}`;
}

function excelDate(value: unknown) {
  let year: number;
  let month: number;
  let day: number;
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    year = value.getUTCFullYear();
    month = value.getUTCMonth() + 1;
    day = value.getUTCDate();
  } else if (typeof value === 'number' && Number.isFinite(value)) {
    const parsed = XLSX.SSF.parse_date_code(value);
    if (!parsed) return null;
    ({ y: year, m: month, d: day } = parsed);
  } else {
    const text = normalize(value);
    const match = /^(\d{2})-(\d{2})-(\d{4})$/.exec(text);
    if (!match) return null;
    day = Number(match[1]);
    month = Number(match[2]);
    year = Number(match[3]);
  }
  const date = new Date(Date.UTC(year, month - 1, day));
  if (date.getUTCFullYear() !== year || date.getUTCMonth() + 1 !== month || date.getUTCDate() !== day) return null;
  return date.toISOString().slice(0, 10);
}

function issueRow(row: number, column: string, reason: string, severity: ImportIssue['severity']): ImportIssue {
  return { row, column, reason, severity };
}

async function parseUpload(file: File) {
  const bytes = await file.arrayBuffer();
  let workbook: XLSX.WorkBook;
  try {
    workbook = XLSX.read(bytes, { type: 'array', cellDates: true });
  } catch {
    throw new ImportValidationError('The file could not be read as a valid Excel or CSV workbook.');
  }
  if (!workbook.SheetNames.length) throw new ImportValidationError('The uploaded file does not contain a worksheet.');
  const sheetName = file.name.toLowerCase().endsWith('.csv')
    ? workbook.SheetNames[0]
    : (workbook.SheetNames.includes('Defaults') ? 'Defaults' : '');
  if (!sheetName) throw new ImportValidationError('The Excel workbook must contain a sheet named Defaults.');
  const sheet = workbook.Sheets[sheetName];
  const values = XLSX.utils.sheet_to_json<unknown[]>(sheet, {
    header: 1,
    raw: true,
    defval: '',
    blankrows: true,
  });
  const headers = (values[0] ?? []).map((value) => normalize(value).replace(/^\uFEFF/, ''));
  if (!headers.length || headers.every((header) => !header)) throw new ImportValidationError('The file is missing its header row.');
  const headerIndexes = new Map<string, number>();
  headers.forEach((header, index) => {
    if (header) headerIndexes.set(normalizeKey(header), index);
  });
  const missing = requiredHeaders.filter((header) => !headerIndexes.has(normalizeKey(header)));
  if (missing.length) throw new ImportValidationError(`Missing required column${missing.length > 1 ? 's' : ''}: ${missing.join(', ')}.`);
  const dataRows = values.slice(1).map((cells, index) => {
    const original: Record<string, unknown> = {};
    headers.forEach((header, columnIndex) => {
      if (header) original[header] = cells[columnIndex] ?? '';
    });
    return { row: index + 2, cells, original };
  }).filter(({ cells }) => cells.some((value) => !isEmpty(value)));
  if (dataRows.length > MAX_ROWS) throw new ImportValidationError(`Upload no more than ${MAX_ROWS.toLocaleString()} non-blank rows at a time.`);

  const allowedHeaders = new Set([...expectedHeaders.map(normalizeKey), 'reason']);
  const issues: ImportIssue[] = [];
  headers.forEach((header) => {
    if (header && !allowedHeaders.has(normalizeKey(header))) {
      issues.push(issueRow(1, header, 'Unrecognized column; it will be ignored.', 'warning'));
    }
  });
  const getCell = (cells: unknown[], key: string) => {
    const index = headerIndexes.get(normalizeKey(key));
    return index === undefined ? '' : cells[index];
  };
  return { dataRows, getCell, issues };
}

async function makePlan(file: File, createMissingTypes: boolean, importingMemberId: string): Promise<ImportPlan> {
  const parsed = await parseUpload(file);
  const { dataRows, getCell, issues } = parsed;
  const db = await getDb();
  const rollNos = Array.from(new Set(dataRows.map(({ cells }) => normalizeRollNo(getCell(cells, 'Roll No'))).filter(Boolean)));
  const [students, members, types, existingDefaults] = await Promise.all([
    db.collection('students').find({
      $or: [
        { rollNoNormalized: { $in: rollNos } },
        { rollNo: { $in: rollNos } },
      ],
    }).project({ _id: 0, rollNo: 1, rollNoNormalized: 1, name: 1 }).toArray(),
    db.collection('members').find({ active: { $ne: false } })
      .project({ _id: 0, memberId: 1, name: 1, team: 1, active: 1 }).toArray(),
    db.collection('default_types').find({ active: { $ne: false } })
      .project({ _id: 0, id: 1, name: 1 }).toArray(),
    rollNos.length
      ? db.collection('defaults').find({
        $or: [
          { rollNoNormalized: { $in: rollNos } },
          { rollNo: { $in: rollNos } },
        ],
      }).project({ _id: 0, rollNo: 1, rollNoNormalized: 1, company: 1, typeId: 1, date: 1 }).toArray()
      : Promise.resolve([]),
  ]);
  const studentByRollNo = new Map<string, typeof students[number]>();
  students.forEach((student) => {
    const key = normalizeRollNo(student.rollNoNormalized || student.rollNo);
    if (key) studentByRollNo.set(key, student);
  });
  const typeByName = new Map<string, typeof types[number]>();
  types.forEach((type) => typeByName.set(normalizeKey(type.name), type));
  const membersByName = new Map<string, typeof members>();
  members.forEach((member) => {
    const key = normalizeKey(member.name);
    membersByName.set(key, [...(membersByName.get(key) ?? []), member]);
  });
  const existingKeys = new Set(existingDefaults.map((record) => [
    normalizeRollNo(record.rollNoNormalized || record.rollNo),
    normalizeKey(record.company),
    String(record.typeId ?? ''),
    String(record.date ?? '').slice(0, 10),
  ].join('\u0000')));
  const seenKeys = new Set<string>();
  const planRows: ImportPlanRow[] = [];
  const rejectedRows: Array<Record<string, unknown>> = [];
  const today = getIndiaDate();

  for (const { row, cells, original } of dataRows) {
    const rowIssues: ImportIssue[] = [];
    const addError = (column: string, reason: string) => rowIssues.push(issueRow(row, column, reason, 'error'));
    const addWarning = (column: string, reason: string) => rowIssues.push(issueRow(row, column, reason, 'warning'));
    const rollNo = normalize(getCell(cells, 'Roll No'));
    const rollNoNormalized = normalizeRollNo(rollNo);
    const company = normalize(getCell(cells, 'Company'));
    const typeName = normalize(getCell(cells, 'Default Type'));
    const date = excelDate(getCell(cells, 'Date'));
    const memberName = normalize(getCell(cells, 'Logged By'));
    const teamInput = normalize(getCell(cells, 'Team'));
    const notes = normalize(getCell(cells, 'Notes'));
    const statusText = normalize(getCell(cells, 'Status')) || 'Logged';
    const dcDateValue = getCell(cells, 'DC Date');
    const dcDate = isEmpty(dcDateValue) ? '' : excelDate(dcDateValue);

    if (!rollNo) addError('Roll No', 'Roll No is required.');
    else if (!studentByRollNo.has(rollNoNormalized)) addError('Roll No', `Roll No ${rollNo} not found in the student roster.`);
    if (!company) addError('Company', 'Company is required.');

    const type = typeByName.get(normalizeKey(typeName));
    if (!typeName) addError('Default Type', 'Default Type is required.');
    else if (!type && createMissingTypes) addWarning('Default Type', `Default type "${typeName}" will be created during import.`);
    else if (!type) addError('Default Type', `Default type "${typeName}" does not exist.`);

    if (!date) addError('Date', 'Enter a real Excel date or a date in DD-MM-YYYY format.');
    else if (date > today) addError('Date', 'Date cannot be in the future.');

    const team = teamInput ? normalizeTeam(teamInput) : '';
    if (teamInput && !['placement_cell', 'summer_internship'].includes(team)) {
      addError('Team', 'Team must be Placement Cell or Summer Internship.');
    }
    const matchingMembers = (memberName
      ? membersByName.get(normalizeKey(memberName)) ?? []
      : members.filter((member) => member.memberId === importingMemberId))
      .filter((member) => !team || normalizeTeam(member.team) === team);
    if (!memberName && !matchingMembers.length) addError('Logged By', 'The importing member could not be resolved.');
    else if (!matchingMembers.length) addError('Logged By', `Active member "${memberName}" not found.`);
    else if (matchingMembers.length > 1) addError('Team', 'More than one active member has this name; specify Team to identify them.');

    const status = statusText.toLowerCase();
    if (!['logged', 'locked', 'removed'].includes(status)) {
      addError('Status', 'Status must be Logged, Locked, or Removed.');
    }
    if (!isEmpty(dcDateValue) && !dcDate) addError('DC Date', 'Enter a real Excel date or a date in DD-MM-YYYY format.');

    const typeId = type?.id;
    if (!rowIssues.some((item) => item.severity === 'error') && date && matchingMembers[0]) {
      const key = [
        rollNoNormalized,
        normalizeKey(company),
        typeId ?? `new:${normalizeKey(typeName)}`,
        date,
      ].join('\u0000');
      if (existingKeys.has(key) || seenKeys.has(key)) {
        addWarning('Roll No', 'Duplicate default; it already exists or appeared earlier in this file and will be skipped.');
      } else {
        seenKeys.add(key);
        planRows.push({
          row,
          original,
          rollNo,
          rollNoNormalized,
          company,
          companyNormalized: normalizeKey(company),
          typeName,
          ...(typeId ? { typeId } : {}),
          loggedByMemberId: matchingMembers[0].memberId,
          date,
          notes,
          status: status as ImportPlanRow['status'],
          dcDate: dcDate || '',
        });
      }
    }
    issues.push(...rowIssues);
    if (rowIssues.some((item) => item.severity === 'error') || rowIssues.some((item) => item.reason.startsWith('Duplicate default'))) {
      const reasons = rowIssues.map((item) => `${item.column}: ${item.reason}`).join(' | ');
      rejectedRows.push({ ...original, Reason: reasons });
    }
  }

  const warningRows = new Set(issues.filter((item) => item.severity === 'warning').map((item) => item.row));
  const errorRows = new Set(issues.filter((item) => item.severity === 'error').map((item) => item.row));
  return {
    rows: planRows,
    issues,
    rejectedRows,
    counts: {
      rowsRead: dataRows.length,
      ready: planRows.length,
      warnings: warningRows.size,
      errors: errorRows.size,
    },
  };
}

function report(plan: ImportPlan) {
  return {
    counts: plan.counts,
    issues: plan.issues,
    rejectedRows: plan.rejectedRows,
  };
}

export async function GET() {
  try {
    const email = await getAuthorizedEmail();
    if (!email) return NextResponse.json({ message: 'Sign in with an active member account.' }, { status: 401 });
    const db = await getDb();
    const batches = await db.collection('import_batches').find({})
      .sort({ importedAt: -1 }).limit(50).toArray();
    const batchIds = batches.map((batch) => String(batch.batchId || batch._id));
    const allDefaultIds = batches.flatMap((batch) => Array.isArray(batch.defaultIds) ? batch.defaultIds : []);
    const [defaults, events] = allDefaultIds.length
      ? await Promise.all([
        db.collection('defaults').find({ id: { $in: allDefaultIds } })
          .project({ _id: 0, id: 1, status: 1, dcId: 1, dcStatus: 1 }).toArray(),
        db.collection('default_events').find({
          defaultId: { $in: allDefaultIds },
          type: { $in: undoBlockingEvents },
        }).project({ _id: 0, defaultId: 1, eventId: 1, type: 1, at: 1 }).toArray(),
      ])
      : [[], []];
    const defaultById = new Map(defaults.map((record) => [record.id, record]));
    const eventsByDefault = new Map<string, typeof events>();
    events.forEach((event) => {
      eventsByDefault.set(event.defaultId, [...(eventsByDefault.get(event.defaultId) ?? []), event]);
    });
    return NextResponse.json(batches.map((batch, index) => {
      const importedAt = batch.importedAt instanceof Date ? batch.importedAt : new Date(batch.importedAt);
      const ids = Array.isArray(batch.defaultIds) ? batch.defaultIds as string[] : [];
      const initialEventIds = new Set(Array.isArray(batch.initialEventIds) ? batch.initialEventIds as string[] : []);
      let undoBlockReason = '';
      if (batch.undoneAt) undoBlockReason = 'This import has already been undone.';
      else {
        for (const id of ids) {
          const record = defaultById.get(id);
          if (!record) continue;
          const initialStatus = batch.initialStatuses?.[id] ?? 'logged';
          if (record.dcId || record.dcStatus === 'Scheduled' || record.status !== initialStatus) {
            undoBlockReason = 'One or more defaults were added to a DC, locked, or removed after import.';
            break;
          }
          const event = (eventsByDefault.get(id) ?? []).find((item) =>
            !initialEventIds.has(item.eventId) && new Date(item.at).getTime() > importedAt.getTime());
          if (event) {
            undoBlockReason = 'One or more defaults were added to a DC, locked, or removed after import.';
            break;
          }
        }
      }
      return {
        batchId: batch.batchId || batchIds[index],
        filename: batch.filename,
        importedBy: batch.importedBy,
        importedAt: batch.importedAt,
        rowsRead: batch.rowsRead,
        imported: batch.imported,
        skipped: batch.skipped,
        errors: batch.errors,
        undoneAt: batch.undoneAt,
        undoBlockReason,
      };
    }));
  } catch (error) {
    console.error('Unable to load defaults import history.', error);
    return NextResponse.json({ message: getApiErrorMessage(error, 'Unable to load defaults import history.') }, { status: 500 });
  }
}

export async function POST(request: Request) {
  try {
    const enteredByEmail = await getAuthorizedEmail();
    if (!enteredByEmail) return NextResponse.json({ message: 'Sign in with an active member account.' }, { status: 401 });
    if (Number(request.headers.get('content-length') ?? 0) > MAX_FILE_BYTES + 65536) {
      return NextResponse.json({ message: 'The upload exceeds the 4 MB limit.' }, { status: 413 });
    }
    const form = await request.formData();
    const file = form.get('file');
    if (!(file instanceof File)) return NextResponse.json({ message: 'Choose an Excel or CSV file.' }, { status: 400 });
    if (file.size > MAX_FILE_BYTES) return NextResponse.json({ message: 'The upload must be under 4 MB.' }, { status: 413 });
    if (!/\.(xlsx|csv)$/i.test(file.name)) return NextResponse.json({ message: 'Upload an .xlsx or .csv file.' }, { status: 400 });
    const mode = form.get('mode');
    if (mode !== 'preview' && mode !== 'import') return NextResponse.json({ message: 'Choose preview or import mode.' }, { status: 400 });
    const createMissingTypes = form.get('createMissingTypes') === 'true';
    const stopOnError = form.get('stopOnError') === 'true';
    const db = await getDb();
    const actor = await db.collection('members').findOne(
      { emailNormalized: enteredByEmail, active: { $ne: false } },
      { projection: { _id: 0, memberId: 1 } },
    );
    if (!actor) return NextResponse.json({ message: 'Your active member record could not be found.' }, { status: 403 });
    const plan = await makePlan(file, createMissingTypes, actor.memberId);
    if (mode === 'preview') return NextResponse.json(report(plan));
    const batchId = randomUUID();
    const now = new Date();
    const shouldStop = stopOnError && plan.counts.errors > 0;
    let imported = 0;
    const initialStatuses: Record<string, string> = {};
    const defaultIds: string[] = [];
    const initialEventIds: string[] = [];

    const importResult = await withMongoTransaction(async (transactionDb, session) => {
      imported = 0;
      defaultIds.length = 0;
      initialEventIds.length = 0;
      Object.keys(initialStatuses).forEach((id) => delete initialStatuses[id]);
      const types = await transactionDb.collection('default_types').find(
        { active: { $ne: false } },
        { session, projection: { _id: 0, id: 1, name: 1 } },
      ).toArray();
      const typeByKey = new Map(types.map((type) => [normalizeKey(type.name), type]));
      if (!shouldStop && createMissingTypes) {
        const missingTypeNames = Array.from(new Set(plan.rows
          .map((row) => row.typeName)
          .filter((name) => !typeByKey.has(normalizeKey(name)))));
        if (missingTypeNames.length) {
          const createdTypes = missingTypeNames.map((name) => ({
            _id: new ObjectId(),
            id: randomUUID(),
            name,
            description: 'Created during Excel import.',
            appliesTo: 'all_companies',
            active: true,
          }));
          await transactionDb.collection<{
            id: string;
            name: string;
            description: string;
            appliesTo: string;
            active: boolean;
          }>('default_types').insertMany(createdTypes, { session });
          createdTypes.forEach((type) => typeByKey.set(normalizeKey(type.name), type));
        }
      }

      const importableRows = shouldStop ? [] : plan.rows;
      const records = importableRows.map((row) => {
        const type = typeByKey.get(normalizeKey(row.typeName));
        if (!type) throw new Error(`Default type "${row.typeName}" could not be resolved during import.`);
        const id = randomUUID();
        const statusDate = row.dcDate ? new Date(`${row.dcDate}T00:00:00.000Z`) : now;
        const record = {
          id,
          rollNo: row.rollNo,
          rollNoNormalized: row.rollNoNormalized,
          memberId: row.loggedByMemberId,
          loggedByMemberId: row.loggedByMemberId,
          enteredByEmail,
          typeId: type.id,
          company: row.company,
          date: row.date,
          status: row.status,
          dcReference: row.dcDate,
          dcId: null,
          dcStatus: row.status === 'logged' ? 'Not scheduled' : 'Decided',
          notes: row.notes,
          source: 'excel_import',
          importBatchId: batchId,
          createdAt: now,
          updatedAt: now,
          updatedByMemberId: actor.memberId,
          ...(row.status === 'locked' ? { lockedAt: statusDate } : {}),
          ...(row.status === 'removed' ? { removedAt: statusDate } : {}),
          ...(row.dcDate ? { dcDate: row.dcDate } : {}),
        };
        initialStatuses[id] = row.status;
        defaultIds.push(id);
        const loggedEventId = randomUUID();
        const events = [{
          eventId: loggedEventId,
          defaultId: id,
          rollNo: row.rollNo,
          type: 'logged',
          memberId: row.loggedByMemberId,
          enteredByEmail,
          at: now,
          importBatchId: batchId,
          actor: actor.memberId,
          reason: 'Imported from Excel.',
          note: '',
          dcReference: row.dcDate,
          changes: {},
        }];
        initialEventIds.push(loggedEventId);
        if (row.status !== 'logged') {
          const statusEventId = randomUUID();
          events.push({
            eventId: statusEventId,
            defaultId: id,
            rollNo: row.rollNo,
            type: row.status,
            memberId: actor.memberId,
            enteredByEmail,
            at: statusDate,
            importBatchId: batchId,
            actor: actor.memberId,
            reason: `Imported with ${row.status} status.`,
            note: '',
            dcReference: row.dcDate,
            changes: { status: { from: 'logged', to: row.status } },
          });
          initialEventIds.push(statusEventId);
        }
        return { row, record, events };
      });

      for (let offset = 0; offset < records.length; offset += chunkSize) {
        const chunk = records.slice(offset, offset + chunkSize);
        await transactionDb.collection('defaults').bulkWrite(chunk.map(({ record }) => ({
          insertOne: { document: record },
        })), { session });
        await transactionDb.collection('default_events').insertMany(
          chunk.flatMap(({ events }) => events),
          { session },
        );
      }
      imported = records.length;
      const autoAdd = records.length
        ? await syncCurrentDcAutoAdd(transactionDb, session, actor.memberId, enteredByEmail, now)
        : undefined;
      const skipped = plan.counts.rowsRead - imported;
      await transactionDb.collection('import_batches').insertOne({
        batchId,
        filename: file.name,
        importedBy: enteredByEmail,
        importedAt: now,
        rowsRead: plan.counts.rowsRead,
        imported,
        skipped,
        errors: plan.counts.errors,
        defaultIds,
        initialStatuses,
        initialEventIds,
      }, { session });
      await transactionDb.collection('audit_log').insertOne({
        action: 'Defaults imported from Excel',
        collection: 'defaults',
        by: enteredByEmail,
        at: now,
        after: { batchId, filename: file.name, rowsRead: plan.counts.rowsRead, imported, skipped, errors: plan.counts.errors, autoAdded: autoAdd?.added ?? 0 },
      }, { session });
      return { autoAdd };
    });

    return NextResponse.json({
      ...report(plan),
      imported,
      skipped: plan.counts.rowsRead - imported,
      stoppedForErrors: shouldStop,
      autoAdd: importResult.autoAdd,
    });
  } catch (error) {
    if (error instanceof ImportValidationError) {
      return NextResponse.json({ message: error.message }, { status: 400 });
    }
    console.error('Unable to import defaults from Excel.', error);
    return NextResponse.json({ message: getApiErrorMessage(error, 'Unable to import defaults from Excel.') }, { status: 500 });
  }
}

export async function DELETE(request: Request) {
  try {
    const email = await getAuthorizedEmail();
    if (!email) return NextResponse.json({ message: 'Sign in with an active member account.' }, { status: 401 });
    const payload = await request.json();
    const batchId = String(payload?.batchId ?? '').trim();
    if (!batchId) return NextResponse.json({ message: 'Import batch ID is required.' }, { status: 400 });
    const db = await getDb();
    const batch = await db.collection('import_batches').findOne({ batchId });
    if (!batch) return NextResponse.json({ message: 'Import batch not found.' }, { status: 404 });
    if (batch.undoneAt) return NextResponse.json({ message: 'This import has already been undone.' }, { status: 409 });
    const ids = Array.isArray(batch.defaultIds) ? batch.defaultIds as string[] : [];
    const records = ids.length ? await db.collection('defaults').find({ id: { $in: ids } }).toArray() : [];
    const importedAt = batch.importedAt instanceof Date ? batch.importedAt : new Date(batch.importedAt);
    const events = ids.length ? await db.collection('default_events').find({
      defaultId: { $in: ids },
      type: { $in: undoBlockingEvents },
      eventId: { $nin: batch.initialEventIds ?? [] },
      at: { $gt: importedAt },
    }).limit(1).toArray() : [];
    const changedStatus = records.some((record) =>
      record.status !== (batch.initialStatuses?.[record.id] ?? 'logged') ||
      Boolean(record.dcId) ||
      record.dcStatus === 'Scheduled',
    );
    if (events.length || changedStatus) {
      return NextResponse.json({
        message: 'Undo is blocked because one or more imported defaults were added to a DC, locked, or removed after import.',
      }, { status: 409 });
    }

    const now = new Date();
    const removed = await withMongoTransaction(async (transactionDb, session) => {
      if (ids.length) {
        await transactionDb.collection('default_events').deleteMany({ defaultId: { $in: ids } }, { session });
      }
      const result = await transactionDb.collection('defaults').deleteMany({ importBatchId: batchId }, { session });
      await transactionDb.collection('import_batches').updateOne(
        { batchId },
        { $set: { undoneAt: now, undoneBy: email, undoneDefaults: result.deletedCount } },
        { session },
      );
      await transactionDb.collection('audit_log').insertOne({
        action: 'Excel defaults import undone',
        collection: 'defaults',
        by: email,
        at: now,
        after: { batchId, defaultsRemoved: result.deletedCount },
      }, { session });
      return result.deletedCount;
    });
    return NextResponse.json({ undone: true, removed });
  } catch (error) {
    console.error('Unable to undo defaults import.', error);
    return NextResponse.json({ message: getApiErrorMessage(error, 'Unable to undo defaults import.') }, { status: 500 });
  }
}
