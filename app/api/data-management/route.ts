import { NextResponse } from 'next/server';
import { getAuthorizedEmail } from '@/lib/authorization';
import { getApiErrorMessage } from '@/lib/database-errors';
import { verifyDataBackupToken } from '@/lib/data-backup-token';
import { getDb } from '@/lib/mongodb';

export const runtime = 'nodejs';
export const dynamic = 'force-dynamic';
export const maxDuration = 30;

const resetCategories = [
  'defaults',
  'dcs',
  'importBatches',
  'auditLog',
  'students',
  'types',
  'members',
] as const;

type ResetCategory = typeof resetCategories[number];

function isResetCategory(value: unknown): value is ResetCategory {
  return typeof value === 'string' && resetCategories.includes(value as ResetCategory);
}

async function getResetCounts(email: string) {
  const db = await getDb();
  const ownMember = await db.collection('members').findOne(
    { emailNormalized: email },
    { projection: { _id: 1 } },
  );
  if (!ownMember) throw new Error('Your active member record could not be found.');

  const [
    defaults,
    defaultEvents,
    dcs,
    dcEntries,
    importBatches,
    auditLog,
    students,
    defaultTypes,
    rules,
    members,
  ] = await Promise.all([
    db.collection('defaults').countDocuments(),
    db.collection('default_events').countDocuments(),
    db.collection('dcs').countDocuments(),
    db.collection('dc_entries').countDocuments(),
    db.collection('import_batches').countDocuments(),
    db.collection('audit_log').countDocuments(),
    db.collection('students').countDocuments(),
    db.collection('default_types').countDocuments(),
    db.collection('rules').countDocuments(),
    db.collection('members').countDocuments({ _id: { $ne: ownMember._id } }),
  ]);
  return {
    defaults: defaults + defaultEvents,
    dcs: dcs + dcEntries,
    importBatches,
    auditLog,
    students,
    types: defaultTypes + rules,
    members,
  };
}

export async function GET() {
  try {
    const email = await getAuthorizedEmail();
    if (!email) return NextResponse.json({ message: 'Sign in with an active member account.' }, { status: 401 });
    const enabled = process.env.ALLOW_DATA_RESET === 'true';
    if (!enabled) return NextResponse.json({ enabled, counts: null });
    return NextResponse.json({ enabled, counts: await getResetCounts(email) });
  } catch (error) {
    console.error('Unable to load data management counts.', error);
    return NextResponse.json({ message: getApiErrorMessage(error, 'Unable to load data management.') }, { status: 500 });
  }
}

export async function POST(request: Request) {
  try {
    const email = await getAuthorizedEmail();
    if (!email) return NextResponse.json({ message: 'Sign in with an active member account.' }, { status: 401 });
    if (process.env.ALLOW_DATA_RESET !== 'true') {
      return NextResponse.json({ message: 'Data reset is disabled.' }, { status: 403 });
    }

    const payload = await request.json();
    const selected = Array.isArray(payload?.selected) ? payload.selected : [];
    if (!selected.length || selected.some((item: unknown) => !isResetCategory(item))) {
      return NextResponse.json({ message: 'Choose at least one valid data category to clear.' }, { status: 400 });
    }
    if (payload?.confirmation !== 'RESET') {
      return NextResponse.json({ message: 'Type RESET exactly to confirm.' }, { status: 400 });
    }
    if (!verifyDataBackupToken(payload?.backupToken, email)) {
      return NextResponse.json({ message: 'Export a fresh backup before clearing data.' }, { status: 400 });
    }

    const db = await getDb();
    const ownMember = await db.collection('members').findOne(
      { emailNormalized: email },
      { projection: { _id: 1 } },
    );
    if (!ownMember) return NextResponse.json({ message: 'Your active member record could not be found.' }, { status: 403 });

    const requested = new Set<ResetCategory>(selected);
    const removed: Record<string, number> = {};
    const errors: string[] = [];
    const clear = async (label: string, collectionName: string, filter: Record<string, unknown> = {}) => {
      try {
        const result = await db.collection(collectionName).deleteMany(filter);
        removed[label] = result.deletedCount;
      } catch (error) {
        console.error(`Unable to clear ${collectionName} during data reset.`, error);
        errors.push(`${collectionName}: ${getApiErrorMessage(error, 'Deletion failed.')}`);
      }
    };

    if (requested.has('defaults')) {
      await clear('defaults', 'defaults');
      await clear('default_events', 'default_events');
    }
    if (requested.has('dcs')) {
      await clear('dcs', 'dcs');
      await clear('dc_entries', 'dc_entries');
    }
    if (requested.has('importBatches')) await clear('import_batches', 'import_batches');
    if (requested.has('auditLog')) await clear('audit_log', 'audit_log');
    if (requested.has('students')) await clear('students', 'students');
    if (requested.has('types')) {
      await clear('default_types', 'default_types');
      await clear('rules', 'rules');
    }
    if (requested.has('members')) {
      await clear('members', 'members', { _id: { $ne: ownMember._id } });
    }

    const at = new Date();
    try {
      await db.collection('audit_log').insertOne({
        action: 'Data reset',
        collection: 'data_management',
        by: email,
        at,
        after: { requested: Array.from(requested), removed, errors },
      });
    } catch (error) {
      console.error('Unable to record the data reset in the audit log.', error);
      return NextResponse.json({
        message: getApiErrorMessage(error, 'Data was cleared, but the reset audit entry could not be written.'),
        removed,
        errors,
      }, { status: 500 });
    }

    return NextResponse.json({ removed, errors, message: errors.length ? 'Reset finished with errors.' : 'Selected data was cleared.' });
  } catch (error) {
    console.error('Unable to reset application data.', error);
    return NextResponse.json({ message: getApiErrorMessage(error, 'Unable to reset application data.') }, { status: 500 });
  }
}
