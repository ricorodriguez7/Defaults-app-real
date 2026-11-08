import { ClientSession, Db, MongoClient } from 'mongodb';
import { randomUUID } from 'node:crypto';

const uri = process.env.MONGODB_URI;
const databaseName = process.env.MONGODB_DATABASE || 'default_tracker';

const defaultTypeSeeds = [
  { id: 't1', name: 'Interview no-show', description: 'Did not attend a scheduled interview.', appliesTo: 'all_companies', active: true },
  { id: 't2', name: 'Late withdrawal', description: 'Withdrew after application deadline.', appliesTo: 'all_companies', active: true },
  { id: 't3', name: 'Test absence', description: 'Absent from a company assessment.', appliesTo: 'all_companies', active: true },
];

type MongoGlobal = typeof globalThis & {
  _mongoClientPromise?: Promise<MongoClient>;
  _databaseSetupPromise?: Promise<void>;
};

const globalState = globalThis as MongoGlobal;

function getClientPromise() {
  if (!uri) {
    throw new Error('MONGODB_URI is not configured.');
  }

  if (!globalState._mongoClientPromise) {
    const clientPromise = new MongoClient(uri, {
      maxPoolSize: 5,
      serverSelectionTimeoutMS: 8000,
    }).connect();
    globalState._mongoClientPromise = clientPromise;
    void clientPromise.catch(() => {
      if (globalState._mongoClientPromise === clientPromise) {
        globalState._mongoClientPromise = undefined;
      }
    });
  }

  return globalState._mongoClientPromise;
}

async function setupDatabase(db: Db) {
  await Promise.all([
    db.collection('students').createIndex({ rollNoNormalized: 1 }, { unique: true }),
    db.collection('defaults').createIndex({ id: 1 }, { unique: true }),
    db.collection('defaults').createIndex({ rollNoNormalized: 1, date: -1 }),
    db.collection('default_types').createIndex({ id: 1 }, { unique: true }),
    db.collection('audit_log').createIndex({ at: -1 }),
    db.collection('default_events').createIndex({ defaultId: 1, at: -1 }),
    db.collection('default_events').createIndex({ rollNo: 1, at: -1 }),
    db.collection('default_events').createIndex({ memberId: 1, at: -1 }),
    db.collection('default_events').createIndex({ eventId: 1 }, { unique: true }),
    db.collection('dcs').createIndex({ date: -1, createdAt: -1 }),
    db.collection('dc_entries').createIndex({ dcId: 1, rollNo: 1 }, { unique: true }),
    db.collection('dc_entries').createIndex(
      { openStudentRollNo: 1 },
      { unique: true, partialFilterExpression: { openStudentRollNo: { $type: 'string' } } },
    ),
  ]);

  await db.collection('default_types').bulkWrite(defaultTypeSeeds.map((type) => ({
    updateOne: {
      filter: { id: type.id },
      update: { $setOnInsert: type },
      upsert: true,
    },
  })));

  const members = db.collection('members');
  const memberCollectionExists = await db.listCollections({ name: 'members' }, { nameOnly: true }).hasNext();
  const oldMemberIndexes = memberCollectionExists ? await members.listIndexes().toArray() : [];
  for (const index of oldMemberIndexes) {
    if (index.name === 'emailNormalized_1' && index.unique && !index.partialFilterExpression) {
      await members.dropIndex(index.name);
    }
  }

  const initialEmails = (process.env.INITIAL_ALLOWED_EMAILS || '')
    .split(',')
    .map((email) => email.trim().toLowerCase())
    .filter(Boolean);
  const hasLegacyAllowlist = await db.listCollections({ name: 'allowed_emails' }, { nameOnly: true }).hasNext();
  const legacyEmails = hasLegacyAllowlist ? await db.collection('allowed_emails').find({}).toArray() : [];
  for (const allowed of legacyEmails) {
    const email = String(allowed.emailNormalized || allowed.email || '').trim().toLowerCase();
    if (!email) continue;
    await members.updateOne(
      { emailNormalized: email },
      {
        $set: { email, emailNormalized: email },
        $setOnInsert: {
          memberId: randomUUID(),
          name: email,
          team: 'unassigned',
          active: allowed.active !== false,
          createdAt: allowed.createdAt || new Date(),
          createdByEmail: 'allowlist migration',
        },
      },
      { upsert: true },
    );
  }
  for (const email of initialEmails) {
    await members.updateOne(
      { emailNormalized: email },
      {
        $set: { email, emailNormalized: email },
        $setOnInsert: {
          memberId: randomUUID(),
          name: email,
          team: 'unassigned',
          active: true,
          createdAt: new Date(),
          createdByEmail: 'environment migration',
        },
      },
      { upsert: true },
    );
  }

  const allMembers = await members.find({}).toArray();
  for (const member of allMembers) {
    const name = String(member.name || '').trim().replace(/\s+/g, ' ');
    const team = String(member.team || 'unassigned').trim().toLowerCase().replace(/[\s-]+/g, '_');
    const storedEmail = typeof member.email === 'string' ? member.email : member.emailNormalized;
    const normalizedEmail = typeof storedEmail === 'string' ? storedEmail.trim().toLowerCase() : '';
    const set: Record<string, unknown> = {
      memberId: member.memberId || randomUUID(),
      name,
      nameNormalized: name.toLowerCase(),
      team,
      teamNormalized: team,
      active: member.active !== false,
    };
    const unset: Record<string, ''> = {};
    if (normalizedEmail) {
      set.email = normalizedEmail;
      set.emailNormalized = normalizedEmail;
    } else {
      unset.email = '';
      unset.emailNormalized = '';
    }
    await members.updateOne(
      { _id: member._id },
      { $set: set, ...(Object.keys(unset).length ? { $unset: unset } : {}) },
    );
  }

  await Promise.all([
    members.createIndex(
      { emailNormalized: 1 },
      { unique: true, partialFilterExpression: { emailNormalized: { $type: 'string' } } },
    ),
    members.createIndex({ nameNormalized: 1, teamNormalized: 1 }, { unique: true }),
  ]);

  await db.collection('defaults').updateMany(
    { dcStatus: { $exists: false } },
    { $set: { dcStatus: 'Not scheduled', dcId: null } },
  );

  if (hasLegacyAllowlist) {
    let migratedCount = 0;
    for (const allowed of legacyEmails) {
      const email = String(allowed.emailNormalized || allowed.email || '').trim().toLowerCase();
      if (email && !(await members.findOne({ emailNormalized: email }, { projection: { _id: 1 } }))) {
        throw new Error('Allowlist migration did not copy every account into members.');
      }
      if (email) migratedCount += 1;
    }
    if (migratedCount) {
      await db.collection('audit_log').insertOne({
        action: 'Legacy sign-in allowlist migrated to members',
        collection: 'members',
        by: 'system migration',
        at: new Date(),
        after: { accounts: migratedCount },
      });
    }
    await db.collection('allowed_emails').drop();
  }

  const existingDefaults = await db.collection('defaults').find({}).toArray();
  if (existingDefaults.length) {
    const now = new Date();
    await db.collection('members').updateOne(
      { memberId: 'imported' },
      { $setOnInsert: { memberId: 'imported', name: 'Imported', team: 'Imported', active: true, createdAt: now } },
      { upsert: true },
    );

    for (const record of existingDefaults) {
      const emailGuess = String(record.enteredByEmail || record.createdBy || (String(record.memberId || '').includes('@') ? record.memberId : '') || '').trim().toLowerCase();
      const memberGuess = String(record.loggedByMemberId || record.memberId || '');
      const candidateMember = (memberGuess
        ? await db.collection('members').findOne({ memberId: memberGuess })
        : null) || (memberGuess.includes('@')
        ? await db.collection('members').findOne({ emailNormalized: memberGuess.toLowerCase() })
        : null) || (!memberGuess && emailGuess
        ? await db.collection('members').findOne({ emailNormalized: emailGuess })
        : null);
      const knownMember = candidateMember &&
        (candidateMember.team === 'placement_cell' || candidateMember.team === 'summer_internship')
        ? candidateMember
        : null;
      const at = record.createdAt instanceof Date ? record.createdAt : now;
      const memberId = String(knownMember?.memberId || 'imported');
      const enteredByEmail = String(record.enteredByEmail || (emailGuess.includes('@') ? emailGuess : 'Imported'));

      await db.collection('defaults').updateOne(
        { _id: record._id },
        {
          $set: {
            loggedByMemberId: memberId,
            enteredByEmail,
            createdAt: at,
            updatedAt: record.updatedAt || at,
            updatedByMemberId: record.updatedByMemberId || memberId,
          },
        },
      );

      const hasLoggedEvent = await db.collection('default_events').findOne(
        { defaultId: record.id, type: 'logged' },
        { projection: { _id: 1, eventId: 1 } },
      );
      if (!hasLoggedEvent) {
        await db.collection('default_events').updateOne(
          { eventId: `legacy:${record.id}` },
          {
            $setOnInsert: {
              eventId: `legacy:${record.id}`,
              defaultId: record.id,
              rollNo: record.rollNo,
              type: 'logged',
              memberId,
              enteredByEmail,
              at,
              reason: 'Existing default imported during history migration.',
              note: '',
              dcReference: record.dcReference || '',
              changes: {},
            },
          },
          { upsert: true },
        );
      } else if (hasLoggedEvent.eventId === `legacy:${record.id}`) {
        await db.collection('default_events').updateOne(
          { eventId: `legacy:${record.id}` },
          { $set: { memberId, enteredByEmail } },
        );
      }

      if (record.status === 'locked' || record.status === 'removed') {
        const statusEventType = record.status;
        const statusAt = record.lockedAt instanceof Date || record.removedAt instanceof Date
          ? (record.lockedAt || record.removedAt)
          : (record.updatedAt instanceof Date ? record.updatedAt : at);
        await db.collection('default_events').updateOne(
          { eventId: `legacy-${statusEventType}:${record.id}` },
          {
            $setOnInsert: {
              eventId: `legacy-${statusEventType}:${record.id}`,
              defaultId: record.id,
              rollNo: record.rollNo,
              type: statusEventType,
              memberId: 'imported',
              enteredByEmail: 'Imported',
              at: statusAt,
              reason: `Existing ${statusEventType} status imported; original reason is unavailable.`,
              note: '',
              dcReference: record.dcReference || '',
              changes: {},
            },
          },
          { upsert: true },
        );
      }
    }
  }
}

export async function getMongoClient() {
  return getClientPromise();
}

export async function getDb() {
  const client = await getMongoClient();
  const db = client.db(databaseName);

  if (!globalState._databaseSetupPromise) {
    const setupPromise = setupDatabase(db);
    globalState._databaseSetupPromise = setupPromise;
    void setupPromise.catch(() => {
      if (globalState._databaseSetupPromise === setupPromise) {
        globalState._databaseSetupPromise = undefined;
      }
    });
  }

  await globalState._databaseSetupPromise;
  return db;
}

export async function withMongoTransaction<T>(
  operation: (db: Db, session: ClientSession) => Promise<T>,
) {
  const client = await getMongoClient();
  const db = client.db(databaseName);
  const session = client.startSession();
  try {
    let result: T | undefined;
    await session.withTransaction(async () => {
      result = await operation(db, session);
    });
    if (result === undefined) throw new Error('MongoDB transaction completed without a result.');
    return result;
  } finally {
    await session.endSession();
  }
}

export async function writeAuditEntry(entry: {
  action: string;
  collection: string;
  by: string;
  before?: unknown;
  after?: unknown;
}) {
  const db = await getDb();
  await db.collection('audit_log').insertOne({
    ...entry,
    at: new Date(),
  });
}
