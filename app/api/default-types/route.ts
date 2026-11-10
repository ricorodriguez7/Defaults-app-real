import { NextResponse } from 'next/server';
import { randomUUID } from 'node:crypto';
import { getAuthorizedEmail } from '@/lib/authorization';
import { getApiErrorMessage } from '@/lib/database-errors';
import { findActorMember } from '@/lib/members';
import { getDb, normalizeDefaultTypeName, withMongoTransaction } from '@/lib/mongodb';

export const runtime = 'nodejs';
export const dynamic = 'force-dynamic';

const normalizeName = (value: unknown) => String(value ?? '').trim().replace(/\s+/g, ' ');
const normalizeAppliesTo = (value: unknown) => String(value ?? '').trim().toLowerCase().replace(/[\s-]+/g, '_');
const validAppliesTo = (value: string) => ['all_companies', 'dream', 'core', 'normal'].includes(value);

export async function GET(request: Request) {
  try {
    const email = await getAuthorizedEmail();
    if (!email) return NextResponse.json({ message: 'Sign in with an authorized account.' }, { status: 401 });

    const db = await getDb();
    const includeInactive = new URL(request.url).searchParams.get('includeInactive') === 'true';
    const types = await db.collection('default_types')
      .find(includeInactive ? {} : { active: { $ne: false } }, { projection: { _id: 0, id: 1, name: 1, description: 1, appliesTo: 1, active: 1, updatedBy: 1, updatedByName: 1, updatedAt: 1 } })
      .sort({ name: 1 })
      .toArray();
    const ids = types.map((type) => type.id);
    const [usage, members] = await Promise.all([
      ids.length ? db.collection('defaults').aggregate([
        { $match: { typeId: { $in: ids } } },
        { $group: { _id: '$typeId', count: { $sum: 1 } } },
      ]).toArray() : [],
      db.collection('members').find({}).project({ _id: 0, memberId: 1, name: 1 }).toArray(),
    ]);
    const usageById = new Map(usage.map((item) => [item._id, item.count]));
    const nameById = new Map(members.map((member) => [member.memberId, member.name]));
    return NextResponse.json(types.map((type) => ({
      ...type,
      appliesTo: type.appliesTo || 'all_companies',
      active: type.active !== false,
      usedIn: usageById.get(type.id) ?? 0,
      lastChangedBy: nameById.get(type.updatedBy) ?? type.updatedByName ?? 'System',
    })));
  } catch (error) {
    console.error('Unable to load default types from MongoDB.', error);
    return NextResponse.json({ message: getApiErrorMessage(error, 'Unable to load default types.') }, { status: 500 });
  }
}

export async function POST(request: Request) {
    try {
      const email = await getAuthorizedEmail();
      if (!email) return NextResponse.json({ message: 'Sign in with an authorized account.' }, { status: 401 });
      const payload = await request.json();
      const name = normalizeName(payload?.name);
      const description = String(payload?.description ?? '').trim();
      const appliesTo = normalizeAppliesTo(payload?.appliesTo || 'all_companies');
      if (!name) return NextResponse.json({ message: 'Default type name is required.' }, { status: 400 });
      if (!validAppliesTo(appliesTo)) return NextResponse.json({ message: 'Applies to must be all companies, Dream, Core, or Normal.' }, { status: 400 });
      const result = await withMongoTransaction(async (db, session) => {
        const actor = await findActorMember(db, email);
        if (!actor) throw new Error('Your active member record could not be found.');
        const duplicate = await db.collection('default_types').findOne(
          { nameNormalized: normalizeDefaultTypeName(name) },
          { session, projection: { _id: 1 } },
        );
        if (duplicate) throw new Error('A default type with this name already exists.');
        const now = new Date();
        const type = {
          id: randomUUID(), name, nameNormalized: normalizeDefaultTypeName(name), description, appliesTo, active: true,
          createdBy: actor.memberId, updatedBy: actor.memberId, updatedByName: actor.name, updatedAt: now,
        };
        await db.collection('default_types').insertOne(type, { session });
        await db.collection('audit_log').insertOne({
          action: 'Default type added', collection: 'default_types', by: email, at: now,
          category: 'important', actor: actor.memberId,
          after: { id: type.id, name, description, appliesTo },
        }, { session });
        return type;
      });
      return NextResponse.json({ type: result }, { status: 201 });
    } catch (error) {
      const message = error instanceof Error ? error.message : '';
      if (message.startsWith('A default type') || message.startsWith('Your active')) {
        return NextResponse.json({ message }, { status: 400 });
      }
      console.error('Unable to add default type.', error);
      return NextResponse.json({ message: getApiErrorMessage(error, 'Unable to add default type.') }, { status: 500 });
    }
  }

export async function PATCH(request: Request) {
    try {
      const email = await getAuthorizedEmail();
      if (!email) return NextResponse.json({ message: 'Sign in with an authorized account.' }, { status: 401 });
      const payload = await request.json();
      const id = String(payload?.id ?? '').trim();
      const action = String(payload?.action ?? 'edit');
      if (!id) return NextResponse.json({ message: 'Default type ID is required.' }, { status: 400 });
      const result = await withMongoTransaction(async (db, session) => {
        const actor = await findActorMember(db, email);
        if (!actor) throw new Error('Your active member record could not be found.');
        const collection = db.collection('default_types');
        const type = await collection.findOne({ id }, { session });
        if (!type) throw new Error('Default type not found.');
        const now = new Date();
        if (action === 'delete') {
          const usedIn = await db.collection('defaults').countDocuments({ typeId: id }, { session });
          if (usedIn) throw new Error('A default type that has been used cannot be deleted; deactivate it instead.');
          await collection.deleteOne({ id }, { session });
          await db.collection('audit_log').insertOne({
            action: 'Default type deleted', collection: 'default_types', by: email, at: now,
            category: 'important', actor: actor.memberId, before: { id, name: type.name },
          }, { session });
          return { deleted: true };
        }

        const update: Record<string, unknown> = {};
        if (action === 'edit') {
          const name = normalizeName(payload?.name);
          const description = String(payload?.description ?? '').trim();
          const appliesTo = normalizeAppliesTo(payload?.appliesTo || 'all_companies');
          if (!name) throw new Error('Default type name is required.');
          if (!validAppliesTo(appliesTo)) throw new Error('Applies to must be all companies, Dream, Core, or Normal.');
          const duplicate = await collection.findOne(
            { nameNormalized: normalizeDefaultTypeName(name), id: { $ne: id } },
            { session, projection: { _id: 1 } },
          );
          if (duplicate) throw new Error('A default type with this name already exists.');
          Object.assign(update, { name, nameNormalized: normalizeDefaultTypeName(name), description, appliesTo });
        } else if (action === 'activate' || action === 'deactivate') {
          update.active = action === 'activate';
        } else {
          throw new Error('Choose edit, activate, deactivate, or delete.');
        }
        Object.assign(update, { updatedBy: actor.memberId, updatedByName: actor.name, updatedAt: now });
        await collection.updateOne({ id }, { $set: update }, { session });
        await db.collection('audit_log').insertOne({
          action: `Default type ${action === 'edit' ? 'edited' : action === 'activate' ? 'reactivated' : 'deactivated'}`,
          collection: 'default_types', by: email, at: now, category: 'important', actor: actor.memberId,
          before: { id, name: type.name, description: type.description, appliesTo: type.appliesTo, active: type.active },
          after: { id, ...update },
        }, { session });
        return { updated: true };
      });
      return NextResponse.json(result);
    } catch (error) {
      const message = error instanceof Error ? error.message : '';
      if (['Default type not found.', 'A default type with this name already exists.', 'A default type that has been used cannot be deleted; deactivate it instead.'].includes(message) ||
        message.startsWith('Default type name') || message.startsWith('Applies to') || message.startsWith('Choose edit') ||
        message.startsWith('Your active')) {
        return NextResponse.json({ message }, { status: message === 'Default type not found.' ? 404 : 400 });
      }
      console.error('Unable to update default type.', error);
      return NextResponse.json({ message: getApiErrorMessage(error, 'Unable to update default type.') }, { status: 500 });
    }
  }
