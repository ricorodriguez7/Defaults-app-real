import type { ClientSession, Db } from 'mongodb';

export type AutoAddResult = {
  added: number;
  defaultsAdded: number;
  failures: string[];
  message: string;
};

export function getIndiaDate(date: Date) {
  const parts = new Intl.DateTimeFormat('en-GB', {
    timeZone: 'Asia/Kolkata',
    year: 'numeric',
    month: '2-digit',
    day: '2-digit',
  }).formatToParts(date);
  const values = Object.fromEntries(parts.map((part) => [part.type, part.value]));
  return `${values.year}-${values.month}-${values.day}`;
}

export async function getCurrentDc(db: Db, session?: ClientSession) {
  const today = getIndiaDate(new Date());
  return db.collection('dcs').findOne(
    { status: 'Planned', date: { $gte: today } },
    { sort: { date: 1, createdAt: 1 }, ...(session ? { session } : {}) },
  );
}

export async function syncCurrentDcAutoAdd(
  db: Db,
  session: ClientSession,
  memberId: string,
  email: string,
  now: Date,
): Promise<AutoAddResult> {
  const dcs = db.collection('dcs');
  const entries = db.collection('dc_entries');
  const defaults = db.collection('defaults');
  const currentDc = await getCurrentDc(db, session);

  const activeDcs = await dcs.find(
    { status: 'Planned' },
    { projection: { _id: 0, id: 1, status: 1 }, session },
  ).toArray();
  const activeDcIds = activeDcs.map((dc) => dc.id);
  await entries.updateMany(
    {
      openStudentRollNo: { $type: 'string' },
      $or: [
        { dcId: { $nin: activeDcIds } },
        { removed: true },
        { excluded: true },
      ],
    },
    { $unset: { openStudentRollNo: '' } },
    { session },
  );
  const activeEntries = activeDcIds.length
    ? await entries.find({ dcId: { $in: activeDcIds }, removed: { $ne: true } }, { session }).toArray()
    : [];
  const activeAssignments = new Map<string, string>();
  const activeDcByDefaultId = new Map<string, string>();
  for (const entry of activeEntries) {
    const rollNo = String(entry.rollNoNormalized ?? entry.rollNo ?? '').trim().toUpperCase();
    if (rollNo && entry.excluded !== true) activeAssignments.set(rollNo, entry.dcId);
    for (const item of entry.decisions ?? []) activeDcByDefaultId.set(item.defaultId, entry.dcId);
  }

  const activeDcIdSet = new Set(activeDcIds);
  const currentEntries = currentDc
    ? await entries.find({ dcId: currentDc.id }, { session }).toArray()
    : [];
  const entryByRollNo = new Map<string, (typeof currentEntries)[number]>();
  for (const entry of currentEntries) {
    const rollNo = String(entry.rollNoNormalized ?? entry.rollNo ?? '').trim().toUpperCase();
    if (rollNo) entryByRollNo.set(rollNo, entry);
  }

  const possiblyStaleDefaults = await defaults.find({
    status: 'logged',
    $or: [
      { dcStatus: 'Scheduled' },
      { dcId: { $exists: true, $ne: null } },
      { id: { $in: Array.from(activeDcByDefaultId.keys()) } },
    ],
  }, { session }).toArray();
  const dcIds = Array.from(new Set(possiblyStaleDefaults.map((record) => record.dcId).filter(Boolean)));
  const linkedDcs = dcIds.length
    ? await dcs.find({ id: { $in: dcIds } }, { projection: { _id: 0, id: 1, status: 1 }, session }).toArray()
    : [];
  const dcById = new Map(linkedDcs.map((dc) => [dc.id, dc]));

  for (const record of possiblyStaleDefaults) {
    const linkedDc = record.dcId ? dcById.get(record.dcId) : undefined;
    if (linkedDc && activeDcIdSet.has(linkedDc.id)) continue;
    const activeAssignmentDcId = activeDcByDefaultId.get(record.id);
    if (activeAssignmentDcId) {
      if (['Scheduled', 'Not scheduled', undefined, null].includes(record.dcStatus)) {
        await defaults.updateOne(
          { id: record.id },
          { $set: { dcId: activeAssignmentDcId, dcStatus: 'Scheduled', updatedAt: now } },
          { session },
        );
      }
      continue;
    }

    const hasDcLink = record.dcId !== undefined && record.dcId !== null;
    const linkedEntry = hasDcLink && record.dcId
      ? await entries.findOne({
        dcId: record.dcId,
        $or: [
          { rollNoNormalized: String(record.rollNoNormalized ?? '').toUpperCase() },
          { rollNo: record.rollNo },
        ],
      }, { session })
      : null;
    const decision = (linkedEntry?.decisions ?? []).find(
      (item: { defaultId: string }) => item.defaultId === record.id,
    )?.decision;
    const hasDecision = linkedEntry?.attendance === 'Present' && decision && decision !== 'Undecided';
    const terminalDc = !linkedDc || linkedDc.status === 'Completed' || linkedDc.status === 'Finalised' || linkedDc.status === 'Cancelled';
    const orphanedScheduledStatus = !record.dcId && record.dcStatus === 'Scheduled';
    const clearTerminalLink = hasDcLink && terminalDc && !hasDecision;
    if (clearTerminalLink || orphanedScheduledStatus) {
      await defaults.updateOne(
        { id: record.id },
        {
          ...(record.dcStatus === 'Scheduled' || orphanedScheduledStatus
            ? { $set: { dcStatus: 'Not scheduled', updatedAt: now } }
            : { $set: { updatedAt: now } }),
          $unset: { dcId: '' },
        },
        { session },
      );
    }
  }

  const result: AutoAddResult = { added: 0, defaultsAdded: 0, failures: [], message: '' };
  if (!currentDc || currentDc.autoAdd === false) {
    result.message = autoAddMessage(result);
    return result;
  }

  const readyDefaults = await defaults.find({
    status: 'logged',
    dcId: null,
    $or: [
      { dcStatus: 'Not scheduled' },
      { dcStatus: { $exists: false } },
      { dcStatus: null },
    ],
  }, { session }).sort({ date: 1, createdAt: 1 }).toArray();

  const defaultsByRollNo = new Map<string, typeof readyDefaults>();
  for (const record of readyDefaults) {
    const rollNo = String(record.rollNoNormalized ?? record.rollNo ?? '').trim().toUpperCase();
    if (!rollNo) {
      result.failures.push(`Default ${record.id} has no student roll number.`);
      continue;
    }
    const group = defaultsByRollNo.get(rollNo) ?? [];
    group.push(record);
    defaultsByRollNo.set(rollNo, group);
  }

  if (defaultsByRollNo.size) {
    const rollNos = Array.from(defaultsByRollNo.keys());
    const students = await db.collection('students').find({
      $or: [
        { rollNoNormalized: { $in: rollNos } },
        { rollNo: { $in: readyDefaults.map((record) => record.rollNo) } },
      ],
    }, { projection: { _id: 0, rollNo: 1, rollNoNormalized: 1, name: 1 }, session }).toArray();
    const studentByRollNo = new Map(
      students.map((student) => [
        String(student.rollNoNormalized ?? student.rollNo).trim().toUpperCase(),
        student,
      ]),
    );

    for (const rollNoNormalized of Array.from(defaultsByRollNo.keys())) {
      const studentDefaults = defaultsByRollNo.get(rollNoNormalized) ?? [];
      const currentEntry = entryByRollNo.get(rollNoNormalized);
      if (currentEntry?.removed === true || currentEntry?.excluded === true) continue;

      const assignedDcId = activeAssignments.get(rollNoNormalized);
      if (assignedDcId && assignedDcId !== currentDc.id) continue;

      const student = studentByRollNo.get(rollNoNormalized);
      if (!student) {
        result.failures.push(`${studentDefaults[0].rollNo}: student is not in the roster.`);
        continue;
      }

      let entry: (typeof currentEntries)[number] | null = currentEntry ?? null;
      let inserted = false;
      if (!entry) {
        const newEntry = {
          dcId: currentDc.id,
          rollNo: student.rollNo,
          rollNoNormalized,
          attendance: 'Pending',
          absentReason: '',
          remarks: '',
          decisions: [],
          source: 'auto',
          excluded: false,
          openStudentRollNo: rollNoNormalized,
          createdAt: now,
          createdBy: memberId,
          removed: false,
          updatedAt: now,
        };
        const upsertResult = await entries.updateOne(
          { dcId: currentDc.id, rollNo: student.rollNo },
          { $setOnInsert: newEntry },
          { upsert: true, session },
        );
        inserted = upsertResult.upsertedCount === 1;
        entry = await entries.findOne({
          dcId: currentDc.id,
          $or: [{ rollNoNormalized }, { rollNo: student.rollNo }],
        }, { session });
        if (!entry) throw new Error(`Unable to create the DC entry for ${student.rollNo}.`);
        if (entry.removed === true || entry.excluded === true) continue;
        if (entry.openStudentRollNo) activeAssignments.set(rollNoNormalized, currentDc.id);
        else {
          await entries.updateOne(
            { _id: entry._id },
            { $set: { openStudentRollNo: rollNoNormalized, removed: false, excluded: false, updatedAt: now } },
            { session },
          );
          entry.openStudentRollNo = rollNoNormalized;
          entry.removed = false;
          entry.excluded = false;
        }
        if (inserted) result.added += 1;
      }
      if (!entry.openStudentRollNo) {
        await entries.updateOne(
          { _id: entry._id },
          { $set: { openStudentRollNo: rollNoNormalized, updatedAt: now } },
          { session },
        );
        entry.openStudentRollNo = rollNoNormalized;
      }
      activeAssignments.set(rollNoNormalized, currentDc.id);

      const existingDefaultIds = new Set(
        (entry.decisions ?? []).map((item: { defaultId: string }) => item.defaultId),
      );
      const newlyAssigned = studentDefaults.filter((record) => !existingDefaultIds.has(record.id));
      if (!newlyAssigned.length) continue;

      const decisions = [
        ...(entry.decisions ?? []),
        ...newlyAssigned.map((record) => ({ defaultId: record.id, decision: 'Undecided' })),
      ];
      await entries.updateOne(
        { _id: entry._id },
        { $set: { decisions, updatedAt: now } },
        { session },
      );
      await defaults.updateMany(
        { id: { $in: newlyAssigned.map((record) => record.id) } },
        { $set: { dcId: currentDc.id, dcStatus: 'Scheduled', updatedAt: now, updatedByMemberId: memberId } },
        { session },
      );

      for (const record of newlyAssigned) {
        await db.collection('default_events').updateOne(
          { eventId: `${record.id}:dc_auto_added:${currentDc.id}` },
          {
            $setOnInsert: {
              eventId: `${record.id}:dc_auto_added:${currentDc.id}`,
              defaultId: record.id,
              rollNo: record.rollNo,
              type: 'dc_auto_added',
              memberId,
              enteredByEmail: email,
              at: now,
              reason: 'Auto-added to current DC',
              note: '',
              dcReference: currentDc.date,
              dcId: currentDc.id,
              changes: { dcStatus: { from: record.dcStatus ?? 'Not scheduled', to: 'Scheduled' } },
            },
          },
          { upsert: true, session },
        );
      }

      result.defaultsAdded += newlyAssigned.length;
    }
  }

  if (result.added || result.defaultsAdded) {
    await db.collection('audit_log').insertOne({
      action: 'DC auto-add synced',
      collection: 'dc_entries',
      by: email,
      at: now,
      after: { dcId: currentDc.id, studentsAdded: result.added, defaultsAdded: result.defaultsAdded },
    }, { session });
  }

  if (result.failures.length) {
    console.error('Automatic DC add could not schedule all ready students.', {
      dcId: currentDc.id,
      failures: result.failures,
    });
  }
  result.message = autoAddMessage(result);
  return result;
}

export function autoAddMessage(result: AutoAddResult) {
  if (!result.failures.length) return `${result.added} students added automatically.`;
  const reasons = result.failures.slice(0, 3).join('; ');
  const remaining = result.failures.length > 3 ? `; and ${result.failures.length - 3} more` : '';
  return `Added ${result.added} students; ${result.failures.length} could not be added: ${reasons}${remaining}`;
}
