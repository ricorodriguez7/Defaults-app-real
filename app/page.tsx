'use client';

import { ChangeEvent, FormEvent, useCallback, useEffect, useMemo, useState } from 'react';
import { signIn, signOut, useSession } from 'next-auth/react';
import * as XLSX from 'xlsx';
import { downloadWorkbook, MAX_EXCEL_BYTES, MAX_EXCEL_FILE_MESSAGE } from '@/lib/excel-download';
import DCSummaryTab from './dc-summary';
import DCTab from './dc-tab';
import DashboardDCTab from './dashboard-dc-tab';

type Student = {
  rollNo: string;
  name: string;
  batch?: 'junior' | 'senior';
  batchLabel?: string;
  numberOfDefaults?: number;
  numberOfRemovedDefaults?: number;
  lockPenaltyStatus?: string;
  studentLocked?: boolean;
};

type Member = {
  memberId: string;
  name: string;
  email?: string;
  team: 'placement_cell' | 'summer_internship' | 'unassigned' | 'imported';
  active: boolean;
  createdAt?: string;
  createdByEmail?: string;
};

type DefaultRecord = {
  id: string;
  rollNo: string;
  company: string;
  typeId: string;
  typeName: string;
  date: string;
  notes: string;
  status: 'logged' | 'locked' | 'removed';
  loggedByMemberId: string;
  loggedByName: string;
  team: string;
  enteredByEmail: string;
  createdAt: string;
  updatedAt: string;
  updatedByMemberId: string;
  updatedByName: string;
  removedAt?: string;
  removeReason?: string;
  removedInDcId?: string;
  lockedAt?: string;
  dcReference?: string;
  dcDate?: string;
  dcStatus?: 'Not scheduled' | 'Scheduled' | 'To discuss' | 'Deferred' | 'Decided';
  discussion?: {
    reason: string;
    discussionDate?: string;
    raisedByName?: string;
    notes?: Array<{ byName: string; at: string; text: string }>;
  };
  deferral?: {
    until: string;
    reason: string;
    history?: Array<{ until: string; reason: string; at: string; byName?: string }>;
  };
};

type DefaultType = { id: string; name: string; description?: string; appliesTo?: string; active?: boolean; usedIn?: number; lastChangedBy?: string };

type EventRecord = {
  eventId: string;
  defaultId: string | null;
  rollNo: string;
  type: string;
  memberId: string;
  memberName: string;
  team: string;
  enteredByEmail: string;
  at: string;
  reason?: string;
  note?: string;
  dcReference?: string;
  changes?: Record<string, { from: unknown; to: unknown }>;
  studentName?: string;
  studentBatch?: 'junior' | 'senior';
  category?: 'important' | 'dc' | 'members' | 'imports';
  actor?: 'system' | 'member';
  actorLabel?: string;
};

type StudentHistoryDefault = DefaultRecord & {
  decisionHistory: Array<{
    dcId: string;
    dcDate: string;
    dcTime: string;
    decision: string;
    until: string;
    remarks: string;
  }>;
};

type ViewName = 'defaults' | 'default-import' | 'students' | 'dc' | 'members' | 'activity' | 'settings';
type DashboardTab = 'overview' | 'dc';
type SettingsSubView = 'data' | 'types';
type ResetCategory = 'defaults' | 'dcs' | 'importBatches' | 'auditLog' | 'students' | 'types' | 'members';
type DefaultImportIssue = { row: number; column: string; reason: string; severity: 'warning' | 'error' };
type DefaultImportPreview = {
  counts: { rowsRead: number; ready: number; warnings: number; errors: number };
  issues: DefaultImportIssue[];
  rejectedRows: Array<Record<string, unknown>>;
  imported?: number;
  skipped?: number;
  stoppedForErrors?: boolean;
};
type ImportBatch = {
  batchId: string;
  filename: string;
  importedBy: string;
  importedAt: string;
  rowsRead: number;
  imported: number;
  skipped: number;
  errors: number;
  undoneAt?: string;
  undoBlockReason: string;
};

const resetCategoryLabels: Array<{ key: ResetCategory; label: string }> = [
  { key: 'defaults', label: 'Defaults and their history' },
  { key: 'dcs', label: 'DCs and DC entries' },
  { key: 'importBatches', label: 'Import batches' },
  { key: 'auditLog', label: 'Audit log' },
  { key: 'students', label: 'Students' },
  { key: 'types', label: 'Default types and penalty rules' },
  { key: 'members', label: 'Members (your own account is preserved)' },
];

const normalizeRollNo = (value: string) => String(value ?? '').trim().toUpperCase();
const teamLabel = (team: string) => team === 'placement_cell' ? 'Placement cell' : team === 'summer_internship' ? 'Summer internship' : team;
const isSelectableTeam = (team: string): team is 'placement_cell' | 'summer_internship' =>
  team === 'placement_cell' || team === 'summer_internship';
const autoAddNotice = (result?: { added?: number; failures?: string[]; message?: string }) =>
  result && ((result.added ?? 0) > 0 || (result.failures?.length ?? 0) > 0) ? ` ${result.message ?? ''}` : '';

async function responseJson(response: Response) {
  const data = await response.json();
  if (!response.ok) throw new Error(data.message || `Request failed (${response.status}).`);
  return data;
}

function parseSheet(file: File) {
  if (file.size > MAX_EXCEL_BYTES) throw new Error(MAX_EXCEL_FILE_MESSAGE);
  return file.arrayBuffer().then((buffer) => {
    const workbook = XLSX.read(buffer, { type: 'array' });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    return XLSX.utils.sheet_to_json<Record<string, unknown>>(sheet, { defval: '' });
  });
}

function jsonBodyWithinLimit(value: unknown) {
  const body = JSON.stringify(value);
  if (new TextEncoder().encode(body).byteLength > MAX_EXCEL_BYTES) throw new Error(MAX_EXCEL_FILE_MESSAGE);
  return body;
}

export default function HomePage() {
  const { data: session, status: sessionStatus } = useSession();
  const accountEmail = session?.user?.email?.toLowerCase() ?? '';
  const [view, setView] = useState<ViewName>('defaults');
  const [dashboardTab, setDashboardTab] = useState<DashboardTab>('overview');
  const [dashboardCounts, setDashboardCounts] = useState({ total: 0, logged: 0, locked: 0, removed: 0, toDiscuss: 0, deferred: 0 });
  const [settingsSubView, setSettingsSubView] = useState<SettingsSubView>('data');
  const [dcSubView, setDcSubView] = useState<'planner' | 'summary'>('planner');
  const [sidebarOpen, setSidebarOpen] = useState(false);
  const [sidebarCollapsed, setSidebarCollapsed] = useState(false);
  const [dashboardLoading, setDashboardLoading] = useState(false);
  const [students, setStudents] = useState<Student[]>([]);
  const [defaults, setDefaults] = useState<DefaultRecord[]>([]);
  const [members, setMembers] = useState<Member[]>([]);
  const [defaultTypes, setDefaultTypes] = useState<DefaultType[]>([]);
  const [rollNumber, setRollNumber] = useState('');
  const [company, setCompany] = useState('');
  const [typeId, setTypeId] = useState('');
  const [date, setDate] = useState(new Date().toISOString().slice(0, 10));
  const [notes, setNotes] = useState('');
  const [loggedByMemberId, setLoggedByMemberId] = useState('');
  const [selectedStudent, setSelectedStudent] = useState<Student | null>(null);
  const [lookupError, setLookupError] = useState('');
  const [statusMessage, setStatusMessage] = useState('');
  const [isSubmitting, setIsSubmitting] = useState(false);
  const [showRemoved, setShowRemoved] = useState(false);
  const [detail, setDetail] = useState<{ record: DefaultRecord; events: EventRecord[] } | null>(null);
  const [editValues, setEditValues] = useState({ company: '', typeId: '', date: '', notes: '', dcReference: '', loggedByMemberId: '' });
  const [editReason, setEditReason] = useState('');
  const [dcNote, setDcNote] = useState('');
  const [discussionReason, setDiscussionReason] = useState('');
  const [discussionDate, setDiscussionDate] = useState('');
  const [studentSearch, setStudentSearch] = useState('');
  const [defaultSearch, setDefaultSearch] = useState('');
  const [defaultStatusFilter, setDefaultStatusFilter] = useState('active');
  const [studentDetail, setStudentDetail] = useState<{ student: Student; defaults: StudentHistoryDefault[] } | null>(null);
  const [activity, setActivity] = useState<EventRecord[]>([]);
  const [activityHasMore, setActivityHasMore] = useState(false);
  const [activityFilters, setActivityFilters] = useState({ memberId: '', team: '', type: '', from: '', to: '', rollNo: '', batch: '', category: 'important', skip: 0 });
  const [activityChip, setActivityChip] = useState<'important' | 'dc' | 'members' | 'imports' | 'all'>('important');
  const [expandedActivity, setExpandedActivity] = useState<string[]>([]);
  const [batchFilter, setBatchFilter] = useState('');
  const [studentBatchFilter, setStudentBatchFilter] = useState('');
  const [rosterSearch, setRosterSearch] = useState('');
  const [selectedStudentRollNos, setSelectedStudentRollNos] = useState<string[]>([]);
  const [bulkStudentBatch, setBulkStudentBatch] = useState('');
  const [memberSearch, setMemberSearch] = useState('');
  const [memberTeamFilter, setMemberTeamFilter] = useState('');
  const [memberStatusFilter, setMemberStatusFilter] = useState('all');
  const [memberForm, setMemberForm] = useState<{ memberId?: string; name: string; team: string; email: string } | null>(null);
  const [memberReason, setMemberReason] = useState('');
  const [memberStatusTarget, setMemberStatusTarget] = useState<{ member: Member; action: 'deactivate' | 'reactivate' } | null>(null);
  const [memberStatusReason, setMemberStatusReason] = useState('');
  const [defaultActionTarget, setDefaultActionTarget] = useState<'lock' | 'unlock' | 'restore' | null>(null);
  const [defaultActionReason, setDefaultActionReason] = useState('');
  const [removeDefaultTarget, setRemoveDefaultTarget] = useState<DefaultRecord | null>(null);
  const [removeDefaultReason, setRemoveDefaultReason] = useState('');
  const [removeDcReference, setRemoveDcReference] = useState('');
  const [studentLockAction, setStudentLockAction] = useState<'lock' | 'unlock' | null>(null);
  const [studentLockReason, setStudentLockReason] = useState('');
  const [bulkRows, setBulkRows] = useState<Record<string, unknown>[]>([]);
  const [bulkPreview, setBulkPreview] = useState<{
    counts: { add: number; update: number; skip: number };
    rows: Array<{ row: number; action: 'add' | 'update' | 'skip'; reason: string; name: string; team: string; email: string }>;
  } | null>(null);
  const [defaultImportFile, setDefaultImportFile] = useState<File | null>(null);
  const [defaultImportPreview, setDefaultImportPreview] = useState<DefaultImportPreview | null>(null);
  const [defaultImportCreateTypes, setDefaultImportCreateTypes] = useState(false);
  const [defaultImportStopOnError, setDefaultImportStopOnError] = useState(false);
  const [defaultImportBusy, setDefaultImportBusy] = useState(false);
  const [importBatches, setImportBatches] = useState<ImportBatch[]>([]);
  const [resetEnabled, setResetEnabled] = useState(false);
  const [resetCounts, setResetCounts] = useState<Record<ResetCategory, number> | null>(null);
  const [resetSelected, setResetSelected] = useState<ResetCategory[]>(['defaults', 'dcs', 'importBatches']);
  const [resetConfirmation, setResetConfirmation] = useState('');
  const [backupToken, setBackupToken] = useState('');
  const [resetBusy, setResetBusy] = useState(false);
  const [resetResult, setResetResult] = useState<{ removed?: Record<string, number>; errors?: string[] } | null>(null);
  const [defaultTypeForm, setDefaultTypeForm] = useState<{ id?: string; name: string; description: string; appliesTo: string } | null>(null);
  const [managedTypes, setManagedTypes] = useState<DefaultType[]>([]);
  const [typeDeleteTarget, setTypeDeleteTarget] = useState<DefaultType | null>(null);

  const studentsByRollNo = useMemo(
    () => new Map(students.map((student) => [normalizeRollNo(student.rollNo), student])),
    [students],
  );
  const activeDefaults = useMemo(() => {
    const query = defaultSearch.trim().toLowerCase();
    return defaults.filter((record) =>
      (showRemoved || record.status !== 'removed') &&
      (defaultStatusFilter === 'all' || (defaultStatusFilter === 'active'
        ? record.status !== 'removed'
        : record.status === defaultStatusFilter)) &&
      (!batchFilter || (batchFilter === 'not_set'
        ? !studentsByRollNo.get(normalizeRollNo(record.rollNo))?.batch
        : studentsByRollNo.get(normalizeRollNo(record.rollNo))?.batch === batchFilter)) &&
      (!query || [record.rollNo, record.company, record.typeName, record.loggedByName].some((value) => value.toLowerCase().includes(query))),
    );
  }, [defaults, showRemoved, defaultSearch, defaultStatusFilter, batchFilter, studentsByRollNo]);
  const includeRemovedDefaults = showRemoved || defaultStatusFilter === 'removed' || defaultStatusFilter === 'all';
  const selectableMembers = useMemo(
    () => members.filter((member) => member.active && isSelectableTeam(member.team)),
    [members],
  );
  const signedInMember = useMemo(
    () => selectableMembers.find((member) => member.email?.toLowerCase() === accountEmail) ?? null,
    [selectableMembers, accountEmail],
  );
  const stats = dashboardCounts;
  const matchedStudent = useMemo(
    () => students.find((student) => normalizeRollNo(student.rollNo) === normalizeRollNo(studentSearch)),
    [students, studentSearch],
  );
  const filteredStudents = useMemo(() => students.filter((student) =>
    (!studentBatchFilter || (studentBatchFilter === 'not_set' ? !student.batch : student.batch === studentBatchFilter)) &&
    (!rosterSearch.trim() || `${student.name} ${student.rollNo}`.toLowerCase().includes(rosterSearch.trim().toLowerCase())),
  ), [students, studentBatchFilter, rosterSearch]);

  const loadDashboard = useCallback(async () => {
    setDashboardLoading(true);
    try {
      const [studentsResponse, defaultsResponse, typesResponse, membersResponse, summaryResponse] = await Promise.all([
        fetch('/api/students', { cache: 'no-store' }),
        fetch(`/api/defaults?includeRemoved=${includeRemovedDefaults}`, { cache: 'no-store' }),
        fetch('/api/default-types', { cache: 'no-store' }),
        fetch('/api/members', { cache: 'no-store' }),
        fetch('/api/dashboard/summary', { cache: 'no-store' }),
      ]);
      const [studentsData, defaultsData, typesData, membersData, summaryData] = await Promise.all([
        responseJson(studentsResponse),
        responseJson(defaultsResponse),
        responseJson(typesResponse),
        responseJson(membersResponse),
        responseJson(summaryResponse),
      ]);
      setStudents(studentsData);
      setDefaults(defaultsData);
      setDashboardCounts(summaryData.counts);
      setDefaultTypes(typesData);
      setMembers(membersData);
      setTypeId((current) => current || typesData[0]?.id || '');
      const selectableMembers = membersData.filter((member: Member) => member.active && isSelectableTeam(member.team));
      const currentMember = selectableMembers.find((member: Member) => member.email?.toLowerCase() === accountEmail);
      setLoggedByMemberId(currentMember?.memberId ?? '');
    } finally {
      setDashboardLoading(false);
    }
  }, [accountEmail, includeRemovedDefaults]);

  useEffect(() => {
    if (sessionStatus === 'authenticated') {
      loadDashboard().catch((error) => setStatusMessage(error instanceof Error ? error.message : 'Unable to load tracker data.'));
    }
  }, [sessionStatus, loadDashboard]);

  async function handleLookup() {
    const trimmed = normalizeRollNo(rollNumber);
    if (!trimmed) {
      setLookupError('Enter a roll number first.');
      setSelectedStudent(null);
      return;
    }
    try {
      const data = await responseJson(await fetch(`/api/students?rollNo=${encodeURIComponent(trimmed)}`, { cache: 'no-store' }));
      if (!data.found) throw new Error('Unknown roll number. Please use a current student roll number.');
      setSelectedStudent(data.student);
      setLookupError('');
    } catch (error) {
      setSelectedStudent(null);
      setLookupError(error instanceof Error ? error.message : 'Unable to look up that student.');
    }
  }

  async function handleSubmit(event: FormEvent<HTMLFormElement>) {
    event.preventDefault();
    if (!selectedStudent) {
      setStatusMessage('Look up a valid student before logging the default.');
      return;
    }
    setIsSubmitting(true);
    setStatusMessage('');
    try {
      const result = await responseJson(await fetch('/api/defaults', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ rollNo: selectedStudent.rollNo, company, typeId, date, notes }),
      }));
      setStatusMessage(`Default logged for ${selectedStudent.name}.${autoAddNotice(result.autoAdd)}`);
      setRollNumber('');
      setCompany('');
      setTypeId(defaultTypes[0]?.id || '');
      setDate(new Date().toISOString().slice(0, 10));
      setNotes('');
      setSelectedStudent(null);
      setLookupError('');
      await loadDashboard();
    } catch (error) {
      setStatusMessage(error instanceof Error ? error.message : 'Unable to log the default.');
    } finally {
      setIsSubmitting(false);
    }
  }

  async function handleRosterUpload(event: ChangeEvent<HTMLInputElement>) {
    const file = event.target.files?.[0];
    if (!file) return;
    try {
      const rows = await parseSheet(file);
      const roster = rows.map((row, index) => ({
        name: String(row.Name ?? row.name ?? '').trim(),
        rollNo: String(row['Roll No'] ?? row['Roll Number'] ?? row['Roll number'] ?? row.rollNo ?? '').trim(),
        batch: String(row.Batch ?? row.batch ?? '').trim(),
        rowNumber: index + 2,
        numberOfDefaults: (row['Number of defaults'] ?? row.numberOfDefaults) === '' ||
          (row['Number of defaults'] ?? row.numberOfDefaults) === undefined
          ? undefined
          : Number(row['Number of defaults'] ?? row.numberOfDefaults),
        lockPenaltyStatus: String(row['Lock / penalty status'] ?? row.lockPenaltyStatus ?? '').trim(),
      })).filter((row) => row.name || row.rollNo || row.batch);
      const body = jsonBodyWithinLimit({ roster });
      await responseJson(await fetch('/api/students', {
        method: 'POST', headers: { 'Content-Type': 'application/json' }, body,
      }));
      setStatusMessage('Student roster updated.');
      await loadDashboard();
    } catch (error) {
      setStatusMessage(error instanceof Error ? error.message : 'Roster upload failed.');
    } finally {
      event.target.value = '';
    }
  }

  async function handleMemberUpload(event: ChangeEvent<HTMLInputElement>) {
    const file = event.target.files?.[0];
    if (!file) return;
    try {
      const rows = await parseSheet(file);
      const uploadedMembers = rows.map((row) => ({
        name: String(row.Name ?? row.name ?? '').trim(),
        email: String(row.Email ?? row.email ?? '').trim(),
        team: String(row.Team ?? row.team ?? '').trim(),
      }));
      const body = jsonBodyWithinLimit({ members: uploadedMembers });
      const preview = await responseJson(await fetch('/api/members/bulk-preview', {
        method: 'POST', headers: { 'Content-Type': 'application/json' }, body,
      }));
      setBulkRows(uploadedMembers);
      setBulkPreview(preview);
      setStatusMessage('Review the member import preview before applying changes.');
    } catch (error) {
      setStatusMessage(error instanceof Error ? error.message : 'Member preview failed.');
    } finally {
      event.target.value = '';
    }
  }

  async function saveMember(event: FormEvent<HTMLFormElement>) {
    event.preventDefault();
    if (!memberForm) return;
    try {
      const editing = Boolean(memberForm.memberId);
      const data = await responseJson(await fetch(
        editing ? `/api/members/${encodeURIComponent(memberForm.memberId!)}` : '/api/members',
        {
          method: editing ? 'PATCH' : 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            ...(editing ? { action: 'edit', reason: memberReason } : {}),
            name: memberForm.name,
            team: memberForm.team,
            email: memberForm.email,
          }),
        },
      ));
      setStatusMessage(editing ? 'Member updated.' : 'Member added.');
      setMemberForm(null);
      setMemberReason('');
      await loadDashboard();
      if (!editing && data.member?.active && isSelectableTeam(data.member.team)) {
        setLoggedByMemberId(data.member.memberId);
      }
    } catch (error) {
      setStatusMessage(error instanceof Error ? error.message : 'Unable to save member.');
    }
  }

  async function changeMemberStatus(member: Member, action: 'deactivate' | 'reactivate') {
    if (action === 'deactivate' && member.email?.toLowerCase() === accountEmail) {
      setStatusMessage('You cannot deactivate your own account.');
      return;
    }
    setMemberStatusTarget({ member, action });
    setMemberStatusReason('');
  }

  async function submitMemberStatus(event: FormEvent<HTMLFormElement>) {
    event.preventDefault();
    if (!memberStatusTarget) return;
    const { member, action } = memberStatusTarget;
    const reason = memberStatusReason.trim();
    if (!reason) return;
    try {
      await responseJson(await fetch(`/api/members/${encodeURIComponent(member.memberId)}`, {
        method: 'PATCH',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ action, reason }),
      }));
      setStatusMessage(`Member ${action}d.`);
      setMemberStatusTarget(null);
      setMemberStatusReason('');
      await loadDashboard();
    } catch (error) {
      setStatusMessage(error instanceof Error ? error.message : `Unable to ${action} member.`);
    }
  }

  async function applyMemberPreview() {
    try {
      const body = jsonBodyWithinLimit({ members: bulkRows });
      const result = await responseJson(await fetch('/api/members/bulk', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body,
      }));
      setStatusMessage(`Member upload complete: ${result.counts.add} added, ${result.counts.update} updated, ${result.counts.skip} skipped.`);
      setBulkRows([]);
      setBulkPreview(null);
      await loadDashboard();
    } catch (error) {
      setStatusMessage(error instanceof Error ? error.message : 'Unable to apply member upload.');
    }
  }

  function downloadMemberTemplate() {
    const workbook = XLSX.utils.book_new();
    const sheet = XLSX.utils.aoa_to_sheet([
      ['Name', 'Team', 'Email'],
      ['Example Member', 'placement_cell', 'member@example.edu'],
      ['Example Intern', 'summer_internship', ''],
    ]);
    XLSX.utils.book_append_sheet(workbook, sheet, 'Members');
    downloadWorkbook(workbook, 'members_sample.xlsx');
  }

  function downloadStudentTemplate() {
    const workbook = XLSX.utils.book_new();
    const sheet = XLSX.utils.aoa_to_sheet([
      ['Roll No', 'Name', 'Batch'],
      ['EXAMPLE001', 'Example Student - replace or delete this row', 'Junior batch'],
    ]);
    XLSX.utils.book_append_sheet(workbook, sheet, 'Students');
    XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet([
      ['Column', 'Required', 'Accepted values'],
      ['Roll No', 'Yes', 'Unique student roll number.'],
      ['Name', 'Yes', 'Student name.'],
      ['Batch', 'No', 'Junior batch or Senior batch. Junior and Senior are also accepted, ignoring case. Blank leaves an existing batch unchanged.'],
    ]), 'Instructions');
    downloadWorkbook(workbook, 'students_sample.xlsx');
  }

  async function loadImportBatches() {
    const batches = await responseJson(await fetch('/api/defaults/import', { cache: 'no-store' }));
    setImportBatches(batches);
  }

  useEffect(() => {
    if (sessionStatus === 'authenticated' && view === 'default-import') {
      loadImportBatches().catch((error) => setStatusMessage(error instanceof Error ? error.message : 'Unable to load import history.'));
    }
  }, [sessionStatus, view]);

  async function previewDefaultImport(file: File, createMissingTypes: boolean) {
    setDefaultImportBusy(true);
    setDefaultImportPreview(null);
    try {
      const form = new FormData();
      form.set('file', file);
      form.set('mode', 'preview');
      form.set('createMissingTypes', String(createMissingTypes));
      form.set('stopOnError', String(defaultImportStopOnError));
      setDefaultImportPreview(await responseJson(await fetch('/api/defaults/import', { method: 'POST', body: form })));
      setStatusMessage('Review the validation results before importing.');
    } catch (error) {
      setStatusMessage(error instanceof Error ? error.message : 'Unable to preview defaults import.');
    } finally {
      setDefaultImportBusy(false);
    }
  }

  async function handleDefaultImportFile(event: ChangeEvent<HTMLInputElement>) {
    const file = event.target.files?.[0];
    event.target.value = '';
    if (!file) return;
    setDefaultImportFile(file);
    await previewDefaultImport(file, defaultImportCreateTypes);
  }

  async function submitDefaultImport() {
    if (!defaultImportFile) return;
    setDefaultImportBusy(true);
    try {
      const form = new FormData();
      form.set('file', defaultImportFile);
      form.set('mode', 'import');
      form.set('createMissingTypes', String(defaultImportCreateTypes));
      form.set('stopOnError', String(defaultImportStopOnError));
      const result: DefaultImportPreview = await responseJson(await fetch('/api/defaults/import', { method: 'POST', body: form }));
      setDefaultImportPreview(result);
      setStatusMessage(result.stoppedForErrors
        ? 'Import stopped because the file contains errors; no rows were imported.'
        : `Import complete: ${result.imported ?? 0} rows imported, ${result.skipped ?? 0} skipped.`);
      await Promise.all([loadImportBatches(), loadDashboard()]);
    } catch (error) {
      setStatusMessage(error instanceof Error ? error.message : 'Unable to import defaults.');
    } finally {
      setDefaultImportBusy(false);
    }
  }

  async function downloadDefaultImportSample() {
    try {
      const response = await fetch('/api/defaults/import/sample', { cache: 'no-store' });
      if (!response.ok) {
        const data = await response.json();
        throw new Error(data.message || `Request failed (${response.status}).`);
      }
      const url = URL.createObjectURL(await response.blob());
      const link = document.createElement('a');
      link.href = url;
      link.download = 'defaults_import_sample.xlsx';
      link.click();
      URL.revokeObjectURL(url);
    } catch (error) {
      setStatusMessage(error instanceof Error ? error.message : 'Unable to download the defaults sample.');
    }
  }

  function downloadImportErrorReport() {
    const rejectedRows = defaultImportPreview?.rejectedRows ?? [];
    if (!rejectedRows.length) return;
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(rejectedRows), 'Rejected rows');
    downloadWorkbook(workbook, 'defaults_import_errors.xlsx');
  }

  async function loadResetCounts() {
    const data = await responseJson(await fetch('/api/data-management', { cache: 'no-store' }));
    setResetEnabled(data.enabled);
    setResetCounts(data.counts);
  }

  async function loadManagedTypes() {
    const data = await responseJson(await fetch('/api/default-types?includeInactive=true', { cache: 'no-store' }));
    setManagedTypes(data);
  }

  useEffect(() => {
    if (sessionStatus === 'authenticated' && view === 'settings' && settingsSubView === 'types') {
      loadManagedTypes().catch((error) => setStatusMessage(error instanceof Error ? error.message : 'Unable to load default types.'));
    }
  }, [sessionStatus, view, settingsSubView]);

  async function saveDefaultType(event: FormEvent<HTMLFormElement>) {
    event.preventDefault();
    if (!defaultTypeForm) return;
    try {
      await responseJson(await fetch('/api/default-types', {
        method: defaultTypeForm.id ? 'PATCH' : 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ ...defaultTypeForm, action: 'edit' }),
      }));
      setDefaultTypeForm(null);
      setStatusMessage(defaultTypeForm.id ? 'Default type updated.' : 'Default type added.');
      await Promise.all([loadManagedTypes(), loadDashboard()]);
    } catch (error) {
      setStatusMessage(error instanceof Error ? error.message : 'Unable to save default type.');
    }
  }

  async function changeDefaultTypeStatus(type: DefaultType, action: 'activate' | 'deactivate' | 'delete') {
    try {
      await responseJson(await fetch('/api/default-types', {
        method: 'PATCH',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ id: type.id, action }),
      }));
      setTypeDeleteTarget(null);
      setStatusMessage(action === 'delete' ? 'Default type deleted.' : `Default type ${action === 'activate' ? 'reactivated' : 'deactivated'}.`);
      await Promise.all([loadManagedTypes(), loadDashboard()]);
    } catch (error) {
      setStatusMessage(error instanceof Error ? error.message : 'Unable to update default type.');
    }
  }

  useEffect(() => {
    if (sessionStatus === 'authenticated' && view === 'settings') {
      loadResetCounts().catch((error) => setStatusMessage(error instanceof Error ? error.message : 'Unable to load data management.'));
    }
  }, [sessionStatus, view]);

  async function exportEverything() {
    try {
      const response = await fetch('/api/data-management/export', { cache: 'no-store' });
      if (!response.ok) {
        const data = await response.json();
        throw new Error(data.message || `Request failed (${response.status}).`);
      }
      const token = response.headers.get('X-Data-Backup-Token');
      if (!token) throw new Error('The server did not confirm the backup export. Reset remains locked.');
      const url = URL.createObjectURL(await response.blob());
      const link = document.createElement('a');
      link.href = url;
      link.download = 'default-tracker-backup.xlsx';
      link.click();
      URL.revokeObjectURL(url);
      setBackupToken(token);
      setResetResult(null);
      setStatusMessage('Backup exported. The reset controls are unlocked for 30 minutes.');
    } catch (error) {
      setStatusMessage(error instanceof Error ? error.message : 'Unable to export all data.');
    }
  }

  function chooseResetPreset(preset: 'test' | 'full') {
    setResetSelected(preset === 'test'
      ? ['defaults', 'dcs', 'importBatches']
      : resetCategoryLabels.map(({ key }) => key));
    setResetResult(null);
  }

  async function performDataReset() {
    setResetBusy(true);
    try {
      const result = await responseJson(await fetch('/api/data-management', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ selected: resetSelected, confirmation: resetConfirmation, backupToken }),
      }));
      setResetResult(result);
      setBackupToken('');
      setResetConfirmation('');
      setStatusMessage(result.errors?.length ? 'Reset completed with errors; review the per-collection results.' : result.message);
      await Promise.all([loadResetCounts(), loadDashboard()]);
    } catch (error) {
      setStatusMessage(error instanceof Error ? error.message : 'Unable to clear selected data.');
    } finally {
      setResetBusy(false);
    }
  }

  async function undoImport(batch: ImportBatch) {
    try {
      const result = await responseJson(await fetch('/api/defaults/import', {
        method: 'DELETE',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ batchId: batch.batchId }),
      }));
      setStatusMessage(`Import undone: ${result.removed} defaults and their history removed.`);
      await Promise.all([loadImportBatches(), loadDashboard()]);
    } catch (error) {
      setStatusMessage(error instanceof Error ? error.message : 'Unable to undo this import.');
    }
  }

  const filteredMembers = useMemo(() => {
    const query = memberSearch.trim().toLowerCase();
    return members.filter((member) =>
      (!query || member.name.toLowerCase().includes(query) || (member.email || '').toLowerCase().includes(query)) &&
      (!memberTeamFilter || member.team === memberTeamFilter) &&
      (memberStatusFilter === 'all' || (memberStatusFilter === 'active' ? member.active : !member.active)),
    );
  }, [members, memberSearch, memberTeamFilter, memberStatusFilter]);
  async function openDefault(record: DefaultRecord) {
    try {
      const data = await responseJson(await fetch(`/api/defaults/${encodeURIComponent(record.id)}`, { cache: 'no-store' }));
      setDetail(data);
      setEditValues({
        company: data.record.company,
        typeId: data.record.typeId,
        date: data.record.date,
        notes: data.record.notes,
        dcReference: data.record.dcReference ?? '',
        loggedByMemberId: data.record.loggedByMemberId,
      });
      setEditReason('');
      setDcNote('');
      setDiscussionReason('');
      setDiscussionDate('');
    } catch (error) {
      setStatusMessage(error instanceof Error ? error.message : 'Unable to load default history.');
    }
  }

  async function updateDefault(action: 'edit' | 'lock' | 'unlock' | 'restore') {
    if (!detail) return;
    if (action !== 'edit') {
      setDefaultActionTarget(action);
      setDefaultActionReason('');
      return;
    }
    try {
      const reason = editReason.trim();
      if (!reason) return;
      const body = { action, reason, ...editValues };
      const data = await responseJson(await fetch(`/api/defaults/${encodeURIComponent(detail.record.id)}`, {
        method: 'PATCH', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(body),
      }));
      setEditReason('');
      await openDefault({ ...detail.record, ...data.updated });
      await loadDashboard();
      setStatusMessage(`Default ${action === 'edit' ? 'updated' : `${action}ed`}.${autoAddNotice(data.autoAdd)}`);
    } catch (error) {
      setStatusMessage(error instanceof Error ? error.message : 'Unable to update default.');
    }
  }

  async function submitDefaultAction(event: FormEvent<HTMLFormElement>) {
    event.preventDefault();
    if (!detail || !defaultActionTarget) return;
    try {
      const data = await responseJson(await fetch(`/api/defaults/${encodeURIComponent(detail.record.id)}`, {
        method: 'PATCH',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ action: defaultActionTarget, reason: defaultActionReason.trim() }),
      }));
      const action = defaultActionTarget;
      setDefaultActionTarget(null);
      setDefaultActionReason('');
      await openDefault({ ...detail.record, ...data.updated });
      await loadDashboard();
      setStatusMessage(`Default ${action === 'restore' ? 'restored' : `${action}ed`}.${autoAddNotice(data.autoAdd)}`);
    } catch (error) {
      setStatusMessage(error instanceof Error ? error.message : 'Unable to update default.');
    }
  }

  async function removeDefault(record: DefaultRecord) {
    setRemoveDefaultTarget(record);
    setRemoveDefaultReason('');
    setRemoveDcReference('');
  }

  async function submitRemoveDefault(event: FormEvent<HTMLFormElement>) {
    event.preventDefault();
    if (!removeDefaultTarget) return;
    const record = removeDefaultTarget;
    try {
      const result = await responseJson(await fetch(`/api/defaults/${encodeURIComponent(record.id)}`, {
        method: 'DELETE',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ reason: removeDefaultReason.trim(), dcReference: removeDcReference.trim() }),
      }));
      setRemoveDefaultTarget(null);
      setDetail(null);
      await loadDashboard();
      setStatusMessage(`Default removed; its history is retained.${autoAddNotice(result.autoAdd)}`);
    } catch (error) {
      setStatusMessage(error instanceof Error ? error.message : 'Unable to remove default.');
    }
  }

  async function saveDcNote() {
    if (!detail) return;
    try {
      await responseJson(await fetch(`/api/defaults/${encodeURIComponent(detail.record.id)}/events`, {
        method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ note: dcNote }),
      }));
      setDcNote('');
      await openDefault(detail.record);
    } catch (error) {
      setStatusMessage(error instanceof Error ? error.message : 'Unable to add DC note.');
    }
  }

  async function flagDefaultForDiscussion(event: FormEvent<HTMLFormElement>) {
    event.preventDefault();
    if (!detail) return;
    try {
      const result = await responseJson(await fetch(`/api/defaults/${encodeURIComponent(detail.record.id)}`, {
        method: 'PATCH',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ action: 'flag_discussion', reason: discussionReason.trim(), discussionDate }),
      }));
      setDiscussionReason('');
      setDiscussionDate('');
      await Promise.all([openDefault(detail.record), loadDashboard()]);
      setStatusMessage(`Default flagged for team discussion.${autoAddNotice(result.autoAdd)}`);
    } catch (error) {
      setStatusMessage(error instanceof Error ? error.message : 'Unable to flag default for discussion.');
    }
  }

  async function loadStudentDetail() {
    if (!matchedStudent) {
      setStatusMessage('Enter a valid roll number from the roster.');
      return;
    }
    try {
      setStudentDetail(await responseJson(await fetch(`/api/students/${encodeURIComponent(matchedStudent.rollNo)}`, { cache: 'no-store' })));
      setView('students');
    } catch (error) {
      setStatusMessage(error instanceof Error ? error.message : 'Unable to load student detail.');
    }
  }

  async function changeStudentBatch(rollNo: string, batch: 'junior' | 'senior' | '') {
    try {
      await responseJson(await fetch(`/api/students/${encodeURIComponent(rollNo)}`, {
        method: 'PATCH', headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ action: 'set_batch', batch }),
      }));
      setStudents((current) => current.map((student) => student.rollNo === rollNo
        ? { ...student, batch: batch || undefined, batchLabel: batch === 'junior' ? 'Junior batch' : batch === 'senior' ? 'Senior batch' : 'Not set' }
        : student));
      if (studentDetail?.student.rollNo === rollNo) {
        setStudentDetail({ ...studentDetail, student: { ...studentDetail.student, batch: batch || undefined, batchLabel: batch === 'junior' ? 'Junior batch' : batch === 'senior' ? 'Senior batch' : 'Not set' } });
      }
      setStatusMessage(`Updated ${rollNo} to ${batch === 'junior' ? 'Junior batch' : 'Senior batch'}.`);
    } catch (error) {
      setStatusMessage(error instanceof Error ? error.message : 'Unable to update student batch.');
    }
  }

  async function updateSelectedStudentBatches() {
    if (!selectedStudentRollNos.length || !bulkStudentBatch) return;
    try {
      const result = await responseJson(await fetch('/api/students/batch', {
        method: 'POST', headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ rollNos: selectedStudentRollNos, batch: bulkStudentBatch }),
      }));
      setStatusMessage(`Updated ${result.updated} students${result.notFound ? `; ${result.notFound} were not found` : ''}.`);
      setSelectedStudentRollNos([]);
      setBulkStudentBatch('');
      await loadDashboard();
    } catch (error) {
      setStatusMessage(error instanceof Error ? error.message : 'Unable to update selected student batches.');
    }
  }

  async function openStudentFromDc(rollNo: string) {
    try {
      const studentData = await responseJson(await fetch(`/api/students/${encodeURIComponent(rollNo)}`, { cache: 'no-store' }));
      setStudentSearch(rollNo);
      setStudentDetail(studentData);
      setView('students');
    } catch (error) {
      setStatusMessage(error instanceof Error ? error.message : 'Unable to load student history.');
    }
  }

  async function toggleStudentLock() {
    if (!studentDetail) return;
    const action = studentDetail.student.studentLocked ? 'unlock' : 'lock';
    setStudentLockAction(action);
    setStudentLockReason('');
  }

  async function submitStudentLock(event: FormEvent<HTMLFormElement>) {
    event.preventDefault();
    if (!studentDetail || !studentLockAction) return;
    const action = studentLockAction;
    try {
      await responseJson(await fetch(`/api/students/${encodeURIComponent(studentDetail.student.rollNo)}`, {
        method: 'PATCH',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ action, reason: studentLockReason.trim() }),
      }));
      setStudentLockAction(null);
      setStudentLockReason('');
      await loadStudentDetail();
      setStatusMessage(`Student ${action}ed.`);
    } catch (error) {
      setStatusMessage(error instanceof Error ? error.message : 'Unable to update student lock.');
    }
  }

  async function loadActivity() {
    const query = new URLSearchParams(Object.entries(activityFilters).filter(([, value]) => value).map(([key, value]) => [key, String(value)]));
    try {
      const result = await responseJson(await fetch(`/api/activity?${query.toString()}`, { cache: 'no-store' }));
      setActivity((current) => activityFilters.skip ? [...current, ...result.events] : result.events);
      setActivityHasMore(result.hasMore);
    } catch (error) {
      setStatusMessage(error instanceof Error ? error.message : 'Unable to load activity.');
    }
  }

  useEffect(() => {
    if (sessionStatus === 'authenticated' && view === 'activity') void loadActivity();
  }, [sessionStatus, view, activityFilters]);

  async function exportWorkbook() {
    try {
      const query = new URLSearchParams(Object.entries(activityFilters).filter(([, value]) => value).map(([key, value]) => [key, String(value)]));
      query.set('limit', '10000');
      const historyResult = await responseJson(await fetch(`/api/activity?${query.toString()}`, { cache: 'no-store' }));
      const history: EventRecord[] = historyResult.events;
      const eventDefaultIds = new Set(history.map((event) => event.defaultId).filter(Boolean));
      const visibleDefaults = defaults.filter((record) =>
        (showRemoved || record.status !== 'removed') &&
        (!activityFilters.rollNo || normalizeRollNo(record.rollNo) === normalizeRollNo(activityFilters.rollNo)) &&
        (!Object.values(activityFilters).some(Boolean) || eventDefaultIds.has(record.id)),
      );
      const selectedRollNos = new Set(history.map((event) => normalizeRollNo(event.rollNo)));
      const summary = students
        .filter((student) => Object.values(activityFilters).every((value) => !value)
          ? true
          : selectedRollNos.has(normalizeRollNo(student.rollNo)))
        .map((student) => ({
          'Roll number': student.rollNo,
          Name: student.name,
          'Number of defaults': student.numberOfDefaults ?? 0,
          'Lock / penalty status': student.lockPenaltyStatus ?? 'Clear',
          Batch: student.batch ? `${student.batch[0].toUpperCase()}${student.batch.slice(1)} batch` : 'Not set',
        }));
      const details = visibleDefaults.map((record) => ({
        'Roll number': record.rollNo,
        Student: students.find((student) => normalizeRollNo(student.rollNo) === normalizeRollNo(record.rollNo))?.name ?? record.rollNo,
        Batch: studentsByRollNo.get(normalizeRollNo(record.rollNo))?.batch ? `${studentsByRollNo.get(normalizeRollNo(record.rollNo))?.batch === 'junior' ? 'Junior' : 'Senior'} batch` : 'Not set',
        Company: record.company,
        Type: record.typeName,
        Date: record.date,
        'Logged by': record.loggedByName,
        Team: teamLabel(record.team),
        'Entered by': record.enteredByEmail,
        'Last updated by': record.updatedByName,
        Status: record.status,
        'Locked/Removed date': record.status === 'locked' ? record.lockedAt ?? '' : record.status === 'removed' ? record.removedAt ?? '' : '',
        'DC reference': record.dcReference ?? '',
        'DC date': record.dcDate ?? '',
        'DC status': record.dcStatus ?? 'Not scheduled',
        'Deferred until': record.deferral?.until === 'next_dc' ? 'Next DC' : record.deferral?.until ?? '',
        'Deferral reason': record.deferral?.reason ?? '',
        'Deferral history': (record.deferral?.history ?? []).map((item) => `${item.until === 'next_dc' ? 'Next DC' : item.until} — ${item.reason} (${item.byName ?? 'Unknown member'} · ${item.at ? new Date(item.at).toLocaleString() : ''})`).join('\n'),
        'Discussion reason': record.discussion?.reason ?? '',
        'Discussion date': record.discussion?.discussionDate ?? '',
        'Discussion notes': (record.discussion?.notes ?? []).map((note) => `${note.at ? new Date(note.at).toLocaleString() : ''} ${note.byName}: ${note.text}`).join('\n'),
        Notes: record.notes,
      }));
      const historyRows = history.map((event) => ({
        'Roll number': event.rollNo,
        Name: event.studentName ?? event.rollNo,
        Batch: event.studentBatch ? `${event.studentBatch === 'junior' ? 'Junior' : 'Senior'} batch` : 'Not set',
        'Event type': event.type,
        Actor: event.actorLabel ?? event.memberName,
        Category: event.category ?? '',
        Team: teamLabel(event.team),
        Account: event.enteredByEmail,
        'Date and time': new Date(event.at).toLocaleString(),
        'Reason or note': event.reason || event.note || '',
        'DC reference': event.dcReference ?? '',
      }));
      const workbook = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(summary), 'Summary');
      XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(details), 'Details');
      XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(historyRows), 'History');
      downloadWorkbook(workbook, 'default-tracker.xlsx');
    } catch (error) {
      setStatusMessage(error instanceof Error ? error.message : 'Unable to export workbook.');
    }
  }

  if (sessionStatus === 'loading') return <main className="mx-auto max-w-xl px-4 py-16 text-center text-slate-600">Checking sign-in...</main>;
  if (!session) {
    return (
      <main className="mx-auto max-w-xl px-4 py-16">
        <section className="rounded-2xl border border-slate-200 bg-white p-8 text-center shadow-sm">
          <p className="text-xs font-semibold uppercase tracking-[0.2em] text-indigo-600">Placement office</p>
          <h1 className="mt-2 text-3xl font-bold text-slate-900">Default Tracker</h1>
          <p className="mt-4 text-slate-600">Sign in using an email address enabled in the active member directory.</p>
          <button onClick={() => void signIn('google')} className="mt-6 rounded-lg bg-indigo-600 px-4 py-2.5 font-semibold text-white hover:bg-indigo-500">Sign in with Google</button>
        </section>
      </main>
    );
  }

  return (
    <main className="min-h-screen bg-slate-100 text-slate-900">
      <div className="flex min-h-screen">
        <aside className={`${sidebarOpen ? 'translate-x-0' : '-translate-x-full'} fixed inset-y-0 left-0 z-40 flex w-72 flex-col border-r border-slate-200 bg-slate-950 text-slate-100 transition md:static md:translate-x-0 ${sidebarCollapsed ? 'md:w-20' : 'md:w-72'}`}>
          <div className="flex items-center justify-between border-b border-slate-800 p-4">
            {!sidebarCollapsed && <div><p className="text-[10px] font-semibold uppercase tracking-[0.25em] text-indigo-300">Placement office</p><h1 className="mt-2 text-xl font-bold text-white">Default Tracker</h1></div>}
            <button type="button" onClick={() => setSidebarCollapsed((current) => !current)} className="hidden h-9 w-9 items-center justify-center rounded-full border border-slate-700 bg-slate-900 text-sm text-slate-200 hover:bg-slate-800 md:inline-flex">{sidebarCollapsed ? '→' : '←'}</button>
          </div>

          <nav aria-label="Main navigation" className="flex-1 space-y-1 p-4">
            {([
              ['defaults', 'Defaults'],
              ['students', 'Students'],
              ['dc', 'Disciplinary Committee'],
              ['members', 'Members'],
              ['activity', 'Activity'],
              ['settings', 'Settings'],
            ] as Array<[ViewName, string]>).map(([key, label]) => (
              <button key={key} type="button" onClick={() => { setView(key); setSidebarOpen(false); if (key === 'dc') setDcSubView('planner'); if (key === 'settings') setSettingsSubView('data'); }} className={`flex w-full items-center gap-3 rounded-xl px-3 py-3 text-left text-sm font-medium transition ${view === key || (key === 'defaults' && view === 'default-import') ? 'bg-indigo-600 text-white shadow-sm' : 'text-slate-300 hover:bg-slate-800 hover:text-white'}`}>
                <span className="inline-flex h-6 w-6 items-center justify-center rounded-md bg-slate-800 text-xs font-semibold">{label.slice(0, 1)}</span>
                {!sidebarCollapsed && <span>{label}</span>}
              </button>
            ))}

            {view === 'dc' && !sidebarCollapsed && (
              <div className="ml-8 mt-2 space-y-1 rounded-xl border border-slate-800 bg-slate-900 p-2">
                {[
                  ['planner', 'Planner'],
                  ['summary', 'Summary'],
                ].map(([key, label]) => (
                  <button key={key} type="button" onClick={() => setDcSubView(key as 'planner' | 'summary')} className={`w-full rounded-lg px-3 py-2 text-left text-sm ${dcSubView === key ? 'bg-slate-700 text-white' : 'text-slate-300 hover:bg-slate-800 hover:text-white'}`}>
                    {label}
                  </button>
                ))}
              </div>
            )}
            {view === 'settings' && !sidebarCollapsed && (
              <div className="ml-8 mt-2 space-y-1 rounded-xl border border-slate-800 bg-slate-900 p-2">
                {([
                  ['data', 'Data management'],
                  ['types', 'Default types'],
                ] as Array<[SettingsSubView, string]>).map(([key, label]) => (
                  <button key={key} type="button" onClick={() => setSettingsSubView(key)} className={`w-full rounded-lg px-3 py-2 text-left text-sm ${settingsSubView === key ? 'bg-slate-700 text-white' : 'text-slate-300 hover:bg-slate-800 hover:text-white'}`}>
                    {label}
                  </button>
                ))}
              </div>
            )}
          </nav>

          <div className="border-t border-slate-800 p-4">
            {!sidebarCollapsed && <p className="mb-3 text-sm text-slate-300">{accountEmail}</p>}
            <button type="button" onClick={() => void signOut()} className={`flex w-full items-center justify-center rounded-xl border border-slate-700 px-3 py-2 text-sm font-medium text-slate-100 hover:bg-slate-800 ${sidebarCollapsed ? 'px-2' : ''}`}>
              {sidebarCollapsed ? 'Out' : 'Sign out'}
            </button>
          </div>
        </aside>

        <div className="flex-1">
          <header className="border-b border-slate-200 bg-white/90 backdrop-blur-sm">
            <div className="flex items-center justify-between gap-4 px-4 py-4 sm:px-6 lg:px-8">
              <div className="flex items-center gap-3">
                <button type="button" className="inline-flex h-10 w-10 items-center justify-center rounded-lg border border-slate-200 bg-white text-lg text-slate-700 md:hidden" onClick={() => setSidebarOpen((current) => !current)} aria-label="Toggle navigation">☰</button>
                <div>
                  <p className="text-[10px] font-semibold uppercase tracking-[0.25em] text-indigo-600">{view === 'dc' ? 'Disciplinary Committee' : view === 'defaults' || view === 'default-import' ? 'Defaults' : view === 'students' ? 'Students' : view === 'members' ? 'Members' : view === 'settings' ? 'Settings' : 'Activity'}</p>
                  <h2 className="mt-1 text-2xl font-bold text-slate-900">{view === 'dc' ? (dcSubView === 'summary' ? 'DC Summary' : 'DC Planner') : view === 'defaults' ? 'Defaults' : view === 'default-import' ? 'Import from Excel' : view === 'students' ? 'Student roster' : view === 'members' ? 'Members' : view === 'settings' ? (settingsSubView === 'types' ? 'Default types' : 'Data management') : 'Activity log'}</h2>
                </div>
              </div>
              <div className="flex flex-wrap items-center gap-2">
                <span className="hidden text-sm text-slate-600 sm:inline">{accountEmail}</span>
                <label className="inline-flex min-h-11 cursor-pointer items-center rounded-lg border border-indigo-200 bg-indigo-50 px-3 py-2 text-sm font-medium text-indigo-700 hover:bg-indigo-100">
                  Upload roster<input type="file" accept=".csv,.xlsx,.xls" className="hidden" onChange={handleRosterUpload} />
                </label>
                <label className="inline-flex min-h-11 cursor-pointer items-center rounded-lg border border-indigo-200 bg-indigo-50 px-3 py-2 text-sm font-medium text-indigo-700 hover:bg-indigo-100">
                  Upload members<input type="file" accept=".csv,.xlsx,.xls" className="hidden" onChange={handleMemberUpload} />
                </label>
                <button onClick={() => void exportWorkbook()} className="inline-flex min-h-11 items-center justify-center rounded-lg bg-slate-900 px-3 py-2 text-sm font-medium text-white hover:bg-slate-700">Export Excel</button>
              </div>
            </div>
          </header>

          <div className="px-4 py-6 sm:px-6 lg:px-8">
            {dashboardLoading && <p role="status" className="mb-4 rounded-lg bg-slate-50 px-4 py-3 text-sm text-slate-600"><span className="loading-spinner" aria-hidden="true" /> Loading tracker data…</p>}
            {statusMessage && <p role="status" className="mb-4 rounded-lg border border-indigo-100 bg-indigo-50 px-4 py-3 text-sm text-indigo-900">{statusMessage}</p>}

      {view === 'defaults' && <>
        <div className="mb-5 flex gap-2 border-b border-slate-200" role="tablist" aria-label="Dashboard tabs">
          <button type="button" role="tab" aria-selected={dashboardTab === 'overview'} onClick={() => setDashboardTab('overview')} className={`border-b-2 px-4 py-3 text-sm font-semibold ${dashboardTab === 'overview' ? 'border-indigo-600 text-indigo-700' : 'border-transparent text-slate-500'}`}>Overview</button>
          <button type="button" role="tab" aria-selected={dashboardTab === 'dc'} onClick={() => setDashboardTab('dc')} className={`border-b-2 px-4 py-3 text-sm font-semibold ${dashboardTab === 'dc' ? 'border-indigo-600 text-indigo-700' : 'border-transparent text-slate-500'}`}>DC</button>
        </div>
        {dashboardTab === 'overview' ? <>
        <section className="mb-8 grid gap-4 sm:grid-cols-2 xl:grid-cols-6">
          {[
            ['Total defaults', stats.total, 'text-slate-900'],
            ['Logged', stats.logged, 'text-emerald-600'],
            ['Locked', stats.locked, 'text-amber-600'],
            ['Removed', stats.removed, 'text-slate-500'],
            ['Awaiting team decision', stats.toDiscuss, 'text-violet-600'],
            ['Deferred', stats.deferred, 'text-sky-600'],
          ].map(([label, count, color]) => (
            <div key={label} className="rounded-2xl border border-slate-200 bg-white p-5 shadow-sm"><p className="text-sm text-slate-500">{label}</p><p className={`mt-3 text-3xl font-bold ${color}`}>{count}</p></div>
          ))}
        </section>
        <section className="mb-8 grid gap-6 lg:grid-cols-[1.2fr_0.8fr]">
          <form onSubmit={handleSubmit} className="rounded-2xl border border-slate-200 bg-white p-6 shadow-sm">
            <div className="mb-5 flex flex-wrap items-center justify-between gap-3"><h2 className="text-xl font-semibold text-slate-900">Log a default</h2><div className="flex flex-wrap items-center gap-2"><button type="button" onClick={() => setView('default-import')} className="rounded-lg border border-indigo-200 bg-indigo-50 px-3 py-2 text-sm font-medium text-indigo-700">Import from Excel</button><span className="rounded-full bg-indigo-100 px-2.5 py-1 text-xs font-medium text-indigo-700">Roll number only</span></div></div>
            <div className="space-y-4">
              <div>
                <label htmlFor="roll-number" className="mb-2 block text-sm font-medium text-slate-700">Student roll number</label>
                <div className="flex gap-2">
                  <input id="roll-number" aria-describedby={lookupError ? 'roll-number-error' : undefined} autoFocus value={rollNumber} onChange={(event) => { setRollNumber(event.target.value); setSelectedStudent(null); setLookupError(''); }} placeholder="e.g. 2026/001" className="w-full rounded-lg border border-slate-300 bg-slate-50 px-3 py-2.5" />
                  <button type="button" onClick={handleLookup} className="rounded-lg bg-slate-900 px-4 py-2.5 text-sm font-medium text-white">Look up</button>
                </div>
                {lookupError && <p id="roll-number-error" role="alert" className="mt-2 text-sm text-rose-600">{lookupError}</p>}
                {selectedStudent && <p className="mt-3 rounded-lg bg-emerald-50 px-3 py-2 text-sm text-emerald-800">Student: <strong>{selectedStudent.name}</strong> · {selectedStudent.batch ? `${selectedStudent.batch === 'junior' ? 'Junior' : 'Senior'} batch` : <strong className="text-amber-700">Batch not set</strong>}</p>}
              </div>
              <label className="block text-sm font-medium text-slate-700">Company<input required value={company} onChange={(event) => setCompany(event.target.value)} className="mt-2 w-full rounded-lg border border-slate-300 bg-slate-50 px-3 py-2.5" /></label>
              <div className="grid gap-4 sm:grid-cols-2">
                <label className="block text-sm font-medium text-slate-700">Default type<select required value={typeId} onChange={(event) => setTypeId(event.target.value)} className="mt-2 w-full rounded-lg border border-slate-300 bg-slate-50 px-3 py-2.5">{defaultTypes.map((type) => <option key={type.id} value={type.id}>{type.name}</option>)}</select></label>
                <label className="block text-sm font-medium text-slate-700">Date<input required type="date" value={date} onChange={(event) => setDate(event.target.value)} className="mt-2 w-full rounded-lg border border-slate-300 bg-slate-50 px-3 py-2.5" /></label>
              </div>
              <p className="rounded-lg bg-slate-50 px-3 py-2.5 text-sm text-slate-700">Logging as: <strong>{signedInMember ? `${signedInMember.name} (${teamLabel(signedInMember.team)})` : 'Active member account required'}</strong></p>
              <label className="block text-sm font-medium text-slate-700">Notes<textarea rows={3} value={notes} onChange={(event) => setNotes(event.target.value)} className="mt-2 w-full rounded-lg border border-slate-300 bg-slate-50 px-3 py-2.5" /></label>
              <div className="text-xs text-slate-500">Entered by (account): {accountEmail}</div>
              <button type="submit" disabled={isSubmitting || !selectedStudent || !signedInMember} className="rounded-lg bg-indigo-600 px-4 py-2.5 text-sm font-semibold text-white disabled:bg-slate-300">{isSubmitting ? 'Logging...' : 'Log default'}</button>
            </div>
          </form>
          <aside className="rounded-2xl border border-slate-200 bg-white p-6 shadow-sm">
            <h2 className="text-xl font-semibold text-slate-900">Student roster</h2>
            <div className="mt-4 max-h-[440px] space-y-2 overflow-auto pr-1">{students.map((student) => <button key={student.rollNo} onClick={() => { setStudentSearch(student.rollNo); setView('students'); setStudentDetail(null); }} className="flex w-full items-center justify-between rounded-lg border border-slate-200 bg-slate-50 px-3 py-2 text-left hover:bg-indigo-50"><span><span className="block font-medium text-slate-800">{student.name}</span><span className="text-xs uppercase text-slate-500">{student.rollNo} · {student.batch ? `${student.batch === 'junior' ? 'Junior' : 'Senior'} batch` : 'Batch not set'}</span></span><span className="text-right text-xs text-slate-600">{student.numberOfDefaults ?? 0} defaults · {student.numberOfRemovedDefaults ?? 0} removed<br />{student.lockPenaltyStatus ?? 'Clear'}</span></button>)}</div>
          </aside>
        </section>
        <section className="rounded-2xl border border-slate-200 bg-white p-6 shadow-sm">
          <div className="sticky-filters mb-4 flex flex-wrap items-end justify-between gap-3 bg-white py-2"><div><h2 className="text-xl font-semibold text-slate-900">Defaults</h2><p className="text-sm text-slate-500">{activeDefaults.length} results</p></div><div className="flex flex-wrap items-end gap-3"><label className="text-sm">Search<input value={defaultSearch} onChange={(event) => setDefaultSearch(event.target.value)} placeholder="Student, company, type…" className="mt-1 min-h-11 rounded-lg border px-3" /></label><label className="text-sm">Batch<select value={batchFilter} onChange={(event) => setBatchFilter(event.target.value)} className="mt-1 min-h-11 rounded-lg border px-3"><option value="">All batches</option><option value="junior">Junior batch</option><option value="senior">Senior batch</option><option value="not_set">Not set</option></select></label><label className="text-sm">Status<select value={defaultStatusFilter} onChange={(event) => setDefaultStatusFilter(event.target.value)} className="mt-1 min-h-11 rounded-lg border px-3"><option value="active">Active defaults</option><option value="all">All statuses</option><option value="logged">Logged</option><option value="locked">Locked</option><option value="removed">Removed</option></select></label></div></div>
          <div className="responsive-table"><table className="min-w-full text-left text-sm"><thead className="border-b bg-slate-50 text-slate-600"><tr>{['Student', 'Batch', 'Company', 'Type', 'Logged by', 'Team', 'Date logged', 'Last updated by', 'Status', ...(defaultStatusFilter === 'removed' || defaultStatusFilter === 'all' || showRemoved ? ['Removed on', 'DC date', 'Removal reason'] : []), 'Actions'].map((header) => <th key={header} className="px-3 py-3 font-medium">{header}</th>)}</tr></thead><tbody>
            {activeDefaults.map((record) => <tr key={record.id} className={`border-b align-top ${record.status === 'removed' ? 'bg-slate-50 text-slate-500' : ''}`}>
              <td data-label="Student" className="px-3 py-3"><button className="text-left" onClick={() => void openDefault(record)}><span className="block font-medium text-indigo-700">{students.find((student) => normalizeRollNo(student.rollNo) === normalizeRollNo(record.rollNo))?.name ?? record.rollNo}</span><span className="text-xs text-slate-500">{record.rollNo}</span></button></td>
              <td data-label="Batch" className="px-3 py-3">{studentsByRollNo.get(normalizeRollNo(record.rollNo))?.batch ? `${studentsByRollNo.get(normalizeRollNo(record.rollNo))?.batch === 'junior' ? 'Junior' : 'Senior'} batch` : <span className="font-medium text-amber-700">Not set</span>}</td><td data-label="Company" className="px-3 py-3">{record.company}</td><td data-label="Type" className="px-3 py-3">{record.typeName}</td><td data-label="Logged by" className="px-3 py-3">{record.loggedByName}</td><td data-label="Team" className="px-3 py-3">{teamLabel(record.team)}</td><td data-label="Date logged" className="px-3 py-3">{new Date(record.createdAt).toLocaleString()}</td><td data-label="Last updated by" className="px-3 py-3">{record.updatedByName}</td>
              <td data-label="Status" className="px-3 py-3"><span className={`status-badge ${record.status === 'locked' ? 'status-pending' : record.status === 'removed' ? 'status-danger' : 'status-success'}`}>{record.status === 'logged' ? 'Logged' : record.status === 'locked' ? 'Locked' : 'Removed'}</span></td>
              {(defaultStatusFilter === 'removed' || defaultStatusFilter === 'all' || showRemoved) && <><td data-label="Removed on" className="px-3 py-3">{record.status === 'removed' && record.removedAt ? new Date(record.removedAt).toLocaleDateString() : '—'}</td><td data-label="DC date" className="px-3 py-3">{record.status === 'removed' ? record.dcDate ?? record.dcReference ?? '—' : '—'}</td><td data-label="Removal reason" className="px-3 py-3">{record.status === 'removed' ? record.removeReason ?? record.notes ?? '—' : '—'}</td></>}
              <td data-label="Actions" className="px-3 py-3"><button onClick={() => void openDefault(record)} className="rounded border px-2 py-1 text-xs">Details</button></td>
            </tr>)}
            {!activeDefaults.length && <tr><td colSpan={defaultStatusFilter === 'removed' || defaultStatusFilter === 'all' || showRemoved ? 13 : 10} className="px-4 py-8 text-center text-slate-500">{defaults.length ? 'No defaults match these filters.' : 'No defaults yet.'}</td></tr>}
          </tbody></table></div>
        </section>
        </> : <DashboardDCTab />}
      </>}

      {view === 'default-import' && <section className="space-y-6">
        <div className="flex flex-wrap items-center justify-between gap-3 rounded-2xl border border-slate-200 bg-white p-5 shadow-sm">
          <div><h2 className="text-xl font-semibold">Import defaults from Excel</h2><p className="mt-1 text-sm text-slate-600">Upload an .xlsx or .csv file with up to 3,000 rows and under 4 MB.</p></div>
          <div className="flex flex-wrap gap-2">
            <button type="button" onClick={() => setView('defaults')} className="rounded-lg border px-3 py-2 text-sm">Back to Defaults</button>
            <button type="button" onClick={() => void downloadDefaultImportSample()} className="rounded-lg bg-indigo-600 px-3 py-2 text-sm font-medium text-white">Download sample Excel</button>
          </div>
        </div>
        <section className="rounded-2xl border border-slate-200 bg-white p-6 shadow-sm">
          <label className="block text-sm font-medium text-slate-700">Choose an Excel or CSV file
            <input type="file" accept=".xlsx,.csv" disabled={defaultImportBusy} onChange={(event) => void handleDefaultImportFile(event)} className="mt-2 block w-full rounded-lg border border-slate-300 p-2" />
          </label>
          {defaultImportFile && <p className="mt-2 text-sm text-slate-600">Selected: {defaultImportFile.name} ({(defaultImportFile.size / (1024 * 1024)).toFixed(2)} MB)</p>}
          <div className="mt-5 grid gap-3 md:grid-cols-2">
            <fieldset className="space-y-2 rounded-lg border border-slate-200 p-3">
              <legend className="px-1 text-sm font-semibold">When some rows have errors</legend>
              <label className="flex items-start gap-2 text-sm"><input type="radio" name="default-import-mode" checked={!defaultImportStopOnError} onChange={() => setDefaultImportStopOnError(false)} /><span>Import valid rows only (default)</span></label>
              <label className="flex items-start gap-2 text-sm"><input type="radio" name="default-import-mode" checked={defaultImportStopOnError} onChange={() => setDefaultImportStopOnError(true)} /><span>Stop if any row has an error</span></label>
            </fieldset>
            <label className="flex items-start gap-2 rounded-lg border border-slate-200 p-3 text-sm">
              <input type="checkbox" checked={defaultImportCreateTypes} onChange={(event) => {
                const checked = event.target.checked;
                setDefaultImportCreateTypes(checked);
                if (defaultImportFile) void previewDefaultImport(defaultImportFile, checked);
              }} />
              <span><strong>Create missing default types automatically</strong><br />Off by default. If enabled, missing types used by valid rows are created during import.</span>
            </label>
          </div>
          {defaultImportBusy && <p role="status" className="mt-4 text-sm text-slate-600">Validating/importing file…</p>}
          {defaultImportPreview && <div className="mt-6 border-t border-slate-200 pt-5">
            <div className="grid gap-3 sm:grid-cols-4">
              {[
                ['Rows read', defaultImportPreview.counts.rowsRead],
                ['Ready to import', defaultImportPreview.counts.ready],
                ['Rows with warnings', defaultImportPreview.counts.warnings],
                ['Rows with errors', defaultImportPreview.counts.errors],
              ].map(([label, count]) => <div key={label} className="rounded-lg bg-slate-50 p-3"><p className="text-sm text-slate-500">{label}</p><p className="mt-1 text-2xl font-semibold">{count}</p></div>)}
            </div>
            {defaultImportPreview.imported !== undefined && <p role="status" className="mt-4 rounded-lg bg-emerald-50 p-3 text-sm text-emerald-900">
              {defaultImportPreview.stoppedForErrors
                ? 'Import stopped because errors were found; no records were saved.'
                : `${defaultImportPreview.imported} rows imported; ${defaultImportPreview.skipped ?? 0} rows skipped.`}
            </p>}
            <div className="mt-4 flex flex-wrap items-center gap-2">
              <button type="button" onClick={() => void submitDefaultImport()} disabled={defaultImportBusy || !defaultImportFile || !defaultImportPreview.counts.rowsRead || defaultImportPreview.imported !== undefined} className="rounded-lg bg-indigo-600 px-4 py-2.5 text-sm font-semibold text-white disabled:bg-slate-300">{defaultImportStopOnError && defaultImportPreview.counts.errors ? 'Confirm (stop if errors)' : 'Confirm import'}</button>
              {defaultImportPreview.rejectedRows.length > 0 && <button type="button" onClick={downloadImportErrorReport} className="rounded-lg border px-3 py-2 text-sm">Download error report</button>}
              <button type="button" onClick={() => { setDefaultImportFile(null); setDefaultImportPreview(null); }} className="rounded-lg border px-3 py-2 text-sm">Clear preview</button>
            </div>
            {defaultImportPreview.issues.length > 0 && <div className="responsive-table mt-5 max-h-96 overflow-auto rounded-lg border border-slate-200">
              <table className="min-w-full text-left text-sm"><thead className="sticky top-0 bg-slate-50"><tr>{['Severity', 'Row', 'Column', 'Problem'].map((heading) => <th key={heading} className="px-3 py-2">{heading}</th>)}</tr></thead>
                <tbody>{defaultImportPreview.issues.map((issue, index) => <tr key={`${issue.row}-${issue.column}-${index}`} className="border-t border-slate-100"><td data-label="Severity" className={`px-3 py-2 font-medium ${issue.severity === 'error' ? 'text-rose-700' : 'text-amber-700'}`}>{issue.severity}</td><td data-label="Row" className="px-3 py-2">{issue.row}</td><td data-label="Column" className="px-3 py-2">{issue.column}</td><td data-label="Problem" className="px-3 py-2">{issue.reason}</td></tr>)}</tbody>
              </table>
            </div>}
          </div>}
        </section>
        <section className="rounded-2xl border border-slate-200 bg-white p-6 shadow-sm">
          <div className="flex flex-wrap items-center justify-between gap-3"><div><h2 className="text-xl font-semibold">Import history</h2><p className="mt-1 text-sm text-slate-500">Undo is available only before imported defaults are used in a DC or changed to locked/removed.</p></div><button type="button" onClick={() => void loadImportBatches()} className="rounded-lg border px-3 py-2 text-sm">Refresh</button></div>
          <div className="responsive-table mt-4"><table className="min-w-full text-left text-sm"><thead><tr>{['Imported', 'File', 'Imported by', 'Rows', 'Result', 'Undo'].map((heading) => <th key={heading} className="px-3 py-2">{heading}</th>)}</tr></thead>
            <tbody>{importBatches.map((batch) => {
              const disabledReason = batch.undoneAt ? 'Already undone.' : batch.undoBlockReason || (!batch.imported ? 'This batch has no imported defaults.' : '');
              return <tr key={batch.batchId} className="border-t border-slate-100">
                <td data-label="Imported" className="px-3 py-3">{batch.importedAt ? new Date(batch.importedAt).toLocaleString() : '—'}</td>
                <td data-label="File" className="px-3 py-3">{batch.filename}</td>
                <td data-label="Imported by" className="px-3 py-3">{batch.importedBy}</td>
                <td data-label="Rows" className="px-3 py-3">{batch.rowsRead}</td>
                <td data-label="Result" className="px-3 py-3">{batch.imported} imported · {batch.skipped} skipped · {batch.errors} errors</td>
                <td data-label="Undo" className="px-3 py-3"><button type="button" title={disabledReason || 'Undo this import'} disabled={Boolean(disabledReason)} onClick={() => void undoImport(batch)} className="rounded border px-2 py-1 text-xs disabled:cursor-not-allowed disabled:opacity-50">{batch.undoneAt ? 'Undone' : 'Undo this import'}</button>{disabledReason && !batch.undoneAt && <p className="mt-1 max-w-xs text-xs text-slate-500">{disabledReason}</p>}</td>
              </tr>;
            })}{!importBatches.length && <tr><td colSpan={6} className="px-3 py-6 text-center text-slate-500">No imports yet.</td></tr>}</tbody>
          </table></div>
        </section>
      </section>}

      {view === 'dc' && dcSubView === 'planner' && <DCTab onRefresh={loadDashboard} onOpenStudent={openStudentFromDc} />}
      {view === 'dc' && dcSubView === 'summary' && <DCSummaryTab />}

      {view === 'students' && <section className="rounded-2xl border border-slate-200 bg-white p-6 shadow-sm">
        <div className="flex flex-wrap items-center justify-between gap-3"><h2 className="text-xl font-semibold">Student history</h2><button type="button" onClick={downloadStudentTemplate} className="rounded-lg border px-3 py-2 text-sm">Download sample Excel</button></div>
        <div className="mt-4 flex items-end gap-2"><label className="w-full max-w-sm text-sm font-medium">Student roll number<input value={studentSearch} onChange={(event) => setStudentSearch(event.target.value)} onKeyDown={(event) => { if (event.key === 'Enter') { event.preventDefault(); void loadStudentDetail(); } }} placeholder="Search by roll number" className="mt-1 w-full rounded-lg border px-3 py-2" /></label><button onClick={() => void loadStudentDetail()} className="rounded-lg bg-indigo-600 px-4 py-2 text-white">Search</button></div>
        <div className="mt-8 border-t pt-5">
          <div className="flex flex-wrap items-end justify-between gap-3">
            <div><h3 className="font-semibold">Roster batches</h3><p className="text-sm text-slate-500">{filteredStudents.length} students · Not set means no batch is assigned.</p></div>
            <div className="flex flex-wrap items-end gap-2">
              <label className="text-sm">Search roster<input value={rosterSearch} onChange={(event) => setRosterSearch(event.target.value)} placeholder="Name or roll number" className="mt-1 min-h-10 rounded-lg border px-3" /></label>
              <label className="text-sm">Filter batch<select value={studentBatchFilter} onChange={(event) => setStudentBatchFilter(event.target.value)} className="mt-1 min-h-10 rounded-lg border px-3"><option value="">All batches</option><option value="junior">Junior batch</option><option value="senior">Senior batch</option><option value="not_set">Not set</option></select></label>
              <label className="text-sm">Set selected to<select value={bulkStudentBatch} onChange={(event) => setBulkStudentBatch(event.target.value)} className="mt-1 min-h-10 rounded-lg border px-3"><option value="">Choose batch</option><option value="junior">Junior batch</option><option value="senior">Senior batch</option><option value="unset">Not set</option></select></label>
              <button type="button" disabled={!selectedStudentRollNos.length || !bulkStudentBatch} onClick={() => void updateSelectedStudentBatches()} className="min-h-10 rounded-lg bg-indigo-600 px-3 text-sm text-white disabled:bg-slate-300">Update {selectedStudentRollNos.length} selected</button>
            </div>
          </div>
          <div className="mt-3 flex items-center gap-2 text-sm"><input type="checkbox" aria-label="Select all filtered students" checked={filteredStudents.length > 0 && filteredStudents.every((student) => selectedStudentRollNos.includes(student.rollNo))} onChange={(event) => setSelectedStudentRollNos(event.target.checked ? Array.from(new Set([...selectedStudentRollNos, ...filteredStudents.map((student) => student.rollNo)])) : selectedStudentRollNos.filter((rollNo) => !filteredStudents.some((student) => student.rollNo === rollNo)))} />Select all visible</div>
          <div className="responsive-table mt-2 max-h-96 overflow-auto"><table className="min-w-full text-left text-sm"><thead className="sticky top-0 border-b bg-slate-50 text-slate-600"><tr><th className="px-3 py-2">Select</th><th className="px-3 py-2">Student</th><th className="px-3 py-2">Roll number</th><th className="px-3 py-2">Batch</th></tr></thead><tbody>
            {filteredStudents.map((student) => <tr key={student.rollNo} className="border-b"><td className="px-3 py-2"><input type="checkbox" aria-label={`Select ${student.name}`} checked={selectedStudentRollNos.includes(student.rollNo)} onChange={(event) => setSelectedStudentRollNos(event.target.checked ? [...selectedStudentRollNos, student.rollNo] : selectedStudentRollNos.filter((rollNo) => rollNo !== student.rollNo))} /></td><td className="px-3 py-2">{student.name}</td><td className="px-3 py-2">{student.rollNo}</td><td className="px-3 py-2">{student.batch ? `${student.batch === 'junior' ? 'Junior' : 'Senior'} batch` : <span className="font-medium text-amber-700">Not set</span>}</td></tr>)}
            {!filteredStudents.length && <tr><td colSpan={4} className="px-3 py-5 text-center text-slate-500">No students match this filter.</td></tr>}
          </tbody></table></div>
        </div>
        {studentDetail && <div className="mt-6">
          <div className="flex flex-wrap items-start justify-between gap-4"><div><h3 className="text-2xl font-semibold">{studentDetail.student.name}</h3><p className="text-slate-500">{studentDetail.student.rollNo}</p><p className="mt-2 text-sm">Defaults: {studentDetail.student.numberOfDefaults} · Removed: {studentDetail.student.numberOfRemovedDefaults ?? 0}</p><p className="mt-1 text-sm">Batch: {studentDetail.student.batch ? `${studentDetail.student.batch === 'junior' ? 'Junior' : 'Senior'} batch` : <span className="font-medium text-amber-700">Not set</span>}</p>          <label className="mt-2 block text-sm">Change batch<select aria-label="Change student batch" value={studentDetail.student.batch ?? ''} onChange={(event) => { if (['', 'junior', 'senior'].includes(event.target.value)) void changeStudentBatch(studentDetail.student.rollNo, event.target.value as 'junior' | 'senior' | ''); }} className="mt-1 rounded-lg border px-3 py-2"><option value="">Not set</option><option value="junior">Junior batch</option><option value="senior">Senior batch</option></select></label></div><button onClick={() => void toggleStudentLock()} className="rounded-lg border px-3 py-2">{studentDetail.student.studentLocked ? 'Unlock student' : 'Lock student'}</button></div>
          <h4 className="mt-6 font-semibold">Active defaults ({studentDetail.defaults.filter((record) => record.status !== 'removed').length})</h4>
          <div className="mt-2 space-y-3">{studentDetail.defaults.filter((record) => record.status !== 'removed').map((record) => (
            <article key={record.id} className="rounded-lg border border-slate-200 p-4">
              <div className="flex flex-wrap items-start justify-between gap-3">
                <div>
                  <h5 className="font-semibold">{record.company} · {record.typeName}</h5>
                  <p className="mt-1 text-sm text-slate-600">Default date: {record.date} · Logged by {record.loggedByName} ({teamLabel(record.team)})</p>
                </div>
                <span className={`status-badge ${record.status === 'locked' ? 'status-pending' : record.status === 'removed' ? 'status-danger' : 'status-success'}`}>
                  {record.status === 'logged' ? 'Logged' : record.status === 'locked' ? 'Locked' : 'Removed'}
                </span>
              </div>
              <div className="mt-3 space-y-2 border-t pt-3">
                {(record.decisionHistory ?? []).length ? record.decisionHistory.map((item, index) => {
                  const decision = item.decision === 'Undecided' ? 'No decision yet' : item.decision;
                  const until = item.decision === 'Defer'
                    ? ` · until ${item.until === 'next_dc' ? 'Next DC' : item.until || 'unspecified'}`
                    : '';
                  return (
                    <div key={`${item.dcId}-${index}`} className="rounded bg-slate-50 p-3 text-sm">
                      <p><strong>DC {item.dcDate}:</strong> {decision}{until}</p>
                      <p className="mt-1 text-slate-600">Remarks: {item.remarks || '—'}</p>
                    </div>
                  );
                }) : <p className="text-sm text-slate-600">No decision yet.</p>}
              </div>
              <button type="button" onClick={() => void openDefault(record)} className="mt-3 rounded border px-3 py-2 text-sm">Open default details</button>
            </article>
          ))}{!studentDetail.defaults.some((record) => record.status !== 'removed') && <p className="rounded-lg bg-slate-50 p-4 text-sm text-slate-500">No active defaults.</p>}</div>
          <h4 className="mt-6 font-semibold text-slate-600">Removed defaults ({studentDetail.defaults.filter((record) => record.status === 'removed').length})</h4>
          <div className="mt-2 space-y-3">{studentDetail.defaults.filter((record) => record.status === 'removed').map((record) => (
            <article key={record.id} className="rounded-lg border border-slate-200 bg-slate-50 p-4 text-slate-500">
              <div className="flex flex-wrap items-start justify-between gap-3"><div><h5 className="font-semibold">{record.company} · {record.typeName}</h5><p className="mt-1 text-sm">Default date: {record.date} · Removed on {record.removedAt ? new Date(record.removedAt).toLocaleDateString() : '—'}</p><p className="text-sm">DC date: {record.dcDate ?? record.dcReference ?? '—'}</p><p className="text-sm">Reason: {record.removeReason ?? record.notes ?? '—'}</p></div><span className="status-badge bg-slate-200 text-slate-700">Removed</span></div>
              <button type="button" onClick={() => void openDefault(record)} className="mt-3 rounded border px-3 py-2 text-sm">Open history</button>
            </article>
          ))}{!studentDetail.defaults.some((record) => record.status === 'removed') && <p className="rounded-lg bg-slate-50 p-4 text-sm text-slate-500">No removed defaults.</p>}</div>
        </div>}
      </section>}

      {view === 'activity' && <section className="rounded-2xl border border-slate-200 bg-white p-6 shadow-sm">
        <div className="sticky-filters bg-white py-2"><h2 className="text-xl font-semibold">Activity</h2><p className="text-sm text-slate-500">{activity.length} results</p></div>
        <div className="mt-3 flex flex-wrap gap-2" role="group" aria-label="Activity category">
          {(['important', 'dc', 'members', 'imports', 'all'] as const).map((category) => <button key={category} type="button" aria-pressed={activityChip === category} onClick={() => { setActivityChip(category); setActivityFilters({ ...activityFilters, category, skip: 0 }); }} className={`rounded-full border px-3 py-1.5 text-sm ${activityChip === category ? 'border-indigo-600 bg-indigo-50 font-semibold text-indigo-700' : 'border-slate-300 text-slate-600'}`}>{category === 'dc' ? 'DC' : category[0].toUpperCase() + category.slice(1)}</button>)}
        </div>
        <div className="sticky-filters mt-4 grid gap-3 bg-white py-2 sm:grid-cols-2 lg:grid-cols-6">
          <label className="text-sm">Member<select aria-label="Member" value={activityFilters.memberId} onChange={(event) => setActivityFilters({ ...activityFilters, memberId: event.target.value, skip: 0 })} className="mt-1 w-full rounded border px-2 py-2"><option value="">All members</option>{members.map((member) => <option key={member.memberId} value={member.memberId}>{member.name}</option>)}</select></label>
          <label className="text-sm">Team<select aria-label="Team" value={activityFilters.team} onChange={(event) => setActivityFilters({ ...activityFilters, team: event.target.value, skip: 0 })} className="mt-1 w-full rounded border px-2 py-2"><option value="">All teams</option><option value="placement_cell">Placement cell</option><option value="summer_internship">Summer internship</option></select></label>
          <label className="text-sm">Event type<select aria-label="Event type" value={activityFilters.type} onChange={(event) => setActivityFilters({ ...activityFilters, type: event.target.value, skip: 0 })} className="mt-1 w-full rounded border px-2 py-2"><option value="">All event types</option>{['logged', 'edited', 'locked', 'unlocked', 'removed', 'restored', 'dc_note', 'dc_scheduled', 'dc_unscheduled', 'dc_attendance', 'dc_decision', 'dc_rescheduled', 'dc_cancelled', 'dc_auto_added', 'student_locked', 'student_unlocked', 'student_batch_updated'].map((type) => <option key={type} value={type}>{type.replaceAll('_', ' ')}</option>)}</select></label>
          <label className="text-sm">From date<input type="date" aria-label="From date" value={activityFilters.from} onChange={(event) => setActivityFilters({ ...activityFilters, from: event.target.value, skip: 0 })} className="mt-1 w-full rounded border px-2 py-2" /></label>
          <label className="text-sm">To date<input type="date" aria-label="To date" value={activityFilters.to} onChange={(event) => setActivityFilters({ ...activityFilters, to: event.target.value, skip: 0 })} className="mt-1 w-full rounded border px-2 py-2" /></label>
          <label className="text-sm">Roll number<input placeholder="Any roll number" value={activityFilters.rollNo} onChange={(event) => setActivityFilters({ ...activityFilters, rollNo: event.target.value, skip: 0 })} className="mt-1 w-full rounded border px-2 py-2" /></label>
          <label className="text-sm">Batch<select value={activityFilters.batch} onChange={(event) => setActivityFilters({ ...activityFilters, batch: event.target.value, skip: 0 })} className="mt-1 w-full rounded border px-2 py-2"><option value="">All batches</option><option value="junior">Junior batch</option><option value="senior">Senior batch</option><option value="not_set">Not set</option></select></label>
        </div>
        <Timeline events={activity} />
        {activityHasMore && <button type="button" onClick={() => setActivityFilters({ ...activityFilters, skip: activity.length })} className="mt-4 rounded-lg border px-4 py-2 text-sm font-medium">Load more</button>}
      </section>}

      {view === 'settings' && <section className="space-y-6">
        <div className="flex gap-2 border-b border-slate-200">
          <button type="button" onClick={() => setSettingsSubView('data')} className={`border-b-2 px-3 py-2 text-sm font-medium ${settingsSubView === 'data' ? 'border-indigo-600 text-indigo-700' : 'border-transparent text-slate-500'}`}>Data management</button>
          <button type="button" onClick={() => setSettingsSubView('types')} className={`border-b-2 px-3 py-2 text-sm font-medium ${settingsSubView === 'types' ? 'border-indigo-600 text-indigo-700' : 'border-transparent text-slate-500'}`}>Default types</button>
        </div>
        {settingsSubView === 'data' && <section className="rounded-2xl border border-slate-200 bg-white p-6 shadow-sm">
          <div className="flex flex-wrap items-start justify-between gap-4">
            <div><h2 className="text-xl font-semibold">Data management</h2><p className="mt-1 text-sm text-slate-600">Review record counts and export a complete workbook backup before clearing test data.</p></div>
            <button type="button" onClick={() => void loadResetCounts()} className="rounded-lg border px-3 py-2 text-sm">Refresh counts</button>
          </div>
          {!resetEnabled
            ? <p role="status" className="mt-5 rounded-lg border border-amber-200 bg-amber-50 p-4 font-medium text-amber-900">Data reset is disabled. Set ALLOW_DATA_RESET=true on the server only while testing, then redeploy.</p>
            : <>
              <div className="mt-5 flex flex-wrap gap-2">
                <button type="button" onClick={() => chooseResetPreset('test')} className="rounded-lg border border-indigo-200 bg-indigo-50 px-3 py-2 text-sm font-medium text-indigo-700">Clear test data</button>
                <button type="button" onClick={() => chooseResetPreset('full')} className="rounded-lg border border-rose-200 bg-rose-50 px-3 py-2 text-sm font-medium text-rose-700">Full reset</button>
              </div>
              <div className="mt-5 space-y-2">
                {resetCategoryLabels.map(({ key, label }) => <label key={key} className="flex items-center justify-between gap-3 rounded-lg border border-slate-200 px-3 py-3">
                  <span className="flex items-center gap-3"><input type="checkbox" checked={resetSelected.includes(key)} onChange={(event) => {
                    setResetSelected((current) => event.target.checked ? [...current, key] : current.filter((value) => value !== key));
                    setResetResult(null);
                  }} /><span>{label}</span></span>
                  <strong className="tabular-nums">{resetCounts?.[key] ?? 0}</strong>
                </label>)}
              </div>
              <div className="mt-5 rounded-lg border border-amber-300 bg-amber-50 p-4 text-sm text-amber-950">
                <p className="font-semibold">This cannot be undone. Export everything to Excel before deleting data.</p>
                <button type="button" onClick={() => void exportEverything()} className="mt-3 rounded-lg bg-slate-900 px-4 py-2 font-semibold text-white">Export everything to Excel</button>
                {backupToken && <p role="status" className="mt-2 text-emerald-800">Backup export confirmed. Reset is unlocked for 30 minutes.</p>}
              </div>
              <label className="mt-5 block text-sm font-medium text-slate-700">Type RESET to confirm
                <input value={resetConfirmation} onChange={(event) => setResetConfirmation(event.target.value)} autoComplete="off" className="mt-2 block w-full max-w-sm rounded-lg border border-slate-300 px-3 py-2.5" />
              </label>
              <button type="button" onClick={() => void performDataReset()} disabled={resetBusy || !backupToken || resetConfirmation !== 'RESET' || !resetSelected.length} className="mt-4 rounded-lg bg-rose-700 px-4 py-2.5 font-semibold text-white disabled:cursor-not-allowed disabled:bg-slate-300">{resetBusy ? 'Clearing selected data…' : 'Delete selected data'}</button>
              {resetResult && <div role="status" className="mt-5 rounded-lg border border-slate-200 bg-slate-50 p-4">
                <h3 className="font-semibold">Records deleted by collection</h3>
                <ul className="mt-2 grid gap-1 text-sm sm:grid-cols-2">{Object.entries(resetResult.removed ?? {}).map(([collection, count]) => <li key={collection}>{collection}: <strong>{count}</strong></li>)}</ul>
                {!!resetResult.errors?.length && <div className="mt-3 text-sm text-rose-700"><p className="font-semibold">Some deletions failed:</p><ul className="list-disc pl-5">{resetResult.errors.map((error, index) => <li key={index}>{error}</li>)}</ul></div>}
                <p className="mt-3 text-sm text-slate-600">A data reset audit entry was recorded. Your signed-in member record was preserved.</p>
              </div>}
            </>}
        </section>}
        {settingsSubView === 'types' && <section className="rounded-2xl border border-slate-200 bg-white p-6 shadow-sm">
          <div className="flex flex-wrap items-center justify-between gap-3">
            <div><h2 className="text-xl font-semibold">Default types</h2><p className="mt-1 text-sm text-slate-500">Types in use can be deactivated but cannot be deleted.</p></div>
            <button type="button" onClick={() => setDefaultTypeForm({ name: '', description: '', appliesTo: 'all_companies' })} className="rounded-lg bg-indigo-600 px-3 py-2 text-sm font-semibold text-white">Add type</button>
          </div>
          {defaultTypeForm && <form onSubmit={saveDefaultType} className="mt-5 grid gap-3 rounded-xl border border-indigo-100 bg-indigo-50/60 p-4 md:grid-cols-4">
            <label className="text-sm">Name *<input required maxLength={100} value={defaultTypeForm.name} onChange={(event) => setDefaultTypeForm({ ...defaultTypeForm, name: event.target.value })} className="mt-1 w-full rounded border px-3 py-2" /></label>
            <label className="text-sm md:col-span-2">Description<input value={defaultTypeForm.description} onChange={(event) => setDefaultTypeForm({ ...defaultTypeForm, description: event.target.value })} className="mt-1 w-full rounded border px-3 py-2" /></label>
            <label className="text-sm">Applies to<select value={defaultTypeForm.appliesTo} onChange={(event) => setDefaultTypeForm({ ...defaultTypeForm, appliesTo: event.target.value })} className="mt-1 w-full rounded border px-3 py-2"><option value="all_companies">All companies</option><option value="dream">Dream</option><option value="core">Core</option><option value="normal">Normal</option></select></label>
            <div className="flex gap-2 md:col-span-4"><button type="submit" className="rounded bg-indigo-600 px-4 py-2 text-sm font-semibold text-white">Save type</button><button type="button" onClick={() => setDefaultTypeForm(null)} className="rounded border bg-white px-4 py-2 text-sm">Cancel</button></div>
          </form>}
          {typeDeleteTarget && <div role="alertdialog" aria-modal="true" className="mt-4 rounded-lg border border-rose-300 bg-rose-50 p-4 text-sm">
            <p>Delete <strong>{typeDeleteTarget.name}</strong>? This cannot be undone.</p>
            <div className="mt-3 flex gap-2"><button type="button" onClick={() => void changeDefaultTypeStatus(typeDeleteTarget, 'delete')} className="rounded bg-rose-700 px-3 py-2 text-white">Delete type</button><button type="button" onClick={() => setTypeDeleteTarget(null)} className="rounded border bg-white px-3 py-2">Cancel</button></div>
          </div>}
          <div className="responsive-table mt-5">
            <table className="min-w-full text-left text-sm"><thead className="border-b bg-slate-50"><tr>{['Name', 'Description', 'Applies to', 'Active', 'Used in', 'Last changed by', 'Actions'].map((heading) => <th key={heading} className="px-3 py-3 font-medium">{heading}</th>)}</tr></thead>
              <tbody>{managedTypes.map((type) => <tr key={type.id} className="border-b border-slate-100">
                <td data-label="Name" className="px-3 py-3 font-medium">{type.name}</td><td data-label="Description" className="px-3 py-3">{type.description || '—'}</td>
                <td data-label="Applies to" className="px-3 py-3">{type.appliesTo === 'all_companies' ? 'All companies' : type.appliesTo}</td>
                <td data-label="Active" className="px-3 py-3">{type.active ? 'Active' : 'Inactive'}</td><td data-label="Used in" className="px-3 py-3">{type.usedIn ?? 0} defaults</td><td data-label="Last changed by" className="px-3 py-3">{type.lastChangedBy || 'System'}</td>
                <td data-label="Actions" className="px-3 py-3"><div className="flex flex-wrap gap-2"><button type="button" onClick={() => setDefaultTypeForm({ id: type.id, name: type.name, description: type.description ?? '', appliesTo: type.appliesTo ?? 'all_companies' })} className="rounded border px-2 py-1 text-xs">Edit</button><button type="button" onClick={() => void changeDefaultTypeStatus(type, type.active ? 'deactivate' : 'activate')} className="rounded border px-2 py-1 text-xs">{type.active ? 'Deactivate' : 'Reactivate'}</button>{!type.usedIn && <button type="button" onClick={() => setTypeDeleteTarget(type)} className="rounded border border-rose-200 px-2 py-1 text-xs text-rose-700">Delete</button>}</div></td>
              </tr>)}{!managedTypes.length && <tr><td colSpan={7} className="px-3 py-8 text-center text-slate-500">No default types have been configured.</td></tr>}</tbody>
            </table>
          </div>
        </section>}
      </section>}

      {view === 'members' && <section className="space-y-6">
        <div className="grid gap-4 sm:grid-cols-2">
          <div className="rounded-2xl border border-slate-200 bg-white p-5 shadow-sm">
            <p className="text-sm text-slate-500">Active placement cell</p>
            <p className="mt-2 text-3xl font-bold">{members.filter((member) => member.active && member.team === 'placement_cell').length}</p>
          </div>
          <div className="rounded-2xl border border-slate-200 bg-white p-5 shadow-sm">
            <p className="text-sm text-slate-500">Active summer internship</p>
            <p className="mt-2 text-3xl font-bold">{members.filter((member) => member.active && member.team === 'summer_internship').length}</p>
          </div>
        </div>
        <section className="rounded-2xl border border-slate-200 bg-white p-6 shadow-sm">
          <div className="flex flex-wrap items-center justify-between gap-3">
            <div><h2 className="text-xl font-semibold">Members</h2><p className="mt-1 text-sm text-slate-500">Manage placement and summer internship members and their sign-in access.</p></div>
            <div className="flex flex-wrap gap-2">
              <button onClick={downloadMemberTemplate} className="rounded-lg border px-3 py-2 text-sm">Download sample Excel</button>
              <label className="inline-flex min-h-11 cursor-pointer items-center rounded-lg border border-indigo-200 bg-indigo-50 px-3 py-2 text-sm font-medium text-indigo-700 hover:bg-indigo-100">Upload Excel<input type="file" accept=".csv,.xlsx,.xls" className="hidden" onChange={handleMemberUpload} /></label>
              <button onClick={() => { setMemberForm({ name: '', team: 'placement_cell', email: '' }); setMemberReason(''); }} className="rounded-lg bg-indigo-600 px-3 py-2 text-sm font-semibold text-white">Add member</button>
            </div>
          </div>

          <div className="sticky-filters mt-5 grid gap-3 bg-white py-2 sm:grid-cols-3">
            <label className="text-sm">Search<input value={memberSearch} onChange={(event) => setMemberSearch(event.target.value)} placeholder="Name or email" className="mt-1 w-full rounded border px-3 py-2" /></label>
            <label className="text-sm">Team<select value={memberTeamFilter} onChange={(event) => setMemberTeamFilter(event.target.value)} className="mt-1 w-full rounded border px-3 py-2"><option value="">All teams</option><option value="placement_cell">Placement cell</option><option value="summer_internship">Summer internship</option></select></label>
            <label className="text-sm">Status<select value={memberStatusFilter} onChange={(event) => setMemberStatusFilter(event.target.value)} className="mt-1 w-full rounded border px-3 py-2"><option value="all">All statuses</option><option value="active">Active</option><option value="inactive">Inactive</option></select></label>
          </div>

          {memberForm && <form onSubmit={saveMember} className="mt-5 rounded-xl border border-indigo-100 bg-indigo-50/60 p-4">
            <div className="mb-3 flex items-center justify-between"><h3 className="font-semibold">{memberForm.memberId ? 'Edit member' : 'Add member'}</h3><button type="button" onClick={() => setMemberForm(null)} className="rounded border bg-white px-2 py-1 text-sm">Cancel</button></div>
            <div className="grid gap-3 sm:grid-cols-3">
              <label className="text-sm">Name *<input required value={memberForm.name} onChange={(event) => setMemberForm({ ...memberForm, name: event.target.value })} className="mt-1 w-full rounded border px-3 py-2" /></label>
              <label className="text-sm">Team *<select required value={memberForm.team} onChange={(event) => setMemberForm({ ...memberForm, team: event.target.value })} className="mt-1 w-full rounded border px-3 py-2"><option value="placement_cell">Placement cell</option><option value="summer_internship">Summer internship</option></select></label>
              <label className="text-sm">Email (optional)<input type="email" value={memberForm.email} onChange={(event) => setMemberForm({ ...memberForm, email: event.target.value })} placeholder="name@college.edu" className="mt-1 w-full rounded border px-3 py-2" /><span className="mt-1 block text-xs text-slate-500">An email enables Google sign-in when active.</span></label>
            </div>
            {memberForm.memberId && <label className="mt-3 block text-sm">Reason for edit *<input required value={memberReason} onChange={(event) => setMemberReason(event.target.value)} className="mt-1 w-full rounded border px-3 py-2" /></label>}
            <button type="submit" className="mt-4 rounded-lg bg-indigo-600 px-4 py-2 font-semibold text-white">Save member</button>
          </form>}

          {bulkPreview && <div className="mt-5 rounded-xl border border-amber-200 bg-amber-50 p-4">
            <div className="flex flex-wrap items-center justify-between gap-3"><div><h3 className="font-semibold">Bulk upload preview</h3><p className="mt-1 text-sm">{bulkPreview.counts.add} rows to add · {bulkPreview.counts.update} to update · {bulkPreview.counts.skip} skipped</p></div><div className="flex gap-2"><button onClick={() => void applyMemberPreview()} disabled={!bulkPreview.counts.add && !bulkPreview.counts.update} className="rounded bg-indigo-600 px-3 py-2 text-sm text-white disabled:bg-slate-300">Apply import</button><button onClick={() => { setBulkPreview(null); setBulkRows([]); }} className="rounded border bg-white px-3 py-2 text-sm">Cancel</button></div></div>
            <div className="responsive-table mt-3 max-h-72 overflow-auto"><table className="min-w-full text-left text-xs"><thead><tr>{['Row', 'Action', 'Name', 'Team', 'Email', 'Reason'].map((heading) => <th key={heading} className="px-2 py-2">{heading}</th>)}</tr></thead><tbody>{bulkPreview.rows.map((row) => <tr key={row.row} className="border-t border-amber-200"><td data-label="Row" className="px-2 py-2">{row.row}</td><td data-label="Action" className="px-2 py-2 capitalize">{row.action}</td><td data-label="Name" className="px-2 py-2">{row.name}</td><td data-label="Team" className="px-2 py-2">{row.team}</td><td data-label="Email" className="px-2 py-2">{row.email || '—'}</td><td data-label="Reason" className="px-2 py-2">{row.reason}</td></tr>)}</tbody></table></div>
          </div>}

          <p className="mt-5 text-sm text-slate-500">{filteredMembers.length} results</p>
          <div className="responsive-table mt-2">
            <table className="min-w-full text-left text-sm">
              <thead className="border-b bg-slate-50 text-slate-600"><tr>{['Name', 'Team', 'Email', 'Status', 'Added on', 'Added by', 'Actions'].map((heading) => <th key={heading} className="px-3 py-3 font-medium">{heading}</th>)}</tr></thead>
              <tbody>{filteredMembers.map((member) => {
                const self = member.email?.toLowerCase() === accountEmail;
                return <tr key={member.memberId} className="border-b">
                  <td data-label="Name" className="px-3 py-3 font-medium">{member.name}{self && <span className="ml-2 text-xs text-indigo-600">(you)</span>}</td>
                  <td data-label="Team" className="px-3 py-3">{teamLabel(member.team)}</td><td data-label="Email" className="px-3 py-3">{member.email || '—'}</td>
                  <td data-label="Status" className="px-3 py-3"><span className={`status-badge ${member.active ? 'status-success' : 'status-danger'}`}>{member.active ? 'Active' : 'Inactive'}</span></td>
                  <td data-label="Added on" className="px-3 py-3">{member.createdAt ? new Date(member.createdAt).toLocaleDateString() : '—'}</td><td data-label="Added by" className="px-3 py-3">{member.createdByEmail || '—'}</td>
                  <td data-label="Actions" className="px-3 py-3"><div className="flex flex-wrap gap-2"><button onClick={() => { setMemberForm({ memberId: member.memberId, name: member.name, team: isSelectableTeam(member.team) ? member.team : 'placement_cell', email: member.email || '' }); setMemberReason(''); }} className="rounded border px-2 py-1 text-xs">Edit</button>{member.active ? <button disabled={self} title={self ? 'You cannot deactivate your own account.' : ''} onClick={() => void changeMemberStatus(member, 'deactivate')} className="rounded border border-rose-200 px-2 py-1 text-xs text-rose-700 disabled:opacity-40">Deactivate</button> : <button onClick={() => void changeMemberStatus(member, 'reactivate')} className="rounded border px-2 py-1 text-xs">Reactivate</button>}</div></td>
                </tr>;
              })}{!filteredMembers.length && <tr><td colSpan={7} className="px-3 py-8 text-center text-slate-500">No members match these filters.</td></tr>}</tbody>
            </table>
          </div>
        </section>
      </section>}

      {memberStatusTarget && <div className="fixed inset-0 z-[60] flex items-center justify-center bg-slate-900/50 p-4" role="presentation">
        <form onSubmit={submitMemberStatus} role="dialog" aria-modal="true" aria-labelledby="member-status-title" className="w-full max-w-lg rounded-2xl bg-white p-6 shadow-xl">
          <h2 id="member-status-title" className="text-xl font-semibold">{memberStatusTarget.action === 'deactivate' ? 'Deactivate' : 'Reactivate'} {memberStatusTarget.member.name}?</h2>
          {memberStatusTarget.action === 'deactivate' && <p className="mt-2 text-sm text-slate-600">They will lose sign-in access.</p>}
          <label className="mt-4 block text-sm font-medium">Reason *
            <textarea autoFocus required value={memberStatusReason} onChange={(event) => setMemberStatusReason(event.target.value)} maxLength={1000} rows={3} className="mt-1 w-full rounded-lg border border-slate-300 px-3 py-2" />
          </label>
          <div className="mt-5 flex justify-end gap-2"><button type="button" className="min-h-11 rounded-lg border px-4 py-2" onClick={() => setMemberStatusTarget(null)}>Cancel</button><button type="submit" className="min-h-11 rounded-lg bg-indigo-600 px-4 py-2 font-semibold text-white">{memberStatusTarget.action === 'deactivate' ? 'Deactivate member' : 'Reactivate member'}</button></div>
        </form>
      </div>}

      {defaultActionTarget && detail && <div className="fixed inset-0 z-[60] flex items-center justify-center bg-slate-900/50 p-4" role="presentation">
        <form onSubmit={submitDefaultAction} role="dialog" aria-modal="true" aria-labelledby="default-action-title" className="w-full max-w-lg rounded-2xl bg-white p-6 shadow-xl">
          <h2 id="default-action-title" className="text-xl font-semibold">{defaultActionTarget === 'restore' ? 'Restore default' : `${defaultActionTarget === 'lock' ? 'Lock' : 'Unlock'} default`}</h2>
          <p className="mt-2 text-sm text-slate-600">{detail.record.company} · {detail.record.rollNo}</p>
          <label className="mt-4 block text-sm font-medium">Reason *
            <textarea autoFocus required value={defaultActionReason} onChange={(event) => setDefaultActionReason(event.target.value)} maxLength={2000} rows={3} className="mt-1 w-full rounded-lg border border-slate-300 px-3 py-2" />
          </label>
          <div className="mt-5 flex justify-end gap-2"><button type="button" className="min-h-11 rounded-lg border px-4 py-2" onClick={() => setDefaultActionTarget(null)}>Cancel</button><button type="submit" className="min-h-11 rounded-lg bg-indigo-600 px-4 py-2 font-semibold text-white">Confirm</button></div>
        </form>
      </div>}

      {removeDefaultTarget && <div className="fixed inset-0 z-[60] flex items-center justify-center bg-slate-900/50 p-4" role="presentation">
        <form onSubmit={submitRemoveDefault} role="dialog" aria-modal="true" aria-labelledby="remove-default-title" className="w-full max-w-lg rounded-2xl bg-white p-6 shadow-xl">
          <h2 id="remove-default-title" className="text-xl font-semibold">Remove default?</h2>
          <p className="mt-2 text-sm text-slate-600">The default for {removeDefaultTarget.rollNo} will be marked removed; its history is retained.</p>
          <label className="mt-4 block text-sm font-medium">Removal reason *
            <textarea autoFocus required value={removeDefaultReason} onChange={(event) => setRemoveDefaultReason(event.target.value)} maxLength={2000} rows={2} className="mt-1 w-full rounded-lg border border-slate-300 px-3 py-2" />
          </label>
          <label className="mt-3 block text-sm font-medium">DC reference *
            <input required value={removeDcReference} onChange={(event) => setRemoveDcReference(event.target.value)} maxLength={200} className="mt-1 min-h-11 w-full rounded-lg border border-slate-300 px-3 py-2" />
          </label>
          <div className="mt-5 flex justify-end gap-2"><button type="button" className="min-h-11 rounded-lg border px-4 py-2" onClick={() => setRemoveDefaultTarget(null)}>Cancel</button><button type="submit" className="min-h-11 rounded-lg bg-rose-700 px-4 py-2 font-semibold text-white">Remove default</button></div>
        </form>
      </div>}

      {studentLockAction && studentDetail && <div className="fixed inset-0 z-[60] flex items-center justify-center bg-slate-900/50 p-4" role="presentation">
        <form onSubmit={submitStudentLock} role="dialog" aria-modal="true" aria-labelledby="student-lock-title" className="w-full max-w-lg rounded-2xl bg-white p-6 shadow-xl">
          <h2 id="student-lock-title" className="text-xl font-semibold">{studentLockAction === 'lock' ? 'Lock' : 'Unlock'} {studentDetail.student.name}?</h2>
          <label className="mt-4 block text-sm font-medium">Reason *
            <textarea autoFocus required value={studentLockReason} onChange={(event) => setStudentLockReason(event.target.value)} maxLength={2000} rows={3} className="mt-1 w-full rounded-lg border border-slate-300 px-3 py-2" />
          </label>
          <div className="mt-5 flex justify-end gap-2"><button type="button" className="min-h-11 rounded-lg border px-4 py-2" onClick={() => setStudentLockAction(null)}>Cancel</button><button type="submit" className="min-h-11 rounded-lg bg-indigo-600 px-4 py-2 font-semibold text-white">Confirm</button></div>
        </form>
      </div>}

      {detail && <div className="fixed inset-0 z-50 flex items-start justify-center overflow-y-auto bg-slate-900/50 p-4" role="dialog" aria-modal="true" aria-label="Default detail">
        <div className="my-8 w-full max-w-3xl rounded-2xl bg-white p-6 shadow-xl">
          <div className="flex items-start justify-between gap-4"><div><p className="text-sm text-slate-500">{detail.record.rollNo} · {detail.record.status}</p><h2 className="text-2xl font-bold">{detail.record.company} — {detail.record.typeName}</h2><p className="mt-1 text-sm text-slate-600">Logged by {detail.record.loggedByName} ({teamLabel(detail.record.team)}) · Entered by {detail.record.enteredByEmail}</p></div><button onClick={() => setDetail(null)} className="rounded border px-3 py-1">Close</button></div>
          <dl className="mt-5 grid gap-3 sm:grid-cols-2">{[['Date', detail.record.date], ['Date logged', new Date(detail.record.createdAt).toLocaleString()], ['Last updated', new Date(detail.record.updatedAt).toLocaleString()], ['Last updated by', detail.record.updatedByName], ['Notes', detail.record.notes], ['DC reference', detail.record.dcReference || '—'], ['DC status', detail.record.dcStatus ?? 'Not scheduled'], ['Deferred until', detail.record.deferral?.until === 'next_dc' ? 'Next DC' : detail.record.deferral?.until ?? '—'], ['Deferral reason', detail.record.deferral?.reason ?? '—'], ['Discussion reason', detail.record.discussion?.reason ?? '—'], ['Discussion date', detail.record.discussion?.discussionDate ?? '—']].map(([label, value]) => <div key={label}><dt className="text-xs uppercase text-slate-500">{label}</dt><dd className="mt-1 whitespace-pre-wrap">{value}</dd></div>)}</dl>
          {detail.record.discussion?.notes?.length ? <section className="mt-4 rounded-lg border p-4"><h3 className="font-semibold">Team discussion notes</h3><ul className="mt-2 space-y-2">{detail.record.discussion.notes.map((note, index) => <li key={`${note.at}-${index}`} className="rounded bg-slate-50 p-3 text-sm"><p>{note.text}</p><p className="mt-1 text-xs text-slate-500">{note.byName} · {new Date(note.at).toLocaleString()}</p></li>)}</ul></section> : null}
          {detail.record.deferral?.history?.length ? <section className="mt-4 rounded-lg border p-4"><h3 className="font-semibold">Deferral history</h3><ul className="mt-2 list-disc space-y-1 pl-5 text-sm">{detail.record.deferral.history.map((item, index) => <li key={`${item.at}-${index}`}>{item.until === 'next_dc' ? 'Next DC' : item.until} — {item.reason} <span className="text-xs text-slate-500">({item.byName ?? 'Unknown member'} · {new Date(item.at).toLocaleString()})</span></li>)}</ul></section> : null}
          {detail.record.status === 'logged' && detail.record.dcStatus !== 'To discuss' && detail.record.dcStatus !== 'Deferred' && <form onSubmit={flagDefaultForDiscussion} className="mt-4 rounded-lg border border-indigo-100 bg-indigo-50/50 p-4">
            <h3 className="font-semibold">Flag for team discussion</h3>
            <div className="mt-3 grid gap-3 sm:grid-cols-2">
              <label className="text-sm sm:col-span-2">Reason *<textarea required value={discussionReason} onChange={(event) => setDiscussionReason(event.target.value)} maxLength={2000} rows={2} className="mt-1 w-full rounded border px-3 py-2" /></label>
              <label className="text-sm">Discussion date (optional)<input type="date" value={discussionDate} onChange={(event) => setDiscussionDate(event.target.value)} className="mt-1 w-full rounded border px-3 py-2" /></label>
              <button type="submit" className="self-end rounded bg-indigo-600 px-3 py-2 text-white">Flag for discussion</button>
            </div>
          </form>}
          {detail.record.status !== 'removed' && <details className="mt-5 rounded-lg border p-4"><summary className="cursor-pointer font-semibold">Edit default</summary><div className="mt-4 grid gap-3 sm:grid-cols-2">
            <label className="text-sm">Company<input value={editValues.company} onChange={(event) => setEditValues({ ...editValues, company: event.target.value })} className="mt-1 w-full rounded border px-2 py-2" /></label>
            <label className="text-sm">Type<select value={editValues.typeId} onChange={(event) => setEditValues({ ...editValues, typeId: event.target.value })} className="mt-1 w-full rounded border px-2 py-2">{defaultTypes.map((type) => <option key={type.id} value={type.id}>{type.name}</option>)}</select></label>
            <label className="text-sm">Date<input type="date" value={editValues.date} onChange={(event) => setEditValues({ ...editValues, date: event.target.value })} className="mt-1 w-full rounded border px-2 py-2" /></label>
              <label className="text-sm">Logged by<select value={editValues.loggedByMemberId} onChange={(event) => setEditValues({ ...editValues, loggedByMemberId: event.target.value })} className="mt-1 w-full rounded border px-2 py-2">{members.filter((member) => member.active && isSelectableTeam(member.team) || member.memberId === editValues.loggedByMemberId).map((member) => <option key={member.memberId} value={member.memberId}>{member.name} — {teamLabel(member.team)}{member.active ? '' : ' (inactive)'}</option>)}</select></label>
            <label className="text-sm sm:col-span-2">Notes<textarea value={editValues.notes} onChange={(event) => setEditValues({ ...editValues, notes: event.target.value })} rows={2} className="mt-1 w-full rounded border px-2 py-2" /></label>
            <label className="text-sm sm:col-span-2">DC reference<input value={editValues.dcReference} onChange={(event) => setEditValues({ ...editValues, dcReference: event.target.value })} className="mt-1 w-full rounded border px-2 py-2" /></label>
            <label className="text-sm sm:col-span-2">Reason (required)<input value={editReason} onChange={(event) => setEditReason(event.target.value)} className="mt-1 w-full rounded border px-2 py-2" /></label>
            <button disabled={detail.record.status === 'locked'} onClick={() => void updateDefault('edit')} className="rounded bg-indigo-600 px-3 py-2 text-white disabled:bg-slate-300">Save edit</button>
          </div>{detail.record.status === 'locked' && <p className="mt-2 text-sm text-amber-700">Unlock before editing.</p>}</details>}
          <div className="mt-4 flex flex-wrap gap-2">
            {detail.record.status === 'logged' && <button onClick={() => void updateDefault('lock')} className="rounded border px-3 py-2">Lock</button>}
            {detail.record.status === 'locked' && <button onClick={() => void updateDefault('unlock')} className="rounded border px-3 py-2">Unlock</button>}
            {detail.record.status === 'logged' && <button onClick={() => void removeDefault(detail.record)} className="rounded border border-rose-200 bg-rose-50 px-3 py-2 text-rose-700">Remove</button>}
            {detail.record.status === 'removed' && <><input placeholder="Reason to restore" value={editReason} onChange={(event) => setEditReason(event.target.value)} className="rounded border px-3 py-2" /><button onClick={() => void updateDefault('restore')} className="rounded border px-3 py-2">Restore</button></>}
          </div>
          <h3 className="mt-7 font-semibold">History</h3><Timeline events={detail.events} />
          <div className="mt-5 border-t pt-4"><label className="block text-sm font-semibold">Add DC note<textarea rows={2} value={dcNote} onChange={(event) => setDcNote(event.target.value)} className="mt-2 w-full rounded border px-3 py-2" /></label><button onClick={() => void saveDcNote()} disabled={!dcNote.trim()} className="mt-2 rounded bg-indigo-600 px-3 py-2 text-white disabled:bg-slate-300">Add note</button><p className="mt-1 text-xs text-slate-500">Only active placement/summer internship members can add notes.</p></div>
        </div>
      </div>}
        </div>
      </div>
    </div>
    </main>
  );
}

function Timeline({ events }: { events: EventRecord[] }) {
  if (!events.length) return <p className="mt-3 text-sm text-slate-500">No history events.</p>;
  const labels: Record<string, string> = {
    logged: 'Default logged', edited: 'Default updated', locked: 'Default locked', unlocked: 'Default unlocked',
    removed: 'Default removed', restored: 'Default restored', dc_auto_added: 'Automatically added to DC',
    dc_scheduled: 'Scheduled for DC', dc_unscheduled: 'Removed from DC', dc_attendance: 'DC attendance updated',
    dc_decision: 'DC decision recorded', dc_rescheduled: 'DC rescheduled', dc_cancelled: 'DC cancelled',
    dc_note: 'DC note added', discuss_flagged: 'Flagged for discussion', discuss_note: 'Discussion note added',
    student_locked: 'Student locked', student_unlocked: 'Student unlocked', student_batch_updated: 'Student batch updated', deferral_due: 'Deferral due',
  };
  const grouped = new Map<string, EventRecord[]>();
  for (const event of events) {
    const timeKey = new Date(event.at).toISOString();
    const key = event.type === 'dc_auto_added' ? `${timeKey}|${event.type}|${event.dcReference ?? ''}` : event.eventId;
    grouped.set(key, [...(grouped.get(key) ?? []), event]);
  }
  const days = new Map<string, EventRecord[][]>();
  for (const group of Array.from(grouped.values())) {
    const day = new Date(group[0].at).toLocaleDateString();
    days.set(day, [...(days.get(day) ?? []), group]);
  }
  return <div className="mt-4 space-y-6">{Array.from(days.entries()).map(([day, groups]) => <section key={day} aria-label={`Activity on ${day}`}>
    <h3 className="sticky top-0 z-10 bg-white py-2 text-sm font-semibold text-slate-600">{day}</h3>
    <ol className="space-y-2 border-l-2 border-indigo-100 pl-4">{groups.map((group) => {
      const event = group[0];
      const title = labels[event.type] ?? event.type.replaceAll('_', ' ');
      const expandable = group.length > 1 || Boolean(event.reason || event.note || event.dcReference || event.changes || event.enteredByEmail);
      return <li key={event.eventId} className="relative rounded-lg border border-slate-200 bg-slate-50 px-3 py-2">
        <span className="absolute -left-[23px] top-4 h-2.5 w-2.5 rounded-full bg-indigo-500" />
        <div className="flex flex-wrap items-center justify-between gap-2">
          {expandable ? <details><summary className="cursor-pointer font-medium">{title}{group.length > 1 ? ` · ${group.length} defaults` : ''}</summary>
            <div className="mt-2 space-y-2 text-sm">
              {group.map((item) => <div key={item.eventId} className="border-t pt-2">
                <p>{item.studentName ?? item.rollNo} ({item.rollNo}){item.studentBatch ? ` · ${item.studentBatch === 'junior' ? 'Junior' : 'Senior'} batch` : ' · Batch not set'}</p>
                <p className="text-xs text-slate-500">{item.actor === 'system' ? 'System' : `${item.memberName} · ${teamLabel(item.team)}`}</p>
                {item.reason && <p><strong>Reason:</strong> {item.reason}</p>}
                {item.note && <p><strong>Note:</strong> {item.note}</p>}
                {item.dcReference && <p><strong>DC:</strong> {item.dcReference}</p>}
                {item.enteredByEmail && <p className="text-xs text-slate-500">Account: {item.enteredByEmail}</p>}
                {item.changes && Object.entries(item.changes).map(([field, change]) => <p key={field} className="text-xs"><strong>{field.replaceAll('_', ' ')}:</strong> {String(change.from)} → {String(change.to)}</p>)}
              </div>)}
            </div>
          </details> : <strong className="font-medium">{title}</strong>}
          <time className="text-xs text-slate-500">{new Date(event.at).toLocaleTimeString()}</time>
        </div>
        {group.length === 1 && <p className="mt-1 text-sm text-slate-600">{event.studentName ?? event.rollNo} ({event.rollNo}) · {event.studentBatch ? `${event.studentBatch === 'junior' ? 'Junior' : 'Senior'} batch` : 'Batch not set'}</p>}
        <p className="mt-1 text-xs text-slate-500">{event.actor === 'system' ? 'System' : event.memberName}</p>
      </li>;
    })}</ol>
  </section>)}</div>;
}
