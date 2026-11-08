'use client';

import { ChangeEvent, FormEvent, useCallback, useEffect, useMemo, useState } from 'react';
import { signIn, signOut, useSession } from 'next-auth/react';
import * as XLSX from 'xlsx';
import { downloadWorkbook, MAX_EXCEL_BYTES, MAX_EXCEL_FILE_MESSAGE } from '@/lib/excel-download';
import DCSummaryTab from './dc-summary';
import DCTab from './dc-tab';

type Student = {
  rollNo: string;
  name: string;
  numberOfDefaults?: number;
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

type DefaultType = { id: string; name: string };

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

type ViewName = 'defaults' | 'students' | 'dc' | 'members' | 'activity';

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
  const [defaultStatusFilter, setDefaultStatusFilter] = useState('all');
  const [studentDetail, setStudentDetail] = useState<{ student: Student; defaults: StudentHistoryDefault[] } | null>(null);
  const [activity, setActivity] = useState<EventRecord[]>([]);
  const [activityFilters, setActivityFilters] = useState({ memberId: '', team: '', type: '', from: '', to: '', rollNo: '' });
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

  const activeDefaults = useMemo(() => {
    const query = defaultSearch.trim().toLowerCase();
    return defaults.filter((record) =>
      (showRemoved || record.status !== 'removed') &&
      (defaultStatusFilter === 'all' || record.status === defaultStatusFilter) &&
      (!query || [record.rollNo, record.company, record.typeName, record.loggedByName].some((value) => value.toLowerCase().includes(query))),
    );
  }, [defaults, showRemoved, defaultSearch, defaultStatusFilter]);
  const includeRemovedDefaults = showRemoved || defaultStatusFilter === 'removed';
  const selectableMembers = useMemo(
    () => members.filter((member) => member.active && isSelectableTeam(member.team)),
    [members],
  );
  const stats = useMemo(() => ({
    total: defaults.filter((record) => record.status !== 'removed').length,
    logged: defaults.filter((record) => record.status === 'logged').length,
    locked: defaults.filter((record) => record.status === 'locked').length,
  }), [defaults]);
  const matchedStudent = useMemo(
    () => students.find((student) => normalizeRollNo(student.rollNo) === normalizeRollNo(studentSearch)),
    [students, studentSearch],
  );

  const loadDashboard = useCallback(async () => {
    setDashboardLoading(true);
    try {
      const [studentsResponse, defaultsResponse, typesResponse, membersResponse] = await Promise.all([
        fetch('/api/students', { cache: 'no-store' }),
        fetch(`/api/defaults?includeRemoved=${includeRemovedDefaults}`, { cache: 'no-store' }),
        fetch('/api/default-types', { cache: 'no-store' }),
        fetch('/api/members', { cache: 'no-store' }),
      ]);
      const [studentsData, defaultsData, typesData, membersData] = await Promise.all([
        responseJson(studentsResponse),
        responseJson(defaultsResponse),
        responseJson(typesResponse),
        responseJson(membersResponse),
      ]);
      setStudents(studentsData);
      setDefaults(defaultsData);
      setDefaultTypes(typesData);
      setMembers(membersData);
      setTypeId((current) => current || typesData[0]?.id || '');
      const selectableMembers = membersData.filter((member: Member) => member.active && isSelectableTeam(member.team));
      const currentMember = selectableMembers.find((member: Member) => member.email?.toLowerCase() === accountEmail);
      setLoggedByMemberId((current) =>
        selectableMembers.some((member: Member) => member.memberId === current)
          ? current
          : currentMember?.memberId || selectableMembers[0]?.memberId || '',
      );
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
        body: JSON.stringify({ rollNo: selectedStudent.rollNo, company, typeId, date, notes, loggedByMemberId }),
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
      const roster = rows.map((row) => ({
        name: String(row.Name ?? row.name ?? '').trim(),
        rollNo: String(row['Roll Number'] ?? row['Roll number'] ?? row.rollNo ?? '').trim(),
        numberOfDefaults: (row['Number of defaults'] ?? row.numberOfDefaults) === '' ||
          (row['Number of defaults'] ?? row.numberOfDefaults) === undefined
          ? undefined
          : Number(row['Number of defaults'] ?? row.numberOfDefaults),
        lockPenaltyStatus: String(row['Lock / penalty status'] ?? row.lockPenaltyStatus ?? '').trim(),
      })).filter((row) => row.name && row.rollNo);
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
    downloadWorkbook(workbook, 'member-template.xlsx');
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
    const query = new URLSearchParams(Object.entries(activityFilters).filter(([, value]) => value));
    try {
      setActivity(await responseJson(await fetch(`/api/activity?${query.toString()}`, { cache: 'no-store' })));
    } catch (error) {
      setStatusMessage(error instanceof Error ? error.message : 'Unable to load activity.');
    }
  }

  useEffect(() => {
    if (sessionStatus === 'authenticated' && view === 'activity') void loadActivity();
  }, [sessionStatus, view, activityFilters]);

  async function exportWorkbook() {
    try {
      const query = new URLSearchParams(Object.entries(activityFilters).filter(([, value]) => value));
      const history: EventRecord[] = await responseJson(await fetch(`/api/activity?${query.toString()}`, { cache: 'no-store' }));
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
        }));
      const details = visibleDefaults.map((record) => ({
        'Roll number': record.rollNo,
        Student: students.find((student) => normalizeRollNo(student.rollNo) === normalizeRollNo(record.rollNo))?.name ?? record.rollNo,
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
        'Event type': event.type,
        Member: event.memberName,
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
            ] as Array<[ViewName, string]>).map(([key, label]) => (
              <button key={key} type="button" onClick={() => { setView(key); setSidebarOpen(false); if (key === 'dc') setDcSubView('planner'); }} className={`flex w-full items-center gap-3 rounded-xl px-3 py-3 text-left text-sm font-medium transition ${view === key ? 'bg-indigo-600 text-white shadow-sm' : 'text-slate-300 hover:bg-slate-800 hover:text-white'}`}>
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
                  <p className="text-[10px] font-semibold uppercase tracking-[0.25em] text-indigo-600">{view === 'dc' ? 'Disciplinary Committee' : view === 'defaults' ? 'Defaults' : view === 'students' ? 'Students' : view === 'members' ? 'Members' : 'Activity'}</p>
                  <h2 className="mt-1 text-2xl font-bold text-slate-900">{view === 'dc' ? (dcSubView === 'summary' ? 'DC Summary' : 'DC Planner') : view === 'defaults' ? 'Defaults' : view === 'students' ? 'Student roster' : view === 'members' ? 'Members' : 'Activity log'}</h2>
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
        <section className="mb-8 grid gap-4 md:grid-cols-3">
          {[['Total defaults', stats.total, 'text-slate-900'], ['Logged', stats.logged, 'text-emerald-600'], ['Locked', stats.locked, 'text-amber-600']].map(([label, count, color]) => (
            <div key={label} className="rounded-2xl border border-slate-200 bg-white p-5 shadow-sm"><p className="text-sm text-slate-500">{label}</p><p className={`mt-3 text-3xl font-bold ${color}`}>{count}</p></div>
          ))}
        </section>
        <section className="mb-8 grid gap-6 lg:grid-cols-[1.2fr_0.8fr]">
          <form onSubmit={handleSubmit} className="rounded-2xl border border-slate-200 bg-white p-6 shadow-sm">
            <div className="mb-5 flex items-center justify-between"><h2 className="text-xl font-semibold text-slate-900">Log a default</h2><span className="rounded-full bg-indigo-100 px-2.5 py-1 text-xs font-medium text-indigo-700">Roll number only</span></div>
            <div className="space-y-4">
              <div>
                <label htmlFor="roll-number" className="mb-2 block text-sm font-medium text-slate-700">Student roll number</label>
                <div className="flex gap-2">
                  <input id="roll-number" aria-describedby={lookupError ? 'roll-number-error' : undefined} autoFocus value={rollNumber} onChange={(event) => { setRollNumber(event.target.value); setSelectedStudent(null); setLookupError(''); }} placeholder="e.g. 2026/001" className="w-full rounded-lg border border-slate-300 bg-slate-50 px-3 py-2.5" />
                  <button type="button" onClick={handleLookup} className="rounded-lg bg-slate-900 px-4 py-2.5 text-sm font-medium text-white">Look up</button>
                </div>
                {lookupError && <p id="roll-number-error" role="alert" className="mt-2 text-sm text-rose-600">{lookupError}</p>}
                {selectedStudent && <p className="mt-3 rounded-lg bg-emerald-50 px-3 py-2 text-sm text-emerald-800">Student: <strong>{selectedStudent.name}</strong></p>}
              </div>
              <label className="block text-sm font-medium text-slate-700">Company<input required value={company} onChange={(event) => setCompany(event.target.value)} className="mt-2 w-full rounded-lg border border-slate-300 bg-slate-50 px-3 py-2.5" /></label>
              <div className="grid gap-4 sm:grid-cols-2">
                <label className="block text-sm font-medium text-slate-700">Default type<select required value={typeId} onChange={(event) => setTypeId(event.target.value)} className="mt-2 w-full rounded-lg border border-slate-300 bg-slate-50 px-3 py-2.5">{defaultTypes.map((type) => <option key={type.id} value={type.id}>{type.name}</option>)}</select></label>
                <label className="block text-sm font-medium text-slate-700">Date<input required type="date" value={date} onChange={(event) => setDate(event.target.value)} className="mt-2 w-full rounded-lg border border-slate-300 bg-slate-50 px-3 py-2.5" /></label>
              </div>
              <label className="block text-sm font-medium text-slate-700">Logged by (member)<select required value={loggedByMemberId} onChange={(event) => setLoggedByMemberId(event.target.value)} className="mt-2 w-full rounded-lg border border-slate-300 bg-slate-50 px-3 py-2.5"><option value="">Select a member</option>{selectableMembers.map((member) => <option key={member.memberId} value={member.memberId}>{member.name} — {teamLabel(member.team)}{member.email?.toLowerCase() === accountEmail ? ' (you)' : ''}</option>)}</select></label>
              {!selectableMembers.length && <p className="text-sm text-amber-700">Add or upload active placement/summer members to enable logging.</p>}
              <label className="block text-sm font-medium text-slate-700">Notes<textarea rows={3} value={notes} onChange={(event) => setNotes(event.target.value)} className="mt-2 w-full rounded-lg border border-slate-300 bg-slate-50 px-3 py-2.5" /></label>
              <div className="text-xs text-slate-500">Entered by (account): {accountEmail}</div>
              <button type="submit" disabled={isSubmitting || !selectedStudent || !selectableMembers.length} className="rounded-lg bg-indigo-600 px-4 py-2.5 text-sm font-semibold text-white disabled:bg-slate-300">{isSubmitting ? 'Logging...' : 'Log default'}</button>
            </div>
          </form>
          <aside className="rounded-2xl border border-slate-200 bg-white p-6 shadow-sm">
            <h2 className="text-xl font-semibold text-slate-900">Student roster</h2>
            <div className="mt-4 max-h-[440px] space-y-2 overflow-auto pr-1">{students.map((student) => <button key={student.rollNo} onClick={() => { setStudentSearch(student.rollNo); setView('students'); setStudentDetail(null); }} className="flex w-full items-center justify-between rounded-lg border border-slate-200 bg-slate-50 px-3 py-2 text-left hover:bg-indigo-50"><span><span className="block font-medium text-slate-800">{student.name}</span><span className="text-xs uppercase text-slate-500">{student.rollNo}</span></span><span className="text-right text-xs text-slate-600">{student.numberOfDefaults ?? 0} defaults<br />{student.lockPenaltyStatus ?? 'Clear'}</span></button>)}</div>
          </aside>
        </section>
        <section className="rounded-2xl border border-slate-200 bg-white p-6 shadow-sm">
          <div className="sticky-filters mb-4 flex flex-wrap items-end justify-between gap-3 bg-white py-2"><div><h2 className="text-xl font-semibold text-slate-900">Defaults</h2><p className="text-sm text-slate-500">{activeDefaults.length} results</p></div><div className="flex flex-wrap items-end gap-3"><label className="text-sm">Search<input value={defaultSearch} onChange={(event) => setDefaultSearch(event.target.value)} placeholder="Student, company, type…" className="mt-1 min-h-11 rounded-lg border px-3" /></label><label className="text-sm">Status<select value={defaultStatusFilter} onChange={(event) => setDefaultStatusFilter(event.target.value)} className="mt-1 min-h-11 rounded-lg border px-3"><option value="all">All statuses</option><option value="logged">Logged</option><option value="locked">Locked</option><option value="removed">Removed</option></select></label><label className="flex min-h-11 items-center gap-2 text-sm text-slate-600"><input type="checkbox" checked={showRemoved} onChange={(event) => setShowRemoved(event.target.checked)} />Show removed</label></div></div>
          <div className="responsive-table"><table className="min-w-full text-left text-sm"><thead className="border-b bg-slate-50 text-slate-600"><tr>{['Student', 'Company', 'Type', 'Logged by', 'Team', 'Date logged', 'Last updated by', 'Status', 'Actions'].map((header) => <th key={header} className="px-3 py-3 font-medium">{header}</th>)}</tr></thead><tbody>
            {activeDefaults.map((record) => <tr key={record.id} className="border-b align-top">
              <td data-label="Student" className="px-3 py-3"><button className="text-left" onClick={() => void openDefault(record)}><span className="block font-medium text-indigo-700">{students.find((student) => normalizeRollNo(student.rollNo) === normalizeRollNo(record.rollNo))?.name ?? record.rollNo}</span><span className="text-xs text-slate-500">{record.rollNo}</span></button></td>
              <td data-label="Company" className="px-3 py-3">{record.company}</td><td data-label="Type" className="px-3 py-3">{record.typeName}</td><td data-label="Logged by" className="px-3 py-3">{record.loggedByName}</td><td data-label="Team" className="px-3 py-3">{teamLabel(record.team)}</td><td data-label="Date logged" className="px-3 py-3">{new Date(record.createdAt).toLocaleString()}</td><td data-label="Last updated by" className="px-3 py-3">{record.updatedByName}</td>
              <td data-label="Status" className="px-3 py-3"><span className={`status-badge ${record.status === 'locked' ? 'status-pending' : record.status === 'removed' ? 'status-danger' : 'status-success'}`}>{record.status === 'logged' ? 'Logged' : record.status === 'locked' ? 'Locked' : 'Removed'}</span></td>
              <td data-label="Actions" className="px-3 py-3"><button onClick={() => void openDefault(record)} className="rounded border px-2 py-1 text-xs">Details</button></td>
            </tr>)}
            {!activeDefaults.length && <tr><td colSpan={9} className="px-4 py-8 text-center text-slate-500">{defaults.length ? 'No defaults match these filters.' : 'No defaults yet.'}</td></tr>}
          </tbody></table></div>
        </section>
      </>}

      {view === 'dc' && dcSubView === 'planner' && <DCTab onRefresh={loadDashboard} onOpenStudent={openStudentFromDc} />}
      {view === 'dc' && dcSubView === 'summary' && <DCSummaryTab />}

      {view === 'students' && <section className="rounded-2xl border border-slate-200 bg-white p-6 shadow-sm">
        <h2 className="text-xl font-semibold">Student history</h2>
        <div className="mt-4 flex items-end gap-2"><label className="w-full max-w-sm text-sm font-medium">Student roll number<input value={studentSearch} onChange={(event) => setStudentSearch(event.target.value)} onKeyDown={(event) => { if (event.key === 'Enter') { event.preventDefault(); void loadStudentDetail(); } }} placeholder="Search by roll number" className="mt-1 w-full rounded-lg border px-3 py-2" /></label><button onClick={() => void loadStudentDetail()} className="rounded-lg bg-indigo-600 px-4 py-2 text-white">Search</button></div>
        {studentDetail && <div className="mt-6">
          <div className="flex flex-wrap items-start justify-between gap-4"><div><h3 className="text-2xl font-semibold">{studentDetail.student.name}</h3><p className="text-slate-500">{studentDetail.student.rollNo}</p><p className="mt-2 text-sm">Defaults: {studentDetail.student.numberOfDefaults}</p></div><button onClick={() => void toggleStudentLock()} className="rounded-lg border px-3 py-2">{studentDetail.student.studentLocked ? 'Unlock student' : 'Lock student'}</button></div>
          <h4 className="mt-6 font-semibold">Default and DC decision history</h4>
          <div className="mt-2 space-y-3">{studentDetail.defaults.map((record) => (
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
          ))}</div>
        </div>}
      </section>}

      {view === 'activity' && <section className="rounded-2xl border border-slate-200 bg-white p-6 shadow-sm">
        <div className="sticky-filters bg-white py-2"><h2 className="text-xl font-semibold">Activity</h2><p className="text-sm text-slate-500">{activity.length} results</p></div>
        <div className="sticky-filters mt-4 grid gap-3 bg-white py-2 sm:grid-cols-2 lg:grid-cols-6">
          <label className="text-sm">Member<select aria-label="Member" value={activityFilters.memberId} onChange={(event) => setActivityFilters({ ...activityFilters, memberId: event.target.value })} className="mt-1 w-full rounded border px-2 py-2"><option value="">All members</option>{members.map((member) => <option key={member.memberId} value={member.memberId}>{member.name}</option>)}</select></label>
          <label className="text-sm">Team<select aria-label="Team" value={activityFilters.team} onChange={(event) => setActivityFilters({ ...activityFilters, team: event.target.value })} className="mt-1 w-full rounded border px-2 py-2"><option value="">All teams</option><option value="placement_cell">Placement cell</option><option value="summer_internship">Summer internship</option></select></label>
          <label className="text-sm">Event type<select aria-label="Event type" value={activityFilters.type} onChange={(event) => setActivityFilters({ ...activityFilters, type: event.target.value })} className="mt-1 w-full rounded border px-2 py-2"><option value="">All event types</option>{['logged', 'edited', 'locked', 'unlocked', 'removed', 'restored', 'dc_note', 'dc_scheduled', 'dc_unscheduled', 'dc_attendance', 'dc_decision', 'dc_rescheduled', 'dc_cancelled', 'student_locked', 'student_unlocked'].map((type) => <option key={type} value={type}>{type.replaceAll('_', ' ')}</option>)}</select></label>
          <label className="text-sm">From date<input type="date" aria-label="From date" value={activityFilters.from} onChange={(event) => setActivityFilters({ ...activityFilters, from: event.target.value })} className="mt-1 w-full rounded border px-2 py-2" /></label>
          <label className="text-sm">To date<input type="date" aria-label="To date" value={activityFilters.to} onChange={(event) => setActivityFilters({ ...activityFilters, to: event.target.value })} className="mt-1 w-full rounded border px-2 py-2" /></label>
          <label className="text-sm">Roll number<input placeholder="Any roll number" value={activityFilters.rollNo} onChange={(event) => setActivityFilters({ ...activityFilters, rollNo: event.target.value })} className="mt-1 w-full rounded border px-2 py-2" /></label>
        </div>
        <Timeline events={activity} />
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
              <button onClick={downloadMemberTemplate} className="rounded-lg border px-3 py-2 text-sm">Download Excel template</button>
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
  return <ol className="mt-4 space-y-3 border-l-2 border-indigo-100 pl-4">{events.map((event) => <li key={event.eventId} className="relative rounded-lg border border-slate-200 bg-slate-50 p-3"><span className="absolute -left-[23px] top-4 h-2.5 w-2.5 rounded-full bg-indigo-500" /><div className="flex flex-wrap justify-between gap-2"><strong className="capitalize">{event.type.replaceAll('_', ' ')}</strong><time className="text-xs text-slate-500">{new Date(event.at).toLocaleString()}</time></div><p className="mt-1 text-sm">{event.memberName} · {teamLabel(event.team)} · {event.enteredByEmail}</p>{event.studentName && <p className="text-sm text-slate-600">{event.studentName} ({event.rollNo})</p>}{event.reason && <p className="mt-2 text-sm"><strong>Reason:</strong> {event.reason}</p>}{event.note && <p className="mt-2 text-sm"><strong>Note:</strong> {event.note}</p>}{event.dcReference && <p className="mt-1 text-xs text-slate-600">DC reference: {event.dcReference}</p>}{event.changes && Object.keys(event.changes).length > 0 && <div className="mt-2 space-y-1 text-xs text-slate-700">{Object.entries(event.changes).map(([field, change]) => <p key={field}><strong>{field}:</strong> {String(change.from)} → {String(change.to)}</p>)}</div>}</li>)}</ol>;
}
