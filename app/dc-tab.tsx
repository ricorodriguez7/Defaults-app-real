'use client';

import { FormEvent, useCallback, useEffect, useMemo, useRef, useState } from 'react';
import * as XLSX from 'xlsx';
import { downloadWorkbook } from '@/lib/excel-download';

type DCStatus = 'Planned' | 'In progress' | 'Completed' | 'Cancelled';
type Attendance = 'Pending' | 'Present' | 'Absent';
type DCDecision = 'Undecided' | 'Remove' | 'Keep' | 'Discuss with team' | 'Defer';

type DCRecord = {
  id: string;
  date: string;
  time: string;
  venue: string;
  status: DCStatus;
  notes: string;
  cancelReason?: string;
  counts: { studentsCalled: number; yetToAttend: number; attended: number; absent: number; defaultsToBeRemoved: number; toDiscuss: number; deferred: number };
};

type AutoAddResult = { added: number; defaultsAdded: number; failures: string[]; message: string };

type PendingDefault = {
  id: string;
  rollNo: string;
  name: string;
  batch?: 'junior' | 'senior';
  company: string;
  typeName: string;
  date: string;
  dcStatus: 'Not scheduled' | 'Scheduled' | 'To discuss' | 'Deferred' | 'Decided';
  discussion?: { reason?: string; discussionDate?: string; notes?: Array<{ byName: string; at: string; text: string }>; raisedByName?: string };
  deferral?: { until?: string; reason?: string; byName?: string; history?: Array<{ until: string; reason: string; at: string; byName?: string }> };
};

type PendingStudent = {
  rollNo: string;
  name: string;
  batch?: 'junior' | 'senior';
  defaultCount: number;
  daysWaiting: number;
  defaults: PendingDefault[];
};

type DCDefault = {
  id: string;
  company: string;
  typeName: string;
  date: string;
  status: 'logged' | 'locked' | 'removed';
  decision: DCDecision;
  reason?: string;
  until?: string;
  discussionDate?: string;
};

type DCEntry = {
  rollNo: string;
  studentName: string;
  batch?: 'junior' | 'senior';
  attendance: Attendance;
  absentReason: string;
  remarks: string;
  defaults: DCDefault[];
};

type DCDetail = { dc: DCRecord; entries: DCEntry[] };
type RosterStudent = { rollNo: string; name: string; batch?: 'junior' | 'senior' };

const fieldClass = 'mt-1 min-h-11 w-full rounded-lg border border-slate-300 bg-white px-3 py-2 text-sm';
const buttonClass = 'inline-flex min-h-11 items-center justify-center rounded-lg px-4 py-2 text-sm font-semibold transition focus-visible:outline focus-visible:outline-2 focus-visible:outline-offset-2';
const normalizeLookup = (value: string) => value.toLocaleLowerCase().replace(/\s+/g, '');
const normalizeName = (value: string) => value.trim().replace(/\s+/g, ' ').toLocaleLowerCase();
const todayInIndia = () => {
  const parts = new Intl.DateTimeFormat('en-GB', {
    timeZone: 'Asia/Kolkata', year: 'numeric', month: '2-digit', day: '2-digit',
  }).formatToParts(new Date());
  const values = Object.fromEntries(parts.map((part) => [part.type, part.value]));
  return `${values.year}-${values.month}-${values.day}`;
};

async function responseJson(response: Response) {
  const data = await response.json();
  if (!response.ok) throw new Error(data.message || `Request failed (${response.status}).`);
  return data;
}

function statusStyle(status: string) {
  if (status === 'Completed' || status === 'Present' || status === 'Keep') return 'status-badge status-success';
  if (status === 'Cancelled' || status === 'Absent' || status === 'Remove') return 'status-badge status-danger';
  if (status === 'Pending' || status === 'Undecided' || status === 'Planned') return 'status-badge status-pending';
  return 'status-badge';
}

export default function DCTab({
  onRefresh,
  onOpenStudent,
}: {
  onRefresh: () => Promise<void>;
  onOpenStudent: (rollNo: string) => Promise<void>;
}) {
  const [dcs, setDcs] = useState<DCRecord[]>([]);
  const [roster, setRoster] = useState<RosterStudent[]>([]);
  const [pendingStudents, setPendingStudents] = useState<PendingStudent[]>([]);
  const [toDiscussDefaults, setToDiscussDefaults] = useState<PendingDefault[]>([]);
  const [deferredDefaults, setDeferredDefaults] = useState<PendingDefault[]>([]);
  const [dueDefaults, setDueDefaults] = useState<PendingDefault[]>([]);
  const [pendingCounts, setPendingCounts] = useState({ ready: 0, toDiscuss: 0, deferred: 0, due: 0 });
  const [pendingFilter, setPendingFilter] = useState<'ready' | 'discussion' | 'deferred' | 'due'>('ready');
  const [pendingBatchFilter, setPendingBatchFilter] = useState('');
  const [detail, setDetail] = useState<DCDetail | null>(null);
  const [studentSearch, setStudentSearch] = useState('');
  const [lookupResult, setLookupResult] = useState<{ query: string; kind: 'outside' | 'missing'; student?: RosterStudent } | null>(null);
  const [attendanceFilter, setAttendanceFilter] = useState<'all' | Attendance>('all');
  const [attendanceBatchFilter, setAttendanceBatchFilter] = useState('');
  const [expandedRollNo, setExpandedRollNo] = useState('');
  const searchInputRef = useRef<HTMLInputElement>(null);
  const [selectedRollNos, setSelectedRollNos] = useState<string[]>([]);
  const [targetDc, setTargetDc] = useState('');
  const [loading, setLoading] = useState(true);
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState('');
  const [message, setMessage] = useState('');
  const [showPlanForm, setShowPlanForm] = useState(false);
  const [plan, setPlan] = useState({ date: todayInIndia(), time: '', venue: '', notes: '' });
  const [reschedule, setReschedule] = useState({ date: '', time: '', venue: '', notes: '' });
  const [rescheduleReason, setRescheduleReason] = useState('');
  const [cancelTarget, setCancelTarget] = useState<DCRecord | null>(null);
  const [cancelReason, setCancelReason] = useState('');
  const [removalTarget, setRemovalTarget] = useState<DCEntry | null>(null);
  const [removalReason, setRemovalReason] = useState('');
  const [confirmBulkAbsence, setConfirmBulkAbsence] = useState(false);
  const [workflowTarget, setWorkflowTarget] = useState<{ record: PendingDefault; action: 'flag_discussion' | 'resolve_discussion' | 'defer' | 'extend_deferral'; outcome?: 'Take to DC' | 'Remove' | 'Defer' } | null>(null);
  const [workflowReason, setWorkflowReason] = useState('');
  const [workflowDate, setWorkflowDate] = useState('');
  const [discussionNotes, setDiscussionNotes] = useState<Record<string, string>>({});
  const [decisionDraft, setDecisionDraft] = useState<{ defaultId: string; rollNo: string; decision: DCDecision; reason: string; until: string; discussionDate: string } | null>(null);

  const loadOverview = useCallback(async () => {
    setLoading(true);
    try {
      const data = await responseJson(await fetch('/api/dcs', { cache: 'no-store' }));
      setDcs(data.dcs);
      setPendingStudents(data.pendingStudents);
      setToDiscussDefaults(data.toDiscussDefaults);
      setDeferredDefaults(data.deferredDefaults);
      setDueDefaults(data.dueDefaults);
      setPendingCounts({
        ready: data.readyDefaultCount,
        toDiscuss: data.toDiscussCount,
        deferred: data.deferredCount,
        due: data.dueCount,
      });
      setTargetDc((current) => data.dcs.some((dc: DCRecord) => dc.id === current && dc.status === 'Planned')
        ? current
        : (data.dcs.find((dc: DCRecord & { isCurrent?: boolean }) => dc.status === 'Planned' && dc.isCurrent) ??
          data.dcs.find((dc: DCRecord) => dc.status === 'Planned'))?.id ?? '');
      if (data.autoAdd?.message) setMessage(data.autoAdd.message);
      setError('');
    } catch (caught) {
      setError(caught instanceof Error ? caught.message : 'Unable to load DC data.');
    } finally {
      setLoading(false);
    }
  }, []);

  function openDc(dcId: string) {
    setStudentSearch('');
    setLookupResult(null);
    setAttendanceFilter('all');
    setExpandedRollNo('');
    setDetail(null);
    void loadDetail(dcId);
  }

  const loadDetail = useCallback(async (dcId: string) => {
    try {
      const data = await responseJson(await fetch(`/api/dcs/${encodeURIComponent(dcId)}`, { cache: 'no-store' }));
      setDetail(data);
      setReschedule({ date: data.dc.date, time: data.dc.time || '', venue: data.dc.venue || '', notes: data.dc.notes || '' });
      if (data.autoAdd?.message) setMessage(data.autoAdd.message);
      setError('');
    } catch (caught) {
      setError(caught instanceof Error ? caught.message : 'Unable to load the DC planner.');
    }
  }, []);

  useEffect(() => { void loadOverview(); }, [loadOverview]);
  useEffect(() => {
    let active = true;
    void fetch('/api/students', { cache: 'no-store' }).then(responseJson)
      .then((students: RosterStudent[]) => { if (active) setRoster(students); })
      .catch((caught) => { if (active) setError(caught instanceof Error ? caught.message : 'Unable to load the student roster.'); });
    return () => { active = false; };
  }, []);

  const pendingCount = pendingStudents.length;
  const pendingRows = pendingFilter === 'ready'
    ? pendingStudents
    : pendingFilter === 'discussion'
      ? toDiscussDefaults
      : pendingFilter === 'deferred'
        ? deferredDefaults
        : dueDefaults;
  const visiblePendingRows = pendingRows.filter((item) =>
    !pendingBatchFilter || (pendingBatchFilter === 'not_set' ? !item.batch : item.batch === pendingBatchFilter),
  );
  const attendanceComplete = Boolean(detail?.entries.length) && detail?.entries.every((entry) => entry.attendance !== 'Pending');
  const attendanceCounts = useMemo(() => ({
    total: detail?.entries.length ?? 0,
    present: detail?.entries.filter((entry) => entry.attendance === 'Present').length ?? 0,
    absent: detail?.entries.filter((entry) => entry.attendance === 'Absent').length ?? 0,
    pending: detail?.entries.filter((entry) => entry.attendance === 'Pending').length ?? 0,
  }), [detail]);
  const attendanceNearlyComplete = attendanceCounts.total > 0 &&
    attendanceCounts.pending > 0 &&
    (attendanceCounts.present + attendanceCounts.absent) / attendanceCounts.total >= 0.8;
  const filteredEntries = useMemo(() => {
    if (!detail) return [];
    const query = normalizeLookup(studentSearch);
    const nameQuery = normalizeName(studentSearch);
    return detail.entries.filter((entry) => {
      const attendanceMatches = (attendanceFilter === 'all' || entry.attendance === attendanceFilter) &&
        (!attendanceBatchFilter || (attendanceBatchFilter === 'not_set' ? !entry.batch : entry.batch === attendanceBatchFilter));
      const searchMatches = !query ||
        normalizeLookup(entry.rollNo).includes(query) ||
        normalizeName(entry.studentName).includes(nameQuery);
      return attendanceMatches && searchMatches;
    });
  }, [detail, studentSearch, attendanceFilter, attendanceBatchFilter]);
  const visiblePendingStudents = useMemo(() => pendingStudents.filter((student) =>
    !pendingBatchFilter || (pendingBatchFilter === 'not_set' ? !student.batch : student.batch === pendingBatchFilter),
  ), [pendingStudents, pendingBatchFilter]);
  const completionSummary = useMemo(() => {
    if (!detail) return null;
    const presentEntries = detail.entries.filter((entry) => entry.attendance === 'Present');
    return {
      present: presentEntries.length,
      absent: detail.entries.filter((entry) => entry.attendance === 'Absent').length,
      remove: presentEntries.reduce((count, entry) => count + entry.defaults.filter((record) => record.decision === 'Remove').length, 0),
      keep: presentEntries.reduce((count, entry) => count + entry.defaults.filter((record) => record.decision === 'Keep').length, 0),
      undecided: presentEntries.reduce((count, entry) => count + entry.defaults.filter((record) => record.decision === 'Undecided').length, 0),
      discuss: presentEntries.reduce((count, entry) => count + entry.defaults.filter((record) => record.decision === 'Discuss with team').length, 0),
      defer: presentEntries.reduce((count, entry) => count + entry.defaults.filter((record) => record.decision === 'Defer').length, 0),
    };
  }, [detail]);

  async function mutateDc(dcId: string, payload: Record<string, unknown>, successMessage?: string): Promise<boolean> {
    setBusy(true);
    setError('');
    setMessage('');
    try {
      const result = await responseJson(await fetch(`/api/dcs/${encodeURIComponent(dcId)}`, {
        method: 'PATCH',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(payload),
      }));
      await Promise.all([loadOverview(), loadDetail(dcId), onRefresh()]);
      if (successMessage) {
        const autoAddMessage = result.autoAdd?.message ? ` ${result.autoAdd.message}` : '';
        setMessage(`${successMessage}${autoAddMessage}`);
      }
      return true;
    } catch (caught) {
      setError(caught instanceof Error ? caught.message : 'Unable to update the DC.');
      return false;
    } finally {
      setBusy(false);
    }
  }

  async function createDc(event: FormEvent<HTMLFormElement>) {
    event.preventDefault();
    setBusy(true);
    setError('');
    try {
      const data = await responseJson(await fetch('/api/dcs', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(plan),
      }));
      setShowPlanForm(false);
      setPlan({ date: todayInIndia(), time: '', venue: '', notes: '' });
      await Promise.all([loadOverview(), onRefresh()]);
      openDc(data.dc.id);
      setMessage(`DC planned. ${data.autoAdd?.message ?? '0 students added automatically.'}`);
    } catch (caught) {
      setError(caught instanceof Error ? caught.message : 'Unable to plan the DC.');
    } finally {
      setBusy(false);
    }
  }

  async function syncNow() {
    setBusy(true);
    setError('');
    setMessage('');
    try {
      const data = await responseJson(await fetch('/api/dcs', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ action: 'sync' }),
      }));
      await Promise.all([loadOverview(), onRefresh()]);
      setMessage((data.autoAdd as AutoAddResult | undefined)?.message ?? '0 students added automatically.');
    } catch (caught) {
      setError(caught instanceof Error ? caught.message : 'Unable to sync students to the current DC.');
    } finally {
      setBusy(false);
    }
  }

  async function addSelectedStudents() {
    if (!targetDc || !selectedRollNos.length) return;
    await mutateDc(targetDc, { action: 'add_students', rollNos: selectedRollNos }, 'Students added to the planned DC.');
    setSelectedRollNos([]);
  }

  async function updateAttendance(entry: DCEntry, attendance: Attendance) {
    if (!detail) return;
    const updated = await mutateDc(detail.dc.id, {
      action: 'attendance',
      rollNo: entry.rollNo,
      attendance,
      absentReason: attendance === 'Absent' ? entry.absentReason : '',
    });
    if (updated) {
      setStudentSearch('');
      setLookupResult(null);
      setExpandedRollNo(entry.rollNo);
      requestAnimationFrame(() => searchInputRef.current?.focus());
    }
  }

  async function saveAbsentReason(entry: DCEntry, absentReason: string) {
    if (!detail || entry.attendance !== 'Absent' || entry.absentReason === absentReason) return;
    await mutateDc(detail.dc.id, { action: 'attendance', rollNo: entry.rollNo, attendance: 'Absent', absentReason });
  }

  async function saveRemarks(entry: DCEntry, remarks: string) {
    if (!detail || entry.remarks === remarks) return;
    await mutateDc(detail.dc.id, { action: 'remarks', rollNo: entry.rollNo, remarks });
  }

  function resolveStudentSearch() {
    if (!detail || !studentSearch.trim()) return;
    if (filteredEntries.length === 1) {
      setExpandedRollNo(filteredEntries[0].rollNo);
      setLookupResult(null);
      return;
    }
    if (filteredEntries.length > 1) return;

    const query = normalizeLookup(studentSearch);
    const nameQuery = normalizeName(studentSearch);
    const student = roster.find((candidate) =>
      normalizeLookup(candidate.rollNo) === query || normalizeName(candidate.name) === nameQuery,
    );
    if (student && !detail.entries.some((entry) => normalizeLookup(entry.rollNo) === normalizeLookup(student.rollNo))) {
      setLookupResult({ query: studentSearch, kind: 'outside', student });
      return;
    }
    if (student && detail.entries.some((entry) => normalizeLookup(entry.rollNo) === normalizeLookup(student.rollNo))) {
      const entry = detail.entries.find((candidate) => normalizeLookup(candidate.rollNo) === normalizeLookup(student.rollNo));
      if (entry) {
        setAttendanceFilter('all');
        setExpandedRollNo(entry.rollNo);
        setLookupResult(null);
      }
      return;
    }
    setLookupResult({ query: studentSearch, kind: 'missing' });
  }

  async function addSearchedStudent() {
    if (!detail || !lookupResult?.student || detail.dc.status !== 'Planned') return;
    const rollNo = lookupResult.student.rollNo;
    if (await mutateDc(detail.dc.id, { action: 'add_students', rollNos: [rollNo] }, `${lookupResult.student.name} added to this DC.`)) {
      setStudentSearch('');
      setLookupResult(null);
      setExpandedRollNo(rollNo);
    }
  }

  async function markPendingAbsent() {
    if (!detail || !attendanceCounts.pending) return;
    setConfirmBulkAbsence(true);
  }

  async function confirmMarkPendingAbsent() {
    if (!detail) return;
    const updated = await mutateDc(detail.dc.id, { action: 'mark_pending_absent' }, 'Pending students marked absent.');
    if (updated) setConfirmBulkAbsence(false);
  }

  function openWorkflow(
    record: PendingDefault,
    action: 'flag_discussion' | 'resolve_discussion' | 'defer' | 'extend_deferral',
    outcome?: 'Take to DC' | 'Remove' | 'Defer',
  ) {
    setWorkflowTarget({ record, action, outcome });
    setWorkflowReason('');
    setWorkflowDate('');
    setError('');
  }

  async function submitWorkflow(event: FormEvent<HTMLFormElement>) {
    event.preventDefault();
    if (!workflowTarget) return;
    const { record, action, outcome } = workflowTarget;
    const reason = workflowReason.trim();
    if (!reason) {
      setError('A reason or discussion note is required.');
      return;
    }
    if ((action === 'defer' || action === 'extend_deferral' || outcome === 'Defer') && !workflowDate) {
      setError('Choose a deferral date or select next DC.');
      return;
    }
    setBusy(true);
    setError('');
    try {
      const result = await responseJson(await fetch(`/api/defaults/${encodeURIComponent(record.id)}`, {
        method: 'PATCH',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          action,
          reason,
          ...(action === 'flag_discussion' ? { discussionDate: workflowDate } : {}),
          ...(action === 'resolve_discussion' ? { outcome, until: outcome === 'Defer' ? workflowDate : '' } : {}),
          ...(action === 'defer' || action === 'extend_deferral' ? { until: workflowDate } : {}),
        }),
      }));
      setWorkflowTarget(null);
      const workflowMessage = action === 'flag_discussion' ? 'Default flagged for team discussion.' :
        action === 'resolve_discussion' ? `Discussion resolved: ${outcome}.` :
          action === 'extend_deferral' ? 'Deferral extended.' : 'Default deferred.';
      setMessage(result.autoAdd?.message
        ? `${workflowMessage} ${result.autoAdd.message}`
        : workflowMessage);
      await Promise.all([loadOverview(), onRefresh()]);
    } catch (caught) {
      setError(caught instanceof Error ? caught.message : 'Unable to update discussion or deferral.');
    } finally {
      setBusy(false);
    }
  }

  async function addDiscussionNote(record: PendingDefault) {
    const note = (discussionNotes[record.id] ?? '').trim();
    if (!note) return;
    setBusy(true);
    setError('');
    try {
      await responseJson(await fetch(`/api/defaults/${encodeURIComponent(record.id)}`, {
        method: 'PATCH',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ action: 'add_discussion_note', note }),
      }));
      setDiscussionNotes((current) => ({ ...current, [record.id]: '' }));
      setMessage('Discussion note added.');
      await Promise.all([loadOverview(), onRefresh()]);
    } catch (caught) {
      setError(caught instanceof Error ? caught.message : 'Unable to add discussion note.');
    } finally {
      setBusy(false);
    }
  }

  async function saveDecisionDraft(event: FormEvent<HTMLFormElement>) {
    event.preventDefault();
    if (!detail || !decisionDraft) return;
    const { defaultId, rollNo, decision, reason, until, discussionDate } = decisionDraft;
    if ((decision === 'Discuss with team' || decision === 'Defer') && !reason.trim()) {
      setError(decision === 'Defer' ? 'A deferral reason is required.' : 'A discussion reason is required.');
      return;
    }
    if (decision === 'Defer' && !until) {
      setError('Choose a deferral date or select next DC.');
      return;
    }
    const saved = await mutateDc(detail.dc.id, {
      action: 'decision',
      defaultId,
      rollNo,
      decision,
      reason: reason.trim(),
      until,
      discussionDate,
    });
    if (saved) setDecisionDraft(null);
  }

  async function chooseDecision(entry: DCEntry, record: DCDefault, decision: DCDecision) {
    if (!detail || record.decision === decision) return;
    if (decision === 'Discuss with team' || decision === 'Defer') {
      setError('');
      setDecisionDraft({
        defaultId: record.id,
        rollNo: entry.rollNo,
        decision,
        reason: '',
        until: '',
        discussionDate: '',
      });
      return;
    }
    await mutateDc(detail.dc.id, {
      action: 'decision',
      rollNo: entry.rollNo,
      defaultId: record.id,
      decision,
      reason: '',
      until: '',
    });
  }

  async function exportDcs(all: boolean) {
    setBusy(true);
    try {
      const records = all ? dcs : detail ? [detail.dc] : [];
      const details = all
        ? await Promise.all(records.map(async (dc) => responseJson(await fetch(`/api/dcs/${encodeURIComponent(dc.id)}`, { cache: 'no-store' }))))
        : detail ? [detail] : [];
      const rows = details.flatMap((item: DCDetail) => item.entries.flatMap((entry) => {
        const defaults = entry.defaults.length ? entry.defaults : [null];
        return defaults.map((record) => ({
          'DC date': item.dc.date,
          'DC time': item.dc.time || '',
          'DC venue': item.dc.venue || '',
          'DC status': item.dc.status,
          'Roll number': entry.rollNo,
          Name: entry.studentName,
          Batch: entry.batch ? `${entry.batch === 'junior' ? 'Junior' : 'Senior'} batch` : 'Not set',
          Attendance: entry.attendance,
          'Absent reason': entry.absentReason,
          'Number of defaults': entry.defaults.length,
          Default: record ? `${record.company} — ${record.typeName}` : '',
          Decision: record?.decision ?? '',
          'Decision reason': record?.reason ?? '',
          'Deferred until': record?.until === 'next_dc' ? 'Next DC' : record?.until ?? '',
          'Discussion date': record?.discussionDate ?? '',
          Remarks: entry.remarks,
        }));
      }));
      const summaries = details.map((item: DCDetail) => {
        const batchEntries = (batch: string) => item.entries.filter((entry) =>
          batch === 'not_set' ? !entry.batch : entry.batch === batch,
        );
        const batchCounts = (batch: string, label: string) => {
          const entries = batchEntries(batch);
          const defaults = entries.flatMap((entry) => entry.defaults);
          return {
            [`${label} called`]: entries.length,
            [`${label} present`]: entries.filter((entry) => entry.attendance === 'Present').length,
            [`${label} absent`]: entries.filter((entry) => entry.attendance === 'Absent').length,
            [`${label} pending`]: entries.filter((entry) => entry.attendance === 'Pending').length,
            [`${label} remove decisions`]: defaults.filter((record) => record.decision === 'Remove').length,
            [`${label} keep decisions`]: defaults.filter((record) => record.decision === 'Keep').length,
          };
        };
        return {
          'DC date': item.dc.date,
          'DC status': item.dc.status,
          Called: item.entries.length,
          Present: item.entries.filter((entry) => entry.attendance === 'Present').length,
          Absent: item.entries.filter((entry) => entry.attendance === 'Absent').length,
          Pending: item.entries.filter((entry) => entry.attendance === 'Pending').length,
          'Absent roll numbers': item.entries.filter((entry) => entry.attendance === 'Absent').map((entry) => entry.rollNo).join(', '),
          ...batchCounts('junior', 'Junior'),
          ...batchCounts('senior', 'Senior'),
          ...batchCounts('not_set', 'Not set'),
        };
      });
      const workbook = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(rows), 'DC meetings');
      XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(summaries), 'Attendance summary');
      downloadWorkbook(workbook, all ? 'all-dc-meetings.xlsx' : 'dc-meeting.xlsx');
    } catch (caught) {
      setError(caught instanceof Error ? caught.message : 'Unable to export DC data.');
    } finally {
      setBusy(false);
    }
  }

  async function completeDc() {
    if (!detail || !completionSummary) return;
    const absentRollNos = detail.entries.filter((entry) => entry.attendance === 'Absent').map((entry) => entry.rollNo);
    const confirmation = `Complete this DC?\n\n${completionSummary.present} present, ${completionSummary.absent} absent, ${completionSummary.remove} defaults to remove, ${completionSummary.keep} to keep, ${completionSummary.undecided} undecided, ${completionSummary.discuss} to discuss, ${completionSummary.defer} deferred.\n\nAbsent students: ${absentRollNos.join(', ') || 'None'}`;
    if (!window.confirm(confirmation)) return;
    await mutateDc(detail.dc.id, { action: 'complete' }, 'DC completed. Present defaults were decided; absent and undecided defaults returned to the pending list.');
  }

  async function cancelDc(dc: DCRecord) {
    setCancelTarget(dc);
    setCancelReason('');
    setError('');
    setMessage('');
  }

  async function confirmCancelDc(event: FormEvent<HTMLFormElement>) {
    event.preventDefault();
    if (!cancelTarget) return;
    const reason = cancelReason.trim();
    if (!reason) {
      setError('Enter a reason before cancelling this DC.');
      return;
    }
    const cancelled = await mutateDc(cancelTarget.id, { action: 'cancel', reason }, 'DC cancelled. Assigned students are back in the pending list.');
    if (cancelled) {
      setCancelTarget(null);
      setCancelReason('');
      setDetail((current) => current?.dc.id === cancelTarget.id ? null : current);
    }
  }

  async function removeStudent(entry: DCEntry) {
    if (!detail) return;
    setRemovalTarget(entry);
    setRemovalReason('');
    setError('');
  }

  async function confirmRemoveStudent(event: FormEvent<HTMLFormElement>) {
    event.preventDefault();
    if (!detail || !removalTarget) return;
    const removed = await mutateDc(detail.dc.id, {
      action: 'remove_student',
      rollNo: removalTarget.rollNo,
      reason: removalReason.trim(),
    }, 'Student removed from the planned DC.');
    if (removed) {
      setRemovalTarget(null);
      setRemovalReason('');
    }
  }

  async function rescheduleDc(event: FormEvent<HTMLFormElement>) {
    event.preventDefault();
    if (!detail) return;
    const reason = rescheduleReason.trim();
    if (!reason) return;
    if (await mutateDc(detail.dc.id, { action: 'reschedule', ...reschedule, reason }, 'DC rescheduled.')) setRescheduleReason('');
  }

  return (
    <div className="space-y-6">
      {error && <p role="alert" className="rounded-xl border border-rose-200 bg-rose-50 p-4 text-sm text-rose-800">{error}</p>}
      {message && <p role="status" className="rounded-xl border border-emerald-200 bg-emerald-50 p-4 text-sm text-emerald-800">{message}</p>}

      {workflowTarget && <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-950/50 p-4" role="presentation">
        <section role="dialog" aria-modal="true" aria-labelledby="default-workflow-title" className="w-full max-w-lg rounded-2xl bg-white p-5 shadow-xl sm:p-6">
          <h2 id="default-workflow-title" className="text-xl font-semibold">
            {workflowTarget.action === 'flag_discussion' ? 'Needs team discussion' :
              workflowTarget.action === 'resolve_discussion' ? `Resolve discussion: ${workflowTarget.outcome}` :
                workflowTarget.action === 'extend_deferral' ? 'Extend deferral' : 'Defer default'}
          </h2>
          <p className="mt-2 text-sm text-slate-600">{workflowTarget.record.name} · {workflowTarget.record.rollNo} · {workflowTarget.record.company}</p>
          <form onSubmit={(event) => void submitWorkflow(event)} className="mt-4 space-y-3">
            <label htmlFor="workflow-reason" className="block text-sm font-medium">
              {workflowTarget.action === 'resolve_discussion' ? 'Required discussion note / outcome reason *' :
                workflowTarget.action === 'flag_discussion' ? 'Reason for discussion *' : 'Deferral reason *'}
              <textarea id="workflow-reason" autoFocus required value={workflowReason} onChange={(event) => setWorkflowReason(event.target.value)} maxLength={2000} rows={3} className={fieldClass} />
            </label>
            {workflowTarget.action === 'flag_discussion' && <label className="block text-sm font-medium">Discussion date (optional)<input type="date" value={workflowDate} onChange={(event) => setWorkflowDate(event.target.value)} className={fieldClass} /></label>}
            {(workflowTarget.action === 'defer' || workflowTarget.action === 'extend_deferral' || workflowTarget.outcome === 'Defer') && <div className="block text-sm font-medium">Deferred until *<select aria-label="Deferred until" value={workflowDate === 'next_dc' ? 'next_dc' : workflowDate ? 'specific' : ''} onChange={(event) => setWorkflowDate(event.target.value === 'next_dc' ? 'next_dc' : '')} className={fieldClass}><option value="">Choose a date or next DC</option><option value="specific">Specific date</option><option value="next_dc">Next DC</option></select>{workflowDate !== 'next_dc' && <input aria-label="Specific deferral date" required type="date" value={workflowDate} onChange={(event) => setWorkflowDate(event.target.value)} className={fieldClass} />}</div>}
            {error && <p role="alert" className="text-sm text-rose-700">{error}</p>}
            <div className="flex flex-wrap justify-end gap-2">
              <button type="button" className={`${buttonClass} btn-secondary`} onClick={() => { setWorkflowTarget(null); setError(''); }} disabled={busy}>Cancel</button>
              <button type="submit" className={`${buttonClass} btn-primary`} disabled={busy}>{busy ? 'Saving…' : 'Save'}</button>
            </div>
          </form>
        </section>
      </div>}

      {cancelTarget && <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-950/50 p-4" role="presentation">
        <section role="dialog" aria-modal="true" aria-labelledby="cancel-dc-title" className="w-full max-w-lg rounded-2xl bg-white p-5 shadow-xl sm:p-6">
          <h2 id="cancel-dc-title" className="text-xl font-semibold">Confirm DC cancellation</h2>
          <p className="mt-2 text-sm text-slate-600">Cancel the DC scheduled for {cancelTarget.date}? All assigned students and their defaults will return to the pending list. This cannot be undone.</p>
          <form onSubmit={(event) => void confirmCancelDc(event)} className="mt-4">
            <label htmlFor="cancel-dc-reason" className="text-sm font-medium">Cancellation reason *</label>
            <textarea id="cancel-dc-reason" autoFocus required value={cancelReason} onChange={(event) => setCancelReason(event.target.value)} maxLength={1000} rows={3} className={fieldClass} />
            {error && <p role="alert" className="mt-2 text-sm text-rose-700">{error}</p>}
            <div className="mt-4 flex flex-wrap justify-end gap-2">
              <button type="button" className={`${buttonClass} btn-secondary`} onClick={() => { setCancelTarget(null); setCancelReason(''); setError(''); }} disabled={busy}>Keep DC</button>
              <button type="submit" className={`${buttonClass} btn-danger`} disabled={busy || !cancelReason.trim()}>{busy ? 'Cancelling…' : 'Confirm cancellation'}</button>
            </div>
          </form>
        </section>
      </div>}

      {removalTarget && <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-950/50 p-4" role="presentation">
        <form onSubmit={confirmRemoveStudent} role="dialog" aria-modal="true" aria-labelledby="remove-student-title" className="w-full max-w-lg rounded-2xl bg-white p-5 shadow-xl sm:p-6">
          <h2 id="remove-student-title" className="text-xl font-semibold">Remove student from DC</h2>
          <p className="mt-2 text-sm text-slate-600">Remove {removalTarget.studentName} ({removalTarget.rollNo}) from the planned DC? The assigned default(s) will return to Ready for DC.</p>
          <label className="mt-4 block text-sm font-medium">Reason *
            <textarea required autoFocus value={removalReason} onChange={(event) => setRemovalReason(event.target.value)} maxLength={1000} rows={3} className={fieldClass} />
          </label>
          <div className="mt-5 flex justify-end gap-2"><button type="button" className={`${buttonClass} btn-secondary`} onClick={() => setRemovalTarget(null)} disabled={busy}>Keep student</button><button type="submit" className={`${buttonClass} btn-danger`} disabled={busy || !removalReason.trim()}>Remove student</button></div>
        </form>
      </div>}

      {confirmBulkAbsence && detail && <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-950/50 p-4" role="presentation">
        <section role="dialog" aria-modal="true" aria-labelledby="bulk-absence-title" className="w-full max-w-lg rounded-2xl bg-white p-5 shadow-xl sm:p-6">
          <h2 id="bulk-absence-title" className="text-xl font-semibold">Mark remaining students absent?</h2>
          <p className="mt-2 text-sm text-slate-600">This will mark {attendanceCounts.pending} pending {attendanceCounts.pending === 1 ? 'student' : 'students'} absent. Students already marked Present or Absent will not change.</p>
          <div className="mt-5 flex justify-end gap-2"><button type="button" className={`${buttonClass} btn-secondary`} onClick={() => setConfirmBulkAbsence(false)} disabled={busy}>Cancel</button><button type="button" className={`${buttonClass} btn-danger`} onClick={() => void confirmMarkPendingAbsent()} disabled={busy}>{busy ? 'Saving…' : `Mark ${attendanceCounts.pending} absent`}</button></div>
        </section>
      </div>}

      <section className="rounded-2xl border border-slate-200 bg-white p-4 shadow-sm sm:p-6">
        <div className="flex flex-wrap items-start justify-between gap-3">
          <div>
            <h2 className="text-xl font-semibold text-slate-900">Pending defaults</h2>
            <p className="mt-1 text-sm text-slate-600"><strong>{pendingCount} students</strong> have defaults ready for a DC</p>
          </div>
          <div className="flex flex-wrap gap-2">
            <button type="button" className={`${buttonClass} btn-secondary`} onClick={() => void syncNow()} disabled={busy}>{busy ? 'Syncing…' : 'Sync now'}</button>
            <button type="button" className={`${buttonClass} btn-secondary`} onClick={() => void exportDcs(true)} disabled={busy || !dcs.length}>Export all DCs</button>
            <button type="button" className={`${buttonClass} btn-primary`} onClick={() => { setShowPlanForm((shown) => !shown); setDetail(null); }}>Plan a DC</button>
          </div>
        </div>

        {pendingCounts.due > 0 && <p role="status" className="mt-4 rounded-xl border border-amber-200 bg-amber-50 p-3 text-sm font-semibold text-amber-900">{pendingCounts.due} deferred {pendingCounts.due === 1 ? 'default is' : 'defaults are'} now due and ready for a DC.</p>}

        <div className="mt-4 flex flex-wrap gap-2" aria-label="Filter pending defaults">
          {([
            ['ready', `Ready for DC (${pendingCounts.ready})`],
            ['discussion', `To discuss (${pendingCounts.toDiscuss})`],
            ['deferred', `Deferred (${pendingCounts.deferred})`],
            ['due', `Due now (${pendingCounts.due})`],
          ] as const).map(([key, label]) => <button key={key} type="button" aria-pressed={pendingFilter === key} className={`${buttonClass} ${pendingFilter === key ? 'btn-primary' : 'btn-secondary'}`} onClick={() => setPendingFilter(key)}>{label}</button>)}
        </div>
        <label className="mt-3 block max-w-xs text-sm">Filter by batch<select value={pendingBatchFilter} onChange={(event) => setPendingBatchFilter(event.target.value)} className={fieldClass}><option value="">All batches</option><option value="junior">Junior batch</option><option value="senior">Senior batch</option><option value="not_set">Not set</option></select></label>

        {showPlanForm && <form onSubmit={createDc} className="mt-5 grid gap-3 rounded-xl border border-indigo-100 bg-indigo-50/60 p-4 sm:grid-cols-2">
          <label className="text-sm font-medium">Date *<input required type="date" value={plan.date} onChange={(event) => setPlan({ ...plan, date: event.target.value })} className={fieldClass} /></label>
          <label className="text-sm font-medium">Time (optional)<input type="time" value={plan.time} onChange={(event) => setPlan({ ...plan, time: event.target.value })} className={fieldClass} /></label>
          <label className="text-sm font-medium">Venue (optional)<input value={plan.venue} onChange={(event) => setPlan({ ...plan, venue: event.target.value })} maxLength={200} className={fieldClass} /></label>
          <label className="text-sm font-medium sm:col-span-2">Notes (optional)<textarea value={plan.notes} onChange={(event) => setPlan({ ...plan, notes: event.target.value })} maxLength={2000} rows={2} className={fieldClass} /></label>
          <div className="flex gap-2 sm:col-span-2"><button disabled={busy} className={`${buttonClass} btn-primary`}>Save DC</button><button type="button" className={`${buttonClass} btn-secondary`} onClick={() => setShowPlanForm(false)}>Cancel</button></div>
        </form>}

        {pendingFilter === 'ready' && pendingCount > 0 && visiblePendingStudents.length > 0 ? <div className="mt-5 space-y-3">
          <div className="grid gap-3 sm:grid-cols-[minmax(0,1fr)_auto]">
            <label className="text-sm font-medium">Planned DC<select value={targetDc} onChange={(event) => setTargetDc(event.target.value)} className={fieldClass}>
              <option value="">Select a planned DC</option>{dcs.filter((dc) => dc.status === 'Planned').map((dc) => <option key={dc.id} value={dc.id}>{dc.date}{dc.time ? ` · ${dc.time}` : ''}{dc.venue ? ` · ${dc.venue}` : ''}</option>)}
            </select></label>
            <button type="button" disabled={!targetDc || !selectedRollNos.length || busy} className={`${buttonClass} btn-primary self-end`} onClick={() => void addSelectedStudents()}>Add selected to DC ({selectedRollNos.length})</button>
          </div>
          <div className="grid gap-3 sm:grid-cols-2 xl:grid-cols-3">
            {visiblePendingStudents.map((student) => <article key={student.rollNo} className="rounded-xl border border-slate-200 p-3 hover:border-indigo-300">
              <label className="flex min-h-11 items-start gap-3">
                <input aria-label={`Select ${student.name}`} type="checkbox" checked={selectedRollNos.includes(student.rollNo)} onChange={(event) => setSelectedRollNos((current) => event.target.checked ? [...current, student.rollNo] : current.filter((rollNo) => rollNo !== student.rollNo))} className="mt-1 h-5 w-5" />
                <span className="min-w-0 flex-1"><span className="flex items-center justify-between gap-2"><strong className="truncate">{student.name}</strong><span className={statusStyle('Pending')}>Ready</span></span><span className="text-xs text-slate-500">{student.rollNo} · {student.batch ? `${student.batch === 'junior' ? 'Junior' : 'Senior'} batch` : 'Batch not set'} · {student.defaultCount} ready defaults · {student.daysWaiting} days waiting</span></span>
              </label>
              <div className="mt-3 space-y-2 border-t pt-2">{student.defaults.map((record) => <div key={record.id} className="rounded-lg bg-slate-50 p-2 text-sm">
                <p className="font-medium">{record.company} · {record.typeName}</p>
                <div className="mt-2 flex flex-wrap gap-2"><button type="button" className={`${buttonClass} btn-secondary`} onClick={() => openWorkflow(record, 'flag_discussion')}>Needs team discussion</button><button type="button" className={`${buttonClass} btn-secondary`} onClick={() => openWorkflow(record, 'defer')}>Defer</button></div>
              </div>)}</div>
            </article>)}
          </div>
        </div> : pendingFilter === 'ready' && pendingCount === 0
          ? <p className="mt-5 rounded-xl bg-slate-50 p-5 text-sm text-slate-600">No defaults are ready for a DC.</p>
          : pendingFilter === 'ready' && pendingCount > 0
            ? <p className="mt-5 rounded-xl bg-slate-50 p-5 text-sm text-slate-600">No students match this batch filter.</p>
            : visiblePendingRows.length === 0
            ? <p className="mt-5 rounded-xl bg-slate-50 p-5 text-sm text-slate-600">No defaults in this category.</p>
            : <div className="mt-5 space-y-3">{(visiblePendingRows as PendingDefault[]).map((record) => <article key={record.id} className="rounded-xl border border-slate-200 p-4">
              <div className="flex flex-wrap items-start justify-between gap-3"><div><h3 className="font-semibold">{record.name} · {record.rollNo}</h3><p className="text-sm text-slate-600">{record.batch ? `${record.batch === 'junior' ? 'Junior' : 'Senior'} batch · ` : 'Batch not set · '}{record.company} · {record.typeName}</p></div><span className={statusStyle(pendingFilter === 'discussion' ? 'To discuss' : pendingFilter === 'deferred' ? 'Deferred' : 'Pending')}>{pendingFilter === 'discussion' ? 'To discuss' : pendingFilter === 'deferred' ? 'Deferred' : 'Due now'}</span></div>
              {pendingFilter === 'discussion' && <div className="mt-3 space-y-2">
                <p className="text-sm"><strong>Reason:</strong> {record.discussion?.reason || '—'}{record.discussion?.discussionDate ? ` · Discussion date ${record.discussion.discussionDate}` : ''}</p>
                {(record.discussion?.notes ?? []).map((note, index) => <p key={`${record.id}-note-${index}`} className="rounded bg-slate-50 p-2 text-sm">{note.text} <span className="text-xs text-slate-500">— {note.byName}</span></p>)}
                <div className="flex flex-wrap gap-2"><input aria-label="Discussion note" value={discussionNotes[record.id] ?? ''} onChange={(event) => setDiscussionNotes((current) => ({ ...current, [record.id]: event.target.value }))} placeholder="Add a dated discussion note" className={fieldClass} /><button type="button" className={`${buttonClass} btn-secondary`} disabled={busy || !(discussionNotes[record.id] ?? '').trim()} onClick={() => void addDiscussionNote(record)}>Add note</button></div>
                <div className="flex flex-wrap gap-2"><button type="button" className={`${buttonClass} btn-primary`} onClick={() => openWorkflow(record, 'resolve_discussion', 'Take to DC')}>Take to DC</button><button type="button" className={`${buttonClass} btn-danger`} onClick={() => openWorkflow(record, 'resolve_discussion', 'Remove')}>Remove</button><button type="button" className={`${buttonClass} btn-secondary`} onClick={() => openWorkflow(record, 'resolve_discussion', 'Defer')}>Defer</button></div>
              </div>}
              {pendingFilter === 'deferred' && <div className="mt-3 space-y-2 text-sm"><p><strong>Until:</strong> {record.deferral?.until === 'next_dc' ? 'Next DC' : record.deferral?.until || '—'}</p><p><strong>Reason:</strong> {record.deferral?.reason || '—'}</p>{record.deferral?.history?.length ? <div><strong>Deferral history</strong><ul className="mt-1 list-disc space-y-1 pl-5">{record.deferral.history.map((item, index) => <li key={`${record.id}-deferral-${index}`}>{item.until === 'next_dc' ? 'Next DC' : item.until} — {item.reason} <span className="text-xs text-slate-500">({item.byName || 'Unknown member'} · {new Date(item.at).toLocaleDateString()})</span></li>)}</ul></div> : null}<button type="button" className={`${buttonClass} btn-secondary`} onClick={() => openWorkflow(record, 'extend_deferral')}>Extend deferral</button></div>}
              {pendingFilter === 'due' && <div className="mt-3 space-y-2 text-sm"><p><strong>Deferred until:</strong> {record.deferral?.until === 'next_dc' ? 'Next DC' : record.deferral?.until || '—'}</p><p><strong>Reason:</strong> {record.deferral?.reason || '—'}</p><p>This default is back in Ready for DC.</p></div>}
            </article>)}
          </div>}
      </section>

      <section className="rounded-2xl border border-slate-200 bg-white p-4 shadow-sm sm:p-6">
        <div className="mb-4 flex flex-wrap items-center justify-between gap-3">
          <div><h2 className="text-xl font-semibold text-slate-900">Disciplinary Committee meetings</h2><p className="mt-1 text-sm text-slate-500">{dcs.length} DCs · {pendingCounts.toDiscuss} defaults to discuss · {pendingCounts.deferred} deferred · newest first</p></div>
          {loading && <span role="status" className="text-sm text-slate-500">Loading DCs…</span>}
        </div>
        {!loading && !dcs.length && <p className="rounded-xl bg-slate-50 p-5 text-sm text-slate-600">No DCs planned yet. Plan a DC to start scheduling students.</p>}
        <div className="grid gap-3 md:grid-cols-2 xl:grid-cols-3">
          {dcs.map((dc) => <article key={dc.id} className={`rounded-xl border p-4 ${detail?.dc.id === dc.id ? 'border-indigo-400 bg-indigo-50/40' : 'border-slate-200'}`}>
            <div className="flex items-start justify-between gap-3">
              <button type="button" className="min-h-11 text-left" onClick={() => { setShowPlanForm(false); openDc(dc.id); }}>
                <strong className="block text-lg">{new Date(`${dc.date}T00:00:00`).toLocaleDateString()}</strong>
                <span className="text-sm text-slate-600">{dc.time || 'Time not set'}{dc.venue ? ` · ${dc.venue}` : ''}</span>
              </button>
              <span className={statusStyle(dc.status)}>{dc.status}</span>
            </div>
            <div className="mt-3 grid grid-cols-2 gap-2 text-sm">
              <span>Students called <strong>{dc.counts.studentsCalled}</strong></span>
              <span>Yet to attend <strong>{dc.counts.yetToAttend}</strong></span>
              <span>Attended <strong>{dc.counts.attended}</strong></span>
              <span>Absent <strong>{dc.counts.absent}</strong></span>
              <span>Defaults to remove <strong>{dc.counts.defaultsToBeRemoved}</strong></span>
              <span>To discuss <strong>{dc.counts.toDiscuss}</strong></span>
              <span>Deferred <strong>{dc.counts.deferred}</strong></span>
            </div>
            {(dc.status === 'Planned' || dc.status === 'In progress') && <div className="mt-3 flex flex-wrap gap-2">
              <button type="button" className={`${buttonClass} btn-secondary`} onClick={() => { setShowPlanForm(false); openDc(dc.id); }}>{dc.status === 'Planned' ? 'Open planner' : 'Open live DC'}</button>
              <button type="button" className={`${buttonClass} btn-danger`} onClick={() => void cancelDc(dc)}>Cancel DC</button>
            </div>}
            {dc.status !== 'Planned' && dc.status !== 'In progress' && <button type="button" className={`${buttonClass} btn-secondary mt-3`} onClick={() => { setShowPlanForm(false); openDc(dc.id); }}>View details</button>}
          </article>)}
        </div>
      </section>

      {detail && <section className="rounded-2xl border border-slate-200 bg-white p-4 shadow-sm sm:p-6">
        <div className="flex flex-wrap items-start justify-between gap-3">
          <div><h2 className="text-2xl font-semibold">DC planner</h2><p className="mt-1 text-slate-600">{detail.dc.date}{detail.dc.time ? ` · ${detail.dc.time}` : ''}{detail.dc.venue ? ` · ${detail.dc.venue}` : ''}</p>
            <div className="mt-2 flex flex-wrap items-center gap-2"><span className={statusStyle(detail.dc.status)}>{detail.dc.status}</span><span className="text-sm text-slate-600">{attendanceCounts.total} students called</span></div>
          </div>
          <div className="flex flex-wrap gap-2"><button type="button" className={`${buttonClass} btn-secondary`} onClick={() => void exportDcs(false)} disabled={busy}>Export this DC</button><button type="button" className={`${buttonClass} btn-secondary`} onClick={() => setDetail(null)}>Close</button></div>
        </div>
        <div className="sticky top-0 z-10 mt-4 rounded-xl border border-slate-200 bg-white/95 p-3 shadow-sm backdrop-blur">
          <label htmlFor="dc-roll-search" className="block text-sm font-semibold">Find a student in this DC</label>
          <div className="mt-1 flex gap-2">
            <input
              ref={searchInputRef}
              id="dc-roll-search"
              type="search"
              inputMode="text"
              autoCapitalize="characters"
              autoCorrect="off"
              spellCheck={false}
              value={studentSearch}
              onChange={(event) => { setStudentSearch(event.target.value); setLookupResult(null); }}
              onKeyDown={(event) => { if (event.key === 'Enter') { event.preventDefault(); resolveStudentSearch(); } }}
              placeholder="Roll number or student name"
              className="min-h-12 w-full rounded-lg border border-slate-300 px-3 text-base"
            />
            {studentSearch && <button type="button" aria-label="Clear search" title="Clear search" className={`${buttonClass} btn-secondary min-w-12 px-3`} onClick={() => { setStudentSearch(''); setLookupResult(null); searchInputRef.current?.focus(); }}>×</button>}
          </div>
          <div className="mt-3 flex flex-wrap gap-2" aria-label="Filter students by attendance">
            {([
              ['all', `All (${attendanceCounts.total})`],
              ['Present', `Present (${attendanceCounts.present})`],
              ['Absent', `Absent (${attendanceCounts.absent})`],
              ['Pending', `Pending (${attendanceCounts.pending})`],
            ] as const).map(([filter, label]) => <button key={filter} type="button" aria-pressed={attendanceFilter === filter} className={`${buttonClass} ${attendanceFilter === filter ? 'btn-primary' : 'btn-secondary'}`} onClick={() => setAttendanceFilter(filter)}>{label}</button>)}
            <span className="self-center text-sm text-slate-600">{filteredEntries.length} of {attendanceCounts.total} students</span>
          </div>
          <label className="mt-3 block max-w-xs text-sm">Filter by batch<select value={attendanceBatchFilter} onChange={(event) => setAttendanceBatchFilter(event.target.value)} className={fieldClass}><option value="">All batches</option><option value="junior">Junior batch</option><option value="senior">Senior batch</option><option value="not_set">Not set</option></select></label>
          {attendanceNearlyComplete && (detail.dc.status === 'Planned' || detail.dc.status === 'In progress') &&
            <button type="button" className={`${buttonClass} btn-danger mt-3`} disabled={busy} onClick={() => void markPendingAbsent()}>
              Mark all {attendanceCounts.pending} pending as absent
            </button>}
          {lookupResult?.query === studentSearch && <p role="status" className="mt-3 rounded-lg bg-amber-50 p-3 text-sm text-amber-900">
            {lookupResult.kind === 'outside' && lookupResult.student
              ? <>{lookupResult.student.name} ({lookupResult.student.rollNo}) is <strong>Not in this DC.</strong> {detail.dc.status === 'Planned' && <button type="button" className={`${buttonClass} btn-primary ml-2`} onClick={() => void addSearchedStudent()} disabled={busy}>Add to this DC</button>}</>
              : 'Roll number not found'}
          </p>}
        </div>
        {detail.dc.status === 'Planned' && <div className="mt-5 flex flex-wrap gap-2">
          <form onSubmit={rescheduleDc} className="grid w-full gap-3 rounded-xl border border-slate-200 p-4 sm:grid-cols-2">
            <h3 className="font-semibold sm:col-span-2">Reschedule planned DC</h3>
            <label className="text-sm font-medium">Date<input required type="date" value={reschedule.date} onChange={(event) => setReschedule({ ...reschedule, date: event.target.value })} className={fieldClass} /></label>
            <label className="text-sm font-medium">Time<input type="time" value={reschedule.time} onChange={(event) => setReschedule({ ...reschedule, time: event.target.value })} className={fieldClass} /></label>
            <label className="text-sm font-medium">Venue<input value={reschedule.venue} onChange={(event) => setReschedule({ ...reschedule, venue: event.target.value })} maxLength={200} className={fieldClass} /></label>
            <label className="text-sm font-medium">Notes<input value={reschedule.notes} onChange={(event) => setReschedule({ ...reschedule, notes: event.target.value })} maxLength={2000} className={fieldClass} /></label>
            <label className="text-sm font-medium sm:col-span-2">Reason for rescheduling *<textarea required value={rescheduleReason} onChange={(event) => setRescheduleReason(event.target.value)} maxLength={1000} rows={2} className={fieldClass} /></label>
            <button disabled={busy} className={`${buttonClass} btn-secondary sm:col-span-2 sm:justify-self-start`}>Reschedule (reason required)</button>
          </form>
        </div>}
        {detail.dc.notes && <p className="mt-3 rounded-lg bg-slate-50 p-3 text-sm text-slate-600">{detail.dc.notes}</p>}
        {!detail.entries.length && <p className="mt-5 rounded-xl bg-slate-50 p-5 text-sm text-slate-600">No students are assigned to this DC yet.</p>}
        {detail.entries.length > 0 && !filteredEntries.length && !lookupResult && <p className="mt-5 rounded-xl bg-slate-50 p-5 text-sm text-slate-600">No students match this search and attendance filter.</p>}
        <div className="mt-5 space-y-3">
          {filteredEntries.map((entry) => <article key={entry.rollNo} className="rounded-xl border border-slate-200 bg-white">
            <button type="button" aria-expanded={expandedRollNo === entry.rollNo} className="flex min-h-14 w-full flex-wrap items-center justify-between gap-3 p-4 text-left hover:bg-slate-50" onClick={() => setExpandedRollNo((current) => current === entry.rollNo ? '' : entry.rollNo)}>
              <span><strong>{entry.studentName}</strong><span className="ml-2 text-sm text-slate-500">{entry.rollNo} · {entry.batch ? `${entry.batch === 'junior' ? 'Junior' : 'Senior'} batch` : 'Batch not set'} · {entry.defaults.length} defaults</span></span>
              <span className="flex items-center gap-2"><span className={statusStyle(entry.attendance)}>{entry.attendance}</span><span aria-hidden="true">{expandedRollNo === entry.rollNo ? '−' : '+'}</span></span>
            </button>
            {expandedRollNo === entry.rollNo && <div className="border-t border-slate-200 p-4">
            <div className="flex flex-wrap items-center justify-between gap-2"><div><h3 className="font-semibold">{entry.studentName}</h3><p className="text-sm text-slate-500">{entry.rollNo}</p></div><button type="button" className={`${buttonClass} btn-secondary`} onClick={() => void onOpenStudent(entry.rollNo)}>Student page &amp; history</button></div>
            <div className="mt-4 grid gap-3 sm:grid-cols-2">
              <label className="text-sm font-medium">Attendance<select disabled={(detail.dc.status !== 'Planned' && detail.dc.status !== 'In progress') || busy} value={entry.attendance} onChange={(event) => void updateAttendance(entry, event.target.value as Attendance)} className={fieldClass}><option>Pending</option><option>Present</option><option>Absent</option></select></label>
              {entry.attendance === 'Absent' && <label className="text-sm font-medium">Absent reason (optional)<input disabled={(detail.dc.status !== 'Planned' && detail.dc.status !== 'In progress') || busy} defaultValue={entry.absentReason} onBlur={(event) => void saveAbsentReason(entry, event.target.value)} placeholder="e.g. informed in advance" maxLength={500} className={fieldClass} /></label>}
            </div>
            {entry.defaults.length > 0 && <div className="mt-4 space-y-3">
              <h4 className="font-semibold">Defaults at this DC</h4>
              {entry.defaults.map((record) => <div key={record.id} className="grid gap-3 rounded-lg bg-slate-50 p-3 sm:grid-cols-[minmax(0,1fr)_12rem] sm:items-center">
                <div><strong>{record.company}</strong><p className="text-sm text-slate-600">{record.typeName} · logged {record.date} · current status {record.status}</p><span className={statusStyle(record.decision)}>{record.decision}</span></div>
                <label className="text-sm font-medium">Decision<select disabled={(detail.dc.status !== 'Planned' && detail.dc.status !== 'In progress') || busy} value={record.decision} onChange={(event) => void chooseDecision(entry, record, event.target.value as DCDecision)} className={fieldClass}><option>Undecided</option><option>Remove</option><option>Keep</option><option>Discuss with team</option><option>Defer</option></select></label>
                {decisionDraft?.defaultId === record.id && <form onSubmit={(event) => void saveDecisionDraft(event)} className="space-y-3 rounded-lg border border-indigo-100 bg-white p-3 sm:col-span-2">
                  <p className="text-sm font-semibold">{decisionDraft.decision === 'Defer' ? 'Deferral details' : 'Team discussion details'}</p>
                  <label className="block text-sm font-medium">Required reason *<textarea required value={decisionDraft.reason} onChange={(event) => setDecisionDraft({ ...decisionDraft, reason: event.target.value })} rows={2} maxLength={2000} className={fieldClass} /></label>
                  {decisionDraft.decision === 'Defer' && <div className="block text-sm font-medium">Deferred until *<select aria-label="Deferred until" value={decisionDraft.until === 'next_dc' ? 'next_dc' : decisionDraft.until ? 'specific' : ''} onChange={(event) => setDecisionDraft({ ...decisionDraft, until: event.target.value === 'next_dc' ? 'next_dc' : '' })} className={fieldClass}><option value="">Choose date or next DC</option><option value="specific">Specific date</option><option value="next_dc">Next DC</option></select>{decisionDraft.until !== 'next_dc' && <input aria-label="Specific deferral date" required type="date" value={decisionDraft.until} onChange={(event) => setDecisionDraft({ ...decisionDraft, until: event.target.value })} className={fieldClass} />}</div>}
                  {decisionDraft.decision === 'Discuss with team' && <label className="block text-sm font-medium">Discussion date (optional)<input type="date" value={decisionDraft.discussionDate} onChange={(event) => setDecisionDraft({ ...decisionDraft, discussionDate: event.target.value })} className={fieldClass} /></label>}
                  <div className="flex flex-wrap gap-2"><button type="submit" className={`${buttonClass} btn-primary`} disabled={busy}>Save decision</button><button type="button" className={`${buttonClass} btn-secondary`} onClick={() => setDecisionDraft(null)} disabled={busy}>Cancel</button></div>
                </form>}
              </div>)}
            </div>}
            <label className="mt-4 block text-sm font-medium">Remarks<textarea disabled={(detail.dc.status !== 'Planned' && detail.dc.status !== 'In progress') || busy} defaultValue={entry.remarks} onBlur={(event) => void saveRemarks(entry, event.target.value)} rows={2} maxLength={2000} placeholder="Reason accepted, warning issued…" className={fieldClass} /></label>
            {detail.dc.status === 'Planned' && <button type="button" className={`${buttonClass} btn-danger mt-3`} onClick={() => void removeStudent(entry)} disabled={busy}>Remove student from DC</button>}
            </div>}
          </article>)}
        </div>
        {(detail.dc.status === 'Planned' || detail.dc.status === 'In progress') && <div className="mt-5 flex flex-wrap items-center justify-between gap-3 border-t border-slate-200 pt-4">
          <p className="text-sm text-slate-600">{attendanceComplete ? 'All students have an attendance status.' : 'Mark every student Present or Absent before completing this DC.'}</p>
          <button type="button" className={`${buttonClass} btn-primary`} disabled={!attendanceComplete || busy} onClick={() => void completeDc()}>Complete DC</button>
        </div>}
        {detail.dc.status === 'Cancelled' && <p className="mt-5 rounded-lg bg-rose-50 p-3 text-sm text-rose-800">Cancelled: {detail.dc.cancelReason}</p>}
      </section>}
    </div>
  );
}
