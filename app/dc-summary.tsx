'use client';

import { useEffect, useMemo, useState } from 'react';

type DcOverview = {
  id: string;
  date: string;
  time: string;
  venue: string;
  status: 'Planned' | 'In progress' | 'Completed' | 'Cancelled';
  notes: string;
  autoAdd?: boolean;
  isCurrent?: boolean;
  counts?: {
    studentsCalled?: number;
    yetToAttend?: number;
    attended?: number;
    absent?: number;
    defaultsToBeRemoved?: number;
    toDiscuss?: number;
    deferred?: number;
  };
};

type Entry = {
  rollNo: string;
  studentName: string;
  batch?: 'junior' | 'senior';
  attendance: 'Pending' | 'Present' | 'Absent';
  absentReason: string;
  remarks: string;
  defaults: Array<{ id: string; company: string; typeName: string; date: string; decision: string; reason?: string; until?: string; discussionDate?: string }>;
};

type DcDetail = { dc: DcOverview; entries: Entry[] };

function sumCounts(items: Array<{ [key: string]: number | string | undefined } | null | undefined>, key: string) {
  return items.reduce((total, item) => total + Number(item?.[key] ?? 0), 0);
}

export default function DCSummaryTab() {
  const [dcs, setDcs] = useState<DcOverview[]>([]);
  const [selectedDcId, setSelectedDcId] = useState('');
  const [detail, setDetail] = useState<DcDetail | null>(null);
  const [range, setRange] = useState({ from: '', to: '' });
  const [decisionFilter, setDecisionFilter] = useState('all');
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState('');
  const [autoAddMessage, setAutoAddMessage] = useState('');

  useEffect(() => {
    let active = true;
    async function loadOverview() {
      setLoading(true);
      try {
        const response = await fetch('/api/dcs', { cache: 'no-store' });
        const data = await response.json();
        if (!response.ok) throw new Error(data.message || 'Unable to load DC summary.');
        if (!active) return;
        const nextDcs = Array.isArray(data.dcs) ? data.dcs : [];
        setDcs(nextDcs);
        setAutoAddMessage(data.autoAdd?.message ?? '');
        const current = nextDcs.find((dc: DcOverview) => dc.isCurrent) ?? nextDcs[0];
        setSelectedDcId((existing) => existing || current?.id || '');
      } catch (caught) {
        if (active) setError(caught instanceof Error ? caught.message : 'Unable to load DC summary.');
      } finally {
        if (active) setLoading(false);
      }
    }
    void loadOverview();
    return () => { active = false; };
  }, []);

  useEffect(() => {
    if (!selectedDcId) {
      setDetail(null);
      return;
    }
    let active = true;
    async function loadDetail() {
      try {
        const response = await fetch(`/api/dcs/${encodeURIComponent(selectedDcId)}`, { cache: 'no-store' });
        const data = await response.json();
        if (!response.ok) throw new Error(data.message || 'Unable to load the selected DC.');
        if (active) setDetail(data);
      } catch (caught) {
        if (active) setError(caught instanceof Error ? caught.message : 'Unable to load the selected DC.');
      }
    }
    void loadDetail();
    return () => { active = false; };
  }, [selectedDcId]);

  const visibleDcList = useMemo(() => {
    return [...dcs]
      .filter((dc) => !range.from || dc.date >= range.from)
      .filter((dc) => !range.to || dc.date <= range.to)
      .sort((a, b) => b.date.localeCompare(a.date) || b.time.localeCompare(a.time));
  }, [dcs, range]);

  const currentDc = dcs.find((dc) => dc.isCurrent) ?? dcs[0];
  const activeDc = detail?.dc ?? dcs.find((dc) => dc.id === selectedDcId) ?? currentDc ?? null;

  const summaryCards = useMemo(() => {
    if (!detail) {
      return {
        studentsCalled: 0,
        present: 0,
        absent: 0,
        pending: 0,
        defaultsRemoved: 0,
        defaultsKept: 0,
        sentForDiscussion: 0,
        deferred: 0,
        undecided: 0,
      };
    }

    const entries = detail.entries ?? [];
    const defaultRows = entries.flatMap((entry) => entry.defaults ?? []);
    return {
      studentsCalled: entries.length,
      present: entries.filter((entry) => entry.attendance === 'Present').length,
      absent: entries.filter((entry) => entry.attendance === 'Absent').length,
      pending: entries.filter((entry) => entry.attendance === 'Pending').length,
      defaultsRemoved: defaultRows.filter((item) => item.decision === 'Remove').length,
      defaultsKept: defaultRows.filter((item) => item.decision === 'Keep').length,
      sentForDiscussion: defaultRows.filter((item) => item.decision === 'Discuss with team').length,
      deferred: defaultRows.filter((item) => item.decision === 'Defer').length,
      undecided: defaultRows.filter((item) => item.decision === 'Undecided').length,
    };
  }, [detail]);
  const batchSummary = useMemo(() => {
    const groups = [
      { key: 'junior', label: 'Junior batch' },
      { key: 'senior', label: 'Senior batch' },
      { key: 'not_set', label: 'Not set' },
    ];
    return groups.map((group) => {
      const entries = (detail?.entries ?? []).filter((entry) =>
        group.key === 'not_set' ? !entry.batch : entry.batch === group.key,
      );
      const defaults = entries.flatMap((entry) => entry.defaults ?? []);
      return {
        ...group,
        called: entries.length,
        present: entries.filter((entry) => entry.attendance === 'Present').length,
        absent: entries.filter((entry) => entry.attendance === 'Absent').length,
        pending: entries.filter((entry) => entry.attendance === 'Pending').length,
        removed: defaults.filter((item) => item.decision === 'Remove').length,
        kept: defaults.filter((item) => item.decision === 'Keep').length,
        undecided: defaults.filter((item) => item.decision === 'Undecided').length,
      };
    });
  }, [detail]);

  const absentees = detail?.entries.filter((entry) => entry.attendance === 'Absent') ?? [];
  const decisionRows = useMemo(() => {
    if (!detail) return [] as Array<{ label: string; students: string[]; count: number }>;
    const groups = new Map<string, { label: string; students: string[] }>();
    for (const entry of detail.entries) {
      for (const defaultItem of entry.defaults ?? []) {
        const label = defaultItem.decision || 'Undecided';
        if (!groups.has(label)) groups.set(label, { label, students: [] });
        groups.get(label)?.students.push(`${entry.rollNo} — ${entry.studentName}`);
      }
    }
    return Array.from(groups.entries()).map(([label, value]) => ({
      label,
      students: value.students,
      count: value.students.length,
    })).filter((row) => decisionFilter === 'all' || row.label === decisionFilter);
  }, [decisionFilter, detail]);

  const carriedForward = detail ? detail.entries.filter((entry) => (entry.defaults ?? []).some((item) => ['Undecided', 'Discuss with team', 'Defer'].includes(item.decision))) : [];

  const overviewTotals = useMemo(() => {
    const dcRows = visibleDcList.map((dc) => ({
      date: dc.date,
      status: dc.status,
      called: dc.counts?.studentsCalled ?? 0,
      present: dc.counts?.attended ?? 0,
      absent: dc.counts?.absent ?? 0,
      removed: dc.counts?.defaultsToBeRemoved ?? 0,
      kept: dc.counts?.attended ?? 0,
      deferred: dc.counts?.deferred ?? 0,
      discussed: dc.counts?.toDiscuss ?? 0,
    }));
    return {
      totalCalled: sumCounts(dcRows, 'called'),
      totalPresent: sumCounts(dcRows, 'present'),
      totalAbsent: sumCounts(dcRows, 'absent'),
      totalRemoved: sumCounts(dcRows, 'removed'),
      totalKept: sumCounts(dcRows, 'kept'),
      totalDeferred: sumCounts(dcRows, 'deferred'),
      totalDiscussed: sumCounts(dcRows, 'discussed'),
      rows: dcRows,
    };
  }, [visibleDcList]);

  const maxChartValue = Math.max(1, ...overviewTotals.rows.map((row) => Math.max(Number(row.called), Number(row.absent))));

  return (
    <section className="space-y-6">
      <div className="rounded-2xl border border-slate-200 bg-white p-4 shadow-sm">
        <div className="flex flex-col gap-3 md:flex-row md:items-end md:justify-between">
          <div>
            <p className="text-xs font-semibold uppercase tracking-[0.2em] text-indigo-600">Disciplinary Committee</p>
            <h2 className="mt-2 text-2xl font-bold text-slate-900">Summary</h2>
          </div>

          <section className="rounded-2xl border border-slate-200 bg-white p-5 shadow-sm">
            <h3 className="text-lg font-semibold text-slate-900">Batch breakdown</h3>
            <div className="responsive-table mt-3"><table className="min-w-full text-left text-sm">
              <thead className="border-b bg-slate-50 text-slate-600"><tr>{['Batch', 'Called', 'Present', 'Absent', 'Pending', 'Remove decisions', 'Keep decisions', 'Undecided'].map((label) => <th key={label} className="px-3 py-2">{label}</th>)}</tr></thead>
              <tbody>{batchSummary.map((row) => <tr key={row.key} className="border-b"><td className="px-3 py-2">{row.label}</td><td className="px-3 py-2">{row.called}</td><td className="px-3 py-2">{row.present}</td><td className="px-3 py-2">{row.absent}</td><td className="px-3 py-2">{row.pending}</td><td className="px-3 py-2">{row.removed}</td><td className="px-3 py-2">{row.kept}</td><td className="px-3 py-2">{row.undecided}</td></tr>)}</tbody>
            </table></div>
          </section>
          <div className="flex flex-col gap-2 sm:flex-row sm:items-center">
            <label className="text-sm text-slate-700">DC
              <select value={selectedDcId} onChange={(event) => setSelectedDcId(event.target.value)} className="mt-1 min-h-11 rounded-lg border border-slate-300 bg-white px-3 py-2 text-sm">
                {dcs.map((dc) => (
                  <option key={dc.id} value={dc.id}>{dc.date} {dc.status === 'Planned' && dc.isCurrent ? '• Current' : ''}</option>
                ))}
              </select>
            </label>
            <button type="button" onClick={() => window.print()} className="min-h-11 rounded-lg border border-slate-200 bg-white px-4 py-2 text-sm font-medium text-slate-700">Print</button>
          </div>
        </div>
      </div>

      {error && <p className="rounded-lg border border-rose-200 bg-rose-50 px-4 py-3 text-sm text-rose-700">{error}</p>}
      {autoAddMessage && <p role="status" className="rounded-lg border border-indigo-100 bg-indigo-50 px-4 py-3 text-sm text-indigo-900">{autoAddMessage}</p>}

      {!loading && activeDc && (
        <>
          <div className="grid gap-4 md:grid-cols-2 xl:grid-cols-5">
            {[
              ['Students called', summaryCards.studentsCalled],
              ['Present', summaryCards.present],
              ['Absent', summaryCards.absent],
              ['Pending', summaryCards.pending],
              ['Defaults removed', summaryCards.defaultsRemoved],
              ['Defaults kept', summaryCards.defaultsKept],
              ['Sent for team discussion', summaryCards.sentForDiscussion],
              ['Deferred', summaryCards.deferred],
              ['Undecided', summaryCards.undecided],
            ].map(([label, value]) => (
              <div key={label} className="rounded-2xl border border-slate-200 bg-white p-4 shadow-sm">
                <p className="text-sm text-slate-500">{label}</p>
                <p className="mt-3 text-3xl font-bold text-slate-900">{value}</p>
              </div>
            ))}
          </div>

          <div className="grid gap-6 xl:grid-cols-[1.1fr_0.9fr]">
            <div className="rounded-2xl border border-slate-200 bg-white p-5 shadow-sm">
              <h3 className="text-lg font-semibold text-slate-900">Absentees</h3>
              <p className="mt-1 text-sm text-slate-500">{absentees.length} total absences</p>
              <div className="mt-4 space-y-2">
                {absentees.length ? absentees.map((entry) => (
                  <div key={entry.rollNo} className="flex items-center justify-between gap-3 rounded-lg border border-slate-200 bg-slate-50 px-3 py-2 text-sm">
                    <span><strong>{entry.rollNo}</strong> · {entry.studentName} · {entry.batch ? `${entry.batch === 'junior' ? 'Junior' : 'Senior'} batch` : 'Not set'}</span>
                    <span className="text-slate-600">{entry.absentReason || 'No reason'}</span>
                  </div>
                )) : <p className="text-sm text-slate-500">No absentees recorded for this DC.</p>}
              </div>
            </div>

            <div className="rounded-2xl border border-slate-200 bg-white p-5 shadow-sm">
              <h3 className="text-lg font-semibold text-slate-900">Decision status</h3>
              <div className="mt-3 flex flex-wrap gap-2">
                {['all', 'Undecided', 'Remove', 'Keep', 'Discuss with team', 'Defer'].map((option) => (
                  <button key={option} type="button" onClick={() => setDecisionFilter(option)} className={`min-h-9 rounded-full px-3 py-1.5 text-xs font-medium ${decisionFilter === option ? 'bg-indigo-600 text-white' : 'bg-slate-100 text-slate-700'}`}>
                    {option === 'all' ? 'All' : option}
                  </button>
                ))}
              </div>
              <div className="mt-4 space-y-3">
                {decisionRows.length ? decisionRows.map((row) => (
                  <div key={row.label} className="rounded-lg border border-slate-200 bg-slate-50 p-3">
                    <div className="flex items-center justify-between gap-3">
                      <span className="font-semibold text-slate-800">{row.label}</span>
                      <span className="text-xs text-slate-500">{row.count}</span>
                    </div>
                    <ul className="mt-2 space-y-1 text-sm text-slate-600">
                      {row.students.slice(0, 6).map((student) => <li key={student}>{student}</li>)}
                    </ul>
                  </div>
                )) : <p className="text-sm text-slate-500">No decisions match the selected filter.</p>}
              </div>
            </div>
          </div>

          <div className="grid gap-6 xl:grid-cols-2">
            <div className="rounded-2xl border border-slate-200 bg-white p-5 shadow-sm">
              <h3 className="text-lg font-semibold text-slate-900">Carried forward</h3>
              <div className="mt-4 space-y-2">
                {carriedForward.length ? carriedForward.map((entry) => (
                  <div key={entry.rollNo} className="rounded-lg border border-slate-200 bg-slate-50 p-3 text-sm">
                    <div className="font-medium text-slate-800">{entry.rollNo} · {entry.studentName}</div>
                    <ul className="mt-1 list-disc pl-4 text-slate-600">
                      {(entry.defaults ?? []).filter((item) => ['Undecided', 'Discuss with team', 'Defer'].includes(item.decision)).map((item) => (
                        <li key={item.id}>{item.typeName} — {item.decision}</li>
                      ))}
                    </ul>
                  </div>
                )) : <p className="text-sm text-slate-500">Nothing is carried forward to the next DC.</p>}
              </div>
            </div>

            <div className="rounded-2xl border border-slate-200 bg-white p-5 shadow-sm">
              <h3 className="text-lg font-semibold text-slate-900">Overall DC overview</h3>
              <div className="mt-3 flex flex-col gap-2 sm:flex-row sm:items-center">
                <input type="date" value={range.from} onChange={(event) => setRange((current) => ({ ...current, from: event.target.value }))} className="min-h-11 rounded-lg border border-slate-300 bg-white px-3 py-2 text-sm" />
                <input type="date" value={range.to} onChange={(event) => setRange((current) => ({ ...current, to: event.target.value }))} className="min-h-11 rounded-lg border border-slate-300 bg-white px-3 py-2 text-sm" />
              </div>
              <div className="mt-4 space-y-3">
                {overviewTotals.rows.map((row) => (
                  <div key={`${row.date}-${row.status}`} className="rounded-lg border border-slate-200 bg-slate-50 p-3">
                    <div className="flex items-center justify-between text-sm text-slate-700">
                      <span>{row.date}</span>
                      <span>{row.status}</span>
                    </div>
                    <div className="mt-2 flex items-end gap-2">
                      <div className="h-2.5 w-full rounded-full bg-slate-200">
                        <div className="h-2.5 rounded-full bg-indigo-500" style={{ width: `${((Number(row.called) || 0) / maxChartValue) * 100}%` }} />
                      </div>
                      <span className="text-xs text-slate-500">{row.called}</span>
                    </div>
                    <div className="mt-2 flex justify-between text-xs text-slate-500">
                      <span>Called {row.called}</span>
                      <span>Absent {row.absent}</span>
                    </div>
                  </div>
                ))}
              </div>
            </div>
          </div>
        </>
      )}

      {!loading && !activeDc && <p className="rounded-lg border border-slate-200 bg-white p-6 text-sm text-slate-600">No DC data is available yet.</p>}
    </section>
  );
}
