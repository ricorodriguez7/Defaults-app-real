'use client';

import { useEffect, useMemo, useState } from 'react';
import * as XLSX from 'xlsx';
import { downloadWorkbook } from '@/lib/excel-download';

type BatchCounts = {
  called: number;
  present: number;
  absent: number;
  removed: number;
  kept: number;
  toDiscuss: number;
  deferred: number;
};

type DashboardDc = {
  id: string;
  date: string;
  status: string;
  source?: string;
  imported?: boolean;
  removedCount: number;
  keptCount: number;
  called: number;
  present: number;
  absent: number;
  toDiscussCount: number;
  deferredCount: number;
  batchCounts?: { junior: BatchCounts; senior: BatchCounts; not_set: BatchCounts };
};

type RemovedDefault = {
  id: string;
  rollNo: string;
  studentName: string;
  batch?: 'junior' | 'senior';
  company: string;
  typeName: string;
  date: string;
  removeReason: string;
};

type DetailResponse = { dc: DashboardDc; records: RemovedDefault[] };
type BatchFilter = 'both' | 'junior' | 'senior';

async function responseJson(response: Response) {
  const data = await response.json();
  if (!response.ok) throw new Error(data.message || `Request failed (${response.status}).`);
  return data;
}

function batchLabel(batch?: string) {
  return batch === 'junior' ? 'Junior' : batch === 'senior' ? 'Senior' : 'Not set';
}

function displayStatus(dc: DashboardDc) {
  if (dc.source === 'import') return 'Finalised (imported)';
  if (dc.status === 'In progress') return 'Held';
  if (dc.status === 'Completed') return 'Finalised';
  return dc.status;
}

export default function DashboardDCTab() {
  const [dcs, setDcs] = useState<DashboardDc[]>([]);
  const [range, setRange] = useState({ from: '', to: '' });
  const [batch, setBatch] = useState<BatchFilter>('both');
  const [selectedId, setSelectedId] = useState('');
  const [detail, setDetail] = useState<DetailResponse | null>(null);
  const [loading, setLoading] = useState(true);
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState('');

  useEffect(() => {
    let active = true;
    void fetch('/api/dashboard/dcs', { cache: 'no-store' }).then(responseJson)
      .then((data: DashboardDc[]) => { if (active) setDcs(data); })
      .catch((caught) => { if (active) setError(caught instanceof Error ? caught.message : 'Unable to load DC data.'); })
      .finally(() => { if (active) setLoading(false); });
    return () => { active = false; };
  }, []);

  useEffect(() => {
    if (!selectedId) {
      setDetail(null);
      return;
    }
    let active = true;
    setBusy(true);
    void fetch(`/api/dashboard/dcs/${encodeURIComponent(selectedId)}`, { cache: 'no-store' }).then(responseJson)
      .then((data: DetailResponse) => { if (active) setDetail(data); })
      .catch((caught) => { if (active) setError(caught instanceof Error ? caught.message : 'Unable to load removed defaults.'); })
      .finally(() => { if (active) setBusy(false); });
    return () => { active = false; };
  }, [selectedId]);

  const visibleDcs = useMemo(() => dcs.filter((dc) =>
    (!range.from || !dc.date || dc.date >= range.from) &&
    (!range.to || !dc.date || dc.date <= range.to),
  ).sort((a, b) => b.date.localeCompare(a.date) || a.id.localeCompare(b.id)), [dcs, range]);

  function countsFor(dc: DashboardDc): BatchCounts {
    if (batch === 'both' || !dc.batchCounts) {
      return {
        called: dc.called,
        present: dc.present,
        absent: dc.absent,
        removed: dc.removedCount,
        kept: dc.keptCount,
        toDiscuss: dc.toDiscussCount,
        deferred: dc.deferredCount,
      };
    }
    return dc.batchCounts[batch];
  }

  const totals = visibleDcs.reduce((sum, dc) => {
    const counts = countsFor(dc);
    return {
      called: sum.called + counts.called,
      present: sum.present + counts.present,
      absent: sum.absent + counts.absent,
      removed: sum.removed + counts.removed,
      kept: sum.kept + counts.kept,
      outstanding: sum.outstanding + counts.toDiscuss + counts.deferred,
    };
  }, { called: 0, present: 0, absent: 0, removed: 0, kept: 0, outstanding: 0 });

  const detailRecords = useMemo(() => detail?.records.filter((record) =>
    batch === 'both' || record.batch === batch,
  ) ?? [], [detail, batch]);

  async function exportWorkbook(records: DashboardDc[]) {
    setBusy(true);
    setError('');
    try {
      const details = await Promise.all(records.map((dc) =>
        fetch(`/api/dashboard/dcs/${encodeURIComponent(dc.id)}`, { cache: 'no-store' }).then(responseJson) as Promise<DetailResponse>,
      ));
      const summaryRows = records.map((dc) => {
        const counts = countsFor(dc);
        return {
          'DC date': dc.date || 'No DC date (imported)',
          Status: displayStatus(dc),
          Batch: batch === 'both' ? 'Both' : batchLabel(batch),
          'Students called': counts.called,
          Present: counts.present,
          Absent: counts.absent,
          'Defaults removed': counts.removed,
          'Defaults kept (locked)': counts.kept,
          'Deferred or sent for discussion': counts.deferred + counts.toDiscuss,
        };
      });
      const removedRows = details.flatMap((item) => item.records
        .filter((record) => batch === 'both' || record.batch === batch)
        .map((record) => ({
          'DC date': item.dc.date || 'No DC date (imported)',
          'Date logged': record.date,
          Student: record.studentName,
          'Roll number': record.rollNo,
          Batch: batchLabel(record.batch),
          Company: record.company,
          'Default type': record.typeName,
          'Removal reason or remarks': record.removeReason || '',
        })));
      const workbook = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(summaryRows), 'DC summary');
      XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(removedRows), 'Removed defaults');
      downloadWorkbook(workbook, records.length === 1 ? `dc-${records[0].date || 'no-date'}-removals.xlsx` : 'dc-removals.xlsx');
    } catch (caught) {
      setError(caught instanceof Error ? caught.message : 'Unable to export DC data.');
    } finally {
      setBusy(false);
    }
  }

  const dateBreakdown = detailRecords.reduce<Record<string, number>>((result, record) => {
    result[record.date || 'Unknown date'] = (result[record.date || 'Unknown date'] ?? 0) + 1;
    return result;
  }, {});

  return <section className="space-y-5">
    <div className="flex flex-wrap items-end justify-between gap-3 rounded-2xl border border-slate-200 bg-white p-4 shadow-sm">
      <div>
        <h2 className="text-xl font-semibold">Disciplinary Committee</h2>
        <p className="mt-1 text-sm text-slate-500">DC outcomes and defaults removed by each meeting.</p>
      </div>
      <div className="flex flex-wrap items-end gap-2">
        <label className="text-sm">From<input type="date" value={range.from} onChange={(event) => setRange({ ...range, from: event.target.value })} className="mt-1 block rounded-lg border px-3 py-2" /></label>
        <label className="text-sm">To<input type="date" value={range.to} onChange={(event) => setRange({ ...range, to: event.target.value })} className="mt-1 block rounded-lg border px-3 py-2" /></label>
        <label className="text-sm">Batch<select value={batch} onChange={(event) => setBatch(event.target.value as BatchFilter)} className="mt-1 block rounded-lg border px-3 py-2"><option value="both">Both</option><option value="junior">Junior</option><option value="senior">Senior</option></select></label>
        <button type="button" disabled={busy || !visibleDcs.length} onClick={() => void exportWorkbook(visibleDcs)} className="min-h-11 rounded-lg border px-3 py-2 text-sm">Export all DCs</button>
      </div>
    </div>
    {error && <p role="alert" className="rounded-lg border border-rose-200 bg-rose-50 p-3 text-sm text-rose-800">{error}</p>}
    {loading ? <p role="status" className="text-sm text-slate-500">Loading DC history…</p> : <div className="responsive-table rounded-2xl border border-slate-200 bg-white shadow-sm">
      <table className="min-w-full table-fixed text-left text-sm">
        <thead className="sticky top-0 z-10 bg-slate-50 text-slate-600"><tr>
          {['DC date', 'Status', 'Students called', 'Present', 'Absent', 'Defaults removed', 'Defaults kept (locked)', 'Deferred / discussion'].map((title) => <th key={title} className="px-3 py-3 font-semibold">{title}</th>)}
        </tr></thead>
        <tbody>
          {visibleDcs.map((dc) => {
            const counts = countsFor(dc);
            return <tr key={dc.id} className="border-t hover:bg-indigo-50/50">
              <td className="px-3 py-3"><button type="button" onClick={() => setSelectedId(dc.id)} className="text-left font-semibold text-indigo-700 underline-offset-2 hover:underline">{dc.date || 'No DC date (imported)'}</button></td>
              <td className="px-3 py-3">{displayStatus(dc)}</td><td className="px-3 py-3">{counts.called}</td><td className="px-3 py-3">{counts.present}</td><td className="px-3 py-3">{counts.absent}</td>
              <td className="px-3 py-3"><button type="button" className="font-semibold text-indigo-700" onClick={() => setSelectedId(dc.id)}>{counts.removed}</button></td>
              <td className="px-3 py-3">{counts.kept}</td><td className="px-3 py-3">{counts.deferred + counts.toDiscuss}</td>
            </tr>;
          })}
          <tr className="border-t-2 border-slate-300 bg-slate-100 font-bold"><td className="px-3 py-3" colSpan={2}>Totals</td><td className="px-3 py-3">{totals.called}</td><td className="px-3 py-3">{totals.present}</td><td className="px-3 py-3">{totals.absent}</td><td className="px-3 py-3">{totals.removed}</td><td className="px-3 py-3">{totals.kept}</td><td className="px-3 py-3">{totals.outstanding}</td></tr>
          {!visibleDcs.length && <tr><td colSpan={8} className="px-3 py-8 text-center text-slate-500">No DCs match this date range.</td></tr>}
        </tbody>
      </table>
    </div>}

    {detail && <section className="rounded-2xl border border-slate-200 bg-white p-5 shadow-sm">
      <div className="flex flex-wrap items-start justify-between gap-3">
        <div><p className="text-sm font-semibold uppercase tracking-wide text-indigo-700">{detail.dc.date || 'No DC date (imported)'}</p><h3 className="mt-1 text-xl font-bold">{detailRecords.length} defaults removed</h3></div>
        <div className="flex gap-2"><button type="button" disabled={busy} onClick={() => void exportWorkbook([detail.dc])} className="min-h-11 rounded-lg border px-3 py-2 text-sm">Export this DC</button><button type="button" onClick={() => setSelectedId('')} className="min-h-11 rounded-lg border px-3 py-2 text-sm">Close details</button></div>
      </div>
      {busy && <p role="status" className="mt-3 text-sm text-slate-500">Loading or exporting…</p>}
      <div className="mt-4 flex flex-wrap gap-2">{Object.entries(dateBreakdown).sort(([a], [b]) => a.localeCompare(b)).map(([date, count]) => <span key={date} className="rounded-full bg-slate-100 px-3 py-1.5 text-sm">{date}: {count} removed</span>)}</div>
      <div className="responsive-table mt-4">
        <table className="min-w-full table-fixed text-left text-sm">
          <thead className="bg-slate-50 text-slate-600"><tr>{['Student', 'Roll number', 'Batch', 'Company', 'Default type', 'Date logged', 'Removal reason / remarks'].map((title) => <th key={title} className="px-3 py-2">{title}</th>)}</tr></thead>
          <tbody>{detailRecords.map((record) => <tr key={record.id} className="border-t text-slate-600"><td className="px-3 py-2">{record.studentName}</td><td className="px-3 py-2">{record.rollNo}</td><td className="px-3 py-2">{batchLabel(record.batch)}</td><td className="px-3 py-2">{record.company}</td><td className="px-3 py-2">{record.typeName}</td><td className="px-3 py-2">{record.date}</td><td className="px-3 py-2">{record.removeReason || '—'}</td></tr>)}
            {!detailRecords.length && <tr><td colSpan={7} className="px-3 py-6 text-center text-slate-500">No removed defaults are recorded for this DC and batch.</td></tr>}
          </tbody>
        </table>
      </div>
    </section>}
  </section>;
}
