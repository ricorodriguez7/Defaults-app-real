'use client';

import { useEffect, useMemo, useState } from 'react';
import * as XLSX from 'xlsx';

type DCRecord = {
  id: string;
  date: string;
  time?: string;
  venue?: string;
  status: 'Held' | 'Finalised';
  heldAt?: string;
  finalisedAt?: string;
  finalisedByName?: string;
  counts?: { studentsCalled?: number; attended?: number; absent?: number; awaitingTeamDecision?: number };
};

type DCDefault = {
  id: string;
  company: string;
  typeName: string;
  date: string;
  recommendation: 'undecided' | 'remove' | 'keep' | 'defer' | 'discuss';
  recommendationNotes: string;
  recommendationByName?: string;
  recommendationAt?: string;
  finalDecision?: '' | 'remove' | 'keep' | 'defer' | 'send_back';
  finalRemarks?: string;
  finalUntil?: string;
  finalByName?: string;
  finalAt?: string;
  discussionNotes: Array<{ text: string; byName?: string; at: string }>;
  until?: string;
};

type DCEntry = {
  rollNo: string;
  studentName: string;
  attendance: 'Pending' | 'Present' | 'Absent';
  absentReason: string;
  remarks: string;
  defaults: DCDefault[];
};

type DCDetail = { dc: DCRecord; entries: DCEntry[] };

const fieldClass = 'mt-1 min-h-11 w-full rounded-lg border border-slate-300 bg-white px-3 py-2 text-sm';
const buttonClass = 'inline-flex min-h-10 items-center justify-center rounded-lg px-3 py-2 text-sm font-semibold transition focus-visible:outline focus-visible:outline-2 focus-visible:outline-offset-2';
const recommendationLabel: Record<DCDefault['recommendation'], string> = {
  undecided: 'Undecided',
  remove: 'To be removed (considered)',
  keep: 'To be kept (not considered)',
  defer: 'Defer',
  discuss: 'Discuss with team',
};
const finalLabel: Record<NonNullable<DCDefault['finalDecision']>, string> = {
  '': 'No final decision',
  remove: 'Remove',
  keep: 'Keep/lock',
  defer: 'Defer',
  send_back: 'Send back to DC',
};

async function responseJson(response: Response) {
  const data = await response.json();
  if (!response.ok) throw new Error(data.message || `Request failed (${response.status}).`);
  return data;
}

export default function DCAfterDiscussionTab() {
  const [dcs, setDcs] = useState<DCRecord[]>([]);
  const [detail, setDetail] = useState<DCDetail | null>(null);
  const [selectedId, setSelectedId] = useState('');
  const [search, setSearch] = useState('');
  const [recommendationFilter, setRecommendationFilter] = useState('all');
  const [drafts, setDrafts] = useState<Record<string, { decision: string; remarks: string; until: string }>>({});
  const [notes, setNotes] = useState<Record<string, string>>({});
  const [selectedDefaults, setSelectedDefaults] = useState<string[]>([]);
  const [loading, setLoading] = useState(true);
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState('');
  const [message, setMessage] = useState('');

  async function loadOverview() {
    const data = await responseJson(await fetch('/api/dcs', { cache: 'no-store' }));
    const records = (data.dcs as DCRecord[]).filter((dc) => dc.status === 'Held' || dc.status === 'Finalised');
    setDcs(records);
    if (!selectedId && records.length) setSelectedId(records[0].id);
  }

  async function loadDetail(id: string) {
    const data = await responseJson(await fetch(`/api/dcs/${encodeURIComponent(id)}`, { cache: 'no-store' }));
    setDetail(data);
    setDrafts((existing) => {
      const next = { ...existing };
      for (const entry of data.entries as DCEntry[]) for (const record of entry.defaults) {
        next[record.id] = {
          decision: record.finalDecision ?? '',
          remarks: record.finalRemarks ?? '',
          until: record.finalUntil ?? '',
        };
      }
      return next;
    });
  }

  useEffect(() => {
    let active = true;
    void (async () => {
      setLoading(true);
      try {
        const data = await responseJson(await fetch('/api/dcs', { cache: 'no-store' }));
        if (!active) return;
        const records = (data.dcs as DCRecord[]).filter((dc) => dc.status === 'Held' || dc.status === 'Finalised');
        setDcs(records);
        setSelectedId((current) => current && records.some((dc) => dc.id === current) ? current : records[0]?.id ?? '');
      } catch (caught) {
        if (active) setError(caught instanceof Error ? caught.message : 'Unable to load held DCs.');
      } finally {
        if (active) setLoading(false);
      }
    })();
    return () => { active = false; };
  }, []);

  useEffect(() => {
    if (!selectedId) {
      setDetail(null);
      return;
    }
    let active = true;
    void (async () => {
      try {
        const data = await responseJson(await fetch(`/api/dcs/${encodeURIComponent(selectedId)}`, { cache: 'no-store' }));
        if (active) {
          setDetail(data);
          setDrafts((existing) => {
            const next = { ...existing };
            for (const entry of data.entries as DCEntry[]) for (const record of entry.defaults) {
              next[record.id] = { decision: record.finalDecision ?? '', remarks: record.finalRemarks ?? '', until: record.finalUntil ?? '' };
            }
            return next;
          });
        }
      } catch (caught) {
        if (active) setError(caught instanceof Error ? caught.message : 'Unable to load this DC.');
      }
    })();
    return () => { active = false; };
  }, [selectedId]);

  const allRows = useMemo(() => (detail?.entries ?? []).flatMap((entry) => entry.defaults.map((record) => ({ entry, record }))), [detail]);
  const discussionRows = useMemo(() => allRows.filter(({ entry, record }) =>
    entry.attendance === 'Present' && record.recommendation !== 'defer',
  ), [allRows]);
  const filteredRows = useMemo(() => {
    const query = search.trim().toLocaleLowerCase();
    return discussionRows.filter(({ entry, record }) =>
      (!query || `${entry.rollNo} ${entry.studentName} ${record.company} ${record.typeName}`.toLocaleLowerCase().includes(query)) &&
      (recommendationFilter === 'all' || record.recommendation === recommendationFilter),
    );
  }, [discussionRows, search, recommendationFilter]);
  const undecidedCount = discussionRows.filter(({ record }) => !record.finalDecision).length;
  const completeCount = discussionRows.length - undecidedCount;
  const deferrals = allRows.filter(({ record }) => record.recommendation === 'defer');
  const absentees = detail?.entries.filter((entry) => entry.attendance === 'Absent') ?? [];
  const isHeld = detail?.dc.status === 'Held';

  async function mutate(action: string, values: Record<string, unknown>, success: string) {
    if (!detail) return false;
    setBusy(true);
    setError('');
    setMessage('');
    try {
      const result = await responseJson(await fetch(`/api/dcs/${encodeURIComponent(detail.dc.id)}`, {
        method: 'PATCH',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ action, ...values }),
      }));
      if (result.autoAdd?.message) setMessage(`${success} ${result.autoAdd.message}`);
      else setMessage(success);
      await Promise.all([loadOverview(), loadDetail(detail.dc.id)]);
      return true;
    } catch (caught) {
      setError(caught instanceof Error ? caught.message : 'Unable to update the DC discussion.');
      return false;
    } finally {
      setBusy(false);
    }
  }

  async function saveNote(entry: DCEntry, record: DCDefault) {
    const note = (notes[record.id] ?? '').trim();
    if (!note) return;
    if (await mutate('discussion_note', { defaultId: record.id, rollNo: entry.rollNo, note }, 'Discussion note saved.')) {
      setNotes((current) => ({ ...current, [record.id]: '' }));
    }
  }

  async function saveFinalDecision(entry: DCEntry, record: DCDefault) {
    const draft = drafts[record.id] ?? { decision: '', remarks: '', until: '' };
    if (!draft.decision) {
      setError('Choose a final decision first.');
      return;
    }
    if (draft.decision === 'defer' && (!draft.remarks.trim() || !draft.until)) {
      setError('A reason and date are required for a final deferral.');
      return;
    }
    await mutate('final_decision', {
      defaultId: record.id,
      finalDecision: draft.decision,
      finalRemarks: draft.remarks.trim(),
      until: draft.until,
    }, 'Final decision saved.');
  }

  async function acceptSelected() {
    if (!detail || !selectedDefaults.length) return;
    setBusy(true);
    setError('');
    setMessage('');
    try {
      for (const id of selectedDefaults) {
        const row = discussionRows.find(({ record }) => record.id === id);
        if (!row || !['remove', 'keep'].includes(row.record.recommendation)) continue;
        await responseJson(await fetch(`/api/dcs/${encodeURIComponent(detail.dc.id)}`, {
          method: 'PATCH',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ action: 'final_decision', defaultId: id, finalDecision: row.record.recommendation, finalRemarks: '' }),
        }));
      }
      setSelectedDefaults([]);
      setMessage('Selected recommendations accepted as final decisions.');
      await Promise.all([loadOverview(), loadDetail(detail.dc.id)]);
    } catch (caught) {
      setError(caught instanceof Error ? caught.message : 'Unable to accept selected recommendations.');
    } finally {
      setBusy(false);
    }
  }

  async function finalise() {
    if (!detail || undecidedCount) return;
    const remove = discussionRows.filter(({ record }) => record.finalDecision === 'remove').length;
    const keep = discussionRows.filter(({ record }) => record.finalDecision === 'keep').length;
    const defer = discussionRows.filter(({ record }) => record.finalDecision === 'defer').length;
    const sendBack = discussionRows.filter(({ record }) => record.finalDecision === 'send_back').length;
    if (!window.confirm(`Finalise this DC?\n\n${remove} Remove · ${keep} Keep/lock · ${defer} Defer · ${sendBack} Send back to DC\n\nDefaults will be updated and this DC will become read-only.`)) return;
    await mutate('finalise', {}, 'DC finalised. Final decisions have been applied.');
  }

  function exportWorkbook(dc: DCDetail) {
    const isDraft = dc.dc.status !== 'Finalised';
    const results = dc.entries.flatMap((entry) => entry.defaults.map((record) => ({
      'Draft status': isDraft ? 'DRAFT — NOT FINALISED' : 'FINAL',
      'DC date': dc.dc.date,
      'Roll number': entry.rollNo,
      Student: entry.studentName,
      Attendance: entry.attendance,
      Company: record.company,
      Default: record.typeName,
      Recommendation: recommendationLabel[record.recommendation],
      'Recommendation notes': record.recommendationNotes,
      'Discussion notes': record.discussionNotes.map((note) => `${note.byName || 'Member'} (${new Date(note.at).toLocaleString()}): ${note.text}`).join('\n'),
      'Final decision': record.finalDecision ? finalLabel[record.finalDecision] : '',
      'Final remarks': record.finalRemarks ?? '',
      'Defer until': record.finalUntil === 'next_dc' ? 'Next DC' : record.finalUntil ?? '',
    })));
    const summary = [{
      'Draft status': isDraft ? 'DRAFT — NOT FINALISED' : 'FINAL',
      'DC date': dc.dc.date,
      Status: dc.dc.status,
      'Students called': dc.entries.length,
      Present: dc.entries.filter((entry) => entry.attendance === 'Present').length,
      Absentees: absentees.length,
      'Awaiting final decision': dc.entries.filter((entry) => entry.attendance === 'Present')
        .flatMap((entry) => entry.defaults).filter((record) => record.recommendation !== 'defer' && !record.finalDecision).length,
      'Finalised by': dc.dc.finalisedByName ?? '',
      'Finalised at': dc.dc.finalisedAt ? new Date(dc.dc.finalisedAt).toLocaleString() : '',
    }];
    const absentRows = dc.entries.filter((entry) => entry.attendance === 'Absent').map((entry) => ({
      'Draft status': isDraft ? 'DRAFT — NOT FINALISED' : 'FINAL',
      'DC date': dc.dc.date,
      'Roll number': entry.rollNo,
      Student: entry.studentName,
      Reason: entry.absentReason,
      Defaults: entry.defaults.map((record) => `${record.company} — ${record.typeName}`).join('; '),
    }));
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(results), 'Results');
    XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(summary), 'Summary');
    XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(absentRows), 'Absentees');
    XLSX.writeFile(workbook, `${isDraft ? 'draft-' : ''}dc-${dc.dc.date}.xlsx`);
  }

  return <section className="space-y-5">
    <header className="rounded-2xl border border-slate-200 bg-white p-5 shadow-sm">
      <p className="text-xs font-semibold uppercase tracking-[0.2em] text-indigo-600">Disciplinary Committee</p>
      <div className="mt-2 flex flex-wrap items-end justify-between gap-3">
        <div><h2 className="text-2xl font-bold text-slate-900">After-DC Discussion</h2><p className="mt-1 text-sm text-slate-600">Review recommendations and record final decisions. Defaults change only when the DC is finalised.</p></div>
        {detail && <button type="button" className={`${buttonClass} border border-slate-300 bg-white text-slate-700`} onClick={() => exportWorkbook(detail)}>Export {isHeld ? 'draft' : 'final'} Excel</button>}
      </div>
      {error && <p role="alert" className="mt-4 rounded-lg border border-rose-200 bg-rose-50 p-3 text-sm text-rose-800">{error}</p>}
      {message && <p role="status" className="mt-4 rounded-lg border border-emerald-200 bg-emerald-50 p-3 text-sm text-emerald-800">{message}</p>}
    </header>

    {loading && <p role="status" className="text-sm text-slate-600">Loading DC meetings…</p>}
    {!loading && !dcs.length && <p className="rounded-xl border border-slate-200 bg-white p-5 text-sm text-slate-600">No held or finalised DC meetings yet.</p>}

    {dcs.length > 0 && <div className="grid gap-5 lg:grid-cols-[minmax(16rem,0.32fr)_minmax(0,1fr)]">
      <aside className="space-y-2">
        {dcs.map((dc) => <button key={dc.id} type="button" onClick={() => { setSelectedId(dc.id); setSelectedDefaults([]); }}
          className={`w-full rounded-xl border p-4 text-left ${selectedId === dc.id ? 'border-indigo-400 bg-indigo-50' : 'border-slate-200 bg-white'} ${dc.status === 'Finalised' ? 'opacity-65' : ''}`}>
          <span className="flex items-center justify-between gap-2"><strong>{dc.date}</strong><span className="status-badge">{dc.status}</span></span>
          <span className="mt-2 block text-sm text-slate-600">{dc.status === 'Finalised' ? 'Finalised' : `${dc.counts?.awaitingTeamDecision ?? 0} awaiting final decision`}</span>
        </button>)}
      </aside>

      {detail && <article className={`min-w-0 rounded-2xl border border-slate-200 bg-white p-4 shadow-sm sm:p-6 ${detail.dc.status === 'Finalised' ? 'opacity-70' : ''}`}>
        <div className="flex flex-wrap items-start justify-between gap-3">
          <div><h3 className="text-xl font-bold">{detail.dc.date} {detail.dc.time ? `· ${detail.dc.time}` : ''}</h3><p className="mt-1 text-sm text-slate-600">{detail.entries.filter((entry) => entry.attendance === 'Present').length} present · {absentees.length} absent</p><span className="mt-2 status-badge">{detail.dc.status}</span></div>
          <p className="text-sm text-slate-600" role="status">{completeCount} of {discussionRows.length} decisions recorded{isHeld ? ` · ${undecidedCount} remaining` : ''}</p>
        </div>

        <div className="mt-4 grid gap-3 sm:grid-cols-2">
          <label className="text-sm font-medium">Search<input type="search" value={search} onChange={(event) => setSearch(event.target.value)} placeholder="Student, roll number, company, default" className={fieldClass} /></label>
          <label className="text-sm font-medium">Recommendation filter<select value={recommendationFilter} onChange={(event) => setRecommendationFilter(event.target.value)} className={fieldClass}>
            <option value="all">All recommendations</option><option value="undecided">Undecided</option><option value="remove">To be removed (considered)</option><option value="keep">To be kept (not considered)</option><option value="discuss">Discuss with team</option>
          </select></label>
        </div>

        {isHeld && <div className="mt-4 flex flex-wrap items-center gap-2 rounded-xl bg-slate-50 p-3">
          <span className="mr-auto text-sm text-slate-700">Select Remove/Keep recommendations to accept them as final.</span>
          <button type="button" className={`${buttonClass} border border-slate-300 bg-white text-slate-700`} disabled={busy || !selectedDefaults.length} onClick={() => void acceptSelected()}>Accept selected recommendations</button>
          <button type="button" className={`${buttonClass} bg-indigo-600 text-white`} disabled={busy || undecidedCount > 0} onClick={() => void finalise()}>Finalise DC</button>
        </div>}

        {filteredRows.map(({ entry, record }) => {
          const draft = drafts[record.id] ?? { decision: '', remarks: '', until: '' };
          const eligibleForBulk = ['remove', 'keep'].includes(record.recommendation);
          return <section key={record.id} className="mt-4 rounded-xl border border-slate-200 p-4">
            <div className="flex flex-wrap items-start justify-between gap-3">
              <div><h4 className="font-semibold">{entry.studentName} · {entry.rollNo}</h4><p className="text-sm text-slate-600">{record.company} · {record.typeName} · logged {record.date}</p></div>
              {isHeld && eligibleForBulk && <label className="flex items-center gap-2 text-sm"><input type="checkbox" checked={selectedDefaults.includes(record.id)} onChange={(event) => setSelectedDefaults((current) => event.target.checked ? [...current, record.id] : current.filter((id) => id !== record.id))} />Select</label>}
            </div>
            <p className="mt-3"><strong>Recommendation:</strong> {recommendationLabel[record.recommendation]}</p>
            {(record.recommendationNotes || record.recommendationByName || record.recommendationAt) && <p className="mt-1 whitespace-pre-wrap text-sm text-slate-600">Notes: {record.recommendationNotes || 'No notes'} {record.recommendationByName ? `· ${record.recommendationByName}` : ''}{record.recommendationAt ? ` · ${new Date(record.recommendationAt).toLocaleString()}` : ''}</p>}
            {record.recommendation === 'defer' && <p className="mt-1 text-sm text-slate-600">Deferred until {record.until === 'next_dc' ? 'Next DC' : record.until || 'date not set'}</p>}
            {record.discussionNotes.map((note, index) => <p key={`${record.id}-note-${index}`} className="mt-2 whitespace-pre-wrap rounded bg-slate-50 p-2 text-sm">{note.text}<span className="mt-1 block text-xs text-slate-500">{note.byName || 'Member'} · {new Date(note.at).toLocaleString()}</span></p>)}
            <p className="mt-3 text-sm"><strong>Final decision:</strong> {record.finalDecision ? `${finalLabel[record.finalDecision]}${record.finalByName ? ` · ${record.finalByName}` : ''}${record.finalAt ? ` · ${new Date(record.finalAt).toLocaleString()}` : ''}` : 'Pending'}</p>
            {isHeld && <>
              <div className="mt-3 grid gap-3 md:grid-cols-2">
                <label className="text-sm font-medium">Discussion note<textarea value={notes[record.id] ?? ''} onChange={(event) => setNotes((current) => ({ ...current, [record.id]: event.target.value }))} rows={2} maxLength={2000} className={fieldClass} /></label>
                <div className="flex items-end"><button type="button" className={`${buttonClass} border border-slate-300 bg-white text-slate-700`} disabled={busy || !(notes[record.id] ?? '').trim()} onClick={() => void saveNote(entry, record)}>Add discussion note</button></div>
              </div>
              <div className="mt-3 grid gap-3 md:grid-cols-2">
                <label className="text-sm font-medium">Final decision<select value={draft.decision} onChange={(event) => setDrafts((current) => ({ ...current, [record.id]: { ...draft, decision: event.target.value } }))} className={fieldClass}>
                  <option value="">Choose final decision</option><option value="remove">Remove</option><option value="keep">Keep/lock</option><option value="defer">Defer</option><option value="send_back">Send back to DC</option>
                </select></label>
                <label className="text-sm font-medium">{draft.decision === 'defer' ? 'Defer until *' : 'Final remarks'}{draft.decision === 'defer'
                  ? <select required value={draft.until === 'next_dc' ? 'next_dc' : draft.until ? 'specific' : ''} onChange={(event) => setDrafts((current) => ({ ...current, [record.id]: { ...draft, until: event.target.value === 'next_dc' ? 'next_dc' : '' } }))} className={fieldClass}><option value="">Choose date</option><option value="specific">Specific date</option><option value="next_dc">Next DC</option></select>
                  : <textarea value={draft.remarks} onChange={(event) => setDrafts((current) => ({ ...current, [record.id]: { ...draft, remarks: event.target.value } }))} rows={2} maxLength={2000} className={fieldClass} />}</label>
                {draft.decision === 'defer' && draft.until !== 'next_dc' && <label className="text-sm font-medium">Deferral date *<input required type="date" value={draft.until} onChange={(event) => setDrafts((current) => ({ ...current, [record.id]: { ...draft, until: event.target.value } }))} className={fieldClass} /></label>}
                {draft.decision === 'defer' && <label className="text-sm font-medium">Deferral reason *<textarea required value={draft.remarks} onChange={(event) => setDrafts((current) => ({ ...current, [record.id]: { ...draft, remarks: event.target.value } }))} rows={2} maxLength={2000} className={fieldClass} /></label>}
              </div>
              <button type="button" className={`${buttonClass} mt-3 bg-indigo-600 text-white`} disabled={busy || !draft.decision} onClick={() => void saveFinalDecision(entry, record)}>Save final decision</button>
            </>}
          </section>;
        })}
        {!filteredRows.length && <p className="mt-5 rounded-lg bg-slate-50 p-4 text-sm text-slate-600">No defaults match these filters.</p>}

        {!!deferrals.length && <section className="mt-6 rounded-xl border border-amber-200 bg-amber-50 p-4">
          <h4 className="font-semibold">Deferred at DC — reference only</h4>
          <p className="mt-1 text-sm text-slate-600">These defaults keep their existing deferral and do not require a team final decision.</p>
          {deferrals.map(({ entry, record }) => <p key={record.id} className="mt-2 text-sm">{entry.rollNo} · {entry.studentName} — {record.company} · {record.typeName} · until {record.until === 'next_dc' ? 'Next DC' : record.until || 'unspecified'}</p>)}
        </section>}

        {!!absentees.length && <section className="mt-6 rounded-xl border border-slate-200 bg-slate-50 p-4">
          <h4 className="font-semibold">Absentees</h4>
          {absentees.map((entry) => <p key={entry.rollNo} className="mt-2 text-sm">{entry.rollNo} · {entry.studentName}{entry.absentReason ? ` — ${entry.absentReason}` : ''}</p>)}
        </section>}
      </article>}
    </div>}
  </section>;
}
