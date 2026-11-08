(() => {
  const STORE_KEY = "campusgrid-default-tracker-v1";
  const teams = ["Placement cell", "Summer internship"];
  const today = new Date().toISOString().slice(0, 10);
  // Initial demo data for the tracker
  const initialData = {
    currentMemberId: "m1",
    students: [
      { rollNo: "22CS014", name: "Riya Sharma" },
      { rollNo: "22CS027", name: "Kabir Patel" },
      { rollNo: "22EE006", name: "Ananya Iyer" },
      { rollNo: "22ME019", name: "Dev Malhotra" },
      { rollNo: "22CS041", name: "Ishita Rao" },
      { rollNo: "22EC011", name: "Arjun Nair" }
    ],
    members: [
      { id: "m1", name: "Aarav Mehta", team: teams[0], email: "aarav.mehta@example.edu" },
      { id: "m2", name: "Meera Shah", team: teams[0], email: "meera.shah@example.edu" },
      { id: "m3", name: "Nikhil Das", team: teams[1], email: "nikhil.das@example.edu" },
      { id: "m4", name: "Sara Khan", team: teams[1], email: "sara.khan@example.edu" }
    ],
    types: [
      { id: "t1", name: "Interview no-show", description: "Did not attend a scheduled interview.", appliesTo: "All companies", active: true },
      { id: "t2", name: "Late withdrawal", description: "Withdrew after the application deadline.", appliesTo: "All companies", active: true },
      { id: "t3", name: "Test absence", description: "Absent from a company assessment.", appliesTo: "All companies", active: true },
      { id: "t4", name: "Process rule violation", description: "Did not follow placement process guidelines.", appliesTo: "All companies", active: true }
    ],
    rules: [
      { id: "r1", threshold: 2, action: "warn", label: "Warning" },
      { id: "r2", threshold: 3, action: "restrict", tier: "Dream", label: "Dream restricted" },
      { id: "r3", threshold: 5, action: "debar", label: "Debarred" }
    ],
    defaults: [
      { id: "d1", studentRollNo: "22CS014", memberId: "m2", typeId: "t1", company: "Northstar Labs", date: "2026-09-16", notes: "Missed the 10:00 AM interview.", status: "Logged", dcReference: "" },
      { id: "d2", studentRollNo: "22CS014", memberId: "m1", typeId: "t2", company: "Aster Technologies", date: "2026-09-22", notes: "Withdrew after the shortlist was shared.", status: "Logged", dcReference: "" },
      { id: "d3", studentRollNo: "22CS027", memberId: "m3", typeId: "t3", company: "Vertex Systems", date: "2026-09-18", notes: "Absent from the online assessment.", status: "Locked", dcReference: "DC-2026-09" },
      { id: "d4", studentRollNo: "22EE006", memberId: "m4", typeId: "t4", company: "BluePeak Energy", date: "2026-09-25", notes: "Shared assessment materials.", status: "Logged", dcReference: "" },
      { id: "d5", studentRollNo: "22CS041", memberId: "m2", typeId: "t2", company: "Northstar Labs", date: "2026-09-28", notes: "Late withdrawal after accepting interview.", status: "Removed", dcReference: "DC-2026-10", reason: "Waived at committee review.", removedAt: "2026-10-04T11:00:00.000Z" }
    ],
    studentLocks: [],
    audit: []
  };

  let state = loadState();
  let activeView = "overview";
  let filters = { search: "", type: "", member: "", team: "", status: "" };
  let importRows = [];
  let importKind = "students";
  let importDuplicateRollNos = [];

  function loadState() {
    try {
      const saved = localStorage.getItem(STORE_KEY);
      if (!saved) return structuredClone(initialData);
      const data = JSON.parse(saved);
      for (const key of Object.keys(initialData)) if (!(key in data)) data[key] = structuredClone(initialData[key]);
      return data;
    } catch (error) {
      console.error("Could not load saved tracker data.", error);
      return structuredClone(initialData);
    }
  }
  function persist() {
    try { localStorage.setItem(STORE_KEY, JSON.stringify(state)); }
    catch (error) { notify("Could not save your changes in this browser.", true); console.error(error); }
  }
  function esc(value = "") {
    return String(value).replace(/[&<>"']/g, char => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[char]));
  }
  function initials(name = "") { return name.split(/\s+/).filter(Boolean).slice(0, 2).map(part => part[0]).join("").toUpperCase() || "?"; }
  function member(id) { return state.members.find(item => item.id === id) || { name: "Unknown member", team: "Unknown team" }; }
  function rollNoKey(rollNo) { return String(rollNo || "").trim().toLowerCase(); }
  function student(rollNo) { return state.students.find(item => rollNoKey(item.rollNo) === rollNoKey(rollNo)) || { rollNo: String(rollNo || "").trim(), name: "Unknown student" }; }
  function typeById(id) { return state.types.find(item => item.id === id) || { name: "Archived type", appliesTo: "All companies" }; }
  function isActiveDefault(item) { return item.status !== "Removed"; }
  function studentDefaults(rollNo) { return state.defaults.filter(item => rollNoKey(item.studentRollNo) === rollNoKey(rollNo) && isActiveDefault(item)); }
  function penaltyFor(rollNo) {
    const count = studentDefaults(rollNo).length;
    return [...state.rules].sort((a, b) => b.threshold - a.threshold).find(rule => count >= Number(rule.threshold)) || null;
  }
  function lockFor(rollNo) {
    const lock = state.studentLocks.find(item => rollNoKey(item.rollNo) === rollNoKey(rollNo) && (!item.endDate || item.endDate >= today));
    return lock || null;
  }
  function audit(action, entity, detail, oldValue = "", newValue = "") {
    state.audit.unshift({ id: uid(), action, entity, detail, oldValue, newValue, memberId: state.currentMemberId, at: new Date().toISOString() });
    persist();
  }
  function uid() { return `${Date.now().toString(36)}${Math.random().toString(36).slice(2, 7)}`; }
  function notify(message, error = false) {
    const toast = document.createElement("div");
    toast.className = `toast${error ? " error" : ""}`;
    toast.textContent = message;
    document.getElementById("toast-region").append(toast);
    setTimeout(() => toast.remove(), 3500);
  }
  function dateLabel(value) {
    if (!value) return "—";
    const date = new Date(`${value.slice(0, 10)}T12:00:00`);
    return date.toLocaleDateString("en", { day: "numeric", month: "short", year: "numeric" });
  }
  function ago(value) {
    const delta = Math.max(0, Date.now() - new Date(value).getTime());
    const mins = Math.floor(delta / 60000);
    if (mins < 1) return "Just now";
    if (mins < 60) return `${mins} min ago`;
    const hours = Math.floor(mins / 60);
    if (hours < 24) return `${hours} hr ago`;
    const days = Math.floor(hours / 24);
    return days < 7 ? `${days} day${days === 1 ? "" : "s"} ago` : dateLabel(value);
  }
  function visibleDefaults() {
    const term = filters.search.trim().toLowerCase();
    return state.defaults.filter(item => {
      const s = student(item.studentRollNo), m = member(item.memberId), t = typeById(item.typeId);
      return (!term || [s.name, s.rollNo, item.company, t.name, m.name, item.notes].join(" ").toLowerCase().includes(term))
        && (!filters.type || item.typeId === filters.type)
        && (!filters.member || item.memberId === filters.member)
        && (!filters.team || m.team === filters.team)
        && (!filters.status || item.status === filters.status);
    }).sort((a, b) => b.date.localeCompare(a.date));
  }
  function statusBadge(status) {
    const cls = status.toLowerCase();
    return `<span class="badge badge-${cls}">${esc(status)}</span>`;
  }
  function penaltyBadge(rollNo) {
    const rule = penaltyFor(rollNo);
    if (!rule) return `<span class="badge badge-clear">Clear</span>`;
    const cls = rule.action === "debar" ? "debarred" : rule.action === "restrict" ? "restricted" : "warning";
    return `<span class="badge badge-${cls}">${esc(rule.label)}</span>`;
  }
  function pageTitle(view) { return ({ overview: "Overview", defaults: "Defaults", students: "Students", settings: "Settings", audit: "Audit log" })[view] || "Overview"; }
  function syncShell() {
    const current = member(state.currentMemberId);
    document.getElementById("user-name").textContent = current.name;
    document.getElementById("user-team").textContent = current.team;
    document.getElementById("user-avatar").textContent = initials(current.name);
    document.getElementById("top-avatar").textContent = initials(current.name);
    document.getElementById("nav-default-count").textContent = state.defaults.filter(isActiveDefault).length;
    document.getElementById("page-crumb").textContent = pageTitle(activeView);
  }
  function heading(eyebrow, title, subtitle, actions = "") {
    return `<div class="page-heading"><div><div class="eyebrow">${eyebrow}</div><h1>${title}</h1><p>${subtitle}</p></div><div class="heading-actions">${actions}</div></div>`;
  }
  function notice() {
    return `<div class="notice"><span class="notice-icon">i</span><span><strong>Demo workspace.</strong> Data is saved only in this browser. There is no sign-in, shared database, or access control; use fictional data until a secure, faculty-approved deployment is connected.</span></div>`;
  }
  function statCard(label, value, foot, icon) {
    return `<div class="stat-card"><div class="stat-top"><span>${label}</span><span class="stat-icon">${icon}</span></div><div class="stat-number">${value}</div><div class="stat-foot">${foot}</div></div>`;
  }
  function personCell(s) {
    return `<div class="student-cell"><span class="student-mini">${esc(initials(s.name))}</span><div><div class="cell-main">${esc(s.name)}</div><div class="cell-sub">${esc(s.rollNo)}</div></div></div>`;
  }
  function defaultRows(items, includeActions = false) {
    if (!items.length) return `<tr><td colspan="${includeActions ? 6 : 5}"><div class="empty-state"><strong>No defaults found</strong>Try changing your search or filters.</div></td></tr>`;
    return items.map(item => {
      const s = student(item.studentRollNo), m = member(item.memberId), t = typeById(item.typeId);
      return `<tr><td>${personCell(s)}</td><td><div class="cell-main">${esc(item.company)}</div><div class="cell-sub">${esc(t.name)}</div></td><td>${dateLabel(item.date)}</td><td>${esc(m.name)}<div class="cell-sub">${esc(m.team)}</div></td><td>${statusBadge(item.status)}</td>${includeActions ? `<td><div class="row-actions">${item.status === "Logged" ? `<button class="button button-small" data-action="edit-default" data-id="${item.id}">Edit</button><button class="button button-small" data-action="lock-default" data-id="${item.id}">Lock</button><button class="button button-small button-quiet" data-action="remove-default" data-id="${item.id}">Remove</button>` : item.status === "Locked" ? `<button class="button button-small" data-action="unlock-default" data-id="${item.id}">Unlock</button>` : `<span class="cell-sub">${esc(item.reason || "Waived")}</span>`}</div></td>` : ""}</tr>`;
    }).join("");
  }
  function activityItems(items, limit = 5) {
    if (!items.length) return `<div class="empty-state"><strong>You're all caught up</strong>Updates will appear here as defaults are managed.</div>`;
    return items.slice(0, limit).map(entry => {
      const who = member(entry.memberId).name;
      return `<div class="activity-item"><span class="activity-dot">${entry.action === "Default logged" ? "+" : entry.action.includes("lock") || entry.action.includes("Lock") ? "⌑" : "↻"}</span><div><div class="activity-text"><strong>${esc(who)}</strong> ${esc(entry.detail)}</div><div class="activity-time">${ago(entry.at)}</div></div></div>`;
    }).join("");
  }
  function ruleItems() {
    if (!state.rules.length) return `<div class="empty-state">No penalty rules configured.</div>`;
    return [...state.rules].sort((a, b) => a.threshold - b.threshold).map(rule => `<div class="rule-item"><span class="rule-count">${rule.threshold}×</span><div class="rule-info"><strong>${esc(rule.label)}</strong><small>${rule.action === "restrict" ? `Restrict ${esc(rule.tier || "tier")} companies` : rule.action === "debar" ? "Debar from placements" : "Issue a warning"}</small></div></div>`).join("");
  }
  function renderOverview() {
    const active = state.defaults.filter(isActiveDefault).length;
    const logged = state.defaults.filter(item => item.status === "Logged").length;
    const locked = state.defaults.filter(item => item.status === "Locked").length;
    const affected = state.students.filter(s => penaltyFor(s.rollNo)).length;
    const latest = state.defaults.slice().sort((a, b) => b.date.localeCompare(a.date)).slice(0, 5);
    const actions = `<button class="button" data-action="export">↧ <span>Export</span></button><button class="button button-primary" data-action="new-default">＋ <span>Log a default</span></button>`;
    return `${heading(new Date().toLocaleDateString("en", { weekday: "long", day: "numeric", month: "long", year: "numeric" }).toUpperCase(), `Good morning, ${esc(member(state.currentMemberId).name.split(" ")[0])}`, "A clear view of defaults across your placement season.", actions)}${notice()}
      <div class="stats-grid">${statCard("Active defaults", active, `<span class="positive">${logged} awaiting DC</span> · ${locked} confirmed`, "▤")}${statCard("Students affected", state.students.filter(s => studentDefaults(s.rollNo).length).length, "With at least one active default", "♙")}${statCard("Penalty flags", affected, "Students at a configured threshold", "⚑")}${statCard("Students tracked", state.students.length, `${state.members.length} members in your workspace`, "◉")}</div>
      <div class="content-grid"><section class="panel"><div class="panel-header"><div><h2>Recent defaults</h2><p>The latest student records across your workspace</p></div><button class="panel-link" data-view="defaults">View all →</button></div><div class="table-wrap"><table><thead><tr><th>Student</th><th>Company / type</th><th>Date</th><th>Logged by</th><th>Status</th></tr></thead><tbody>${defaultRows(latest)}</tbody></table></div></section>
      <div class="side-stack"><section class="panel"><div class="panel-header"><div><h2>Penalty rules</h2><p>Current active thresholds</p></div><button class="panel-link" data-view="settings">Edit</button></div><div class="rule-list">${ruleItems()}</div></section><section class="panel"><div class="panel-header"><div><h2>Recent activity</h2><p>Latest changes by your team</p></div><button class="panel-link" data-view="audit">View log</button></div><div class="activity-list">${activityItems(state.audit)}</div></section></div></div>`;
  }
  function renderDefaults() {
    const items = visibleDefaults();
    const actions = `<button class="button" data-action="import" data-kind="defaults">↥ <span>Import Excel</span></button><button class="button" data-action="export">↧ <span>Export Excel</span></button><button class="button button-primary" data-action="new-default">＋ <span>Log a default</span></button>`;
    return `${heading("RECORDS", "Defaults", "Search, review and manage student defaults.", actions)}${notice()}<section class="panel"><div class="filters"><label class="search-box"><span>⌕</span><input class="field" id="filter-search" value="${esc(filters.search)}" placeholder="Search student, company, or keyword…"></label>
      <select class="filter-select" id="filter-type"><option value="">All types</option>${state.types.map(t => `<option value="${t.id}" ${filters.type === t.id ? "selected" : ""}>${esc(t.name)}</option>`).join("")}</select>
      <select class="filter-select" id="filter-member"><option value="">All members</option>${state.members.map(m => `<option value="${m.id}" ${filters.member === m.id ? "selected" : ""}>${esc(m.name)}</option>`).join("")}</select>
      <select class="filter-select" id="filter-team"><option value="">All teams</option>${teams.map(t => `<option ${filters.team === t ? "selected" : ""}>${t}</option>`).join("")}</select>
      <select class="filter-select" id="filter-status"><option value="">All statuses</option>${["Logged", "Locked", "Removed"].map(s => `<option ${filters.status === s ? "selected" : ""}>${s}</option>`).join("")}</select></div>
      <div class="table-wrap"><table><thead><tr><th>Student</th><th>Company / type</th><th>Date</th><th>Logged by</th><th>Status</th><th>Actions</th></tr></thead><tbody id="defaults-body">${defaultRows(items, true)}</tbody></table></div>
      <div class="table-footer"><span id="defaults-count">Showing ${items.length} of ${state.defaults.length} records</span><span>Removed records are retained for audit</span></div></section>`;
  }
  function renderStudents() {
    const query = filters.search.toLowerCase().trim();
    const list = state.students.filter(s => `${s.name} ${s.rollNo}`.toLowerCase().includes(query)).sort((a, b) => a.name.localeCompare(b.name));
    const actions = `<button class="button" data-action="import" data-kind="students">↥ <span>Upload student list</span></button><button class="button button-primary" data-action="student-lock">♙ <span>Lock student</span></button>`;
    return `${heading("DIRECTORY", "Students", "Student names and roll numbers from your uploaded roster.", actions)}${notice()}<div class="filters" style="padding:0 0 13px;border:0"><label class="search-box"><span>⌕</span><input class="field" id="student-search" value="${esc(filters.search)}" placeholder="Search name or roll number…"></label><button class="button" data-action="import" data-kind="students">Upload Excel</button></div>
      ${list.length ? `<div class="student-grid">${list.map(s => { const count = studentDefaults(s.rollNo).length; const lock = lockFor(s.rollNo); return `<article class="student-card"><div class="student-card-head"><span class="student-mini">${esc(initials(s.name))}</span><div><div class="student-name">${esc(s.name)}</div><div class="student-roll">${esc(s.rollNo)}</div></div></div><div class="student-card-stats"><span>${count} active default${count === 1 ? "" : "s"}</span>${penaltyBadge(s.rollNo)}</div>${lock ? `<div class="cell-sub" style="margin-top:9px;color:#bd414d">Manually locked${lock.endDate ? ` until ${dateLabel(lock.endDate)}` : ""} · ${esc(lock.reason)}</div><button class="button button-small" style="margin-top:8px" data-action="unlock-student" data-id="${lock.id}">Unlock student</button>` : `<button class="button button-small" style="margin-top:9px" data-action="student-lock" data-roll="${esc(s.rollNo)}">Lock placement access</button>`}</article>`; }).join("")}</div>` : `<div class="panel empty-state"><strong>No students found</strong>Upload a student roster to get started.</div>`}
      <div class="table-footer" style="padding-left:0"><span>${list.length} of ${state.students.length} students</span><span>Roster re-uploads update names without removing defaults</span></div>`;
  }
  function renderSettings() {
    const types = state.types.map(t => `<div class="setting-row"><div class="setting-detail"><strong>${esc(t.name)}</strong><small>${esc(t.description || "No description")} · ${esc(t.appliesTo)}</small></div><div style="display:flex;align-items:center;gap:10px"><button class="button button-small button-quiet" data-action="edit-type" data-id="${t.id}">Edit</button><button class="toggle ${t.active ? "on" : ""}" data-action="toggle-type" data-id="${t.id}" aria-label="Toggle ${esc(t.name)}"></button></div></div>`).join("");
    const rules = [...state.rules].sort((a, b) => a.threshold - b.threshold).map(rule => `<div class="setting-row"><div class="setting-detail"><strong>${rule.threshold} defaults — ${esc(rule.label)}</strong><small>${rule.action === "restrict" ? `Restrict ${esc(rule.tier || "tier")} companies` : rule.action === "debar" ? "Debar from placements" : "Issue a warning"}</small></div><button class="button button-small button-quiet" data-action="edit-rule" data-id="${rule.id}">Edit</button></div>`).join("");
    const actions = `<button class="button" data-action="import" data-kind="members">↥ <span>Upload members</span></button>`;
    return `${heading("CONFIGURATION", "Settings", "Manage default types, penalty thresholds and your member directory.", actions)}${notice()}<div class="settings-layout"><div><section class="panel settings-panel"><div class="panel-header"><div><h2>Default types</h2><p>Inactive types stay attached to historical records</p></div><button class="button button-small" data-action="new-type">＋ Add type</button></div><div class="settings-list">${types || `<div class="empty-state">No default types yet.</div>`}</div></section>
      <section class="panel settings-panel"><div class="panel-header"><div><h2>Penalty rules</h2><p>Applied to active defaults from the effective count</p></div><button class="button button-small" data-action="new-rule">＋ Add rule</button></div><div class="settings-list">${rules || `<div class="empty-state">No penalty rules yet.</div>`}</div></section></div>
      <div><section class="panel settings-panel"><div class="panel-header"><div><h2>Member directory</h2><p>Everyone has the same access in this prototype</p></div><button class="button button-small" data-action="import" data-kind="members">Upload</button></div><div class="settings-list">${state.members.map(m => `<div class="setting-row"><div class="student-cell"><span class="student-mini">${esc(initials(m.name))}</span><div class="setting-detail"><strong>${esc(m.name)}</strong><small>${esc(m.team)}${m.email ? ` · ${esc(m.email)}` : ""}</small></div></div></div>`).join("") || `<div class="empty-state">No members uploaded.</div>`}</div></section>
      <section class="panel settings-panel"><div class="panel-header"><div><h2>Data management</h2><p>Import existing records or export your current filters</p></div></div><div class="settings-list"><div class="setting-row"><div class="setting-detail"><strong>Import existing defaults</strong><small>Validate an Excel sheet before adding records.</small></div><button class="button button-small" data-action="import" data-kind="defaults">Import</button></div><div class="setting-row"><div class="setting-detail"><strong>Export workbook</strong><small>Summary and detail sheets for committee review.</small></div><button class="button button-small" data-action="export">Export</button></div><div class="setting-row"><div class="setting-detail"><strong>Local demo data</strong><small>Clear this browser's saved tracker data.</small></div><button class="button button-small button-danger" data-action="reset">Reset</button></div></div></section></div></div>`;
  }
  function renderAudit() {
    const entries = state.audit;
    const items = entries.length ? entries.map(entry => `<div class="audit-row"><span class="audit-icon">${entry.action.includes("logged") ? "+" : entry.action.toLowerCase().includes("lock") ? "⌑" : "↻"}</span><div class="audit-copy"><strong>${esc(member(entry.memberId).name)}</strong> ${esc(entry.action.toLowerCase())} ${esc(entry.entity)}${entry.detail ? ` — ${esc(entry.detail)}` : ""}${entry.oldValue || entry.newValue ? `<div class="audit-meta">${entry.oldValue ? `From: ${esc(entry.oldValue)}` : ""}${entry.oldValue && entry.newValue ? " · " : ""}${entry.newValue ? `To: ${esc(entry.newValue)}` : ""}</div>` : ""}<div class="audit-meta">${esc(member(entry.memberId).team)}</div></div><span class="audit-time">${ago(entry.at)}</span></div>`).join("") : `<div class="empty-state"><strong>No changes recorded yet</strong>Actions taken from now on will be recorded here.</div>`;
    return `${heading("TRANSPARENCY", "Audit log", "A traceable history of changes made in this browser.", `<button class="button" data-action="export-audit">↧ <span>Export log</span></button>`)}${notice()}<section class="panel"><div class="panel-header"><div><h2>All activity</h2><p>${entries.length} recorded change${entries.length === 1 ? "" : "s"} · newest first</p></div></div><div class="audit-list">${items}</div></section>`;
  }
  function render() {
    syncShell();
    const page = document.getElementById("page-content");
    page.innerHTML = ({ overview: renderOverview, defaults: renderDefaults, students: renderStudents, settings: renderSettings, audit: renderAudit })[activeView]();
    document.querySelectorAll(".nav-link").forEach(button => button.classList.toggle("active", button.dataset.view === activeView));
  }
  function setView(view) {
    activeView = view;
    filters.search = "";
    document.getElementById("sidebar").classList.remove("open");
    render();
    window.scrollTo(0, 0);
  }
  function modal(title, subtitle, content, wide = false) {
    document.getElementById("modal-root").innerHTML = `<section class="modal${wide ? " modal-wide" : ""}" role="dialog" aria-modal="true" aria-label="${esc(title)}"><div class="modal-header"><div><h2>${esc(title)}</h2><p>${esc(subtitle)}</p></div><button class="modal-close" data-action="close-modal" aria-label="Close">×</button></div>${content}</section>`;
  }
  function closeModal() { document.getElementById("modal-root").innerHTML = ""; }
  function field(label, name, value = "", type = "text", required = false, extra = "") {
    return `<div class="form-field"><label for="form-${name}">${label}${required ? " *" : ""}</label><input class="field" id="form-${name}" name="${name}" type="${type}" value="${esc(value)}" ${required ? "required" : ""} ${extra}></div>`;
  }
  function selectField(label, name, options, selected = "", required = true) {
    return `<div class="form-field"><label for="form-${name}">${label}${required ? " *" : ""}</label><select class="field" id="form-${name}" name="${name}" ${required ? "required" : ""}>${options.map(option => `<option value="${esc(option.value)}" ${selected === option.value ? "selected" : ""}>${esc(option.label)}</option>`).join("")}</select></div>`;
  }
  function formModal(title, subtitle, formId, fields, submitLabel, wide = false) {
    modal(title, subtitle, `<form class="modal-form" id="${formId}"><div class="form-grid">${fields}</div><div class="modal-actions"><button type="button" class="button" data-action="close-modal">Cancel</button><button class="button button-primary" type="submit">${submitLabel}</button></div></form>`, wide);
  }
  function openDefaultForm() {
    if (!state.students.length || !state.members.length || !state.types.some(t => t.active)) {
      notify("Add students, members, and an active default type before logging a record.", true);
      return;
    }
    const fields = `<div class="form-field full"><label for="form-studentRollNo">Roll number *</label><input class="field" id="form-studentRollNo" name="studentRollNo" type="text" required autocomplete="off" autocapitalize="characters" spellcheck="false" placeholder="Enter or paste roll number" aria-describedby="roll-lookup"><span class="form-help" id="roll-lookup" aria-live="polite">Enter a roll number to look up the student.</span></div>`
      + selectField("Logged by", "memberId", state.members.map(m => ({ value: m.id, label: `${m.name} · ${m.team}` })), state.currentMemberId)
      + field("Company", "company", "", "text", true, 'placeholder="Company name"')
      + selectField("Default type", "typeId", state.types.filter(t => t.active).map(t => ({ value: t.id, label: t.name })))
      + field("Date", "date", today, "date", true)
      + `<div class="form-field full"><label for="form-notes">Notes</label><textarea id="form-notes" name="notes" placeholder="Add relevant context for the DC meeting…"></textarea></div>`;
    formModal("Log a default", "Record an incident for committee review.", "default-form", fields, "Save default");
    const submit = document.querySelector("#default-form button[type=submit]");
    submit.disabled = true;
  }
  function openEditDefaultForm(id) {
    const record = state.defaults.find(item => item.id === id);
    if (!record || record.status !== "Logged") return;
    const fields = selectField("Student", "studentRollNo", state.students.map(s => ({ value: s.rollNo, label: `${s.name} · ${s.rollNo}` })), student(record.studentRollNo).rollNo)
      + selectField("Logged by", "memberId", state.members.map(m => ({ value: m.id, label: `${m.name} · ${m.team}` })), record.memberId)
      + field("Company", "company", record.company, "text", true)
      + selectField("Default type", "typeId", state.types.map(t => ({ value: t.id, label: `${t.name}${t.active ? "" : " (inactive)"}` })), record.typeId)
      + field("Date", "date", record.date, "date", true)
      + `<div class="form-field full"><label for="form-notes">Notes</label><textarea id="form-notes" name="notes">${esc(record.notes || "")}</textarea></div>`;
    formModal("Edit default", "Only logged defaults can be edited. Changes are recorded in the audit log.", "edit-default-form", fields, "Save changes");
    document.getElementById("edit-default-form").dataset.id = id;
  }
  function openReasonForm(kind, id) {
    const isRemoval = kind === "remove";
    const item = state.defaults.find(d => d.id === id);
    if (!item) return;
    const fields = `<div class="form-field full"><label for="form-reason">Reason *</label><textarea class="field" id="form-reason" name="reason" required placeholder="${isRemoval ? "Why is this default being waived or reversed?" : "Why is this locked default being unlocked?"}"></textarea></div>`
      + field("DC meeting date or reference", "dcReference", "", "text", true, 'placeholder="e.g. DC-2026-10 or 27 Oct 2026"');
    formModal(isRemoval ? "Remove default" : "Unlock default", isRemoval ? "Removal reduces the count but keeps the record in the audit history." : "Unlocking allows this record to be edited or removed again.", "reason-form", fields, isRemoval ? "Remove default" : "Unlock default");
    document.getElementById("reason-form").dataset.kind = kind;
    document.getElementById("reason-form").dataset.id = id;
  }
  function openStudentLock(rollNo = "") {
    if (!state.students.length) { notify("Upload the student roster first.", true); return; }
    const fields = selectField("Student", "rollNo", state.students.map(s => ({ value: s.rollNo, label: `${s.name} · ${s.rollNo}` })), rollNo || state.students[0].rollNo)
      + field("Reason", "reason", "", "text", true, 'placeholder="Reason for manual placement lock"')
      + field("End date", "endDate", "", "date")
      + `<div class="form-field full"><span class="form-help">Leave the end date blank for an ongoing lock. The student's access flag is independent of count-based penalty rules.</span></div>`;
    formModal("Lock student", "Restrict placement access until the lock expires or is lifted.", "student-lock-form", fields, "Lock placement access");
  }
  function openStudentUnlock(id) {
    const lock = state.studentLocks.find(item => item.id === id);
    if (!lock) return;
    const fields = `<div class="form-field full"><label for="form-reason">Reason for unlocking *</label><textarea class="field" id="form-reason" name="reason" required placeholder="Reason the placement lock is being lifted"></textarea></div>`;
    formModal("Unlock student", `Lift the placement access lock for ${student(lock.rollNo).name}.`, "student-unlock-form", fields, "Unlock student");
    document.getElementById("student-unlock-form").dataset.id = id;
  }
  function openTypeForm(id = "") {
    const existing = state.types.find(t => t.id === id) || { name: "", description: "", appliesTo: "All companies", active: true };
    const fields = field("Type name", "name", existing.name, "text", true)
      + field("Applies to", "appliesTo", existing.appliesTo, "text", true, 'placeholder="All companies, Dream, Core…"')
      + `<div class="form-field full"><label for="form-description">Description</label><textarea id="form-description" name="description">${esc(existing.description)}</textarea></div>`
      + `<div class="form-field full"><label class="setting-row" style="padding:0;border:0;justify-content:flex-start;gap:8px"><input type="checkbox" name="active" ${existing.active ? "checked" : ""}> Active and available for new records</label></div>`;
    formModal(id ? "Edit default type" : "Add default type", "Type changes are recorded in the audit log.", "type-form", fields, id ? "Save changes" : "Add type");
    document.getElementById("type-form").dataset.id = id;
  }
  function openRuleForm(id = "") {
    const existing = state.rules.find(r => r.id === id) || { threshold: "", action: "warn", tier: "Dream", label: "" };
    const fields = field("Threshold (number of defaults)", "threshold", existing.threshold, "number", true, 'min="1" step="1"')
      + selectField("Action", "action", [{ value: "warn", label: "Warn" }, { value: "restrict", label: "Restrict tier" }, { value: "debar", label: "Debar" }], existing.action)
      + field("Tier (for restrict action)", "tier", existing.tier || "Dream")
      + field("Student label", "label", existing.label, "text", true, 'placeholder="Shown next to the student"');
    formModal(id ? "Edit penalty rule" : "Add penalty rule", "New rules apply to counts from the time they are saved.", "rule-form", fields, id ? "Save rule" : "Add rule");
    document.getElementById("rule-form").dataset.id = id;
  }
  function openImport(kind) {
    importKind = kind;
    importRows = [];
    const label = kind === "students" ? "student roster" : kind === "members" ? "member directory" : "existing defaults";
    const hint = kind === "students" ? "Required columns: roll number and name." : kind === "members" ? "Required columns: name and team (Placement cell or Summer internship). Optional: email." : "Required columns: roll number, company, type, and date. Member may be matched by name; unmatched types will be flagged.";
    modal(`Upload ${label}`, "Choose an Excel workbook (.xlsx or .xls) or CSV file.", `<div class="modal-form"><div class="file-drop"><strong>Choose a spreadsheet to preview</strong><span>${hint}</span><input type="file" id="import-file" accept=".xlsx,.xls,.csv"></div><div id="import-preview"></div><div class="modal-actions"><button class="button" data-action="close-modal">Cancel</button><button class="button button-primary" id="confirm-import" disabled>Import rows</button></div></div>`, true);
    document.getElementById("import-file").addEventListener("change", handleImportFile);
  }
  function normalizeKey(key) { return String(key || "").toLowerCase().replace(/[^a-z0-9]/g, ""); }
  function readColumn(row, aliases) {
    const lookup = new Map(Object.entries(row).map(([key, value]) => [normalizeKey(key), value]));
    for (const alias of aliases) {
      const value = lookup.get(normalizeKey(alias));
      if (value !== undefined && value !== null && String(value).trim()) return String(value).trim();
    }
    return "";
  }
  function normalizeDate(value) {
    if (typeof value === "number" && window.XLSX?.SSF) {
      const parsed = window.XLSX.SSF.parse_date_code(value);
      if (parsed) return `${parsed.y}-${String(parsed.m).padStart(2, "0")}-${String(parsed.d).padStart(2, "0")}`;
    }
    const text = String(value || "").trim();
    if (/^\d{4}-\d{2}-\d{2}/.test(text)) return text.slice(0, 10);
    const date = new Date(text);
    return Number.isNaN(date.getTime()) ? "" : date.toISOString().slice(0, 10);
  }
  async function handleImportFile(event) {
    const file = event.target.files[0];
    if (!file) return;
    try {
      if (!window.XLSX) throw new Error("The spreadsheet reader is unavailable. Check your internet connection and try again.");
      const data = await file.arrayBuffer();
      const workbook = window.XLSX.read(data, { type: "array", cellDates: true });
      const sheet = workbook.Sheets[workbook.SheetNames[0]];
      const source = window.XLSX.utils.sheet_to_json(sheet, { defval: "", raw: false });
      importRows = source.map((row, index) => normalizeImportRow(row, index + 2)).filter(row => row);
      importDuplicateRollNos = [];
      if (importKind === "students") {
        const rowsByRollNo = new Map();
        for (const row of importRows) {
          const key = rollNoKey(row.rollNo);
          if (!key) continue;
          if (!rowsByRollNo.has(key)) rowsByRollNo.set(key, []);
          rowsByRollNo.get(key).push(row);
        }
        for (const duplicateRows of rowsByRollNo.values()) {
          if (duplicateRows.length < 2) continue;
          const displayRollNo = duplicateRows[0].rollNo;
          importDuplicateRollNos.push(displayRollNo);
          duplicateRows.forEach(row => { row.error = `Duplicate roll number: ${displayRollNo}`; });
        }
      }
      const invalid = importRows.filter(row => row.error);
      const duplicateMessage = importDuplicateRollNos.length ? `Upload rejected. Duplicate roll number${importDuplicateRollNos.length === 1 ? "" : "s"}: ${importDuplicateRollNos.join(", ")}.` : "";
      document.getElementById("import-preview").innerHTML = `${duplicateMessage ? `<div class="form-help" role="alert" style="margin-top:11px;color:#bd414d">${esc(duplicateMessage)}</div>` : ""}<div class="form-help" style="margin-top:11px">${source.length} rows found · ${importRows.length - invalid.length} valid · ${invalid.length} need attention. Invalid rows will not be imported.</div><div class="mini-table"><table><thead><tr><th>Row</th><th>Preview</th><th>Validation</th></tr></thead><tbody>${importRows.slice(0, 12).map(row => `<tr><td>${row.row}</td><td>${esc(row.preview)}</td><td>${row.error ? `<span style="color:#bd414d">${esc(row.error)}</span>` : "Ready"}</td></tr>`).join("")}</tbody></table></div>${source.length > 12 ? `<div class="form-help" style="margin-top:6px">Showing first 12 rows.</div>` : ""}`;
      const button = document.getElementById("confirm-import");
      button.disabled = importDuplicateRollNos.length > 0 || !importRows.some(row => !row.error);
      button.textContent = `Import ${importRows.filter(row => !row.error).length} rows`;
      if (duplicateMessage) notify(duplicateMessage, true);
    } catch (error) {
      notify(error.message || "Could not read the spreadsheet.", true);
    }
  }
  function normalizeImportRow(row, rowNumber) {
    const result = { row: rowNumber };
    if (importKind === "students") {
      result.rollNo = readColumn(row, ["roll number", "roll no", "roll_no", "rollno", "student id"]);
      result.name = readColumn(row, ["name", "student name", "full name"]);
      result.preview = `${result.name || "—"} · ${result.rollNo || "—"}`;
      if (!result.rollNo || !result.name) result.error = "Name and roll number required";
    } else if (importKind === "members") {
      result.name = readColumn(row, ["name", "member name", "full name"]);
      result.team = readColumn(row, ["team", "group"]);
      result.email = readColumn(row, ["email", "college email", "email address"]);
      if (result.team.toLowerCase().includes("placement")) result.team = teams[0];
      if (result.team.toLowerCase().includes("summer") || result.team.toLowerCase().includes("intern")) result.team = teams[1];
      result.preview = `${result.name || "—"} · ${result.team || "—"}`;
      if (!result.name || !teams.includes(result.team)) result.error = "Name and valid team required";
    } else {
      result.rollNo = readColumn(row, ["roll number", "roll no", "roll_no", "rollno", "student roll number"]);
      result.company = readColumn(row, ["company", "company name"]);
      result.typeName = readColumn(row, ["type", "default type", "default"]);
      result.date = normalizeDate(readColumn(row, ["date", "default date", "incident date"]));
      result.memberName = readColumn(row, ["member", "logged by", "member name"]);
      result.team = readColumn(row, ["team"]);
      result.notes = readColumn(row, ["notes", "details"]);
      result.preview = `${result.rollNo || "—"} · ${result.company || "—"} · ${result.date || "—"}`;
      const missing = [];
      if (!result.rollNo) missing.push("roll number");
      if (!result.company) missing.push("company");
      if (!result.typeName) missing.push("type");
      if (!result.date) missing.push("valid date");
      if (missing.length) result.error = `Missing ${missing.join(", ")}`;
      else if (!state.students.some(s => rollNoKey(s.rollNo) === rollNoKey(result.rollNo))) result.error = "Student not in roster";
      else if (!state.types.some(t => t.name.toLowerCase() === result.typeName.toLowerCase())) result.error = "Unknown default type";
      else if (result.memberName && !state.members.some(m => m.name.toLowerCase() === result.memberName.toLowerCase())) result.error = "Unknown member";
    }
    return result;
  }
  function commitImport() {
    if (importKind === "students" && importDuplicateRollNos.length) {
      notify(`Upload rejected. Duplicate roll number${importDuplicateRollNos.length === 1 ? "" : "s"}: ${importDuplicateRollNos.join(", ")}.`, true);
      return;
    }
    const valid = importRows.filter(row => !row.error);
    if (!valid.length) return;
    if (importKind === "students") {
      const merged = new Map(state.students.map(item => [rollNoKey(item.rollNo), item]));
      valid.forEach(row => {
        const key = rollNoKey(row.rollNo);
        const existing = merged.get(key);
        if (existing) Object.assign(existing, { rollNo: row.rollNo, name: row.name });
        else merged.set(key, { rollNo: row.rollNo, name: row.name });
      });
      state.students = [...merged.values()];
      audit("Students imported", "roster", `${valid.length} roster rows processed`, "", `${state.students.length} students`);
    } else if (importKind === "members") {
      const merged = new Map(state.members.map(item => [item.name.toLowerCase(), item]));
      valid.forEach(row => {
        const existing = merged.get(row.name.toLowerCase());
        merged.set(row.name.toLowerCase(), { id: existing?.id || uid(), name: row.name, team: row.team, email: row.email || existing?.email || "" });
      });
      state.members = [...merged.values()];
      if (!state.members.some(item => item.id === state.currentMemberId)) state.currentMemberId = state.members[0]?.id || "";
      audit("Members imported", "directory", `${valid.length} member rows processed`, "", `${state.members.length} members`);
    } else {
      let added = 0;
      let skipped = 0;
      valid.forEach(row => {
        const type = state.types.find(item => item.name.toLowerCase() === row.typeName.toLowerCase());
        const matched = row.memberName ? state.members.find(item => item.name.toLowerCase() === row.memberName.toLowerCase()) : null;
        const studentRecord = state.students.find(s => rollNoKey(s.rollNo) === rollNoKey(row.rollNo));
        const duplicate = state.defaults.some(item => rollNoKey(item.studentRollNo) === rollNoKey(studentRecord.rollNo)
          && item.company.toLowerCase() === row.company.toLowerCase() && item.typeId === type.id && item.date === row.date);
        if (duplicate) { skipped += 1; return; }
        const entry = { id: uid(), studentRollNo: studentRecord.rollNo, memberId: matched?.id || state.currentMemberId, typeId: type.id, company: row.company, date: row.date, notes: row.notes, status: "Locked", dcReference: "Imported" };
        state.defaults.push(entry);
        added += 1;
      });
      audit("Defaults imported", "records", `${added} validated records imported as locked${skipped ? `; ${skipped} duplicates skipped` : ""}`, "", `${added} records`);
      valid.importedCount = added;
      valid.skippedCount = skipped;
    }
    persist();
    closeModal();
    render();
    const imported = valid.importedCount ?? valid.length;
    const skipped = valid.skippedCount ?? 0;
    notify(`Imported ${imported} row${imported === 1 ? "" : "s"}${skipped ? `; skipped ${skipped} duplicate${skipped === 1 ? "" : "s"}` : ""}.`);
  }
  function exportWorkbook() {
    if (!window.XLSX) { notify("Excel export is unavailable. Check your internet connection and try again.", true); return; }
    const items = visibleDefaults();
    const hasFilters = Object.values(filters).some(Boolean);
    const filteredRolls = new Set(items.map(item => rollNoKey(item.studentRollNo)));
    if (filters.search) {
      state.students.filter(s => `${s.name} ${s.rollNo}`.toLowerCase().includes(filters.search.toLowerCase())).forEach(s => filteredRolls.add(rollNoKey(s.rollNo)));
    }
    const summary = state.students.filter(s => !hasFilters || filteredRolls.has(rollNoKey(s.rollNo))).map(s => {
      const count = items.filter(item => rollNoKey(item.studentRollNo) === rollNoKey(s.rollNo) && isActiveDefault(item)).length;
      const lock = lockFor(s.rollNo);
      return { "Roll number": s.rollNo, Name: s.name, "Number of defaults": count, "Lock / penalty status": lock ? `Manually locked: ${lock.reason}` : penaltyFor(s.rollNo)?.label || "Clear" };
    });
    const details = items.map(item => {
      const s = student(item.studentRollNo), m = member(item.memberId), t = typeById(item.typeId);
      return { "Roll number": s.rollNo, Student: s.name, Company: item.company, Type: t.name, Date: item.date, "Logged by": m.name, Team: m.team, Status: item.status, "DC reference": item.dcReference || "", Reason: item.reason || "", Notes: item.notes || "" };
    });
    const workbook = window.XLSX.utils.book_new();
    window.XLSX.utils.book_append_sheet(workbook, window.XLSX.utils.json_to_sheet(summary), "Summary");
    window.XLSX.utils.book_append_sheet(workbook, window.XLSX.utils.json_to_sheet(details), "Details");
    window.XLSX.writeFile(workbook, `default-tracker-${today}.xlsx`);
    notify("Workbook exported with Summary and Details sheets.");
  }
  function exportAudit() {
    if (!window.XLSX) { notify("Excel export is unavailable. Check your internet connection and try again.", true); return; }
    const rows = state.audit.map(item => ({ "Date and time": item.at, Member: member(item.memberId).name, Team: member(item.memberId).team, Action: item.action, Entity: item.entity, Detail: item.detail, "Old value": item.oldValue, "New value": item.newValue }));
    const workbook = window.XLSX.utils.book_new();
    window.XLSX.utils.book_append_sheet(workbook, window.XLSX.utils.json_to_sheet(rows), "Audit log");
    window.XLSX.writeFile(workbook, `default-tracker-audit-${today}.xlsx`);
  }
  function onSubmit(event) {
    const form = event.target;
    if (!form.matches("form")) return;
    event.preventDefault();
    const data = Object.fromEntries(new FormData(form));
    if (form.id === "default-form") {
      const studentRecord = student(data.studentRollNo);
      if (studentRecord.name === "Unknown student" || !data.studentRollNo.trim()) {
        const input = form.elements.studentRollNo;
        input.setCustomValidity("Roll number not found in the dataset");
        input.reportValidity();
        document.getElementById("roll-lookup").textContent = "Roll number not found in the dataset";
        document.getElementById("roll-lookup").className = "form-help lookup-error";
        return;
      }
      const record = { id: uid(), studentRollNo: studentRecord.rollNo, memberId: data.memberId, typeId: data.typeId, company: data.company.trim(), date: data.date, notes: data.notes.trim(), status: "Logged", dcReference: "" };
      state.defaults.unshift(record);
      audit("Default logged", "default", `logged ${typeById(record.typeId).name} for ${student(record.studentRollNo).name} at ${record.company}`, "", "Logged");
      closeModal(); render(); notify("Default logged and added to the audit trail.");
    } else if (form.id === "edit-default-form") {
      const record = state.defaults.find(item => item.id === form.dataset.id);
      if (!record || record.status !== "Logged") return;
      const before = { student: student(record.studentRollNo).name, rollNo: record.studentRollNo, company: record.company, type: typeById(record.typeId).name, date: record.date, member: member(record.memberId).name, notes: record.notes || "" };
      Object.assign(record, { studentRollNo: data.studentRollNo, memberId: data.memberId, typeId: data.typeId, company: data.company.trim(), date: data.date, notes: data.notes.trim() });
      const after = { student: student(record.studentRollNo).name, rollNo: record.studentRollNo, company: record.company, type: typeById(record.typeId).name, date: record.date, member: member(record.memberId).name, notes: record.notes };
      audit("Default updated", "default", `${after.student} · ${after.company}`, JSON.stringify(before), JSON.stringify(after));
      persist(); closeModal(); render(); notify("Default updated and recorded in the audit trail.");
    } else if (form.id === "reason-form") {
      const item = state.defaults.find(d => d.id === form.dataset.id);
      const previous = item.status;
      if (form.dataset.kind === "remove") {
        item.status = "Removed"; item.reason = data.reason.trim(); item.dcReference = data.dcReference.trim(); item.removedAt = new Date().toISOString();
        audit("Default removed", "default", `${student(item.studentRollNo).name} · ${typeById(item.typeId).name}; ${item.dcReference}`, previous, `Removed: ${item.reason}`);
      } else {
        item.status = "Logged";
        audit("Default unlocked", "default", `${student(item.studentRollNo).name} · ${data.dcReference}; ${data.reason.trim()}`, previous, "Logged");
      }
      persist(); closeModal(); render(); notify(form.dataset.kind === "remove" ? "Default removed; counts and labels have been recalculated." : "Default unlocked.");
    } else if (form.id === "student-lock-form") {
      if (lockFor(data.rollNo)) { notify("This student already has an active placement lock.", true); return; }
      const entry = { id: uid(), rollNo: data.rollNo, reason: data.reason.trim(), endDate: data.endDate || "" };
      state.studentLocks.push(entry);
      audit("Student locked", "student", `${student(data.rollNo).name}; ${entry.reason}${entry.endDate ? ` until ${entry.endDate}` : ""}`, "", "Placement access locked");
      persist(); closeModal(); render(); notify("Student placement access locked.");
    } else if (form.id === "student-unlock-form") {
      const lock = state.studentLocks.find(item => item.id === form.dataset.id);
      if (!lock) return;
      state.studentLocks = state.studentLocks.filter(item => item.id !== lock.id);
      audit("Student unlocked", "student", `${student(lock.rollNo).name}; ${data.reason.trim()}`, "Placement access locked", "Unlocked");
      persist(); closeModal(); render(); notify("Student placement access unlocked.");
    } else if (form.id === "type-form") {
      const id = form.dataset.id;
      const existing = state.types.find(item => item.id === id);
      const type = { id: id || uid(), name: data.name.trim(), description: data.description.trim(), appliesTo: data.appliesTo.trim(), active: new FormData(form).has("active") };
      const oldValue = existing ? JSON.stringify({ name: existing.name, description: existing.description, appliesTo: existing.appliesTo, active: existing.active }) : "";
      if (existing) Object.assign(existing, type); else state.types.push(type);
      audit(existing ? "Default type updated" : "Default type created", "settings", type.name, oldValue, JSON.stringify({ name: type.name, description: type.description, appliesTo: type.appliesTo, active: type.active }));
      persist(); closeModal(); render(); notify("Default type saved.");
    } else if (form.id === "rule-form") {
      const id = form.dataset.id;
      const existing = state.rules.find(item => item.id === id);
      const rule = { id: id || uid(), threshold: Number(data.threshold), action: data.action, tier: data.tier.trim() || "Dream", label: data.label.trim() };
      const oldValue = existing ? JSON.stringify({ threshold: existing.threshold, action: existing.action, tier: existing.tier, label: existing.label }) : "";
      if (existing) Object.assign(existing, rule); else state.rules.push(rule);
      state.rules.sort((a, b) => a.threshold - b.threshold);
      audit(existing ? "Penalty rule updated" : "Penalty rule created", "settings", `${rule.threshold} defaults · ${rule.label}`, oldValue, JSON.stringify({ threshold: rule.threshold, action: rule.action, tier: rule.tier, label: rule.label }));
      persist(); closeModal(); render(); notify("Penalty rule saved. Student labels now reflect the active rules.");
    }
  }
  function handleAction(button) {
    const { action, id, kind, roll } = button.dataset;
    if (action === "close-modal") closeModal();
    else if (action === "new-default") openDefaultForm();
    else if (action === "lock-default") {
      const item = state.defaults.find(d => d.id === id);
      if (!item || item.status !== "Logged") return;
      item.status = "Locked";
      audit("Default locked", "default", `${student(item.studentRollNo).name} · ${typeById(item.typeId).name}`, "Logged", "Locked");
      persist(); render(); notify("Default locked after committee review.");
    } else if (action === "edit-default") openEditDefaultForm(id);
    else if (action === "unlock-default") openReasonForm("unlock", id);
    else if (action === "remove-default") openReasonForm("remove", id);
    else if (action === "student-lock") openStudentLock(roll);
    else if (action === "unlock-student") {
      openStudentUnlock(id);
    } else if (action === "new-type") openTypeForm();
    else if (action === "edit-type") openTypeForm(id);
    else if (action === "toggle-type") {
      const type = state.types.find(t => t.id === id);
      if (!type) return;
      const old = type.active ? "Active" : "Inactive";
      type.active = !type.active;
      audit("Default type updated", "settings", type.name, old, type.active ? "Active" : "Inactive");
      persist(); render(); notify(`${type.name} is now ${type.active ? "active" : "inactive"}.`);
    } else if (action === "new-rule") openRuleForm();
    else if (action === "edit-rule") openRuleForm(id);
    else if (action === "import") openImport(kind);
    else if (action === "export") exportWorkbook();
    else if (action === "export-audit") exportAudit();
    else if (action === "reset") {
      if (confirm("Clear all tracker data saved in this browser and restore the fictional demo data?")) {
        state = structuredClone(initialData); persist(); render(); notify("Demo data restored.");
      }
    } else if (action === "switch-user") {
      if (!state.members.length) { notify("Upload the member list first.", true); return; }
      const currentIndex = state.members.findIndex(m => m.id === state.currentMemberId);
      state.currentMemberId = state.members[(currentIndex + 1) % state.members.length].id;
      persist(); render(); notify(`Demo member switched to ${member(state.currentMemberId).name}.`);
    }
  }
  document.addEventListener("click", event => {
    const view = event.target.closest("[data-view]");
    if (view) { event.preventDefault(); setView(view.dataset.view); return; }
    const action = event.target.closest("[data-action]");
    if (action) handleAction(action);
  });
  document.addEventListener("submit", onSubmit);
  document.getElementById("mobile-menu").addEventListener("click", () => document.getElementById("sidebar").classList.toggle("open"));
  document.getElementById("switch-user").addEventListener("click", () => handleAction({ dataset: { action: "switch-user" } }));
  document.getElementById("top-avatar").addEventListener("click", () => handleAction({ dataset: { action: "switch-user" } }));
  document.getElementById("modal-root").addEventListener("click", event => { if (event.target === event.currentTarget) closeModal(); });
  document.getElementById("page-content").addEventListener("input", event => {
    if (event.target.id === "filter-search" || event.target.id === "student-search") {
      filters.search = event.target.value;
      if (activeView === "defaults") {
        const rows = visibleDefaults();
        document.getElementById("defaults-body").innerHTML = defaultRows(rows, true);
        document.getElementById("defaults-count").textContent = `Showing ${rows.length} of ${state.defaults.length} records`;
      } else render();
    }
  });
  document.getElementById("page-content").addEventListener("change", event => {
    if (event.target.id?.startsWith("filter-")) {
      filters[event.target.id.slice(7)] = event.target.value;
      const rows = visibleDefaults();
      document.getElementById("defaults-body").innerHTML = defaultRows(rows, true);
      document.getElementById("defaults-count").textContent = `Showing ${rows.length} of ${state.defaults.length} records`;
    }
  });
  document.getElementById("modal-root").addEventListener("click", event => {
    if (event.target.id === "confirm-import") commitImport();
  });
  document.getElementById("modal-root").addEventListener("input", event => {
    if (event.target.id !== "form-studentRollNo") return;
    const input = event.target;
    const normalized = input.value.trim();
    const found = normalized ? state.students.find(item => rollNoKey(item.rollNo) === rollNoKey(normalized)) : null;
    const lookup = document.getElementById("roll-lookup");
    const submit = document.querySelector("#default-form button[type=submit]");
    input.setCustomValidity(normalized && !found ? "Roll number not found in the dataset" : "");
    input.removeAttribute("aria-invalid");
    if (!normalized) {
      lookup.textContent = "Enter a roll number to look up the student.";
      lookup.className = "form-help";
      input.removeAttribute("aria-invalid");
    } else if (found) {
      lookup.textContent = found.name;
      lookup.className = "form-help lookup-success";
    } else {
      lookup.textContent = "Roll number not found in the dataset";
      lookup.className = "form-help lookup-error";
      input.setAttribute("aria-invalid", "true");
    }
    if (submit) submit.disabled = !found;
  });
  document.addEventListener("keydown", event => { if (event.key === "Escape") closeModal(); });
  if ("serviceWorker" in navigator && location.protocol.startsWith("http")) {
    navigator.serviceWorker.register("./service-worker.js").catch(error => console.error("Could not register the offline app shell.", error));
  }
  render();
})();
