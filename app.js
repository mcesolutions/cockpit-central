/* Cockpit Central (vanilla SPA)
 * - Auth: MSAL Browser (Microsoft CDN)
 * - Data: Microsoft Lists (SharePoint List) via Microsoft Graph
 */

const $ = (sel, root=document) => root.querySelector(sel);
const $$ = (sel, root=document) => Array.from(root.querySelectorAll(sel));

const escapeHtml = (s) => String(s ?? "").replace(/[&<>"']/g, (c) => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));

const pad2 = (n) => String(n).padStart(2,'0');
const toISODate = (d) => d ? `${d.getFullYear()}-${pad2(d.getMonth()+1)}-${pad2(d.getDate())}` : "";
const parseISO = (s) => {
  if (!s) return null;
  const d = new Date(s);
  return isNaN(d.getTime()) ? null : d;
};
const fmtDate = (s) => {
  const d = parseISO(s);
  if (!d) return "";
  return `${d.getFullYear()}-${pad2(d.getMonth()+1)}-${pad2(d.getDate())}`;
};
const nowISO = () => new Date().toISOString();

const debounce = (fn, ms=300) => {
  let t;
  return (...args) => {
    clearTimeout(t);
    t = setTimeout(() => fn(...args), ms);
  };
};

const APP = {
  el: null,
  cfg: null,
  msal: null,
  account: null,
  token: null,
  tokenExpiresAt: 0,
  tasks: [],
  loading: true,
  error: null,
  route: { path: '/', params: {} },
  drag: { taskId: null, fromStatus: null },
};

// Microsoft Graph scopes required for Lists (delegated)
const GRAPH_SCOPES = ['User.Read','Sites.ReadWrite.All'];

function msalInteractionKey(){
  try {
    const cid = APP?.cfg?.clientId || window?.COCKPIT_CONFIG?.clientId || '';
    return cid ? `msal.${cid}.interaction.status` : null;
  } catch { return null; }
}

function isInteractionInProgress(){
  const k = msalInteractionKey();
  if (!k) return false;
  try {
    return sessionStorage.getItem(k) === 'interaction_in_progress';
  } catch { return false; }
}

function clearStaleInteractionFlag(){
  const k = msalInteractionKey();
  if (!k) return;
  try {
    const v = sessionStorage.getItem(k);
    if (v !== 'interaction_in_progress') return;

    // If there's no MSAL response in the URL hash, consider the flag stale and clear it.
    const h = String(location.hash || '');
    const looksLikeMsalResponse = h.startsWith('#code=') || h.startsWith('#error=') || h.includes('client_info=') || h.includes('state=');
    if (!looksLikeMsalResponse) {
      sessionStorage.removeItem(k);
    }
  } catch {}
}

function getCfg(){
  const cfg = window.COCKPIT_CONFIG;
  if (!cfg) return null;
  const required = ['tenantId','clientId','redirectUri','siteId','listId'];
  for (const k of required) {
    if (!cfg[k] || String(cfg[k]).startsWith('YOUR_')) return null;
  }
  return cfg;
}

function setAppRoot(){
  APP.el = document.getElementById('app');
}

function setHash(path){
  if (!path.startsWith('#')) path = '#' + path;
  if (location.hash === path) return;
  location.hash = path;
}

function parseRoute(){
  const h = location.hash.replace(/^#/, '') || '/';
  const parts = h.split('?');
  const path = parts[0];
  const seg = path.split('/').filter(Boolean);
  if (seg.length === 0) return { path: '/', params: {} };
  if (seg[0] === 'pole' && seg[1]) return { path: '/pole', params: { pole: seg[1] } };
  if (seg[0] === 'settings') return { path: '/settings', params: {} };
  return { path: '/', params: {} };
}

function toast(msg, type='info'){
  const host = document.body;
  const t = document.createElement('div');
  t.className = `toast toast--${type}`;
  t.innerHTML = `<div class="toast__dot"></div><div class="toast__msg">${escapeHtml(msg)}</div>`;
  host.appendChild(t);
  setTimeout(() => t.classList.add('toast--in'), 10);
  setTimeout(() => { t.classList.remove('toast--in'); t.classList.add('toast--out'); }, 3200);
  setTimeout(() => t.remove(), 3800);
}

async function initAuth(){
  if (!window.msal || !window.msal.PublicClientApplication) {
    throw new Error('MSAL non chargé. Vérifie ta connexion et le script CDN.');
  }
  const cfg = APP.cfg;
  APP.msal = new window.msal.PublicClientApplication({
    auth: {
      clientId: cfg.clientId,
      authority: `https://login.microsoftonline.com/${cfg.tenantId}`,
      redirectUri: cfg.redirectUri,
    },
    cache: {
      cacheLocation: 'localStorage',
      storeAuthStateInCookie: false,
    },
  });

  // Handle redirect login
  await APP.msal.initialize?.();

  // If an interaction was started but never completed (e.g., user canceled or redirectUri mismatch),
  // MSAL can get stuck in interaction_in_progress. Clear stale flags before processing redirect.
  clearStaleInteractionFlag();

  let resp = null;
  try {
    resp = await APP.msal.handleRedirectPromise?.();
  } catch (e) {
    console.warn('MSAL redirect handling failed (non-blocking)', e);
  }
  if (resp && resp.account) APP.account = resp.account;
  const accounts = APP.msal.getAllAccounts();
  if (!APP.account && accounts && accounts.length) APP.account = accounts[0];

  // MSAL (responseMode=fragment) returns auth data in the URL hash (e.g. #code=...).
  // Our router also uses the hash (#/...). After successful redirect handling, restore the prior route.
  const h = String(location.hash || '');
  if (h.startsWith('#code=') || h.startsWith('#error=') || h.includes('client_info=')) {
    const saved = sessionStorage.getItem('cc_post_login_hash');
    if (saved) sessionStorage.removeItem('cc_post_login_hash');
    location.hash = (saved && saved.startsWith('#/')) ? saved : '#/';
  }

  // Keep MSAL's active account aligned with our state
  if (APP.account && APP.msal.setActiveAccount) {
    APP.msal.setActiveAccount(APP.account);
  }
}

async function login(){
  if (isInteractionInProgress()) {
    toast('Connexion en cours…', 'info');
    return;
  }
  // Use redirect flow for maximum reliability (no stuck popup window)
  // Note: after redirect, initAuth() will process the response and set APP.account.
  try {
    // Remember where the user was before redirect.
    sessionStorage.setItem('cc_post_login_hash', location.hash || '#/');
  } catch (_) {}
  await APP.msal.loginRedirect({ scopes: GRAPH_SCOPES });
}

async function logout(){
  try {
    await APP.msal.logoutRedirect({ account: APP.account || undefined });
  } catch (e) {
    console.error(e);
  }
}

async function getToken(){
  const scopes = GRAPH_SCOPES;
  const now = Date.now();

  // If we already have a cached token, reuse it
  if (APP.token && now < APP.tokenExpiresAt - 30_000) return APP.token;

  // Rehydrate account from MSAL cache if needed (common after refresh)
  if (!APP.account && APP.msal?.getAllAccounts) {
    const accounts = APP.msal.getAllAccounts();
    if (accounts && accounts.length) APP.account = accounts[0];
  }

  // Align active account
  if (APP.account && APP.msal?.setActiveAccount) {
    APP.msal.setActiveAccount(APP.account);
  }

  // If still not connected, trigger interactive login once
  if (!APP.account) {
    throw new Error('Non connecté');
  }

  try {
    const resp = await APP.msal.acquireTokenSilent({ scopes, account: APP.account });
    APP.token = resp.accessToken;
    APP.tokenExpiresAt = resp.expiresOn ? resp.expiresOn.getTime() : (now + 45*60*1000);
    return APP.token;
  } catch (e) {
    // Final fallback: redirect (most reliable on static hosting)
    if (isInteractionInProgress()) {
      throw new Error('Connexion en cours. Termine l\'authentification (onglet Microsoft), puis réessaie.');
    }
    await APP.msal.acquireTokenRedirect({ scopes, account: APP.account });
    throw e;
  }
}

async function graphFetch(url, opts={}){
  const token = await getToken();
  const headers = new Headers(opts.headers || {});
  headers.set('Authorization', `Bearer ${token}`);
  headers.set('Accept', 'application/json');
  if (opts.body && !headers.has('Content-Type')) headers.set('Content-Type', 'application/json');
  const res = await fetch(url, { ...opts, headers });
  if (!res.ok) {
    const text = await res.text();
    throw new Error(`${res.status} ${res.statusText}: ${text}`);
  }
  if (res.status === 204) return null;
  return await res.json();
}

const FIELD = {
  Title: 'Title',
  Pole: 'Pole',
  Status: 'Status',
  DueDate: 'DueDate',
  Priority: 'Priority',
  Notes: 'Notes',
  LinkUrl: 'LinkUrl',
  SortOrder: 'SortOrder',
};


// Runtime column discovery (bulletproof): we only write optional fields if they exist in the target list.
APP.columns = null;
APP.columnsNorm = null;
APP.fieldInternal = APP.fieldInternal || {};

async function loadListColumns(){
  if (APP.columns && APP.columnsNorm) return;
  const { siteId, listId } = APP.cfg;
  const url = `https://graph.microsoft.com/v1.0/sites/${encodeURIComponent(siteId)}/lists/${encodeURIComponent(listId)}/columns?$select=name,displayName`;
  try {
    const data = await graphFetch(url, { method: 'GET' });
    const cols = data?.value || [];
    APP.columns = {};
    APP.columnsNorm = {};
    for (const c of cols) {
      if (!c?.name) continue;
      APP.columns[c.name] = c.displayName || c.name;
      // normalize internal name and display name for tolerant matching
      APP.columnsNorm[normKey(c.name)] = c.name;
      if (c.displayName) APP.columnsNorm[normKey(c.displayName)] = c.name;
    }

    // Resolve internal name for Due Date if present (French/English variants)
    const dueCandidates = [
      'Echeance',
      'Échéance',
      "Date d'échéance",
      'Date echeance',
      'Due date',
      'DueDate',
      'Deadline',
    ];
    APP.fieldInternal.DueDate = resolveInternalName(dueCandidates);
  } catch (e) {
    // If columns can't be loaded (permissions / transient), keep the app usable.
    APP.columns = {};
    APP.columnsNorm = {};
    APP.fieldInternal.DueDate = null;
    console.warn('loadListColumns failed:', e);
  }
}

function resolveInternalName(candidates){
  if (!APP.columnsNorm) return null;
  for (const c of candidates) {
    const hit = APP.columnsNorm[normKey(c)];
    if (hit) return hit;
  }
  return null;
}


function extractUnknownFieldName(err){
  const msg = String(err?.message || err || '');
  // Typical Graph error: Field 'DueDate' is not recognized
  const m = msg.match(/Field\s+'([^']+)'\s+is\s+not\s+recognized/i);
  if (m) return m[1];
  // Sometimes the JSON is embedded after a ':'
  const idx = msg.indexOf('{"error"');
  if (idx >= 0) {
    try {
      const j = JSON.parse(msg.slice(idx));
      const message = j?.error?.message;
      const m2 = String(message || '').match(/Field\s+'([^']+)'\s+is\s+not\s+recognized/i);
      if (m2) return m2[1];
    } catch {}
  }
  return null;
}

async function graphFetchWithUnknownFieldRetry(url, options, ctxLabel=''){ 
  try {
    return await graphFetch(url, options);
  } catch (e) {
    const unknown = extractUnknownFieldName(e);
    if (!unknown) throw e;

    // Try stripping the unknown field from payload and retry once.
    try {
      const body = options?.body ? JSON.parse(options.body) : null;
      if (body?.fields && Object.prototype.hasOwnProperty.call(body.fields, unknown)) {
        delete body.fields[unknown];
        toast(`Champ "${unknown}" absent dans la liste : valeur ignorée.`, 'warn');
        return await graphFetch(url, { ...options, body: JSON.stringify(body) });
      }
    } catch {}

    throw e;
  }
}

// Fields in Microsoft Lists can have different internal names depending on how they were created.
// This helper tries the preferred key(s) first, then falls back to a normalized match on any field key.
function pickField(fields, preferredKeys){
  for (const k of preferredKeys) {
    if (fields && fields[k] != null && fields[k] !== '') return fields[k];
  }
  if (!fields) return undefined;
  const want = new Set(preferredKeys.map(normKey));
  for (const k of Object.keys(fields)) {
    if (want.has(normKey(k))) {
      const v = fields[k];
      if (v != null && v !== '') return v;
    }
  }
  return undefined;
}

function normKey(v){
  return String(v ?? '')
    .trim()
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '');
}

function normalizePole(raw){
  const s = normKey(raw);
  if (!s) return '';
  if (s === 'bcs' || s.includes('bien') || s.includes('chez') || s.includes('soi')) return 'BCS';
  if (s === 'evo' || s.includes('evolumis')) return 'EVO';
  if (s === 'perso' || s === 'personnel' || s === 'personal' || s.includes('perso')) return 'PERSO';
  return String(raw).trim();
}

function normalizeStatus(raw){
  const s = normKey(raw);
  if (!s) return 'Backlog';
  if (s === 'backlog' || s === 'todo' || s.includes('a faire') || s.includes('to do')) return 'Backlog';
  if (s === 'encours' || s.includes('en cours') || s.includes('in progress')) return 'EnCours';
  if (s === 'enattente' || s.includes('en attente') || s.includes('waiting') || s.includes('blocked')) return 'EnAttente';
  if (s === 'termine' || s.includes('termine') || s.includes('done') || s.includes('completed')) return 'Termine';
  return String(raw).trim();
}

function normalizePriority(raw){
  const s = normKey(raw);
  if (!s) return 'P2';
  if (s === 'p1' || s === '1' || s.includes('urgent') || s.includes('crit')) return 'P1';
  if (s === 'p2' || s === '2' || s.includes('high')) return 'P2';
  if (s === 'p3' || s === '3' || s.includes('low')) return 'P3';
  return String(raw).trim();
}

function mapFromListItem(item){
  const f = item.fields || {};
  const title = pickField(f, [FIELD.Title, 'Title', 'Titre']);
  const pole = pickField(f, [FIELD.Pole, 'Pole', 'Pôle', 'PoleKey', 'PoleId']);
  const status = pickField(f, [FIELD.Status, 'Status', 'Statut']);
  const dueDate = pickField(f, [FIELD.DueDate, 'DueDate', 'Echeance', 'Échéance', 'Echéance', 'Due', 'Date']);
  const priority = pickField(f, [FIELD.Priority, 'Priority', 'Priorite', 'Priorité']);
  const notes = pickField(f, [FIELD.Notes, 'Notes', 'Note', 'Commentaires', 'Commentaire']);
  const sortOrder = pickField(f, [FIELD.SortOrder, 'SortOrder', 'Order', 'Ordre']);
  const link = pickField(f, [FIELD.LinkUrl, 'LinkUrl', 'Lien', 'URL', 'Url']);
  return {
    id: item.id,
    title: title || '',
    pole: normalizePole(pole || ''),
    status: normalizeStatus(status || 'Backlog'),
    dueDate: dueDate || '',
    priority: normalizePriority(priority || 'P2'),
    notes: notes || '',
    linkUrl: (link && (link.Url || link.url || link)) || '',
    sortOrder: Number(sortOrder ?? 0),
    raw: item,
  };
}

function mapToFields(task){
  const fields = {};
  fields[FIELD.Title] = task.title;
  fields[FIELD.Pole] = task.pole;
  fields[FIELD.Status] = task.status;
  // Due date is optional. Only write if the list has a compatible column.
  if (task.dueDate) {
    const dueInternal = APP.fieldInternal?.DueDate;
    if (dueInternal) {
      fields[dueInternal] = task.dueDate;
    } else {
      // Keep task creation/update functional even if the tenant/list doesn't have a due date column.
      // (No throw; we simply ignore the due date.)
    }
  }
  fields[FIELD.Priority] = task.priority;
  fields[FIELD.Notes] = task.notes || '';
  if (task.linkUrl) {
    fields[FIELD.LinkUrl] = { Url: task.linkUrl, Description: '' };
  } else {
    fields[FIELD.LinkUrl] = null;
  }
  fields[FIELD.SortOrder] = Number(task.sortOrder ?? 0);
  return fields;
}

async function loadTasks(){
  await loadListColumns();
  const { siteId, listId } = APP.cfg;
  // IMPORTANT: $expand must be prefixed with '$' otherwise Graph ignores it.
  const url = `https://graph.microsoft.com/v1.0/sites/${encodeURIComponent(siteId)}/lists/${encodeURIComponent(listId)}/items?$top=500&$expand=fields`;
  const data = await graphFetch(url);
  const items = (data.value || []).map(mapFromListItem);
  // default sort
  items.sort((a,b) => (a.pole.localeCompare(b.pole)) || (a.status.localeCompare(b.status)) || (a.sortOrder-b.sortOrder) || (a.title.localeCompare(b.title)));
  APP.tasks = items;
}

async function updateTaskFields(itemId, partialFields){
  await loadListColumns();
  const { siteId, listId } = APP.cfg;

  // Bulletproof: remap DueDate to the actual internal field name if present; otherwise drop it.
  if (partialFields && Object.prototype.hasOwnProperty.call(partialFields, FIELD.DueDate)) {
    const val = partialFields[FIELD.DueDate];
    delete partialFields[FIELD.DueDate];
    const dueInternal = APP.fieldInternal?.DueDate;
    if (dueInternal) {
      partialFields[dueInternal] = val;
    } else {
      toast('Colonne Échéance absente : date ignorée.', 'warn');
    }
  }

  const url = `https://graph.microsoft.com/v1.0/sites/${encodeURIComponent(siteId)}/lists/${encodeURIComponent(listId)}/items/${encodeURIComponent(itemId)}/fields`;
  await graphFetchWithUnknownFieldRetry(url, { method: 'PATCH', body: JSON.stringify(partialFields) }, 'update');
}

async function createTask(task){
  await loadListColumns();
  if (task?.dueDate && !APP.fieldInternal?.DueDate) {
    toast('Colonne Échéance absente : date ignorée (tâche créée quand même).', 'warn');
  }

  const { siteId, listId } = APP.cfg;
  const url = `https://graph.microsoft.com/v1.0/sites/${encodeURIComponent(siteId)}/lists/${encodeURIComponent(listId)}/items`;
  const body = { fields: mapToFields(task) };
  const created = await graphFetchWithUnknownFieldRetry(url, { method: 'POST', body: JSON.stringify(body) }, 'create');
  const mapped = mapFromListItem(created);
  APP.tasks.push(mapped);
  return mapped;
}

async function deleteTask(itemId){
  const { siteId, listId } = APP.cfg;
  const url = `https://graph.microsoft.com/v1.0/sites/${encodeURIComponent(siteId)}/lists/${encodeURIComponent(listId)}/items/${encodeURIComponent(itemId)}`;
  await graphFetch(url, { method: 'DELETE' });
  APP.tasks = APP.tasks.filter(t => t.id !== itemId);
}

function poleMeta(key){
  return (APP.cfg.poles || []).find(p => p.key === key) || { key, label: key, emoji: '' };
}

function statusMeta(key){
  return (APP.cfg.statuses || []).find(s => s.key === key) || { key, label: key };
}

function priorityMeta(key){
  return (APP.cfg.priorities || []).find(p => p.key === key) || { key, label: key };
}

function kpiForPole(poleKey){
  const tasks = APP.tasks.filter(t => t.pole === poleKey);
  const open = tasks.filter(t => t.status !== 'Termine');
  const inProgress = tasks.filter(t => t.status === 'EnCours');
  const due7 = open.filter(t => {
    if (!t.dueDate) return false;
    const d = parseISO(t.dueDate);
    if (!d) return false;
    const diff = (d.getTime() - Date.now()) / (1000*60*60*24);
    return diff >= -1 && diff <= 7.01;
  });
  return { total: tasks.length, open: open.length, inProgress: inProgress.length, due7: due7.length };
}

function topTasksForPole(poleKey, n=5){
  return APP.tasks
    .filter(t => t.pole === poleKey && t.status !== 'Termine')
    .sort((a,b) => {
      const ad = a.dueDate ? parseISO(a.dueDate)?.getTime() : Infinity;
      const bd = b.dueDate ? parseISO(b.dueDate)?.getTime() : Infinity;
      return (ad - bd) || (a.status.localeCompare(b.status)) || (a.sortOrder - b.sortOrder) || a.title.localeCompare(b.title);
    })
    .slice(0,n);
}

function badgeForStatus(statusKey){
  const s = statusKey;
  if (s === 'EnCours') return `<span class="badge badge--good">En cours</span>`;
  if (s === 'Backlog') return `<span class="badge">Backlog</span>`;
  if (s === 'EnAttente') return `<span class="badge badge--warn">En attente</span>`;
  if (s === 'Termine') return `<span class="badge badge--muted">Terminé</span>`;
  return `<span class="badge">${escapeHtml(s)}</span>`;
}

function badgeForPriority(p){
  if (p === 'P1') return `<span class="badge badge--bad">P1</span>`;
  if (p === 'P2') return `<span class="badge badge--warn">P2</span>`;
  return `<span class="badge">${escapeHtml(p || 'P3')}</span>`;
}

function icon(name){
  const icons = {
    bolt: `<svg class="i" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M13 2L3 14h9l-1 8 10-12h-9l1-8z"/></svg>`,
    link: `<svg class="i" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M10 13a5 5 0 0 1 0-7l1-1a5 5 0 0 1 7 7l-1 1"/><path d="M14 11a5 5 0 0 1 0 7l-1 1a5 5 0 0 1-7-7l1-1"/></svg>`,
    folder: `<svg class="i" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M3 7h5l2 2h11v10a2 2 0 0 1-2 2H3a2 2 0 0 1-2-2V9a2 2 0 0 1 2-2z"/></svg>`,
    calendar: `<svg class="i" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="4" width="18" height="18" rx="2" ry="2"/><line x1="16" y1="2" x2="16" y2="6"/><line x1="8" y1="2" x2="8" y2="6"/><line x1="3" y1="10" x2="21" y2="10"/></svg>`,
    table: `<svg class="i" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M9 3H3v18h6V3z"/><path d="M21 3h-6v18h6V3z"/><path d="M9 9h6"/><path d="M9 15h6"/></svg>`,
    kanban: `<svg class="i" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="4" width="6" height="16" rx="2"/><rect x="10" y="4" width="4" height="10" rx="2"/><rect x="15" y="4" width="6" height="13" rx="2"/></svg>`,
    plus: `<svg class="i" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><line x1="12" y1="5" x2="12" y2="19"/><line x1="5" y1="12" x2="19" y2="12"/></svg>`,
    refresh: `<svg class="i" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M21 12a9 9 0 1 1-3-6.7"/><path d="M21 3v6h-6"/></svg>`,
    settings: `<svg class="i" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M12 15.5a3.5 3.5 0 1 0 0-7 3.5 3.5 0 0 0 0 7z"/><path d="M19.4 15a7.8 7.8 0 0 0 .1-2l2-1.2-2-3.5-2.3.6a7.6 7.6 0 0 0-1.7-1L15 4h-6l-.5 2.9a7.6 7.6 0 0 0-1.7 1L4.5 7.3l-2 3.5L4.5 12a7.8 7.8 0 0 0 .1 2l-2 1.2 2 3.5 2.3-.6a7.6 7.6 0 0 0 1.7 1L9 20h6l.5-2.9a7.6 7.6 0 0 0 1.7-1l2.3.6 2-3.5L19.4 15z"/></svg>`,
    logout: `<svg class="i" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M9 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h4"/><polyline points="16 17 21 12 16 7"/><line x1="21" y1="12" x2="9" y2="12"/></svg>`,
    login: `<svg class="i" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M15 3h4a2 2 0 0 1 2 2v14a2 2 0 0 1-2 2h-4"/><polyline points="10 17 15 12 10 7"/><line x1="15" y1="12" x2="3" y2="12"/></svg>`,
    trash: `<svg class="i" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><polyline points="3 6 5 6 21 6"/><path d="M19 6l-1 14a2 2 0 0 1-2 2H8a2 2 0 0 1-2-2L5 6"/><path d="M10 11v6"/><path d="M14 11v6"/><path d="M9 6V4a1 1 0 0 1 1-1h4a1 1 0 0 1 1 1v2"/></svg>`,
  };
  return icons[name] || '';
}


function chipForStatus(statusKey){
  const s = statusKey;
  if (s === 'EnCours') return `<span class="chip chip--good">En cours</span>`;
  if (s === 'Backlog') return `<span class="chip">Backlog</span>`;
  if (s === 'EnAttente') return `<span class="chip chip--warn">En attente</span>`;
  if (s === 'Termine') return `<span class="chip">Terminé</span>`;
  return `<span class="chip">${escapeHtml(s)}</span>`;
}

function chipForPriority(p){
  if (p === 'P1') return `<span class="chip chip--bad">P1</span>`;
  if (p === 'P2') return `<span class="chip chip--warn">P2</span>`;
  return `<span class="chip">${escapeHtml(p || 'P3')}</span>`;
}

function layout(content){
  const user = APP.account ? (APP.account.name || APP.account.username || '') : '';
  return `
    <div class="shell">
      <div class="topbar">
        <div class="topbar__inner">
          <div class="brand" onclick="window.location.hash='#/'" role="button" tabindex="0">
            <div class="logo"></div>
            <div class="brand__title">
              <b>${escapeHtml(APP.cfg.appName || 'Cockpit Central')}</b>
              <span>Centre de commande • Microsoft Lists</span>
            </div>
          </div>

          <div class="actions">
            <a class="pill" href="#/settings" title="Réglages">${icon('settings')}<span>Réglage</span></a>
            <button class="pill" id="btnRefresh" title="Rafraîchir">${icon('refresh')}<span>Sync</span></button>
            ${APP.account
              ? `<button class="pill pill--primary" id="btnLogout" title="Déconnexion">${icon('logout')}<span>${escapeHtml(user || 'Déconnexion')}</span></button>`
              : `<button class="pill pill--primary" id="btnLogin" title="Connexion">${icon('login')}<span>Connexion</span></button>`
            }
          </div>
        </div>
      </div>

      <div class="content">
        ${content}
      </div>

      <div class="footer">© ${new Date().getFullYear()} Cockpit Central • Zéro blabla, juste du contrôle.</div>
    </div>
  `;
}

function viewHome(){
  const cards = (APP.cfg.poles || []).map(p => {
    const kpi = kpiForPole(p.key);
    const top = topTasksForPole(p.key, 5);
    const appUrl = p.key === 'BCS' ? APP.cfg.links?.bienChezSoiApp : (p.key === 'EVO' ? APP.cfg.links?.evolumisApp : APP.cfg.links?.persoApp);

    const preview = top.length
      ? top.map(t => `
        <div class="item">
          <div class="item__title">
            <b title="${escapeHtml(t.title)}">${escapeHtml(t.title)}</b>
            <span>${t.dueDate ? `Échéance: ${escapeHtml(fmtDate(t.dueDate))}` : 'Sans échéance'}</span>
          </div>
          <div style="display:flex; gap:8px; align-items:center;">
            ${chipForPriority(t.priority)}
            ${chipForStatus(t.status)}
          </div>
        </div>
      `).join('')
      : `<div class="small">Aucune tâche ouverte. Soit tu es à ton top… soit tu procrastines très bien 😄</div>`;

    return `
      <a class="card bigbtn" href="#/pole/${escapeHtml(p.key)}" data-pole="${escapeHtml(p.key)}">
        <div class="card__inner">
          <div class="bigbtn__head">
            <div class="bigbtn__title">
              <b>${escapeHtml(p.label)}</b>
              <span>${escapeHtml(p.key)} • ${kpi.open} ouvertes • ${kpi.due7} cette semaine</span>
            </div>
            <div class="badges">
              <span class="badge">${escapeHtml(p.emoji || '')} Pôle</span>
              ${appUrl ? `<span class="badge">${icon('link')} Ouvrir l'app</span>` : ``}
            </div>
          </div>

          <div class="kpis">
            <div class="kpi"><b>${kpi.inProgress}</b><span>En cours</span></div>
            <div class="kpi"><b>${kpi.open}</b><span>Ouvertes</span></div>
            <div class="kpi"><b>${kpi.total}</b><span>Total</span></div>
          </div>

          <div class="preview">
            <h4>Tâches prioritaires</h4>
            ${preview}
          </div>
        </div>
      </a>
    `;
  }).join('');

  return layout(`
    <div class="hero">
      <div class="card__inner">
        <div class="hero__top">
          <div>
            <h1>Ton cockpit, version 2026.</h1>
            <p>3 gros boutons. Des tâches dynamiques. Kanban / Calendrier / Table. Backend Microsoft Lists, prêt pour la prod.</p>
          </div>
          <div class="badges">
            <span class="badge">Sécurité: Entra ID</span>
            <span class="badge">Data: SharePoint List</span>
            <span class="badge">UI: Premium glass</span>
          </div>
        </div>
      </div>
    </div>

    <div class="grid3">
      ${cards}
    </div>
  `);
}

function viewSettings(){
  const cfg = window.COCKPIT_CONFIG || {};
  const txt = escapeHtml(JSON.stringify(cfg, null, 2));
  return layout(`
    <div class="card">
      <div class="card__inner">
        <div class="row">
          <div>
            <b style="font-size:16px;">Réglages</b>
            <div class="small">Tout se configure dans <span style="font-family:ui-monospace, SFMono-Regular, Menlo, monospace;">config.js</span>. Voici l'état actuel.</div>
          </div>
          <a class="pill" href="#/">Retour</a>
        </div>
        <div class="sep"></div>
        <div class="tablewrap" style="min-height:220px;">
          <pre style="margin:0; padding:14px; white-space:pre;">${txt}</pre>
        </div>
      </div>
    </div>
  `);
}

function viewPole(poleKey){
  const p = poleMeta(poleKey);
  const mode = APP.poleView?.[poleKey] || 'kanban';

  const tasks = APP.tasks
    .filter(t => t.pole === poleKey)
    .sort((a,b) => (a.status.localeCompare(b.status)) || (a.sortOrder - b.sortOrder) || a.title.localeCompare(b.title));

  const setMode = (m) => {
    APP.poleView = APP.poleView || {};
    APP.poleView[poleKey] = m;
    render();
  };

  const header = `
    <div class="card">
      <div class="card__inner">
        <div class="row">
          <div>
            <b style="font-size:16px;">${escapeHtml(p.emoji || '')} ${escapeHtml(p.label)}</b>
            <div class="small">Gestion détaillée • glisser-déposer • édition rapide</div>
          </div>
          <div class="row" style="justify-content:flex-end;">
            ${APP.cfg.links?.sharePointFolderUrl ? `<a class="pill" target="_blank" rel="noopener" href="${escapeHtml(APP.cfg.links.sharePointFolderUrl)}">${icon('folder')}<span>Dossier</span></a>` : ``}
            ${APP.cfg.links?.listWebUrl ? `<a class="pill" target="_blank" rel="noopener" href="${escapeHtml(APP.cfg.links.listWebUrl)}">${icon('link')}<span>Liste</span></a>` : ``}
            <button class="pill pill--primary" id="btnAddTask">${icon('plus')}<span>Nouvelle tâche</span></button>
          </div>
        </div>
        <div class="sep"></div>
        <div class="tabs">
          <button class="tab ${mode==='kanban'?'active':''}" data-tab="kanban">${icon('kanban')} Kanban</button>
          <button class="tab ${mode==='calendar'?'active':''}" data-tab="calendar">${icon('calendar')} Calendrier</button>
          <button class="tab ${mode==='table'?'active':''}" data-tab="table">${icon('table')} Table</button>
          <a class="tab" href="#/">← Accueil</a>
        </div>
      </div>
    </div>
  `;

  const body = mode === 'calendar'
    ? viewPoleCalendar(tasks)
    : (mode === 'table' ? viewPoleTable(tasks) : viewPoleKanban(tasks));

  // After render, wire events
  setTimeout(() => {
    $$('.tab[data-tab]').forEach(btn => btn.addEventListener('click', () => setMode(btn.dataset.tab)));
    const add = $('#btnAddTask');
    if (add) add.addEventListener('click', async () => {
      const title = prompt('Titre de la tâche');
      if (!title) return;
      const task = { title, pole: poleKey, status: 'Backlog', dueDate: null, priority: 'P2', notes: '', linkUrl: '', sortOrder: Date.now() };
      try {
        await createTask(task);
        if (!APP.account) { await login(); }
          await loadTasks();
          toast('Tâche créée ✅', 'good');
        render();
      } catch (e) {
        console.error(e);
        toast('Erreur création (voir console)', 'bad');
      }
    });
  }, 0);

  return layout(header + body);
}

function viewPoleKanban(tasks){
  const statuses = (APP.cfg.statuses || []).map(s => s.key);
  const cols = statuses.map(sk => {
    const meta = statusMeta(sk);
    const items = tasks.filter(t => t.status === sk);
    const cards = items.map(t => `
      <div class="cardtask" draggable="true" data-id="${escapeHtml(t.id)}">
        <b>${escapeHtml(t.title)}</b>
        <div class="meta">
          <span>${t.dueDate ? escapeHtml(fmtDate(t.dueDate)) : '—'}</span>
          <span>${escapeHtml(t.priority || '')}</span>
        </div>
      </div>
    `).join('');
    return `
      <div class="col">
        <div class="col__head">
          <b>${escapeHtml(meta.label)}</b>
          <div class="col__count">${items.length}</div>
        </div>
        <div class="dropzone" data-status="${escapeHtml(sk)}">${cards || ''}</div>
      </div>
    `;
  }).join('');

  // wire DnD after render
  setTimeout(() => {
    $$('.cardtask').forEach(el => {
      el.addEventListener('dragstart', (e) => {
        const id = el.dataset.id;
        e.dataTransfer.setData('text/plain', id);
      });
    });

    $$('.dropzone').forEach(zone => {
      zone.addEventListener('dragover', (e) => { e.preventDefault(); zone.style.borderColor='rgba(255,255,255,.22)'; });
      zone.addEventListener('dragleave', () => { zone.style.borderColor='rgba(255,255,255,.12)'; });
      zone.addEventListener('drop', async (e) => {
        e.preventDefault();
        zone.style.borderColor='rgba(255,255,255,.12)';
        const id = e.dataTransfer.getData('text/plain');
        const newStatus = zone.dataset.status;
        try {
          await updateTaskFields(id, { [FIELD.Status]: newStatus });
          if (!APP.account) { await login(); }
          await loadTasks();
          toast('Mise à jour ✅', 'good');
          render();
        } catch (err) {
          console.error(err);
          toast('Erreur déplacement (voir console)', 'bad');
        }
      });
    });
  }, 0);

  return `
    <div class="card">
      <div class="card__inner">
        <div class="kanban">${cols}</div>
      </div>
    </div>
  `;
}

function viewPoleTable(tasks){
  const rows = tasks.map(t => `
    <tr>
      <td><div class="cell-edit" contenteditable="true" data-id="${escapeHtml(t.id)}" data-field="${FIELD.Title}">${escapeHtml(t.title)}</div></td>
      <td>${escapeHtml(t.status)}</td>
      <td>${escapeHtml(t.priority)}</td>
      <td><div class="cell-edit" contenteditable="true" data-id="${escapeHtml(t.id)}" data-field="${FIELD.DueDate}">${escapeHtml(t.dueDate ? fmtDate(t.dueDate) : '')}</div></td>
      <td><div class="cell-edit" contenteditable="true" data-id="${escapeHtml(t.id)}" data-field="${FIELD.Notes}">${escapeHtml(t.notes || '')}</div></td>
    </tr>
  `).join('');

  // wire inline edits
  setTimeout(() => {
    $$('.cell-edit').forEach(el => {
      const save = debounce(async () => {
        const id = el.dataset.id;
        const field = el.dataset.field;
        let val = el.textContent.trim();
        if (field === FIELD.DueDate) {
          val = val ? val : null;
        }
        try {
          await updateTaskFields(id, { [field]: val });
          toast('Enregistré ✅', 'good');
          await loadTasks();
        } catch (e) {
          console.error(e);
          toast('Erreur enregistrement', 'bad');
        }
      }, 450);
      el.addEventListener('input', save);
    });
  }, 0);

  return `
    <div class="card">
      <div class="card__inner">
        <div class="tablewrap">
          <table>
            <thead><tr><th>Titre</th><th>Statut</th><th>Priorité</th><th>Échéance</th><th>Notes</th></tr></thead>
            <tbody>${rows || ''}</tbody>
          </table>
        </div>
      </div>
    </div>
  `;
}

function viewPoleCalendar(tasks){
  // show current week (Mon..Sun)
  const now = new Date();
  const day = (now.getDay() + 6) % 7; // Monday=0
  const monday = new Date(now);
  monday.setDate(now.getDate() - day);
  const days = Array.from({length:7}, (_,i) => {
    const d = new Date(monday);
    d.setDate(monday.getDate() + i);
    const iso = toISODate(d);
    const items = tasks.filter(t => t.dueDate && fmtDate(t.dueDate) === iso);
    return { d, iso, items };
  });

  const cells = days.map(({d, iso, items}) => {
    const label = `${pad2(d.getDate())}/${pad2(d.getMonth()+1)}`;
    const name = ['Lun','Mar','Mer','Jeu','Ven','Sam','Dim'][(d.getDay()+6)%7];
    const list = items.length
      ? `<div class="day__items">${items.map(t => `<div class="day__task" title="${escapeHtml(t.title)}">${escapeHtml(t.title)}</div>`).join('')}</div>`
      : `<div class="small">—</div>`;
    return `
      <div class="day">
        <div class="day__top"><span>${name}</span><span>${label}</span></div>
        ${list}
      </div>
    `;
  }).join('');

  return `
    <div class="card">
      <div class="card__inner">
        <div class="cal">${cells}</div>
      </div>
    </div>
  `;
}

function renderBootError(msg){
  APP.el.innerHTML = `
    <div class="boot">
      <div class="spinner"></div>
      <div class="boot__title">Cockpit Central — Configuration requise</div>
      <div class="boot__sub">${escapeHtml(msg || '')}</div>
      <div class="row" style="justify-content:center;">
        <a class="pill pill--primary" href="#/settings">Ouvrir /settings</a>
      </div>
    </div>
  `;
}

function render(){
  APP.route = parseRoute();
  let html = '';

  if (APP.route.path === '/settings') {
    html = viewSettings();
  } else if (APP.route.path === '/pole') {
    html = viewPole(APP.route.params.pole);
  } else {
    html = viewHome();
  }

  APP.el.innerHTML = html;

  // bind topbar actions
  const btnLogin = $('#btnLogin');
  const btnLogout = $('#btnLogout');
  const btnRefresh = $('#btnRefresh');

  if (btnLogin) btnLogin.addEventListener('click', () => login());
  if (btnLogout) btnLogout.addEventListener('click', () => logout());
  if (btnRefresh) btnRefresh.addEventListener('click', async () => {
    try {
      if (!APP.account) { await login(); }
      await loadTasks();
      const total = APP.tasks.length;
      const classified = APP.tasks.filter(t => !!t.pole).length;
      if (total === 0) {
        toast('Sync OK, mais 0 tâche trouvée. (Vérifie la liste / accès)', 'warn');
      } else if (classified === 0) {
        toast(`Sync OK (${total}). Aucune tâche classée par pôle — vérifie la colonne "Pole".`, 'warn');
      } else {
        toast(`Synchronisé ✅ (${total})`, 'good');
      }
      render();
    } catch (e) {
      console.error(e);
      toast('Erreur de sync (voir console)', 'bad');
    }
  });
}

async function boot(){
  try {
    setAppRoot();
    APP.cfg = getCfg();
    if (!APP.cfg) {
      renderBootError('config.js est incomplet. Renseigne tenantId, clientId, redirectUri, siteId, listId.');
      return;
    }

    await initAuth();
    if (APP.account) {
      try {
        await loadTasks();
      } catch (e) {
        console.error(e);
        toast('Connecté, mais impossible de charger les tâches (permissions?)', 'warn');
      }
    }

    window.addEventListener('hashchange', render);
    render();
  } catch (e) {
    console.error(e);
    renderBootError(String(e?.message || e));
  }
}

boot();
