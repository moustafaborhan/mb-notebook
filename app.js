"use strict";
/* ============================================================
   MB Notebook — Version 2.0
   Google sign-in only. Your notebook IS your Google Drive folder:
     MBorhan-Cloud_Word / <Section folder> / <Page = Google Doc>
   Edits save to the Google Doc a couple of seconds after you stop
   typing; other devices pick changes up within ~15 seconds.
   ============================================================ */
const APP_VERSION = "2.0.2";
const CONFIG = window.MBN_CONFIG || {};
const ROOT_ID = CONFIG.rootId || "12VpHmTOvgahjI8jEMMMC2NBYjS2Os5dO";
const ROOT_NAME = CONFIG.rootName || "MBorhan-Cloud_Word";
const SCOPE = "https://www.googleapis.com/auth/drive";
const API = "https://www.googleapis.com/drive/v3";
const UPLOAD = "https://www.googleapis.com/upload/drive/v3";
const DOC = "application/vnd.google-apps.document";
const FOLDER = "application/vnd.google-apps.folder";
const DOCX = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
const POLL_MS = window.__MBN_POLL || 15000;
const CDN = "https://cdn.jsdelivr.net/npm/";
const LIB = {
  tiny: CDN + "tinymce@7.9.1/",
  marked: CDN + "marked@12.0.2/marked.min.js",
  turndown: CDN + "turndown@7.2.0/dist/turndown.js",
  gis: "https://accounts.google.com/gsi/client",
};
const COLORS = ["#4986E7", "#16A765", "#9A9CFF", "#F691B2", "#FF7537", "#FAD165", "#42D692", "#8F8F8F"];  // Drive folder colors
const DEVICE = /Android/i.test(navigator.userAgent) ? "Android"
  : /Windows/i.test(navigator.userAgent) ? "Windows"
  : /iPhone|iPad/i.test(navigator.userAgent) ? "iPhone"
  : /Macintosh/i.test(navigator.userAgent) ? "Mac" : "browser";

const $ = (id) => document.getElementById(id);
const isPhone = () => matchMedia("(max-width: 760px)").matches;
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
const esc = (s) => String(s ?? "").replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
const dq = (s) => String(s).replace(/\\/g, "\\\\").replace(/'/g, "\\'");
const fmtSize = (n) => n < 1024 ? n + " B" : n < 1048576 ? (n / 1024).toFixed(0) + " KB" : (n / 1048576).toFixed(1) + " MB";
function fmtWhen(t) {
  if (!t) return "";
  const d = new Date(t), now = new Date();
  const time = d.toLocaleTimeString([], { hour: "2-digit", minute: "2-digit" });
  if (d.toDateString() === now.toDateString()) return time;
  const y = new Date(now); y.setDate(now.getDate() - 1);
  if (d.toDateString() === y.toDateString()) return "Yesterday " + time;
  return d.toLocaleDateString([], { day: "numeric", month: "short", year: d.getFullYear() === now.getFullYear() ? undefined : "numeric" });
}
function lsGet(k) { try { const v = localStorage.getItem(k); return v ? JSON.parse(v) : null; } catch { return null; } }
function lsSet(k, v) { try { localStorage.setItem(k, JSON.stringify(v)); return true; } catch { return false; } }
function lsDel(k) { try { localStorage.removeItem(k); } catch {} }
const scripts = new Map();
function loadScript(src) {
  if (!scripts.has(src)) scripts.set(src, new Promise((res, rej) => {
    const s = document.createElement("script"); s.src = src; s.async = true; s.onload = res;
    s.onerror = () => { scripts.delete(src); rej({ code: "offline", message: "Couldn't load part of the app. Check your internet connection." }); };
    document.head.appendChild(s);
  }));
  return scripts.get(src);
}

/* ---------- tiny IndexedDB key-value store (offline copies + unsaved edits) ---------- */
const idb = (() => {
  let dbp;
  const open = () => dbp || (dbp = new Promise((res, rej) => {
    try {
      const r = indexedDB.open("mb-notebook", 1);
      r.onupgradeneeded = () => r.result.createObjectStore("kv");
      r.onsuccess = () => res(r.result); r.onerror = () => rej(r.error);
    } catch (e) { rej(e); }
  }));
  const tx = async (mode, fn) => {
    try {
      const db = await open();
      return await new Promise((res, rej) => { const t = db.transaction("kv", mode); const q = fn(t.objectStore("kv")); t.oncomplete = () => res(q && q.result); t.onerror = () => rej(t.error); });
    } catch { return undefined; }
  };
  return {
    get: (k) => tx("readonly", (s) => s.get(k)),
    set: (k, v) => tx("readwrite", (s) => s.put(v, k)),
    del: (k) => tx("readwrite", (s) => s.delete(k)),
    keys: () => tx("readonly", (s) => s.getAllKeys()),
  };
})();

/* ---------- state ---------- */
const S = {
  sections: new Map(), pages: new Map(), treeAt: 0, loaded: false,
  curSection: lsGet("mbn2:lastSection"), cur: null,   // cur = {id, version, savedHtml, savedTitle}
  ed: null, edReady: false, loadingContent: false,
  dirty: false, saving: false, saveTimer: null, pendingSave: false,
  search: "", searchHits: null, online: navigator.onLine, syncState: "", attachments: [],
};

/* ---------- toast / menus / sheets ---------- */
let toastTimer;
function toast(msg, ms = 3200) { const t = $("toast"); t.textContent = msg; t.hidden = false; clearTimeout(toastTimer); toastTimer = setTimeout(() => (t.hidden = true), ms); }
function closeMenus() { document.querySelectorAll(".menu").forEach((m) => m.remove()); }
function openMenu(anchor, items) {
  closeMenus();
  const m = document.createElement("div"); m.className = "menu"; m.setAttribute("role", "menu");
  for (const it of items) {
    if (it === "-") { m.appendChild(document.createElement("hr")); continue; }
    const b = document.createElement("button"); b.setAttribute("role", "menuitem");
    if (it.danger) b.className = "danger";
    b.textContent = it.label;
    if (it.hint) { const s = document.createElement("small"); s.textContent = it.hint; b.appendChild(s); }
    b.onclick = (e) => { e.stopPropagation(); closeMenus(); it.run(); };
    m.appendChild(b);
  }
  document.body.appendChild(m);
  const r = anchor.getBoundingClientRect(), mw = m.offsetWidth, mh = m.offsetHeight;
  let left = Math.max(8, Math.min(r.right - mw, window.innerWidth - mw - 8));
  let top = r.bottom + 4; if (top + mh > window.innerHeight - 8) top = Math.max(8, r.top - mh - 4);
  m.style.left = left + "px"; m.style.top = top + "px";
  m.querySelector("button")?.focus();
}
document.addEventListener("click", (e) => { if (!e.target.closest(".menu")) closeMenus(); });
document.addEventListener("keydown", (e) => { if (e.key === "Escape") { closeMenus(); closeSheet(null); } });
let sheetResolve = null;
function closeSheet(v) { if (!sheetResolve) return; const r = sheetResolve; sheetResolve = null; $("scrim").hidden = true; $("scrim").innerHTML = ""; r(v); }
function sheet({ title, text, fields = [], ok = "OK", danger = false, cancel = "Cancel", body = [], extra }) {
  closeSheet(null);
  return new Promise((resolve) => {
    sheetResolve = resolve;
    const sc = $("scrim"); sc.innerHTML = ""; sc.hidden = false;
    const f = document.createElement("form"); f.className = "sheet"; f.setAttribute("role", "dialog"); f.setAttribute("aria-modal", "true");
    const h = document.createElement("h3"); h.textContent = title; f.appendChild(h);
    if (text) { const p = document.createElement("p"); p.textContent = text; f.appendChild(p); }
    for (const n of body) f.appendChild(n);
    const vals = {};
    for (const fd of fields) {
      const l = document.createElement("label"); l.textContent = fd.label;
      if (fd.type === "color") {
        vals[fd.id] = fd.value;
        const sw = document.createElement("div"); sw.className = "swatches";
        COLORS.forEach((c) => {
          const b = document.createElement("button"); b.type = "button"; b.style.background = c;
          b.setAttribute("aria-label", "Color " + c); b.setAttribute("aria-pressed", String(c.toLowerCase() === String(fd.value).toLowerCase()));
          b.onclick = () => { vals[fd.id] = c; sw.querySelectorAll("button").forEach((x) => x.setAttribute("aria-pressed", String(x === b))); };
          sw.appendChild(b);
        });
        l.appendChild(sw);
      } else if (fd.type === "select") {
        const s = document.createElement("select"); s.id = "f_" + fd.id;
        fd.options.forEach((o) => { const op = document.createElement("option"); op.value = o.value; op.textContent = o.label; s.appendChild(op); });
        s.value = fd.value ?? ""; l.appendChild(s);
      } else {
        const i = document.createElement("input"); i.type = "text"; i.id = "f_" + fd.id; i.value = fd.value ?? ""; i.maxLength = 120; i.autocomplete = "off"; l.appendChild(i);
      }
      f.appendChild(l);
    }
    const a = document.createElement("div"); a.className = "actions";
    if (extra) { const x = document.createElement("button"); x.type = "button"; x.className = "btn"; x.textContent = extra.label; x.onclick = () => closeSheet(extra.value); a.appendChild(x); }
    if (cancel) { const c = document.createElement("button"); c.type = "button"; c.className = "btn"; c.textContent = cancel; c.onclick = () => closeSheet(null); a.appendChild(c); }
    const o = document.createElement("button"); o.type = "submit"; o.className = "btn " + (danger ? "danger" : "primary"); o.textContent = ok; a.appendChild(o);
    f.appendChild(a);
    f.onsubmit = (e) => { e.preventDefault(); for (const fd of fields) if (fd.type !== "color") vals[fd.id] = $("f_" + fd.id).value; closeSheet(Object.keys(vals).length ? vals : true); };
    sc.onclick = (e) => { if (e.target === sc) closeSheet(null); };
    sc.appendChild(f);
    (f.querySelector("input[type=text]") || o).focus();
    f.querySelector("input[type=text]")?.select();
  });
}
function busy(text, frac) { $("loading").hidden = false; $("loadingText").textContent = text; $("loadingBar").style.width = frac == null ? "15%" : Math.round(frac * 100) + "%"; }
function idle() { $("loading").hidden = true; }

/* ============================================================
   Google sign-in (Google Identity Services, token model)
   ============================================================ */
const AUTH = { token: null, exp: 0, email: lsGet("mbn2:email"), client: null, asking: false, waiters: [] };
(function restoreToken() {
  const t = lsGet("mbn2:token");
  if (t && t.exp > Date.now() + 60000) { AUTH.token = t.t; AUTH.exp = t.exp; }
})();
const tokenValid = () => !!AUTH.token && AUTH.exp > Date.now() + 30000;

async function initAuth() {
  if (!CONFIG.clientId || /PASTE/.test(CONFIG.clientId)) { showGate("setup"); return false; }
  try { await loadScript(LIB.gis); }
  catch { if (!tokenValid()) showGate("offline"); return false; }
  AUTH.client = google.accounts.oauth2.initTokenClient({
    client_id: CONFIG.clientId, scope: SCOPE,
    callback: onToken,
    error_callback: (err) => { AUTH.asking = false; if (err && err.type === "popup_closed") toast("Sign-in was closed before it finished."); else if (err && err.type === "popup_failed_to_open") toast("Allow pop-ups for this site, then tap Sign in again.", 6000); },
  });
  return true;
}
function signIn(fresh) {
  if (!AUTH.client) { toast("Sign-in isn't ready yet. Check your internet connection."); return; }
  AUTH.asking = true;
  AUTH.client.requestAccessToken({ prompt: fresh || !AUTH.email ? "consent" : "", login_hint: AUTH.email || undefined });
}
async function onToken(r) {
  AUTH.asking = false;
  if (!r || r.error) { toast("Google sign-in didn't finish: " + (r && (r.error_description || r.error) || "unknown error")); return; }
  if (window.google?.accounts?.oauth2?.hasGrantedAllScopes && !google.accounts.oauth2.hasGrantedAllScopes(r, SCOPE)) {
    showGate("scope"); return;
  }
  AUTH.token = r.access_token; AUTH.exp = Date.now() + (Number(r.expires_in) || 3599) * 1000;
  lsSet("mbn2:token", { t: AUTH.token, exp: AUTH.exp });
  const waiters = AUTH.waiters.splice(0); waiters.forEach((w) => w());
  $("reconnect").hidden = true;
  if (!AUTH.email) {
    try { const a = await api("/about", { query: { fields: "user(emailAddress,displayName)" } }); AUTH.email = a.user.emailAddress; lsSet("mbn2:email", AUTH.email); } catch {}
  }
  renderAccount();
  if (!S.started) startNotebook(); else { refreshTree(); if (S.dirty) scheduleSave(300); pushOfflineDrafts(); }
}
function needAuth() {
  if (!S.started) { showGate("signin"); return; }
  $("reconnect").hidden = false;
  setSync("bad", "Google sign-in needed");
}
function signOut() {
  const t = AUTH.token;
  AUTH.token = null; AUTH.exp = 0; lsDel("mbn2:token");
  try { if (t) google.accounts.oauth2.revoke(t, () => {}); } catch {}
  lsDel("mbn2:email"); AUTH.email = null;
  location.reload();
}
// Google access lasts an hour; renew it on your next tap shortly before it ends (no password needed)
document.addEventListener("pointerdown", () => {
  if (AUTH.client && AUTH.token && !AUTH.asking && AUTH.exp - Date.now() < 5 * 60000) signIn(false);
}, true);

function showGate(kind) {
  const g = $("gate"); g.hidden = false;
  const msg = $("gateText"), btn = $("gateBtn");
  btn.hidden = false;
  if (kind === "setup") { msg.textContent = "Setup isn't finished: the Google Client ID is missing from config.js."; btn.hidden = true; }
  else if (kind === "offline") { msg.textContent = "You're offline. Connect to the internet to sign in the first time."; btn.hidden = true; }
  else if (kind === "scope") { msg.textContent = "MB Notebook needs permission to see and edit your Google Drive files so it can keep your pages in " + ROOT_NAME + ". Tap Sign in and tick the Google Drive box."; }
  else if (kind === "root") { msg.textContent = "The folder " + ROOT_NAME + " can't be found in your Google Drive, or it's in the trash. Restore it, then reload."; btn.hidden = true; }
  else msg.textContent = "Sign in with your Google account. Your pages are saved as Google Docs in your Drive folder " + ROOT_NAME + ".";
}

/* ============================================================
   Google Drive API
   ============================================================ */
async function api(path, { method = "GET", query, body, headers = {}, upload = false, raw = false, json } = {}) {
  if (!tokenValid()) { needAuth(); throw { code: "auth" }; }
  if (!navigator.onLine) throw { code: "offline" };
  const url = (upload ? UPLOAD : API) + path + (query ? "?" + new URLSearchParams(query) : "");
  if (json !== undefined) { body = JSON.stringify(json); headers = { ...headers, "Content-Type": "application/json; charset=UTF-8" }; }
  let r;
  try { r = await fetch(url, { method, headers: { Authorization: "Bearer " + AUTH.token, ...headers }, body }); }
  catch { throw { code: "offline" }; }
  if (r.status === 401) { AUTH.token = null; lsDel("mbn2:token"); needAuth(); throw { code: "auth" }; }
  if (!r.ok) {
    const j = await r.json().catch(() => ({}));
    const reason = j.error?.errors?.[0]?.reason || "";
    if (r.status === 429 || /rateLimit/i.test(reason)) throw { code: "rate", retry: true };
    if (r.status >= 500) throw { code: "server", retry: true };
    if (r.status === 404) throw { code: "not_found", message: j.error?.message };
    if (r.status === 403 && /storageQuota/i.test(reason)) throw { code: "quota" };
    if (r.status === 403 && /exportSizeLimit/i.test(reason)) throw { code: "too_big" };
    throw { code: "bad", status: r.status, message: j.error?.message || ("HTTP " + r.status) };
  }
  if (raw) return r;
  const ct = r.headers.get("content-type") || "";
  return ct.includes("json") ? r.json() : r.text();
}
async function retry(fn, tries = 3) {
  for (let a = 0; ; a++) {
    try { return await fn(); }
    catch (e) { if (!e || !e.retry || a >= tries - 1) throw e; await sleep(800 * (a + 1) + Math.random() * 600); }
  }
}
function multipart(meta, content, type) {
  const b = "mbn" + Math.random().toString(36).slice(2);
  const body = new Blob([
    "--" + b + "\r\nContent-Type: application/json; charset=UTF-8\r\n\r\n" + JSON.stringify(meta) + "\r\n--" + b + "\r\nContent-Type: " + type + "\r\n\r\n",
    content, "\r\n--" + b + "--",
  ]);
  return { body, headers: { "Content-Type": "multipart/related; boundary=" + b } };
}
async function uploadNew(meta, content, type, fields) {
  const size = content.size ?? new Blob([content]).size;
  if (size <= 5 * 1048576) {
    const mp = multipart(meta, content, type);
    return retry(() => api("/files", { method: "POST", upload: true, query: { uploadType: "multipart", fields }, body: mp.body, headers: mp.headers }));
  }
  // large files: resumable upload
  const start = await api("/files", { method: "POST", upload: true, raw: true, query: { uploadType: "resumable", fields }, json: meta, headers: { "X-Upload-Content-Type": type } });
  const loc = start.headers.get("Location");
  if (!loc) throw { code: "bad", message: "Upload couldn't start" };
  const r = await fetch(loc, { method: "PUT", headers: { "Content-Type": type }, body: content });
  if (!r.ok) throw { code: r.status >= 500 ? "server" : "bad", message: "Upload failed (" + r.status + ")" };
  return r.json();
}
const PAGE_FIELDS = "id,name,parents,modifiedTime,version,appProperties,webViewLink";
const SEC_FIELDS = "id,name,folderColorRgb,appProperties,createdTime";

/* ============================================================
   Notebook tree: sections (folders) + pages (Google Docs)
   ============================================================ */
function setSync(kind, text) { $("syncDot").className = "dot " + (kind || ""); $("syncText").textContent = text; S.syncState = kind; }
function cacheTree() {
  lsSet("mbn2:tree", { sections: [...S.sections.values()], pages: [...S.pages.values()], at: S.treeAt });
}
function loadCachedTree() {
  const t = lsGet("mbn2:tree"); if (!t) return false;
  S.sections = new Map(t.sections.map((s) => [s.id, s]));
  S.pages = new Map(t.pages.map((p) => [p.id, p]));
  S.treeAt = t.at || 0;
  return true;
}
function toSection(f) { return { id: f.id, name: f.name, color: f.folderColorRgb || COLORS[7], order: Number(f.appProperties?.o ?? 9999), createdAt: Date.parse(f.createdTime || 0) || 0 }; }
function toPage(f) {
  const ap = f.appProperties || {};
  return { id: f.id, title: f.name, sectionId: (f.parents || [])[0], modifiedAt: Date.parse(f.modifiedTime) || 0, version: String(f.version || ""),
    url: f.webViewLink, device: ap.d || "", words: Number(ap.w || 0), preview: ap.p || "", saveId: ap.s || "" };
}
let treeBusy = null;
async function refreshTree() {
  if (treeBusy) return treeBusy;
  treeBusy = (async () => {
    try {
      const secs = await retry(() => api("/files", { query: { q: `'${ROOT_ID}' in parents and mimeType='${FOLDER}' and trashed=false`, fields: "files(" + SEC_FIELDS + ")", pageSize: "1000", orderBy: "createdTime" } }));
      const sections = new Map(secs.files.map((f) => [f.id, toSection(f)]));
      if (!sections.size) {
        const f = await api("/files", { method: "POST", query: { fields: SEC_FIELDS }, json: { name: "Quick Notes", mimeType: FOLDER, parents: [ROOT_ID], folderColorRgb: COLORS[0], appProperties: { o: "0" } } });
        sections.set(f.id, toSection(f));
      }
      const ids = [...sections.keys()], pages = new Map();
      for (let i = 0; i < ids.length; i += 30) {
        const part = ids.slice(i, i + 30).map((id) => `'${id}' in parents`).join(" or ");
        let token;
        do {
          const r = await retry(() => api("/files", { query: { q: `mimeType='${DOC}' and trashed=false and (${part})`, fields: "nextPageToken,files(" + PAGE_FIELDS + ")", pageSize: "1000", ...(token ? { pageToken: token } : {}) } }));
          for (const f of r.files) pages.set(f.id, toPage(f));
          token = r.nextPageToken;
        } while (token);
      }
      S.sections = sections; S.pages = pages; S.treeAt = Date.now(); S.loaded = true;
      cacheTree();
      if (!S.curSection || !S.sections.has(S.curSection)) S.curSection = sortedSections()[0]?.id || null;
      setSync("ok", "Synced with Google Drive · " + fmtWhen(Date.now()));
      renderSections(); renderPages();
      if (S.cur) onCurrentPageMeta();
    } catch (e) {
      if (e.code === "not_found") { showGate("root"); }
      else if (e.code === "offline") setSync("warn", "Offline · changes are kept on this device");
      else if (e.code !== "auth") setSync("warn", "Couldn't reach Google Drive. Retrying…");
      throw e;
    } finally { treeBusy = null; }
  })();
  return treeBusy.catch(() => {});
}
const sortedSections = () => [...S.sections.values()].sort((a, b) => a.order - b.order || a.createdAt - b.createdAt || a.name.localeCompare(b.name));
const pagesIn = (sid) => [...S.pages.values()].filter((p) => p.sectionId === sid).sort((a, b) => b.modifiedAt - a.modifiedAt);

/* ---------- rendering ---------- */
function renderSections() {
  const ul = $("sectionList"); ul.innerHTML = "";
  const list = sortedSections();
  if (!list.length) { const li = document.createElement("li"); li.className = "empty"; li.textContent = S.loaded ? "No sections yet." : "Loading your sections…"; ul.appendChild(li); }
  for (const s of list) {
    const li = document.createElement("li"); li.className = "sec" + (s.id === S.curSection && !S.search ? " active" : ""); li.tabIndex = 0;
    const tab = document.createElement("span"); tab.className = "tab"; tab.style.background = s.color;
    const nm = document.createElement("span"); nm.className = "name"; nm.textContent = s.name || "Untitled section";
    const ct = document.createElement("span"); ct.className = "count"; ct.textContent = pagesIn(s.id).length;
    const more = document.createElement("button"); more.className = "icon-btn more"; more.setAttribute("aria-label", "Section options");
    more.innerHTML = '<svg viewBox="0 0 24 24" fill="currentColor"><circle cx="5" cy="12" r="1.8"/><circle cx="12" cy="12" r="1.8"/><circle cx="19" cy="12" r="1.8"/></svg>';
    more.onclick = (e) => { e.stopPropagation(); sectionMenu(more, s); };
    li.append(tab, nm, ct, more);
    const pick = () => { selectSection(s.id); if (isPhone()) setView("pages"); };
    li.onclick = pick; li.onkeydown = (e) => { if (e.key === "Enter") pick(); };
    ul.appendChild(li);
  }
  const cs = S.sections.get(S.curSection);
  $("sectionTitle").textContent = S.search ? "Search results" : (cs?.name || "Notebook");
  $("sectionTab").style.background = S.search ? "transparent" : (cs?.color || "transparent");
  $("newPageBtn").disabled = !cs;
}
function renderPages() {
  const ul = $("pageList"); ul.innerHTML = "";
  const q = S.search.trim();
  let list;
  if (q) list = (S.searchHits || []).map((id) => S.pages.get(id)).filter(Boolean);
  else list = S.curSection ? pagesIn(S.curSection) : [];
  if (!list.length) {
    const li = document.createElement("li"); li.className = "empty";
    if (q) li.innerHTML = S.searchHits ? "<strong>No matches</strong>Search looks through the titles and text of all your pages." : "Searching…";
    else if (!S.loaded) li.textContent = "Loading pages…";
    else li.innerHTML = "<strong>No pages in this section yet</strong>Tap <em>New page</em> to start writing, or <em>Open file</em> to bring in a Word, PDF or text document.";
    ul.appendChild(li); return;
  }
  for (const p of list) {
    const li = document.createElement("li"); li.className = "pg" + (S.cur && S.cur.id === p.id ? " active" : ""); li.tabIndex = 0;
    const live = S.cur && S.cur.id === p.id ? $("titleInput").value : p.title;
    const t = document.createElement("div"); t.className = "t"; t.textContent = live || "Untitled page";
    const pv = document.createElement("div"); pv.className = "p"; pv.textContent = p.preview || " ";
    const m = document.createElement("div"); m.className = "m";
    if (q) { const sec = S.sections.get(p.sectionId); const tb = document.createElement("span"); tb.className = "tab"; tb.style.background = sec?.color || COLORS[7]; m.appendChild(tb); m.append((sec?.name || "") + " · "); }
    m.append(fmtWhen(p.modifiedAt) + (p.device ? " · " + p.device : "") + (p.words ? " · " + p.words.toLocaleString() + " words" : ""));
    li.append(t, pv, m);
    const open = () => openPage(p.id);
    li.onclick = open; li.onkeydown = (e) => { if (e.key === "Enter") open(); };
    ul.appendChild(li);
  }
}
function setView(v) { $("app").dataset.view = v; }
function selectSection(id) {
  S.curSection = id; S.search = ""; S.searchHits = null; $("searchInput").value = "";
  lsSet("mbn2:lastSection", id); renderSections(); renderPages();
}
function renderAccount() { $("accountText").textContent = AUTH.email ? AUTH.email : "Google account"; }

/* ---------- search across all pages (Drive full-text) ---------- */
let searchSeq = 0;
async function runSearch(q) {
  S.search = q; S.searchHits = null; renderSections(); renderPages();
  if (!q.trim()) return;
  const seq = ++searchSeq;
  const local = [...S.pages.values()].filter((p) => (p.title + " " + p.preview).toLowerCase().includes(q.toLowerCase())).map((p) => p.id);
  try {
    const ids = [...S.sections.keys()].slice(0, 40).map((id) => `'${id}' in parents`).join(" or ");
    const r = await api("/files", { query: { q: `fullText contains '${dq(q)}' and mimeType='${DOC}' and trashed=false and (${ids})`, fields: "files(id)", pageSize: "100" } });
    if (seq !== searchSeq) return;
    S.searchHits = [...new Set([...local, ...r.files.map((f) => f.id)])];
  } catch { if (seq !== searchSeq) return; S.searchHits = local; }
  renderPages();
}

/* ============================================================
   Sections
   ============================================================ */
async function addSection() {
  const used = new Set([...S.sections.values()].map((s) => s.color.toLowerCase()));
  const v = await sheet({ title: "New section", text: "A section is a folder inside " + ROOT_NAME + ".", fields: [
    { id: "name", label: "Name", value: "" },
    { id: "color", label: "Color", type: "color", value: COLORS.find((c) => !used.has(c.toLowerCase())) || COLORS[0] }], ok: "Create section" });
  if (!v) return;
  const order = Math.max(0, ...[...S.sections.values()].map((s) => (s.order < 9999 ? s.order : 0))) + 1;
  try {
    const f = await api("/files", { method: "POST", query: { fields: SEC_FIELDS }, json: { name: v.name.trim() || "New section", mimeType: FOLDER, parents: [ROOT_ID], folderColorRgb: v.color, appProperties: { o: String(order) } } });
    S.sections.set(f.id, toSection(f)); cacheTree(); selectSection(f.id); if (isPhone()) setView("pages");
  } catch (e) { reportError(e, "create the section"); }
}
function sectionMenu(anchor, s) {
  openMenu(anchor, [
    { label: "Rename or recolor", run: async () => {
      const v = await sheet({ title: "Edit section", fields: [{ id: "name", label: "Name", value: s.name }, { id: "color", label: "Color", type: "color", value: s.color }], ok: "Save" });
      if (!v) return;
      try { const f = await api("/files/" + s.id, { method: "PATCH", query: { fields: SEC_FIELDS }, json: { name: v.name.trim() || s.name, folderColorRgb: v.color } }); Object.assign(s, toSection({ ...f, appProperties: { o: String(s.order) } })); cacheTree(); renderSections(); }
      catch (e) { reportError(e, "update the section"); }
    } },
    { label: "Move up", run: () => moveSection(s, -1) },
    { label: "Move down", run: () => moveSection(s, 1) },
    "-",
    { label: "Delete section", danger: true, run: () => deleteSection(s) },
  ]);
}
async function moveSection(s, dir) {
  const list = sortedSections(); const i = list.findIndex((x) => x.id === s.id); const j = i + dir;
  if (j < 0 || j >= list.length) return;
  [list[i], list[j]] = [list[j], list[i]];
  try {
    for (let k = 0; k < list.length; k++) if (list[k].order !== k) { await api("/files/" + list[k].id, { method: "PATCH", query: { fields: "id" }, json: { appProperties: { o: String(k) } } }); list[k].order = k; }
    cacheTree(); renderSections();
  } catch (e) { reportError(e, "reorder sections"); }
}
async function deleteSection(s) {
  if (S.sections.size <= 1) { toast("Keep at least one section. Create another one first."); return; }
  const n = pagesIn(s.id).length;
  const v = await sheet({ title: "Delete “" + s.name + "”?", text: (n ? "Its " + n + " page" + (n > 1 ? "s" : "") + " and attachments move to your Google Drive trash with it. " : "") + "You can restore it from the Drive trash for 30 days.", ok: "Delete section", danger: true });
  if (!v) return;
  try {
    await api("/files/" + s.id, { method: "PATCH", query: { fields: "id" }, json: { trashed: true } });
    if (S.cur && S.pages.get(S.cur.id)?.sectionId === s.id) closePage();
    S.sections.delete(s.id); for (const p of pagesIn(s.id)) S.pages.delete(p.id);
    if (S.curSection === s.id) S.curSection = sortedSections()[0]?.id || null;
    cacheTree(); renderSections(); renderPages(); toast("Section moved to the Drive trash");
  } catch (e) { reportError(e, "delete the section"); }
}

/* ============================================================
   Google Docs <-> editor HTML
   ============================================================ */
const DROP_PROPS = /^(padding.*|margin(?!-left).*|line-height|orphans|widows|height|max-width|min-height|page-break.*|-webkit.*|text-decoration-skip.*|transform|overflow|display|border-spacing|text-indent)$/i;
function parseStyle(str) {
  const out = [];
  for (const part of String(str || "").split(/;(?=(?:[^"]*"[^"]*")*[^"]*$)/)) {
    const i = part.indexOf(":"); if (i < 0) continue;
    const p = part.slice(0, i).trim().toLowerCase(), v = part.slice(i + 1).trim();
    if (p && v) out.push([p, v]);
  }
  return out;
}
function cleanStyle(el, isHeading, isCell) {
  const keep = [];
  for (const [p, v] of parseStyle(el.getAttribute("style"))) {
    const lv = v.toLowerCase().replace(/\s+/g, " ");
    if (DROP_PROPS.test(p) && !(isCell && /^(border|padding)/.test(p))) continue;
    if (/^border/.test(p) && !isCell && el.tagName !== "TABLE") continue;
    if (p === "margin-left") { if (el.tagName === "P" && parseFloat(v) > 0) keep.push(["padding-left", v]); continue; }
    if (p === "color" && /^(#000000|#000|black|rgb\(0, ?0, ?0\)|inherit)$/.test(lv)) continue;
    if (p === "font-weight" && /^(400|normal)$/.test(lv)) continue;
    if (p === "font-style" && lv === "normal") continue;
    if (p === "text-decoration" && /^(none|inherit)$/.test(lv)) continue;
    if (p === "vertical-align" && lv === "baseline") continue;
    if (p === "font-size" && lv === "11pt") continue;
    if (p === "font-family" && /^["']?calibri["']?$/.test(lv)) continue;
    if (p === "text-align" && (lv === "left" || lv === "start")) continue;
    if (p === "background-color" && /^(#ffffff|#fff|white|transparent)$/.test(lv) && !isCell) continue;
    if (isHeading && /^(font-size|font-weight|font-family|color)$/.test(p)) continue;
    keep.push([p, v]);
  }
  if (keep.length) el.setAttribute("style", keep.map(([p, v]) => p + ": " + v).join("; ") + ";");
  else el.removeAttribute("style");
}
function unwrap(el) { const p = el.parentNode; while (el.firstChild) p.insertBefore(el.firstChild, el); p.removeChild(el); }

/** Google Docs HTML export -> clean editor HTML */
function fromGoogle(raw) {
  const doc = new DOMParser().parseFromString(raw, "text/html");
  const body = doc.body;
  body.querySelectorAll("style,script,meta,title").forEach((n) => n.remove());
  // links: unwrap Google's redirect
  body.querySelectorAll("a").forEach((a) => {
    const h = a.getAttribute("href");
    if (!h) { unwrap(a); return; }
    const m = h.match(/^https:\/\/www\.google\.com\/url\?q=([^&]+)/);
    if (m) a.setAttribute("href", decodeURIComponent(m[1]));
    a.removeAttribute("style"); a.setAttribute("target", "_blank");
  });
  // picture wrappers
  body.querySelectorAll("img").forEach((img) => {
    const w = parseFloat(img.style.width), h = parseFloat(img.style.height);
    if (w) img.setAttribute("width", Math.round(w)); if (h) img.setAttribute("height", Math.round(h));
    img.removeAttribute("style"); img.removeAttribute("title");
    const wrap = img.parentElement;
    if (wrap && wrap.tagName === "SPAN" && /inline-block/.test(wrap.getAttribute("style") || "")) wrap.removeAttribute("style");
  });
  // page breaks
  body.querySelectorAll("hr").forEach((hr) => { if (/page-break-before/.test(hr.getAttribute("style") || "")) hr.replaceWith(doc.createComment(" pagebreak ")); });
  // Docs "title" paragraph -> Title heading
  body.querySelectorAll("p.title").forEach((p) => { const h = doc.createElement("h1"); h.innerHTML = p.innerHTML; p.replaceWith(h); });
  // list levels from Docs classes (lst-kix_xxx-N) -> real nesting
  const lists = [...body.querySelectorAll(":scope > ul, :scope > ol")];
  let prev = null;
  for (const l of lists) {
    const m = (l.getAttribute("class") || "").match(/lst-kix_[\w]+-(\d+)/);
    const lvl = m ? Number(m[1]) : 0;
    l.dataset.lvl = lvl;
    const sib = l.previousElementSibling;
    if (lvl > 0 && prev && (sib === prev || sib?.dataset?.lvlRoot === "1")) {
      // attach under the last item one level up
      let host = prev, hostLvl = Number(prev.dataset.lvl);
      let li = host.lastElementChild;
      while (hostLvl < lvl - 1 && li) {
        const sub = li.querySelector(":scope > ul:last-of-type, :scope > ol:last-of-type");
        if (!sub) break; host = sub; hostLvl++; li = sub.lastElementChild;
      }
      if (li) { li.appendChild(l); continue; }
    }
    if (lvl === 0 && prev && sib === prev && prev.tagName === l.tagName && !l.classList.contains("start")) {
      while (l.firstChild) prev.appendChild(l.firstChild); l.remove(); continue;   // same list continues
    }
    prev = l;
  }
  body.querySelectorAll("ul,ol").forEach((l) => { delete l.dataset.lvl; l.removeAttribute("class"); if (l.tagName === "OL" && l.getAttribute("start") === "1") l.removeAttribute("start"); l.removeAttribute("style"); });
  body.querySelectorAll("li").forEach((li) => li.removeAttribute("class"));
  // styles
  body.querySelectorAll("[style]").forEach((el) => {
    const heading = /^H[1-6]$/.test(el.tagName) || !!el.closest("h1,h2,h3,h4,h5,h6");
    const cell = el.tagName === "TD" || el.tagName === "TH";
    cleanStyle(el, heading, cell);
  });
  body.querySelectorAll("[class]").forEach((el) => el.removeAttribute("class"));
  body.querySelectorAll("[id]").forEach((el) => el.removeAttribute("id"));
  // span styles -> semantic tags
  body.querySelectorAll("span").forEach((sp) => {
    const st = sp.style;
    let node = sp;
    const wrapIn = (tag) => { const w = doc.createElement(tag); sp.parentNode.insertBefore(w, sp); w.appendChild(sp); };
    if (/^(700|bold)$/.test(st.fontWeight)) { st.removeProperty("font-weight"); wrapIn("strong"); }
    if (st.fontStyle === "italic") { st.removeProperty("font-style"); wrapIn("em"); }
    if (st.verticalAlign === "super") { st.removeProperty("vertical-align"); wrapIn("sup"); }
    if (st.verticalAlign === "sub") { st.removeProperty("vertical-align"); wrapIn("sub"); }
    if (!sp.getAttribute("style")) sp.removeAttribute("style");
    if (!sp.attributes.length) unwrap(node);
  });
  // to-do lines (☐ / ☑) -> checklist
  const isTodo = (p) => p && p.tagName === "P" && /^\s*[☐☑☒]/.test(p.textContent);
  for (const p of [...body.querySelectorAll(":scope > p")]) {
    if (!p.isConnected || !isTodo(p)) continue;
    const ul = doc.createElement("ul"); ul.className = "todo"; p.before(ul);
    let cur = p;
    while (isTodo(cur)) {
      const next = cur.nextElementSibling;
      const li = doc.createElement("li");
      const done = /^\s*[☑☒]/.test(cur.textContent);
      if (done) li.className = "done";
      // strip the leading box character from the first text node
      const walker = doc.createTreeWalker(cur, NodeFilter.SHOW_TEXT);
      let tn; while ((tn = walker.nextNode())) { if (tn.nodeValue.trim()) { tn.nodeValue = tn.nodeValue.replace(/^\s*[☐☑☒]\s?/, ""); break; } }
      li.innerHTML = cur.innerHTML; ul.appendChild(li); cur.remove(); cur = next;
    }
  }
  body.querySelectorAll("ul").forEach((ul) => {
    const items = [...ul.children].filter((n) => n.tagName === "LI");
    if (!items.length || !items.every((li) => /^\s*[☐☑☒]/.test(li.textContent))) return;
    ul.className = "todo";
    for (const li of items) {
      if (/^\s*[☑☒]/.test(li.textContent)) li.className = "done";
      const w = doc.createTreeWalker(li, NodeFilter.SHOW_TEXT); let tn;
      while ((tn = w.nextNode())) { if (tn.nodeValue.trim()) { tn.nodeValue = tn.nodeValue.replace(/^\s*[☐☑☒]\s?/, ""); break; } }
    }
  });
  // empty Docs paragraphs -> blank lines
  body.querySelectorAll("p").forEach((p) => { if (!p.textContent.trim() && !p.querySelector("img,br")) p.innerHTML = "<br>"; });
  return body.innerHTML.replace(/<!-- pagebreak -->/g, "<!-- pagebreak -->");
}

/** editor HTML -> HTML that Google Docs imports well */
function toGoogle(html, title) {
  const doc = new DOMParser().parseFromString("<body>" + html + "</body>", "text/html");
  const b = doc.body;
  b.querySelectorAll("ul.todo").forEach((ul) => {
    const frag = doc.createDocumentFragment();
    for (const li of [...ul.children]) {
      li.querySelectorAll("ul,ol").forEach((n) => n.remove());
      const p = doc.createElement("p"); p.setAttribute("style", "margin-left:30pt");
      p.innerHTML = (li.classList.contains("done") ? "☑ " : "☐ ") + li.innerHTML; frag.appendChild(p);
    }
    ul.replaceWith(frag);
  });
  b.querySelectorAll("blockquote").forEach((q) => {
    const ps = q.querySelectorAll("p").length ? [...q.querySelectorAll("p")] : null;
    const out = doc.createDocumentFragment();
    (ps || [q]).forEach((src) => { const p = doc.createElement("p"); p.setAttribute("style", "margin-left:36pt;font-style:italic;color:#4A5565"); p.innerHTML = src.innerHTML; out.appendChild(p); });
    q.replaceWith(out);
  });
  b.querySelectorAll("table").forEach((t) => { if (!t.getAttribute("border")) t.setAttribute("border", "1"); });
  b.querySelectorAll("p[style*='padding-left']").forEach((p) => { p.style.marginLeft = p.style.paddingLeft; p.style.removeProperty("padding-left"); });
  let out = b.innerHTML.replace(/<!--\s*pagebreak\s*-->/g, '<hr style="page-break-before:always;display:none">');
  return '<html><head><meta charset="utf-8"><title>' + esc(title) + '</title><style>body{font-family:Calibri,Carlito,sans-serif;font-size:11pt}h1{font-size:20pt}h2{font-size:16pt}h3{font-size:13pt}h4{font-size:12pt}</style></head><body>' + out + "</body></html>";
}
function textOf(html) { const d = document.createElement("div"); d.innerHTML = html; return (d.textContent || "").replace(/\s+/g, " ").trim(); }
function previewOf(text) {  // appProperties hold ~120 bytes
  let p = text.slice(0, 90); while (new TextEncoder().encode("p" + p).length > 118) p = p.slice(0, -1); return p;
}

/* Was this page changed somewhere else since `base` (our last save or open)?
   Google Docs bumps a Doc's version/modified time on its own right after a save,
   so we look at who saved (MB Notebook save id) and allow a short settle window. */
function changedElsewhere(meta, base) {
  if (!meta || !base) return false;
  if (meta.saveId && meta.saveId !== (base.saveId || "")) return true;   // another device saved with MB Notebook
  return (meta.modifiedAt || 0) > (base.writeAt || 0) + 20000;            // edited in Google Docs itself
}
const newSaveId = () => DEVICE.slice(0, 1) + Date.now().toString(36) + Math.random().toString(36).slice(2, 6);

/* ============================================================
   Pages: open / save / create / delete / move
   ============================================================ */
async function readPage(id) {
  const raw = await retry(() => api("/files/" + id + "/export", { query: { mimeType: "text/html" } }));
  return fromGoogle(raw);
}
async function openPage(id, { html: given } = {}) {
  const meta = S.pages.get(id); if (!meta) return;
  if (S.cur && S.cur.id !== id) await flushSave();
  await editorReady;
  S.cur = { id, version: meta.version, savedHtml: null, savedTitle: meta.title || "", saveId: meta.saveId || "", writeAt: meta.modifiedAt || 0 };
  S.dirty = false; clearTimeout(S.saveTimer);
  if (meta.sectionId !== S.curSection && !S.search) S.curSection = meta.sectionId;
  lsSet("mbn2:lastPage", id);
  $("noPage").hidden = true; setPageControls(true);
  $("titleInput").value = meta.title || "";
  renderMeta(); renderSections(); renderPages();
  if (isPhone()) setView("editor");

  let html = given, from = "cloud";
  const cached = await idb.get("page:" + id);
  const draft = await idb.get("draft:" + id);
  if (html == null) {
    if (cached && cached.version === meta.version) { html = cached.html; from = "cache"; }
    else {
      busy("Opening page…");
      try { html = await readPage(id); idb.set("page:" + id, { html, version: meta.version, title: meta.title }); }
      catch (e) {
        if (cached) { html = cached.html; from = "stale"; }
        else { idle(); reportError(e, "open this page"); closePage(); return; }
      }
      idle();
    }
  }
  if (!S.cur || S.cur.id !== id) return;
  let restored = false;
  if (draft && given == null && !changedElsewhere(meta, draft.base || { saveId: meta.saveId, writeAt: 0 })) {
    html = draft.html; restored = true; if (draft.title != null) $("titleInput").value = draft.title;
  } else if (draft && given == null) {
    // the page changed elsewhere after these offline edits: keep both
    await idb.del("draft:" + id);
    saveConflictCopy(meta, draft);
  }
  setEditorContent(html);
  S.cur.savedHtml = given != null ? null : (from === "stale" ? S.ed.getContent() : S.ed.getContent());
  if (restored) { S.cur.savedHtml = null; S.dirty = true; setSave("unsaved"); scheduleSave(300); toast("Restored edits you made on this device that hadn't reached Google Drive yet."); }
  else if (from === "stale") { setSave("offline"); toast("You're offline, so this is the copy saved on this device."); }
  else setSave(given != null ? "unsaved" : "saved");
  loadAttachments(id);
}
function setEditorContent(html) {
  S.loadingContent = true; S.ed.setContent(html || ""); S.ed.undoManager.clear(); S.ed.getWin()?.scrollTo(0, 0); S.loadingContent = false;
}
function setPageControls(on) {
  for (const id of ["titleInput", "attachBtn", "exportBtn", "pageMenuBtn"]) $(id).disabled = !on;
  if (S.ed) S.ed.mode.set(on ? "design" : "readonly");
}
function closePage() {
  S.cur = null; S.dirty = false; clearTimeout(S.saveTimer);
  $("noPage").hidden = false; $("titleInput").value = ""; setPageControls(false);
  $("metaRow").innerHTML = ""; setSave("idle"); lsDel("mbn2:lastPage"); renderPages();
  if (isPhone()) setView("pages");
}
function renderMeta() {
  const row = $("metaRow"); row.innerHTML = "";
  if (!S.cur) return;
  const meta = S.pages.get(S.cur.id); if (!meta) return;
  const sec = S.sections.get(meta.sectionId);
  const info = document.createElement("span");
  info.textContent = (sec ? sec.name + " · " : "") + (meta.modifiedAt ? "Edited " + fmtWhen(meta.modifiedAt) + (meta.device ? " on " + meta.device : "") : "");
  row.appendChild(info);
  if (meta.url) {
    const c = document.createElement("span"); c.className = "chip drive";
    const a = document.createElement("a"); a.href = meta.url; a.target = "_blank"; a.rel = "noopener"; a.textContent = "Open in Google Docs";
    c.appendChild(a); c.style.paddingRight = "10px"; row.appendChild(c);
  }
  for (const f of S.attachments) {
    const c = document.createElement("span"); c.className = "chip";
    const a = document.createElement("a"); a.href = f.webViewLink; a.target = "_blank"; a.rel = "noopener"; a.textContent = "📎 " + f.name;
    const sz = document.createElement("span"); sz.className = "sz"; sz.textContent = f.size ? fmtSize(Number(f.size)) : "";
    const x = document.createElement("button"); x.type = "button"; x.setAttribute("aria-label", "Remove " + f.name); x.textContent = "×"; x.onclick = () => removeAttachment(f);
    c.append(a, sz, x); row.appendChild(c);
  }
}
/* another device (or Google Docs) changed the page we have open */
async function onCurrentPageMeta() {
  const meta = S.pages.get(S.cur.id);
  if (!meta) { if (S.loaded) { toast("This page was deleted or moved to the trash on another device."); closePage(); } return; }
  if (document.activeElement !== $("titleInput") && !S.dirty) $("titleInput").value = meta.title || "";
  renderMeta();
  if (!S.saving && changedElsewhere(meta, S.cur)) {
    if (!S.dirty) {
      const id = S.cur.id, keep = S.ed.getWin()?.scrollY || 0;
      try {
        const html = await readPage(id);
        if (!S.cur || S.cur.id !== id || S.dirty) return;
        S.cur.version = meta.version; S.cur.saveId = meta.saveId; S.cur.writeAt = meta.modifiedAt;
        setEditorContent(html); S.cur.savedHtml = S.ed.getContent(); S.cur.savedTitle = meta.title;
        idb.set("page:" + id, { html, version: meta.version, title: meta.title });
        S.ed.getWin()?.scrollTo(0, keep);
        toast("Updated with changes from " + (meta.device || "another device"));
      } catch {}
    } else S.cur.remoteChanged = true;   // you have unsaved edits: you'll be asked when this page saves
  }
}

/* ---------- saving ---------- */
function setSave(state) {
  const map = { idle: ["", ""], saved: ["ok", "Saved"], saving: ["", "Saving…"], unsaved: ["warn", "Unsaved"], offline: ["warn", "Saved on this device"], error: ["bad", "Not saved"] };
  const [k, t] = map[state];
  $("saveDot").className = "dot " + k; $("saveText").textContent = t; $("saveStatus").hidden = state === "idle";
  $("saveStatus").title = state === "offline" ? "Will save to Google Drive when you're back online" : t;
}
function markDirty() {
  if (S.loadingContent || !S.cur) return;
  S.dirty = true; setSave("unsaved"); scheduleSave(); writeDraftSoon();
}
function saveDelay() { const n = (S.ed && S.ed.getBody()?.innerHTML.length) || 0; return Math.min(10000, 2000 + Math.floor(n / 200000) * 1000); }
function scheduleSave(ms) { clearTimeout(S.saveTimer); S.saveTimer = setTimeout(() => saveNow(), ms ?? saveDelay()); }
let draftTimer;
function writeDraftSoon() { clearTimeout(draftTimer); draftTimer = setTimeout(writeDraft, 600); }
function writeDraft() {
  if (!S.cur || !S.dirty) return;
  idb.set("draft:" + S.cur.id, { html: S.ed.getContent(), title: $("titleInput").value, base: { saveId: S.cur.saveId, writeAt: S.cur.writeAt }, t: Date.now() });
}
async function saveNow({ force = false } = {}) {
  if (!S.cur || !S.edReady) return;
  if (S.saving) { S.pendingSave = true; return; }
  const pid = S.cur.id, meta = S.pages.get(pid);
  const html = S.ed.getContent();
  const title = $("titleInput").value.trim() || "Untitled page";
  const contentChanged = html !== S.cur.savedHtml, titleChanged = title !== S.cur.savedTitle;
  if (!contentChanged && !titleChanged) { S.dirty = false; setSave("saved"); return; }
  if (!tokenValid() || !navigator.onLine) { writeDraft(); setSave("offline"); return; }
  S.saving = true; setSave("saving"); writeDraft();
  try {
    // did someone change this page since we opened it?
    if (!force) {
      const now = await api("/files/" + pid, { query: { fields: "modifiedTime,appProperties,trashed" } });
      if (now.trashed) { S.saving = false; toast("This page is in the Drive trash. Restore it from Google Drive to keep editing."); setSave("error"); return; }
      const remote = { saveId: now.appProperties?.s || "", modifiedAt: Date.parse(now.modifiedTime) || 0 };
      if (changedElsewhere(remote, S.cur)) {
        S.saving = false;
        const v = await sheet({ title: "This page changed on another device",
          text: "It was edited somewhere else while you were editing here. Keep your version here (the other changes are replaced), or load the other version (your edits here are saved as a separate copy).",
          ok: "Keep my version", extra: { label: "Load other version", value: "theirs" }, cancel: "" });
        if (v === "theirs") { await saveConflictCopy(meta, { html, title }); S.dirty = false; await idb.del("draft:" + pid); S.pages.get(pid).version = ""; return reopen(pid); }
        return saveNow({ force: true });
      }
    }
    const text = textOf(html);
    const sid = newSaveId();
    const metaBody = { name: title, appProperties: { d: DEVICE, w: String(text ? text.split(" ").length : 0), p: previewOf(text), s: sid } };
    let f;
    if (contentChanged) {
      const mp = multipart(metaBody, toGoogle(html, title), "text/html");
      f = await retry(() => api("/files/" + pid, { method: "PATCH", upload: true, query: { uploadType: "multipart", fields: PAGE_FIELDS }, body: mp.body, headers: mp.headers }));
    } else {
      f = await retry(() => api("/files/" + pid, { method: "PATCH", query: { fields: PAGE_FIELDS }, json: metaBody }));
    }
    const p = toPage(f);
    S.pages.set(pid, p); cacheTree();
    if (S.cur && S.cur.id === pid) { S.cur.version = p.version; S.cur.saveId = sid; S.cur.writeAt = p.modifiedAt; S.cur.savedTitle = title; if (contentChanged) S.cur.savedHtml = html; S.cur.remoteChanged = false; }
    idb.set("page:" + pid, { html, version: p.version, title });
    const still = S.cur && S.cur.id === pid && (S.ed.getContent() !== S.cur.savedHtml || ($("titleInput").value.trim() || "Untitled page") !== S.cur.savedTitle);
    S.dirty = !!still;
    if (!still) await idb.del("draft:" + pid);
    setSave(still ? "unsaved" : "saved");
    renderPages(); renderMeta();
    if (still) scheduleSave(500);
  } catch (e) {
    console.warn("save failed", e);
    S.dirty = true;
    if (e.code === "offline" || e.code === "auth") setSave("offline");
    else { setSave("error"); reportError(e, "save this page"); scheduleSave(15000); }
  } finally {
    S.saving = false;
    if (S.pendingSave) { S.pendingSave = false; scheduleSave(200); }
  }
}
async function reopen(pid) { S.cur = null; await openPage(pid); }
async function flushSave() {
  clearTimeout(S.saveTimer);
  if (S.cur && (S.dirty || ($("titleInput").value.trim() && $("titleInput").value.trim() !== S.cur.savedTitle))) await saveNow();
  while (S.saving) await sleep(50);
}
async function saveConflictCopy(meta, draft) {
  try {
    const title = (draft.title || meta.title || "Untitled page") + " (edited on " + DEVICE + ")";
    const mp = multipart({ name: title, mimeType: DOC, parents: [meta.sectionId], appProperties: { d: DEVICE } }, toGoogle(draft.html, title), "text/html");
    const f = await api("/files", { method: "POST", upload: true, query: { uploadType: "multipart", fields: PAGE_FIELDS }, body: mp.body, headers: mp.headers });
    S.pages.set(f.id, toPage(f)); cacheTree(); renderPages();
    toast("Your other edits were kept as “" + title + "”.", 6000);
  } catch (e) { reportError(e, "keep your other edits"); }
}
/* edits made offline on pages you're not looking at */
async function pushOfflineDrafts() {
  const keys = (await idb.keys()) || [];
  for (const k of keys) {
    if (!String(k).startsWith("draft:")) continue;
    const pid = k.slice(6); if (S.cur && S.cur.id === pid) continue;
    const d = await idb.get(k), meta = S.pages.get(pid);
    if (!d || !meta) continue;
    try {
      if (!changedElsewhere(meta, d.base || { saveId: meta.saveId, writeAt: 0 })) {
        const mp = multipart({ name: d.title || meta.title, appProperties: { d: DEVICE, s: newSaveId() } }, toGoogle(d.html, d.title || meta.title), "text/html");
        const f = await api("/files/" + pid, { method: "PATCH", upload: true, query: { uploadType: "multipart", fields: PAGE_FIELDS }, body: mp.body, headers: mp.headers });
        S.pages.set(pid, toPage(f)); idb.set("page:" + pid, { html: d.html, version: String(f.version), title: f.name });
      } else await saveConflictCopy(meta, d);
      await idb.del(k);
    } catch (e) { if (e.code === "offline" || e.code === "auth") return; }
  }
  cacheTree(); renderPages();
}

async function createPage({ title = "Untitled page", html = "", sectionId = S.curSection } = {}) {
  if (!sectionId) { toast("Create a section first."); return null; }
  const mp = multipart({ name: title, mimeType: DOC, parents: [sectionId], appProperties: { d: DEVICE, w: "0", p: "", s: newSaveId() } }, toGoogle(html, title), "text/html");
  const f = await api("/files", { method: "POST", upload: true, query: { uploadType: "multipart", fields: PAGE_FIELDS }, body: mp.body, headers: mp.headers });
  S.pages.set(f.id, toPage(f)); cacheTree();
  return f.id;
}
async function newPage() {
  try {
    busy("Creating page…");
    const id = await createPage(); idle();
    if (id) { await openPage(id, { html: "" }); S.cur.savedHtml = S.ed.getContent(); setSave("saved"); setTimeout(() => { $("titleInput").focus(); $("titleInput").select(); }, 50); }
  } catch (e) { idle(); reportError(e, "create a page"); }
}
async function deletePage() {
  const p = S.cur && S.pages.get(S.cur.id); if (!p) return;
  const v = await sheet({ title: "Delete this page?", text: "“" + (p.title || "Untitled page") + "” and its attachments move to your Google Drive trash on all your devices. You can restore them from the Drive trash for 30 days.", ok: "Delete page", danger: true });
  if (!v) return;
  const atts = S.attachments.slice();
  S.dirty = false; clearTimeout(S.saveTimer); closePage();
  try {
    await api("/files/" + p.id, { method: "PATCH", query: { fields: "id" }, json: { trashed: true } });
    for (const a of atts) api("/files/" + a.id, { method: "PATCH", query: { fields: "id" }, json: { trashed: true } }).catch(() => {});
    S.pages.delete(p.id); idb.del("page:" + p.id); idb.del("draft:" + p.id); cacheTree(); renderPages(); renderSections();
    toast("Page moved to the Drive trash");
  } catch (e) { reportError(e, "delete the page"); }
}
async function movePage() {
  const p = S.cur && S.pages.get(S.cur.id); if (!p) return;
  const v = await sheet({ title: "Move page", fields: [{ id: "sec", label: "Move to section", type: "select", value: p.sectionId, options: sortedSections().map((s) => ({ value: s.id, label: s.name })) }], ok: "Move" });
  if (!v || v.sec === p.sectionId) return;
  try {
    await flushSave();
    const f = await api("/files/" + p.id, { method: "PATCH", query: { addParents: v.sec, removeParents: p.sectionId, fields: PAGE_FIELDS } });
    for (const a of S.attachments) api("/files/" + a.id, { method: "PATCH", query: { addParents: v.sec, removeParents: p.sectionId, fields: "id" } }).catch(() => {});
    S.pages.set(p.id, toPage(f)); if (S.cur) S.cur.version = String(f.version);
    S.curSection = v.sec; lsSet("mbn2:lastSection", v.sec); cacheTree(); renderSections(); renderPages(); renderMeta();
    toast("Moved to " + (S.sections.get(v.sec)?.name || "section"));
  } catch (e) { reportError(e, "move the page"); }
}
async function duplicatePage() {
  const p = S.cur && S.pages.get(S.cur.id); if (!p) return;
  await flushSave();
  try { busy("Copying page…"); const id = await createPage({ title: (p.title || "Untitled page") + " (copy)", html: S.ed.getContent(), sectionId: p.sectionId }); idle(); if (id) openPage(id); }
  catch (e) { idle(); reportError(e, "copy the page"); }
}

/* ---------- attachments: real files next to the page in its section folder ---------- */
async function loadAttachments(pid) {
  S.attachments = [];
  try {
    const r = await api("/files", { query: { q: `trashed=false and appProperties has { key='mbPage' and value='${dq(pid)}' }`, fields: "files(id,name,size,mimeType,webViewLink)", pageSize: "100" } });
    if (S.cur && S.cur.id === pid) { S.attachments = r.files; renderMeta(); }
  } catch {}
}
async function attachFile(file) {
  if (!S.cur) return;
  const meta = S.pages.get(S.cur.id);
  busy("Uploading " + file.name + "…");
  try {
    const f = await uploadNew({ name: file.name, parents: [meta.sectionId], appProperties: { mbPage: meta.id } }, file, file.type || "application/octet-stream", "id,name,size,mimeType,webViewLink");
    S.attachments.push(f); renderMeta(); toast("Attached " + file.name);
  } catch (e) { reportError(e, "attach the file"); }
  idle();
}
async function removeAttachment(f) {
  const v = await sheet({ title: "Remove attachment?", text: "“" + f.name + "” moves to your Google Drive trash.", ok: "Remove", danger: true });
  if (!v) return;
  try { await api("/files/" + f.id, { method: "PATCH", query: { fields: "id" }, json: { trashed: true } }); S.attachments = S.attachments.filter((x) => x.id !== f.id); renderMeta(); }
  catch (e) { reportError(e, "remove the attachment"); }
}

/* ---------- open files: Google converts Word and PDF into Google Docs ---------- */
async function importFile(file) {
  const ext = (file.name.split(".").pop() || "").toLowerCase();
  const title = file.name.replace(/\.[^.]+$/, "");
  const sectionId = S.curSection;
  if (!sectionId) { toast("Create a section first."); return; }
  try {
    let id;
    if (ext === "docx" || ext === "doc" || ext === "odt" || ext === "rtf" || ext === "pdf") {
      busy("Converting " + file.name + " with Google Docs…");
      const type = file.type || (ext === "pdf" ? "application/pdf" : ext === "docx" ? DOCX : "application/octet-stream");
      const f = await uploadNew({ name: title, mimeType: DOC, parents: [sectionId], appProperties: { d: DEVICE } }, file, type, PAGE_FIELDS);
      S.pages.set(f.id, toPage(f)); id = f.id;
      if (ext === "pdf") {
        busy("Keeping the original PDF…", 0.9);
        try { await uploadNew({ name: file.name, parents: [sectionId], appProperties: { mbPage: f.id } }, file, "application/pdf", "id"); } catch {}
      }
    } else if (ext === "md" || ext === "markdown" || ext === "txt") {
      busy("Reading " + file.name + "…");
      const text = await file.text();
      let html;
      if (ext === "txt") html = text.replace(/\r\n?/g, "\n").split(/\n{2,}/).map((p) => "<p>" + esc(p).replace(/\n/g, "<br>") + "</p>").join("");
      else { await loadScript(LIB.marked); html = window.marked.parse(text).replace(/<li>\s*\[ \]\s*/g, "<li>☐ ").replace(/<li>\s*\[x\]\s*/gi, "<li>☑ "); }
      id = await createPage({ title, html, sectionId });
    } else { idle(); toast("Open Word (.docx, .doc), PDF, Markdown or text files."); return; }
    cacheTree(); renderSections(); renderPages();
    idle();
    await openPage(id);
    toast("Opened " + file.name + " as a new page");
  } catch (e) { idle(); reportError(e, "open " + file.name); }
}

/* ---------- export: Google Docs does the converting ---------- */
function exportMenu(anchor) {
  openMenu(anchor, [
    { label: "Word document", hint: ".docx", run: () => exportAs("docx") },
    { label: "PDF", hint: ".pdf", run: () => exportAs("pdf") },
    { label: "Markdown", hint: ".md", run: () => exportAs("md") },
    { label: "Plain text", hint: ".txt", run: () => exportAs("txt") },
  ]);
}
const safeName = (t) => (t || "Untitled page").replace(/[\\/:*?"<>|]+/g, " ").trim().slice(0, 80) || "page";
function download(blob, name) {
  const a = document.createElement("a"); a.href = URL.createObjectURL(blob); a.download = name;
  document.body.appendChild(a); a.click(); setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 4000);
}
async function exportAs(kind) {
  if (!S.cur) return;
  await flushSave();
  const title = $("titleInput").value.trim() || "Untitled page";
  try {
    busy("Preparing " + kind.toUpperCase() + "…");
    if (kind === "md") {
      await loadScript(LIB.turndown);
      const td = new window.TurndownService({ headingStyle: "atx", codeBlockStyle: "fenced", bulletListMarker: "-" });
      td.keep(["table"]);
      td.addRule("todo", { filter: (n) => n.nodeName === "LI" && n.parentNode.classList?.contains("todo"), replacement: (c, n) => "- [" + (n.classList.contains("done") ? "x" : " ") + "] " + c.trim() + "\n" });
      download(new Blob(["# " + title + "\n\n" + td.turndown(S.ed.getContent())], { type: "text/markdown" }), safeName(title) + ".md");
    } else {
      const mime = { docx: DOCX, pdf: "application/pdf", txt: "text/plain" }[kind];
      const r = await api("/files/" + S.cur.id + "/export", { query: { mimeType: mime }, raw: true });
      download(await r.blob(), safeName(title) + "." + kind);
    }
  } catch (e) { reportError(e, "export the page"); }
  idle();
}

function reportError(e, what) {
  console.warn(what, e);
  const c = e && e.code;
  if (c === "auth") return;  // reconnect bar is showing
  if (c === "offline") toast("You're offline, so I couldn't " + what + ". Try again when you're connected.");
  else if (c === "quota") toast("Your Google Drive storage is full, so I couldn't " + what + ".", 6000);
  else if (c === "too_big") toast("This page is too large for Google Docs to export.", 6000);
  else if (c === "rate") toast("Google Drive is busy. Wait a few seconds and try again.");
  else if (c === "not_found") toast("That item no longer exists in Google Drive.");
  else toast("Couldn't " + what + (e && e.message ? ": " + e.message : "."), 6000);
}

/* ============================================================
   Editor (TinyMCE, Word-style)
   ============================================================ */
const FONT_CSS = "https://fonts.googleapis.com/css2?family=Carlito:ital,wght@0,400;0,700;1,400;1,700&family=Arimo:ital,wght@0,400;0,700;1,400;1,700&family=Tinos:ital,wght@0,400;0,700;1,400;1,700&family=Cousine:wght@400;700&family=Cairo:wght@400;700&family=Noto+Naskh+Arabic:wght@400;700&family=Roboto:ital,wght@0,400;0,700;1,400&family=Merriweather:ital,wght@0,400;0,700;1,400&display=swap";
const PAPER_CSS = `
html { background: #E9ECF1; }
body { box-sizing: border-box; max-width: 8.27in; min-height: 11.69in; margin: 28px auto 60px; padding: 0.9in 1in; background: #fff; color: #1B1F24;
  box-shadow: 0 1px 3px rgba(20,30,45,.14), 0 10px 30px rgba(20,30,45,.08);
  font-family: Calibri, Carlito, "Segoe UI", Arial, sans-serif; font-size: 11pt; line-height: 1.45; }
/* Full width (default): the page fills the window like OneNote */
html.full { background: #fff; min-height: 100%; }
html.full body { max-width: none; min-height: 0; margin: 0; padding: 28px 48px 120px; box-shadow: none; }
@media (max-width: 760px) { html { background: #fff; } body, html.full body { margin: 0; padding: 16px 16px 80px; box-shadow: none; min-height: 0; max-width: none; } }
h1, h2, h3, h4 { line-height: 1.2; margin: 1.1em 0 .4em; }
h1 { font-size: 20pt; color: #1F3A63; } h2 { font-size: 16pt; color: #1F4E8C; } h3 { font-size: 13pt; color: #1F4E8C; } h4 { font-size: 12pt; color: #1F4E8C; }
p { margin: 0 0 .5em; }
blockquote { margin: .8em 0; padding: .2em 0 .2em 1em; border-left: 3px solid #C5CEDB; color: #4A5565; }
table { border-collapse: collapse; } td, th { border: 1px solid #B9C2CF; padding: 4px 8px; vertical-align: top; } th { background: #F1F4F8; }
img { max-width: 100%; height: auto; } a { color: #1D5BB8; }
pre { background: #F4F6F9; padding: 10px 12px; border-radius: 4px; white-space: pre-wrap; font-family: "Courier New", Cousine, monospace; font-size: 10pt; }
hr { border: 0; border-top: 1px solid #C5CEDB; margin: 1.2em 0; }
ul.todo { list-style: none; padding-left: 4px; }
ul.todo > li { position: relative; padding-left: 28px; margin: .25em 0; }
ul.todo > li::before { content: ""; position: absolute; left: 0; top: .18em; width: 15px; height: 15px; border: 1.5px solid #6B7686; border-radius: 4px; background: #fff; cursor: pointer; }
ul.todo > li.done { color: #7A8494; text-decoration: line-through; }
ul.todo > li.done::before { background: #1D5BB8 url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 24 24' fill='none' stroke='white' stroke-width='3.5'%3E%3Cpath d='M5 12.5l4.5 4.5L19 7'/%3E%3C/svg%3E") center/11px no-repeat; border-color: #1D5BB8; }
[dir=rtl] ul.todo > li, ul.todo[dir=rtl] > li { padding-left: 0; padding-right: 28px; }
[dir=rtl] ul.todo > li::before, ul.todo[dir=rtl] > li::before { left: auto; right: 0; }
`;
const FONT_FORMATS = [
  "Calibri=Calibri,Carlito,sans-serif", "Arial=Arial,Arimo,Helvetica,sans-serif", "Times New Roman='Times New Roman',Tinos,Times,serif",
  "Georgia=Georgia,serif", "Courier New='Courier New',Cousine,monospace", "Roboto=Roboto,sans-serif", "Merriweather=Merriweather,serif",
  "Cairo (Arabic)=Cairo,sans-serif", "Noto Naskh Arabic='Noto Naskh Arabic',serif",
].join(";");
let resolveEditor; const editorReady = new Promise((r) => (resolveEditor = r));
const isDark = () => { const t = document.documentElement.getAttribute("data-theme"); return t ? t === "dark" : matchMedia("(prefers-color-scheme: dark)").matches; };

/* shrink big pictures before they go into a page, so saves stay quick */
function shrinkImage(blob) {
  return new Promise((res) => {
    const url = URL.createObjectURL(blob), img = new Image();
    img.onload = () => {
      const max = 1600, s = Math.min(1, max / Math.max(img.width, img.height));
      if (s === 1 && blob.size < 400000) { URL.revokeObjectURL(url); const r = new FileReader(); r.onload = () => res(r.result); r.readAsDataURL(blob); return; }
      const c = document.createElement("canvas"); c.width = Math.round(img.width * s); c.height = Math.round(img.height * s);
      c.getContext("2d").drawImage(img, 0, 0, c.width, c.height); URL.revokeObjectURL(url);
      res(c.toDataURL(blob.type === "image/png" && blob.size < 1500000 ? "image/png" : "image/jpeg", 0.85));
    };
    img.onerror = () => { URL.revokeObjectURL(url); const r = new FileReader(); r.onload = () => res(r.result); r.readAsDataURL(blob); };
    img.src = url;
  });
}

async function initEditor() {
  const phone = isPhone();
  await tinymce.init({
    selector: "#editor", license_key: "gpl",
    skin: isDark() ? "oxide-dark" : "oxide", content_css: ["default", FONT_CSS], content_style: PAPER_CSS,
    height: "100%", resize: false, branding: false, promotion: false,
    plugins: "advlist anchor autolink charmap code directionality fullscreen image insertdatetime link lists nonbreaking pagebreak quickbars searchreplace table visualblocks wordcount",
    menubar: phone ? false : "edit view insert format table tools",
    menu: { tools: { title: "Tools", items: "searchreplace wordcount code visualblocks" } },
    toolbar: "undo redo | blocks fontfamily fontsize | bold italic underline strikethrough | forecolor backcolor | alignleft aligncenter alignright alignjustify | bullist numlist todo | outdent indent lineheight | link image table | hr pagebreak charmap insertdatetime | ltr rtl | searchreplace removeformat",
    toolbar_mode: phone ? "scrolling" : "wrap", toolbar_sticky: true, statusbar: true, elementpath: false,
    font_family_formats: FONT_FORMATS,
    font_size_formats: "8pt 9pt 10pt 10.5pt 11pt 12pt 13pt 14pt 16pt 18pt 20pt 22pt 24pt 28pt 36pt 48pt 72pt",
    block_formats: "Normal text=p; Heading 1=h1; Heading 2=h2; Heading 3=h3; Heading 4=h4; Quote=blockquote; Code=pre",
    line_height_formats: "1 1.15 1.45 1.5 2 2.5 3",
    color_map: ["1B1F24","Black","5A6575","Grey","C0392B","Red","D2652A","Orange","B48A12","Gold","2E8B57","Green","1D5BB8","Blue","7A4FD6","Purple","FFF59D","Highlight yellow","C8F7C5","Highlight green","FFD6E0","Highlight pink","D6E6FF","Highlight blue","FFFFFF","White"],
    insertdatetime_formats: ["%d %B %Y", "%d/%m/%Y", "%H:%M", "%A %d %B %Y, %H:%M"],
    table_default_attributes: { border: "1" },
    table_default_styles: { "border-collapse": "collapse", width: "100%" },
    visual: false,
    quickbars_insert_toolbar: false,
    quickbars_selection_toolbar: phone ? false : "bold italic underline | forecolor backcolor | quicklink blockquote",
    contextmenu: "link image table",
    image_advtab: true, automatic_uploads: true, paste_data_images: true,
    images_upload_handler: (info) => shrinkImage(info.blob()),
    file_picker_types: "image",
    file_picker_callback: (cb) => {
      const inp = document.createElement("input"); inp.type = "file"; inp.accept = "image/*";
      inp.onchange = async () => { const f = inp.files[0]; if (f) cb(await shrinkImage(f), { alt: f.name.replace(/\.[^.]+$/, "") }); };
      inp.click();
    },
    link_default_target: "_blank", link_assume_external_targets: "https", browser_spellcheck: true, convert_urls: false,
    setup(ed) {
      S.ed = ed;
      ed.ui.registry.addToggleButton("todo", {
        icon: "checklist", tooltip: "To-do list", onAction: () => toggleTodo(ed),
        onSetup: (api) => { const h = () => api.setActive(!!ed.dom.getParent(ed.selection.getNode(), "ul.todo")); ed.on("NodeChange", h); return () => ed.off("NodeChange", h); },
      });
      ed.on("click", (e) => {
        const li = e.target.closest && e.target.closest("ul.todo > li");
        if (!li || e.target !== li) return;
        const r = li.getBoundingClientRect(), rtl = li.closest("[dir=rtl]");
        const x = rtl ? r.right - e.clientX : e.clientX - r.left;
        if (x <= 22) { ed.undoManager.transact(() => ed.dom.toggleClass(li, "done")); markDirty(); }
      });
      ed.on("input change undo redo Paste", () => markDirty());
      ed.on("ExecCommand", (e) => { if (!/^(mceFocus|SelectAll|mceRepaint|mceSearchReplace|mceVisualBlocks|mceFullScreen|mceWordCount|mceCodeEditor)$/i.test(e.command || "")) markDirty(); });
      ed.on("keyup", (e) => { if (e.key && e.key.length === 1) markDirty(); });
      ed.addShortcut("meta+s", "Save now", () => { clearTimeout(S.saveTimer); saveNow(); });
      ed.on("init", () => { S.edReady = true; applyLayout(); ed.mode.set("readonly"); resolveEditor(); });
    },
  });
}
/* page layout: full width (fills the window) or print layout (A4 sheet) */
const layout = () => lsGet("mbn2:layout") || "full";
function applyLayout() { const d = S.ed && S.ed.getDoc(); if (d) d.documentElement.classList.toggle("full", layout() === "full"); }
function toggleLayout() {
  lsSet("mbn2:layout", layout() === "full" ? "print" : "full"); applyLayout();
  toast(layout() === "full" ? "Full width: the page fills the window" : "Print layout: shows the page as an A4 sheet");
}
function toggleTodo(ed) {
  const ul = ed.dom.getParent(ed.selection.getNode(), "ul");
  if (ul && ed.dom.hasClass(ul, "todo")) { ed.undoManager.transact(() => { ed.dom.removeClass(ul, "todo"); ul.querySelectorAll("li.done").forEach((l) => ed.dom.removeClass(l, "done")); }); markDirty(); return; }
  ed.undoManager.transact(() => { if (!ul) ed.execCommand("InsertUnorderedList"); const u = ed.dom.getParent(ed.selection.getNode(), "ul"); if (u) ed.dom.addClass(u, "todo"); });
  markDirty();
}

/* ============================================================
   Wiring + start
   ============================================================ */
function pageMenu(anchor) {
  openMenu(anchor, [
    { label: "Save now", hint: "Ctrl+S", run: () => { clearTimeout(S.saveTimer); saveNow(); } },
    { label: "Move to another section", run: movePage },
    { label: "Duplicate page", run: duplicatePage },
    { label: "Find and replace", hint: "Ctrl+F", run: () => S.ed.execCommand("SearchReplace") },
    { label: layout() === "full" ? "Print layout (A4 page)" : "Full width page", run: toggleLayout },
    "-",
    { label: "Delete page", danger: true, run: deletePage },
  ]);
}
function accountMenu(anchor) {
  openMenu(anchor, [
    { label: "Open " + ROOT_NAME + " in Google Drive", run: () => window.open("https://drive.google.com/drive/folders/" + ROOT_ID, "_blank", "noopener") },
    { label: "Sync now", run: () => refreshTree().then(() => toast("Up to date with Google Drive")) },
    "-",
    { label: "Sign out", danger: true, run: async () => { await flushSave(); signOut(); } },
  ]);
}
function wire() {
  $("addSectionBtn").onclick = addSection;
  $("newPageBtn").onclick = newPage; $("emptyNewBtn").onclick = newPage;
  const openPicker = () => { if (!S.curSection) { toast("Create a section first."); return; } $("importInput").value = ""; $("importInput").click(); };
  $("openFileBtn").onclick = openPicker; $("emptyOpenBtn").onclick = openPicker;
  $("importInput").onchange = () => { const f = $("importInput").files[0]; if (f) importFile(f); };
  $("attachBtn").onclick = () => { $("attachInput").value = ""; $("attachInput").click(); };
  $("attachInput").onchange = () => { const f = $("attachInput").files[0]; if (f) attachFile(f); };
  $("exportBtn").onclick = (e) => { e.stopPropagation(); exportMenu($("exportBtn")); };
  $("pageMenuBtn").onclick = (e) => { e.stopPropagation(); pageMenu($("pageMenuBtn")); };
  $("accountBtn").onclick = (e) => { e.stopPropagation(); accountMenu($("accountBtn")); };
  $("gateBtn").onclick = () => signIn(false);
  $("reconnectBtn").onclick = () => signIn(false);
  $("backToSections").onclick = () => setView("sections");
  $("backToPages").onclick = async () => { await flushSave(); setView("pages"); };
  $("titleInput").addEventListener("input", () => { if (!S.cur) return; renderPages(); S.dirty = true; setSave("unsaved"); scheduleSave(); writeDraftSoon(); });
  $("titleInput").addEventListener("keydown", (e) => { if (e.key === "Enter") { e.preventDefault(); S.ed?.focus(); } });
  let st;
  $("searchInput").addEventListener("input", () => { clearTimeout(st); st = setTimeout(() => runSearch($("searchInput").value), 350); });
  for (const zone of [$("pagesPane"), $("noPage")]) {
    zone.addEventListener("dragover", (e) => { if (e.dataTransfer?.types?.includes("Files")) e.preventDefault(); });
    zone.addEventListener("drop", (e) => { const f = e.dataTransfer?.files?.[0]; if (f) { e.preventDefault(); importFile(f); } });
  }
  window.addEventListener("online", () => { setSync("", "Back online · syncing…"); refreshTree(); pushOfflineDrafts(); if (S.dirty) scheduleSave(300); });
  window.addEventListener("offline", () => setSync("warn", "Offline · changes are kept on this device"));
  window.addEventListener("focus", () => { if (S.started && Date.now() - S.treeAt > 5000) refreshTree(); });
  document.addEventListener("visibilitychange", () => {
    if (document.hidden) { if (S.dirty) { writeDraft(); saveNow(); } }
    else if (S.started) refreshTree();
  });
  window.addEventListener("beforeunload", (e) => { if (S.dirty) { writeDraft(); e.preventDefault(); e.returnValue = ""; } });
}

async function startNotebook() {
  if (S.started) return;
  S.started = true;
  $("gate").hidden = true;
  renderAccount();
  if (loadCachedTree()) { S.loaded = true; if (!S.curSection || !S.sections.has(S.curSection)) S.curSection = sortedSections()[0]?.id; renderSections(); renderPages(); }
  await refreshTree();
  pushOfflineDrafts();
  setInterval(() => { if (!document.hidden && navigator.onLine && tokenValid()) refreshTree(); }, POLL_MS);
  const last = lsGet("mbn2:lastPage");
  if (!isPhone() && last && S.pages.has(last)) openPage(last);
}

async function boot() {
  $("appVersion").textContent = "Version " + APP_VERSION;
  wire(); setSave("idle"); renderSections(); renderPages();
  if ("serviceWorker" in navigator && location.protocol === "https:") navigator.serviceWorker.register("sw.js").catch(() => {});
  initEditor().catch((e) => { console.error(e); toast("The editor couldn't load. Check your internet connection and reload.", 8000); });
  const ready = await initAuth();
  if (tokenValid()) startNotebook();
  else if (ready) showGate("signin");
  else if (lsGet("mbn2:tree") && !navigator.onLine) startNotebook();   // offline with a cached notebook
}
boot();
