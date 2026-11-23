'use strict';
/* Aufmaß-App Schreinerei Patrick Weiss
   - offline: alle Daten in IndexedDB auf dem Tablet
   - Nedo mEsstronic 2: sendet als Bluetooth-Tastatur in das aktive Feld
   - Bosch GLM (z. B. 150-27 C): Web Bluetooth, Messwert springt ins aktive Feld
   - Upload als OneNote-Seite in den Projekt-Abschnitt (Notizbuch "Kunden Termine") */

const TENANT_DEFAULT = 'adaf1c24-d253-4d9a-9d2a-20398cfda7b3';
const NOTEBOOK = 'Kunden Termine';
const SCOPES = 'Notes.ReadWrite User.Read offline_access';
const GRAPH = 'https://graph.microsoft.com/v1.0';

const TYPES = {
  raum:    { name: 'Raum',                  f: [['L', 'Länge'], ['B', 'Breite'], ['H', 'Höhe']] },
  wand:    { name: 'Wand',                  f: [['B', 'Breite'], ['H', 'Höhe']] },
  nische:  { name: 'Nische / Einbauschrank', f: [['Bo', 'Breite oben'], ['Bm', 'Breite Mitte'], ['Bu', 'Breite unten'], ['Hl', 'Höhe links'], ['Hm', 'Höhe Mitte'], ['Hr', 'Höhe rechts'], ['T', 'Tiefe']] },
  tuer:    { name: 'Innentür',              f: [['B', 'Rohbau Breite'], ['H', 'Rohbau Höhe ab OFF'], ['W', 'Wandstärke']], anschlag: true },
  fenster: { name: 'Fenster / Haustür',     f: [['B', 'Rohbau Breite'], ['H', 'Rohbau Höhe'], ['Br', 'Brüstungshöhe'], ['Lt', 'Laibungstiefe']], anschlag: true },
  frei:    { name: 'Freies Maß',            f: [['M1', 'Maß 1'], ['M2', 'Maß 2'], ['M3', 'Maß 3']] },
  skizze:  { name: 'Skizze / Grundriss',    f: [] },
};
// Bild = { bg: Foto oder null, shapes: [...], png: fertiges Bild }; alte Einträge waren nur Foto-Strings
const pic = p => typeof p === 'string' ? { bg: p, shapes: [], png: p } : p;
const skLines = p => pic(p).shapes.filter(s => s.t === 'line');
const skText = (it) => it.photos.map((p, k) => { const ls = skLines(p).filter(s => s.mm != null); return ls.length ? `Bild ${k + 1}: ` + ls.map(s => `${s.n} = ${fmtMM(s.mm)}`).join(', ') : ''; }).filter(Boolean);

// ---------- Hilfen ----------
const $ = s => document.querySelector(s);
const esc = s => String(s ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const uid = () => Date.now().toString(36) + Math.random().toString(36).slice(2, 7);
const fmtDate = t => new Date(t).toLocaleDateString('de-DE');
const fmtMM = v => (v == null || v === '' || isNaN(v)) ? '' : String(Math.round(v));
function toast(msg, ms = 2500) { const t = $('#toast'); t.textContent = msg; t.classList.remove('hidden'); clearTimeout(t._h); t._h = setTimeout(() => t.classList.add('hidden'), ms); }
function lsGet(k) { try { return localStorage.getItem(k); } catch { return null; } }
function lsSet(k, v) { try { v == null ? localStorage.removeItem(k) : localStorage.setItem(k, v); } catch {} }

// Text vom Messgerät/Tastatur -> mm
function parseMM(str, unit) {
  if (str == null) return null;
  let s = String(str).trim().toLowerCase();
  if (!s) return null;
  let u = unit || 'auto';
  if (/mm\s*$/.test(s)) u = 'mm'; else if (/cm\s*$/.test(s)) u = 'cm'; else if (/m\s*$/.test(s)) u = 'm';
  s = s.replace(/[^0-9,.\-]/g, '');
  // "1.234,5" -> 1234.5 ; "1,234" -> 1.234
  if (s.includes(',') && s.includes('.')) s = s.replace(/\./g, '').replace(',', '.'); else s = s.replace(',', '.');
  const n = parseFloat(s);
  if (isNaN(n)) return null;
  if (u === 'auto') u = (s.includes('.') && n < 30) ? 'm' : 'mm';
  return u === 'm' ? n * 1000 : u === 'cm' ? n * 10 : n;
}

// ---------- IndexedDB ----------
let _db;
function db() {
  if (_db) return Promise.resolve(_db);
  return new Promise((res, rej) => {
    const r = indexedDB.open('aufmass', 1);
    r.onupgradeneeded = () => { r.result.createObjectStore('aufmass', { keyPath: 'id' }); r.result.createObjectStore('kv'); };
    r.onsuccess = () => res(_db = r.result);
    r.onerror = () => rej(r.error);
  });
}
async function tx(store, mode, fn) {
  const d = await db();
  return new Promise((res, rej) => {
    const t = d.transaction(store, mode); const s = t.objectStore(store);
    const out = fn(s); t.oncomplete = () => res(out && 'result' in out ? out.result : out); t.onerror = () => rej(t.error);
  });
}
const A = {
  all: () => tx('aufmass', 'readonly', s => s.getAll()),
  get: id => tx('aufmass', 'readonly', s => s.get(id)),
  put: a => { a.updated = Date.now(); return tx('aufmass', 'readwrite', s => s.put(a)); },
  del: id => tx('aufmass', 'readwrite', s => s.delete(id)),
};
const KV = {
  get: k => tx('kv', 'readonly', s => s.get(k)),
  set: (k, v) => tx('kv', 'readwrite', s => s.put(v, k)),
};

// ---------- Einstellungen ----------
let settings = { name: '', clientId: '', tenantId: TENANT_DEFAULT, unit: 'auto' };
async function loadSettings() { Object.assign(settings, (await KV.get('settings')) || {}); }
async function saveSettings() { await KV.set('settings', settings); }

// ---------- Anmeldung Microsoft (PKCE, ohne Fremdbibliothek) ----------
const auth = { access: null, exp: 0 };
const b64url = buf => btoa(String.fromCharCode(...new Uint8Array(buf))).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');
const redirectUri = () => location.origin + location.pathname;
const authority = () => `https://login.microsoftonline.com/${settings.tenantId || TENANT_DEFAULT}/oauth2/v2.0`;

async function login() {
  if (!settings.clientId) { toast('Zuerst App-ID in den Einstellungen eintragen'); return; }
  if (!navigator.onLine) { toast('Anmelden geht nur mit Internet'); return; }
  const verifier = b64url(crypto.getRandomValues(new Uint8Array(48)));
  const challenge = b64url(await crypto.subtle.digest('SHA-256', new TextEncoder().encode(verifier)));
  const state = uid();
  lsSet('pkce', JSON.stringify({ verifier, state }));
  const p = new URLSearchParams({ client_id: settings.clientId, response_type: 'code', redirect_uri: redirectUri(), scope: SCOPES, code_challenge: challenge, code_challenge_method: 'S256', state, prompt: 'select_account' });
  location.href = `${authority()}/authorize?${p}`;
}
async function tokenRequest(body) {
  const r = await fetch(`${authority()}/token`, { method: 'POST', headers: { 'Content-Type': 'application/x-www-form-urlencoded' }, body: new URLSearchParams({ client_id: settings.clientId, scope: SCOPES, ...body }) });
  const j = await r.json();
  if (!r.ok) throw new Error(j.error_description || j.error || 'Anmeldung fehlgeschlagen');
  auth.access = j.access_token; auth.exp = Date.now() + (j.expires_in - 120) * 1000;
  if (j.refresh_token) await KV.set('refresh', j.refresh_token);
  return auth.access;
}
async function handleRedirect() {
  const q = new URLSearchParams(location.search);
  if (!q.has('code') && !q.has('error')) return;
  history.replaceState(null, '', redirectUri());
  if (q.has('error')) { toast('Anmeldung: ' + (q.get('error_description') || q.get('error')), 6000); return; }
  const pk = JSON.parse(lsGet('pkce') || '{}'); lsSet('pkce', null);
  if (pk.state !== q.get('state')) { toast('Anmeldung abgebrochen (state)'); return; }
  try { await tokenRequest({ grant_type: 'authorization_code', code: q.get('code'), redirect_uri: redirectUri(), code_verifier: pk.verifier }); toast('Angemeldet'); refreshProjects(); }
  catch (e) { toast(e.message, 6000); }
}
async function getToken() {
  if (auth.access && Date.now() < auth.exp) return auth.access;
  const rt = await KV.get('refresh');
  if (!rt) throw new Error('LOGIN');
  try { return await tokenRequest({ grant_type: 'refresh_token', refresh_token: rt }); }
  catch (e) { await KV.set('refresh', null); throw new Error('LOGIN'); }
}
async function graph(path, opt = {}) {
  const tok = await getToken();
  const r = await fetch(path.startsWith('http') ? path : GRAPH + path, { ...opt, headers: { Authorization: 'Bearer ' + tok, ...(opt.headers || {}) } });
  if (!r.ok) { let m = r.status + ''; try { m += ' ' + (await r.json()).error.message; } catch {} throw new Error('Graph ' + m); }
  return r.status === 204 ? null : r.json();
}

// ---------- Projekte = Abschnitte im Notizbuch "Kunden Termine" ----------
let projects = [];
async function loadProjects() { projects = (await KV.get('projects')) || []; }
async function refreshProjects(silent) {
  try {
    const nb = await graph(`/me/onenote/notebooks?$filter=displayName eq '${encodeURIComponent(NOTEBOOK)}'&$select=id,displayName`);
    if (!nb.value.length) throw new Error('Notizbuch "' + NOTEBOOK + '" nicht gefunden');
    let url = `/me/onenote/notebooks/${nb.value[0].id}/sections?$select=id,displayName`, out = [];
    while (url) { const r = await graph(url); out.push(...r.value); url = r['@odata.nextLink']; }
    const num = s => { const m = s.displayName.match(/(\d{5,})\s*$/); return m ? +m[1] : 0; };
    projects = out.map(s => ({ id: s.id, name: s.displayName })).sort((a, b) => num(b) - num(a) || a.name.localeCompare(b.name));
    await KV.set('projects', projects); await KV.set('projectsAt', Date.now());
    if (!silent) toast(projects.length + ' Projekte geladen');
  } catch (e) { if (!silent) toast(e.message === 'LOGIN' ? 'Bitte anmelden (Einstellungen)' : e.message, 5000); }
}

// ---------- Bosch GLM über Web Bluetooth ----------
// Protokoll nach Community-Analyse GLM 50 C / 100 C; am GLM 150-27 C mit der Diagnose prüfen.
const GLM_SVC = '02a6c0d0-0451-4000-b000-fb3210111989';
const GLM_CHR = '02a6c0d1-0451-4000-b000-fb3210111989';
const GLM_AUTOSYNC = [0xC0, 0x55, 0x02, 0x01, 0x00, 0x1A];
const laser = { dev: null, chr: null, log: [] };
function laserLog(s) { laser.log.push(new Date().toLocaleTimeString('de-DE') + '  ' + s); if (laser.log.length > 80) laser.log.shift(); const el = $('#laserlog'); if (el) el.textContent = laser.log.join('\n'); }
const hex = dv => [...new Uint8Array(dv.buffer, dv.byteOffset, dv.byteLength)].map(b => b.toString(16).padStart(2, '0')).join(' ');
function setLaserState(txt, cls) { const b = $('#btnLaser'); b.textContent = txt; b.className = 'pill ' + (cls || ''); }

async function laserConnect(all) {
  if (!navigator.bluetooth) { toast('Bluetooth geht nur in Chrome auf Android/Windows', 5000); return; }
  try {
    const opt = all ? { acceptAllDevices: true, optionalServices: [GLM_SVC] }
                    : { filters: [{ namePrefix: 'GLM' }, { namePrefix: 'Bosch' }, { services: [GLM_SVC] }], optionalServices: [GLM_SVC] };
    const dev = await navigator.bluetooth.requestDevice(opt);
    laser.dev = dev; laserLog('Gerät: ' + (dev.name || dev.id));
    dev.addEventListener('gattserverdisconnected', () => { laser.chr = null; setLaserState('Laser getrennt', 'busy'); laserLog('getrennt'); setTimeout(laserReconnect, 1500); });
    await laserAttach();
  } catch (e) { if (e.name !== 'NotFoundError') { toast('Laser: ' + e.message, 5000); laserLog('Fehler: ' + e.message); } }
}
async function laserAttach() {
  setLaserState('verbinde…', 'busy');
  const srv = await laser.dev.gatt.connect();
  let svc;
  try { svc = await srv.getPrimaryService(GLM_SVC); }
  catch (e) {
    laserLog('GLM-Dienst nicht gefunden. Sichtbare Dienste:');
    try { (await srv.getPrimaryServices()).forEach(s => laserLog('  ' + s.uuid)); } catch (x) { laserLog('  (keine Liste: ' + x.message + ')'); }
    setLaserState('Laser: unbekannt', 'busy'); return;
  }
  const chr = await svc.getCharacteristic(GLM_CHR);
  chr.addEventListener('characteristicvaluechanged', ev => onLaserData(ev.target.value));
  await chr.startNotifications();
  await chr.writeValue(new Uint8Array(GLM_AUTOSYNC));
  laser.chr = chr; laserLog('verbunden, Auto-Sync an');
  setLaserState((laser.dev.name || 'Laser') + ' ✓', 'on');
}
async function laserReconnect() {
  if (!laser.dev || laser.chr) return;
  try { await laserAttach(); } catch (e) { laserLog('Neuverbindung: ' + e.message); setLaserState('Laser verbinden'); }
}
function onLaserData(dv) {
  laserLog('RX ' + hex(dv));
  // Messpaket: C0 55 10 06 ... Entfernung als float32 (little endian, Meter) ab Byte 7
  if (dv.byteLength >= 11 && dv.getUint8(0) === 0xC0 && dv.getUint8(1) === 0x55 && dv.getUint8(2) === 0x10) {
    const m = dv.getFloat32(7, true);
    if (m > 0 && m < 300) { laserLog('Messung ' + (m * 1000).toFixed(1) + ' mm'); applyMeasurement(m * 1000); }
  }
}

// Messwert -> aktives Feld, danach nächstes leeres Feld
let lastField = null;
function applyMeasurement(mm) {
  const f = lastField && document.body.contains(lastField) ? lastField : null;
  if (!f) { toast('Messung ' + fmtMM(mm) + ' mm – erst ein Feld antippen', 3500); return; }
  f.value = fmtMM(mm); f.dispatchEvent(new Event('change')); f.classList.remove('flash'); void f.offsetWidth; f.classList.add('flash');
  if (!f.dataset.stay) focusNext(f);
}
function focusNext(f) {
  const all = [...document.querySelectorAll('input[data-m]')];
  const i = all.indexOf(f);
  const nxt = all.slice(i + 1).find(x => !x.value) || all[i + 1];
  if (nxt) { nxt.focus(); lastField = nxt; } else f.blur();
}

// ---------- Ansichten ----------
const nav = [];
function go(view, arg, push = true) { if (push && state.view) nav.push([state.view, state.arg]); state.view = view; state.arg = arg; render(); window.scrollTo(0, 0); }
function back() { const p = nav.pop(); if (p) { state.view = p[0]; state.arg = p[1]; render(); } }
const state = { view: null, arg: null };

async function render() {
  $('#btnBack').classList.toggle('hidden', !nav.length);
  const v = $('#view');
  if (state.view === 'list') return renderList(v);
  if (state.view === 'aufmass') return renderAufmass(v, await A.get(state.arg));
  if (state.view === 'item') return renderItem(v, await A.get(state.arg.a), state.arg.i);
  if (state.view === 'settings') return renderSettings(v);
  if (state.view === 'sketch') return renderSketch(v, await A.get(state.arg.a), state.arg.i, state.arg.k);
}

async function renderList(v) {
  $('#title').textContent = 'Aufmaß';
  const all = (await A.all()).sort((a, b) => b.updated - a.updated);
  const open = all.filter(a => a.status !== 'hochgeladen').length;
  v.innerHTML = `
    <div class="card">
      <div class="row"><button class="btn" id="new">+ Neues Aufmaß</button><button class="btn sec" id="sync">Hochladen (${open})</button><button class="btn sec" id="set">Einstellungen</button></div>
    </div>
    <div class="card list">${all.length ? all.map(a => `
      <div class="item" data-id="${a.id}"><div class="t"><b>${esc(projName(a) || 'ohne Projekt')}</b><div class="muted">${fmtDate(a.created)} · ${a.items.length} Positionen</div></div>
      <span class="badge ${a.status === 'hochgeladen' ? 'done' : 'open'}">${a.status === 'hochgeladen' ? 'in OneNote' : 'offen'}</span></div>`).join('') : '<div class="muted">Noch keine Aufmaße.</div>'}
    </div>`;
  $('#new').onclick = async () => { const a = { id: uid(), created: Date.now(), projekt: null, projektText: '', notiz: '', items: [], status: 'offen' }; await A.put(a); go('aufmass', a.id); };
  $('#sync').onclick = () => uploadAll(false);
  $('#set').onclick = () => go('settings');
  v.querySelectorAll('.item').forEach(el => el.onclick = () => go('aufmass', el.dataset.id));
}
const projName = a => a.projekt ? a.projekt.name : a.projektText;

function summary(it) {
  const t = TYPES[it.type];
  const parts = t.f.filter(([k]) => it.v[k] != null && it.v[k] !== '').map(([k, l]) => `${l} ${fmtMM(it.v[k])}`);
  if (it.anschlag) parts.push('DIN ' + it.anschlag);
  parts.push(...skText(it));
  return parts.join(' · ');
}

async function renderAufmass(v, a) {
  if (!a) return back();
  $('#title').textContent = projName(a) || 'Neues Aufmaß';
  const raeume = [...new Set(a.items.map(i => i.raum || ''))];
  const opts = projects.map(p => `<option value="${esc(p.id)}" ${a.projekt && a.projekt.id === p.id ? 'selected' : ''}>${esc(p.name)}</option>`).join('');
  v.innerHTML = `
    <div class="card">
      <h2>Projekt</h2>
      <input id="psearch" placeholder="Projekt suchen (Name oder Nummer)…">
      <select id="proj" size="1"><option value="">– Projekt wählen (${projects.length}) –</option>${opts}</select>
      <label>oder frei eintragen (wird dann nicht automatisch zugeordnet)</label>
      <input id="ptext" value="${esc(a.projektText)}" placeholder="z. B. Müller Küche">
      <label>Notiz zum Aufmaß</label><textarea id="anotiz">${esc(a.notiz)}</textarea>
      ${a.status === 'hochgeladen' ? `<p class="muted">In OneNote seit ${fmtDate(a.uploaded)}${a.pageUrl ? ` · <a href="${esc(a.pageUrl)}" target="_blank">öffnen</a>` : ''}. Änderungen → erneut hochladen erzeugt neue Seite.</p>` : ''}
    </div>
    ${raeume.map(r => `<div class="card"><h2>${esc(r || 'ohne Raum')}</h2><div class="list">${a.items.map((it, i) => (it.raum || '') !== r ? '' : `
      <div class="item" data-i="${i}"><div class="t"><b>${esc(it.label || TYPES[it.type].name)}</b><div class="muted">${esc(summary(it)) || 'noch keine Maße'}${it.photos.length ? ' · 📷 ' + it.photos.length : ''}</div></div>›</div>`).join('')}</div></div>`).join('')}
    <div class="card"><h2>Position hinzufügen</h2>
      <label>Raum</label><input id="raum" list="raeume" value="${esc(raeume[raeume.length - 1] || '')}" placeholder="z. B. Küche, Flur EG"><datalist id="raeume">${raeume.map(r => `<option value="${esc(r)}">`).join('')}</datalist>
      <div class="row" style="margin-top:10px">${Object.entries(TYPES).map(([k, t]) => `<button class="btn sec" data-t="${k}">${t.name}</button>`).join('')}</div>
    </div>
    <div class="card row"><button class="btn" id="up">In OneNote hochladen</button><button class="btn sec" id="share">Als Datei teilen</button><button class="btn del" id="del">Aufmaß löschen</button></div>`;
  const save = () => A.put(a);
  const sel = $('#proj');
  $('#psearch').oninput = e => { const q = e.target.value.toLowerCase(); [...sel.options].forEach((o, i) => { if (i) o.hidden = q && !o.text.toLowerCase().includes(q); }); const first = [...sel.options].find((o, i) => i && !o.hidden); if (q && first) { sel.value = first.value; sel.onchange(); } };
  sel.onchange = () => { const p = projects.find(x => x.id === sel.value); a.projekt = p ? { id: p.id, name: p.name } : null; $('#title').textContent = projName(a) || 'Neues Aufmaß'; save(); };
  $('#ptext').onchange = e => { a.projektText = e.target.value.trim(); save(); };
  $('#anotiz').onchange = e => { a.notiz = e.target.value; save(); };
  v.querySelectorAll('[data-t]').forEach(b => b.onclick = async () => {
    const t = b.dataset.t, raum = $('#raum').value.trim();
    const n = a.items.filter(i => i.type === t && (i.raum || '') === raum).length + 1;
    a.items.push({ id: uid(), type: t, raum, label: TYPES[t].name + (n > 1 ? ' ' + n : ''), v: {}, notiz: '', photos: [], anschlag: '' });
    await save(); go('item', { a: a.id, i: a.items.length - 1 });
    if (t === 'skizze') go('sketch', { a: a.id, i: a.items.length - 1, k: -1 });
  });
  v.querySelectorAll('.item[data-i]').forEach(el => el.onclick = () => go('item', { a: a.id, i: +el.dataset.i }));
  $('#up').onclick = () => uploadOne(a.id, true);
  $('#share').onclick = () => shareFile(a);
  $('#del').onclick = async () => { if (confirm('Aufmaß wirklich löschen? Das kann nicht rückgängig gemacht werden.')) { await A.del(a.id); back(); } };
}

async function renderItem(v, a, i) {
  const it = a && a.items[i];
  if (!it) return back();
  const t = TYPES[it.type];
  $('#title').textContent = (it.raum ? it.raum + ' – ' : '') + it.label;
  v.innerHTML = `
    <div class="card">
      <div class="row"><div><label>Bezeichnung</label><input id="lbl" value="${esc(it.label)}"></div><div><label>Raum</label><input id="raum" value="${esc(it.raum)}"></div></div>
    </div>
    <div class="card ${t.f.length ? '' : 'hidden'}"><h2>${t.name}</h2><p class="muted">Feld antippen, dann messen (Nedo oder Laser). Der Wert springt danach ins nächste Feld.</p>
      <div class="m">${t.f.map(([k, l]) => `<div class="f"><label>${l}</label><input data-m="${k}" inputmode="decimal" autocomplete="off" value="${fmtMM(it.v[k])}"><span class="u">mm</span></div>`).join('')}</div>
      ${t.anschlag ? `<label>Anschlag / Öffnungsrichtung</label><select id="ans"><option value="">–</option>${['links', 'rechts', 'zweiflügelig', 'Schiebetür'].map(o => `<option ${it.anschlag === o ? 'selected' : ''}>${o}</option>`).join('')}</select>` : ''}
    </div>
    <div class="card"><h2>Fotos & Skizzen</h2>
      <p class="muted">Bild antippen = darauf zeichnen und Maße per Messgerät eintragen.</p>
      <div class="row"><label class="btn sec" style="text-align:center;margin:0">📷 Foto<input id="foto" type="file" accept="image/*" capture="environment" hidden></label><button class="btn sec" id="skz">✏️ Neue Skizze</button></div>
      <div class="photos">${it.photos.map((p, k) => `<img src="${pic(p).png}" data-k="${k}">`).join('')}</div>
      ${skText(it).map(s => `<div class="muted">${esc(s)}</div>`).join('')}
      <label>Notiz</label>
      <textarea id="notiz" placeholder="Besonderheiten, Steckdosen, Leitungen, Schrägen…">${esc(it.notiz)}</textarea>
    </div>
    <div class="card row"><button class="btn" id="ok">Fertig</button><button class="btn sec" id="copy">Kopieren</button><button class="btn del" id="del">Löschen</button></div>`;
  const save = () => A.put(a);
  v.querySelectorAll('input[data-m]').forEach(inp => {
    inp.onfocus = () => { lastField = inp; };
    inp.onchange = () => { const mm = parseMM(inp.value, settings.unit); it.v[inp.dataset.m] = mm == null ? null : Math.round(mm); inp.value = fmtMM(it.v[inp.dataset.m]); save(); };
    inp.onkeydown = e => { if (e.key === 'Enter') { e.preventDefault(); inp.onchange(); focusNext(inp); } };
  });
  const first = v.querySelector('input[data-m]'); if (first && !first.value) { first.focus(); lastField = first; }
  $('#lbl').onchange = e => { it.label = e.target.value; save(); };
  $('#raum').onchange = e => { it.raum = e.target.value.trim(); save(); };
  $('#notiz').onchange = e => { it.notiz = e.target.value; save(); };
  if ($('#ans')) $('#ans').onchange = e => { it.anschlag = e.target.value; save(); };
  $('#foto').onchange = async e => { const f = e.target.files[0]; if (!f) return; const d = await shrink(f); it.photos.push({ bg: d, shapes: [], png: d }); await save(); render(); };
  $('#skz').onclick = () => go('sketch', { a: a.id, i, k: -1 });
  v.querySelectorAll('.photos img').forEach(img => img.onclick = () => go('sketch', { a: a.id, i, k: +img.dataset.k }));
  $('#ok').onclick = async () => { document.activeElement && document.activeElement.blur(); await save(); back(); };
  $('#copy').onclick = async () => { const c = JSON.parse(JSON.stringify(it)); c.id = uid(); c.label += ' (Kopie)'; c.photos = []; a.items.push(c); await save(); state.arg = { a: a.id, i: a.items.length - 1 }; render(); };
  $('#del').onclick = async () => { if (confirm('Position löschen?')) { a.items.splice(i, 1); await save(); back(); } };
}

// ---------- Skizzen-Editor ----------
// Werkzeuge: Maßlinie (ziehen → Wert per Laser/Nedo/Tastatur), Stift, Text. Linie antippen = auswählen.
const COLORS = ['#d32f2f', '#111111', '#1565c0'];
async function renderSketch(v, a, i, k) {
  const it = a && a.items[i];
  if (!it) return back();
  const src = k >= 0 && it.photos[k] ? pic(it.photos[k]) : { bg: null, shapes: [] };
  const sk = { bg: src.bg, shapes: JSON.parse(JSON.stringify(src.shapes)) };
  $('#title').textContent = (it.raum ? it.raum + ' – ' : '') + it.label + ' – Skizze';
  v.innerHTML = `
    <div class="sktools card">
      <div class="row">
        <button class="btn sec small" data-tool="line">📏 Maßlinie</button>
        <button class="btn sec small" data-tool="pen">✏️ Stift</button>
        <button class="btn sec small" data-tool="rect">▭ Rechteck</button>
        <button class="btn sec small" data-tool="text">T Text</button>
        ${COLORS.map(c => `<button class="btn sec small sw" data-c="${c}" style="background:${c}"></button>`).join('')}
        <button class="btn sec small" id="undo">↶</button>
      </div>
    </div>
    <div class="skwrap"><canvas id="cv"></canvas></div>
    <div class="card skbar">
      <div class="row">
        <div style="flex:2"><label id="selname">Maßlinie ziehen, dann messen</label>
          <div class="m" style="display:block"><div class="f"><input id="skm" data-stay="1" inputmode="decimal" autocomplete="off" placeholder="—"><span class="u">mm</span></div></div></div>
        <button class="btn del small" id="seldel" style="flex:0 0 auto">Linie löschen</button>
      </div>
      <div class="row" style="margin-top:10px"><button class="btn" id="ok">Speichern</button><button class="btn sec" id="cancel">Abbrechen</button>${k >= 0 ? '<button class="btn del" id="del">Bild löschen</button>' : ''}</div>
    </div>`;
  const cv = $('#cv'), ctx = cv.getContext('2d'), inp = $('#skm');
  let bgImg = null, tool = 'line', color = COLORS[0], sel = -1, cur = null;
  if (sk.bg) { bgImg = await new Promise(r => { const im = new Image(); im.onload = () => r(im); im.onerror = () => r(null); im.src = sk.bg; }); }
  cv.width = bgImg ? bgImg.width : 1600; cv.height = bgImg ? bgImg.height : 1100;
  const lw = Math.max(3, cv.width / 400), fs = Math.max(26, cv.width / 38);
  const setTool = t => { tool = t; v.querySelectorAll('[data-tool]').forEach(b => b.classList.toggle('on', b.dataset.tool === t)); };
  const setColor = c => { color = c; v.querySelectorAll('[data-c]').forEach(b => b.classList.toggle('on', b.dataset.c === c)); };
  setTool('line'); setColor(color);
  const nextLetter = () => { const n = sk.shapes.filter(s => s.t === 'line').length; return n < 26 ? String.fromCharCode(65 + n) : 'L' + (n + 1); };

  function label(s, hi) {
    const mx = (s.a[0] + s.b[0]) / 2, my = (s.a[1] + s.b[1]) / 2;
    let ang = Math.atan2(s.b[1] - s.a[1], s.b[0] - s.a[0]); if (ang > Math.PI / 2 || ang < -Math.PI / 2) ang += Math.PI;
    const txt = s.n + (s.mm != null ? ' = ' + fmtMM(s.mm) : ' = ?');
    ctx.save(); ctx.translate(mx, my); ctx.rotate(ang); ctx.font = `bold ${fs}px Arial`;
    const w = ctx.measureText(txt).width + fs * 0.6;
    ctx.fillStyle = hi ? '#80BD01' : 'rgba(255,255,255,.92)'; ctx.fillRect(-w / 2, -fs * 1.55, w, fs * 1.25);
    ctx.fillStyle = hi ? '#fff' : s.c; ctx.textAlign = 'center'; ctx.textBaseline = 'middle'; ctx.fillText(txt, 0, -fs * 0.92);
    ctx.restore();
  }
  function draw(final) {
    ctx.fillStyle = '#fff'; ctx.fillRect(0, 0, cv.width, cv.height);
    if (bgImg) ctx.drawImage(bgImg, 0, 0);
    else { ctx.strokeStyle = '#e6e9e0'; ctx.lineWidth = 1; for (let x = 0; x < cv.width; x += 50) { ctx.beginPath(); ctx.moveTo(x, 0); ctx.lineTo(x, cv.height); ctx.stroke(); } for (let y = 0; y < cv.height; y += 50) { ctx.beginPath(); ctx.moveTo(0, y); ctx.lineTo(cv.width, y); ctx.stroke(); } }
    const all = cur ? [...sk.shapes, cur] : sk.shapes;
    all.forEach((s, idx) => {
      const hi = !final && idx === sel;
      ctx.strokeStyle = hi ? '#80BD01' : s.c; ctx.lineWidth = hi ? lw * 1.8 : lw; ctx.lineCap = 'round'; ctx.lineJoin = 'round';
      if (s.t === 'pen') { ctx.beginPath(); s.pts.forEach((p, j) => j ? ctx.lineTo(p[0], p[1]) : ctx.moveTo(p[0], p[1])); ctx.stroke(); }
      else if (s.t === 'rect') { ctx.strokeRect(Math.min(s.a[0], s.b[0]), Math.min(s.a[1], s.b[1]), Math.abs(s.b[0] - s.a[0]), Math.abs(s.b[1] - s.a[1])); }
      else if (s.t === 'text') { ctx.font = `bold ${fs}px Arial`; ctx.fillStyle = s.c; ctx.textBaseline = 'middle'; ctx.fillText(s.s, s.p[0], s.p[1]); }
      else if (s.t === 'line') {
        ctx.beginPath(); ctx.moveTo(s.a[0], s.a[1]); ctx.lineTo(s.b[0], s.b[1]); ctx.stroke();
        const ang = Math.atan2(s.b[1] - s.a[1], s.b[0] - s.a[0]) + Math.PI / 2, t = fs * 0.6;
        [s.a, s.b].forEach(p => { ctx.beginPath(); ctx.moveTo(p[0] - Math.cos(ang) * t, p[1] - Math.sin(ang) * t); ctx.lineTo(p[0] + Math.cos(ang) * t, p[1] + Math.sin(ang) * t); ctx.stroke(); });
        label(s, hi);
      }
    });
  }
  function select(idx) {
    sel = idx; const s = sk.shapes[idx];
    if (s && s.t === 'line') { $('#selname').textContent = `Maß ${s.n} – jetzt messen oder eintippen`; inp.value = fmtMM(s.mm); lastField = inp; if (!laser.chr) inp.focus({ preventScroll: true }); }
    else { sel = -1; $('#selname').textContent = 'Maßlinie ziehen, dann messen'; inp.value = ''; lastField = null; }
    draw();
  }
  const pt = e => { const r = cv.getBoundingClientRect(); return [(e.clientX - r.left) * cv.width / r.width, (e.clientY - r.top) * cv.height / r.height]; };
  const distSeg = (p, a, b) => { const dx = b[0] - a[0], dy = b[1] - a[1], l = dx * dx + dy * dy || 1; let t = ((p[0] - a[0]) * dx + (p[1] - a[1]) * dy) / l; t = Math.max(0, Math.min(1, t)); return Math.hypot(p[0] - a[0] - t * dx, p[1] - a[1] - t * dy); };
  const hit = p => { let best = -1, bd = fs * 1.5; sk.shapes.forEach((s, idx) => { if (s.t === 'line') { const d = distSeg(p, s.a, s.b); if (d < bd) { bd = d; best = idx; } } }); return best; };

  cv.onpointerdown = e => {
    e.preventDefault(); cv.setPointerCapture(e.pointerId); const p = pt(e);
    if (tool === 'text') { const s = prompt('Text:'); if (s) { sk.shapes.push({ t: 'text', p, s, c: color }); draw(); } return; }
    if (tool === 'pen') cur = { t: 'pen', pts: [p], c: color };
    else cur = { t: tool, a: p, b: p, c: color, n: tool === 'line' ? nextLetter() : undefined, mm: null };
  };
  cv.onpointermove = e => { if (!cur) return; const p = pt(e); if (cur.t === 'pen') cur.pts.push(p); else cur.b = p; draw(); };
  cv.onpointerup = e => {
    if (!cur) return; const c = cur; cur = null;
    const len = c.t === 'pen' ? c.pts.length : Math.hypot(c.b[0] - c.a[0], c.b[1] - c.a[1]);
    if ((c.t === 'pen' && len < 3) || (c.t !== 'pen' && len < fs)) { select(hit(c.a)); return; } // Antippen = Linie wählen
    sk.shapes.push(c);
    if (c.t === 'line') select(sk.shapes.length - 1); else draw();
  };
  inp.onchange = () => { const s = sk.shapes[sel]; if (!s || s.t !== 'line') { inp.value = ''; return; } const mm = parseMM(inp.value, settings.unit); s.mm = mm == null ? null : Math.round(mm); inp.value = fmtMM(s.mm); draw(); };
  inp.onkeydown = e => { if (e.key === 'Enter') { e.preventDefault(); inp.onchange(); select(-1); } };
  v.querySelectorAll('[data-tool]').forEach(b => b.onclick = () => setTool(b.dataset.tool));
  v.querySelectorAll('[data-c]').forEach(b => b.onclick = () => setColor(b.dataset.c));
  $('#undo').onclick = () => { sk.shapes.pop(); select(-1); };
  $('#seldel').onclick = () => { if (sel >= 0) { sk.shapes.splice(sel, 1); select(-1); } };
  $('#cancel').onclick = () => back();
  $('#ok').onclick = async () => {
    inp.onchange(); draw(true);
    const out = { bg: sk.bg, shapes: sk.shapes, png: cv.toDataURL('image/jpeg', 0.85) };
    if (k >= 0) it.photos[k] = out; else it.photos.push(out);
    lastField = null; await A.put(a); back();
  };
  if ($('#del')) $('#del').onclick = async () => { if (confirm('Bild löschen?')) { it.photos.splice(k, 1); lastField = null; await A.put(a); back(); } };
  draw();
}

function shrink(file, max = 1600) {
  return new Promise((res, rej) => {
    const img = new Image(); img.onerror = rej;
    img.onload = () => { const s = Math.min(1, max / Math.max(img.width, img.height)); const c = document.createElement('canvas'); c.width = img.width * s; c.height = img.height * s; c.getContext('2d').drawImage(img, 0, 0, c.width, c.height); URL.revokeObjectURL(img.src); res(c.toDataURL('image/jpeg', 0.8)); };
    img.src = URL.createObjectURL(file);
  });
}

async function renderSettings(v) {
  $('#title').textContent = 'Einstellungen';
  const at = await KV.get('projectsAt'); const rt = await KV.get('refresh');
  v.innerHTML = `
    <div class="card"><h2>Allgemein</h2>
      <label>Aufmaß von (Name)</label><input id="sname" value="${esc(settings.name)}">
      <label>Einheit bei Tastatur-Eingabe / Nedo</label>
      <select id="sunit">${[['auto', 'automatisch (Komma unter 30 = Meter, sonst mm)'], ['mm', 'immer mm'], ['cm', 'immer cm'], ['m', 'immer Meter']].map(([k, l]) => `<option value="${k}" ${settings.unit === k ? 'selected' : ''}>${l}</option>`).join('')}</select>
    </div>
    <div class="card"><h2>OneNote</h2>
      <p class="muted">Status: ${rt ? 'angemeldet' : 'nicht angemeldet'} · Projekte: ${projects.length}${at ? ' (Stand ' + new Date(at).toLocaleString('de-DE') + ')' : ''}</p>
      <div class="row"><button class="btn" id="login">Anmelden</button><button class="btn sec" id="proj">Projekte aktualisieren</button><button class="btn sec" id="logout">Abmelden</button></div>
      <label>App-ID (Entra-App-Registrierung)</label><input id="scid" value="${esc(settings.clientId)}" placeholder="00000000-0000-0000-0000-000000000000">
      <label>Mandant (Tenant-ID)</label><input id="stid" value="${esc(settings.tenantId)}">
      <p class="muted">Rückleitungs-URL für die App-Registrierung (Plattform „Single-Page-Anwendung“): <b>${esc(redirectUri())}</b></p>
    </div>
    <div class="card"><h2>Laser-Diagnose</h2>
      <p class="muted">Falls der Bosch nicht gefunden wird: „Alle Geräte“ verbinden, messen, und das Protokoll an Claude schicken.</p>
      <div class="row"><button class="btn sec" id="lall">Alle Geräte anzeigen</button><button class="btn sec" id="ltest">Testwert 1234 mm</button></div>
      <pre class="log" id="laserlog">${esc(laser.log.join('\n')) || '–'}</pre>
    </div>
    <p class="muted">Version 1 · Schreinerei Patrick Weiss</p>`;
  $('#sname').onchange = e => { settings.name = e.target.value.trim(); saveSettings(); };
  $('#sunit').onchange = e => { settings.unit = e.target.value; saveSettings(); };
  $('#scid').onchange = e => { settings.clientId = e.target.value.trim(); saveSettings(); };
  $('#stid').onchange = e => { settings.tenantId = e.target.value.trim() || TENANT_DEFAULT; saveSettings(); };
  $('#login').onclick = login;
  $('#proj').onclick = async () => { await refreshProjects(); render(); };
  $('#logout').onclick = async () => { await KV.set('refresh', null); auth.access = null; render(); };
  $('#lall').onclick = () => laserConnect(true);
  $('#ltest').onclick = () => onLaserData(testPacket(1.234));
}
function testPacket(m) { const b = new DataView(new ArrayBuffer(16)); [0xC0, 0x55, 0x10, 0x06].forEach((x, i) => b.setUint8(i, x)); b.setFloat32(7, m, true); return b; }

// ---------- OneNote-Seite ----------
function pageHtml(a, withImgs) {
  const raeume = [...new Set(a.items.map(i => i.raum || ''))];
  const td = 'style="border:1px solid #bbb;padding:4px 6px;font-family:Arial;font-size:11pt;vertical-align:top"';
  const th = 'style="border:1px solid #bbb;padding:4px 6px;font-family:Arial;font-size:11pt;background-color:#80BD01;color:#ffffff"';
  let img = 0; const parts = [];
  const body = raeume.map(r => `
    <h2 style="font-family:Arial;color:#5f8c00">${esc(r || 'ohne Raum')}</h2>
    <table style="border-collapse:collapse"><tr><th ${th}>Position</th><th ${th}>Maße (mm)</th><th ${th}>Notiz</th></tr>
    ${a.items.filter(i => (i.raum || '') === r).map(it => `<tr><td ${td}><b>${esc(it.label)}</b></td><td ${td}>${TYPES[it.type].f.filter(([k]) => it.v[k] != null).map(([k, l]) => `${esc(l)}: <b>${fmtMM(it.v[k])}</b>`).join('<br/>')}${it.anschlag ? '<br/>Anschlag: ' + esc(it.anschlag) : ''}${skText(it).map(s => '<br/>' + esc(s)).join('')}</td><td ${td}>${esc(it.notiz).replace(/\n/g, '<br/>')}</td></tr>`).join('')}
    </table>
    ${a.items.filter(i => (i.raum || '') === r).flatMap(it => it.photos.map((ph, k) => { const p = pic(ph).png; const n = 'img' + img++; parts.push([n, p]); return `<p style="font-family:Arial;font-size:9pt">${esc(it.label)} – Bild ${k + 1}</p>` + (withImgs ? `<img src="${withImgs === 'inline' ? p : 'name:' + n}" width="480"/>` : ''); })).join('')}`).join('');
  const title = `Aufmaß ${fmtDate(a.created)}${settings.name ? ' – ' + settings.name : ''}`;
  const html = `<!DOCTYPE html><html><head><title>${esc(title)}</title><meta name="created" content="${new Date(a.created).toISOString()}"/></head><body>
    <p style="font-family:Arial;font-size:14pt;color:#80BD01"><b>Schreinerei Patrick Weiss – Aufmaß</b></p>
    <p style="font-family:Arial;font-size:11pt">Projekt: <b>${esc(projName(a) || '–')}</b><br/>Datum: ${fmtDate(a.created)}${settings.name ? '<br/>Aufgenommen von: ' + esc(settings.name) : ''}</p>
    ${a.notiz ? `<p style="font-family:Arial;font-size:11pt">${esc(a.notiz).replace(/\n/g, '<br/>')}</p>` : ''}
    ${body}</body></html>`;
  return { html, parts, title };
}
const dataUrlBlob = d => { const [h, b] = d.split(','); const bin = atob(b); const u = new Uint8Array(bin.length); for (let i = 0; i < bin.length; i++) u[i] = bin.charCodeAt(i); return new Blob([u], { type: h.match(/:(.*?);/)[1] }); };

async function uploadOne(id, interactive) {
  const a = await A.get(id);
  if (!a.projekt) { if (interactive) toast('Bitte ein Projekt aus der Liste wählen', 4000); return false; }
  if (!navigator.onLine) { if (interactive) toast('Kein Internet – bleibt gespeichert und wird später hochgeladen', 4000); return false; }
  try {
    const { html, parts } = pageHtml(a, 'parts');
    const fd = new FormData();
    fd.append('Presentation', new Blob([html], { type: 'text/html' }), 'page.html');
    parts.forEach(([n, d]) => fd.append(n, dataUrlBlob(d), n + '.jpg'));
    const r = await graph(`/me/onenote/sections/${a.projekt.id}/pages`, { method: 'POST', body: fd });
    a.status = 'hochgeladen'; a.uploaded = Date.now(); a.pageUrl = r.links && r.links.oneNoteWebUrl && r.links.oneNoteWebUrl.href;
    await A.put(a);
    if (interactive) { toast('In OneNote abgelegt: ' + a.projekt.name, 4000); render(); }
    return true;
  } catch (e) {
    if (interactive) toast(e.message === 'LOGIN' ? 'Bitte anmelden (Einstellungen)' : e.message, 6000);
    return false;
  }
}
async function uploadAll(silent) {
  const open = (await A.all()).filter(a => a.status !== 'hochgeladen' && a.projekt && a.items.length);
  if (!open.length) { if (!silent) toast('Nichts offen (Aufmaße ohne Projekt bleiben liegen)'); return; }
  let ok = 0; for (const a of open) if (await uploadOne(a.id, false)) ok++;
  if (!silent || ok) toast(`${ok} von ${open.length} hochgeladen` + (ok < open.length ? ' – ggf. anmelden' : ''), 4000);
  if (state.view === 'list') render();
}
async function shareFile(a) {
  const { html, title } = pageHtml(a, 'inline');
  const file = new File([html], title.replace(/[^\wäöüÄÖÜß .-]/g, '_') + '.html', { type: 'text/html' });
  if (navigator.canShare && navigator.canShare({ files: [file] })) { try { await navigator.share({ files: [file], title }); } catch {} }
  else { const u = URL.createObjectURL(file); const l = document.createElement('a'); l.href = u; l.download = file.name; l.click(); setTimeout(() => URL.revokeObjectURL(u), 5000); }
}

// ---------- Start ----------
function netState() { const on = navigator.onLine; const s = $('#stNet'); s.textContent = on ? 'online' : 'offline'; s.className = 'pill ' + (on ? 'on' : 'busy'); }
window.addEventListener('online', () => { netState(); uploadAll(true); refreshProjects(true); });
window.addEventListener('offline', netState);
$('#btnBack').onclick = back;
$('#btnLaser').onclick = () => laser.chr ? toast('Laser ist verbunden') : laserConnect(false);

(async () => {
  if ('serviceWorker' in navigator) navigator.serviceWorker.register('sw.js').catch(() => {});
  netState();
  await loadSettings(); await loadProjects();
  await handleRedirect();
  go('list', null, false);
  if (navigator.onLine) { uploadAll(true); refreshProjects(true); }
})();
