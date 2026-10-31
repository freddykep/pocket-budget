import { ApiError, displayAmount, errorMessage, formatEuro, parseAmount, parseSpeech, totalExpenses } from './util.js';

const $app = document.getElementById('app');
const $sheet = document.getElementById('sheet');
const $toast = document.getElementById('toast');

const state = { me: null, data: null, lastSync: 0, showAll: false, loading: false, loadError: null, homeActive: false };
const RECENT = 30;

// ---- Hilfsfunktionen -------------------------------------------------------------------------

// Baut DOM-Knoten ohne innerHTML: Texte aus Excel werden nie als HTML interpretiert.
function h(tag, props = {}, ...kids) {
  const el = document.createElement(tag);
  for (const [k, v] of Object.entries(props)) {
    if (v == null || v === false) continue;
    if (k === 'class') el.className = v;
    else if (k === 'text') el.textContent = v;
    else if (k.startsWith('on')) el.addEventListener(k.slice(2), v);
    else el.setAttribute(k, v === true ? '' : v);
  }
  for (const kid of kids.flat()) if (kid != null && kid !== false) el.append(kid);
  return el;
}

let toastTimer;
function toast(text) {
  $toast.textContent = text;
  $toast.classList.add('show');
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => $toast.classList.remove('show'), 3500);
}

async function api(path, body) {
  if (!navigator.onLine) throw new ApiError('offline');
  const ctl = new AbortController();
  const timer = setTimeout(() => ctl.abort(), 40_000);
  try {
    const res = await fetch(path, {
      ...(body === undefined ? {} : { method: 'POST', headers: { 'content-type': 'application/json' }, body: JSON.stringify(body) }),
      signal: ctl.signal,
      credentials: 'same-origin',
    });
    const data = await res.json().catch(() => ({}));
    if (!res.ok) throw new ApiError(data.error?.code || 'internal', data.error?.message, data.error);
    return data;
  } catch (e) {
    if (e instanceof ApiError) throw e;
    throw new ApiError(e.name === 'AbortError' ? 'timeout' : 'offline');
  } finally {
    clearTimeout(timer);
  }
}

const authExpired = (e) => e.code === 'not_signed_in' || e.code === 'reauth';
const time = (ms) => new Date(ms).toLocaleTimeString('de-DE', { hour: '2-digit', minute: '2-digit' });

function closeSheet() {
  if ($sheet.open) $sheet.close();
}
$sheet.addEventListener('click', (e) => {
  if (e.target === $sheet) closeSheet(); // Klick auf den abgedunkelten Hintergrund
});

function openSheet(...content) {
  $sheet.replaceChildren(h('div', { class: 'sheet' }, ...content));
  if (!$sheet.open) $sheet.showModal();
}

function goLogin(err) {
  state.me = null;
  state.data = null;
  closeSheet();
  renderLogin(errorMessage(err));
}

// ---- Anmeldung -------------------------------------------------------------------------------

function renderLogin(message, isInfo = false) {
  state.homeActive = false;
  $app.replaceChildren(
    h('main', { class: 'page' },
      h('div', { class: 'hero' },
        h('img', { src: '/icons/logo.svg', alt: '' }),
        h('h1', { text: 'Pocket Budget' }),
        h('p', { text: 'Ausgaben schnell erfassen – direkt in deiner Excel-Datei auf OneDrive.' }),
        message && (isInfo ? h('p', { class: 'note', text: message }) : h('div', { class: 'msg-error', role: 'alert', text: message })),
        h('a', { class: 'btn block', href: '/api/auth/login', text: 'Mit Microsoft OneDrive verbinden' }),
        h('p', { class: 'hint', text: 'Du meldest dich bei Microsoft an. Pocket Budget sieht dein Passwort nie.' }),
      )),
  );
}

// ---- Dateiauswahl ----------------------------------------------------------------------------

async function renderFilePicker(stack = []) {
  state.homeActive = false;
  const hasFile = !!state.me?.selectedFile;
  const folder = stack.at(-1)?.id || '';
  const top = h('div', { class: 'topbar' },
    stack.length || hasFile
      ? h('button', { class: 'icon-btn', 'aria-label': 'Zurück', onclick: () => (stack.length ? renderFilePicker(stack.slice(0, -1)) : bootHome()), text: '←' })
      : null,
    h('h1', { text: 'Excel-Datei wählen' }));
  const body = h('div', {}, h('div', { class: 'card skeleton' }));
  $app.replaceChildren(h('main', { class: 'page' }, top, body));

  try {
    const { items } = await api(`/api/files${folder ? `?folder=${encodeURIComponent(folder)}` : ''}`);
    const list = h('ul', { class: 'list' });
    for (const it of items) {
      const open = it.isFolder ? () => renderFilePicker([...stack, { id: it.id, name: it.name }]) : () => selectFile(it);
      list.append(h('li', {}, h('button', { class: 'item', onclick: open },
        h('div', { class: 'main' }, h('div', { class: 'title', text: `${it.isFolder ? '📁' : '📄'} ${it.name}` })),
        h('span', { class: 'amount', text: '›' }))));
    }
    body.replaceChildren(
      h('div', { class: 'crumbs', text: `OneDrive / ${stack.map((s) => s.name).join(' / ')}` }),
      items.length ? list : h('div', { class: 'card empty', text: 'Keine Ordner oder Excel-Dateien (.xlsx) hier.' }),
      h('p', { class: 'hint', text: 'Wähle die Excel-Datei, in der deine Ausgaben gespeichert werden.' }),
    );
  } catch (e) {
    if (authExpired(e)) return goLogin(e);
    body.replaceChildren(h('div', { class: 'card error-card', role: 'alert' }, h('p', { text: errorMessage(e) }),
      h('button', { class: 'btn small', onclick: () => renderFilePicker(stack), text: 'Erneut versuchen' })));
  }
}

async function selectFile(item) {
  try {
    await api('/api/file/select', { itemId: item.id });
    state.me = { ...state.me, selectedFile: { name: item.name } };
    state.data = null;
    toast(`Datei gewählt: ${item.name}`);
    bootHome();
  } catch (e) {
    if (authExpired(e)) return goLogin(e);
    toast(errorMessage(e));
  }
}

// ---- Startseite ------------------------------------------------------------------------------

function bootHome() {
  renderHome();
  return loadData();
}

function renderHome() {
  state.homeActive = true;
  $app.replaceChildren(
    h('main', { class: 'page' },
      h('div', { class: 'topbar' },
        h('h1', {}, h('img', { src: '/icons/logo.svg', alt: '' }), 'Pocket Budget'),
        h('button', { id: 'refresh', class: 'icon-btn', 'aria-label': 'Aktualisieren', onclick: () => loadData(), }, h('span', { text: '↻' })),
        h('button', { class: 'icon-btn', 'aria-label': 'Verbindung und Einstellungen', onclick: openSettings, text: '⚙' })),
      h('div', { id: 'content' })),
    h('div', { class: 'fab-bar' }, h('div', { class: 'fab-inner' },
      h('button', { class: 'btn grow', onclick: () => openEntrySheet(null), text: '＋ Ausgabe erfassen' }),
      h('button', { class: 'btn secondary icon', 'aria-label': 'Per Sprache erfassen', onclick: openVoiceSheet, text: '🎤' }))),
  );
  paint();
}

function paint() {
  const $content = document.getElementById('content');
  if (!$content) return;
  const nodes = [];

  if (state.loadError && !state.data) {
    nodes.push(h('div', { class: 'card error-card', role: 'alert' },
      h('p', { text: errorMessage(state.loadError) }),
      h('button', { class: 'btn small', onclick: () => loadData(), text: 'Erneut versuchen' })));
  } else if (!state.data) {
    nodes.push(h('div', { class: 'card skeleton' }));
  } else {
    const d = state.data;
    const stale = state.loadError ? h('div', { class: 'card error-card', role: 'alert' }, h('p', { text: `Aktualisierung fehlgeschlagen: ${errorMessage(state.loadError)}` })) : null;
    nodes.push(
      stale,
      h('div', { class: 'card summary' },
        h('div', { class: 'label', text: 'Gesamtsumme' }),
        h('div', { class: 'total', text: formatEuro(totalExpenses(d.summary.total)) }),
        h('div', { class: 'meta', text: `${d.summary.count} ${d.summary.count === 1 ? 'Ausgabe' : 'Ausgaben'} · Stand ${time(state.lastSync)}` })),
    );
    if (d.missingIds > 0) {
      nodes.push(h('div', { class: 'card banner' },
        h('p', { text: `${d.missingIds} ${d.missingIds === 1 ? 'Eintrag hat' : 'Einträge haben'} noch keine technische ID. Sie wird für Änderungen und den Abgleich gebraucht und in einer zusätzlichen Spalte der Excel-Datei abgelegt.` }),
        h('button', { class: 'btn small', onclick: () => openAssignIds(d.missingIds), text: 'IDs ergänzen' })));
    }
    nodes.push(h('div', { class: 'section-title', text: 'Zuletzt erfasst' }));
    if (!d.rows.length) {
      nodes.push(h('div', { class: 'card empty', text: 'Noch keine Ausgaben. Tippe unten auf „Ausgabe erfassen".' }));
    } else {
      const recent = [...d.rows].reverse();
      const visible = state.showAll ? recent : recent.slice(0, RECENT);
      const list = h('ul', { class: 'list' });
      for (const r of visible) {
        list.append(h('li', {}, h('button', { class: 'item', disabled: !r.id, onclick: () => openEntrySheet(r) },
          h('div', { class: 'main' },
            h('div', { class: 'title', text: r.ausgabe || '(ohne Text)' }),
            r.ort ? h('div', { class: 'sub', text: r.ort }) : null),
          h('span', { class: 'amount', text: displayAmount(r.betrag) }))));
      }
      nodes.push(list);
      if (recent.length > RECENT) {
        nodes.push(h('button', { class: 'link-btn', onclick: () => { state.showAll = !state.showAll; paint(); }, text: state.showAll ? 'Weniger anzeigen' : `Alle ${recent.length} anzeigen` }));
      }
    }
  }
  $content.replaceChildren(...nodes.filter(Boolean));
}

async function loadData() {
  if (state.loading) return;
  state.loading = true;
  state.loadError = null;
  document.getElementById('refresh')?.classList.add('spin');
  if (!state.data) paint();
  try {
    state.data = await api('/api/excel/read');
    state.lastSync = Date.now();
  } catch (e) {
    if (authExpired(e)) {
      state.loading = false;
      return goLogin(e);
    }
    state.loadError = e;
  } finally {
    state.loading = false;
    document.getElementById('refresh')?.classList.remove('spin');
  }
  paint();
}

// ---- Erfassen / Ändern -----------------------------------------------------------------------

function openEntrySheet(entry, prefill = null) {
  const isEdit = !!entry;
  // Eine ID pro geöffnetem Formular: Wiederholte Speicherversuche erzeugen keine Dublette.
  const id = isEdit ? entry.id : crypto.randomUUID();
  const values = {
    ausgabe: entry?.ausgabe ?? prefill?.ausgabe ?? '',
    ort: entry?.ort ?? prefill?.ort ?? '',
    betrag: entry ? String(Math.abs(entry.betrag)).replace('.', ',') : prefill?.betrag ? String(prefill.betrag).replace('.', ',') : '',
  };
  // Standard: Ausgabe (negativ). Nur die Checkbox macht daraus eine Rückerstattung (positiv).
  let refund = entry ? entry.betrag > 0 : false;
  showForm();

  function showForm(errors = {}) {
    const field = (key, label, attrs) => h('label', { class: `field${errors[key] ? ' invalid' : ''}` },
      h('span', { text: label }),
      h('input', { name: key, value: values[key], autocomplete: 'off', ...attrs }),
      errors[key] ? h('div', { class: 'err', role: 'alert', text: errors[key] }) : null);

    const form = h('form', { novalidate: true, onsubmit: (ev) => {
      ev.preventDefault();
      const fd = new FormData(ev.target);
      for (const k of Object.keys(values)) values[k] = String(fd.get(k) ?? '');
      refund = fd.has('refund');
      const errs = {};
      if (!values.ausgabe.trim()) errs.ausgabe = 'Bitte eine Beschreibung eingeben.';
      else if (values.ausgabe.length > 200) errs.ausgabe = 'Maximal 200 Zeichen.';
      if (values.ort.length > 100) errs.ort = 'Maximal 100 Zeichen.';
      const amount = parseAmount(values.betrag);
      if (Number.isNaN(amount)) errs.betrag = 'Bitte einen Betrag wie 84,50 eingeben (maximal 2 Nachkommastellen).';
      else if (amount <= 0) errs.betrag = 'Der Betrag muss größer als 0 sein.';
      if (Object.keys(errs).length) return showForm(errs);
      showReview(amount);
    } },
      field('ausgabe', 'Ausgabe', { required: true, autocapitalize: 'sentences', enterkeyhint: 'next', placeholder: 'z. B. Wocheneinkauf' }),
      field('ort', 'Ort / Händler (optional)', { autocapitalize: 'sentences', enterkeyhint: 'next', placeholder: 'z. B. REWE' }),
      field('betrag', 'Betrag in Euro', { inputmode: 'decimal', enterkeyhint: 'done', placeholder: '0,00' }),
      h('label', { class: 'check' },
        h('input', { type: 'checkbox', name: 'refund', checked: refund }),
        h('span', { text: 'Positiver Betrag (Rückerstattung)' })),
      h('div', { class: 'actions' },
        h('button', { type: 'button', class: 'btn ghost', onclick: closeSheet, text: 'Abbrechen' }),
        h('button', { type: 'submit', class: 'btn', text: 'Prüfen' })),
      isEdit ? h('div', { class: 'danger-zone' },
        h('button', { type: 'button', class: 'btn ghost block', onclick: showDelete, text: 'Eintrag löschen' })) : null);
    openSheet(
      h('h2', { text: isEdit ? 'Ausgabe ändern' : 'Ausgabe erfassen' }),
      prefill?.note ? h('p', { class: 'note', text: prefill.note }) : null,
      form,
    );
    form.querySelector(errors.ausgabe ? '[name=ausgabe]' : errors.ort ? '[name=ort]' : errors.betrag ? '[name=betrag]' : values.ausgabe ? '[name=betrag]' : '[name=ausgabe]')?.focus();
  }

  function showDelete() {
    const errBox = h('div', { class: 'msg-error', role: 'alert', hidden: true });
    let mode = 'delete';
    const delBtn = h('button', { class: 'btn danger', text: 'Endgültig löschen', onclick: async () => {
      if (mode === 'reload') { closeSheet(); await loadData(); return; }
      delBtn.disabled = backBtn.disabled = true;
      delBtn.textContent = 'Löscht…';
      errBox.hidden = true;
      try {
        await api('/api/excel/delete', { id, expectedHash: entry.hash });
        closeSheet();
        toast('Eintrag gelöscht');
        await loadData();
      } catch (e) {
        if (authExpired(e)) return goLogin(e);
        errBox.textContent = errorMessage(e);
        errBox.hidden = false;
        delBtn.disabled = backBtn.disabled = false;
        delBtn.textContent = 'Erneut versuchen';
        if (e.code === 'conflict') {
          mode = 'reload';
          delBtn.textContent = 'Aktuelle Daten laden';
        }
      }
    } });
    const backBtn = h('button', { class: 'btn ghost', onclick: () => showForm(), text: 'Zurück' });
    openSheet(
      h('h2', { text: 'Eintrag löschen?' }),
      h('dl', { class: 'review' },
        h('div', {}, h('dt', { text: 'Ausgabe' }), h('dd', { text: entry.ausgabe || '–' })),
        h('div', {}, h('dt', { text: 'Ort' }), h('dd', { text: entry.ort || '–' })),
        h('div', { class: 'big' }, h('dt', { text: 'Betrag' }), h('dd', { text: displayAmount(entry.betrag) }))),
      h('p', { class: 'note', text: 'Der Eintrag wird aus deiner Excel-Datei entfernt. Die Einträge darunter rücken eine Zeile nach oben.' }),
      errBox,
      h('div', { class: 'actions' }, backBtn, delBtn),
    );
  }

  function showReview(amount) {
    const errBox = h('div', { class: 'msg-error', role: 'alert', hidden: true });
    let mode = 'save';
    const saveBtn = h('button', { class: 'btn', text: 'Speichern', onclick: async () => { if (mode === 'save') return save(); closeSheet(); await loadData(); } });
    const backBtn = h('button', { class: 'btn ghost', onclick: () => showForm(), text: 'Zurück' });

    async function save() {
      saveBtn.disabled = backBtn.disabled = true;
      saveBtn.textContent = 'Speichert…';
      errBox.hidden = true;
      const payload = { ausgabe: values.ausgabe.trim(), ort: values.ort.trim(), betrag: amount, refund };
      try {
        if (isEdit) await api('/api/excel/update', { id, expectedHash: entry.hash, values: payload });
        else await api('/api/excel/append', { id, ...payload });
        closeSheet();
        toast('In Excel gespeichert');
        await loadData();
        // Kontrolle: der Eintrag muss nach dem erneuten Einlesen wirklich in der Datei stehen.
        if (!isEdit && state.data && !state.data.rows.some((r) => r.id === id)) {
          toast('Achtung: Der Eintrag wurde in der Datei nicht gefunden. Bitte in Excel prüfen.');
        }
      } catch (e) {
        if (authExpired(e)) return goLogin(e);
        let text = errorMessage(e);
        if (!isEdit && (e.code === 'offline' || e.code === 'timeout')) {
          text += ' Falls unklar ist, ob gespeichert wurde: Erneut speichern ist sicher, es entsteht kein doppelter Eintrag.';
        }
        errBox.textContent = text;
        errBox.hidden = false;
        saveBtn.disabled = backBtn.disabled = false;
        saveBtn.textContent = 'Erneut speichern';
        if (e.code === 'conflict' || e.code === 'row_not_found') {
          mode = 'reload';
          saveBtn.textContent = 'Aktuelle Daten laden';
        }
      }
    }

    openSheet(
      h('h2', { text: 'Bitte prüfen' }),
      h('dl', { class: 'review' },
        h('div', {}, h('dt', { text: 'Ausgabe' }), h('dd', { text: values.ausgabe.trim() })),
        h('div', {}, h('dt', { text: 'Ort' }), h('dd', { text: values.ort.trim() || '–' })),
        h('div', { class: 'big' }, h('dt', { text: 'Betrag' }), h('dd', { text: refund ? `+${formatEuro(amount)}` : formatEuro(amount) }))),
      h('p', { class: 'note', text: refund
        ? 'Wird als Rückerstattung (positiver Betrag) in deiner Excel-Datei gespeichert.'
        : 'Wird als Ausgabe (negativer Betrag) in deiner Excel-Datei gespeichert.' }),
      errBox,
      h('div', { class: 'actions' }, backBtn, saveBtn),
    );
  }
}

// ---- Spracheingabe --------------------------------------------------------------------------

const speechApi = () => window.SpeechRecognition || window.webkitSpeechRecognition;

const SPEECH_ERRORS = {
  'not-allowed': 'Der Mikrofonzugriff wurde verweigert. Bitte erlaube das Mikrofon in den Einstellungen deines Browsers.',
  'service-not-allowed': 'Die Spracherkennung ist in diesem Browser nicht erlaubt.',
  'no-speech': 'Ich habe nichts gehört. Bitte noch einmal versuchen und deutlich sprechen.',
  'audio-capture': 'Es wurde kein Mikrofon gefunden.',
  network: 'Die Spracherkennung braucht eine Internetverbindung.',
  'language-not-supported': 'Deutsch wird von der Spracherkennung dieses Browsers nicht unterstützt.',
};

function openVoiceSheet() {
  const SpeechRecognition = speechApi();
  if (!SpeechRecognition) {
    return openSheet(
      h('h2', { text: 'Spracheingabe' }),
      h('p', { text: 'Dieser Browser unterstützt keine Spracheingabe. Am besten funktioniert sie in Chrome, Edge und Safari. Du kannst die Ausgabe stattdessen von Hand erfassen.' }),
      h('div', { class: 'actions' },
        h('button', { class: 'btn ghost', onclick: closeSheet, text: 'Schließen' }),
        h('button', { class: 'btn', onclick: () => openEntrySheet(null), text: 'Von Hand erfassen' })),
    );
  }

  const live = h('p', { class: 'transcript', 'aria-live': 'polite', text: 'Tippe auf „Aufnahme starten" und sage zum Beispiel: „Wocheneinkauf bei Rewe 84 Euro 50".' });
  const errBox = h('div', { class: 'msg-error', role: 'alert', hidden: true });
  const start = h('button', { class: 'btn', text: '🎤 Aufnahme starten', onclick: () => listen() });
  let rec = null;
  let transcript = '';

  function listen() {
    errBox.hidden = true;
    transcript = '';
    rec = new SpeechRecognition();
    rec.lang = 'de-DE';
    rec.interimResults = true;
    rec.continuous = false;
    rec.maxAlternatives = 1;
    let failed = false;
    rec.onstart = () => { start.disabled = true; start.textContent = 'Ich höre zu…'; live.textContent = '…'; };
    rec.onresult = (ev) => {
      transcript = [...ev.results].map((r) => r[0].transcript).join(' ').trim();
      live.textContent = transcript || '…';
    };
    rec.onerror = (ev) => {
      failed = true;
      errBox.textContent = SPEECH_ERRORS[ev.error] || `Spracherkennung fehlgeschlagen (${ev.error}).`;
      errBox.hidden = false;
    };
    rec.onend = () => {
      start.disabled = false;
      start.textContent = '🎤 Erneut aufnehmen';
      if (failed || !transcript) {
        if (!failed) { errBox.textContent = SPEECH_ERRORS['no-speech']; errBox.hidden = false; }
        return;
      }
      const r = parseSpeech(transcript);
      closeSheet();
      // Nur ein Vorschlag: Der Benutzer prüft und bestätigt in den nächsten Schritten.
      openEntrySheet(null, { ...r, note: `Erkannt: „${r.transcript}". Bitte prüfen und bei Bedarf korrigieren.` });
    };
    try {
      rec.start();
    } catch {
      errBox.textContent = 'Die Aufnahme konnte nicht gestartet werden.';
      errBox.hidden = false;
    }
  }

  openSheet(
    h('h2', { text: 'Spracheingabe' }),
    live,
    errBox,
    h('p', { class: 'note', text: 'Die Spracherkennung läuft über deinen Browser (Chrome: Google, Safari: Apple). Dabei wird die Aufnahme an den Browser-Hersteller gesendet. Pocket Budget speichert keine Aufnahmen und speichert nichts, bevor du es bestätigst.' }),
    h('div', { class: 'actions' }, h('button', { class: 'btn ghost', onclick: () => { rec?.abort(); closeSheet(); }, text: 'Abbrechen' }), start),
  );
}

// ---- IDs ergänzen ----------------------------------------------------------------------------

function openAssignIds(count, then) {
  const errBox = h('div', { class: 'msg-error', role: 'alert', hidden: true });
  const go = h('button', { class: 'btn', text: 'IDs ergänzen', onclick: async () => {
    go.disabled = true;
    go.textContent = 'Ergänze…';
    try {
      const r = await api('/api/excel/assign-ids', {});
      closeSheet();
      toast(`${r.assigned} ${r.assigned === 1 ? 'ID' : 'IDs'} ergänzt`);
      await loadData();
      then?.();
    } catch (e) {
      if (authExpired(e)) return goLogin(e);
      errBox.textContent = errorMessage(e);
      errBox.hidden = false;
      go.disabled = false;
      go.textContent = 'Erneut versuchen';
    }
  } });
  openSheet(
    h('h2', { text: 'IDs ergänzen' }),
    h('p', { text: `Für ${count} ${count === 1 ? 'Eintrag' : 'Einträge'} wird in der Spalte „ID" der Excel-Datei eine technische Kennung ergänzt. Vorhandene IDs und alle anderen Inhalte bleiben unverändert.` }),
    errBox,
    h('div', { class: 'actions' }, h('button', { class: 'btn ghost', onclick: closeSheet, text: 'Abbrechen' }), go),
  );
}

// ---- Excel leeren ----------------------------------------------------------------------------

async function openClearSheet() {
  openSheet(h('h2', { text: 'Excel leeren' }), h('p', { text: 'Lade aktuellen Stand…' }));
  let d;
  try {
    d = await api('/api/excel/read');
    state.data = d;
    state.lastSync = Date.now();
    paint();
  } catch (e) {
    if (authExpired(e)) return goLogin(e);
    return openSheet(h('h2', { text: 'Excel leeren' }), h('div', { class: 'msg-error', role: 'alert', text: errorMessage(e) }),
      h('div', { class: 'actions' }, h('button', { class: 'btn ghost', onclick: closeSheet, text: 'Schließen' })));
  }

  if (!d.rows.length) {
    return openSheet(h('h2', { text: 'Excel leeren' }), h('p', { text: 'Die Liste ist bereits leer.' }),
      h('div', { class: 'actions' }, h('button', { class: 'btn', onclick: closeSheet, text: 'OK' })));
  }
  if (d.missingIds > 0) {
    return openSheet(h('h2', { text: 'Excel leeren' }),
      h('p', { text: `${d.missingIds} ${d.missingIds === 1 ? 'Eintrag hat' : 'Einträge haben'} noch keine ID. Zum Leeren muss die Datei eindeutig als Erfassungsliste erkannt werden.` }),
      h('div', { class: 'actions' },
        h('button', { class: 'btn ghost', onclick: closeSheet, text: 'Abbrechen' }),
        h('button', { class: 'btn', onclick: () => openAssignIds(d.missingIds, openClearSheet), text: 'IDs ergänzen' })));
  }

  const errBox = h('div', { class: 'msg-error', role: 'alert', hidden: true });
  const go = h('button', { class: 'btn danger', disabled: true, text: 'Excel jetzt leeren' });
  const check = h('input', { type: 'checkbox', id: 'confirm-clear', onchange: (ev) => { go.disabled = !ev.target.checked; } });
  let reload = false;
  go.addEventListener('click', async () => {
    if (reload) return openClearSheet();
    go.disabled = true;
    go.textContent = 'Leere…';
    try {
      const r = await api('/api/excel/clear', { expectedListHash: d.listHash });
      closeSheet();
      toast(`Excel geleert (${r.cleared} ${r.cleared === 1 ? 'Eintrag' : 'Einträge'})`);
      await loadData();
    } catch (e) {
      if (authExpired(e)) return goLogin(e);
      errBox.textContent = errorMessage(e);
      errBox.hidden = false;
      go.textContent = 'Excel jetzt leeren';
      go.disabled = !check.checked;
      if (e.code === 'conflict') {
        reload = true;
        go.textContent = 'Neu laden';
        go.disabled = false;
      }
    }
  });

  openSheet(
    h('h2', { text: 'Excel leeren' }),
    h('dl', { class: 'review' },
      h('div', {}, h('dt', { text: 'Einträge' }), h('dd', { text: String(d.summary.count) })),
      h('div', { class: 'big' }, h('dt', { text: 'Gesamtsumme' }), h('dd', { text: formatEuro(totalExpenses(d.summary.total)) }))),
    h('p', { class: 'note', text: 'Entfernt alle Einträge aus der Excel-Datei. Überschriften, Formatierung und verbundene Zellen bleiben erhalten. Zurückholen geht nur über den Versionsverlauf in OneDrive.' }),
    h('label', { class: 'item' }, check, h('span', { text: 'Ich habe die Ausgaben in meine Haushalts-Excel übernommen.' })),
    errBox,
    h('div', { class: 'actions' }, h('button', { class: 'btn ghost', onclick: closeSheet, text: 'Abbrechen' }), go),
  );
}

// ---- Verbindung / Einstellungen --------------------------------------------------------------

function openSettings() {
  const me = state.me;
  openSheet(
    h('h2', { text: 'Verbindung' }),
    h('dl', { class: 'kv' },
      h('div', {}, h('dt', { text: 'Microsoft-Konto' }), h('dd', { text: `${me.displayName ?? ''} (${me.account ?? ''})` })),
      h('div', {}, h('dt', { text: 'Excel-Datei' }), h('dd', { text: me.selectedFile?.name ?? '–' }))),
    h('div', { class: 'stack' },
      h('button', { class: 'btn secondary', onclick: () => { closeSheet(); renderFilePicker(); }, text: 'Andere Datei wählen' }),
      h('button', { class: 'btn ghost', onclick: openDisconnect, text: 'Verbindung trennen' })),
    h('div', { class: 'danger-zone' },
      h('h3', { text: 'Excel leeren' }),
      h('p', { text: 'Entfernt alle erfassten Ausgaben aus der Excel-Datei, zum Beispiel nach der Übernahme in die Haushalts-Excel.' }),
      h('button', { class: 'btn ghost block', onclick: openClearSheet, text: 'Excel leeren…' })),
    h('div', { class: 'actions' }, h('button', { class: 'btn ghost', onclick: closeSheet, text: 'Schließen' })),
  );
}

function openDisconnect() {
  const go = h('button', { class: 'btn danger', text: 'Trennen', onclick: async () => {
    go.disabled = true;
    try {
      await api('/api/auth/logout', {});
      closeSheet();
      state.me = null;
      state.data = null;
      renderLogin('Verbindung getrennt. Deine Excel-Datei bleibt unverändert.', true);
    } catch (e) {
      go.disabled = false;
      toast(errorMessage(e));
    }
  } });
  openSheet(
    h('h2', { text: 'Verbindung trennen?' }),
    h('p', { text: 'Pocket Budget löscht die gespeicherte Verbindung zu deinem OneDrive. Die Excel-Datei selbst bleibt unverändert. Zum erneuten Arbeiten musst du dich neu verbinden.' }),
    h('div', { class: 'actions' }, h('button', { class: 'btn ghost', onclick: closeSheet, text: 'Abbrechen' }), go),
  );
}

// ---- Start -----------------------------------------------------------------------------------

async function boot() {
  const params = new URLSearchParams(location.search);
  const urlError = params.get('error');
  if (urlError) history.replaceState(null, '', '/');
  try {
    state.me = await api('/api/me');
  } catch (e) {
    if (e.code === 'not_signed_in') return renderLogin(urlError);
    if (e.code === 'reauth') return renderLogin(errorMessage(e));
    return renderLogin(urlError || errorMessage(e));
  }
  if (!state.me.selectedFile) return renderFilePicker();
  bootHome();
}

// Beim Zurückkehren in die App den aktuellen Excel-Stand holen (Änderungen aus Excel werden sichtbar).
document.addEventListener('visibilitychange', () => {
  if (!document.hidden && state.homeActive && !$sheet.open && Date.now() - state.lastSync > 30_000) loadData();
});
window.addEventListener('online', () => {
  if (state.homeActive && (state.loadError || !state.data)) loadData();
});

if ('serviceWorker' in navigator) navigator.serviceWorker.register('/sw.js').catch(() => {});
boot();
