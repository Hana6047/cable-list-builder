'use strict';
/* =============================================================================
   Cable List Builder – IO LIST -> CABLE LIST (feuille EQUIPMENT)
   Port JavaScript de generate_cable_list.py (même logique, même résultat).

   Règles :
   - chaque loop démarre au FACP et retourne au FACP ;
   - l'ordre des appareils = l'ordre des lignes de l'IO LIST ;
   - un appareil sans adresse de boucle (FD, PSH, SOV, LS, CT, CAM...) est câblé
     en dérivation depuis le dernier appareil de boucle qui le précède ;
     la boucle continue ensuite depuis ce module, pas depuis la dérivation.
   ============================================================================= */

/* ---------- Paramètres par défaut (modifiables dans la page) ---------- */
const DEFAULTS = {
  cableType: '2x1,5mmq',
  shield: 'SH',
  armored: '-',
  facpName: 'FACP',
  facpDescription: '',
  removeText: ['-PORTA TAGLIAFUOCO'],
  hubPrefixes: ['CCP'],
  addressMode: 'io',            // 'io' = adresse de l'IO LIST | 'type' = table TYPE_ADDRESS
};

const TYPE_ADDRESS = {
  MCP: '004', MCPO: '017', ML: '017', ALM: '013', MIO: '014',
  SD: '015', ASD: '003', SOA: '002',
};

const HEADER_ALIASES = {
  loop:    ['LOOP'],
  prefix:  ['PREFISSO', 'PREFIX'],
  number:  ["NUMERO DELL'APPARECCHIATURA", 'NUMERO', 'NUMBER'],
  address: ['INDIRIZZO', 'ADDRESS'],
  desc:    ['DESCRIZIONE', 'DESCRIPTION'],
  area:    ['AREA', 'ZONE'],
};

/* =============================================================================
   1. LECTURE DE L'IO LIST (logique pure, sans DOM)
   ============================================================================= */
const norm = (v) => (v == null ? '' : String(v).replace(/\s+/g, ' ').trim().toUpperCase());

function clean(v) {
  const s = v == null ? '' : String(v).trim();
  return s === '-' || s === '--' ? '' : s;
}

const escapeRe = (s) => s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');

/** Valeur ExcelJS (texte, nombre, formule, rich text...) -> valeur simple. */
function cellText(v) {
  if (v == null) return null;
  if (v instanceof Date) return v.toISOString();
  if (typeof v === 'object') {
    if (v.error) return null;
    if ('result' in v) return cellText(v.result);
    if (v.richText) return v.richText.map((t) => t.text).join('');
    if (v.text !== undefined) return cellText(v.text);
    return String(v);
  }
  return v;
}

/** Cherche la ligne d'en-têtes et les colonnes utiles dans une matrice de cellules. */
function findHeader(matrix) {
  const limit = Math.min(matrix.length, 15);
  for (let r = 0; r < limit; r++) {
    const heads = (matrix[r] || []).map(norm);
    const cols = {};
    for (const [field, aliases] of Object.entries(HEADER_ALIASES)) {
      let idx = -1;
      for (const a of aliases) { idx = heads.indexOf(a); if (idx >= 0) break; }          // exact
      if (idx < 0) for (const a of aliases) { idx = heads.findIndex((h) => h.startsWith(a)); if (idx >= 0) break; }
      if (idx >= 0) cols[field] = idx;
    }
    if ('loop' in cols && 'prefix' in cols && 'number' in cols) return { row: r, cols };
  }
  return null;
}

/** matrix = tableau de lignes (tableaux de valeurs). Retourne les appareils ou null. */
function parseIoMatrix(matrix, opts) {
  const hdr = findHeader(matrix);
  if (!hdr) return null;
  const devices = [];
  for (let r = hdr.row + 1; r < matrix.length; r++) {
    const row = matrix[r] || [];
    const get = (f) => (f in hdr.cols ? row[hdr.cols[f]] : null);
    const prefix = clean(get('prefix'));
    if (!prefix) continue;
    const rawNum = get('number');
    const num = typeof rawNum === 'number' ? String(Math.trunc(rawNum)).padStart(2, '0') : clean(rawNum);
    let desc = clean(get('desc'));
    for (const rm of opts.removeText) {
      if (rm) desc = desc.replace(new RegExp(escapeRe(rm), 'gi'), '');
    }
    desc = desc.replace(/\s+/g, ' ').trim().replace(/^[\s-]+|[\s-]+$/g, '');
    devices.push({
      srcRow: r + 1, loop: clean(get('loop')), prefix, num,
      addr: clean(get('address')), desc, area: clean(get('area')),
    });
  }
  return devices;
}

/* =============================================================================
   2. CONSTRUCTION DES BOUCLES  (FACP -> ... -> FACP)
   ============================================================================= */
function buildRows(devices, opts) {
  const warnings = [];
  const loops = new Map();
  let current = null;

  for (const d of devices) {
    if (d.loop) current = d.loop;
    if (current === null) {
      warnings.push(`IO LIST ligne ${d.srcRow} (${d.prefix}-${d.num}) : aucun loop avant cet appareil, ignoré`);
      continue;
    }
    d.onLoop = Boolean(d.loop && d.addr);       // appareil adressé sur la boucle
    if (!loops.has(current)) loops.set(current, []);
    loops.get(current).push(d);
  }

  const rows = [];
  const hubs = new Set(opts.hubPrefixes);
  const nextRow = () => rows.length + 2;        // n° de ligne Excel de la prochaine ligne

  for (const [loop, devs] of loops) {
    let prevRow = null;                         // dernier appareil de boucle
    let hubRow = null;                          // concentrateur actif (CCP)
    let lastArea = devs[0].area;
    const tagCount = new Map();

    for (const d of devs) {
      const tag = `${d.prefix}-${d.num}`;
      tagCount.set(tag, (tagCount.get(tag) || 0) + 1);
      const r = nextRow();
      let srcRow;
      if (d.onLoop) {
        srcRow = prevRow;                       // null => FACP
        prevRow = r; hubRow = null;
      } else {
        const parent = hubRow || prevRow;
        if (parent === null) {
          warnings.push(`${loop} ${tag} : dérivation sans appareil précédent, câblée au ${opts.facpName}`);
          srcRow = null;
        } else {
          srcRow = parent;
        }
        if (hubs.has(d.prefix)) hubRow = r;
      }
      if (d.desc === '') warnings.push(`${loop} ${tag} : description vide dans l'IO LIST`);
      lastArea = d.area || lastArea;
      rows.push({
        tag, loop: d.loop, kind: 'device', onLoop: d.onLoop, srcRow,
        addr: opts.addressMode === 'type' ? (TYPE_ADDRESS[d.prefix] || '') : d.addr,
        desc: d.desc, area: d.area,
      });
    }
    for (const [t, n] of tagCount) {
      if (n > 1) warnings.push(`${loop} : tag « ${t} » présent ${n} fois, étiquettes de câble en double`);
    }
    // ligne de retour au FACP
    rows.push({
      tag: opts.facpName, loop, kind: 'return', onLoop: true, srcRow: prevRow,
      addr: '', desc: opts.facpDescription, area: lastArea,
    });
  }

  // étiquettes (utilisées pour l'aperçu et comme valeur en cache des formules)
  for (const row of rows) row.label = row.loop && row.kind === 'device' ? `${row.loop}-${row.tag}` : row.tag;
  for (const row of rows) row.srcLabel = row.srcRow === null ? opts.facpName : rows[row.srcRow - 2].label;

  return { rows, loops, warnings };
}

/* =============================================================================
   3. ÉCRITURE DU FICHIER EXCEL (ExcelJS : styles, fusions, formules)
   ============================================================================= */
function buildWorkbook(rows, opts, ExcelLib) {
  const wb = new ExcelLib.Workbook();
  const ws = wb.addWorksheet('EQUIPMENT');

  const font = { name: 'Calibri', size: 14 };
  const headFont = { name: 'Calibri', size: 14, bold: true, color: { argb: 'FFFFFFFF' } };
  const solid = (argb) => ({ type: 'pattern', pattern: 'solid', fgColor: { argb } });
  const thin = { style: 'thin' };
  const border = { top: thin, left: thin, bottom: thin, right: thin };
  const center = { horizontal: 'center', vertical: 'middle', wrapText: true };

  const layout = [
    [1, 1, 'ARTICOLO'], [2, 2, 'ATTREZZATURA'], [3, 3, 'INDIRIZZO'], [4, 4, 'LOOP n°'],
    [5, 5, 'ETICHETTA PER CAVI'], [6, 9, 'DA'], [10, 10, 'A'], [11, 12, 'TIPO'],
    [13, 13, 'SHIELD'], [14, 14, 'BLINDATO'], [15, 15, 'LUNGHEZZA (m)'],
    [16, 19, 'DESCRIZIONE'], [20, 23, 'AREA'],
  ];

  function put(r, c1, c2, value, o = {}) {
    for (let c = c1; c <= c2; c++) {
      const cell = ws.getCell(r, c);
      cell.border = border; cell.alignment = center; cell.font = o.font || font;
      if (o.fill) cell.fill = o.fill;
    }
    const master = ws.getCell(r, c1);
    master.value = value === '' || value === undefined ? null : value;
    if (o.numFmt) master.numFmt = o.numFmt;
    if (c2 > c1) ws.mergeCells(r, c1, r, c2);
  }

  for (const [c1, c2, title] of layout) put(1, c1, c2, title, { font: headFont, fill: solid('FF4472C4') });
  ws.getRow(1).height = 40;

  const yellow = solid('FFFFFF00');
  rows.forEach((d, i) => {
    const r = i + 2;
    const labelFormula = d.loop && d.kind === 'device'
      ? { formula: `CONCATENATE(D${r},"-",B${r})`, result: d.label }
      : { formula: `CONCATENATE(B${r})`, result: d.label };
    put(r, 1, 1, i + 1, { numFmt: '0' });
    put(r, 2, 2, d.tag, { numFmt: '@' });
    put(r, 3, 3, d.addr, { numFmt: '@' });
    put(r, 4, 4, d.loop, { numFmt: '@' });
    put(r, 5, 5, { formula: `J${r}`, result: d.label });
    put(r, 6, 9, d.srcRow === null ? opts.facpName : { formula: `J${d.srcRow}`, result: d.srcLabel });
    put(r, 10, 10, labelFormula);
    put(r, 11, 12, opts.cableType, { fill: yellow });      // jaune = à vérifier / modifier
    put(r, 13, 13, opts.shield);
    put(r, 14, 14, opts.armored);
    put(r, 15, 15, null, { fill: yellow, numFmt: '0' });    // jaune = longueur à saisir
    put(r, 16, 19, d.desc);
    put(r, 20, 23, d.area);
    ws.getRow(r).height = 20.1;
  });

  // barre de fermeture du tableau
  const end = rows.length + 2;
  for (let c = 1; c <= 23; c++) ws.getCell(end, c).fill = solid('FF000000');
  ws.mergeCells(end, 1, end, 23);
  ws.getRow(end).height = 6;

  const widths = [10, 16, 12, 10, 26, 7, 7, 7, 7, 26, 8, 8, 10, 13, 17, 30, 10, 10, 10, 20, 12, 12, 12];
  widths.forEach((w, i) => { ws.getColumn(i + 1).width = w; });

  ws.views = [{ state: 'frozen', ySplit: 1 }];
  ws.autoFilter = `A1:W${end - 1}`;
  ws.pageSetup = { orientation: 'landscape', fitToPage: true, fitToWidth: 1, fitToHeight: 0 };
  return wb;
}

/* Export pour les tests Node (sans effet dans le navigateur) */
if (typeof module !== 'undefined' && module.exports) {
  module.exports = { DEFAULTS, parseIoMatrix, buildRows, buildWorkbook, findHeader, cellText };
}

/* =============================================================================
   4. INTERFACE (navigateur uniquement)
   ============================================================================= */
if (typeof document !== 'undefined') {
  const $ = (id) => document.getElementById(id);
  const state = { fileName: '', candidates: [], result: null };

  const el = {
    drop: $('drop'), file: $('file'), fileInfo: $('file-info'), sheetRow: $('sheet-row'),
    sheet: $('sheet'), status: $('status'), results: $('results'), empty: $('empty'),
    stats: $('stats'), loopBody: $('loop-body'), warnBox: $('warn-box'), warnList: $('warn-list'),
    warnCount: $('warn-count'), preview: $('preview-body'), download: $('download'),
    addressMode: $('opt-address'), cableType: $('opt-cable'), facp: $('opt-facp'),
    remove: $('opt-remove'), hubs: $('opt-hubs'), reset: $('opt-reset'),
  };

  function setStatus(msg, kind) {
    el.status.textContent = msg || '';
    el.status.dataset.kind = kind || '';
  }

  function readOptions() {
    const split = (s) => s.split(',').map((x) => x.trim()).filter(Boolean);
    return {
      ...DEFAULTS,
      addressMode: el.addressMode.value,
      cableType: el.cableType.value.trim() || DEFAULTS.cableType,
      facpName: el.facp.value.trim() || DEFAULTS.facpName,
      removeText: split(el.remove.value),
      hubPrefixes: split(el.hubs.value.toUpperCase()),
    };
  }

  function resetOptions() {
    el.addressMode.value = DEFAULTS.addressMode;
    el.cableType.value = DEFAULTS.cableType;
    el.facp.value = DEFAULTS.facpName;
    el.remove.value = DEFAULTS.removeText.join(', ');
    el.hubs.value = DEFAULTS.hubPrefixes.join(', ');
  }

  function sheetToMatrix(ws) {
    const matrix = [];
    const cols = ws.columnCount;
    for (let r = 1; r <= ws.rowCount; r++) {
      const row = ws.getRow(r);
      const arr = [];
      for (let c = 1; c <= cols; c++) arr.push(cellText(row.getCell(c).value));
      matrix.push(arr);
    }
    return matrix;
  }

  async function loadFile(file) {
    if (!file) return;
    if (typeof ExcelJS === 'undefined') {
      setStatus('La bibliothèque ExcelJS n’a pas pu être chargée. Vérifie la connexion internet puis recharge la page.', 'error');
      return;
    }
    if (!/\.xlsx$/i.test(file.name)) {
      setStatus('Format non pris en charge : importe un fichier .xlsx (si ton IO list est en .xls, enregistre-la d’abord en .xlsx).', 'error');
      return;
    }
    setStatus('Lecture du fichier…');
    try {
      const wb = new ExcelJS.Workbook();
      await wb.xlsx.load(await file.arrayBuffer());
      state.fileName = file.name;
      state.candidates = [];
      wb.eachSheet((ws) => {
        const matrix = sheetToMatrix(ws);
        if (findHeader(matrix)) state.candidates.push({ name: ws.name, matrix });
      });
      el.fileInfo.textContent = file.name;
      if (!state.candidates.length) {
        el.results.hidden = true; el.empty.hidden = false; el.download.disabled = true;
        setStatus('Aucune feuille ne contient les colonnes LOOP, PREFISSO et NUMERO DELL’APPARECCHIATURA. Vérifie la ligne d’en-têtes de l’IO list.', 'error');
        return;
      }
      el.sheet.innerHTML = '';
      state.candidates.forEach((c, i) => el.sheet.add(new Option(c.name, i)));
      el.sheetRow.hidden = state.candidates.length < 2;
      generate();
    } catch (err) {
      console.error(err);
      setStatus('Impossible de lire ce fichier : il est peut-être protégé par mot de passe ou endommagé.', 'error');
    }
  }

  function generate() {
    const cand = state.candidates[Number(el.sheet.value) || 0];
    if (!cand) return;
    const opts = readOptions();
    const devices = parseIoMatrix(cand.matrix, opts);
    const res = buildRows(devices, opts);
    state.result = { ...res, opts, sheet: cand.name, deviceCount: devices.length };
    render();
    setStatus(`Cable list prête : ${res.rows.length} câbles pour ${res.loops.size} loops.`, 'ok');
  }

  const td = (text, cls) => {
    const c = document.createElement('td');
    c.textContent = text == null ? '' : text;
    if (cls) c.className = cls;
    return c;
  };

  function render() {
    const { rows, loops, warnings, deviceCount } = state.result;
    el.empty.hidden = true; el.results.hidden = false; el.download.disabled = false;

    // chiffres clés
    el.stats.innerHTML = '';
    [[deviceCount, 'appareils lus'], [loops.size, 'loops'], [rows.length, 'câbles générés']].forEach(([n, label]) => {
      const d = document.createElement('div');
      const b = document.createElement('strong'); b.textContent = n;
      const s = document.createElement('span'); s.textContent = label;
      d.append(b, s); el.stats.append(d);
    });

    // résumé par loop
    el.loopBody.innerHTML = '';
    for (const [loop, devs] of loops) {
      const onLoop = devs.filter((d) => d.onLoop).length;
      const tr = document.createElement('tr');
      tr.append(td(loop, 'k'), td(onLoop), td(devs.length - onLoop), td(devs.length + 1));
      el.loopBody.append(tr);
    }

    // avertissements
    el.warnBox.hidden = warnings.length === 0;
    el.warnCount.textContent = warnings.length === 1 ? '1 point à vérifier' : `${warnings.length} points à vérifier`;
    el.warnList.innerHTML = '';
    warnings.forEach((w) => { const li = document.createElement('li'); li.textContent = w; el.warnList.append(li); });

    // aperçu
    el.preview.innerHTML = '';
    const frag = document.createDocumentFragment();
    rows.forEach((d, i) => {
      const tr = document.createElement('tr');
      if (d.kind === 'return') tr.className = 'ret';
      else if (!d.onLoop) tr.className = 'spur';
      tr.append(td(i + 1, 'n'), td(d.label, 'k'), td(d.addr), td(d.srcLabel), td(d.desc), td(d.area));
      frag.append(tr);
    });
    el.preview.append(frag);
  }

  async function download() {
    if (!state.result) return;
    try {
      const wb = buildWorkbook(state.result.rows, state.result.opts, ExcelJS);
      const buf = await wb.xlsx.writeBuffer();
      const blob = new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url; a.download = 'CABLE_LIST_generated.xlsx';
      document.body.append(a); a.click(); a.remove();
      setTimeout(() => URL.revokeObjectURL(url), 2000);
      setStatus('CABLE_LIST_generated.xlsx téléchargé.', 'ok');
    } catch (err) {
      console.error(err);
      setStatus('Le fichier Excel n’a pas pu être créé. Recharge la page et réessaie.', 'error');
    }
  }

  /* événements */
  el.file.addEventListener('change', (e) => loadFile(e.target.files[0]));
  ['dragenter', 'dragover'].forEach((t) => el.drop.addEventListener(t, (e) => { e.preventDefault(); el.drop.classList.add('over'); }));
  ['dragleave', 'drop'].forEach((t) => el.drop.addEventListener(t, (e) => { e.preventDefault(); el.drop.classList.remove('over'); }));
  el.drop.addEventListener('drop', (e) => loadFile(e.dataTransfer.files[0]));
  el.sheet.addEventListener('change', generate);
  [el.addressMode, el.cableType, el.facp, el.remove, el.hubs].forEach((i) =>
    i.addEventListener('change', () => { if (state.candidates.length) generate(); }));
  el.reset.addEventListener('click', () => { resetOptions(); if (state.candidates.length) generate(); });
  el.download.addEventListener('click', download);

  resetOptions();
}
