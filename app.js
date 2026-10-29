
'use strict';

/* =============================================================================
   Cable List Builder – IO LIST -> CABLE LIST (EQUIPMENT sheet)
   ============================================================================= */

/* ---------- Options et paramètres par défaut ---------- */
const DEFAULTS = {
  cableType: '2x1,5mmq',
  powerCableType: '2x2,5mmq',
  fdCableType: '2x (2x1,5mmq)',
  shield: 'SH',
  armored: '-',
  facpName: 'FACP',
  facpDescription: '',
  removeText: ['-PORTA TAGLIAFUOCO'],
  hubPrefixes: ['CCP'],
};

const HEADER_ALIASES = {
  loop: ['LOOP'],
  prefix: ['PREFISSO', 'PREFIX'],
  number: ["NUMERO DELL'APPARECCHIATURA", 'NUMERO', 'NUMBER'],
  address: ['INDIRIZZO', 'ADDRESS'],
  desc: ['DESCRIZIONE', 'DESCRIPTION'],
  area: ['AREA', 'ZONE'],
};

/* =============================================================================
   1. PARSING DE LA LISTE IO (IO LIST)
   ============================================================================= */

const norm = (v) =>
  (v == null ? '' : String(v).replace(/\s+/g, ' ').trim().toUpperCase());

function clean(v) {
  const s = v == null ? '' : String(v).trim();
  return s === '-' || s === '--' ? '' : s;
}

const escapeRe = (s) => s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');

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

function findHeader(matrix) {
  const limit = Math.min(matrix.length, 15);

  for (let r = 0; r < limit; r++) {
    const heads = (matrix[r] || []).map(norm);
    const cols = {};

    for (const [field, aliases] of Object.entries(HEADER_ALIASES)) {
      let idx = -1;

      for (const a of aliases) {
        idx = heads.indexOf(a);
        if (idx >= 0) break;
      }

      if (idx < 0) {
        for (const a of aliases) {
          idx = heads.findIndex((h) => h.startsWith(a));
          if (idx >= 0) break;
        }
      }

      if (idx >= 0) cols[field] = idx;
    }

    if ('loop' in cols && 'prefix' in cols && 'number' in cols) {
      return { row: r, cols };
    }
  }

  return null;
}

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

    const num =
      typeof rawNum === 'number'
        ? String(Math.trunc(rawNum)).padStart(2, '0')
        : clean(rawNum);

    let desc = clean(get('desc'));

    for (const rm of opts.removeText) {
      if (rm) desc = desc.replace(new RegExp(escapeRe(rm), 'gi'), '');
    }

    desc = desc
      .replace(/\s+/g, ' ')
      .trim()
      .replace(/^[\s-]+|[\s-]+$/g, '');

    devices.push({
      srcRow: r + 1,
      loop: clean(get('loop')),
      prefix,
      num,
      addr: clean(get('address')),
      desc,
      area: clean(get('area')),
    });
  }

  return devices;
}

/* =============================================================================
   2. CONSTRUCTION DE LA LISTE DES CÂBLES
   ============================================================================= */

function buildRows(devices, opts) {
  const warnings = [];
  const loops = new Map();
  let current = null;

  for (const d of devices) {
    if (d.loop) current = d.loop;

    if (current === null) {
      warnings.push(
        `Ligne IO ${d.srcRow} (${d.prefix}-${d.num}): pas de boucle définie, ignoré`
      );
      continue;
    }

    d.onLoop = Boolean(d.loop && d.addr);

    if (!loops.has(current)) loops.set(current, []);
    loops.get(current).push(d);
  }

  const rows = [];
  const hubs = new Set(opts.hubPrefixes);

  for (const [loop, devs] of loops) {
    let prevRowIndex = null;
    let hubRowIndex = null;
    let lastArea = devs[0].area;

    const tagCount = new Map();
    const powerDevices = [];

    for (const d of devs) {
      const tag = `${d.prefix}-${d.num}`;
      const fullTag = d.loop && d.onLoop ? `${d.loop}-${tag}` : tag;

      tagCount.set(fullTag, (tagCount.get(fullTag) || 0) + 1);

      const rIndex = rows.length + 2;
      let srcRowIndex;

      if (d.onLoop) {
        srcRowIndex = prevRowIndex;
        prevRowIndex = rIndex;
        hubRowIndex = null;
      } else {
        const parent = hubRowIndex || prevRowIndex;
        srcRowIndex = parent;

        if (hubs.has(d.prefix)) hubRowIndex = rIndex;
      }

      lastArea = d.area || lastArea;

      // Câble MIO du FD : conserver le type de câble original.
      let cableType = opts.cableType;

      if (d.prefix === 'FD') {
        cableType = opts.fdCableType;
      }

      const rowObj = {
        tag: fullTag,
        rawTag: tag,
        prefix: d.prefix,
        num: d.num,
        loop: d.loop,
        kind: 'device',
        onLoop: d.onLoop,
        srcRowIndex,
        addr: d.addr,
        desc: d.desc,
        area: d.area,
        cableType,
        shield: opts.shield,
        isPower: false,
      };

      rows.push(rowObj);

      if (d.prefix === 'FD' || d.prefix === 'SOA') {
        powerDevices.push(rowObj);
      }
    }

    // -------------------------------------------------------------------------
    // CÂBLES ALM / PUISSANCE
    //
    // Chaque FD conserve sa ligne MIO et reçoit une ligne ALM distincte.
    // Le type du câble ALM de chaque FD est TBA et doit être surligné.
    // Les autres câbles de puissance conservent leur type habituel.
    // -------------------------------------------------------------------------

    const almDev = devs.find((d) => d.prefix === 'ALM');

    if (
      powerDevices.length > 0 &&
      (almDev || powerDevices.some((pd) => pd.prefix === 'FD'))
    ) {
      let currentSourceTag = almDev
        ? `${almDev.loop}-${almDev.prefix}-${almDev.num}`
        : 'TBA';

      for (const pd of powerDevices) {
        const isFdAlm = pd.prefix === 'FD';

        // Sans ALM déclaré dans la boucle, générer tout de même
        // le câble ALM des FD. Les autres câbles restent inchangés.
        if (!almDev && !isFdAlm) continue;

        const pLabel = isFdAlm
          ? `${pd.rawTag}-ALM`
          : `${pd.tag}-P`;

        const targetLabel = pd.tag;

        rows.push({
          tag: pLabel,
          rawTag: pLabel,
          prefix: pd.prefix,
          num: pd.num,
          loop: pd.loop,
          kind: 'power',
          onLoop: false,
          srcLabel: currentSourceTag,
          targetLabel,
          desc: pd.desc,
          area: pd.area,

          // ALM des FD = TBA ; autres câbles = type de puissance habituel.
          cableType: isFdAlm ? 'TBA' : opts.powerCableType,

          // Utilisé pour le surlignage dans l'aperçu.
          highlightCableType: isFdAlm,

          shield: '-',
          isPower: true,
        });

        currentSourceTag = targetLabel;
      }
    }

    // Bouclage retour FACP
    rows.push({
      tag: opts.facpName,
      rawTag: opts.facpName,
      loop,
      kind: 'return',
      onLoop: true,
      srcRowIndex: prevRowIndex,
      addr: '',
      desc: opts.facpDescription,
      area: lastArea,
      cableType: opts.cableType,
      shield: opts.shield,
      isPower: false,
    });
  }

  // Résolution des paires DA -> A
  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];

    if (row.isPower) {
      row.daLabel = row.srcLabel;
      row.aLabel = row.targetLabel;
    } else if (row.srcRowIndex === null) {
      row.daLabel = opts.facpName;
      row.aLabel = row.tag;
    } else {
      const srcRowObj = rows[row.srcRowIndex - 2];

      row.daLabel = srcRowObj ? srcRowObj.tag : opts.facpName;
      row.aLabel = row.tag;
    }
  }

  return { rows, loops, warnings };
}

/* =============================================================================
   3. CRÉATION DU FICHIER EXCEL FINAL
   ============================================================================= */

function buildWorkbook(rows, opts, ExcelLib) {
  const wb = new ExcelLib.Workbook();
  const ws = wb.addWorksheet('EQUIPMENT');

  const font = { name: 'Calibri', size: 14 };

  const headFont = {
    name: 'Calibri',
    size: 14,
    bold: true,
    color: { argb: 'FFFFFFFF' },
  };

  const solid = (argb) => ({
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb },
  });

  const thin = { style: 'thin' };

  const border = {
    top: thin,
    left: thin,
    bottom: thin,
    right: thin,
  };

  const center = {
    horizontal: 'center',
    vertical: 'middle',
    wrapText: true,
  };

  const layout = [
    [1, 1, 'ARTICOLO'],
    [2, 2, 'ETICHETTA PER CAVI'],
    [3, 3, 'DA'],
    [4, 4, 'A'],
    [5, 6, 'TIPO'],
    [7, 7, 'SHIELD'],
    [8, 8, 'LUNGHEZZA (m)'],
    [9, 9, 'DESCRIZIONE'],
    [10, 10, 'AREA'],
  ];

  function put(r, c1, c2, value, o = {}) {
    for (let c = c1; c <= c2; c++) {
      const cell = ws.getCell(r, c);

      cell.border = border;
      cell.alignment = center;
      cell.font = o.font || font;

      if (o.fill) cell.fill = o.fill;
    }

    const master = ws.getCell(r, c1);

    master.value =
      value === '' || value === undefined ? null : value;

    if (o.numFmt) master.numFmt = o.numFmt;

    if (c2 > c1) ws.mergeCells(r, c1, r, c2);
  }

  // En-tête
  for (const [c1, c2, title] of layout) {
    put(1, c1, c2, title, {
      font: headFont,
      fill: solid('FF4472C4'),
    });
  }

  ws.getRow(1).height = 40;

  const yellow = solid('FFFFFF00');

  // Lignes de données
  rows.forEach((d, i) => {
    const r = i + 2;
    const numStr = String(i + 1).padStart(2, '0');

    put(r, 1, 1, numStr, { numFmt: '@' });
    put(r, 2, 2, d.tag, { numFmt: '@' });
    put(r, 3, 3, d.daLabel);
    put(r, 4, 4, d.aLabel);

    // TBA est explicitement surligné.
    // Le fond jaune existant des types de câbles est conservé.
    put(r, 5, 6, d.cableType, {
      fill: d.highlightCableType ? yellow : yellow,
      font: d.highlightCableType
        ? { name: 'Calibri', size: 14, bold: true }
        : font,
    });

    put(r, 7, 7, d.shield);
    put(r, 8, 8, null, { fill: yellow, numFmt: '0' });
    put(r, 9, 9, d.desc);
    put(r, 10, 10, d.area);

    ws.getRow(r).height = 20.1;
  });

  const columnWidths = [
    10, 25, 20, 22, 12, 12, 12, 16, 45, 30,
  ];

  columnWidths.forEach((w, i) => {
    ws.getColumn(i + 1).width = w;
  });

  ws.views = [{ state: 'frozen', ySplit: 1 }];
  ws.autoFilter = `A1:J${rows.length + 1}`;

  ws.pageSetup = {
    orientation: 'landscape',
    fitToPage: true,
    fitToWidth: 1,
    fitToHeight: 0,
  };

  return wb;
}

if (typeof module !== 'undefined' && module.exports) {
  module.exports = {
    DEFAULTS,
    parseIoMatrix,
    buildRows,
    buildWorkbook,
    findHeader,
    cellText,
  };
}

/* =============================================================================
   4. INTERACTION NAVIGATEUR ET UI
   ============================================================================= */

if (typeof document !== 'undefined') {
  const $ = (id) => document.getElementById(id);

  const state = {
    fileName: '',
    candidates: [],
    result: null,
  };

  const el = {
    drop: $('drop'),
    file: $('file'),
    fileInfo: $('file-info'),
    sheetRow: $('sheet-row'),
    sheet: $('sheet'),
    status: $('status'),
    results: $('results'),
    empty: $('empty'),
    stats: $('stats'),
    preview: $('preview-body'),
    download: $('download'),
    cableType: $('opt-cable'),
    facp: $('opt-facp'),
    remove: $('opt-remove'),
    hubs: $('opt-hubs'),
    reset: $('opt-reset'),
  };

  function setStatus(msg, kind) {
    if (!el.status) return;

    el.status.textContent = msg || '';
    el.status.dataset.kind = kind || '';
  }

  function readOptions() {
    const split = (s) =>
      s.split(',').map((x) => x.trim()).filter(Boolean);

    return {
      ...DEFAULTS,

      cableType: el.cableType
        ? el.cableType.value.trim() || DEFAULTS.cableType
        : DEFAULTS.cableType,

      facpName: el.facp
        ? el.facp.value.trim() || DEFAULTS.facpName
        : DEFAULTS.facpName,

      removeText: el.remove
        ? split(el.remove.value)
        : DEFAULTS.removeText,

      hubPrefixes: el.hubs
        ? split(el.hubs.value.toUpperCase())
        : DEFAULTS.hubPrefixes,
    };
  }

  function resetOptions() {
    if (el.cableType) {
      el.cableType.value = DEFAULTS.cableType;
    }

    if (el.facp) {
      el.facp.value = DEFAULTS.facpName;
    }

    if (el.remove) {
      el.remove.value = DEFAULTS.removeText.join(', ');
    }

    if (el.hubs) {
      el.hubs.value = DEFAULTS.hubPrefixes.join(', ');
    }
  }

  function sheetToMatrix(ws) {
    const matrix = [];
    const cols = ws.columnCount;

    for (let r = 1; r <= ws.rowCount; r++) {
      const row = ws.getRow(r);
      const arr = [];

      for (let c = 1; c <= cols; c++) {
        arr.push(cellText(row.getCell(c).value));
      }

      matrix.push(arr);
    }

    return matrix;
  }

  async function loadFile(file) {
    if (!file) return;

    if (typeof ExcelJS === 'undefined') {
      setStatus('ExcelJS non chargé.', 'error');
      return;
    }

    setStatus('Lecture en cours...');

    try {
      const wb = new ExcelJS.Workbook();

      await wb.xlsx.load(await file.arrayBuffer());

      state.fileName = file.name;
      state.candidates = [];

      wb.eachSheet((ws) => {
        const matrix = sheetToMatrix(ws);

        if (findHeader(matrix)) {
          state.candidates.push({
            name: ws.name,
            matrix,
          });
        }
      });

      if (el.fileInfo) {
        el.fileInfo.textContent = file.name;
      }

      if (!state.candidates.length) {
        if (el.results) el.results.hidden = true;

        setStatus('Format IO LIST non reconnu.', 'error');
        return;
      }

      if (el.sheet) {
        el.sheet.innerHTML = '';

        state.candidates.forEach((c, i) => {
          el.sheet.add(new Option(c.name, i));
        });
      }

      generate();
    } catch (err) {
      console.error(err);
      setStatus('Erreur de lecture du fichier.', 'error');
    }
  }

  function generate() {
    const cand =
      state.candidates[
        Number(el.sheet ? el.sheet.value : 0) || 0
      ];

    if (!cand) return;

    const opts = readOptions();
    const devices = parseIoMatrix(cand.matrix, opts);

    if (!devices) {
      setStatus('Impossible de lire la liste IO.', 'error');
      return;
    }

    const res = buildRows(devices, opts);

    state.result = {
      ...res,
      opts,
      sheet: cand.name,
      deviceCount: devices.length,
    };

    render();

    setStatus(
      `Succès : ${res.rows.length} câbles générés.`,
      'ok'
    );
  }

  const td = (text, cls) => {
    const c = document.createElement('td');

    c.textContent = text == null ? '' : text;

    if (cls) c.className = cls;

    // Surlignage jaune pour le type TBA du câble ALM FD.
    if (cls === 'tba-highlight') {
      c.style.backgroundColor = '#FFFF00';
      c.style.fontWeight = 'bold';
      c.style.color = '#000000';
    }

    return c;
  };

  function render() {
    const { rows } = state.result;

    if (el.empty) el.empty.hidden = true;
    if (el.results) el.results.hidden = false;
    if (el.download) el.download.disabled = false;

    if (el.preview) {
      el.preview.innerHTML = '';

      const frag = document.createDocumentFragment();

      rows.forEach((d, i) => {
        const tr = document.createElement('tr');

        if (d.kind === 'return') {
          tr.className = 'ret';
        } else if (d.isPower) {
          tr.className = 'power';
        }

        tr.append(
          td(String(i + 1).padStart(2, '0'), 'n'),
          td(d.tag, 'k'),
          td(d.daLabel),
          td(d.aLabel),
          td(
            d.cableType,
            d.highlightCableType ? 'tba-highlight' : ''
          ),
          td(d.desc),
          td(d.area)
        );

        frag.append(tr);
      });

      el.preview.append(frag);
    }
  }

  async function download() {
    if (!state.result) return;

    try {
      const wb = buildWorkbook(
        state.result.rows,
        state.result.opts,
        ExcelJS
      );

      const buf = await wb.xlsx.writeBuffer();

      const blob = new Blob([buf], {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      });

      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');

      a.href = url;
      a.download = 'CABLE_LIST.xlsx';

      document.body.append(a);
      a.click();
      a.remove();

      setTimeout(() => URL.revokeObjectURL(url), 2000);

      setStatus('CABLE_LIST.xlsx généré et téléchargé.', 'ok');
    } catch (err) {
      console.error(err);
      setStatus('Erreur lors du téléchargement.', 'error');
    }
  }

  if (el.file) {
    el.file.addEventListener('change', (e) => {
      loadFile(e.target.files[0]);
    });
  }

  if (el.download) {
    el.download.addEventListener('click', download);
  }

  if (el.sheet) {
    el.sheet.addEventListener('change', generate);
  }

  if (el.reset) {
    el.reset.addEventListener('click', () => {
      resetOptions();

      if (state.candidates.length) {
        generate();
      }
    });
  }

  resetOptions();
}
