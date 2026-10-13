const $ = (id) => document.getElementById(id);

const STANDARD_FIELDS = [
  { key: 'id', label: 'ID Cliente', aliases: ['id cliente', 'id_cliente', 'cliente id', 'codigo cliente', 'codigo_cliente', 'cedula', 'documento', 'dni', 'identificacion', 'nif'] },
  { key: 'name', label: 'Nombre', aliases: ['nombre', 'nombres', 'nombre completo', 'nombre_completo', 'full name', 'name', 'cliente'] },
  { key: 'email', label: 'Email', aliases: ['email', 'correo', 'correo electronico', 'e-mail', 'mail'] },
  { key: 'phone', label: 'Teléfono', aliases: ['telefono', 'teléfono', 'tel', 'celular', 'movil', 'móvil', 'phone', 'whatsapp'] },
  { key: 'address', label: 'Dirección', aliases: ['direccion', 'dirección', 'domicilio', 'address', 'calle'] },
  { key: 'city', label: 'Ciudad', aliases: ['ciudad', 'municipio', 'localidad', 'city'] },
  { key: 'date', label: 'Fecha de registro', aliases: ['fecha registro', 'fecha_registro', 'fecha de registro', 'fecha alta', 'fecha_alta', 'registro', 'created at'] }
];

const state = {
  datasets: [],
  records: [],
  groups: [],
  filter: 'all',
  sourceFilter: '',
  search: '',
  page: 0,
  selected: null,
  detailView: 'grouped',
  lastRun: null,
  currentStep: 1
};

const fieldByKey = (key) => STANDARD_FIELDS.find((field) => field.key === key);
const normalize = (value) => String(value ?? '')
  .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
  .toLowerCase().trim().replace(/\s+/g, ' ');
const digits = (value) => String(value ?? '').replace(/\D/g, '');
const escapeHtml = (value) => String(value ?? '').replace(/[&<>"']/g, (char) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[char]));
const formatNumber = (value) => new Intl.NumberFormat('es').format(value);
const formatBytes = (bytes) => bytes < 1024 ? `${bytes} B` : bytes < 1048576 ? `${(bytes / 1024).toFixed(0)} KB` : `${(bytes / 1048576).toFixed(1)} MB`;

function normalizedField(value, field) {
  if (field === 'phone') return digits(value);
  if (field === 'id') return normalize(value).replace(/[^a-z0-9]/g, '');
  return normalize(value);
}

function parseCSV(text, forcedSeparator) {
  const firstLine = text.split(/\r?\n/, 1)[0] || '';
  const separators = forcedSeparator ? [forcedSeparator] : [',', ';', '\t', '|'];
  const separator = separators.sort((a, b) => firstLine.split(b).length - firstLine.split(a).length)[0];
  const rows = [];
  let row = [], value = '', quoted = false;

  for (let index = 0; index < text.length; index += 1) {
    const char = text[index];
    if (char === '"') {
      if (quoted && text[index + 1] === '"') { value += '"'; index += 1; }
      else quoted = !quoted;
    } else if (!quoted && char === separator) {
      row.push(value); value = '';
    } else if (!quoted && (char === '\n' || char === '\r')) {
      if (char === '\r' && text[index + 1] === '\n') index += 1;
      row.push(value);
      if (row.some((cell) => String(cell).trim())) rows.push(row);
      row = []; value = '';
    } else value += char;
  }
  row.push(value);
  if (row.some((cell) => String(cell).trim())) rows.push(row);
  return rows;
}

function matrixToDataset(matrix, title, file, metadata = {}) {
  const copy = matrix.map((row) => [...row]);
  const rawHeaders = copy.shift() || [];
  const headers = rawHeaders.map((header, index) => String(header || `Columna ${index + 1}`).trim());
  const rows = copy
    .filter((row) => row.some((cell) => String(cell ?? '').trim()))
    .map((row) => Object.fromEntries(headers.map((header, index) => [header, String(row[index] ?? '')])));
  const dataset = { id: crypto.randomUUID(), title, filename: file.name, size: file.size, extension: file.name.split('.').pop().toLowerCase(), headers, rows, map: {}, ...metadata };
  autoMap(dataset);
  return dataset;
}

function autoMap(dataset) {
  STANDARD_FIELDS.forEach((field) => {
    const exact = dataset.headers.find((header) => field.aliases.some((alias) => normalize(header) === normalize(alias)));
    const partial = dataset.headers.find((header) => field.aliases.some((alias) => normalize(header).includes(normalize(alias)) || normalize(alias).includes(normalize(header))));
    dataset.map[field.key] = exact || partial || '';
  });
}

async function loadScript(url, ready) {
  if (ready()) return;
  await new Promise((resolve, reject) => {
    const script = document.createElement('script');
    script.src = url;
    script.onload = resolve;
    script.onerror = () => reject(new Error('No se pudo cargar la biblioteca necesaria para este formato.'));
    document.head.append(script);
  });
}

function pdfLineCells(items) {
  if (!items.length) return [];
  const heights = items.map((item) => Math.abs(item.height || item.transform?.[3] || 10)).filter(Boolean).sort((a, b) => a - b);
  const medianHeight = heights[Math.floor(heights.length / 2)] || 10;
  const gapLimit = Math.max(9, medianHeight * 1.25);
  const sorted = [...items].sort((a, b) => a.transform[4] - b.transform[4]);
  const cells = [];
  sorted.forEach((item) => {
    const text = String(item.str || '').trim();
    if (!text) return;
    const x = item.transform[4];
    const width = Number(item.width) || Math.max(4, text.length * medianHeight * .45);
    const previous = cells[cells.length - 1];
    if (!previous || x - previous.end > gapLimit) cells.push({ x, end: x + width, text });
    else {
      previous.text += `${previous.text.endsWith('-') ? '' : ' '}${text}`;
      previous.end = Math.max(previous.end, x + width);
    }
  });
  return cells;
}

function groupPdfLines(items, pageNumber) {
  const lines = [];
  [...items].filter((item) => String(item.str || '').trim()).sort((a, b) => b.transform[5] - a.transform[5] || a.transform[4] - b.transform[4]).forEach((item) => {
    const y = item.transform[5];
    let line = lines.find((candidate) => Math.abs(candidate.y - y) <= 3);
    if (!line) { line = { y, pageNumber, items: [] }; lines.push(line); }
    line.items.push(item);
  });
  return lines.sort((a, b) => b.y - a.y).map((line) => ({ ...line, cells: pdfLineCells(line.items) }));
}

async function extractPdfMatrix(pdf) {
  const lines = [];
  for (let pageNumber = 1; pageNumber <= pdf.numPages; pageNumber += 1) {
    const page = await pdf.getPage(pageNumber);
    const content = await page.getTextContent({ normalizeWhitespace: true });
    lines.push(...groupPdfLines(content.items, pageNumber));
  }
  if (!lines.length) throw new Error('El PDF no contiene texto seleccionable. Aplica OCR al documento y vuelve a intentarlo.');

  const structured = lines.map((line, index) => ({ ...line, index })).filter((line) => line.cells.length >= 2);
  const frequencies = new Map();
  structured.forEach((line) => frequencies.set(line.cells.length, (frequencies.get(line.cells.length) || 0) + 1));
  const targetColumns = [...frequencies.entries()].sort((a, b) => b[1] - a[1] || b[0] - a[0])[0]?.[0] || 0;
  const headerLine = structured.find((line) => line.cells.length === targetColumns);

  if (targetColumns < 2 || !headerLine || (frequencies.get(targetColumns) || 0) < 2) {
    const textRows = lines.map((line) => line.items.sort((a, b) => a.transform[4] - b.transform[4]).map((item) => String(item.str || '').trim()).filter(Boolean).join(' ')).filter(Boolean);
    return { matrix: [['Texto'], ...textRows.map((text) => [text])], mode: 'text' };
  }

  const header = headerLine.cells.map((cell, index) => cell.text || `Columna ${index + 1}`);
  const starts = headerLine.cells.map((cell) => cell.x);
  const boundaries = starts.slice(1).map((start, index) => (starts[index] + start) / 2);
  const rows = [];
  lines.slice(headerLine.index + 1).forEach((line) => {
    const row = Array.from({ length: header.length }, () => '');
    line.items.sort((a, b) => a.transform[4] - b.transform[4]).forEach((item) => {
      const text = String(item.str || '').trim();
      if (!text) return;
      const center = item.transform[4] + (Number(item.width) || 0) / 2;
      let column = boundaries.findIndex((boundary) => center < boundary);
      if (column < 0) column = header.length - 1;
      row[column] = `${row[column]} ${text}`.trim();
    });
    const isHeader = row.every((value, index) => normalize(value) === normalize(header[index]));
    if (!isHeader && row.filter(Boolean).length >= 2) rows.push(row);
  });
  if (!rows.length) return { matrix: [['Texto'], ...lines.map((line) => [line.cells.map((cell) => cell.text).join(' · ')])], mode: 'text' };
  return { matrix: [header, ...rows], mode: 'table' };
}

async function readFile(file) {
  const extension = file.name.split('.').pop().toLowerCase();
  if (['csv', 'tsv', 'txt'].includes(extension)) {
    return [matrixToDataset(parseCSV(await file.text(), extension === 'tsv' ? '\t' : undefined), file.name, file)];
  }
  if (['xlsx', 'xls'].includes(extension)) {
    await loadScript('https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js', () => Boolean(window.XLSX));
    const workbook = XLSX.read(await file.arrayBuffer(), { type: 'array', cellDates: true });
    return workbook.SheetNames.map((sheetName) => matrixToDataset(
      XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1, defval: '', raw: false }),
      workbook.SheetNames.length > 1 ? `${file.name} · ${sheetName}` : file.name,
      file
    ));
  }
  if (extension === 'pdf') {
    await loadScript('https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js', () => Boolean(window.pdfjsLib));
    pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';
    const pdf = await pdfjsLib.getDocument({ data: new Uint8Array(await file.arrayBuffer()) }).promise;
    const extracted = await extractPdfMatrix(pdf);
    return [matrixToDataset(extracted.matrix, file.name, file, { pdfMode: extracted.mode, pdfPages: pdf.numPages })];
  }
  throw new Error('Formato no compatible.');
}

async function addFiles(files) {
  if (!files.length) return;
  $('loadStatus').textContent = 'Procesando archivos…';
  const errors = [];
  for (const file of files) {
    try { state.datasets.push(...await readFile(file)); }
    catch (error) { errors.push(`${file.name}: ${error.message}`); }
  }
  renderFiles();
  renderMapping();
  renderCompareControls();
  renderSourceFilter();
  updateProgress();
  const textPdfs = state.datasets.filter((dataset) => dataset.pdfMode === 'text').length;
  const success = `${state.datasets.length} ${state.datasets.length === 1 ? 'fuente cargada' : 'fuentes cargadas'} correctamente.${textPdfs ? ` ${textPdfs} PDF se interpretó como texto libre y no bloqueará la comparación.` : ''}`;
  $('loadStatus').textContent = errors.length ? `${success} ${errors.join(' · ')}` : success;
  $('files').value = '';
}

function renderFiles() {
  if (!state.datasets.length) {
    $('fileList').innerHTML = '<div class="empty-files"><span>Aún no hay archivos</span><small>Agrega al menos dos para comenzar</small></div>';
    $('previewData').disabled = true;
    $('autoMapAll').disabled = true;
    return;
  }
  $('previewData').disabled = false;
  $('autoMapAll').disabled = false;
  $('fileList').innerHTML = state.datasets.map((dataset) => {
    const type = ['xlsx', 'xls'].includes(dataset.extension) ? 'XLS' : dataset.extension.toUpperCase();
    const iconClass = dataset.extension === 'pdf' ? 'pdf' : ['csv', 'tsv'].includes(dataset.extension) ? 'csv' : dataset.extension === 'txt' ? 'txt' : '';
    const pdfLabel = dataset.pdfMode === 'table' ? ' · tabla detectada' : dataset.pdfMode === 'text' ? ' · texto libre' : '';
    return `<article class="file-item"><span class="file-icon ${iconClass}">${type}</span><span class="file-info"><strong title="${escapeHtml(dataset.title)}">${escapeHtml(dataset.title)}</strong><small>${formatNumber(dataset.rows.length)} filas · ${formatBytes(dataset.size)}${pdfLabel}</small></span><button class="remove-file" type="button" data-remove="${dataset.id}" aria-label="Quitar ${escapeHtml(dataset.title)}">×</button></article>`;
  }).join('');
  $('fileList').querySelectorAll('[data-remove]').forEach((button) => button.addEventListener('click', () => {
    state.datasets = state.datasets.filter((dataset) => dataset.id !== button.dataset.remove);
    state.groups = []; state.records = []; state.selected = null;
    renderFiles(); renderMapping(); renderCompareControls(); renderSourceFilter(); renderResults(); updateProgress();
  }));
}

function renderMapping() {
  const table = $('mappingTable');
  if (!state.datasets.length) { table.className = 'mapping-table empty-mapping'; table.innerHTML = ''; return; }
  table.className = 'mapping-table';
  const columns = `128px repeat(${state.datasets.length}, minmax(145px, 1fr))`;
  const header = `<div class="map-row map-head" style="grid-template-columns:${columns}"><div class="map-cell">Campo estándar</div>${state.datasets.map((dataset) => `<div class="map-cell" title="${escapeHtml(dataset.title)}">${escapeHtml(dataset.title)}</div>`).join('')}</div>`;
  const rows = STANDARD_FIELDS.map((field) => `<div class="map-row" style="grid-template-columns:${columns}"><div class="map-cell map-standard">${field.label}</div>${state.datasets.map((dataset) => `<div class="map-cell"><select data-dataset="${dataset.id}" data-field="${field.key}" aria-label="${field.label} en ${escapeHtml(dataset.title)}"><option value="">Sin asignar</option>${dataset.headers.map((headerName) => `<option value="${escapeHtml(headerName)}" ${dataset.map[field.key] === headerName ? 'selected' : ''}>${escapeHtml(headerName)}</option>`).join('')}</select></div>`).join('')}</div>`).join('');
  table.innerHTML = header + rows;
  table.querySelectorAll('select').forEach((select) => select.addEventListener('change', () => {
    const dataset = state.datasets.find((item) => item.id === select.dataset.dataset);
    dataset.map[select.dataset.field] = select.value;
  }));
}

function renderCompareControls() {
  const currentPrimary = $('primaryKey').value || '';
  $('primaryKey').innerHTML = `<option value="">Análisis integral automático</option>${STANDARD_FIELDS.map((field) => `<option value="${field.key}" ${currentPrimary === field.key ? 'selected' : ''}>${field.label}</option>`).join('')}`;
  $('compareFields').innerHTML = STANDARD_FIELDS.map((field) => {
    const mode = field.key === 'phone' || field.key === 'address' ? 'flexible' : field.key === 'date' ? 'days' : 'exact';
    const options = mode === 'flexible'
      ? '<option value="flexible">Coincidencia flexible</option><option value="exact">Coincidencia exacta</option>'
      : mode === 'days'
        ? '<option value="days">Diferencia máxima de 7 días</option><option value="exact">Misma fecha</option>'
        : '<option value="exact">Coincidencia exacta</option><option value="flexible">Coincidencia flexible</option>';
    return `<label class="compare-field"><input type="checkbox" data-compare="${field.key}" checked><span>${field.label}</span><select class="compare-mode" data-mode="${field.key}" aria-label="Modo de ${field.label}">${options}<option value="ignore">Ignorar diferencias</option></select></label>`;
  }).join('');
  $('runCompare').disabled = state.datasets.length < 2;
  $('selectAll').checked = true;
  $('selectAll').indeterminate = false;
}

function comparisonConfig() {
  return [...document.querySelectorAll('[data-compare]:checked')].map((input) => ({
    field: input.dataset.compare,
    mode: document.querySelector(`[data-mode="${input.dataset.compare}"]`)?.value || 'exact'
  })).filter((config) => config.mode !== 'ignore');
}

function renderSourceFilter() {
  const current = state.sourceFilter;
  $('sourceFilter').innerHTML = '<option value="">Todos los archivos</option>' + state.datasets.map((dataset) => `<option value="${dataset.id}" ${dataset.id === current ? 'selected' : ''}>${escapeHtml(dataset.title)}</option>`).join('');
}

function buildRecords() {
  state.records = state.datasets.flatMap((dataset) => dataset.rows.map((raw, rowIndex) => ({
    id: `${dataset.id}:${rowIndex}`,
    sourceId: dataset.id,
    source: dataset.title,
    extension: dataset.extension,
    rowNumber: rowIndex + 2,
    values: Object.fromEntries(STANDARD_FIELDS.map((field) => [field.key, dataset.map[field.key] ? raw[dataset.map[field.key]] ?? '' : '']))
  })));
}

function similarity(left, right) {
  const a = normalize(left), b = normalize(right);
  if (!a || !b) return 0;
  if (a === b) return 1;
  let previous = Array.from({ length: b.length + 1 }, (_, index) => index);
  for (let x = 1; x <= a.length; x += 1) {
    const current = [x];
    for (let y = 1; y <= b.length; y += 1) current[y] = Math.min(current[y - 1] + 1, previous[y] + 1, previous[y - 1] + (a[x - 1] === b[y - 1] ? 0 : 1));
    previous = current;
  }
  return 1 - previous[b.length] / Math.max(a.length, b.length);
}

const FIELD_WEIGHTS = { id: .30, email: .23, phone: .18, name: .15, address: .07, city: .04, date: .03 };
const isReliablePhone = (value) => {
  const phone = digits(value).slice(-8);
  return phone.length >= 7 && new Set(phone).size >= 3 && !/^(01234567|12345678|87654321|00000000)$/.test(phone);
};

function valueScore(left, right, field, mode) {
  const a = normalizedField(left, field), b = normalizedField(right, field);
  if (!a || !b) return null;
  if (mode === 'days') {
    const dateA = new Date(left), dateB = new Date(right);
    if (Number.isNaN(dateA.getTime()) || Number.isNaN(dateB.getTime())) return a === b ? 1 : 0;
    const days = Math.abs(dateA - dateB) / 86400000;
    return days <= 7 ? Math.max(.55, 1 - days / 15) : 0;
  }
  if (a === b) return 1;
  if (mode !== 'flexible') return 0;
  if (field === 'phone') return a.slice(-8) === b.slice(-8) ? 1 : 0;
  return similarity(a, b);
}

function compareRecordPair(left, right, compared, preferred) {
  let weighted = 0, weightTotal = 0, evidence = 0, strongMatches = 0;
  const scores = {};
  compared.forEach(({ field, mode }) => {
    const score = valueScore(left.values[field], right.values[field], field, mode);
    if (score === null) return;
    const weight = FIELD_WEIGHTS[field] || .05;
    weighted += score * weight;
    weightTotal += weight;
    evidence += 1;
    scores[field] = score;
    if (score >= .98 && (['id', 'email'].includes(field) || (field === 'phone' && isReliablePhone(left.values.phone) && isReliablePhone(right.values.phone)))) strongMatches += 1;
  });
  const score = weightTotal ? weighted / weightTotal : 0;
  const preferredMatch = Boolean(preferred && scores[preferred] >= .98);
  const exactIdentity = scores.id >= .98 || scores.email >= .98;
  const phoneSupported = scores.phone >= .98 && evidence >= 2 && isReliablePhone(left.values.phone) && isReliablePhone(right.values.phone);
  const related = preferredMatch || exactIdentity || phoneSupported || (evidence >= 2 && score >= .62);
  const idConflict = left.values.id && right.values.id && normalizedField(left.values.id, 'id') !== normalizedField(right.values.id, 'id');
  const emailConflict = left.values.email && right.values.email && normalize(left.values.email) !== normalize(right.values.email);
  const conflict = related && (idConflict || emailConflict) && (strongMatches > 0 || score >= .55);
  return { score, evidence, related, conflict };
}

function blockingKeys(record, compared) {
  const enabled = new Set(compared.map(({ field }) => field));
  const keys = [];
  const id = normalizedField(record.values.id, 'id');
  const email = normalize(record.values.email);
  const phone = digits(record.values.phone);
  const name = normalize(record.values.name).replace(/[^a-z0-9 ]/g, '');
  const city = normalize(record.values.city);
  const address = normalize(record.values.address).replace(/[^a-z0-9 ]/g, '');
  if (enabled.has('id') && id.length >= 4) keys.push(`id:${id}`);
  if (enabled.has('email') && email) {
    keys.push(`email:${email}`);
    const local = email.split('@')[0];
    if (local.length >= 6) keys.push(`email-local:${local}`);
  }
  if (enabled.has('phone') && isReliablePhone(phone)) keys.push(`phone:${phone.slice(-8)}`);
  if (enabled.has('name') && name.length >= 4) {
    const compact = name.replace(/\s/g, '');
    keys.push(`name:${name}`);
    if (compact.length >= 8) keys.push(`name-prefix:${compact.slice(0, 8)}`);
    const tokens = name.split(' ').filter((token) => token.length >= 3);
    if (tokens.length) keys.push(`name-token:${tokens[tokens.length - 1]}`);
    if (city && enabled.has('city')) keys.push(`name-city:${compact.slice(0, 5)}:${city}`);
  }
  if (enabled.has('address') && address.length >= 8) keys.push(`address:${address.replace(/\s/g, '').slice(0, 14)}`);
  return [...new Set(keys)];
}

function groupRecords() {
  const compared = comparisonConfig();
  const preferred = $('primaryKey').value;
  const parent = state.records.map((_, index) => index);
  const find = (index) => { while (parent[index] !== index) { parent[index] = parent[parent[index]]; index = parent[index]; } return index; };
  const unite = (left, right) => { const a = find(left), b = find(right); if (a !== b) parent[b] = a; };
  const buckets = new Map();
  state.records.forEach((record, index) => blockingKeys(record, compared).forEach((key) => {
    if (!buckets.has(key)) buckets.set(key, []);
    buckets.get(key).push(index);
  }));

  const candidatePairs = new Set();
  buckets.forEach((bucket) => {
    if (bucket.length > 120) return;
    for (let left = 0; left < bucket.length; left += 1) {
      for (let right = left + 1; right < bucket.length; right += 1) {
        const a = bucket[left], b = bucket[right];
        if (state.records[a].sourceId === state.records[b].sourceId) continue;
        candidatePairs.add(a < b ? `${a}:${b}` : `${b}:${a}`);
      }
    }
  });

  const linked = [];
  candidatePairs.forEach((pair) => {
    const [left, right] = pair.split(':').map(Number);
    const assessment = compareRecordPair(state.records[left], state.records[right], compared, preferred);
    if (!assessment.related) return;
    unite(left, right);
    linked.push({ left, right, ...assessment });
  });

  const components = new Map();
  state.records.forEach((record, index) => {
    const root = find(index);
    if (!components.has(root)) components.set(root, []);
    components.get(root).push(index);
  });
  const edgesByRoot = new Map();
  linked.forEach((edge) => {
    const root = find(edge.left);
    if (!edgesByRoot.has(root)) edgesByRoot.set(root, []);
    edgesByRoot.get(root).push(edge);
  });
  state.groups = [...components.entries()].map(([root, indices], groupIndex) => {
    const records = indices.map((index) => state.records[index]);
    const edges = edgesByRoot.get(root) || [];
    const canonical = Object.fromEntries(STANDARD_FIELDS.map((field) => [field.key, records.find((record) => String(record.values[field.key]).trim())?.values[field.key] || '']));
    let status = 'Diferente';
    if (records.length > 1) {
      if (edges.some((edge) => edge.conflict)) status = 'Conflicto';
      else if (edges.length && edges.every((edge) => edge.score >= .9)) status = 'Exacto';
      else status = 'Posible';
    }
    const averageScore = edges.length ? edges.reduce((sum, edge) => sum + edge.score, 0) / edges.length : 0;
    return { id: crypto.randomUUID(), key: preferred ? canonical[preferred] : canonical.id || canonical.email || '', records, canonical, status, confidence: averageScore, sources: [...new Set(records.map((record) => record.source))], number: groupIndex + 1 };
  });
}

function fieldHasDifference(records, field, mode) {
  const values = records.map((record) => record.values[field]).filter((value) => String(value).trim());
  if (values.length < 2) return false;
  if (mode === 'days') {
    const dates = values.map((value) => new Date(value)).filter((date) => !Number.isNaN(date.getTime()));
    if (dates.length < 2) return new Set(values.map(normalize)).size > 1;
    return (Math.max(...dates.map(Number)) - Math.min(...dates.map(Number))) / 86400000 > 7;
  }
  if (mode === 'flexible') {
    if (field === 'phone') return new Set(values.map((value) => digits(value).slice(-8))).size > 1;
    const reference = values[0];
    return values.slice(1).some((value) => similarity(reference, value) < .82);
  }
  return new Set(values.map((value) => normalizedField(value, field))).size > 1;
}

function runComparison() {
  if (state.datasets.length < 2) return;
  if (!comparisonConfig().length) {
    $('runStatus').textContent = 'Selecciona al menos un campo para realizar el análisis.';
    return;
  }
  $('runStatus').textContent = 'Analizando y agrupando registros…';
  requestAnimationFrame(() => {
    buildRecords();
    if (!state.records.length) {
      $('runStatus').textContent = 'Los archivos no contienen registros para comparar.';
      return;
    }
    groupRecords();
    state.filter = 'all'; state.sourceFilter = ''; state.search = ''; state.page = 0; state.selected = state.groups[0]?.id || null; state.detailView = 'grouped';
    state.lastRun = new Date();
    $('search').value = '';
    $('sourceFilter').value = '';
    document.querySelectorAll('.detail-tabs button').forEach((button) => button.classList.toggle('active', button.dataset.detailView === 'grouped'));
    renderResults();
    $('runStatus').textContent = `Listo: ${formatNumber(state.records.length)} registros analizados sin exigir una clave común.`;
    goToStep(4);
  });
}

function filteredGroups() {
  const needle = normalize(state.search);
  return state.groups.filter((group) =>
    (state.filter === 'all' || group.status === state.filter)
    && (!state.sourceFilter || group.records.some((record) => record.sourceId === state.sourceFilter))
    && (!needle || normalize(Object.values(group.canonical).join(' ') + group.sources.join(' ')).includes(needle))
  );
}

function renderResults() {
  const counts = Object.fromEntries(['Exacto', 'Posible', 'Diferente', 'Conflicto'].map((status) => [status, state.groups.filter((group) => group.status === status).length]));
  $('countAll').textContent = formatNumber(state.groups.length);
  $('countExact').textContent = formatNumber(counts.Exacto);
  $('countPossible').textContent = formatNumber(counts.Posible);
  $('countExclusive').textContent = formatNumber(counts.Diferente);
  $('countConflict').textContent = formatNumber(counts.Conflicto);
  $('resultsSummary').textContent = state.groups.length ? `Se analizaron ${formatNumber(state.records.length)} registros · ${formatNumber(state.groups.length)} grupos resultantes` : 'Carga archivos para comenzar el análisis.';
  $('search').disabled = !state.groups.length;
  $('filterButton').disabled = !state.groups.length;
  $('exportCsv').disabled = !state.groups.length;
  $('exportXlsx').disabled = !state.groups.length;
  $('exportSummary').textContent = state.groups.length ? `${formatNumber(state.records.length)} registros quedaron organizados en ${formatNumber(state.groups.length)} grupos. Puedes exportar el consolidado completo.` : 'Ejecuta una comparación para habilitar las descargas.';
  if (!state.groups.length) {
    $('filterPopover').hidden = true;
    $('filterButton').setAttribute('aria-expanded', 'false');
  }

  const all = filteredGroups();
  const pageSize = 8;
  const pages = Math.max(1, Math.ceil(all.length / pageSize));
  state.page = Math.min(state.page, pages - 1);
  const visible = all.slice(state.page * pageSize, (state.page + 1) * pageSize);
  if (!visible.length) {
    $('resultsBody').innerHTML = `<tr class="empty-row"><td colspan="8"><svg aria-hidden="true"><use href="#i-compare"></use></svg><strong>${state.groups.length ? 'No hay coincidencias con este filtro' : 'Los resultados aparecerán aquí'}</strong><span>${state.groups.length ? 'Prueba otra búsqueda o categoría.' : 'Agrega tus archivos, revisa el mapeo y ejecuta la comparación.'}</span></td></tr>`;
  } else {
    $('resultsBody').innerHTML = visible.map((group) => `<tr data-group="${group.id}" class="${state.selected === group.id ? 'selected' : ''}"><td><span class="status-chip ${group.status}">${group.status}</span></td><td>${escapeHtml(group.canonical.id || '—')}</td><td>${escapeHtml(group.canonical.name || '—')}</td><td>${escapeHtml(group.canonical.email || '—')}</td><td>${escapeHtml(group.canonical.phone || '—')}</td><td>${escapeHtml(group.canonical.city || '—')}</td><td><span class="source-dots">${group.sources.slice(0, 3).map((source, index) => `<span class="source-dot" title="${escapeHtml(source)}">${index + 1}</span>`).join('')}</span></td><td><button class="row-action" type="button" aria-label="Ver detalle">•••</button></td></tr>`).join('');
    $('resultsBody').querySelectorAll('[data-group]').forEach((row) => row.addEventListener('click', () => {
      state.selected = row.dataset.group; renderResults();
    }));
  }
  $('paginationText').textContent = all.length ? `Mostrando ${state.page * pageSize + 1}–${Math.min((state.page + 1) * pageSize, all.length)} de ${formatNumber(all.length)} registros` : 'Mostrando 0 registros';
  $('previousPage').disabled = state.page === 0;
  $('nextPage').disabled = state.page >= pages - 1;
  $('pageButtons').innerHTML = `<b>${state.page + 1}</b>${pages > 1 ? `<span>de ${pages}</span>` : ''}`;
  renderDetail();
}

function renderDetail() {
  const group = state.groups.find((item) => item.id === state.selected);
  if (!group) {
    $('groupNumber').textContent = 'Sin selección';
    $('detailContent').innerHTML = '<div class="detail-empty"><svg aria-hidden="true"><use href="#i-users"></use></svg><strong>Selecciona un resultado</strong><p>Aquí verás los registros de cada fuente y podrás revisar sus diferencias.</p></div>';
    return;
  }
  $('groupNumber').textContent = `Grupo #${group.number}`;
  if (state.detailView === 'history') {
    const fields = comparisonConfig().map(({ field }) => fieldByKey(field)?.label).filter(Boolean).join(', ') || 'Ninguno';
    $('detailContent').innerHTML = `<div class="detail-summary"><div><h3>Resumen del análisis</h3><small>Trazabilidad de esta agrupación durante la sesión actual</small></div></div><article class="history-card"><div><span>Procesado</span><b>${state.lastRun ? state.lastRun.toLocaleString('es') : '—'}</b></div><div><span>Identificador preferido</span><b>${escapeHtml(fieldByKey($('primaryKey').value)?.label || 'Automático')}</b></div><div><span>Campos comparados</span><b>${escapeHtml(fields)}</b></div><div><span>Confianza media</span><b>${group.records.length > 1 ? `${Math.round(group.confidence * 100)} %` : 'Sin relación detectada'}</b></div><div><span>Fuentes vinculadas</span><b>${escapeHtml(group.sources.join(' · '))}</b></div><div><span>Resultado</span><b>${group.status}</b></div></article>`;
    return;
  }
  const title = state.detailView === 'individual' ? 'Registros individuales' : `${escapeHtml(group.canonical.id || 'Sin ID')} – ${escapeHtml(group.canonical.name || 'Sin nombre')}`;
  $('detailContent').innerHTML = `<div class="detail-summary"><div><h3><span class="status-chip ${group.status}">${group.status}</span> ${title}</h3><small>${group.records.length} ${group.records.length === 1 ? 'registro' : 'registros'} en este grupo</small></div></div>${group.records.map((record) => {
    const type = ['xlsx', 'xls'].includes(record.extension) ? 'XLS' : record.extension.toUpperCase();
    const iconClass = record.extension === 'pdf' ? 'pdf' : ['csv', 'tsv'].includes(record.extension) ? 'csv' : record.extension === 'txt' ? 'txt' : '';
    const visibleFields = state.detailView === 'individual' ? STANDARD_FIELDS : STANDARD_FIELDS.filter((field) => ['name', 'email', 'phone', 'address'].includes(field.key));
    return `<article class="detail-source"><span class="file-icon ${iconClass}">${type}</span><strong title="${escapeHtml(record.source)}">${escapeHtml(record.source)}</strong><div class="detail-fields">${visibleFields.map((field) => `<span>${field.label}</span><b>${escapeHtml(record.values[field.key] || '—')}</b>`).join('')}</div></article>`;
  }).join('')}`;
}

function consolidatedRows() {
  return state.groups.map((group) => ({
    estado: group.status,
    id_cliente: group.canonical.id,
    nombre: group.canonical.name,
    email: group.canonical.email,
    telefono: group.canonical.phone,
    direccion: group.canonical.address,
    ciudad: group.canonical.city,
    fecha_registro: group.canonical.date,
    fuentes: group.sources.join(' / '),
    registros_agrupados: group.records.length
  }));
}

function csvText(rows) {
  const headers = Object.keys(rows[0] || {});
  return '\ufeff' + [headers, ...rows.map((row) => headers.map((header) => row[header] ?? ''))]
    .map((row) => row.map((cell) => `"${String(cell).replace(/"/g, '""')}"`).join(','))
    .join('\r\n');
}

function download(blob, filename) {
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url; link.download = filename; link.click();
  setTimeout(() => URL.revokeObjectURL(url), 500);
}

async function exportExcel() {
  await loadScript('https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js', () => Boolean(window.XLSX));
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(consolidatedRows()), 'Consolidado');
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(state.records.map((record) => ({ archivo: record.source, fila: record.rowNumber, ...record.values }))), 'Registros originales');
  XLSX.writeFile(workbook, 'CompClick_resultados.xlsx');
}

function showPreview() {
  const dataset = state.datasets[0];
  if (!dataset) return;
  $('previewCaption').textContent = `${dataset.title} · primeras ${Math.min(20, dataset.rows.length)} filas`;
  $('previewHead').innerHTML = `<tr>${dataset.headers.map((header) => `<th>${escapeHtml(header)}</th>`).join('')}</tr>`;
  $('previewBody').innerHTML = dataset.rows.slice(0, 20).map((row) => `<tr>${dataset.headers.map((header) => `<td>${escapeHtml(row[header] || '')}</td>`).join('')}</tr>`).join('') || `<tr><td colspan="${Math.max(1, dataset.headers.length)}">El archivo no contiene registros.</td></tr>`;
  $('previewDialog').showModal();
}

function reachableStep() {
  if (state.groups.length) return 5;
  if (state.datasets.length >= 2) return 3;
  if (state.datasets.length) return 2;
  return 1;
}

function updateProgress() {
  document.querySelectorAll('.progress-step').forEach((button) => {
    const step = Number(button.dataset.step);
    button.classList.toggle('active', step === state.currentStep);
    button.classList.toggle('completed', step < state.currentStep && step <= reachableStep());
  });
  document.querySelectorAll('.flow-screen').forEach((screen) => screen.classList.toggle('active', Number(screen.dataset.flowScreen) === state.currentStep));
  document.querySelectorAll('.nav-item').forEach((item) => {
    const step = Number(item.dataset.goStep);
    const active = (state.currentStep === 1 && step === 1) || (state.currentStep === 2 && step === 2) || (state.currentStep === 3 && step === 3) || (state.currentStep >= 4 && step === 4);
    item.classList.toggle('active', active);
    if (active) item.setAttribute('aria-current', 'page'); else item.removeAttribute('aria-current');
  });
  $('flowBack').disabled = state.currentStep === 1;
  $('flowNext').hidden = state.currentStep === 5;
  const labels = { 1: 'Siguiente: revisar mapeo →', 2: 'Siguiente: configurar análisis →', 3: 'Analizar archivos →', 4: 'Siguiente: exportar →' };
  $('flowNext').textContent = labels[state.currentStep] || 'Siguiente →';
  $('flowNext').disabled = (state.currentStep === 1 && state.datasets.length < 2) || (state.currentStep === 3 && state.datasets.length < 2) || (state.currentStep === 4 && !state.groups.length);
  const hints = {
    1: state.datasets.length < 2 ? 'Carga al menos dos fuentes para comenzar.' : `${state.datasets.length} fuentes listas para revisar.`,
    2: 'Los campos sin asignar se omiten; no bloquean el análisis.',
    3: 'La comparación usa el registro completo y conserva lo no relacionado como diferente.',
    4: 'Selecciona una fila para revisar la evidencia de cada grupo.',
    5: 'Las descargas se generan localmente en tu navegador.'
  };
  $('flowHint').textContent = hints[state.currentStep];
}

function goToStep(requested) {
  const allowed = reachableStep();
  const target = Math.max(1, Math.min(Number(requested) || 1, allowed));
  state.currentStep = target;
  updateProgress();
  document.querySelector('main').scrollIntoView({ behavior: 'smooth', block: 'start' });
}

$('files').addEventListener('change', (event) => addFiles([...event.target.files]));
const dropZone = $('dropZone');
['dragenter', 'dragover'].forEach((eventName) => dropZone.addEventListener(eventName, (event) => { event.preventDefault(); dropZone.classList.add('dragging'); }));
['dragleave', 'drop'].forEach((eventName) => dropZone.addEventListener(eventName, (event) => { event.preventDefault(); dropZone.classList.remove('dragging'); }));
dropZone.addEventListener('drop', (event) => addFiles([...event.dataTransfer.files]));

$('selectAll').addEventListener('change', (event) => {
  event.target.indeterminate = false;
  document.querySelectorAll('[data-compare]').forEach((input) => { input.checked = event.target.checked; });
});
$('compareFields').addEventListener('change', (event) => {
  if (!event.target.matches('[data-compare]')) return;
  const checks = [...document.querySelectorAll('[data-compare]')];
  const checked = checks.filter((input) => input.checked).length;
  $('selectAll').checked = checked === checks.length;
  $('selectAll').indeterminate = checked > 0 && checked < checks.length;
});
$('autoMapAll').addEventListener('click', () => {
  state.datasets.forEach(autoMap);
  renderMapping();
  $('loadStatus').textContent = 'Mapeo automático actualizado.';
});
$('previewData').addEventListener('click', showPreview);
$('closePreview').addEventListener('click', () => $('previewDialog').close());
$('previewDialog').addEventListener('click', (event) => { if (event.target === $('previewDialog')) $('previewDialog').close(); });
$('runCompare').addEventListener('click', runComparison);
$('search').addEventListener('input', (event) => { state.search = event.target.value; state.page = 0; renderResults(); });
$('filterButton').addEventListener('click', () => {
  const willOpen = $('filterPopover').hidden;
  $('filterPopover').hidden = !willOpen;
  $('filterButton').setAttribute('aria-expanded', String(willOpen));
});
$('sourceFilter').addEventListener('change', (event) => { state.sourceFilter = event.target.value; state.page = 0; renderResults(); });
$('clearFilters').addEventListener('click', () => {
  state.filter = 'all'; state.sourceFilter = ''; state.search = ''; state.page = 0;
  $('search').value = ''; $('sourceFilter').value = ''; $('filterPopover').hidden = true; $('filterButton').setAttribute('aria-expanded', 'false');
  document.querySelectorAll('.status-tab').forEach((item) => item.classList.toggle('active', item.dataset.filter === 'all'));
  renderResults();
});
$('statusTabs').addEventListener('click', (event) => {
  const button = event.target.closest('[data-filter]');
  if (!button) return;
  state.filter = button.dataset.filter; state.page = 0;
  document.querySelectorAll('.status-tab').forEach((item) => item.classList.toggle('active', item === button));
  renderResults();
});
$('previousPage').addEventListener('click', () => { if (state.page > 0) { state.page -= 1; renderResults(); } });
$('nextPage').addEventListener('click', () => { state.page += 1; renderResults(); });
$('exportCsv').addEventListener('click', () => {
  download(new Blob([csvText(consolidatedRows())], { type: 'text/csv;charset=utf-8' }), 'CompClick_resultados.csv');
});
$('exportXlsx').addEventListener('click', () => {
  exportExcel().catch((error) => alert(error.message));
});

document.querySelectorAll('[data-go-step]').forEach((button) => button.addEventListener('click', () => goToStep(button.dataset.goStep)));
$('flowBack').addEventListener('click', () => goToStep(state.currentStep - 1));
$('flowNext').addEventListener('click', () => {
  if (state.currentStep === 3) runComparison();
  else goToStep(state.currentStep + 1);
});

document.querySelectorAll('.detail-tabs button').forEach((button) => button.addEventListener('click', () => {
  document.querySelectorAll('.detail-tabs button').forEach((item) => item.classList.toggle('active', item === button));
  state.detailView = button.dataset.detailView;
  renderDetail();
}));

$('themeToggle').addEventListener('click', () => {
  document.body.classList.toggle('light');
  localStorage.setItem('compclick-theme', document.body.classList.contains('light') ? 'light' : 'dark');
});
if (localStorage.getItem('compclick-theme') === 'light') document.body.classList.add('light');

renderFiles();
renderMapping();
renderCompareControls();
renderSourceFilter();
renderResults();
goToStep(1);
