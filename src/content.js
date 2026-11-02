// Content conversion: OneNote XHTML <-> Markdown.
//
// Markdown is the interchange format because OneNote's own HTML is both
// expensive to write and expensive to read back: textContent welds blocks
// together, and hand-written inline styles cost ~8x the tokens.
//
// Output is limited to the documented supported subset — see
// https://learn.microsoft.com/en-us/graph/onenote-input-output-html
// Anything OneNote silently drops (pre, code, blockquote, hr, th, rowspan,
// CSS classes, table borders) is mapped onto something it keeps.
import { marked } from 'marked';
import { parseHtml, collapse, normText, listOf } from './dom.js';

// Image marker for reads. OneNote images are authenticated resources that
// cannot be re-submitted, so this is deliberately not a working URL.
export const IMAGE_PLACEHOLDER = 'onenote:image';

// ---------------------------------------------------------------------------
// House style presets.
//
// Table borders are not a style choice: a border on <table> or <td> comes back
// as border:0px, so shading is the only way to delimit rows. Headings are:
// send no style and OneNote applies its own look per level.
// ---------------------------------------------------------------------------

const HOUSE_PRESET = {
  h1: 'font-size:16pt;color:#1e4e79;margin-top:0pt;margin-bottom:0pt',
  h2: 'font-size:14pt;color:#2e75b5;margin-top:0pt;margin-bottom:0pt',
  h3: 'font-size:12pt;color:#2e75b5;font-weight:bold;margin-top:0pt;margin-bottom:0pt',
  // font-style:normal is required: OneNote forces italic on h4/h6 elements and
  // demotes this object's style to a wrapping <span>, which cannot un-italicise
  // its ancestor.
  h4: 'font-size:11pt;color:#595959;font-weight:bold;font-style:normal;margin-top:0pt;margin-bottom:0pt',
  headerCell: 'background-color:#d9e2f3;font-weight:bold',
  cellEven: 'background-color:white',
  cellOdd: 'background-color:#f2f2f2',
  code: 'font-family:Consolas;font-size:10pt',
  codeInline: 'font-family:Consolas;background-color:#f2f2f2',
  quote: 'color:#595959;font-style:italic',
  lastListItem: 'margin-bottom:12pt',
};

// Empty heading styles are the point: OneNote then applies its own look, h3-h6
// get no wrapping <span>, and there is less to strip back out on read.
const NATIVE_PRESET = { ...HOUSE_PRESET, h1: '', h2: '', h3: '', h4: '' };

const STYLE_PRESETS = { house: HOUSE_PRESET, native: NATIVE_PRESET };

// An unknown name is a startup error, not a silent fallback.
const REQUESTED_STYLE = process.env.ONENOTE_STYLE || 'house';
if (!STYLE_PRESETS[REQUESTED_STYLE]) {
  throw new Error(
    `Unknown ONENOTE_STYLE "${REQUESTED_STYLE}". Choose one of: ${Object.keys(STYLE_PRESETS).join(', ')}.`
  );
}

/** The active style. Every writer reads this. */
export const HOUSE = STYLE_PRESETS[REQUESTED_STYLE];

export const escapeHtml = (s) =>
  String(s)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');

/** Wrap plain text in paragraphs; blank line = new paragraph. */
export const wrapText = (text) =>
  String(text)
    .split(/\n{2,}/)
    .map((par) => `<p>${escapeHtml(par).replace(/\n/g, '<br />')}</p>`)
    .join('');

export const slugify = (s) =>
  String(s).toLowerCase().normalize('NFKD').replace(/[\u0300-\u036f]/g, '')
    .replace(/[^a-z0-9]+/g, '-').replace(/^-+|-+$/g, '').slice(0, 40) || 'section';

// jsdom emits HTML5 void tags; OneNote wants XHTML. Idempotent, so every write
// can funnel through it.
const VOID_TAGS = /<(br|img|hr|meta|input|col|area|base|link)\b([^>]*?)\s*\/?>/gi;
export const toXhtml = (html) => String(html).replace(VOID_TAGS, (_m, tag, attrs) => `<${tag}${attrs} />`);

// ===========================================================================
// HTML -> Markdown  (reading a page)
// ===========================================================================

const styleOf = (el) => (el.getAttribute && el.getAttribute('style')) || '';
const hasStyle = (el, re) => re.test(styleOf(el));

const IS_BOLD = (el, t) => t === 'b' || t === 'strong' || hasStyle(el, /font-weight:\s*bold/i);
const IS_ITALIC = (el, t) => t === 'i' || t === 'em' || t === 'cite' || hasStyle(el, /font-style:\s*italic/i);
const IS_UNDER = (el, t) => t === 'u' || hasStyle(el, /text-decoration:\s*underline/i);
const IS_STRIKE = (el, t) => ['del', 's', 'strike'].includes(t) || hasStyle(el, /line-through/i);
const IS_CODE = (el) => hasStyle(el, /font-family:\s*(consolas|courier|monospace)/i);

// Escape only what would otherwise be read back as Markdown syntax.
const escapeMd = (s) => s.replace(/([\\`*_[\]])/g, '\\$1');

function inlineToMd(node, opts) {
  const TEXT_NODE = 3, ELEMENT_NODE = 1;
  if (node.nodeType === TEXT_NODE) return escapeMd(node.textContent);
  if (node.nodeType !== ELEMENT_NODE) return '';

  const tag = node.tagName.toLowerCase();
  // MathML round-trips verbatim: Markdown cannot represent it and OneNote
  // renders it natively.
  if (tag === 'math') return node.outerHTML;
  if (tag === 'br') return '\n';
  if (tag === 'img') {
    const alt = node.getAttribute('alt') || 'image';
    return `![${escapeMd(alt)}](${IMAGE_PLACEHOLDER})`;
  }
  if (tag === 'object') return `[file: ${escapeMd(node.getAttribute('data-attachment') || 'attachment')}]`;
  if (tag === 'iframe') return `[video: ${node.getAttribute('data-original-src') || ''}]`;

  const inner = [...node.childNodes].map((c) => inlineToMd(c, opts)).join('');
  if (!inner.trim()) return inner;

  if (tag === 'a') {
    const href = node.getAttribute('href');
    return href ? `[${inner}](${href})` : inner;
  }
  if (IS_CODE(node)) return '`' + inner.replace(/\\([\\`*_[\]])/g, '$1') + '`';
  if (tag === 'sup') return `<sup>${inner}</sup>`;
  if (tag === 'sub') return `<sub>${inner}</sub>`;

  // Order matters: bold outermost so **_x_** nests predictably.
  let out = inner;
  if (IS_STRIKE(node, tag)) out = `~~${out}~~`;
  if (IS_UNDER(node, tag)) out = `<u>${out}</u>`;
  if (IS_ITALIC(node, tag)) out = `*${out}*`;
  if (IS_BOLD(node, tag)) out = `**${out}**`;
  return out;
}

const inlineOf = (el, opts) => [...el.childNodes].map((c) => inlineToMd(c, opts)).join('').trim();

const rowsOf = (table) => [...table.querySelectorAll(':scope > tr, :scope > tbody > tr')];
const cellsOf = (tr) => [...tr.children].filter((c) => /^(td|th)$/i.test(c.tagName));

// OneNote check boxes are data-tag="to-do" / "to-do:completed". The attribute
// stays on a <p>, but a list item's tag is moved onto a <span> inside it.
// https://learn.microsoft.com/en-us/graph/onenote-note-tags
const todoBox = (el) => {
  const tag = el.getAttribute('data-tag') || el.querySelector('[data-tag]')?.getAttribute('data-tag');
  if (!tag || !/^to-do\b/.test(tag)) return null;
  return tag.includes(':completed') ? '[x]' : '[ ]';
};

// OneNote has ~20 note tags besides to-do, each an icon in the app. None has a
// Markdown equivalent, so surface them read-only: writing "[tag: x]" back does
// not attach a tag, the same way "[file: …]" does not create an attachment.
// https://learn.microsoft.com/en-us/graph/onenote-note-tags
const otherTag = (el) => {
  const tag = el.getAttribute('data-tag');
  return tag && !/^to-do\b/.test(tag) ? tag : null;
};

function tableToMd(table, opts) {
  const rows = rowsOf(table);
  if (!rows.length) return '';
  const width = Math.max(...rows.map((r) => cellsOf(r).length));
  const render = (tr) => {
    const cells = cellsOf(tr).map((td) => inlineOf(td, opts).replace(/\|/g, '\\|').replace(/\n/g, '<br />'));
    while (cells.length < width) cells.push('');
    return `| ${cells.join(' | ')} |`;
  };
  // OneNote has no <th>: row 0 is the header by convention. Its cells are bold,
  // but the header row already conveys that, so keeping the ** would round-trip
  // as doubly-bold for nothing.
  const header = render(rows[0]).replace(/\*\*([^*|]+)\*\*/g, '$1');
  const out = [header, `| ${Array(width).fill('---').join(' | ')} |`];
  for (const tr of rows.slice(1)) out.push(render(tr));
  return out.join('\n');
}

function listToMd(list, opts, depth = 0) {
  const ordered = list.tagName.toLowerCase() === 'ol';
  const pad = '  '.repeat(depth);
  const out = [];
  let n = 1;
  for (const li of [...list.children].filter((c) => c.tagName.toLowerCase() === 'li')) {
    const nested = [...li.children].filter((c) => /^(ul|ol)$/i.test(c.tagName));
    const clone = li.cloneNode(true);
    [...clone.children].filter((c) => /^(ul|ol)$/i.test(c.tagName)).forEach((c) => c.remove());
    const text = inlineOf(clone, opts).replace(/\n/g, ' ');
    const box = todoBox(li);
    out.push(`${pad}${box ? `- ${box}` : ordered ? `${n++}.` : '-'} ${text}`);
    for (const sub of nested) out.push(listToMd(sub, opts, depth + 1));
  }
  return out.join('\n');
}

function blockToMd(el, opts) {
  const tag = el.tagName.toLowerCase();
  // A display equation sits at block level; round-trip it verbatim.
  if (tag === 'math') return el.outerHTML;
  if (/^h[1-6]$/.test(tag)) {
    const text = inlineOf(el, opts);
    if (!text) return '';
    // OneNote keeps the heading's bold as a wrapping <span>, which would read
    // back as "### **Text**". Strip only a wrap around the whole text: bold on
    // part of a heading is real emphasis.
    const wholly = text.match(/^\*\*([\s\S]+)\*\*$/);
    return `${'#'.repeat(Number(tag[1]))} ${wholly ? wholly[1] : text}`;
  }
  if (tag === 'table') return tableToMd(el, opts);
  if (tag === 'ul' || tag === 'ol') return listToMd(el, opts);
  if (tag === 'div') return blocksToMd(el, opts);
  if (tag === 'br') return '';
  if (tag === 'img' || tag === 'object' || tag === 'iframe') return inlineToMd(el, opts);
  const text = inlineOf(el, opts); // p, and anything else with text in it
  if (!text) return text;
  const box = todoBox(el);
  if (box) return `- ${box} ${text}`;
  const other = otherTag(el);
  return other ? `[tag: ${other}] ${text}` : text;
}

const TASK_LINE = /^- \[[ x]\] /;
const lastLineOf = (s) => s.slice(s.lastIndexOf('\n') + 1);

const blocksToMd = (container, opts) => {
  const parts = [...container.children].map((el) => blockToMd(el, opts)).filter((s) => s && s.trim());
  // Consecutive check boxes are one list, so they are joined without the blank
  // line that would read back as a loose list.
  return parts.reduce((acc, part, i) => {
    if (i === 0) return part;
    const glue = TASK_LINE.test(part) && TASK_LINE.test(lastLineOf(parts[i - 1])) ? '\n' : '\n\n';
    return acc + glue + part;
  }, '');
};

/** OneNote page HTML to Markdown. */
export async function htmlToMarkdown(html) {
  const doc = await parseHtml(html);
  const md = blocksToMd(doc.body, {});
  return md.replace(/\n{3,}/g, '\n\n').trim();
}

// Row labels are capped so the outline stays cheap.
const ROW_LABELS = { count: 40, chars: 40 };

/** Headings, table shapes and size — enough to aim an edit without reading the body. */
export async function outlineOf(html) {
  const doc = await parseHtml(html);

  const headings = [...doc.querySelectorAll('h1,h2,h3,h4,h5,h6')]
    .map((h) => ({ level: Number(h.tagName[1]), text: collapse(h.textContent) }))
    .filter((h) => h.text);

  const headingAbove = (el) => {
    for (let n = el.previousElementSibling; n; n = n.previousElementSibling) {
      if (/^h[1-6]$/i.test(n.tagName)) return collapse(n.textContent);
    }
    const parent = el.parentElement;
    return parent && parent !== doc.body ? headingAbove(parent) : null;
  };

  const tables = [...doc.querySelectorAll('table')].map((t, index) => {
    const { grid, hasHeader } = tableToGrid(t);
    const out = { index, rows: grid.length, columns: Math.max(0, ...grid.map((r) => r.length)) };
    const heading = headingAbove(t);
    if (heading) out.afterHeading = heading;
    if (hasHeader) out.headers = grid[0].map((c) => c.text);
    // The first column is what editTable addresses rows by.
    const labels = grid.slice(hasHeader ? 1 : 0)
      .slice(0, ROW_LABELS.count)
      .map((r) => (r[0] ? r[0].text : ''))
      .map((label) => (label.length > ROW_LABELS.chars ? label.slice(0, ROW_LABELS.chars) + '…' : label));
    if (labels.some(Boolean)) out.rowLabels = labels;
    return out;
  });

  const tags = [...doc.querySelectorAll('[data-tag]')].map((el) => el.getAttribute('data-tag'));
  const otherTags = tags.filter((t) => !/^to-do\b/.test(t));
  const out = {
    headings,
    tables,
    images: doc.querySelectorAll('img').length,
    equations: doc.querySelectorAll('math').length,
    checkboxes: tags.filter((t) => /^to-do\b/.test(t)).length,
    chars: collapse(doc.body.textContent).length,
  };
  // Per kind rather than a total: "3 important, 1 critical" says whether the
  // page is worth reading.
  if (otherTags.length) {
    out.tags = otherTags.reduce((acc, t) => ({ ...acc, [t]: (acc[t] || 0) + 1 }), {});
  }
  return out;
}

/** Plain text with block boundaries kept. */
export async function htmlToText(html) {
  const md = await htmlToMarkdown(html);
  return md
    .replace(/^#{1,6}\s+/gm, '')
    .replace(/\*\*|~~|`/g, '')
    .replace(/\\([\\`*_[\]])/g, '$1')
    .trim();
}

// ===========================================================================
// Markdown -> OneNote XHTML  (writing a page)
// ===========================================================================

/**
 * Build a table. A cell is a string (escaped) or { html } (verbatim), so a cell
 * carried over from an existing table keeps its formatting.
 */
export function renderTable({ headers, rows, dataId }) {
  const cell = (c, style) => {
    const inner = c == null ? '' : typeof c === 'object' ? c.html : escapeHtml(String(c));
    return `<td style="${style}">${inner}</td>`;
  };
  const out = [];
  if (headers && headers.length) {
    out.push(`<tr>${headers.map((h) => cell(h, HOUSE.headerCell)).join('')}</tr>`);
  }
  (rows || []).forEach((row, i) => {
    const style = i % 2 === 0 ? HOUSE.cellEven : HOUSE.cellOdd;
    out.push(`<tr>${row.map((c) => cell(c, style)).join('')}</tr>`);
  });
  // No style="border:..." — OneNote rejects the border property on input tables.
  return `<table${dataId ? ` data-id="${escapeHtml(dataId)}"` : ''}>${out.join('')}</table>`;
}

/**
 * Read a <table> as a grid of { html, text } cells, plus whether row 0 is a
 * header. Cells carry their HTML because OneNote cannot patch a <tr> or <td>:
 * an edit rewrites the table, and re-escaping would flatten it.
 */
export function tableToGrid(table) {
  const trs = rowsOf(table);
  const grid = trs.map((tr) =>
    cellsOf(tr).map((td) => ({ html: td.innerHTML, text: td.textContent.trim() }))
  );
  const isBold = (td) =>
    /font-weight:\s*bold/i.test(td.getAttribute('style') || '') ||
    /font-weight:\s*bold|<b[\s>]|<strong[\s>]/i.test(td.innerHTML);
  // Only call row 0 a header if it looks like one, so editing a headerless
  // table does not promote its first row.
  const firstRow = trs.length ? cellsOf(trs[0]) : [];
  const hasHeader = firstRow.length > 0 && firstRow.every(isBold);
  return { grid, hasHeader };
}

/**
 * A column by 0-based index or header name. Naming needs a header row: with
 * none there is nothing to match, and defaulting to column 0 would corrupt it.
 */
export function resolveColumn(ref, grid, hasHeader) {
  const width = Math.max(0, ...grid.map((r) => r.length));
  if (typeof ref === 'number') {
    if (ref < 0 || ref >= width) throw new Error(`Column ${ref} does not exist (table has ${width} column(s)).`);
    return ref;
  }
  if (!hasHeader) {
    throw new Error(`Column "${ref}" cannot be matched by name: this table has no header row. Use a 0-based index.`);
  }
  const headers = grid[0].map((c) => c.text);
  const needle = normText(ref);
  const match = (pred) => headers.map((h, i) => ({ h, i })).filter(({ h }) => pred(normText(h)));
  const exact = match((h) => h === needle);
  const hits = exact.length ? exact : match((h) => h.includes(needle));
  if (!hits.length) throw new Error(`No column header matches "${ref}". Headers: ${listOf(headers)}.`);
  if (hits.length > 1) {
    throw new Error(`"${ref}" matches ${hits.length} columns (${listOf(hits.map((x) => x.h))}). Use a 0-based index.`);
  }
  return hits[0].i;
}

/**
 * A row by 0-based index or by text in it. The first column is tried first, so
 * a value that appears elsewhere still selects the row it labels.
 */
export function resolveRow(ref, grid, hasHeader) {
  if (typeof ref === 'number') {
    if (ref < 0 || ref >= grid.length) throw new Error(`Row ${ref} does not exist (table has ${grid.length} row(s)).`);
    return ref;
  }
  const needle = normText(ref);
  const body = grid
    .map((r, i) => ({ i, label: r[0] ? r[0].text : '', all: r.map((c) => c.text).join(' ') }))
    .slice(hasHeader ? 1 : 0);
  const exact = body.filter((r) => normText(r.label) === needle);
  const partial = body.filter((r) => normText(r.label).includes(needle));
  const hits = exact.length ? exact : partial.length ? partial : body.filter((r) => normText(r.all).includes(needle));
  if (!hits.length) throw new Error(`No row matches "${ref}". First column: ${listOf(body.map((r) => r.label))}.`);
  if (hits.length > 1) {
    throw new Error(`"${ref}" matches ${hits.length} rows (${listOf(hits.map((r) => r.label))}). Use a 0-based index.`);
  }
  return hits[0].i;
}

/**
 * Apply cell, row and column edits. References resolve against the table as the
 * caller read it, before any splice moves things. The result is rectangular.
 */
export function editGrid(grid, hasHeader, ops = {}) {
  const cell = (v) => ({ html: inlineMarkdown(v), text: String(v ?? '') });
  const rowRef = (r) => resolveRow(r, grid, hasHeader);
  const colRef = (c) => resolveColumn(c, grid, hasHeader);

  const sets = (ops.setCells || []).map((s) => ({ r: rowRef(s.row), c: colRef(s.col), value: s.value }));
  const dropRows = [...new Set((ops.deleteRows || []).map(rowRef))].sort((a, b) => b - a);
  const dropCols = [...new Set((ops.deleteColumns || []).map(colRef))].sort((a, b) => b - a);
  // Deletions run first and shift everything after them, so an insertion point
  // has to be carried across them or it lands N places too far right.
  const past = (at, dropped) => (at === null ? null : at - dropped.filter((d) => d < at).length);
  const atRow = past(ops.atRow === undefined ? null : rowRef(ops.atRow), dropRows);
  const atCol = past(ops.atCol === undefined ? null : colRef(ops.atCol), dropCols);

  const out = grid.map((r) => [...r]);
  for (const { r, c, value } of sets) {
    while (out[r].length <= c) out[r].push(cell(''));
    out[r][c] = cell(value);
  }
  for (const c of dropCols) for (const row of out) row.splice(c, 1);
  // Column values are positional against the rows, so columns go in before any
  // row is removed or inserted.
  (ops.addColumns || []).forEach((values, n) => {
    const width = Math.max(0, ...out.map((r) => r.length));
    const at = atCol === null ? width : Math.min(atCol + n, width);
    out.forEach((r, i) => r.splice(at, 0, cell(values[i] ?? '')));
  });
  for (const r of dropRows) out.splice(r, 1);
  if (ops.addRows && ops.addRows.length) {
    out.splice(atRow === null ? out.length : Math.min(atRow, out.length), 0,
      ...ops.addRows.map((r) => r.map(cell)));
  }
  const width = Math.max(0, ...out.map((r) => r.length));
  for (const r of out) while (r.length < width) r.push(cell(''));
  return { grid: out, width };
}

/** Map generic HTML onto the constructs OneNote accepts. */
function applyHouseStyle(doc) {
  const warnings = [];
  const q = (sel) => [...doc.querySelectorAll(sel)];

  // data-id survives updates (generated ids do not) and can be targeted later,
  // so two headings that slugify alike must not share one anchor.
  const anchors = new Set();
  for (const h of q('h1,h2,h3,h4,h5,h6')) {
    const level = Math.min(Number(h.tagName[1]), 4);
    // An empty preset value means "let OneNote decide" — that has to be no
    // attribute at all, not style="", which asks for no styling.
    const style = HOUSE['h' + level];
    if (style) h.setAttribute('style', style);
    let anchor = h.getAttribute('data-id');
    if (!anchor) {
      const base = 'h-' + slugify(h.textContent);
      anchor = base;
      for (let n = 2; anchors.has(anchor); n++) anchor = `${base}-${n}`;
      h.setAttribute('data-id', anchor);
    }
    anchors.add(anchor);
  }

  // <pre> is unsupported. Render monospace lines, keeping indentation as
  // non-breaking spaces, which HTML would otherwise collapse.
  for (const pre of q('pre')) {
    const text = pre.textContent.replace(/\n$/, '');
    const lines = text.split('\n').map((l) => {
      const body = escapeHtml(l).replace(/^ +/, (m) => '&#160;'.repeat(m.length));
      return `<p><span style="${HOUSE.code}">${body || '&#160;'}</span></p>`;
    });
    const holder = doc.createElement('div');
    holder.innerHTML = lines.join('');
    pre.replaceWith(...holder.childNodes);
  }
  for (const code of q('code')) {
    const span = doc.createElement('span');
    span.setAttribute('style', HOUSE.codeInline);
    span.textContent = code.textContent;
    code.replaceWith(span);
  }

  // <blockquote> is not supported input — flatten to styled paragraphs.
  for (const bq of q('blockquote')) {
    const holder = doc.createElement('div');
    for (const child of [...bq.children]) {
      const p = doc.createElement('p');
      p.innerHTML = `<span style="${HOUSE.quote}">${child.innerHTML}</span>`;
      holder.appendChild(p);
    }
    bq.replaceWith(...holder.childNodes);
  }

  // <hr> is not supported input — the closest supported thing is a blank line.
  for (const hr of q('hr')) hr.replaceWith(doc.createElement('p'));

  q('table').forEach((table, tableIndex) => {
    table.removeAttribute('style'); // border in style is rejected on input
    // OneNote drops rowspan/colspan without a word, which shifts every cell
    // after the merge into the wrong column.
    if (table.querySelector('[colspan],[rowspan]')) {
      warnings.push('OneNote does not support rowspan/colspan; merged cells were flattened.');
    }
    if (!table.getAttribute('data-id')) table.setAttribute('data-id', 't-' + (tableIndex + 1));
    // OneNote has no <th>: convert to a bold <td>.
    for (const th of [...table.querySelectorAll('th')]) {
      const td = doc.createElement('td');
      td.innerHTML = th.innerHTML;
      td.setAttribute('style', HOUSE.headerCell);
      th.replaceWith(td);
    }
    // <thead>/<tbody> carry no meaning for OneNote — unwrap them.
    for (const sec of [...table.querySelectorAll('thead,tbody,tfoot')]) sec.replaceWith(...sec.childNodes);
    const rows = [...table.querySelectorAll(':scope > tr')];
    // Striping is counted from the first BODY row, so a headerless table starts
    // on white here exactly as it does in renderTable.
    const headerRow = !!rows[0] && [...rows[0].children].every((c) => c.getAttribute('style') === HOUSE.headerCell);
    rows.forEach((tr, i) => {
      if (i === 0 && headerRow) return;
      const style = (i - (headerRow ? 1 : 0)) % 2 === 0 ? HOUSE.cellEven : HOUSE.cellOdd;
      for (const td of [...tr.children]) if (!td.getAttribute('style')) td.setAttribute('style', style);
    });
  });

  // Loose lists come out of Markdown as <li><p>…</p></li>; OneNote wants text.
  for (const li of q('li')) {
    for (const p of [...li.children].filter((c) => c.tagName.toLowerCase() === 'p')) {
      p.replaceWith(...p.childNodes);
    }
  }

  // A data-tag on a <ul>/<ol> overrides every child <li>, so a check box inside
  // a list is dropped. Paragraphs are also what Microsoft recommends.
  // https://learn.microsoft.com/en-us/graph/onenote-note-tags
  for (const list of q('ul,ol')) {
    const items = [...list.children].filter((c) => c.tagName.toLowerCase() === 'li');
    const boxes = items.map((li) => li.querySelector('input[type="checkbox"]'));
    if (!boxes.some(Boolean)) continue;
    for (const box of boxes) if (box) box.remove(); // <input> is not supported input HTML
    if (list.querySelector('ul,ol')) {
      // Flattening to paragraphs would lose the nesting, so keep the list and
      // say what happened rather than quietly dropping either one.
      warnings.push('Check boxes in a nested list are not supported; those items became plain bullets.');
      continue;
    }
    const holder = doc.createElement('div');
    items.forEach((li, i) => {
      const p = doc.createElement('p');
      p.innerHTML = li.innerHTML.replace(/^\s+/, '');
      if (boxes[i]) p.setAttribute('data-tag', boxes[i].hasAttribute('checked') ? 'to-do:completed' : 'to-do');
      holder.appendChild(p);
    });
    list.replaceWith(...holder.childNodes);
  }

  // Trailing space on the last item of an OUTERMOST list only: on a nested list
  // it opens a gap mid-list instead of after it.
  for (const list of q('ul,ol')) {
    if (list.closest('li')) continue;
    const items = [...list.children].filter((c) => c.tagName.toLowerCase() === 'li');
    const last = items[items.length - 1];
    if (last) last.setAttribute('style', HOUSE.lastListItem);
  }

  for (const img of q('img')) {
    const src = img.getAttribute('src') || '';
    if (src.startsWith(IMAGE_PLACEHOLDER) || /graph\.microsoft\.com/.test(src)) {
      warnings.push(`Image "${img.getAttribute('alt') || 'image'}" cannot be re-submitted this way.`);
    }
  }

  return [...new Set(warnings)];
}

// A run of consecutive lines that all look like table rows, outside any code
// fence — a fenced block full of pipes is a note ABOUT tables, not a table.
const FENCED = /^[ \t]*```[\s\S]*?^[ \t]*```/gm;
const TABLE_RUNS = /^[ \t]*\|.*\|[ \t]*(?:\r?\n[ \t]*\|.*\|[ \t]*)+$/gm;
const DELIMITER_CELL = /^:?-+:?$/;

/** The cells of one Markdown table row: split on pipes that are not escaped. */
const markdownCells = (line) =>
  line.trim().replace(/^\|/, '').replace(/\|$/, '').split(/(?<!\\)\|/);

/**
 * Warn about Markdown tables that will not survive the trip to the page.
 *
 * Both failures here are silent ones: GFM needs a "| --- |" row or the lines
 * are just a paragraph full of pipe characters, and it drops any cell that has
 * no column to go in. Either way the page quietly ends up missing what was
 * written, so the tables are checked against their SOURCE — by the time the
 * HTML exists, the dropped cells are already gone.
 */
function checkMarkdownTables(md) {
  const warnings = [];
  for (const [run] of md.replace(FENCED, '').matchAll(TABLE_RUNS)) {
    const lines = run.split(/\r?\n/).filter((l) => l.trim());
    const header = markdownCells(lines[0]);
    const delimiter = lines[1] ? markdownCells(lines[1]) : [];
    if (!delimiter.length || !delimiter.every((c) => DELIMITER_CELL.test(c.trim()))) {
      warnings.push(
        'Lines shaped like a table were not parsed as one: a Markdown table needs a ' +
        '"| --- | --- |" separator row under its header.'
      );
      continue;
    }
    if (delimiter.length !== header.length) {
      warnings.push(
        `A table header has ${header.length} columns but its "| --- |" row has ${delimiter.length}. ` +
        'GFM needs them to match, so those lines were not parsed as a table.'
      );
      continue;
    }
    const wide = lines.slice(2).filter((l) => markdownCells(l).length > header.length).length;
    if (wide) {
      warnings.push(
        `${wide} row(s) of a table have more cells than its ${header.length}-column header; ` +
        'the extra cells were dropped. Widen the header, or escape a literal "|" as "\\|".'
      );
    }
  }
  return warnings;
}

/**
 * Warn about two OneNote import quirks confirmed by writing minimal content
 * through the live Graph API: neither is something this server does to the
 * markup, and neither is caught by MathML validation, because the MathML
 * sent really is valid — OneNote just does not store it the way it arrived.
 *
 *   - Two display-block equations with nothing but whitespace between them
 *     are silently MERGED into one <math>, concatenating their content with
 *     no separator: "f(x)=x^r" next to "f'(x)=rx^(r-1)" becomes one
 *     nonsensical run-on formula.
 *   - A display-block equation inside a list item, at any nesting depth, is
 *     silently DUPLICATED: the original stays where it was written, and an
 *     identical copy is appended as a stray sibling right after the
 *     OUTERMOST list closes. Reading such a page back and writing it
 *     straight back writes the duplicate right back in, so it compounds on
 *     every "read, edit, write back" cycle instead of staying put.
 */
function checkMathPlacement(body) {
  const warnings = [];
  // Every <math>, not just display="block": OneNote promotes a standalone
  // equation to block on import either way, so filtering here only produced
  // false negatives on the shapes people actually write.
  const equations = [...body.querySelectorAll('math')];
  const isMath = (n) => n && n.nodeType === 1 && n.localName === 'math';
  const isBlank = (n) => n.nodeType === 3 && !n.textContent.trim();

  const adjacent = equations.some((m) => {
    let n = m.nextSibling;
    while (n && isBlank(n)) n = n.nextSibling;
    return isMath(n);
  });
  if (adjacent) {
    warnings.push(
      'Two equations sit next to each other with nothing between them: OneNote merges adjacent ' +
      'equations into one, concatenating their content with no separator. Put a word or a ' +
      'sentence between them.'
    );
  }

  // The trigger is the equation being the only content of the list item, at any
  // depth. With text beside it the equation stays put; alone, OneNote emits a
  // stray second copy outside the list.
  const aloneInItem = equations.some((m) => {
    const li = m.closest('li');
    if (!li) return false;
    const rest = li.cloneNode(true);
    // A nested list belongs to its own items, not to this one's content.
    for (const nested of rest.querySelectorAll('ul,ol')) nested.remove();
    for (const eq of rest.querySelectorAll('math')) eq.remove();
    return !rest.textContent.trim();
  });
  if (aloneInItem) {
    warnings.push(
      'An equation is the only content of a list item: OneNote duplicates it as a stray copy ' +
      'outside the list. Put text beside it in the item (e.g. "a) <equation>"), or move the ' +
      'equation out of the list.'
    );
  }
  return warnings;
}

/**
 * Render one line of inline Markdown — bold, links, code, raw MathML — as the
 * inside of a table cell. Block syntax is not applicable inside a <td>.
 */
export const inlineMarkdown = (text) =>
  marked.parseInline(String(text ?? ''), { gfm: true, async: false })
    .replace(/<code>([\s\S]*?)<\/code>/g, `<span style="${HOUSE.codeInline}">$1</span>`);

/**
 * Convert Markdown to OneNote-ready XHTML.
 * Raw HTML in the Markdown (including <math>) passes through untouched.
 */
export async function markdownToOneNote(md) {
  // Shield MathML from the Markdown parser: a blank line inside a <math> block
  // would split it across paragraphs.
  const maths = [];
  const shielded = String(md).replace(/<math[\s\S]*?<\/math>/gi, (m) => {
    maths.push(m);
    return `MATHBLOCK${maths.length - 1}ENDMATHBLOCK`;
  });

  const rawHtml = marked.parse(shielded, { gfm: true, breaks: true, async: false });
  const doc = await parseHtml(`<body>${rawHtml}</body>`);
  const warnings = [...applyHouseStyle(doc), ...checkMarkdownTables(shielded)];

  let html = toXhtml(doc.body.innerHTML);
  html = html.replace(/MATHBLOCK(\d+)ENDMATHBLOCK/g, (_m, i) => maths[Number(i)]);
  // marked wraps a standalone equation in its own <p>; OneNote hoists it out
  // anyway, so the wrapper is bytes that change nothing.
  html = html.replace(/<p>(\s*<math[\s\S]*?<\/math>\s*)<\/p>/gi, '$1');
  // Only the substituted-back html has real <math> elements to inspect — the
  // doc above still holds the shielded MATHBLOCK…ENDMATHBLOCK text tokens.
  if (maths.length) {
    const withMath = await parseHtml(`<body>${html}</body>`);
    warnings.push(...checkMathPlacement(withMath.body));
  }
  // Deduped: two malformed tables, or two mispositioned equations, in one
  // write share the same diagnosis.
  return { html, warnings: [...new Set(warnings)] };
}

/**
 * Resolve caller-supplied content of a declared format into OneNote XHTML.
 * `format` is explicit — the old "does it look like HTML?" sniff mis-fired on
 * ordinary prose such as "if x<y then y>x".
 */
export async function toOneNoteHtml(content, format = 'markdown') {
  if (content == null || content === '') return { html: '', warnings: [] };
  if (format === 'html') return { html: String(content), warnings: [] };
  if (format === 'text') return { html: wrapText(content), warnings: [] };
  return markdownToOneNote(content);
}

/**
 * Refuse a write that would silently delete images the caller never saw.
 * Read-edit-write of a page with images loses them, because OneNote images are
 * authenticated resource endpoints that cannot be re-submitted as www links.
 */
export function assertNoLostImages(newHtml, existingHtml, { allowImageLoss } = {}) {
  if (allowImageLoss) return;
  const existing = (existingHtml.match(/<img\b/gi) || []).length;
  if (!existing) return;
  const kept = (newHtml.match(/<img\b/gi) || []).length;
  if (kept >= existing) return;
  throw new Error(
    `This would drop ${existing - kept} of ${existing} image(s) on the page: OneNote images are ` +
    `authenticated resources that cannot be re-submitted in a write. Edit around them with ` +
    `insertNearText / replaceElement / editTable, re-add them with addImageToPage, or pass ` +
    `allowImageLoss:true if losing them is intended.`
  );
}
