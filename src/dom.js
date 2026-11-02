// jsdom cold-imports in ~3.5 s, so load it lazily: the server boots fast and
// only the first HTML-parsing call pays.
let _JSDOM = null;

export async function getJSDOM() {
  if (!_JSDOM) ({ JSDOM: _JSDOM } = await import('jsdom'));
  return _JSDOM;
}

/** Parse an HTML string and return its document. */
export async function parseHtml(html) {
  const JSDOM = await getJSDOM();
  return new JSDOM(html).window.document;
}

/** Parse a string as XML (strict). Returns the document; caller checks for <parsererror>. */
export async function parseXml(xml) {
  const JSDOM = await getJSDOM();
  return new JSDOM(xml, { contentType: 'application/xml' }).window.document;
}

// Navigating and editing a parsed page.
//
// OneNote cannot patch below block level: there is no delete action, and tr,
// td and span support no update actions. Anything finer than one block element
// has to be spliced here and written back.

/** OneNote wraps page content in a single absolutely-positioned div. */
export const contentContainer = (body) => {
  const divs = body.querySelectorAll(':scope > div');
  return divs.length === 1 ? divs[0] : body;
};

// ---------------------------------------------------------------------------
// Naming things on a page
//
// Every tool that takes a text fragment resolves it the same way: exact match
// beats a contains, and anything still ambiguous is an error listing the
// candidates rather than a write to whichever came first.
// ---------------------------------------------------------------------------

/** Collapse runs of whitespace, so a label reads the way the page does. */
export const collapse = (s) => String(s ?? '').replace(/\s+/g, ' ').trim();

/** …and fold case too, for matching a fragment against it. */
export const normText = (s) => collapse(s).toLowerCase();

/** Candidate labels for an error message, bounded so it stays readable. */
export const listOf = (labels) =>
  labels.slice(0, 12).map((l) => `"${l}"`).join(', ') + (labels.length > 12 ? ', …' : '');

/**
 * The one candidate a fragment names. Exactness wins: "Aufgaben" selects the
 * heading Aufgaben, not Aufgabenliste. Two matches are an error, never a coin
 * flip.
 */
export function pickOne(candidates, fragment, what, plural = `${what}s`) {
  const needle = normText(fragment);
  const exact = candidates.filter((c) => normText(c.text) === needle);
  const hits = exact.length ? exact : candidates.filter((c) => normText(c.text).includes(needle));
  if (!hits.length) {
    const known = candidates.length ? ` On this page: ${listOf(candidates.map((c) => c.text))}.` : '';
    throw new Error(`No ${what} matches "${fragment}".` + known);
  }
  if (hits.length > 1) {
    throw new Error(
      `"${fragment}" matches ${hits.length} ${plural} (${listOf(hits.map((c) => c.text))}). ` +
      'Use the full text, or edit with insertNearText / findAndReplace.'
    );
  }
  return hits[0];
}

const isHeading = (el) => /^h[1-6]$/i.test(el.tagName);

/**
 * The span of one heading's section: the heading, then siblings up to the next
 * heading of equal or higher level. Returns the sibling list and [start, end).
 */
export function sectionSpan(container, heading) {
  const nodes = [...container.children];
  const candidates = nodes
    .map((el, index) => ({ el, index, text: collapse(el.textContent) }))
    .filter((c) => isHeading(c.el));
  const { index: start } = pickOne(candidates, heading, 'heading');
  const level = Number(nodes[start].tagName[1]);
  let end = start + 1;
  while (end < nodes.length && !(isHeading(nodes[end]) && Number(nodes[end].tagName[1]) <= level)) end++;
  return { nodes, start, end };
}

/**
 * Put `html` into the section under `heading`, in place. The heading is kept
 * except by "delete": "replace" swaps the content below it, "prepend" adds
 * directly under it, "append" adds at the end of the section, and "delete"
 * removes the heading and the section together.
 */
export function spliceSection(doc, heading, html, mode = 'replace') {
  const { nodes, start, end } = sectionSpan(contentContainer(doc.body), heading);
  // OneNote has no delete action, so removing a section means removing the
  // nodes here and writing the body back — the same read-modify-write the
  // other modes already do.
  if (mode === 'delete') {
    for (let i = start; i < end; i++) nodes[i].remove();
    return;
  }
  const holder = doc.createElement('div');
  holder.innerHTML = html;
  const added = [...holder.childNodes];
  if (mode === 'append') {
    nodes[end - 1].after(...added); // end - 1 is the heading when the section is empty
  } else {
    if (mode !== 'prepend') for (let i = start + 1; i < end; i++) nodes[i].remove();
    nodes[start].after(...added);
  }
}

// OneNote check boxes are data-tag="to-do" / "to-do:completed".
// https://learn.microsoft.com/en-us/graph/onenote-note-tags
const CHECKBOX = '[data-tag^="to-do"]';
const REPLACEABLE_BLOCK = /^(p|li|h[1-6])$/i;

/**
 * Every check box under `root`, as { tagged, block, text }.
 *
 * `tagged` carries the data-tag — on the paragraph itself, or on a <span>
 * when the box lives in a list. `block` is the nearest ancestor OneNote will
 * accept as a replace target, since a <span> supports no update actions at
 * all; it is null when nothing on the way up qualifies, which is the caller's
 * cue to fall back to a wider edit rather than to write something that fails.
 */
export function checkboxesIn(root) {
  return [...root.querySelectorAll(CHECKBOX)].map((tagged) => {
    let block = tagged;
    while (block && !(block.getAttribute('id') && REPLACEABLE_BLOCK.test(block.tagName))) {
      block = block.parentElement;
    }
    return { tagged, block, text: collapse(tagged.textContent) };
  });
}

/** Every text node under `root`, skipping MathML — its text is structural. */
export function textNodesUnder(root) {
  const out = [];
  (function walk(node) {
    for (const child of node.childNodes) {
      if (child.nodeType === 3) out.push(child);
      else if (child.nodeType === 1 && child.tagName.toLowerCase() !== 'math') walk(child);
    }
  })(root);
  return out;
}

/**
 * Replace text inside already-selected nodes, leaving markup untouched. Returns
 * the count and the nodes that changed, so the caller can scope its write.
 */
export function replaceInTextNodes(nodes, find, replace, { all = true } = {}) {
  let occurrences = 0;
  const touched = [];
  for (const node of nodes) {
    const n = node.data.split(find).length - 1;
    if (!n) continue;
    if (!all) {
      node.data = node.data.replace(find, replace);
      return { occurrences: 1, touched: [node] };
    }
    node.data = node.data.split(find).join(replace);
    occurrences += n;
    touched.push(node);
  }
  return { occurrences, touched };
}
