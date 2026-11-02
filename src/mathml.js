// ---------------------------------------------------------------------------
// MathML validation — catches the silent-corruption bugs OneNote produces from
// malformed equations (non-well-formed XML, wrong child counts).
// ---------------------------------------------------------------------------
import { parseXml } from './dom.js';

const MATHML_ELEMENTS = new Set([
  'math', 'mrow', 'mi', 'mo', 'mn', 'ms', 'mtext', 'mspace', 'mfrac', 'msqrt',
  'mroot', 'msub', 'msup', 'msubsup', 'munder', 'mover', 'munderover',
  'mmultiscripts', 'mtable', 'mtr', 'mtd', 'mfenced', 'mpadded', 'mphantom',
  'menclose', 'mstyle', 'merror', 'mglyph', 'mlabeledtr', 'maligngroup',
  'malignmark', 'none', 'mprescripts', 'semantics', 'annotation', 'annotation-xml',
]);

// Presentation elements with a FIXED number of child elements.
const MATHML_ARITY = {
  mfrac: 2, msub: 2, msup: 2, msubsup: 3, munder: 2, mover: 2, munderover: 3, mroot: 2,
};

/** Every <math>…</math> block in a string. */
export const findMathBlocks = (s) => (s && s.match(/<math[\s\S]*?<\/math>/gi)) || [];

export async function validateMathMLString(mathml) {
  const errors = [], warnings = [];
  let xdoc;
  try {
    xdoc = await parseXml(mathml);
  } catch (e) {
    return { valid: false, errors: ['XML parse failed: ' + e.message], warnings };
  }
  const perr = xdoc.querySelector('parsererror');
  if (perr) {
    errors.push('Not well-formed XML: ' + perr.textContent.replace(/\s+/g, ' ').trim().slice(0, 160));
    return { valid: false, errors, warnings };
  }
  const root = xdoc.documentElement;
  if (!root) return { valid: false, errors: ['Empty / unparseable.'], warnings };
  if (root.nodeName.toLowerCase() !== 'math') {
    errors.push(`Root element is <${root.nodeName}>, expected <math>.`);
  } else if (!root.getAttribute('xmlns')) {
    warnings.push('Missing xmlns="http://www.w3.org/1998/Math/MathML" on <math>.');
  }
  (function walk(el) {
    const name = el.nodeName.toLowerCase();
    if (!MATHML_ELEMENTS.has(name)) warnings.push(`Unknown element <${name}> (may not render).`);
    const kids = [...el.children];
    if (MATHML_ARITY[name] !== undefined && kids.length !== MATHML_ARITY[name]) {
      errors.push(`<${name}> needs exactly ${MATHML_ARITY[name]} child elements but has ${kids.length}.`);
    }
    kids.forEach(walk);
  })(root);
  return { valid: errors.length === 0, errors, warnings };
}

/** Validate every <math> block in a content string. hardError => refuse to write. */
export async function checkMathInContent(content) {
  if (!content || !/<math[\s>]/i.test(content)) return { hardError: null, warnings: [] };
  const hard = [], warnings = [];
  for (const block of findMathBlocks(content)) {
    const r = await validateMathMLString(block);
    if (!r.valid) hard.push(...r.errors);
    warnings.push(...r.warnings);
  }
  return { hardError: hard.length ? hard.join(' | ') : null, warnings };
}
