// The block surgery behind editPageSection and findAndReplace's text mode.
// OneNote has no delete action and cannot patch a <tr>, <td> or <span>, so
// every edit finer than one block element is spliced here — which makes this
// the code most worth pinning down without a Graph round trip.
import { test } from 'node:test';
import assert from 'node:assert/strict';
import {
  parseHtml, contentContainer, sectionSpan, spliceSection, textNodesUnder, replaceInTextNodes,
  checkboxesIn, pickOne,
} from '../src/dom.js';

const PAGE = '<html><body data-absolute-enabled="true"><div data-id="_default">'
  + '<h1>Titel</h1><p>intro</p>'
  + '<h2>Alpha</h2><p>a1</p><h3>Alpha detail</h3><p>a2</p>'
  + '<h2>Beta</h2><p>b1</p>'
  + '</div></body></html>';

const load = async (html = PAGE) => {
  const doc = await parseHtml(html);
  return { doc, container: contentContainer(doc.body) };
};
const tags = (container) => [...container.children].map((el) => el.tagName.toLowerCase() + ':' + el.textContent);

test('the content container is the page wrapper div, or the body', async () => {
  const { doc, container } = await load();
  assert.equal(container.getAttribute('data-id'), '_default');
  const bare = await parseHtml('<body><p>a</p><p>b</p></body>');
  assert.equal(contentContainer(bare.body), bare.body);
  assert.equal(doc.body.tagName.toLowerCase(), 'body');
});

test('a section runs to the next heading of equal or higher level', async () => {
  const { container } = await load();
  const { nodes, start, end } = sectionSpan(container, 'Alpha');
  // The h3 belongs to the Alpha section; the following h2 ends it.
  assert.deepEqual(nodes.slice(start, end).map((n) => n.tagName.toLowerCase()), ['h2', 'p', 'h3', 'p']);
});

test('the last section runs to the end of the page', async () => {
  const { container } = await load();
  const { nodes, start, end } = sectionSpan(container, 'Beta');
  assert.equal(end, nodes.length);
  assert.deepEqual(nodes.slice(start, end).map((n) => n.tagName.toLowerCase()), ['h2', 'p']);
});

test('deleting a section takes the heading with it', async () => {
  // OneNote has no delete action, so this is the only way to remove a run of
  // blocks: drop them here and write the body back.
  const { doc, container } = await load();
  spliceSection(doc, 'Alpha', '', 'delete');
  assert.deepEqual(tags(container), ['h1:Titel', 'p:intro', 'h2:Beta', 'p:b1']);
});

test('a missing heading fails loudly rather than editing the wrong place', async () => {
  const { container } = await load();
  assert.throws(() => sectionSpan(container, 'Omega'), /No heading matches "Omega"/);
  // The error names the headings that do exist, so the next call can be right.
  assert.throws(() => sectionSpan(container, 'Omega'), /"Alpha detail"/);
});

test('an exact heading beats one that merely contains the fragment', async () => {
  // "Alpha" is a prefix of "Alpha detail". Editing the wrong section here is
  // silent and destructive, so exactness has to win.
  const { container } = await load();
  const { nodes, start } = sectionSpan(container, 'Alpha');
  assert.equal(nodes[start].textContent, 'Alpha');
});

test('an ambiguous heading is an error, not a coin flip', async () => {
  const { container } = await load(
    '<body><div><h2>Notes</h2><p>a</p><h2>Notes</h2><p>b</p></div></body>'
  );
  assert.throws(() => sectionSpan(container, 'Notes'), /matches 2 headings/);
  // A fragment that is ambiguous only as a substring is caught the same way.
  const { container: c2 } = await load(
    '<body><div><h2>Woche 1</h2><p>a</p><h2>Woche 2</h2><p>b</p></div></body>'
  );
  assert.throws(() => sectionSpan(c2, 'Woche'), /matches 2 headings/);
  assert.equal(sectionSpan(c2, 'Woche 2').nodes[sectionSpan(c2, 'Woche 2').start].textContent, 'Woche 2');
});

test('replacing a section keeps the heading and its neighbours', async () => {
  const { doc, container } = await load();
  spliceSection(doc, 'Alpha', '<p>new</p>', 'replace');
  assert.deepEqual(tags(container), [
    'h1:Titel', 'p:intro', 'h2:Alpha', 'p:new', 'h2:Beta', 'p:b1',
  ]);
});

test('appending puts content at the end of the section, not the page', async () => {
  const { doc, container } = await load();
  spliceSection(doc, 'Alpha', '<p>new</p>', 'append');
  assert.deepEqual(tags(container), [
    'h1:Titel', 'p:intro', 'h2:Alpha', 'p:a1', 'h3:Alpha detail', 'p:a2', 'p:new', 'h2:Beta', 'p:b1',
  ]);
});

test('prepending puts content directly under the heading', async () => {
  const { doc, container } = await load();
  spliceSection(doc, 'Alpha', '<p>new</p>', 'prepend');
  assert.deepEqual(tags(container), [
    'h1:Titel', 'p:intro', 'h2:Alpha', 'p:new', 'p:a1', 'h3:Alpha detail', 'p:a2', 'h2:Beta', 'p:b1',
  ]);
});

test('appending to an empty section lands between the two headings', async () => {
  const { doc, container } = await load(
    '<html><body><div><h2>Gamma</h2><h2>Delta</h2><p>d</p></div></body></html>'
  );
  spliceSection(doc, 'Gamma', '<p>new</p>', 'append');
  assert.deepEqual(tags(container), ['h2:Gamma', 'p:new', 'h2:Delta', 'p:d']);
});

test('splicing a whole subtree in keeps its structure', async () => {
  const { doc, container } = await load();
  spliceSection(doc, 'Beta', '<h3>Sub</h3><ul><li>x</li></ul>', 'replace');
  assert.deepEqual(tags(container).slice(-3), ['h2:Beta', 'h3:Sub', 'ul:x']);
});

test('text nodes inside equations are left alone', async () => {
  const doc = await parseHtml(
    '<body><p>a <b>b</b> c</p><math xmlns="x"><mi>b</mi></math></body>'
  );
  assert.deepEqual(textNodesUnder(doc.body).map((n) => n.data), ['a ', 'b', ' c']);
});

test('replacing text keeps the markup around it', async () => {
  const doc = await parseHtml('<body><p>Montag: <b>Analysis</b> am Montag</p></body>');
  const nodes = textNodesUnder(doc.body).filter((n) => n.data.includes('Montag'));
  const { occurrences, touched } = replaceInTextNodes(nodes, 'Montag', 'Dienstag');
  assert.equal(occurrences, 2);
  assert.equal(touched.length, 2);
  assert.equal(doc.body.innerHTML, '<p>Dienstag: <b>Analysis</b> am Dienstag</p>');
});

test('all:false replaces only the first occurrence', async () => {
  const doc = await parseHtml('<body><p>x x</p><p>x</p></body>');
  const nodes = textNodesUnder(doc.body);
  const { occurrences } = replaceInTextNodes(nodes, 'x', 'y', { all: false });
  assert.equal(occurrences, 1);
  assert.equal(doc.body.innerHTML, '<p>y x</p><p>x</p>');
});

test('a replacement is inserted as text, so markup characters cannot escape', async () => {
  const doc = await parseHtml('<body><p>value</p></body>');
  replaceInTextNodes(textNodesUnder(doc.body), 'value', 'a < b & c');
  assert.equal(doc.body.innerHTML, '<p>a &lt; b &amp; c</p>');
});

// ---------------------------------------------------------------------------
// Check boxes
//
// Ticking one used to mean reading the whole page as HTML just to learn an
// element id. These are the two shapes OneNote returns them in.
// ---------------------------------------------------------------------------

const TASKS = '<body><div>'
  + '<p data-tag="to-do" id="p:{g}{1}">Übungsblatt 5</p>'
  + '<p data-tag="to-do:completed" id="p:{g}{2}">Skript lesen</p>'
  + '<ul><li id="li:{g}{3}"><span data-tag="to-do" id="span:{g}{4}">Formeln lernen</span></li></ul>'
  + '<p id="p:{g}{5}">not a task</p>'
  + '</div></body>';

test('every check box is found, whatever element carries its tag', async () => {
  const { doc } = await load(TASKS);
  const boxes = checkboxesIn(doc.body);
  assert.deepEqual(boxes.map((b) => b.text), ['Übungsblatt 5', 'Skript lesen', 'Formeln lernen']);
});

test('the write targets a block OneNote can replace, never the <span>', async () => {
  // <span> supports no update actions at all, so a list check box has to be
  // written back through the <li> that holds it.
  const { doc } = await load(TASKS);
  const boxes = checkboxesIn(doc.body);
  assert.deepEqual(boxes.map((b) => b.block.tagName.toLowerCase()), ['p', 'p', 'li']);
  assert.equal(boxes[2].block.getAttribute('id'), 'li:{g}{3}');
});

test('a check box with nothing replaceable above it reports rather than writes', async () => {
  // No generated ids: nothing here can be a PATCH target, and the caller needs
  // to know that before it builds a command OneNote will reject.
  const { doc } = await load('<body><div><p data-tag="to-do">Task</p></div></body>');
  assert.equal(checkboxesIn(doc.body)[0].block, null);
});

test('a check box is named the same way a heading is', async () => {
  const { doc } = await load(TASKS);
  const boxes = checkboxesIn(doc.body);
  const box = (t) => pickOne(boxes, t, 'check box', 'check boxes');
  assert.equal(box('formeln').text, 'Formeln lernen');
  assert.throws(() => box('le'), /matches 2 check boxes/);
  assert.throws(() => box('Nichts'), /No check box matches/);
});
