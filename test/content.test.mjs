import { test } from 'node:test';
import assert from 'node:assert/strict';
import {
  htmlToMarkdown, htmlToText, markdownToOneNote, toOneNoteHtml, renderTable,
  tableToGrid, inlineMarkdown, editGrid, outlineOf, escapeHtml, wrapText, slugify,
  assertNoLostImages, toXhtml, HOUSE,
} from '../src/content.js';
import { parseHtml } from '../src/dom.js';
import { GUIDE, TOPICS, renderGuide } from '../src/guide.js';

// Real OneNote output HTML is minified — no whitespace between tags. This is
// the exact shape that the old textContent-based reader collapsed into mush.
const PAGE = '<html><body data-absolute-enabled="true"><div data-id="_default">'
  + '<h1 style="font-size:16pt;color:#1e4e79">Wochenplan</h1>'
  + '<p>Übersicht der Termine.</p>'
  + '<table style="border:0px">'
  + '<tr><td style="border:0px"><span style="font-weight:bold">Tag</span></td>'
  + '<td style="border:0px"><span style="font-weight:bold">Thema</span></td></tr>'
  + '<tr><td style="border:0px">Montag</td><td style="border:0px">Analysis</td></tr>'
  + '<tr><td style="border:0px">Dienstag</td><td style="border:0px">Lineare Algebra</td></tr>'
  + '</table>'
  + '<h2 style="font-size:14pt;color:#2e75b5">Aufgaben</h2>'
  + '<ul><li>Kapitel 3 lesen</li><li>Übungsblatt 4</li></ul>'
  + '</div></body></html>';

test('reading a page preserves block structure', async () => {
  const md = await htmlToMarkdown(PAGE);
  assert.match(md, /^# Wochenplan$/m);
  assert.match(md, /^## Aufgaben$/m);
  assert.match(md, /^\| Tag \| Thema \|$/m);
  assert.match(md, /^\| Montag \| Analysis \|$/m);
  assert.match(md, /^- Kapitel 3 lesen$/m);
  assert.match(md, /^- Übungsblatt 4$/m);
  // The regression this replaces: list items must not be welded together.
  assert.doesNotMatch(md, /lesenÜbungsblatt/);
});

test('bold spans become markdown emphasis, not lost styling', async () => {
  const md = await htmlToMarkdown('<body><p><span style="font-weight:bold">Hi</span> there</p></body>');
  assert.equal(md, '**Hi** there');
});

test('text mode still strips markers but keeps boundaries', async () => {
  const txt = await htmlToText(PAGE);
  assert.match(txt, /^Wochenplan$/m);
  assert.doesNotMatch(txt, /lesenÜbungsblatt/);
  assert.doesNotMatch(txt, /\*\*/);
});

test('markdown writes tables without the unsupported style border', async () => {
  const { html } = await markdownToOneNote('| A | B |\n| --- | --- |\n| 1 | 2 |');
  assert.match(html, /<table/);
  assert.doesNotMatch(html, /style="border/);
  assert.doesNotMatch(html, /<th[\s>]/); // OneNote has no <th>
  assert.match(html, new RegExp(HOUSE.headerCell.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')));
});

test('markdown table zebra-stripes body rows', async () => {
  const { html } = await markdownToOneNote('| A |\n| --- |\n| 1 |\n| 2 |');
  assert.match(html, /background-color:white/);
  assert.match(html, /background-color:#f2f2f2/);
});

test('headings get the house style and a stable data-id anchor', async () => {
  const { html } = await markdownToOneNote('# Kapitel Eins\n\ntext');
  assert.match(html, /<h1 style="font-size:16pt;color:#1e4e79/);
  assert.match(html, /data-id="h-kapitel-eins"/);
});

test('unsupported tags are mapped, never emitted raw', async () => {
  const { html } = await markdownToOneNote('```\ncode line\n```\n\n> quoted\n\n---\n');
  for (const tag of ['<pre', '<code', '<blockquote', '<hr']) {
    assert.ok(!html.includes(tag), `${tag} must not reach OneNote`);
  }
  assert.match(html, /font-family:Consolas/);
});

test('MathML survives a markdown round trip verbatim', async () => {
  const math = '<math xmlns="http://www.w3.org/1998/Math/MathML"><mfrac><mi>a</mi><mi>b</mi></mfrac></math>';
  const { html } = await markdownToOneNote(`Rate:\n\n${math}\n\nDone.`);
  assert.ok(html.includes(math), 'math block must pass through unchanged');
  const backToMd = await htmlToMarkdown(`<body>${html}</body>`);
  assert.ok(backToMd.includes('<math'), 'math must survive the read back');
});

test('read -> edit -> write round trip keeps the document shape', async () => {
  const md = await htmlToMarkdown(PAGE);
  const edited = md.replace('Analysis', 'Analysis II');
  const { html } = await markdownToOneNote(edited);
  const again = await htmlToMarkdown(`<body>${html}</body>`);
  assert.match(again, /^# Wochenplan$/m);
  assert.match(again, /^## Aufgaben$/m);
  assert.match(again, /Analysis II/);
  assert.match(again, /^- Kapitel 3 lesen$/m);
});

test('void elements are self-closed for XHTML', async () => {
  const { html } = await markdownToOneNote('line one\nline two');
  assert.match(html, /<br \/>/);
  assert.doesNotMatch(html, /<br>/);
});

test('escapeHtml escapes quotes so attributes cannot break out', () => {
  assert.equal(escapeHtml('The "big" idea'), 'The &quot;big&quot; idea');
  assert.equal(escapeHtml('a & b < c'), 'a &amp; b &lt; c');
});

test('explicit format replaces the old html-sniffing heuristic', async () => {
  // "if x<y then y>x" matched /<[a-z][\s\S]*>/i and was injected as raw markup.
  const { html } = await toOneNoteHtml('if x<y then y>x', 'text');
  assert.match(html, /&lt;y then y&gt;/);
  const asHtml = await toOneNoteHtml('<p>kept</p>', 'html');
  assert.equal(asHtml.html, '<p>kept</p>');
});

test('renderTable and tableToGrid are inverses', async () => {
  const html = renderTable({ headers: ['A', 'B'], rows: [['1', '2'], ['3', '4']] });
  const doc = await parseHtml(`<body>${html}</body>`);
  const { grid, hasHeader } = tableToGrid(doc.querySelector('table'));
  assert.equal(hasHeader, true);
  assert.deepEqual(grid.map((r) => r.map((c) => c.text)), [['A', 'B'], ['1', '2'], ['3', '4']]);
});

test('a cell keeps its formatting when another cell is edited', async () => {
  // The regression this guards: editTable used to read cells as textContent, so
  // changing one cell flattened the bold, links and equations in all the others.
  const doc = await parseHtml(PAGE);
  const { grid, hasHeader } = tableToGrid(doc.querySelector('table'));
  grid[1][1] = { html: inlineMarkdown('**Analysis II**'), text: 'Analysis II' };
  const html = renderTable({ headers: hasHeader ? grid[0] : null, rows: hasHeader ? grid.slice(1) : grid });
  assert.match(html, /font-weight:bold">Tag<\/span>/, 'header cell markup survived');
  assert.match(html, /<strong>Analysis II<\/strong>/, 'the edited cell rendered inline Markdown');
  assert.match(html, /Lineare Algebra/, 'untouched rows survived');
});

test('a table with no header row is not given one by an edit', async () => {
  const doc = await parseHtml('<body><table><tr><td>a</td></tr><tr><td>b</td></tr></table></body>');
  const { grid, hasHeader } = tableToGrid(doc.querySelector('table'));
  assert.equal(hasHeader, false);
  const html = renderTable({ headers: null, rows: grid });
  assert.ok(!html.includes(HOUSE.headerCell), 'a headerless table must stay headerless');
});

const gridOf = async (html) => {
  const doc = await parseHtml(`<body>${html}</body>`);
  return tableToGrid(doc.querySelector('table'));
};
const textOf = (grid) => grid.map((r) => r.map((c) => c.text));

test('a cell can be addressed by header name and by row text', async () => {
  const { grid, hasHeader } = await gridOf(renderTable({
    headers: ['Tag', 'Thema'], rows: [['Montag', 'Analysis'], ['Dienstag', 'Lineare Algebra']],
  }));
  const { grid: next } = editGrid(grid, hasHeader, {
    setCells: [{ row: 'Dienstag', col: 'Thema', value: '**Topologie**' }],
  });
  assert.match(next[2][1].html, /<strong>Topologie<\/strong>/);
  assert.equal(next[1][1].text, 'Analysis', 'the other rows are untouched');
});

test('an ambiguous or unknown reference is an error, not a wrong write', async () => {
  const { grid, hasHeader } = await gridOf(renderTable({ headers: ['A'], rows: [['x'], ['x']] }));
  const set = (row, col) => () => editGrid(grid, hasHeader, { setCells: [{ row, col, value: '1' }] });
  assert.throws(set('x', 'A'), /matches 2 rows/);
  assert.throws(set('zzz', 'A'), /No row matches/);
  assert.throws(set(0, 'nope'), /No column header matches/);
  assert.throws(set(9, 0), /Row 9 does not exist/);
});

test('a table with no header row can only be addressed by index', async () => {
  const { grid, hasHeader } = await gridOf('<table><tr><td>a</td></tr></table>');
  assert.equal(hasHeader, false);
  assert.throws(
    () => editGrid(grid, hasHeader, { setCells: [{ row: 0, col: 'A', value: '1' }] }),
    /no header row/
  );
});

test('columns can be added and deleted', async () => {
  const { grid, hasHeader } = await gridOf(renderTable({
    headers: ['Tag', 'Thema'], rows: [['Montag', 'Analysis']],
  }));
  const { grid: next, width } = editGrid(grid, hasHeader, {
    addColumns: [['Raum', 'B12']],
    deleteColumns: ['Thema'],
  });
  assert.equal(width, 2);
  assert.deepEqual(textOf(next), [['Tag', 'Raum'], ['Montag', 'B12']]);
});

test('every reference is resolved against the table as it was read', async () => {
  // Deleting a row shifts every row under it. Both edits must still land where
  // the caller meant them, whatever order they are applied in.
  const { grid, hasHeader } = await gridOf(renderTable({ headers: ['K'], rows: [['a'], ['b'], ['c']] }));
  const { grid: next } = editGrid(grid, hasHeader, {
    setCells: [{ row: 'c', col: 0, value: 'C' }],
    deleteRows: ['a'],
  });
  assert.deepEqual(textOf(next), [['K'], ['b'], ['C']]);
});

test('a new row wider than the table widens the table', async () => {
  const { grid, hasHeader } = await gridOf(renderTable({ headers: ['A'], rows: [['1']] }));
  const { grid: next, width } = editGrid(grid, hasHeader, { addRows: [['2', 'extra']] });
  assert.equal(width, 2);
  assert.deepEqual(textOf(next), [['A', ''], ['1', ''], ['2', 'extra']]);
});

test('a headerless table is striped from its first row', async () => {
  // The regression: striping counted from row 0 as if row 0 were a header, so
  // a headerless table came out grey-first here and white-first in renderTable.
  const { html } = await markdownToOneNote('<table><tr><td>a</td></tr><tr><td>b</td></tr></table>');
  assert.deepEqual(html.match(/background-color:[^"';]+/g), [
    'background-color:white', 'background-color:#f2f2f2',
  ]);
});

test('inline Markdown in a cell is rendered, and stray markup escaped', () => {
  assert.equal(inlineMarkdown('**b** [l](http://x)'), '<strong>b</strong> <a href="http://x">l</a>');
  assert.equal(inlineMarkdown('a & b < c'), 'a &amp; b &lt; c');
  assert.match(inlineMarkdown('`x`'), /font-family:Consolas/);
});

test('task lists become OneNote check boxes and read back as task lists', async () => {
  const { html } = await markdownToOneNote('- [ ] open\n- [x] done\n- [ ] third');
  // A data-tag on a <ul> overrides every child <li>, so check boxes must be
  // paragraphs. <input> is not supported input HTML and must never be sent.
  assert.doesNotMatch(html, /<input/);
  assert.doesNotMatch(html, /<ul/);
  assert.match(html, /<p data-tag="to-do">open<\/p>/);
  assert.match(html, /<p data-tag="to-do:completed">done<\/p>/);
  const back = await htmlToMarkdown(`<body>${html}</body>`);
  assert.equal(back, '- [ ] open\n- [x] done\n- [ ] third');
});

test('check boxes OneNote moved onto a span are still read as check boxes', async () => {
  // This is the shape the API returns for a list-level data-tag.
  const md = await htmlToMarkdown('<body><ul><li><span data-tag="to-do">a</span></li>'
    + '<li><span data-tag="to-do:completed">b</span></li></ul></body>');
  assert.equal(md, '- [ ] a\n- [x] b');
});

test('check boxes in a nested list degrade to bullets, and say so', async () => {
  const { html, warnings } = await markdownToOneNote('- [ ] parent\n  - [ ] child');
  assert.doesNotMatch(html, /<input/);
  assert.match(warnings.join(' '), /nested list/);
});

test('headings that slugify alike get distinct anchors', async () => {
  const { html } = await markdownToOneNote('# Notes\n\na\n\n# Notes\n\nb');
  assert.deepEqual(html.match(/data-id="[^"]+"/g), ['data-id="h-notes"', 'data-id="h-notes-2"']);
});

test('a table without its separator row warns instead of writing pipes', async () => {
  // The failure this catches: without the "| --- |" row marked emits a
  // paragraph, and the page silently gets a wall of pipe characters.
  const broken = await markdownToOneNote('| Tag | Thema |\n| Montag | Analysis |');
  assert.doesNotMatch(broken.html, /<table/);
  assert.match(broken.warnings.join(' '), /separator row/);

  const good = await markdownToOneNote('| Tag | Thema |\n| --- | --- |\n| Montag | Analysis |');
  assert.match(good.html, /<table/);
  assert.deepEqual(good.warnings, [], 'a well-formed table must not warn');
});

test('merged cells warn rather than silently shifting columns', async () => {
  // Raw HTML embedded in Markdown still goes through the house-style pass.
  // format:"html" deliberately does not; the server instructions cover that.
  const { warnings } = await markdownToOneNote('<table><tr><td colspan="2">wide</td></tr></table>');
  assert.match(warnings.join(' '), /rowspan\/colspan/);
});

test('the outline maps a page for a fraction of its size', async () => {
  const outline = await outlineOf(PAGE);
  assert.deepEqual(outline.headings, [
    { level: 1, text: 'Wochenplan' },
    { level: 2, text: 'Aufgaben' },
  ]);
  // afterHeading and rowLabels are what editTable is aimed with: with them, a
  // cell can be changed without ever reading the table.
  assert.deepEqual(outline.tables, [{
    index: 0, rows: 3, columns: 2, afterHeading: 'Wochenplan',
    headers: ['Tag', 'Thema'], rowLabels: ['Montag', 'Dienstag'],
  }]);
  assert.equal(outline.images, 0);
  assert.ok(outline.chars > 0);
  assert.ok(JSON.stringify(outline).length < PAGE.length, 'an outline must be cheaper than the page');
});

test('renderTable escapes cell content', () => {
  const html = renderTable({ headers: ['H'], rows: [['<script>x</script>']] });
  assert.ok(!html.includes('<script>'));
  assert.match(html, /&lt;script&gt;/);
});

test('a write that would silently drop images is refused', () => {
  const existing = '<p>a</p><img src="https://graph.microsoft.com/x/$value" />';
  assert.throws(() => assertNoLostImages('<p>a</p>', existing), /drop 1 of 1 image/);
  assert.doesNotThrow(() => assertNoLostImages('<p>a</p>', existing, { allowImageLoss: true }));
  assert.doesNotThrow(() => assertNoLostImages('<p>a</p>', '<p>no images here</p>'));
});

test('wrapText and slugify behave', () => {
  assert.equal(wrapText('a\n\nb'), '<p>a</p><p>b</p>');
  assert.equal(wrapText('a\nb'), '<p>a<br />b</p>');
  assert.equal(slugify('Übungsblatt 4 — Teil A'), 'ubungsblatt-4-teil-a');
  assert.equal(slugify('!!!'), 'section');
});

// ---------------------------------------------------------------------------
// Silent losses in the Markdown -> page direction
//
// GFM drops a cell that has no column to go in, and pipe-shaped lines without
// a separator row are not a table at all. Both fail without an error, so the
// only thing standing between them and a quietly wrong page is a warning.
// ---------------------------------------------------------------------------

test('a table row with more cells than its header warns instead of losing them', async () => {
  const { html, warnings } = await markdownToOneNote('| A | B |\n| --- | --- |\n| 1 | 2 | 3 |\n');
  assert.ok(!html.includes('>3<'), 'GFM really does drop the extra cell');
  assert.match(warnings.join(' '), /more cells than its 2-column header/);
});

test('a table with no separator row is reported, not silently written as prose', async () => {
  const { warnings } = await markdownToOneNote('| A | B |\n| 1 | 2 |\n');
  assert.match(warnings.join(' '), /separator row/);
});

test('a header and separator row of different widths is reported', async () => {
  const { warnings } = await markdownToOneNote('| A | B |\n| --- |\n| 1 | 2 |\n');
  assert.match(warnings.join(' '), /2 columns but its "\| --- \|" row has 1/);
});

test('a well-formed table warns about nothing', async () => {
  assert.deepEqual((await markdownToOneNote('| A | B |\n| --- | --- |\n| 1 | 2 |\n')).warnings, []);
  // An escaped pipe is one cell, not two, so it must not read as a wide row.
  assert.deepEqual((await markdownToOneNote('| A | B |\n|---|---|\n| P(A \\| B) | x |\n')).warnings, []);
});

test('pipes inside a code fence are a note about tables, not a broken table', async () => {
  // A note explaining Markdown tables would otherwise warn on every write.
  assert.deepEqual((await markdownToOneNote('```\n| a | b |\n| c | d |\n```\n')).warnings, []);
});

test('trailing list spacing goes after the list, not inside it', async () => {
  const { html } = await markdownToOneNote('- one\n  - a\n  - b\n- two\n');
  const doc = await parseHtml(`<body>${html}</body>`);
  const spaced = [...doc.querySelectorAll(`li[style*="${HOUSE.lastListItem}"]`)];
  assert.equal(spaced.length, 1, 'exactly one item carries the trailing space');
  assert.equal(spaced[0].textContent.trim(), 'two', 'and it is the last item of the OUTER list');
});

test('an insertion point survives deletions made in the same call', async () => {
  // atCol names a column in the table as the caller read it. The deletion
  // happens first and shifts everything after it, so without carrying the
  // index across, the new column lands one place too far right.
  const doc = await parseHtml(
    '<table>'
    + '<tr><td style="font-weight:bold">H1</td><td style="font-weight:bold">H2</td>'
    + '<td style="font-weight:bold">H3</td></tr>'
    + '<tr><td>a1</td><td>a2</td><td>a3</td></tr></table>'
  );
  const { grid, hasHeader } = tableToGrid(doc.querySelector('table'));
  const text = (g) => g.map((r) => r.map((c) => c.text));

  assert.deepEqual(
    text(editGrid(grid, hasHeader, { addColumns: [['H4', 'a4']], atCol: 'H3' }).grid),
    [['H1', 'H2', 'H4', 'H3'], ['a1', 'a2', 'a4', 'a3']]
  );
  assert.deepEqual(
    text(editGrid(grid, hasHeader, {
      deleteColumns: ['H1'], addColumns: [['H4', 'a4']], atCol: 'H3',
    }).grid),
    [['H2', 'H4', 'H3'], ['a2', 'a4', 'a3']],
    'the new column still lands before H3'
  );
  // Deleting the very column you insert at means "put the new one where that
  // one was".
  assert.deepEqual(
    text(editGrid(grid, hasHeader, {
      deleteColumns: ['H2'], addColumns: [['NEU', 'n']], atCol: 'H2',
    }).grid),
    [['H1', 'NEU', 'H3'], ['a1', 'n', 'a3']]
  );
});

test('a row insertion point survives deletions in the same call', async () => {
  const doc = await parseHtml(
    '<table><tr><td style="font-weight:bold">Name</td></tr>'
    + '<tr><td>alpha</td></tr><tr><td>beta</td></tr><tr><td>gamma</td></tr></table>'
  );
  const { grid, hasHeader } = tableToGrid(doc.querySelector('table'));
  const out = editGrid(grid, hasHeader, { deleteRows: ['alpha'], addRows: [['neu']], atRow: 'gamma' });
  assert.deepEqual(out.grid.map((r) => r[0].text), ['Name', 'beta', 'neu', 'gamma']);
});

test('void tags are closed for OneNote, and closing them again changes nothing', () => {
  // Every PATCH funnels through this, including commands serialised out of a
  // parsed page by jsdom — which writes HTML5 void tags.
  assert.equal(toXhtml('<p>a<br>b</p>'), '<p>a<br />b</p>');
  assert.equal(toXhtml('<img src="x" alt="y">'), '<img src="x" alt="y" />');
  // Idempotence is what makes it safe to apply to content that is already XHTML.
  const already = '<p>a<br />b</p><img src="x" />';
  assert.equal(toXhtml(already), already);
  // A page title is not markup and must survive untouched.
  assert.equal(toXhtml('Kapitel 4 < 5'), 'Kapitel 4 < 5');
});

// ---------------------------------------------------------------------------
// Headings h3-h6: what OneNote actually stores, not just what was sent
//
// Fixtures below are copied verbatim from a real page's GET response after
// createPage wrote HOUSE.h3/h4 through it. OneNote's importer will not let an
// h3-h6 element carry the requested colour/weight itself: it substitutes its
// OWN built-in per-level look on the tag and keeps the request only as a
// <span> wrapping the whole heading text. Two consequences follow, and both
// are invisible unless you write through the real API and read the result:
//   - that built-in look is ITALIC for h4 and h6, so a house style that never
//     mentions italic still comes out slanted unless it says font-style:normal
//   - the wrapping span makes an unedited heading indistinguishable from one
//     where the user bolded every character of it by hand
// ---------------------------------------------------------------------------

test('HOUSE cancels the italic OneNote silently forces onto h4 and h6', () => {
  // h3 and h5 were never affected — OneNote's own built-in look for those two
  // levels has no italic, so nothing to cancel.
  assert.doesNotMatch(HOUSE.h3, /font-style/);
  assert.match(HOUSE.h4, /font-style:\s*normal/);
});

test('reading back an OneNote-stored h3/h4 does not show a bogus **bold**', async () => {
  // OneNote wraps the whole text of h3/h4 in a bold span; h1/h2 are the two
  // levels it leaves untouched, so they never had this problem.
  const h3 = '<h3 style="font-size:12pt;color:#5b9bd5">'
    + '<span style="color:#2e75b5;font-weight:bold">Überschrift 3</span></h3>';
  const h4 = '<h4 style="font-size:12pt;color:#5b9bd5;font-style:italic">'
    + '<span style="font-size:11pt;color:#595959;font-weight:bold">Überschrift 4</span></h4>';
  assert.equal(await htmlToMarkdown(`<body>${h3}</body>`), '### Überschrift 3');
  assert.equal(await htmlToMarkdown(`<body>${h4}</body>`), '#### Überschrift 4');
});

test('bold on PART of a heading is a real user choice and survives', async () => {
  const h2 = '<h2>Kapitel <span style="font-weight:bold">1</span>: Overview</h2>';
  assert.equal(await htmlToMarkdown(`<body>${h2}</body>`), '## Kapitel **1**: Overview');
});

// ---------------------------------------------------------------------------
// Note tags other than to-do
//
// OneNote has ~20 built-in paragraph tags (Wichtig, Frage, Idee, Kritisch, a
// project colour, …), each shown as an icon in the app. None has a Markdown
// syntax, and unlike to-do there was previously no marker at all — a page
// someone had tagged "Kritisch" read back as indistinguishable plain text.
// ---------------------------------------------------------------------------

test('a non-to-do tag is surfaced, not silently dropped', async () => {
  const html = '<p data-tag="important">Wichtig (Sternchen)</p>'
    + '<p data-tag="critical">Kritisch (Ausrufezeichen)</p>'
    + '<p>Untagged paragraph.</p>';
  assert.equal(
    await htmlToMarkdown(`<body>${html}</body>`),
    '[tag: important] Wichtig (Sternchen)\n\n[tag: critical] Kritisch (Ausrufezeichen)\n\nUntagged paragraph.'
  );
});

test('outlineOf counts other tags per kind, and omits the key when there are none', async () => {
  const html = '<p data-tag="important">a</p><p data-tag="important">b</p><p data-tag="idea">c</p>'
    + '<p data-tag="to-do">d</p>';
  const withTags = await outlineOf(`<body>${html}</body>`);
  assert.deepEqual(withTags.tags, { important: 2, idea: 1 });
  assert.equal(withTags.checkboxes, 1, 'to-do is still counted separately, not folded into tags');

  const withoutTags = await outlineOf('<body><p>plain</p></body>');
  assert.ok(!('tags' in withoutTags), 'no tags key when the page has none');
});

// ---------------------------------------------------------------------------
// MathML placement: two OneNote import quirks, confirmed by writing minimal
// content through the live Graph API. Neither is caught by MathML validation
// (the MathML sent is perfectly valid) and neither is this server's own
// doing — OneNote just does not store either shape the way it was written.
// ---------------------------------------------------------------------------

const BLOCK_F = '<math display="block" xmlns="http://www.w3.org/1998/Math/MathML"><mi>f</mi></math>';
const BLOCK_G = '<math display="block" xmlns="http://www.w3.org/1998/Math/MathML"><mi>g</mi></math>';
const INLINE_F = '<math xmlns="http://www.w3.org/1998/Math/MathML"><mi>f</mi></math>';
const INLINE_G = '<math xmlns="http://www.w3.org/1998/Math/MathML"><mi>g</mi></math>';

test('two adjacent block equations warn: OneNote merges them into one', async () => {
  const { warnings } = await markdownToOneNote(`Text.\n\n${BLOCK_F}\n\n${BLOCK_G}\n`);
  assert.match(warnings.join(' '), /merges adjacent equations/);
});

test('block equations separated by real content do not warn', async () => {
  const md = `${BLOCK_F}\n\nEin Satz dazwischen.\n\n${BLOCK_G}\n`;
  assert.deepEqual((await markdownToOneNote(md)).warnings, []);
});

// What actually triggers the duplication is the equation being the ONLY thing
// in the list item — established live, case by case: alone in a flat item and
// alone in a nested item both duplicate, while an equation with any text
// beside it stays put at either depth. display="block" is irrelevant, because
// OneNote promotes a standalone equation to block on import regardless.

test('an equation alone in a list item warns, nested or not', async () => {
  const nested = await markdownToOneNote(`1. Aufgabe eins\n   1. ${BLOCK_F}\n`);
  assert.match(nested.warnings.join(' '), /only content of a list item/);
  // Written inline and alone in a FLAT item: this duplicated live too, and is
  // exactly the shape a display="block" filter used to miss.
  const flat = await markdownToOneNote(`1. ${INLINE_F}\n2. zwei\n`);
  assert.match(flat.warnings.join(' '), /only content of a list item/);
});

test('an equation with text beside it in the item does not warn', async () => {
  const flat = `1. Ergebnis ist ${INLINE_F} für alle x\n`;
  assert.deepEqual((await markdownToOneNote(flat)).warnings, []);
  const nested = `1. Oberpunkt\n   1. Text davor ${INLINE_F} Text danach\n`;
  assert.deepEqual((await markdownToOneNote(nested)).warnings, []);
});

test('adjacent equations warn whether written inline or as blocks', async () => {
  // Both merge live into one <math>; only the block form used to be caught.
  const asBlocks = await markdownToOneNote(`Text.\n\n${BLOCK_F}\n\n${BLOCK_G}\n`);
  assert.match(asBlocks.warnings.join(' '), /merges adjacent/);
  const asInline = await markdownToOneNote(`${INLINE_F}\n\n${INLINE_G}\n`);
  assert.match(asInline.warnings.join(' '), /merges adjacent/);
});

test('a lone block equation, outside any list, does not warn', async () => {
  assert.deepEqual((await markdownToOneNote(`Text davor.\n\n${BLOCK_F}\n`)).warnings, []);
});

// ---------------------------------------------------------------------------
// Raw HTML inline in otherwise-normal Markdown (not format:"html" for the
// whole page). This is what makes OneNote's other note tags reachable at all
// without a dedicated tool: CommonMark passes a recognised block-level tag
// through untouched, exactly like it already does for <math>.
// ---------------------------------------------------------------------------

test('a raw tag written directly in Markdown passes through untouched', async () => {
  const md = '# Kategorien\n\n<p data-tag="important">Wichtig (Sternchen)</p>\n\nNormaler Text danach.\n';
  const { html, warnings } = await markdownToOneNote(md);
  assert.ok(html.includes('<p data-tag="important">Wichtig (Sternchen)</p>'));
  assert.ok(html.includes('<p>Normaler Text danach.</p>'));
  assert.deepEqual(warnings, []);
});

test('raw HTML placed inside a list breaks the list, and does not warn about it', async () => {
  // A documented CommonMark caveat, not a bug: an HTML block interrupts list
  // parsing, so content after it inside the same list is no longer a list at
  // all. There is nothing to validate against here (the input is not
  // malformed Markdown), so this is a behaviour-pinning test, not a warning.
  const md = '- eins\n<p data-tag="critical">Kritisch</p>\n- zwei\n';
  const { html } = await markdownToOneNote(md);
  assert.ok(html.includes('<li style="margin-bottom:12pt">eins</li>'), 'first item became its own one-item list');
  assert.ok(html.includes('- zwei'), 'the line after the raw block is no longer parsed as a list item');
});

// ---------------------------------------------------------------------------
// getAuthoringGuide's content — a comprehensive companion to the terse,
// always-sent SERVER_INSTRUCTIONS, so it is fine for this to be long. What
// matters is that a topic can be read alone, and that "no topic" really means
// "everything," not an empty default.
// ---------------------------------------------------------------------------

test('every topic renders standalone, and the whole guide is their concatenation', () => {
  for (const topic of TOPICS) {
    const section = renderGuide(topic);
    assert.ok(section.length > 100, `${topic} should not be a stub`);
  }
  const whole = renderGuide();
  for (const topic of TOPICS) assert.ok(whole.includes(GUIDE[topic]), `${topic} missing from the full guide`);
});

test('an unknown topic is a clear error, not a silent empty string', () => {
  assert.throws(() => renderGuide('nonexistent'), /Unknown topic "nonexistent"/);
});

test('the guide carries the rules a model cannot derive from the docs', () => {
  // Each of these is a behaviour of the live service, not of Markdown, so the
  // guide is the only place a model learns it without hitting it first.
  assert.match(GUIDE.math, /MERGE into one/);
  assert.match(GUIDE.math, /DUPLICATED/);
  assert.match(GUIDE.math, /getMathExpressions is also the cheap way to CHECK/);
  assert.match(GUIDE.tables, /extra cells are dropped/);
  assert.match(GUIDE.tables, /Borders are impossible/);
  assert.match(GUIDE.tags, /data-tag="important"/);
  assert.match(GUIDE.tags, /never inside a list/);
  assert.match(GUIDE.tools, /no copy API/);
});

test('the guide tells a model how to structure a page, not just how to avoid bugs', () => {
  // Avoiding corruption is the floor. This topic is what separates a page that
  // merely renders from one worth reading.
  assert.match(GUIDE.documents, /Quellen|Sources/);
  assert.match(GUIDE.documents, /compares ACROSS/);
  assert.ok(TOPICS.includes('documents'));
});

test('the guide stays free of provenance and internals a model cannot act on', () => {
  // It is written for the model that has to act on it: how a fact was
  // established, and what the server does internally, change nothing about
  // what gets written and were costing a third of the guide's length.
  const whole = renderGuide();
  for (const noise of [/confirmed by writing/i, /live API/i, /this session/i, /wrapping <span>/i]) {
    assert.doesNotMatch(whole, noise, `guide should not explain ${noise}`);
  }
});

// ---------------------------------------------------------------------------
// Style presets
//
// A preset is chosen once at module load (code default, or ONENOTE_STYLE), not
// per call — a per-request parameter would cost schema bytes on every single
// request to express something that changes maybe once in a notebook's life.
// Re-importing with a cache-busting specifier is what lets one test process
// observe more than one preset.
// ---------------------------------------------------------------------------

const withStyle = async (value, tag) => {
  const before = process.env.ONENOTE_STYLE;
  if (value === undefined) delete process.env.ONENOTE_STYLE;
  else process.env.ONENOTE_STYLE = value;
  try {
    return await import(`../src/content.js?preset=${tag}`);
  } finally {
    if (before === undefined) delete process.env.ONENOTE_STYLE;
    else process.env.ONENOTE_STYLE = before;
  }
};

test('both presets exist and differ only where OneNote gives a choice', async () => {
  // Observed through HOUSE rather than a preset table: the presets are an
  // implementation detail, HOUSE is what every caller actually reads.
  const house = (await withStyle(undefined, 'house-cmp')).HOUSE;
  const native = (await withStyle('native', 'native-cmp')).HOUSE;
  // Headings are the real choice: OneNote supplies a full per-level look when
  // no style is sent, so "native" opts into it by saying nothing.
  for (const level of ['h1', 'h2', 'h3', 'h4']) {
    assert.ok(house[level].length > 0, `house.${level} should be explicit`);
    assert.equal(native[level], '', `native.${level} should defer to OneNote`);
  }
  // Table shading is NOT a choice — a border is stripped by OneNote whether it
  // is sent on the table or the cell, so shading is the only way to delimit
  // rows and both presets must keep it.
  for (const key of ['headerCell', 'cellEven', 'cellOdd', 'code', 'codeInline', 'quote']) {
    assert.equal(native[key], house[key], `${key} should be shared, not a style choice`);
  }
});

test('the native preset writes headings with no style attribute at all', async () => {
  const { markdownToOneNote: nativeMd } = await withStyle('native', 'native');
  const { html } = await nativeMd('# Eins\n## Zwei\n### Drei\n#### Vier\n');
  const doc = await parseHtml(`<body>${html}</body>`);
  for (const h of doc.querySelectorAll('h1,h2,h3,h4')) {
    assert.equal(h.getAttribute('style'), null, `${h.tagName} must carry no style, not style=""`);
    assert.ok(h.getAttribute('data-id'), 'the anchor is still stamped');
  }
  // Everything OneNote does not supply for free is still applied.
  const { html: table } = await nativeMd('| A |\n|---|\n| 1 |\n');
  assert.match(table, /background-color:#d9e2f3/, 'header shading survives in native');
});

test('the default preset still states every heading explicitly', async () => {
  const { markdownToOneNote: houseMd } = await withStyle(undefined, 'default');
  const { html } = await houseMd('# Eins\n### Drei\n');
  const doc = await parseHtml(`<body>${html}</body>`);
  assert.match(doc.querySelector('h1').getAttribute('style'), /font-size:16pt/);
  assert.match(doc.querySelector('h3').getAttribute('style'), /font-weight:bold/);
});

test('an unknown preset name fails at startup instead of silently falling back', async () => {
  await assert.rejects(
    () => withStyle('sparkly', 'bogus'),
    /Unknown ONENOTE_STYLE "sparkly"/
  );
});

test('the guide callout recipe round-trips byte-for-byte', async () => {
  // Verified against the live API: a TWO-row one-column table gives a shaded
  // bold title over a white body. A ONE-row table is all header, so every word
  // is bolded and then stripped again on read — which is why the guide says
  // not to use one. Both halves are pinned here so the recipe cannot rot.
  const callout = '| ⚠ Warnung |\n| --- |\n| Text der Warnung. |';
  const { html } = await markdownToOneNote(callout);
  assert.match(html, new RegExp(`background-color:${HOUSE.headerCell.match(/#[0-9a-f]{6}/i)[0]}`));
  assert.equal(await htmlToMarkdown(`<body>${html}</body>`), callout);

  // The one-row form: a real table, but with a header and no body, so the
  // whole box is bolded and shaded and the bold is then stripped on read.
  const lone = '| \u26a0 Warnung |\n| --- |';
  const oneRow = await markdownToOneNote(lone);
  assert.match(oneRow.html, /font-weight:bold/, 'everything in a lone row is header-bold');
  assert.equal(
    await htmlToMarkdown(`<body>${oneRow.html}</body>`),
    lone,
    'the header bold is an artifact and comes back off'
  );
});
