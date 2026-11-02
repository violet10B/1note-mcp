// ---------------------------------------------------------------------------
// The long-form authoring guide
//
// SERVER_INSTRUCTIONS (onenote-mcp.mjs) is the terse version, paid for on
// every request. This is the comprehensive one, behind a tool call, so its
// content costs nothing until a model actually reads it.
//
// Written for the model that has to ACT on it. Every line should change what
// gets written; nothing here explains how the server works internally, or how
// a fact was established, because neither changes the output. Facts that came
// from writing real content through the live API are stated as rules, not as
// findings — the provenance belongs in the commit history and the tests.
// ---------------------------------------------------------------------------

export const GUIDE = {
  overview: `OVERVIEW

Markdown in, Markdown out. Write Markdown and the server converts it and applies the house style — do not hand-write styled HTML.

Before editing a page you did not just write, call getPage(format:"outline"): a cheap map of the headings, table shapes, headers and row labels, usually enough to aim editTable / editPageSection / setCheckbox without reading the body at all.

Writing tools can return "warnings" alongside success. A warning means the write went through but something did not survive as written. Read it — it names the fix.`,

  documents: `WRITING A GOOD PAGE

STRUCTURE
  - Open with one "# " heading naming the topic, then "## " per section.
  - Do not open with a generic "Overview" heading. Lead with the substance.
  - Number section headings ("## 1. …") when the order carries meaning — a procedure, a derivation, a syllabus. Leave them unnumbered when it does not.
  - Close with a "## Quellen" / "## Sources" section, one source per line, each tagged web-site-to-visit (see "tags").

TABLES VS LISTS — the most common mistake
  A table earns its place when the reader compares ACROSS its columns. Two columns of "thing → what it is" is a list wearing a table costume.
    weak    | Item | Purpose |                          — a description list
    strong  | Salt from | Hydrolysis? | pH | Example |   — four dimensions, each column comparable top to bottom
  If a second column only restates what the first implies, write a bulleted list instead. If you can name a third and fourth dimension, the table was the right call.
  Bulleted list for unordered points, numbered list for steps, table only for repeated structure.

MAKE IT ACTIONABLE
  - "- [ ] " check boxes for what the reader will actually tick off — a shopping list, prerequisites, a revision plan. Not for reference material.
  - A callout table (see "styling") for a warning or a caveat that must not be missed. Not for ordinary prose.
  - Note tags to mark what matters: important, question, idea, and web-site-to-visit on each source.

BE SPECIFIC
  A page of notes is worth more than an encyclopedia entry. Prefer concrete numbers, names, thresholds and worked examples over advice that would be true of any page on the subject. If a sentence would survive unchanged in a different document, cut it or sharpen it.`,

  tables: `TABLES

GFM: a header row, then a "| --- | --- |" row with EXACTLY as many cells as the header, then the data rows. Two failures are silent, and both are reported in warnings:
  - no separator row, or one of a different width → not parsed as a table at all, just a paragraph of pipe characters
  - a data row with more cells than the header → the extra cells are dropped

Escape a literal "|" inside a cell as "\\|". There is no <th>: row 0 is the header and is shaded and bolded for you. Borders are impossible — OneNote strips them from tables and cells alike — so the shading is what delimits rows.

Never rewrite a table by hand to change it. OneNote cannot patch a <tr> or a <td>, so every cell edit replaces the whole table; editTable does exactly that and addresses rows and columns by 0-based index OR by header/row text (exact match wins over partial; ambiguous text is an error listing the candidates). getPage(format:"outline") already lists each table's headers and row labels, which is what lets you aim a cell edit without reading the table first.`,

  math: `MATH

Write MathML inline: <math xmlns="http://www.w3.org/1998/Math/MathML">…</math>. Do not bother choosing display="block" — OneNote promotes a standalone equation to block either way.

Two placements corrupt an equation. Both are caught and reported in warnings before the write:
  - Two equations with nothing but whitespace between them MERGE into one, concatenating their content. Put a sentence between them; a blank line is not enough.
  - An equation that is the ONLY content of a list item is DUPLICATED as a stray copy outside the list. Put text beside it in the item, e.g. "a) <equation>". With any text alongside, it stays put.

To edit an existing formula, take the exact fragment from getMathExpressions: OneNote normalises what it stores, so the string you originally sent will not match. getMathExpressions is also the cheap way to CHECK a formula landed intact — a successful write only means OneNote accepted the command, not that it stored what you sent. getPage(format:"text") flattens equations; read markdown or html when they matter.`,

  tags: `NOTE TAGS

"- [ ]" and "- [x]" become real, tickable check boxes. setCheckbox ticks or unticks them by their text — several in one call, without reading the page.

OneNote's other tags have no Markdown syntax: important, question, idea, critical, contact, address, phone-number, web-site-to-visit, remember-for-later, project-a, project-b, movie-to-see, book-to-read, music-to-listen-to and more. Write the raw tag as its own block inside otherwise-normal Markdown:

  <p data-tag="important">Der wichtige Satz.</p>

Keep it between blocks, never inside a list — an HTML block ends list parsing, so the rest of the list stops being a list. On read they come back as "[tag: important] the text" (a marker only; writing that back attaches nothing), and getPage(format:"outline") counts them per kind.`,

  styling: `STYLING

Write Markdown and the styling is handled: "#" through "######" are styled automatically and each gets a stable data-id anchor. Do not hand-write heading styles — OneNote overrides them for h3-h6 anyway.

If you do write raw HTML: inline style= only, and only background-color, color, font-family, font-size, font-style (normal|italic), font-weight (normal|bold), text-align, text-decoration. No CSS classes, no <style> blocks. OneNote silently drops <th>, <pre>, <code>, <blockquote>, <hr>, rowspan and colspan, and strips table borders.

CALLOUT BOX — a one-column table with a title row and a body row:

  | ⚠ Warnung |
  | --- |
  | Text der Warnung. |

The title row is shaded and bolded automatically; the body row stays plain, so **bold** inside it reads as real emphasis. Use 📓 Anmerkungen, 🛈 Information, 💡 Tipp the same way. Do NOT put the whole callout in a single row with <br /> between the lines: a one-row table is all header, so every word comes out bold.

Spacing is automatic — headings and lists carry their own, so do not add empty paragraphs to fake it. "---" is not a horizontal rule (OneNote drops <hr>); separate sections with a heading.`,

  images: `IMAGES

addImageToPage uploads from a local filePath or a url and does the binary work itself. Never paste base64, and never write your own <img src="https://…"> except inside createPage's initial content. png/jpg/gif/bmp/tiff, up to about 3.5 MB.

Images already on a page are authenticated resources that cannot be re-submitted, so any whole-body rewrite that would drop them — replacePage, or editPageSection over a section containing one — is refused unless you pass allowImageLoss:true. Prefer a narrower edit that goes around the image.`,

  tools: `CHOOSING A TOOL

Narrowest first; each of these avoids re-sending the page:
  editTable       one table: cells, rows, columns, by index or by header/row text
  setCheckbox     tick or untick check boxes by their text
  editPageSection replace, append to, prepend to or delete one heading's content
  findAndReplace  swap an exact HTML fragment, or a run of plain text
  insertNearText  insert before/after the paragraph containing a fragment
  replaceElement  replace one block by its generated id
  addToPage       append or prepend to the page
  replacePage     rewrite the whole body — only when nothing narrower fits

Reading: getPage(format:"outline") for a map, getPageSection for one heading, getMathExpressions for the formulas alone, searchPages(inContent:true) to find a page by what is in it rather than its title (scope it to a section or notebook). batchUpdatePages applies one update to many pages at once.

Copy and move: personal accounts have no copy API, so copyPageToSection re-creates the page instead — new id, new creation date, and images cannot come along (allowImageLoss:true to accept that). A whole section has no stand-in: create it, then copy its pages.`,
};

export const TOPICS = Object.keys(GUIDE);

/** The requested section, or every section joined, when no topic is given. */
export function renderGuide(topic) {
  if (topic == null) return TOPICS.map((t) => GUIDE[t]).join('\n\n');
  if (!GUIDE[topic]) throw new Error(`Unknown topic "${topic}". Choose from: ${TOPICS.join(', ')}.`);
  return GUIDE[topic];
}
