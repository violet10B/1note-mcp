# OneNote MCP Server

An MCP server for Microsoft OneNote, using the Graph API. Pages are read and
written as Markdown. It uses the public Graph Explorer client id, so no Azure
app registration is needed.

Derived from [danosb/onenote-mcp](https://github.com/danosb/onenote-mcp) and
[azure-onenote-mcp-server](https://github.com/ZubeidHendricks/azure-onenote-mcp-server).
See [Credits](#credits).

## Requirements

Node.js 20 or newer. `marked@18` requires it, and the repo has an `.nvmrc`.

A Microsoft account with OneNote. Personal accounts work. See
[Account limitations](#account-limitations).

## Install

```bash
git clone https://github.com/violet10B/1note-mcp.git
cd 1note-mcp
npm install
npm run auth
```

## Client registration

Claude Desktop config lives at `%APPDATA%\Claude\claude_desktop_config.json` on
Windows, and `~/Library/Application Support/Claude/claude_desktop_config.json`
on macOS. Merge this in and restart the app:

```json
{
  "mcpServers": {
    "onenote": {
      "command": "node",
      "args": ["/absolute/path/to/1note-mcp/onenote-mcp.mjs"]
    }
  }
}
```

On Windows `command` has to be the full path to the executable, for example
`C:\\Program Files\\nodejs\\node.exe`. On macOS, `command -v node` prints it.

Cursor uses the same `mcpServers` object.

## Authentication

`npm run auth` starts the device-code flow. Open `microsoft.com/devicelogin`,
enter the code, sign in. The token cache goes to `.msal-cache.json` with mode
`0600`. It refreshes for about 90 days.

The `authenticate` tool runs the same flow from a client.

## Configuration

Copy `.env.example` to `.env` if you need any of these. `.env` is read from the
repo root, not the working directory.

| Variable | Default | Purpose |
|---|---|---|
| `ONENOTE_STYLE` | `house` | Heading preset: `house` or `native`. An unknown value fails at startup. |
| `GRAPH_ACCESS_TOKEN` | none | Raw token, used only without an MSAL session. Expires in about an hour and does not refresh. |
| `ONENOTE_CLIENT_ID` | Graph Explorer client id | Use a private app registration, or survive Microsoft retiring the public one. |


## Content model

`getPage` and all writing tools default to Markdown. The converter emits only
the [supported HTML/CSS subset](https://learn.microsoft.com/en-us/graph/onenote-input-output-html):

| Markdown | Emitted as |
|---|---|
| `# Heading` | `<h1>` with inline style and a `data-id` anchor |
| GFM table | `<table>` with shaded header and zebra cells |
| ` ```code``` ` | `<p><span style="font-family:Consolas">` |
| `> quote` | `<p><span style="color:#595959;font-style:italic">` |
| `- [ ] task` | `<p data-tag="to-do">` |
| `---` | an empty paragraph |

OneNote drops `<th>`, `<pre>`, `<code>`, `<blockquote>`, `<hr>`, `<style>`, CSS
classes and table borders on input. Raw HTML passes through, so `<math>` and
`<sup>`/`<sub>` work as written.

Two Markdown errors fail without an error message. Writing tools return
`warnings` for both: a table with no `| --- |` separator row, and any cell in a
row wider than its header. Both checks run against the Markdown source.

`format: "html"` skips the converter. `format: "text"` escapes the content and
wraps it in paragraphs.

### Style presets

Set in code or with `ONENOTE_STYLE`. Not per call.

| Preset | Headings |
|---|---|
| `house` (default) | every level set explicitly |
| `native` | no style attribute, so OneNote applies its own look |

`native` sends fewer bytes. It skips the wrapping `<span>` OneNote adds when an
h3 to h6 element carries its own style.

Table borders come back as `border:0px`, so shading is the only way to separate
rows. Both presets use it.

### MathML

Every write validates MathML first: well-formedness plus child counts for
`mfrac`, `msub`, `msup` and similar. Invalid markup is rejected.

Two placements corrupt an equation after OneNote stores it. This server does not
cause either, and validation does not catch either:

- Two equations with nothing between them merge into one. A blank line does not
  stop this. Put text between them.
- An equation that is the only content of a list item gets duplicated outside
  the list. Reading the page and writing it back adds another copy. Write
  `a) <equation>` instead.

OneNote promotes a standalone equation to block on import whatever `display`
says, so both apply to inline math as well. Both are checked before the write and
reported in `warnings`.

`getPage(format:"text")` flattens equations to characters.

## Tools

### Guide

| Tool | Description |
|---|---|
| `getAuthoringGuide` | Long-form authoring reference. Optional `topic` (`overview`, `documents`, `tables`, `math`, `tags`, `styling`, `images`, `tools`) reads one section. |

Holds what the `instructions` field cannot, since `instructions` is re-sent on
every call.

### Notebooks and sections

| Tool | Description |
|---|---|
| `listNotebooks` | List notebooks, optionally filtered by name |
| `getNotebookTree` | Notebook / section-group / section hierarchy in one call |
| `listSections` | Sections with notebook and group attribution |
| `copySectionToNotebook` | Async. Polls up to 60 s. |

Also: `createNotebook`, `createSection`, `createSectionGroup`.

### Reading

| Tool | Description |
|---|---|
| `listPages` | List pages in a section; `previewChars` adds content snippets |
| `getPage` | Read a page as `markdown` (default), `outline`, `text` or `html` |
| `getPageSection` | Read the content under one heading |
| `getMathExpressions` | Read the `<math>` blocks with context |
| `searchPages` | Match titles, or bodies with `inContent` |

### Writing

Use the narrowest tool that fits.

| Tool | Description |
|---|---|
| `editTable` | Set cells, add or delete rows and columns |
| `editPageSection` | Replace, append, prepend or delete one heading's content |
| `setCheckbox` | Tick or untick check boxes by text |
| `findAndReplace` | Swap an exact HTML fragment or a run of visible text |
| `insertNearText` | Insert before or after the paragraph matching a fragment |
| `replaceElement` | Replace one block element by id |
| `addToPage` | Append or prepend content |
| `replacePage` | Replace the whole body. Widest edit. |
| `createPage` | Optional `lang` for spell-check. |
| `addImageToPage` | From a local path or URL |
| `batchUpdatePages` | One change applied to many pages |

Also: `setPageTitle`.

### Managing and auth

| Tool | Description |
|---|---|
| `copyPageToSection` | Copy to another section, or move |
| `authenticate` | Start the device-code flow |
| `saveAccessToken` | Raw token. About an hour, no refresh. |

Also: `deletePage`. Pages are the only thing Graph lets this server delete.

## Addressing by text

Headings, table rows, columns and check boxes take a text fragment instead of an
index (`pickOne` in [src/dom.js](src/dom.js)).

- An exact match beats a substring match. `"Tasks"` selects the heading `Tasks`,
  not `Tasks archive`.
- An ambiguous fragment returns an error listing the candidates. Nothing is
  written.

Rows and columns also take a 0-based index.

## Table editing

`editTable` finds a table by `nearText` or `tableIndex`, then rewrites that one
`<table>`:

```jsonc
{ "pageId": "p1ar6", "nearText": "Inventory",
  "setCells": [{ "row": "Widgets", "col": "Stock", "value": "12" }],
  "addColumns": [["Reorder point", "5", "10", "20"]],
  "deleteRows": ["Fasteners"] }
```

`row` takes an index or text in the row. `col` takes an index or a header name.
References resolve against the table as read, so `setCells` and `deleteRows` in
one call both land where intended.

OneNote cannot PATCH a `<tr>` or `<td>`, only a whole `<table>` by id, so every
cell edit rewrites the table.

- Untouched cells keep their markup. Cells carry across as HTML.
- Cell values are inline Markdown: `**bold**`, links, `<math>`.
- Row 0 counts as a header only if it already looks like one.
- `atCol` and `deleteColumns` pair with `atRow` and `deleteRows`.
- The result is rectangular.

## Section editing

`editPageSection` finds a heading, splices content into that section on the
server, and writes it back:

```jsonc
{ "pageId": "p1ar6", "heading": "Tasks", "mode": "append",
  "content": "- [ ] Write the report" }
```

`mode` is `replace` (default), `append`, `prepend` or `delete`. `delete` removes
the heading and its content, and needs no `content`. The other modes keep the
heading. OneNote has no delete action in its PATCH API, so `delete` splices the
nodes and writes the body back.

Images get the same guard as `replacePage`. Pass `allowImageLoss: true` to
override.

## Check boxes

`- [ ]` and `- [x]` become `<p data-tag="to-do">`. `setCheckbox` ticks by text
in one PATCH, however many boxes:

```jsonc
{ "pageId": "p1ar6", "text": ["Write the report", "Review the draft"] }
```

`checked: false` unticks. Nothing is written until every named box resolves.

OneNote has about 20 other note tags: important, question, idea, critical and so
on. None has Markdown syntax and no tool sets one, but they can be read.
`getPage` emits `[tag: important] paragraph text`, and
`getPage(format:"outline")` counts them per kind. Writing the marker back does
not attach a tag.

A raw tag passes through as a CommonMark HTML block. A raw tag inside a list
breaks list parsing, so keep tagged lines between blocks.

## Formula editing

1. `getMathExpressions(pageId)` returns each `<math>` block with its stored
   `mathml` and the containing block's element id.
2. `findAndReplace(pageId, find, replace)` swaps that fragment.

OneNote normalises MathML on write. `<mi>dy</mi>` comes back as
`<mrow><mi>d</mi><mi>y</mi></mrow>`. Take `find` from `getMathExpressions`, not
from the string you sent.

`findAndReplace` tries an exact HTML fragment first, then a run of visible text.
A replacement containing markup is refused. The result reports `mode`,
`occurrences` and `remaining`.

The PATCH scopes to the smallest replaceable block: `<p>`, heading, list, table
or nested `<div>`. The page's top-level positioned `<div>` is excluded because
OneNote has no `replace` for it, and targeting it corrupts the page. If nothing
narrower contains the fragment, the tool falls back to a whole-body replace and
reports `scope: "body"`.

`data-id` survives a PATCH. `id` from `includeIDs=true` does not: it is
snapshot-bound and needs a re-read after each edit. OneNote discards any `id` you
send. The converter stamps `data-id` on every heading and table.

## Images

`addImageToPage(pageId, { filePath | url, nearText?, position?, width?, height?, altText? })`.

OneNote needs raw binary in a `multipart/form-data` part, and does not fetch URLs
during a PATCH. The server reads the bytes and uploads them.
- Types: PNG, JPEG, GIF, BMP, TIFF. Limit 3.5 MB.
- A public `<img src="https://…">` works only in `createPage`.
- Images already on a page cannot be re-submitted. A body rewrite that would drop
  them is refused unless you pass `allowImageLoss: true`.
- `url` must be `http` or `https` and must not resolve to a loopback, private or
  link-local host ([src/images.js](src/images.js)).

## Token efficiency

- Tools return short handles (`p1ar6`, `s1wdi`, `nb1n8s`) instead of 85-character
  Graph ids, kept in `.id-map.json`. Raw ids work as input.
- Page content and section listings cache for 60 s and clear on write.
- `getPage(format:"outline")` lists headings, table headers and row labels
  without the body, which is enough to address a cell. `getPageSection`,
  `getMathExpressions`, `listPages(previewChars)` and `searchPages(previewChars)`
  narrow further.
- `editPageSection`, `editTable` and `findAndReplace` read, modify and write on
  the server. Only the new content crosses the wire.
- Tool schemas are re-sent on every request, so
  [test/server.test.mjs](test/server.test.mjs) enforces a byte budget.
- Graph errors reduce to `code` and `message`.
- Reading tools carry `readOnlyHint`. `deletePage` carries `destructiveHint`.

## Account limitations

These do not work on personal Microsoft accounts:

- Full-text page search. `searchPages(inContent: true)` reads pages instead,
  capped at 150 per call. Scope it with `sectionId` or `notebookId`.
- Deleting sections, notebooks or section groups. Only pages can be deleted.
- Copy and move. `copyToSection` and `copyToNotebook` return `501`.
  `copyPageToSection` falls back to re-creating the page from its HTML and
  reports `re-created`. Images cannot be re-submitted, so a page with any is
  refused unless you pass `allowImageLoss: true`. `copySectionToNotebook` cannot
  be emulated, and tells the caller to copy its pages instead.
- `.All` scopes. This server asks for `Notes.ReadWrite` and `Notes.Create`.

## Development

```bash
npm test
```

`node --test` with no arguments; a glob would need shell-specific quoting. The
suite runs offline. Layout:

- `onenote-mcp.mjs` holds the caches, content preparation and tool registrations.
- `src/content.js` converts Markdown to OneNote XHTML and back.
- `src/dom.js` handles text matching, section spans, splicing and check boxes.
- `src/ids.js`, `src/mathml.js`, `src/graph.js`, `src/images.js` mint handles,
  validate MathML, make Graph requests, and guard the image URL.

CI runs the suite on Node 20, 22 and 24.

## Local files

`.msal-cache.json` holds the refresh token and lives about 90 days.
`.access-token.txt` only appears if you used `saveAccessToken`, and lasts an
hour. `.id-map.json` is the handle lookup and is safe to delete. All three are
gitignored. The two token files are written `0600`.

## Troubleshooting

| Symptom | Cause |
|---|---|
| `401` from any tool | Token expired. Run `npm run auth`. |
| Server missing in Claude Desktop | The app was closed, not quit. Quit from the tray and reopen. |
| `node` not found on Windows | Use the full path in `command`, e.g. `C:\\Program Files\\nodejs\\node.exe`. |
| `EBADENGINE` or a startup syntax error | Node 18 or older. Run `nvm use`. |

## Credits

- [danosb/onenote-mcp](https://github.com/danosb/onenote-mcp). Direct ancestor.
  The device-code flow, the Graph request layer and the client setup come from
  here.
- [ZubeidHendricks/azure-onenote-mcp-server](https://github.com/ZubeidHendricks/azure-onenote-mcp-server).
  What danosb's work builds on.

Particular to this repository: the Markdown content model, style presets, table
addressing by row and column name, MathML validation, check-box handling, the
authoring guide, and the `src/` and `test/` trees.

## License

MIT. See [LICENSE](LICENSE).



