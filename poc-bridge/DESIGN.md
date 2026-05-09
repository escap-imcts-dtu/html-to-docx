# DocxBag — lossless DOCX↔HTML bridge (design)

**Status:** Day 1 design.
**Sibling implementations:**
- `../html-to-docx/` — HTML→DOCX (TurboDocx fork + our fragment API)
- `../mammoth.js/` — DOCX→HTML (mammoth fork; will be extended for lossless attrs)
- `../html-to-docx/poc-bridge/` — this design lives here, bridge code lives here

## Why a sidecar instead of a pure round-trip

Audit of `Permit to Stay_Family_Old_version.docx` (the smallest HCRU MFA template):
- 35 OOXML parts in the .docx
- 93 SDT Content Controls in the body alone (the form-field mechanism)
- 12 customXml parts (Word's databindings + item-prop metadata)
- Glossary subtree (AutoText / Building Blocks)
- 2 header parts; theme; settings; fontTable; footnotes; endnotes
- All are dropped by mammoth's default body output. None has a clean HTML equivalent.

Reconstructing all of that from HTML alone — even with `data-ooxml-*` annotations —
is complex, brittle, and not obviously possible (custom XML, glossary entries, theme
references, etc. live outside the runs).

The sidecar approach trades that brittleness for a simple invariant: **only the body
of `word/document.xml` is editable in HTML; every other byte of the .docx is preserved
verbatim**. Combined with `data-ooxml-*` attributes for the body's own OOXML
constructs (SDT identity, run properties, paragraph styles, list refs), we get true
losslessness for the realistic edit shape — fill / amend a template — while keeping
the implementation tractable.

## What is "body"?

In v1, **body = inner of `<w:body>` in `word/document.xml`, excluding `<w:sectPr>`**.
- The body content is a sequence of `<w:p>` and `<w:tbl>` elements.
- `<w:sectPr>` (page geometry, header/footer references, columns, etc.) is part of
  the shell.
- Headers (`word/header*.xml`), footers (`word/footer*.xml`), footnotes, endnotes,
  glossary, comments — all shell. Authors do not edit those in the visual editor.

A future v2 may add per-part body extraction for headers/footers if a real use case
emerges. The DocxBag API is shaped to allow this without breaking callers.

## Lossless contract

For the round-trip `docx → DocxBag → HTML → (edit) → HTML → DocxBag.replaceBody → docx`:

1. **Every part of the output zip that is not `word/document.xml` is byte-identical
   to the input.** No exceptions. Verified by SHA-256 per part.

2. **Within `word/document.xml`, the shell is byte-identical**: XML declaration, the
   `<w:document …>` opening tag (with all its `xmlns:*` attrs), the `<w:body>`
   wrapper, and `<w:sectPr>…</w:sectPr>` at the end. Only the children of `<w:body>`
   between the start and `<w:sectPr>` change.

3. **If the user makes no edits to the HTML**, the new body XML is structurally
   equivalent to the original: same SDT identifiers in the same positions, same
   paragraph/run boundaries, same content. Whitespace differences inside element
   markup are tolerated; semantic content is preserved.

4. **If the user makes edits**, edits are reflected in the new body XML; SDTs that
   the user did not delete are preserved by identity.

5. **What is explicitly NOT preserved**:
   - SDTs the user deleted in the editor.
   - Run-level formatting the user changed (e.g., applied bold to a previously plain
     run) — the OUTPUT reflects the user's edit.
   - Inline images the user removed.

   These are user-intent losses, not roundtrip losses.

## Architecture

```
┌──────────┐  load   ┌─────────────────────────────┐
│ DOCX bytes├────────►│ DocxBag                     │
└──────────┘         │  • parts: Map<path,Buffer>  │
                     │  • body: string (extracted) │
                     │  • shell-of-document: pre,  │
                     │    post (string snippets)   │
                     └────┬─────────────────┬──────┘
                          │ getBody()        │ replaceBody(newXml)
                          ▼                  ▲
                     ┌────────────────────────────────┐
                     │  body XML (the editable part)  │
                     └────┬─────────────────┬─────────┘
                          │ docx→html       │ html→ooxml
                          ▼                  ▲
                     ┌─────────────────┐ ┌─────────────────┐
                     │ extended mammoth│ │ extended html-  │
                     │  (Day 2)        │ │ to-docx (Day 3) │
                     └─────────────────┘ └─────────────────┘
```

## DocxBag API (v1, locked-in for Day 1)

```js
class DocxBag {
  constructor(buffer)            // buffer: input .docx Buffer
  static async fromBuffer(buf)   // async factory (returns DocxBag instance)

  // Body access (read-only side)
  getBodyContent(): string       // children of <w:body> minus <w:sectPr>
  getBodyShell(): { pre: string, post: string }
                                 // pre = "<?xml…?>...<w:document …><w:body>"
                                 // post = "<w:sectPr>…</w:sectPr></w:body></w:document>"

  // Body mutation
  replaceBodyContent(newXml: string): void
                                 // replace inner of <w:body> (everything before <w:sectPr>)

  // Shell parts (read-only)
  listParts(): string[]          // every entry in the zip, sorted
  getPart(path: string): Buffer | undefined
                                 // raw bytes; for binary parts (images) this is the image
  getPartText(path: string): string | undefined
                                 // utf-8 string; convenience for XML parts

  // Output
  serialize(): Buffer            // emit .docx Buffer with body replaced
}
```

## `data-ooxml-*` attribute schema (locked-in for Day 2)

The HTML emitted by docx→html carries these attributes when relevant:

| HTML element | data-ooxml-* attribute | Purpose | Example value |
|---|---|---|---|
| any | `data-ooxml-sdt-id` | SDT internal id | `100001` |
| any | `data-ooxml-sdt-tag` | SDT `<w:tag w:val>` | `recipientName` |
| any | `data-ooxml-sdt-alias` | SDT `<w:alias w:val>` | `Recipient` |
| any | `data-ooxml-sdt-placeholder` | inner display text from `<w:sdtContent>` | `Title1` |
| `<span>` | `data-ooxml-rpr` | base64 of `<w:rPr>` for runs with non-trivial properties | … |
| `<p>` | `data-ooxml-pstyle` | `<w:pStyle w:val>` | `Heading2` |
| `<p>` | `data-ooxml-ppr-extra` | base64 of `<w:pPr>` minus pStyle (for unusual props) | … |
| `<ol>`/`<ul>`/`<li>` | `data-ooxml-numid` | `<w:numId w:val>` | `3` |
| `<ol>`/`<ul>`/`<li>` | `data-ooxml-ilvl` | `<w:ilvl w:val>` | `0` |
| `<a>` | `data-ooxml-rid` | hyperlink rId in `word/_rels/document.xml.rels` | `rId7` |
| `<img>` | `data-ooxml-rid` | image embed rId | `rId8` |

**Why base64 for `data-ooxml-rpr` and `data-ooxml-ppr-extra`?** Run/paragraph
properties have many small subelements (`<w:color>`, `<w:sz>`, `<w:rFonts>`, etc.)
that don't have HTML equivalents and whose enumeration would require a long flat
namespace of attrs. Base64'd OOXML chunks let us stash whatever's there without
designing a custom serialization. The reverse converter just decodes and splices.

**Why `<span>` for runs?** Mammoth uses semantic HTML (`<strong>`, `<em>`) for
common run properties; we layer `<span data-ooxml-rpr="…">` underneath only when a
run has properties that don't map to a clean HTML tag. This keeps the output usable
as plain HTML for non-roundtripping consumers (preview, search index, etc.).

## Round-trip identity tests (Day 1 + Day 4)

**Day 1 test (this work):**
- `bag = await DocxBag.fromBuffer(originalDocx)`
- `body = bag.getBodyContent()`
- `bag.replaceBodyContent(body)`  // identity replacement
- `out = bag.serialize()`
- For each part in the input zip:
  - if `path !== 'word/document.xml'`: assert `sha256(input[path]) == sha256(out[path])`
  - if `path === 'word/document.xml'`: assert shell preserved (regex / parse), body content equal modulo whitespace

**Day 4 test (lossless roundtrip):**
- `bag = await DocxBag.fromBuffer(originalDocx)`
- `html = await docxToLosslessHtml(bag.getBodyContent())`
- `newBody = await losslessHtmlToOoxml(html)`
- `bag.replaceBodyContent(newBody)`
- `out = bag.serialize()`
- Same per-part assertions as Day 1, with body content compared via canonical OOXML
  semantic equivalence (XML normalization + element-tree compare, ignoring
  whitespace).

## Out of scope (for now)

- Editing headers / footers / glossary in HTML. They survive but are not editable
  via this bridge. A v2 can extend the API.
- Tracked changes, comments, footnote/endnote insertion via HTML.
- DOCX features that have no HTML cousin: VML shapes, Office Math (we'd preserve
  them as base64 OOXML chunks but won't render them in the editor).

## Open questions (decide on Day 2 with evidence)

1. Should we extend mammoth in-tree, or wrap it with a post-processor that walks
   the input OOXML separately and grafts `data-ooxml-*` attrs onto mammoth's
   output? The wrap approach is less invasive but requires a second OOXML parse.

2. For SDTs that mammoth outputs as plain text: do we wrap mammoth's `<p>foo</p>`
   in `<span data-ooxml-sdt-tag="…">foo</span>`, or do we replace the run entirely
   with our own `<span>`? Affects how Tiptap displays it.

3. For the existing `transformSdtsToBraces` pipeline in ESCAP — does our bridge
   coexist (different code paths for "edit" vs "render") or do we want to unify?

These are real design questions; we will answer them with code and tests on Day 2.
