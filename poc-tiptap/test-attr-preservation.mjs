// Day 5 test: data-ooxml-* attributes survive a Tiptap round-trip.
//
// For each fixture HTML (synthetic + real bridge output):
//   1. Parse to ProseMirror JSON via generateJSON() with our extensions
//   2. Render back to HTML via generateHTML()
//   3. Verify every data-ooxml-* attribute that was in the input
//      is also in the output
//
// Pass criterion: 100% attribute preservation across all fixtures.
//
// Run:
//   node --import ./poc-tiptap/jsdom-bootstrap.mjs poc-tiptap/test-attr-preservation.mjs

import { readFileSync } from 'node:fs';
import { generateJSON, generateHTML } from '@tiptap/core';
import StarterKit from '@tiptap/starter-kit';
import Underline from '@tiptap/extension-underline';

import { preserveOoxmlExtensions } from './preserve-ooxml-extensions.mjs';

const extensions = [
  // Disable nodes/marks we replace with our preserving versions
  // (paragraph, heading, link) AND nodes we don't want at all in
  // body content (tightening the schema later, but for now just
  // avoid duplicates with our own extensions).
  StarterKit.configure({
    paragraph: false,
    heading: false,
    link: false,
    underline: false,
  }),
  Underline,
  ...preserveOoxmlExtensions,
];

// -------------------------------------------------------------------
// Fixtures
// -------------------------------------------------------------------

const synthetic = [
  {
    name: 'paragraph with pstyle',
    html: '<p data-ooxml-pstyle="Heading2">hello</p>',
  },
  {
    name: 'paragraph with ppr-extra blob',
    html: '<p data-ooxml-ppr-extra="ZmFrZWJsb2I=">x</p>',
  },
  {
    name: 'paragraph with numId/ilvl',
    html: '<p data-ooxml-numid="3" data-ooxml-ilvl="0">item</p>',
  },
  {
    name: 'heading with style',
    html: '<h2 data-ooxml-pstyle="Title">A title</h2>',
  },
  {
    name: 'inline SDT',
    html:
      '<p>Dear <span class="ooxml-sdt" data-ooxml-sdt-tag="recipient" data-ooxml-sdt-id="42">Recipient</span>,</p>',
  },
  {
    name: 'block SDT around two paragraphs',
    html:
      '<div class="ooxml-sdt block" data-ooxml-sdt-tag="contractBlock"><p>first</p><p>second</p></div>',
  },
  {
    name: 'rPr mark on a span',
    html:
      '<p>foo <span data-ooxml-rpr="ZmFrZQ==">bar</span> baz</p>',
  },
  {
    name: 'smartTag inline wrapper',
    html:
      '<p>TO <span class="ooxml-smarttag" data-ooxml-st-uri="ifinger" data-ooxml-st-element="data">WHOM IT MAY CONCERN</span></p>',
  },
  {
    name: 'opaque blob',
    html: '<div class="ooxml-blob" data-ooxml-blob="ZmFrZQ=="></div>',
  },
  {
    name: 'page break inside a paragraph',
    html:
      '<p>before<span class="ooxml-pagebreak" data-ooxml-br-type="page"></span>after</p>',
  },
  {
    name: 'tab marker',
    html: '<p>before<span class="ooxml-tab" data-ooxml-tab="">    </span>after</p>',
  },
  {
    name: 'drawing image',
    html:
      '<p><img class="ooxml-drawing" alt="" data-ooxml-embed-rid="rId7" data-ooxml-drawing="ZHJhd2luZw=="></p>',
  },
  {
    name: 'link with rid',
    html:
      '<p>see <a href="https://x" data-ooxml-rid="rId5">here</a> for more</p>',
  },
  {
    name: 'table with tbl-pr',
    html:
      '<table data-ooxml-tbl-pr="dGJsCg=="><tbody><tr data-ooxml-tr-pr="dHJSCg=="><td data-ooxml-tc-pr="dGNzCg=="><p>cell</p></td></tr></tbody></table>',
  },
  {
    name: 'nested formatting (strong > em > u)',
    html: '<p><strong><em><u>three styles</u></em></strong></p>',
  },
];

// -------------------------------------------------------------------
// Real bridge HTML — pick a few of the generated previews
// -------------------------------------------------------------------

const bridgeHtmlDir =
  '/Users/mario002e/IdeaProjects/bsh-migration/html-to-docx/poc-bridge/out/html';

const bridgeFixtures = [
  'Permit to Stay_Family_Old_version.html',
  'Official Custom Clearance ESCAP.html',
  'Roadtax IFAD.html',
];

for (const f of bridgeFixtures) {
  let raw;
  try {
    raw = readFileSync(`${bridgeHtmlDir}/${f}`, 'utf8');
  } catch {
    continue;
  }
  // The preview wrapper has <html><head>...<body>BODY</body></html>;
  // pull just the body inner content for parsing.
  const bodyMatch = raw.match(/<body>([\s\S]*?)<\/body>/i);
  if (!bodyMatch) continue;
  // Strip the H1 title that the preview prepends.
  const bodyInner = bodyMatch[1].replace(/^\s*<h1[^>]*>[^<]*<\/h1>\s*/i, '');
  synthetic.push({ name: `bridge fixture: ${f}`, html: bodyInner });
}

// -------------------------------------------------------------------
// Test runner
// -------------------------------------------------------------------

const dataAttrRegex = /data-ooxml-[a-z0-9-]+="[^"]*"/g;
const collectAttrs = (html) =>
  new Set([...html.matchAll(dataAttrRegex)].map((m) => m[0]));

let pass = 0;
let fail = 0;

for (const { name, html } of synthetic) {
  let json, htmlOut, errMsg;
  try {
    json = generateJSON(html, extensions);
    htmlOut = generateHTML(json, extensions);
  } catch (e) {
    errMsg = e.message;
  }

  if (errMsg) {
    fail += 1;
    console.log(`[FAIL] ${name.padEnd(50)} threw: ${errMsg}`);
    continue;
  }

  const inAttrs = collectAttrs(html);
  const outAttrs = collectAttrs(htmlOut);
  const lost = [...inAttrs].filter((a) => !outAttrs.has(a));
  const added = [...outAttrs].filter((a) => !inAttrs.has(a));

  if (lost.length === 0) {
    pass += 1;
    const noteAdded = added.length > 0 ? ` (added ${added.length})` : '';
    console.log(`[PASS] ${name.padEnd(50)} ${inAttrs.size} attrs preserved${noteAdded}`);
  } else {
    fail += 1;
    console.log(
      `[FAIL] ${name.padEnd(50)} lost ${lost.length}/${inAttrs.size}: ${lost.slice(0, 3).join(', ')}…`,
    );
  }
}

console.log(`\nAttribute preservation: ${pass}/${pass + fail}`);
process.exit(fail === 0 ? 0 : 1);
