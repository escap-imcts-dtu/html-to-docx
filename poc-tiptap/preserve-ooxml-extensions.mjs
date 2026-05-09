// Tiptap extensions that preserve our `data-ooxml-*` attributes
// through the editor.
//
// Tiptap (and ProseMirror underneath) strips unknown attributes by
// default — they're not in the schema. To round-trip the OOXML
// metadata our DOCX↔HTML bridge depends on, we have to declare every
// data-attribute as a real schema attribute on the node/mark that
// carries it.
//
// Two flavours of extension:
//   1. EXTENDED nodes — Paragraph / Heading / Link / Table / etc.
//      already exist in starter-kit; we extend them with extra
//      `addAttributes()` for our data-*.
//   2. NEW nodes/marks for OOXML constructs that have no Tiptap
//      cousin: SDT (inline + block), smartTag, customXml, opaque
//      blobs, page/column breaks, OOXML tabs, the rPr mark.
//
// All exported as a single array `preserveOoxmlExtensions` so callers
// can spread it into the extensions list:
//
//   const editor = new Editor({
//     extensions: [
//       StarterKit.configure({ paragraph: false, heading: false }),
//       Underline,
//       Link.configure({ ... }),
//       Table, TableRow, TableCell, TableHeader,
//       ...preserveOoxmlExtensions,
//     ],
//   });
//
// We unconfigure StarterKit's Paragraph/Heading because we replace
// them with our preserving versions.

import { Node, Mark } from '@tiptap/core';
import { Paragraph } from '@tiptap/extension-paragraph';
import { Heading } from '@tiptap/extension-heading';
import { Link } from '@tiptap/extension-link';
import { Table } from '@tiptap/extension-table';
import { TableRow } from '@tiptap/extension-table-row';
import { TableCell } from '@tiptap/extension-table-cell';
import { TableHeader } from '@tiptap/extension-table-header';

// -------------------------------------------------------------------
// Shared helper: declarative attribute spec
// -------------------------------------------------------------------
//
// Tiptap wants each preserved attribute as { default, parseHTML, renderHTML }.
// Most of ours follow the identity pattern: read data-foo, write data-foo.
// dataAttr() builds the spec for that pattern.

const dataAttr = (htmlName) => ({
  default: null,
  // Use hasAttribute to distinguish "absent" (return null) from
  // "present-but-empty" (return ''). The empty-string case matters
  // for marker attributes like data-ooxml-tab="" whose presence —
  // not whose value — is the carrier of meaning.
  parseHTML: (el) => (el.hasAttribute(htmlName) ? el.getAttribute(htmlName) : null),
  renderHTML: (attrs) => {
    const key = camelKeyForHtml(htmlName);
    const v = attrs[key];
    if (v == null) return {};
    return { [htmlName]: String(v) };
  },
});

// Tiptap stores attributes by camelCased key in JS, but the HTML name
// stays kebab-cased. We need to derive the JS key from the html name
// for the renderHTML lookup.
const camelKeyForHtml = (htmlName) =>
  htmlName.replace(/-([a-z])/g, (_, c) => c.toUpperCase());

// Build a `addAttributes()` body from a list of html attribute names.
// The keys in the returned object are camelCased html names so they
// match what dataAttr().renderHTML expects.
const dataAttrs = (htmlNames) => () => {
  const out = {};
  for (const html of htmlNames) {
    out[camelKeyForHtml(html)] = dataAttr(html);
  }
  return out;
};

const mergeAttrs = (parentFn, ourFn) => () => ({
  ...(parentFn?.() ?? {}),
  ...ourFn(),
});

// -------------------------------------------------------------------
// Extended built-in nodes
// -------------------------------------------------------------------

const PARAGRAPH_DATA_ATTRS = [
  'data-ooxml-pstyle',
  'data-ooxml-ppr-extra',
  'data-ooxml-numid',
  'data-ooxml-ilvl',
];

// preserveWhitespace: 'full' — needed everywhere our walker emits
// text verbatim (which preserves OOXML's xml:space="preserve" runs).
// Without it ProseMirror collapses leading/trailing spaces on text
// nodes that abut other text nodes, breaking text content fidelity
// (e.g. "Expiry" + " Passport1" → "ExpiryPassport1").

const withPreservedWhitespace = (parseRules) =>
  parseRules.map((r) => ({ ...r, preserveWhitespace: 'full' }));

export const PreservingParagraph = Paragraph.extend({
  addAttributes() {
    return mergeAttrs(this.parent, dataAttrs(PARAGRAPH_DATA_ATTRS))();
  },
  parseHTML() {
    return withPreservedWhitespace(this.parent?.() ?? [{ tag: 'p' }]);
  },
});

export const PreservingHeading = Heading.extend({
  addAttributes() {
    return mergeAttrs(this.parent, dataAttrs(PARAGRAPH_DATA_ATTRS))();
  },
  parseHTML() {
    return withPreservedWhitespace(this.parent?.() ?? []);
  },
});

const LINK_DATA_ATTRS = ['data-ooxml-rid', 'data-ooxml-anchor'];

export const PreservingLink = Link.extend({
  addAttributes() {
    return mergeAttrs(this.parent, dataAttrs(LINK_DATA_ATTRS))();
  },
});

const TABLE_DATA_ATTRS = ['data-ooxml-tbl-pr', 'data-ooxml-tbl-grid'];
export const PreservingTable = Table.extend({
  addAttributes() {
    return mergeAttrs(this.parent, dataAttrs(TABLE_DATA_ATTRS))();
  },
});

export const PreservingTableRow = TableRow.extend({
  addAttributes() {
    return mergeAttrs(this.parent, dataAttrs(['data-ooxml-tr-pr']))();
  },
});

export const PreservingTableCell = TableCell.extend({
  addAttributes() {
    return mergeAttrs(this.parent, dataAttrs(['data-ooxml-tc-pr']))();
  },
});

// -------------------------------------------------------------------
// New OOXML nodes
// -------------------------------------------------------------------

const SDT_DATA_ATTRS = [
  'data-ooxml-sdt-id',
  'data-ooxml-sdt-tag',
  'data-ooxml-sdt-alias',
  'data-ooxml-sdt-docpart',
  'data-ooxml-sdt-pr',
];

export const SdtInline = Node.create({
  name: 'sdtInline',
  group: 'inline',
  inline: true,
  content: 'inline*',
  selectable: true,
  addAttributes: dataAttrs(SDT_DATA_ATTRS),
  parseHTML() {
    return [
      {
        tag: 'span.ooxml-sdt',
        preserveWhitespace: 'full',
        getAttrs: (el) => {
          const cls = el.getAttribute('class') || '';
          if (/\bblock\b/.test(cls)) return false;
          return {};
        },
      },
    ];
  },
  renderHTML({ HTMLAttributes }) {
    return ['span', { class: 'ooxml-sdt', ...HTMLAttributes }, 0];
  },
});

export const SdtBlock = Node.create({
  name: 'sdtBlock',
  group: 'block',
  content: 'block*',
  defining: true,
  addAttributes: dataAttrs(SDT_DATA_ATTRS),
  parseHTML() {
    return [{ tag: 'div.ooxml-sdt.block', preserveWhitespace: 'full' }];
  },
  renderHTML({ HTMLAttributes }) {
    return ['div', { class: 'ooxml-sdt block', ...HTMLAttributes }, 0];
  },
});

const SMARTTAG_DATA_ATTRS = [
  'data-ooxml-st-uri',
  'data-ooxml-st-element',
  'data-ooxml-st-pr',
];

export const SmartTagInline = Node.create({
  name: 'smartTagInline',
  group: 'inline',
  inline: true,
  content: 'inline*',
  addAttributes: dataAttrs(SMARTTAG_DATA_ATTRS),
  parseHTML() {
    return [{ tag: 'span.ooxml-smarttag', preserveWhitespace: 'full' }];
  },
  renderHTML({ HTMLAttributes }) {
    return ['span', { class: 'ooxml-smarttag', ...HTMLAttributes }, 0];
  },
});

const CUSTOMXML_DATA_ATTRS = [
  'data-ooxml-cx-uri',
  'data-ooxml-cx-element',
  'data-ooxml-cx-pr',
];

export const CustomXmlInline = Node.create({
  name: 'customXmlInline',
  group: 'inline',
  inline: true,
  content: 'inline*',
  addAttributes: dataAttrs(CUSTOMXML_DATA_ATTRS),
  parseHTML() {
    return [{ tag: 'span.ooxml-customxml', preserveWhitespace: 'full' }];
  },
  renderHTML({ HTMLAttributes }) {
    return ['span', { class: 'ooxml-customxml', ...HTMLAttributes }, 0];
  },
});

// Opaque OOXML blob — INLINE atom, holds a base64-encoded chunk we
// can splice back. Most opaque w:* elements (fldSimple, fldChar,
// instrText, pict, proofErr, …) live INSIDE a paragraph in OOXML, so
// the placeholder must be inline-valid in the editor schema.
//
// (Block-level blobs are emitted via OoxmlDrawingBlob below; bare
//  block-level `<div class="ooxml-blob">` is also accepted for
//  backwards compat with older bridge HTML.)
export const OoxmlBlob = Node.create({
  name: 'ooxmlBlob',
  group: 'inline',
  inline: true,
  atom: true,
  selectable: true,
  addAttributes: dataAttrs(['data-ooxml-blob']),
  parseHTML() {
    return [
      {
        tag: 'span.ooxml-blob',
        getAttrs: (el) => {
          const cls = el.getAttribute('class') || '';
          // Don't grab the marker class (zero-width inline that
          // already has its own node).
          if (/\booxml-marker\b/.test(cls)) return false;
          return {};
        },
      },
    ];
  },
  renderHTML({ HTMLAttributes }) {
    return ['span', { class: 'ooxml-blob', ...HTMLAttributes }];
  },
});

// Backwards-compat for older bridge HTML that emitted block-level
// <div class="ooxml-blob">. Same wire format; different schema slot.
export const OoxmlBlobBlock = Node.create({
  name: 'ooxmlBlobBlock',
  group: 'block',
  atom: true,
  selectable: true,
  addAttributes: dataAttrs(['data-ooxml-blob']),
  parseHTML() {
    return [
      {
        tag: 'div.ooxml-blob',
        getAttrs: (el) => {
          const cls = el.getAttribute('class') || '';
          if (/\booxml-drawing-blob\b/.test(cls)) return false;
          return {};
        },
      },
    ];
  },
  renderHTML({ HTMLAttributes }) {
    return ['div', { class: 'ooxml-blob', ...HTMLAttributes }];
  },
});

export const OoxmlDrawingBlob = Node.create({
  name: 'ooxmlDrawingBlob',
  group: 'block',
  atom: true,
  addAttributes: dataAttrs(['data-ooxml-blob']),
  parseHTML() {
    return [{ tag: 'div.ooxml-drawing-blob' }];
  },
  renderHTML({ HTMLAttributes }) {
    return ['div', { class: 'ooxml-blob ooxml-drawing-blob', ...HTMLAttributes }];
  },
});

export const OoxmlMarker = Node.create({
  name: 'ooxmlMarker',
  group: 'inline',
  inline: true,
  atom: true,
  addAttributes: dataAttrs(['data-ooxml-blob']),
  parseHTML() {
    return [{ tag: 'span.ooxml-marker' }];
  },
  renderHTML({ HTMLAttributes }) {
    return ['span', { class: 'ooxml-marker', ...HTMLAttributes }];
  },
});

export const OoxmlPagebreak = Node.create({
  name: 'ooxmlPagebreak',
  group: 'inline',
  inline: true,
  atom: true,
  addAttributes: dataAttrs(['data-ooxml-br-type']),
  parseHTML() {
    return [{ tag: 'span.ooxml-pagebreak' }];
  },
  renderHTML({ HTMLAttributes }) {
    return ['span', { class: 'ooxml-pagebreak', ...HTMLAttributes }];
  },
});

export const OoxmlColumnbreak = Node.create({
  name: 'ooxmlColumnbreak',
  group: 'inline',
  inline: true,
  atom: true,
  addAttributes: dataAttrs(['data-ooxml-br-type']),
  parseHTML() {
    return [{ tag: 'span.ooxml-columnbreak' }];
  },
  renderHTML({ HTMLAttributes }) {
    return ['span', { class: 'ooxml-columnbreak', ...HTMLAttributes }];
  },
});

export const OoxmlTab = Node.create({
  name: 'ooxmlTab',
  group: 'inline',
  inline: true,
  atom: true,
  addAttributes: dataAttrs(['data-ooxml-tab']),
  parseHTML() {
    return [{ tag: 'span.ooxml-tab' }];
  },
  renderHTML({ HTMLAttributes }) {
    return [
      'span',
      { class: 'ooxml-tab', ...HTMLAttributes },
      // Visible representation: 4 spaces (matches what bodyOoxmlToHtml emits).
      '    ',
    ];
  },
});

export const OoxmlDrawingImg = Node.create({
  name: 'ooxmlDrawingImg',
  group: 'inline',
  inline: true,
  atom: true,
  addAttributes: dataAttrs([
    'src',
    'alt',
    'data-ooxml-embed-rid',
    'data-ooxml-drawing',
  ]),
  parseHTML() {
    return [{ tag: 'img.ooxml-drawing' }];
  },
  renderHTML({ HTMLAttributes }) {
    return ['img', { class: 'ooxml-drawing', ...HTMLAttributes }];
  },
});

// -------------------------------------------------------------------
// rPr wrapper (run-property preservation)
// -------------------------------------------------------------------
//
// Implemented as a Node, not a Mark. Marks of the same type collapse
// when nested in ProseMirror — only the inner one survives. Our HTML
// sometimes has nested <span data-ooxml-rpr=A><span data-ooxml-rpr=B>
// when a child run inherited part of its parent's properties; both
// must survive. A Node preserves the wrapper structure literally.

export const OoxmlRpr = Node.create({
  name: 'ooxmlRpr',
  group: 'inline',
  inline: true,
  content: 'inline*',
  addAttributes: dataAttrs(['data-ooxml-rpr']),
  parseHTML() {
    return [
      {
        tag: 'span[data-ooxml-rpr]',
        // OOXML uses `<w:t xml:space="preserve">` to retain
        // leading/trailing spaces inside a run. Our walker drops the
        // attribute when emitting HTML (it's purely textual on the way
        // out) but ProseMirror still needs to be told NOT to collapse
        // whitespace inside our wrapper, otherwise text like
        // `<span> Passport1</span>` loses the leading space and the
        // SDT-display-text round-trip breaks.
        preserveWhitespace: 'full',
      },
    ];
  },
  renderHTML({ HTMLAttributes }) {
    return ['span', HTMLAttributes, 0];
  },
});

// -------------------------------------------------------------------
// Aggregated export
// -------------------------------------------------------------------

export const preserveOoxmlExtensions = [
  PreservingParagraph,
  PreservingHeading,
  PreservingLink,
  PreservingTable,
  PreservingTableRow,
  TableHeader, // header cells are uncommon in our HTML output but needed by Table
  PreservingTableCell,
  SdtInline,
  SdtBlock,
  SmartTagInline,
  CustomXmlInline,
  OoxmlBlob,
  OoxmlBlobBlock,
  OoxmlDrawingBlob,
  OoxmlMarker,
  OoxmlPagebreak,
  OoxmlColumnbreak,
  OoxmlTab,
  OoxmlDrawingImg,
  OoxmlRpr,
];

export default preserveOoxmlExtensions;
