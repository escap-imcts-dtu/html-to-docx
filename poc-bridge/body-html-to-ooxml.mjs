// Body HTML → OOXML (reverse of body-ooxml-to-html.mjs).
//
// Walks the HTML produced by bodyOoxmlToHtml — possibly edited by a
// Tiptap-style editor — and emits OOXML that's structurally equivalent
// to the original.
//
// Lossless contract (DESIGN.md):
//   - data-ooxml-* attributes carry preserved OOXML metadata.
//   - Where present, that metadata is the source of truth.
//   - Where ABSENT (because the user added new content), we emit
//     reasonable default OOXML based on the semantic HTML.
//
// Key invariants enforced on output:
//   1. Top-level body content is a sequence of <w:p> and <w:tbl>
//      (and structural extras like <w:sdt> when block-level).
//   2. Every <w:p> may have <w:pPr> first, followed by runs / inline elements.
//   3. Every <w:r> may have <w:rPr> first, followed by content
//      (<w:t>, <w:tab>, <w:br>, <w:drawing>, …).
//   4. Text with leading/trailing whitespace gets xml:space="preserve".
//
// Property stacking: Tiptap-style nested formatting (e.g.
// <strong><em>foo</em></strong>) is flattened into a single <w:r>
// with combined <w:rPr> children. Outer = earlier in stack; inner
// overrides on conflicts.

import { DOMParser, XMLSerializer } from '@xmldom/xmldom';

// -------------------------------------------------------------------
// Constants
// -------------------------------------------------------------------

const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

const isElement = (n) => n && n.nodeType === 1;
const isText = (n) => n && n.nodeType === 3;

const HEADING_TAG_RE = /^h([1-6])$/i;

// -------------------------------------------------------------------
// Public API
// -------------------------------------------------------------------

/**
 * @param {string} html
 *   HTML produced by bodyOoxmlToHtml (Tiptap-edited or not).
 * @returns {string} body OOXML — sequence of <w:p>/<w:tbl>/etc.
 *   suitable for splicing into DocxBag.replaceBodyContent().
 */
export function bodyHtmlToOoxml(html) {
  if (typeof html !== 'string') {
    throw new TypeError('bodyHtmlToOoxml expects a string');
  }
  if (html.length === 0) return '';

  // Wrap in synthetic root for a clean fragment parse. xmldom requires
  // a single root; we strip it after.
  const wrapped =
    `<__bagroot xmlns="http://www.w3.org/1999/xhtml">` + sanitizeForXmlParse(html) + `</__bagroot>`;
  const dom = new DOMParser({
    onError: (level, msg) => {
      if (level === 'fatalError' || level === 'error') {
        throw new Error(`bodyHtmlToOoxml: parse ${level}: ${msg.split('\n')[0]}`);
      }
    },
  }).parseFromString(wrapped, 'text/xml');

  const root = dom.documentElement;
  if (!root) throw new Error('bodyHtmlToOoxml: no root after parse');
  return blockify(root.childNodes);
}

// -------------------------------------------------------------------
// HTML quirks → XML-safe
// -------------------------------------------------------------------

/**
 * The HTML we emit is XML-safe by construction (escapeText/escapeAttr
 * cover &/</>) BUT user edits via Tiptap may include HTML void
 * elements without self-closing slashes ( e.g. <br>, <img …>, <hr>),
 * which choke an XML parser. Rewrite those to self-closing form.
 *
 * Also normalize a few entities that XML doesn't recognize by name
 * (HTML's &nbsp; etc.) to numeric form.
 */
function sanitizeForXmlParse(html) {
  return html
    .replace(/<(br|img|hr|input|meta|link|area|base|col|embed|param|source|track|wbr)\b([^>]*?)(?<!\/)>/gi, '<$1$2/>')
    .replace(/&nbsp;/g, '&#160;')
    .replace(/&copy;/g, '&#169;')
    .replace(/&reg;/g, '&#174;')
    .replace(/&trade;/g, '&#8482;')
    .replace(/&hellip;/g, '&#8230;')
    .replace(/&mdash;/g, '&#8212;')
    .replace(/&ndash;/g, '&#8211;')
    .replace(/&lsquo;/g, '&#8216;')
    .replace(/&rsquo;/g, '&#8217;')
    .replace(/&ldquo;/g, '&#8220;')
    .replace(/&rdquo;/g, '&#8221;');
}

// -------------------------------------------------------------------
// Block-vs-inline dispatch
// -------------------------------------------------------------------

const isBlockElement = (el) => {
  if (!isElement(el)) return false;
  const name = el.localName.toLowerCase();
  if (['p', 'table', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6'].includes(name)) return true;
  if (name === 'div') {
    const cls = el.getAttribute('class') || '';
    if (/\b(ooxml-sdt\s+block|ooxml-blob|ooxml-drawing-blob)\b/.test(cls)) return true;
    // Plain <div> from a Tiptap edit — treat as paragraph wrapper
    return true;
  }
  return false;
};

/**
 * Convert a sequence of HTML child nodes into a sequence of
 * top-level body elements (paragraphs, tables, blobs, ...).
 * Inline content not enclosed in a block wrapper is auto-wrapped
 * in a <w:p>.
 */
function blockify(childNodes) {
  const out = [];
  let inlineBuf = [];

  const flushInline = () => {
    if (inlineBuf.length === 0) return;
    const innerXml = inlineify(inlineBuf, []);
    if (innerXml.length > 0) {
      out.push(`<w:p>${innerXml}</w:p>`);
    }
    inlineBuf = [];
  };

  for (let i = 0; i < childNodes.length; i += 1) {
    const child = childNodes[i];
    if (isText(child)) {
      // Whitespace-only text between blocks is irrelevant.
      if ((child.nodeValue ?? '').trim() === '') continue;
      inlineBuf.push(child);
      continue;
    }
    if (!isElement(child)) continue;
    if (isBlockElement(child)) {
      flushInline();
      out.push(emitBlock(child));
    } else {
      inlineBuf.push(child);
    }
  }
  flushInline();
  return out.join('');
}

function emitBlock(el) {
  const name = el.localName.toLowerCase();
  if (name === 'p' || HEADING_TAG_RE.test(name)) return emitParagraph(el);
  if (name === 'table') return emitTable(el);
  if (name === 'div') {
    const cls = el.getAttribute('class') || '';
    if (/\booxml-sdt\s+block\b/.test(cls)) return emitSdt(el, /*isBlock*/ true);
    // Block-level blob (legacy — older HTML used <div class="ooxml-blob">).
    if (/\booxml-blob\b/.test(cls) || /\booxml-drawing-blob\b/.test(cls)) {
      return decodeAttrAsXml(el, 'data-ooxml-blob');
    }
    // Plain div from Tiptap — emit as paragraph wrapping its inline content.
    const inner = inlineify([...el.childNodes], []);
    return `<w:p>${inner}</w:p>`;
  }
  // Fallback: unknown block — emit a paragraph wrapping its text.
  const text = el.textContent || '';
  return `<w:p><w:r><w:t${needsXmlSpace(text) ? ' xml:space="preserve"' : ''}>${escapeText(text)}</w:t></w:r></w:p>`;
}

// -------------------------------------------------------------------
// Paragraph
// -------------------------------------------------------------------

function emitParagraph(p) {
  const name = p.localName.toLowerCase();
  const headingMatch = name.match(HEADING_TAG_RE);

  const explicitPstyle = p.getAttribute('data-ooxml-pstyle') || '';
  const pStyle = explicitPstyle || (headingMatch ? `Heading${headingMatch[1]}` : '');
  const pprExtraB64 = p.getAttribute('data-ooxml-ppr-extra') || '';
  const numId = p.getAttribute('data-ooxml-numid') || '';
  const ilvl = p.getAttribute('data-ooxml-ilvl') || '';

  const pPrParts = [];
  if (pStyle) pPrParts.push(`<w:pStyle w:val="${escapeAttr(pStyle)}"/>`);
  if (numId) {
    pPrParts.push(
      `<w:numPr><w:ilvl w:val="${escapeAttr(ilvl || '0')}"/><w:numId w:val="${escapeAttr(numId)}"/></w:numPr>`,
    );
  }
  if (pprExtraB64) {
    // The blob contains a full <w:pPr xmlns:w="...">…</w:pPr>; strip
    // the wrapper and splice its children.
    pPrParts.push(unwrapXmlChildren(decodeBase64Utf8(pprExtraB64)));
  }
  const pPrXml = pPrParts.length > 0 ? `<w:pPr>${pPrParts.join('')}</w:pPr>` : '';

  const inner = inlineify([...p.childNodes], []);
  return `<w:p>${pPrXml}${inner}</w:p>`;
}

// -------------------------------------------------------------------
// Inline content
// -------------------------------------------------------------------

/**
 * Convert an array of inline-context HTML nodes into a sequence of
 * <w:r>/<w:hyperlink>/<w:sdt>/etc. The rPrStack accumulates
 * formatting in scope (each frame is a list of <w:rPr> child XML
 * fragments).
 */
function inlineify(nodes, rPrStack) {
  const out = [];
  for (let i = 0; i < nodes.length; i += 1) {
    const n = nodes[i];
    if (isText(n)) {
      const text = n.nodeValue || '';
      if (text === '') continue;
      out.push(emitTextRun(text, rPrStack));
      continue;
    }
    if (!isElement(n)) continue;
    out.push(emitInlineElement(n, rPrStack));
  }
  return out.join('');
}

function emitInlineElement(el, rPrStack) {
  const tag = el.localName.toLowerCase();
  const cls = el.getAttribute('class') || '';

  // SDT (inline OR block; here we handle inline span variant)
  if (/\booxml-sdt\b/.test(cls)) {
    return emitSdt(el, /*isBlock*/ false);
  }
  // Deprecated transparent inline wrappers (smartTag, customXml).
  if (/\booxml-smarttag\b/.test(cls)) {
    return emitTransparentWrapper(el, rPrStack, 'smartTag', 'st');
  }
  if (/\booxml-customxml\b/.test(cls)) {
    return emitTransparentWrapper(el, rPrStack, 'customXml', 'cx');
  }
  // Blob marker (zero-width, holds e.g. bookmarkStart in base64).
  if (/\booxml-marker\b/.test(cls)) {
    const b64 = el.getAttribute('data-ooxml-blob') || '';
    return b64 ? decodeBase64Utf8(b64) : '';
  }
  // Inline opaque blob (fldSimple / fldChar / instrText / pict / etc.)
  if (/\booxml-blob\b/.test(cls)) {
    const b64 = el.getAttribute('data-ooxml-blob') || '';
    return b64 ? decodeBase64Utf8(b64) : '';
  }
  // Page / column break markers
  if (/\booxml-pagebreak\b/.test(cls)) {
    return `<w:r>${composeRPr(rPrStack)}<w:br w:type="page"/></w:r>`;
  }
  if (/\booxml-columnbreak\b/.test(cls)) {
    return `<w:r>${composeRPr(rPrStack)}<w:br w:type="column"/></w:r>`;
  }
  // Tab marker
  if (/\booxml-tab\b/.test(cls)) {
    return `<w:r>${composeRPr(rPrStack)}<w:tab/></w:r>`;
  }
  // Drawing image
  if (tag === 'img' && /\booxml-drawing\b/.test(cls)) {
    const dB64 = el.getAttribute('data-ooxml-drawing') || '';
    const drawing = dB64 ? decodeBase64Utf8(dB64) : '';
    if (drawing) {
      return `<w:r>${composeRPr(rPrStack)}${drawing}</w:r>`;
    }
    return '';
  }
  // Plain user-added <img> (no preserved drawing) — best-effort: skip
  // (a Tiptap-uploaded image needs a media-asset pipeline, not in scope)
  if (tag === 'img') return '';

  // Soft line break
  if (tag === 'br') return `<w:r>${composeRPr(rPrStack)}<w:br/></w:r>`;

  // Hyperlink
  if (tag === 'a') return emitHyperlink(el, rPrStack);

  // Semantic formatting tags push onto the rPr stack.
  if (tag === 'strong' || tag === 'b') return inlineify([...el.childNodes], pushFlag(rPrStack, 'b'));
  if (tag === 'em' || tag === 'i') return inlineify([...el.childNodes], pushFlag(rPrStack, 'i'));
  if (tag === 'u') return inlineify([...el.childNodes], pushFlag(rPrStack, 'u'));
  if (tag === 's' || tag === 'strike' || tag === 'del') return inlineify([...el.childNodes], pushFlag(rPrStack, 'strike'));

  // <span data-ooxml-rpr="..."> — push the decoded rPr children
  if (tag === 'span') {
    const rprB64 = el.getAttribute('data-ooxml-rpr') || '';
    if (rprB64) {
      const childrenXml = unwrapXmlChildren(decodeBase64Utf8(rprB64));
      return inlineify([...el.childNodes], pushChildrenXml(rPrStack, childrenXml));
    }
    // Plain span without preserved rPr — recurse without changes.
    return inlineify([...el.childNodes], rPrStack);
  }

  // Unknown inline element — fall through, treat as transparent.
  return inlineify([...el.childNodes], rPrStack);
}

function emitTextRun(text, rPrStack) {
  return `<w:r>${composeRPr(rPrStack)}<w:t${needsXmlSpace(text) ? ' xml:space="preserve"' : ''}>${escapeText(text)}</w:t></w:r>`;
}

function emitTransparentWrapper(el, rPrStack, ooxmlTag, attrPrefix) {
  const uri = el.getAttribute(`data-ooxml-${attrPrefix}-uri`) || '';
  const element = el.getAttribute(`data-ooxml-${attrPrefix}-element`) || '';
  const prB64 = el.getAttribute(`data-ooxml-${attrPrefix}-pr`) || '';
  const inner = inlineify([...el.childNodes], rPrStack);
  const wrapperAttrs = [];
  if (uri) wrapperAttrs.push(`w:uri="${escapeAttr(uri)}"`);
  if (element) wrapperAttrs.push(`w:element="${escapeAttr(element)}"`);
  const prXml = prB64 ? decodeBase64Utf8(prB64) : '';
  const attrStr = wrapperAttrs.length > 0 ? ' ' + wrapperAttrs.join(' ') : '';
  return `<w:${ooxmlTag}${attrStr}>${prXml}${inner}</w:${ooxmlTag}>`;
}

function emitHyperlink(a, rPrStack) {
  const rId = a.getAttribute('data-ooxml-rid') || '';
  const anchor = a.getAttribute('data-ooxml-anchor') || '';
  const inner = inlineify([...a.childNodes], rPrStack);
  const attrs = [];
  if (rId) attrs.push(`r:id="${escapeAttr(rId)}"`);
  if (anchor) attrs.push(`w:anchor="${escapeAttr(anchor)}"`);
  return `<w:hyperlink${attrs.length ? ' ' + attrs.join(' ') : ''}>${inner}</w:hyperlink>`;
}

// -------------------------------------------------------------------
// SDT
// -------------------------------------------------------------------

function emitSdt(el, isBlock) {
  const sdtPrB64 = el.getAttribute('data-ooxml-sdt-pr') || '';
  let sdtPrXml = '';
  if (sdtPrB64) {
    sdtPrXml = decodeBase64Utf8(sdtPrB64);
  } else {
    // Reconstruct minimal sdtPr from any individual data-ooxml-sdt-* attrs.
    const tag = el.getAttribute('data-ooxml-sdt-tag') || '';
    const alias = el.getAttribute('data-ooxml-sdt-alias') || '';
    const id = el.getAttribute('data-ooxml-sdt-id') || '';
    const docPart = el.getAttribute('data-ooxml-sdt-docpart') || '';
    const parts = [];
    if (alias) parts.push(`<w:alias w:val="${escapeAttr(alias)}"/>`);
    if (tag) parts.push(`<w:tag w:val="${escapeAttr(tag)}"/>`);
    if (id) parts.push(`<w:id w:val="${escapeAttr(id)}"/>`);
    if (docPart) {
      parts.push(`<w:placeholder><w:docPart w:val="${escapeAttr(docPart)}"/></w:placeholder>`);
    }
    sdtPrXml = parts.length > 0 ? `<w:sdtPr>${parts.join('')}</w:sdtPr>` : '';
  }

  // Inner content of sdtContent: block SDT contains block-level
  // elements (paragraphs/tables); inline SDT contains inline runs.
  const inner = isBlock
    ? blockify([...el.childNodes])
    : inlineify([...el.childNodes], []);
  return `<w:sdt>${sdtPrXml}<w:sdtContent>${inner}</w:sdtContent></w:sdt>`;
}

// -------------------------------------------------------------------
// Table
// -------------------------------------------------------------------

function emitTable(table) {
  const tblPrB64 = table.getAttribute('data-ooxml-tbl-pr') || '';
  const tblGridB64 = table.getAttribute('data-ooxml-tbl-grid') || '';
  const tblPr = tblPrB64 ? decodeBase64Utf8(tblPrB64) : '';
  const tblGrid = tblGridB64 ? decodeBase64Utf8(tblGridB64) : '';

  const rows = [];
  const trs = collectDescendantsByTag(table, 'tr');
  for (const tr of trs) rows.push(emitTableRow(tr));

  return `<w:tbl>${tblPr}${tblGrid}${rows.join('')}</w:tbl>`;
}

function emitTableRow(tr) {
  const trPrB64 = tr.getAttribute('data-ooxml-tr-pr') || '';
  const trPr = trPrB64 ? decodeBase64Utf8(trPrB64) : '';
  const cells = [];
  const tds = childrenByTag(tr, 'td');
  for (const td of tds) cells.push(emitTableCell(td));
  // Empty rows still need a cell to be valid OOXML.
  if (cells.length === 0) cells.push('<w:tc><w:p/></w:tc>');
  return `<w:tr>${trPr}${cells.join('')}</w:tr>`;
}

function emitTableCell(td) {
  const tcPrB64 = td.getAttribute('data-ooxml-tc-pr') || '';
  const tcPr = tcPrB64 ? decodeBase64Utf8(tcPrB64) : '';
  const inner = blockify([...td.childNodes]);
  // OOXML requires at least one <w:p> in every cell.
  const safeInner = inner.includes('<w:p') || inner.includes('<w:tbl') ? inner : '<w:p/>';
  return `<w:tc>${tcPr}${safeInner}</w:tc>`;
}

// -------------------------------------------------------------------
// rPr stack management
// -------------------------------------------------------------------

/**
 * Each stack frame is either:
 *   { kind: 'flag', name: 'b'|'i'|'u'|'strike' }
 *   { kind: 'xml', children: '<w:rFonts …/><w:sz …/>' }   // pre-serialized rPr children
 *
 * composeRPr() collapses the stack into a single <w:rPr>…</w:rPr>.
 * Order: outer frames first (so inner frames can override duplicate
 * properties; for our flag set, duplicates are idempotent so order
 * doesn't change semantics, but we keep the outer-first convention
 * to match what Word writes).
 */

function pushFlag(stack, name) {
  return [...stack, { kind: 'flag', name }];
}

function pushChildrenXml(stack, childrenXml) {
  return [...stack, { kind: 'xml', children: childrenXml }];
}

function composeRPr(stack) {
  if (stack.length === 0) return '';
  const seenFlags = new Set();
  const parts = [];
  for (const frame of stack) {
    if (frame.kind === 'flag') {
      if (!seenFlags.has(frame.name)) {
        parts.push(`<w:${frame.name}/>`);
        seenFlags.add(frame.name);
      }
    } else if (frame.kind === 'xml' && frame.children) {
      parts.push(frame.children);
    }
  }
  if (parts.length === 0) return '';
  return `<w:rPr>${parts.join('')}</w:rPr>`;
}

// -------------------------------------------------------------------
// Helpers
// -------------------------------------------------------------------

const escapeAttr = (s) =>
  String(s)
    .replace(/&/g, '&amp;')
    .replace(/"/g, '&quot;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');

const escapeText = (s) =>
  String(s)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');

const decodeBase64Utf8 = (b64) => Buffer.from(b64, 'base64').toString('utf8');

const needsXmlSpace = (text) => {
  if (text.length === 0) return false;
  const first = text.charCodeAt(0);
  const last = text.charCodeAt(text.length - 1);
  // Whitespace = space, tab, LF, CR
  return (
    first === 32 || first === 9 || first === 10 || first === 13 ||
    last === 32 || last === 9 || last === 10 || last === 13
  );
};

/**
 * Strip the outer wrapper of an XML fragment like
 *   <w:rPr xmlns:w="…">…children…</w:rPr>
 * and return just the "…children…" part.
 *
 * We DON'T re-parse — that would re-serialize and lose attribute
 * order. Instead, we string-strip the first `>` and the last `</…>`.
 */
function unwrapXmlChildren(xml) {
  if (!xml) return '';
  const firstClose = xml.indexOf('>');
  const lastOpen = xml.lastIndexOf('<');
  if (firstClose < 0 || lastOpen < firstClose) return xml;
  return xml.slice(firstClose + 1, lastOpen);
}

/**
 * Decode a base64 attribute and return its content as XML string,
 * suitable for splicing inline. The decoded XML is one well-formed
 * element (e.g. <w:bookmarkStart …/> or <w:fldSimple>…</w:fldSimple>).
 * We just return it as-is — the parent <w:p> wrapper handles namespace
 * scoping.
 */
function decodeAttrAsXml(el, attrName) {
  const b64 = el.getAttribute(attrName);
  if (!b64) return '';
  return decodeBase64Utf8(b64);
}

function childrenByTag(parent, tagName) {
  const out = [];
  for (let i = 0; i < parent.childNodes.length; i += 1) {
    const c = parent.childNodes[i];
    if (isElement(c) && c.localName.toLowerCase() === tagName) out.push(c);
  }
  return out;
}

function collectDescendantsByTag(root, tagName) {
  const out = [];
  const stack = [root];
  while (stack.length > 0) {
    const n = stack.pop();
    if (!n) continue;
    if (isElement(n) && n.localName.toLowerCase() === tagName) {
      out.push(n);
      // Don't recurse into <tr> (no nested <tr> in OOXML)
      continue;
    }
    if (n.childNodes) {
      for (let i = n.childNodes.length - 1; i >= 0; i -= 1) stack.push(n.childNodes[i]);
    }
  }
  return out;
}

export default bodyHtmlToOoxml;
