// Body OOXML → lossless HTML.
//
// Walks the children of <w:body> (sans <w:sectPr>) and emits HTML
// that:
//   1. Renders sensibly in a browser / WYSIWIG editor (semantic markup
//      where possible: <p>, <strong>, <em>, <table>, <a>, <img>).
//   2. Preserves enough OOXML metadata (via `data-ooxml-*` attributes
//      and base64 blob fallbacks) for a reverse converter to
//      reconstruct semantically-equivalent OOXML.
//
// Companion: body-html-to-ooxml.mjs (Day 3) consumes this output.
// Schema: ./DESIGN.md  → "data-ooxml-* attribute schema"
//
// Lossless strategy
// -----------------
// For every OOXML element we emit:
//   • Common, well-known properties (b, i, u, strike) → semantic HTML
//     tags (<strong>, <em>, <u>, <s>) for readability.
//   • Anything we cannot map cleanly is base64'd into an attribute on
//     the same element so the reverse pass can re-emit it verbatim:
//       <span data-ooxml-rpr="<base64>">…</span>     (run properties)
//       <p data-ooxml-ppr-extra="<base64>">…</p>     (paragraph props minus pStyle)
//       <table data-ooxml-tblpr="<base64>">…</table> (table properties)
//   • Elements with no HTML cousin (drawings/objects/fields beyond a
//     few we know about) are wrapped in a placeholder element carrying
//     the entire OOXML XML as base64:
//       <div class="ooxml-blob" data-ooxml-blob="<base64>"></div>
//     The reverse pass swaps these back at the same position.
//
// Whitespace
// ----------
// OOXML's <w:t> uses xml:space="preserve" to retain leading/trailing
// space. HTML normalizes whitespace by default. We preserve full
// content as the text node; the reverse converter restores the
// xml:space="preserve" attribute when needed (text starts/ends with
// whitespace).
//
// Why not extend mammoth?
// -----------------------
// Mammoth is excellent at OOXML→clean HTML for *display*, but its
// design intentionally drops Word-specific metadata (SDT identity, run
// properties beyond a small whitelist). For lossless round-trip we
// need to PRESERVE that metadata. A custom walker is simpler than
// adding a parallel "preserve metadata" mode to mammoth's transforms.

import { DOMParser, XMLSerializer } from '@xmldom/xmldom';

// -------------------------------------------------------------------
// XML / HTML escaping helpers
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

const b64 = (s) => Buffer.from(s, 'utf8').toString('base64');

// -------------------------------------------------------------------
// Public API
// -------------------------------------------------------------------

/**
 * Convert body XML (children of <w:body> minus <w:sectPr>) to HTML.
 *
 * @param {string} bodyXml
 * @param {object} [options]
 * @param {Map<string,string>} [options.relations]
 *   rId → target. Used to resolve hyperlinks (and image media paths).
 * @returns {string} HTML
 */
export function bodyOoxmlToHtml(bodyXml, options = {}) {
  if (typeof bodyXml !== 'string') {
    throw new TypeError('bodyOoxmlToHtml expects a string');
  }
  if (bodyXml.length === 0) return '';

  const relations = options.relations ?? new Map();

  // Parse as fragment: wrap with a synthetic root that declares the
  // common namespaces so the parser can resolve prefixes properly.
  const wrapped = wrapForParse(bodyXml);
  const dom = new DOMParser({
    onError: (level, msg) => {
      // Surface fatal errors; warnings are noisy on real Word output.
      if (level === 'fatalError' || level === 'error') {
        throw new Error(`bodyOoxmlToHtml: parse ${level}: ${msg.split('\n')[0]}`);
      }
    },
  }).parseFromString(wrapped, 'text/xml');

  const root = dom.documentElement;
  if (!root) throw new Error('bodyOoxmlToHtml: no root after parse');

  const out = [];
  for (let i = 0; i < root.childNodes.length; i += 1) {
    const child = root.childNodes[i];
    out.push(emitNode(child, { relations }));
  }
  return out.filter((s) => s !== '').join('');
}

// -------------------------------------------------------------------
// Wrapping: make the body XML parseable as a standalone fragment
// -------------------------------------------------------------------

// Namespaces commonly seen in real Word output. We pre-declare a
// generous set on the parse wrapper so xmldom doesn't choke on
// elements/attributes that use them. Anything not in this set will
// trigger a NamespaceError — caller should add to this map.
const KNOWN_NAMESPACES = {
  w: 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
  r: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
  w14: 'http://schemas.microsoft.com/office/word/2010/wordml',
  w15: 'http://schemas.microsoft.com/office/word/2012/wordml',
  w16: 'http://schemas.microsoft.com/office/word/2018/wordml',
  w16cex: 'http://schemas.microsoft.com/office/word/2018/wordml/cex',
  w16cid: 'http://schemas.microsoft.com/office/word/2016/wordml/cid',
  w16sdtdh: 'http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash',
  w16se: 'http://schemas.microsoft.com/office/word/2015/wordml/symex',
  wp: 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
  wp14: 'http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing',
  wpc: 'http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas',
  wpg: 'http://schemas.microsoft.com/office/word/2010/wordprocessingGroup',
  wpi: 'http://schemas.microsoft.com/office/word/2010/wordprocessingInk',
  wps: 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape',
  a: 'http://schemas.openxmlformats.org/drawingml/2006/main',
  a14: 'http://schemas.microsoft.com/office/drawing/2010/main',
  a16: 'http://schemas.microsoft.com/office/drawing/2014/main',
  pic: 'http://schemas.openxmlformats.org/drawingml/2006/picture',
  o: 'urn:schemas-microsoft-com:office:office',
  v: 'urn:schemas-microsoft-com:vml',
  m: 'http://schemas.openxmlformats.org/officeDocument/2006/math',
  mc: 'http://schemas.openxmlformats.org/markup-compatibility/2006',
  aink: 'http://schemas.microsoft.com/office/drawing/2016/ink',
  am3d: 'http://schemas.microsoft.com/office/drawing/2017/model3d',
  oel: 'http://schemas.microsoft.com/office/2019/extlst',
  cx: 'http://schemas.microsoft.com/office/drawing/2014/chartex',
  cx1: 'http://schemas.microsoft.com/office/drawing/2015/9/8/chartex',
  cx2: 'http://schemas.microsoft.com/office/drawing/2015/10/21/chartex',
  cx3: 'http://schemas.microsoft.com/office/drawing/2016/5/9/chartex',
  cx4: 'http://schemas.microsoft.com/office/drawing/2016/5/10/chartex',
  cx5: 'http://schemas.microsoft.com/office/drawing/2016/5/11/chartex',
  cx6: 'http://schemas.microsoft.com/office/drawing/2016/5/12/chartex',
  cx7: 'http://schemas.microsoft.com/office/drawing/2016/5/13/chartex',
  cx8: 'http://schemas.microsoft.com/office/drawing/2016/5/14/chartex',
  w10: 'urn:schemas-microsoft-com:office:word',
};

const PARSE_WRAPPER_OPEN =
  '<__bagroot ' +
  Object.entries(KNOWN_NAMESPACES)
    .map(([prefix, uri]) => `xmlns:${prefix}="${uri}"`)
    .join(' ') +
  '>';
const PARSE_WRAPPER_CLOSE = '</__bagroot>';

const wrapForParse = (xml) => PARSE_WRAPPER_OPEN + xml + PARSE_WRAPPER_CLOSE;

// -------------------------------------------------------------------
// xmldom doesn't preserve namespace prefixes the way we want, so we
// match by the LOCAL name and the namespace URI explicitly.
// -------------------------------------------------------------------

const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

const isElement = (n) => n && n.nodeType === 1;
const isText = (n) => n && n.nodeType === 3;

const isW = (n, localName) =>
  isElement(n) && n.namespaceURI === W_NS && n.localName === localName;

const HEADING_STYLE_RE = /^Heading([1-6])$/i;

// -------------------------------------------------------------------
// Element dispatch
// -------------------------------------------------------------------

function emitNode(node, ctx) {
  if (isText(node)) return emitText(node);
  if (!isElement(node)) return '';
  if (node.namespaceURI !== W_NS) {
    // Foreign element inside the body (mc:AlternateContent, etc.).
    // Treat as opaque blob.
    return emitBlob(node);
  }
  switch (node.localName) {
    case 'p':
      return emitParagraph(node, ctx);
    case 'r':
      return emitRun(node, ctx);
    case 't':
      return emitText(node.firstChild ?? null);
    case 'tab':
      return '<span class="ooxml-tab" data-ooxml-tab="">    </span>';
    case 'br':
      return emitBr(node);
    case 'sdt':
      return emitSdt(node, ctx);
    case 'tbl':
      return emitTable(node, ctx);
    case 'hyperlink':
      return emitHyperlink(node, ctx);
    case 'bookmarkStart':
    case 'bookmarkEnd':
    case 'commentRangeStart':
    case 'commentRangeEnd':
    case 'commentReference':
    case 'proofErr':
    case 'lastRenderedPageBreak':
      // Markers — preserve as zero-width markers for the reverse pass.
      return emitMarker(node);
    case 'fldSimple':
    case 'instrText':
    case 'fldChar':
      // Field codes — preserve verbatim as blob (display reconstructed
      // by the reverse pass; for now the editor sees no representation).
      return emitBlob(node);
    case 'drawing':
      return emitDrawing(node, ctx);
    case 'pict':
      return emitBlob(node);
    case 'smartTag':
    case 'customXml':
      // Deprecated Word semantic / customXml inline wrapper. Holds
      // inline runs as direct children. Transparent for display
      // (we emit just its inner runs) but the wrapper attributes are
      // preserved so we can rebuild it on the reverse trip.
      return emitTransparentInlineWrapper(node, ctx);
    case 'pPr':
    case 'rPr':
      // Should not be reached at top level — handled by parent.
      return '';
    default:
      // Unknown w:* element — preserve as blob.
      return emitBlob(node);
  }
}

// -------------------------------------------------------------------
// Specific emitters
// -------------------------------------------------------------------

function emitText(textNode) {
  if (!textNode) return '';
  return escapeText(textNode.nodeValue ?? '');
}

function emitBr(node) {
  const type = node.getAttributeNS(W_NS, 'type') || node.getAttribute('w:type') || '';
  // Page break is semantically distinct from line break.
  if (type === 'page') {
    return '<span class="ooxml-pagebreak" data-ooxml-br-type="page"></span>';
  }
  if (type === 'column') {
    return '<span class="ooxml-columnbreak" data-ooxml-br-type="column"></span>';
  }
  return '<br>';
}

function emitMarker(node) {
  const ser = new XMLSerializer().serializeToString(node);
  return `<span class="ooxml-marker" data-ooxml-blob="${escapeAttr(b64(ser))}"></span>`;
}

function emitBlob(node) {
  const ser = new XMLSerializer().serializeToString(node);
  // <span> not <div>: most opaque w:* elements (fldSimple, instrText,
  // fldChar, proofErr, pict, …) sit INSIDE a paragraph, so the
  // placeholder must be inline-valid. A span at body level is also
  // valid in our HTML schema (we only consume it via the bridge,
  // never via an HTML5 validator).
  return `<span class="ooxml-blob" data-ooxml-blob="${escapeAttr(b64(ser))}"></span>`;
}

function emitParagraph(p, ctx) {
  const pPr = childByLocal(p, 'pPr');
  const pStyle = pPr ? childByLocal(pPr, 'pStyle') : null;
  const pStyleVal = pStyle ? attrVal(pStyle, 'val') : '';
  const numPr = pPr ? childByLocal(pPr, 'numPr') : null;

  // Build inner content from <w:r>, <w:hyperlink>, <w:sdt>, etc.
  // Skip <w:pPr> — we already extracted what we need.
  const innerParts = [];
  for (let i = 0; i < p.childNodes.length; i += 1) {
    const c = p.childNodes[i];
    if (isW(c, 'pPr')) continue;
    innerParts.push(emitNode(c, ctx));
  }
  const inner = innerParts.join('');

  // Extract pPr attributes for round-trip. We strip <w:pStyle> (reflected
  // in the data-pstyle attr) and stash the remainder if non-empty.
  const pPrExtraAttr = pPr ? extractPPrExtra(pPr) : '';

  // Map heading styles to <h1>…<h6> for editor friendliness; everything
  // else stays as <p>.
  const heading = mapStyleToHeading(pStyleVal);

  const tag = heading ?? 'p';
  const attrs = [];
  if (pStyleVal && !heading) attrs.push(`data-ooxml-pstyle="${escapeAttr(pStyleVal)}"`);
  if (heading) attrs.push(`data-ooxml-pstyle="${escapeAttr(pStyleVal)}"`);
  if (pPrExtraAttr) attrs.push(`data-ooxml-ppr-extra="${escapeAttr(pPrExtraAttr)}"`);
  if (numPr) {
    const numId = childByLocal(numPr, 'numId');
    const ilvl = childByLocal(numPr, 'ilvl');
    if (numId) attrs.push(`data-ooxml-numid="${escapeAttr(attrVal(numId, 'val'))}"`);
    if (ilvl) attrs.push(`data-ooxml-ilvl="${escapeAttr(attrVal(ilvl, 'val'))}"`);
  }
  const attrStr = attrs.length > 0 ? ' ' + attrs.join(' ') : '';
  return `<${tag}${attrStr}>${inner}</${tag}>`;
}

function emitRun(r, ctx) {
  const rPr = childByLocal(r, 'rPr');

  // Build inner content from <w:t>, <w:tab>, <w:br>, <w:drawing>, etc.
  const innerParts = [];
  for (let i = 0; i < r.childNodes.length; i += 1) {
    const c = r.childNodes[i];
    if (isW(c, 'rPr')) continue;
    innerParts.push(emitNode(c, ctx));
  }
  let inner = innerParts.join('');

  if (!rPr) return inner; // run with no properties → just its content

  // Detect simple props that map to semantic HTML tags.
  const hasB = !!childByLocal(rPr, 'b');
  const hasI = !!childByLocal(rPr, 'i');
  const hasU = !!childByLocal(rPr, 'u');
  const hasStrike = !!childByLocal(rPr, 'strike');

  // Collect rPr children that are NOT in the "simple" set.
  const extraRPr = filterChildrenLocal(rPr, (c) => {
    return !['b', 'i', 'u', 'strike'].includes(c.localName);
  });

  // Wrap with semantic tags first (tightest-bound formatting innermost).
  if (hasStrike) inner = `<s>${inner}</s>`;
  if (hasU) inner = `<u>${inner}</u>`;
  if (hasI) inner = `<em>${inner}</em>`;
  if (hasB) inner = `<strong>${inner}</strong>`;

  if (extraRPr.length === 0) {
    return inner;
  }

  // Build a synthetic <w:rPr> containing only the "extra" children and
  // base64 it into data-ooxml-rpr.
  const rPrFragment = serializeAsRpr(extraRPr);
  return `<span data-ooxml-rpr="${escapeAttr(b64(rPrFragment))}">${inner}</span>`;
}

function emitSdt(sdt, ctx) {
  const sdtPr = childByLocal(sdt, 'sdtPr');
  const sdtContent = childByLocal(sdt, 'sdtContent');

  // Identifiers (any may be missing).
  const tagEl = sdtPr ? childByLocal(sdtPr, 'tag') : null;
  const aliasEl = sdtPr ? childByLocal(sdtPr, 'alias') : null;
  const idEl = sdtPr ? childByLocal(sdtPr, 'id') : null;
  const placeholderEl = sdtPr ? childByLocal(sdtPr, 'placeholder') : null;
  const docPartEl = placeholderEl ? childByLocal(placeholderEl, 'docPart') : null;

  const tagVal = tagEl ? attrVal(tagEl, 'val') : '';
  const aliasVal = aliasEl ? attrVal(aliasEl, 'val') : '';
  const idVal = idEl ? attrVal(idEl, 'val') : '';
  const docPartVal = docPartEl ? attrVal(docPartEl, 'val') : '';

  // Stash the FULL sdtPr as base64 — there are many sub-elements
  // (tabIndex, showingPlcHdr, lock, dataBinding, …) we can't enumerate
  // exhaustively.
  const sdtPrXml = sdtPr ? new XMLSerializer().serializeToString(sdtPr) : '';

  // Block-vs-inline detection: an SDT is block-level if its content
  // contains ANY block element ANYWHERE in its descendants — not just
  // as a direct child. Real templates nest <w:sdt><w:sdt><w:p>... so
  // the outer SDT must also be block-level even though its DIRECT
  // children are SDTs not paragraphs. Without this, an inline <span>
  // ends up wrapping a block <div>, which the editor's HTML schema
  // rejects (auto-corrected by closing the span early, which strands
  // trailing content and pads in extra paragraphs).
  const isBlock = sdtContent ? hasBlockDescendant(sdtContent) : false;

  // Inner content of sdtContent.
  let innerHtml = '';
  if (sdtContent) {
    const contentParts = [];
    for (let i = 0; i < sdtContent.childNodes.length; i += 1) {
      contentParts.push(emitNode(sdtContent.childNodes[i], ctx));
    }
    innerHtml = contentParts.join('');
  }

  const attrs = [];
  if (idVal) attrs.push(`data-ooxml-sdt-id="${escapeAttr(idVal)}"`);
  if (tagVal) attrs.push(`data-ooxml-sdt-tag="${escapeAttr(tagVal)}"`);
  if (aliasVal) attrs.push(`data-ooxml-sdt-alias="${escapeAttr(aliasVal)}"`);
  if (docPartVal) attrs.push(`data-ooxml-sdt-docpart="${escapeAttr(docPartVal)}"`);
  if (sdtPrXml) attrs.push(`data-ooxml-sdt-pr="${escapeAttr(b64(sdtPrXml))}"`);

  const cls = isBlock ? 'ooxml-sdt block' : 'ooxml-sdt';
  const tag = isBlock ? 'div' : 'span';
  return `<${tag} class="${cls}" ${attrs.join(' ')}>${innerHtml}</${tag}>`;
}

// True if `el` has any descendant that would render as a block-level
// HTML element (<p>, <table>, or another block-level <w:sdt>).
function hasBlockDescendant(el) {
  if (!el?.childNodes) return false;
  for (let i = 0; i < el.childNodes.length; i += 1) {
    const c = el.childNodes[i];
    if (!isElement(c)) continue;
    if (isW(c, 'p') || isW(c, 'tbl')) return true;
    if (isW(c, 'sdt')) {
      // Recurse into nested SDT's content.
      const inner = childByLocal(c, 'sdtContent');
      if (inner && hasBlockDescendant(inner)) return true;
    }
    // Recurse into other w:* containers (smartTag, customXml, …).
    if (hasBlockDescendant(c)) return true;
  }
  return false;
}

function emitTable(tbl, ctx) {
  const tblPr = childByLocal(tbl, 'tblPr');
  const tblGrid = childByLocal(tbl, 'tblGrid');
  const tblPrXml = tblPr ? new XMLSerializer().serializeToString(tblPr) : '';
  const tblGridXml = tblGrid ? new XMLSerializer().serializeToString(tblGrid) : '';

  const rows = [];
  for (let i = 0; i < tbl.childNodes.length; i += 1) {
    const c = tbl.childNodes[i];
    if (!isW(c, 'tr')) continue;
    rows.push(emitTableRow(c, ctx));
  }

  const attrs = [];
  if (tblPrXml) attrs.push(`data-ooxml-tbl-pr="${escapeAttr(b64(tblPrXml))}"`);
  if (tblGridXml) attrs.push(`data-ooxml-tbl-grid="${escapeAttr(b64(tblGridXml))}"`);
  return `<table${attrs.length ? ' ' + attrs.join(' ') : ''}><tbody>${rows.join('')}</tbody></table>`;
}

function emitTableRow(tr, ctx) {
  const trPr = childByLocal(tr, 'trPr');
  const trPrXml = trPr ? new XMLSerializer().serializeToString(trPr) : '';
  const cells = [];
  for (let i = 0; i < tr.childNodes.length; i += 1) {
    const c = tr.childNodes[i];
    if (!isW(c, 'tc')) continue;
    cells.push(emitTableCell(c, ctx));
  }
  const attrs = trPrXml
    ? ` data-ooxml-tr-pr="${escapeAttr(b64(trPrXml))}"`
    : '';
  return `<tr${attrs}>${cells.join('')}</tr>`;
}

function emitTableCell(tc, ctx) {
  const tcPr = childByLocal(tc, 'tcPr');
  const tcPrXml = tcPr ? new XMLSerializer().serializeToString(tcPr) : '';
  const inner = [];
  for (let i = 0; i < tc.childNodes.length; i += 1) {
    const c = tc.childNodes[i];
    if (isW(c, 'tcPr')) continue;
    inner.push(emitNode(c, ctx));
  }
  const attrs = tcPrXml
    ? ` data-ooxml-tc-pr="${escapeAttr(b64(tcPrXml))}"`
    : '';
  return `<td${attrs}>${inner.join('')}</td>`;
}

function emitHyperlink(link, ctx) {
  const rId = link.getAttributeNS(R_NS, 'id') || link.getAttribute('r:id') || '';
  const anchor = link.getAttributeNS(W_NS, 'anchor') || link.getAttribute('w:anchor') || '';
  const href = rId ? ctx.relations.get(rId) ?? '' : anchor ? `#${anchor}` : '';

  const inner = [];
  for (let i = 0; i < link.childNodes.length; i += 1) {
    inner.push(emitNode(link.childNodes[i], ctx));
  }
  const attrs = [];
  if (href) attrs.push(`href="${escapeAttr(href)}"`);
  if (rId) attrs.push(`data-ooxml-rid="${escapeAttr(rId)}"`);
  if (anchor) attrs.push(`data-ooxml-anchor="${escapeAttr(anchor)}"`);
  return `<a ${attrs.join(' ')}>${inner.join('')}</a>`;
}

function emitTransparentInlineWrapper(node, ctx) {
  // smartTag attributes live on the element itself: w:uri, w:element.
  // customXml uses the same shape.
  const uri = attrVal(node, 'uri');
  const element = attrVal(node, 'element');

  // Optional sub-element holding properties (smartTagPr / customXmlPr).
  // We base64 it as-is so the reverse pass can splice it back exactly.
  const prLocal = node.localName === 'smartTag' ? 'smartTagPr' : 'customXmlPr';
  const prEl = childByLocal(node, prLocal);
  const prB64 = prEl ? b64(new XMLSerializer().serializeToString(prEl)) : '';

  // Recurse into children, skipping the *Pr child (already captured).
  const innerParts = [];
  for (let i = 0; i < node.childNodes.length; i += 1) {
    const c = node.childNodes[i];
    if (isW(c, prLocal)) continue;
    innerParts.push(emitNode(c, ctx));
  }
  const inner = innerParts.join('');

  const isSmart = node.localName === 'smartTag';
  const cls = isSmart ? 'ooxml-smarttag' : 'ooxml-customxml';
  const prefix = isSmart ? 'st' : 'cx'; // short to keep attrs tidy
  const attrs = [`class="${cls}"`];
  if (uri) attrs.push(`data-ooxml-${prefix}-uri="${escapeAttr(uri)}"`);
  if (element) attrs.push(`data-ooxml-${prefix}-element="${escapeAttr(element)}"`);
  if (prB64) attrs.push(`data-ooxml-${prefix}-pr="${escapeAttr(prB64)}"`);
  return `<span ${attrs.join(' ')}>${inner}</span>`;
}

function emitDrawing(drawing, ctx) {
  // Try to find an embedded picture via descendant <a:blip r:embed>.
  const blip = findDescendantNS(
    drawing,
    'http://schemas.openxmlformats.org/drawingml/2006/main',
    'blip',
  );
  const embedRId = blip
    ? blip.getAttributeNS(R_NS, 'embed') || blip.getAttribute('r:embed') || ''
    : '';

  const fullDrawing = new XMLSerializer().serializeToString(drawing);
  if (embedRId) {
    // Render as an <img> for editor display, but stash the FULL drawing
    // OOXML so the reverse pass can splice it back exactly.
    const target = ctx.relations.get(embedRId) ?? '';
    const src = target ? `data:image/placeholder,rid=${embedRId}` : '';
    const attrs = [
      'class="ooxml-drawing"',
      `data-ooxml-embed-rid="${escapeAttr(embedRId)}"`,
      `data-ooxml-drawing="${escapeAttr(b64(fullDrawing))}"`,
      `alt=""`,
    ];
    if (src) attrs.push(`src="${escapeAttr(src)}"`);
    return `<img ${attrs.join(' ')}>`;
  }
  // Non-image drawing (shape, chart, etc.) — preserve as blob.
  return `<div class="ooxml-blob ooxml-drawing-blob" data-ooxml-blob="${escapeAttr(b64(fullDrawing))}"></div>`;
}

// -------------------------------------------------------------------
// Helpers — DOM walking, name lookup
// -------------------------------------------------------------------

function childByLocal(parent, localName) {
  for (let i = 0; i < parent.childNodes.length; i += 1) {
    const c = parent.childNodes[i];
    if (isElement(c) && c.namespaceURI === W_NS && c.localName === localName) {
      return c;
    }
  }
  return null;
}

function filterChildrenLocal(parent, pred) {
  const out = [];
  for (let i = 0; i < parent.childNodes.length; i += 1) {
    const c = parent.childNodes[i];
    if (isElement(c) && c.namespaceURI === W_NS && pred(c)) out.push(c);
  }
  return out;
}

function attrVal(el, localName) {
  // Word writes attributes in the w: namespace (e.g. <w:val>). xmldom
  // exposes them either via getAttributeNS(W_NS, 'val') or via the
  // qualified name 'w:val' depending on parser quirks; try both.
  return (
    el.getAttributeNS(W_NS, localName) ||
    el.getAttribute(`w:${localName}`) ||
    el.getAttribute(localName) ||
    ''
  );
}

function findDescendantNS(root, ns, localName) {
  const stack = [root];
  while (stack.length > 0) {
    const n = stack.pop();
    if (!n) continue;
    if (isElement(n) && n.namespaceURI === ns && n.localName === localName) return n;
    if (n.childNodes) {
      for (let i = 0; i < n.childNodes.length; i += 1) stack.push(n.childNodes[i]);
    }
  }
  return null;
}

function mapStyleToHeading(styleVal) {
  const m = styleVal.match(HEADING_STYLE_RE);
  return m ? `h${m[1]}` : null;
}

/** Serialize a <w:rPr> containing only the supplied children. */
function serializeAsRpr(children) {
  const ser = new XMLSerializer();
  const parts = children.map((c) => ser.serializeToString(c));
  return `<w:rPr xmlns:w="${W_NS}">${parts.join('')}</w:rPr>`;
}

/** Extract pPr minus pStyle and serialize the remainder if any children remain. */
function extractPPrExtra(pPr) {
  const remaining = filterChildrenLocal(pPr, (c) => c.localName !== 'pStyle');
  if (remaining.length === 0) return '';
  return b64(
    `<w:pPr xmlns:w="${W_NS}">${remaining.map((c) => new XMLSerializer().serializeToString(c)).join('')}</w:pPr>`,
  );
}

export default bodyOoxmlToHtml;
