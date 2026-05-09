// DocxBag — sidecar that holds an entire .docx and exposes only the
// body of word/document.xml as the editable surface. Every other byte
// (every other zip entry, plus the shell of document.xml: declaration,
// root attrs, sectPr) survives a round-trip byte-for-byte.
//
// Design: ./DESIGN.md
//
// Public API (locked-in for Day 1):
//   const bag = await DocxBag.fromBuffer(docxBuffer);
//   const body = bag.getBodyContent();             // string (children of <w:body> minus <w:sectPr>)
//   bag.replaceBodyContent(newBodyXml);            // string in, no return
//   const shell = bag.getBodyShell();              // { pre, post }
//   const partNames = bag.listParts();             // string[]
//   const data = bag.getPart(path);                // Buffer | undefined
//   const text = bag.getPartText(path);            // string | undefined
//   const out = await bag.serialize();             // Buffer
//
// Implementation choices:
//   - Uses jszip (matches both forks: mammoth + html-to-docx).
//   - Body boundaries detected by a depth-aware token scanner over
//     word/document.xml — NOT by re-serializing through a DOM. This
//     guarantees the shell bytes are preserved verbatim.
//   - The last <w:sectPr> that is a direct child of <w:body> is the
//     boundary between body content and shell-post. Section-break
//     <w:sectPr> elements that live INSIDE a paragraph's <w:pPr> are
//     correctly classified as part of the body content.

import JSZip from 'jszip';

const DOCUMENT_PATH = 'word/document.xml';

export class DocxBag {
  /**
   * @param {JSZip} zip
   * @param {Map<string, Uint8Array>} parts  raw bytes per path
   * @param {string} documentXml             the in-memory document.xml as a string
   * @param {{ preEnd: number, postStart: number }} bounds
   *   Byte offsets within documentXml: body content is documentXml.slice(preEnd, postStart).
   */
  constructor(zip, parts, documentXml, bounds) {
    this._zip = zip;
    this._parts = parts;
    this._documentXml = documentXml;
    this._bounds = bounds;
  }

  /** Async factory. Reads the .docx and pre-computes body bounds. */
  static async fromBuffer(buffer) {
    if (!Buffer.isBuffer(buffer) && !(buffer instanceof Uint8Array)) {
      throw new TypeError('DocxBag.fromBuffer expects a Buffer or Uint8Array');
    }
    const zip = await JSZip.loadAsync(buffer);

    // Snapshot all parts as raw bytes. We hold these so non-document
    // parts can be rewritten verbatim on serialize() — no JSZip
    // re-encoding (which would change bytes via deflate variation).
    const parts = new Map();
    const entries = Object.entries(zip.files);
    for (const [path, file] of entries) {
      if (file.dir) continue;
      const data = await file.async('uint8array');
      parts.set(path, data);
    }

    const docBytes = parts.get(DOCUMENT_PATH);
    if (!docBytes) {
      throw new Error(`DocxBag: input is not a valid .docx — missing ${DOCUMENT_PATH}`);
    }
    const documentXml = bufToUtf8(docBytes);
    const bounds = findBodyBounds(documentXml);

    return new DocxBag(zip, parts, documentXml, bounds);
  }

  // -------------------------------------------------------------------
  // Body access / mutation
  // -------------------------------------------------------------------

  getBodyContent() {
    return this._documentXml.slice(this._bounds.preEnd, this._bounds.postStart);
  }

  getBodyShell() {
    return {
      pre: this._documentXml.slice(0, this._bounds.preEnd),
      post: this._documentXml.slice(this._bounds.postStart),
    };
  }

  replaceBodyContent(newBodyXml) {
    if (typeof newBodyXml !== 'string') {
      throw new TypeError('replaceBodyContent expects a string');
    }
    const { pre, post } = this.getBodyShell();
    const newDoc = pre + newBodyXml + post;
    this._documentXml = newDoc;
    // Recompute bounds in case sectPr position shifted (it can if the
    // user appended content after the sectPr's expected byte; in
    // practice it shouldn't, but be defensive).
    this._bounds = findBodyBounds(newDoc);
  }

  // -------------------------------------------------------------------
  // Shell parts (read-only)
  // -------------------------------------------------------------------

  listParts() {
    return [...this._parts.keys()].sort();
  }

  getPart(path) {
    const v = this._parts.get(path);
    return v ? Buffer.from(v) : undefined;
  }

  getPartText(path) {
    const v = this._parts.get(path);
    return v ? bufToUtf8(v) : undefined;
  }

  // -------------------------------------------------------------------
  // Output
  // -------------------------------------------------------------------

  /**
   * Serialize back to a .docx Buffer. Every part except
   * word/document.xml is written from its original bytes (no JSZip
   * re-encoding of the payload). word/document.xml is written from the
   * current in-memory string (which == original if no replaceBody was
   * called).
   *
   * Uses STORE compression for byte-determinism — same choice as
   * ESCAP's existing fillDocx (per ADR-016).
   */
  async serialize() {
    const out = new JSZip();
    const sortedPaths = [...this._parts.keys()].sort();
    for (const path of sortedPaths) {
      if (path === DOCUMENT_PATH) {
        out.file(path, this._documentXml);
      } else {
        out.file(path, this._parts.get(path));
      }
    }
    const buf = await out.generateAsync({
      type: 'nodebuffer',
      compression: 'STORE',
    });
    return buf;
  }
}

// -------------------------------------------------------------------
// Body bounds detection
// -------------------------------------------------------------------

/**
 * Scan word/document.xml and find:
 *   preEnd     — byte index just after the body's opening tag close
 *   postStart  — byte index of the body-level sectPr start (or
 *                of the body's closing tag if there is no body-level sectPr).
 *
 * The body-level sectPr is the last sectPr that is a *direct* child
 * of the body. Section-break sectPrs nested inside a paragraph's pPr
 * are correctly classified as body content (depth-aware scan).
 *
 * Implementation: walk every w:* tag using matchAll, track depth
 * relative to the body element. Record the position of any sectPr
 * seen at body's child level. Return the last such position; if none,
 * return the position of the closing body tag.
 *
 * This is a single linear pass; no DOM parse, no re-serialization.
 *
 * Throws if the body open or close are missing.
 */
function findBodyBounds(xml) {
  const bodyOpenRe = /<w:body(?:\s[^>]*)?>/;
  const bodyOpenMatch = bodyOpenRe.exec(xml);
  if (!bodyOpenMatch) throw new Error('DocxBag: w:body opening tag not found in document.xml');
  const preEnd = bodyOpenMatch.index + bodyOpenMatch[0].length;

  const bodyCloseIdx = xml.lastIndexOf('</w:body>');
  if (bodyCloseIdx < preEnd) {
    throw new Error('DocxBag: w:body closing tag missing or before opening tag');
  }

  // Tag scanner over the body region. We only care about w:* tags;
  // other namespaces (a:, pic:, wp:, etc.) live INSIDE a w-element
  // so they don't affect our depth tracking.
  const tagRe = /<(\/?)w:([A-Za-z0-9_]+)((?:\s[^>]*?)?)(\/?)>/g;
  const region = xml.slice(preEnd, bodyCloseIdx);

  let depth = 0; // depth relative to body's children
  let lastBodyLevelSectPrStart = -1;

  for (const m of region.matchAll(tagRe)) {
    const isClose = m[1] === '/';
    const isSelfClose = m[4] === '/';
    const tagName = m[2];

    if (isClose) {
      depth -= 1;
      continue;
    }

    // The element about to open sits at the current depth; its
    // children are at depth+1.
    if (depth === 0 && tagName === 'sectPr') {
      lastBodyLevelSectPrStart = preEnd + m.index;
    }

    if (!isSelfClose) {
      depth += 1;
    }
  }

  const postStart =
    lastBodyLevelSectPrStart >= 0 ? lastBodyLevelSectPrStart : bodyCloseIdx;
  return { preEnd, postStart };
}

// -------------------------------------------------------------------
// Helpers
// -------------------------------------------------------------------

function bufToUtf8(u8) {
  return Buffer.isBuffer(u8) ? u8.toString('utf8') : Buffer.from(u8).toString('utf8');
}

export default DocxBag;
