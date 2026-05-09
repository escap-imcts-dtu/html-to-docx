// HtmlModule — open-source replacement for docxtemplater's paid HTML
// module. Recognises:
//   {~html}   block-level: replaces the enclosing <w:p>
//   {~~html}  inline-level: replaces the placeholder in-place inside its <w:p>
//
// Both pull an HTML string from the render scope (or from the
// `htmlValues` map passed to the module constructor), convert it to
// OOXML via `htmlToOoxmlFragment` from this fork, and splice the
// result into the host document.
//
// docxtemplater hooks used:
//   - matchers()  : claim '~' and '~~' prefixes, attaching `htmlMode`
//   - postparse() : for block tags, expand to enclosing <w:p>
//   - resolve()   : async — convert HTML to OOXML, stash on part
//   - render()    : sync — return the stashed OOXML as `value`
//
// Side effects (image media, list numbering defs, hyperlink rels)
// are accumulated during resolve() and exposed via takePending() so
// a postrender pass can commit them to the docx zip. (Mergers land
// in the next files: rels-merger.mjs, numbering-merger.mjs.)

// Use the built ESM bundle so we don't depend on Rollup-style
// extensionless imports inside src/. Run `npm run build` in the
// fork root before invoking the harness.
import { htmlToOoxmlFragment } from '../dist/html-to-docx.esm.js';

import { expandToParagraph } from './expand-paragraph.mjs';
import { extractInlineRuns } from './inline-extract.mjs';
import { mergeNumbering } from './numbering-merger.mjs';
import { mergeRelsAndMedia } from './rels-merger.mjs';

const MODULE_NAME = 'html';
const BLOCK_PREFIX = '~';
const INLINE_PREFIX = '~~';

const isHtmlPart = (part) =>
  part &&
  part.module === MODULE_NAME &&
  (part.htmlMode === 'block' || part.htmlMode === 'inline');

const emptyPending = () => ({ media: [], relationships: [], numbering: [] });

const debug = (...args) => {
  if (process.env.HTML_MODULE_DEBUG) {
    // eslint-disable-next-line no-console
    console.log('[HtmlModule]', ...args);
  }
};

export class HtmlModule {
  constructor(options = {}) {
    this.name = 'EscapHtmlModule';
    this.modules = [];
    this.options = { silentOnMissing: true, ...options };
    this.pending = emptyPending();
    // Per-iteration resolved cache. docxtemplater calls resolve()
    // once per (placeholder, scope) and provides a unique
    // `options.resolvedId`. Render() gets the same id, so we look
    // up here to avoid clobbering values when the same `part` object
    // is reused across loop iterations.
    this.resolvedMap = new Map();
    // For markers (numId / rId remap): each resolve call gets a
    // short numeric index so its numIds and rIds are namespaced.
    // Without this, multiple html-to-docx calls all generate numId 1,
    // numId 2, etc., and the merger collapses them to a single ID.
    this.callIndexFor = new Map(); // resolvedId → integer
    this.nextCallIndex = 0;
  }

  _callIndex(resolvedId) {
    let idx = this.callIndexFor.get(resolvedId);
    if (idx == null) {
      idx = this.nextCallIndex;
      this.nextCallIndex += 1;
      this.callIndexFor.set(resolvedId, idx);
    }
    return idx;
  }

  // -------------------------------------------------------------------
  // public helpers
  // -------------------------------------------------------------------

  /** Snapshot accumulated zip mutations and reset the buffer. */
  takePending() {
    const out = this.pending;
    this.pending = emptyPending();
    return out;
  }

  /**
   * Apply accumulated zip mutations to the docxtemplater output zip.
   * Call AFTER `await doc.renderAsync(...)` and BEFORE
   * `doc.getZip().generate(...)`.
   *
   * Performs three things:
   *   1. Merge our list-numbering definitions into word/numbering.xml,
   *      assigning fresh numIds.
   *   2. Merge our hyperlink/image relationships into
   *      word/_rels/document.xml.rels, assigning fresh rIds.
   *      Write image media into word/media/.
   *   3. Walk word/document.xml and replace our placeholder markers
   *      with the freshly-assigned numIds and rIds.
   */
  commit(zip) {
    const pending = this.takePending();

    const numResult = mergeNumbering(zip, pending.numbering);
    const relResult = mergeRelsAndMedia(zip, pending);

    // Final step: rewrite document.xml to (a) replace ID placeholders
    // and (b) ensure the <w:document> root declares all namespaces our
    // spliced content references (wp, a, pic, etc.). Word's parser
    // is strict about undeclared namespace prefixes; LibreOffice is
    // more forgiving but we should be valid for both.
    const docPath = 'word/document.xml';
    let docXml = zip.file(docPath)?.asText();
    if (docXml) {
      docXml = unmarkIds(docXml, numResult.numIdRemap, relResult.relIdRemap);
      docXml = ensureRequiredNamespaces(docXml);
      zip.file(docPath, docXml);
    }

    return {
      numbering: numResult,
      rels: relResult,
    };
  }

  // -------------------------------------------------------------------
  // docxtemplater module hooks
  // -------------------------------------------------------------------

  set(_options) {
    // Called by docxtemplater multiple times during parse/render. We
    // intentionally do NOT reset `pending` here — it would wipe work
    // accumulated during resolve() before postrender can read it.
  }

  matchers() {
    // Longer prefix first so '~~' wins on the priority tie-break.
    return [
      [INLINE_PREFIX, MODULE_NAME, { htmlMode: 'inline' }],
      [BLOCK_PREFIX, MODULE_NAME, { htmlMode: 'block' }],
    ];
  }

  postparse(postparsed, _modules, _options) {
    debug(
      'postparse in:',
      postparsed.map((p, i) => ({
        i,
        type: p.type,
        module: p.module,
        htmlMode: p.htmlMode,
        value: p.value,
        contentSnippet:
          p.type === 'content' ? String(p.value).slice(0, 60) : undefined,
      })),
    );

    const { parts, errors } = expandToParagraph(postparsed, {
      moduleName: MODULE_NAME,
      shouldExpand: (p) => isHtmlPart(p) && p.htmlMode === 'block',
    });

    debug(
      'postparse out:',
      parts.map((p, i) => ({
        i,
        type: p.type,
        module: p.module,
        htmlMode: p.htmlMode,
        value: p.value,
        expanded: p.expanded ? `<expanded ${p.expanded.length} parts>` : undefined,
      })),
      'errors:',
      errors,
    );

    return parts;
  }

  /**
   * Async pre-pass. docxtemplater awaits this during renderAsync()
   * before calling render(). We do the HTML→OOXML conversion here
   * and stash the result on the part so render() can read it sync.
   *
   * IMPORTANT: must return a falsy value (not a resolved Promise) for
   * parts we do not own. docxtemplater's resolve dispatcher does
   * `if (moduleResolved) return moduleResolved` — a Promise that
   * resolves to null is still truthy, which would short-circuit the
   * default scope resolution and turn `{recipient}` into `undefined`.
   */
  resolve(part, options) {
    if (!isHtmlPart(part)) return false;

    const key = options.resolvedId;
    if (key == null) {
      // Fallback: if docxtemplater didn't give us a resolvedId for
      // some reason, fall back to part identity. This works for
      // simple non-loop templates but breaks under loops.
      // Should never trigger in practice with current docxtemplater.
      debug('resolve: no resolvedId, falling back to part identity');
    }

    const html = this._resolveHtmlValue(part, options);
    if (html == null || html === '') {
      this.resolvedMap.set(key, emptyResolved());
      return Promise.resolve(null);
    }

    const callIdx = this._callIndex(key);

    return htmlToOoxmlFragment(html).then(
      (fragment) => {
        // Tag numId / r:id values inside the body XML with markers so
        // commit() can remap them to the host's ID space without
        // collisions across multiple fragment renders. Markers are
        // namespaced by callIdx so two fragments both using numId=1
        // get distinct final numIds.
        const markedBody = markIds(fragment.bodyXml, callIdx);
        const { inlineXml } = extractInlineRuns(markedBody);
        this.resolvedMap.set(key, {
          bodyXml: markedBody,
          inlineXml,
          fragment,
        });
        this._recordPending(fragment, callIdx);
        debug('resolve: ok for', part.value, 'callIdx=', callIdx, '(', markedBody.length, 'chars)');
        return null;
      },
      (cause) => {
        this.resolvedMap.set(key, {
          ...emptyResolved(),
          error: wrapError(
            `HTML→OOXML conversion failed for "${part.value}"`,
            cause,
          ),
        });
        return null;
      },
    );
  }

  render(part, options) {
    if (!isHtmlPart(part)) return null;

    const key = options.resolvedId;
    const resolved = this.resolvedMap.get(key);

    if (process.env.HTML_MODULE_DEBUG) {
      const snippet = resolved?.bodyXml?.slice(0, 80) ?? '<no body>';
      // eslint-disable-next-line no-console
      console.log(
        '[HtmlModule.render] part=' + part.value,
        'mode=' + part.htmlMode,
        'lIndex=' + part.lIndex,
        'key=' + key,
        'snippet=' + JSON.stringify(snippet),
      );
    }

    if (resolved && resolved.error) {
      return { value: '', errors: [resolved.error] };
    }

    if (!resolved) {
      if (this.options.silentOnMissing) return { value: '', errors: [] };
      return {
        value: '',
        errors: [
          new Error(
            `HTML resolve() did not run for "${part.value}" — call doc.renderAsync() instead of doc.render().`,
          ),
        ],
      };
    }

    if (resolved.bodyXml === '' && this.options.silentOnMissing) {
      return { value: '', errors: [] };
    }

    if (part.htmlMode === 'block') {
      return { value: resolved.bodyXml, errors: [] };
    }

    // Inline mode: the placeholder is sitting inside the host
    // paragraph's <w:r><w:t>...</w:t></w:r>. Inserting our <w:r>
    // children verbatim would nest them inside <w:t>, which is
    // invalid OOXML. Close the host's <w:t> and <w:r> before our
    // content, then re-open them after so any trailing text in the
    // same paragraph flows correctly.
    //
    // We always add `xml:space="preserve"` on the reopened <w:t>;
    // it's a safe over-broad default. (A tighter implementation
    // would inspect the original host <w:t>'s attributes during
    // postparse.)
    const wrapped =
      `</w:t></w:r>${resolved.inlineXml}<w:r><w:t xml:space="preserve">`;
    return { value: wrapped, errors: [] };
  }

  // -------------------------------------------------------------------
  // internals
  // -------------------------------------------------------------------

  _resolveHtmlValue(part, options) {
    if (
      this.options.htmlValues &&
      Object.prototype.hasOwnProperty.call(this.options.htmlValues, part.value)
    ) {
      const fromMap = this.options.htmlValues[part.value];
      if (typeof fromMap === 'string') return fromMap;
      if (fromMap == null) return null;
    }
    try {
      const v = options.scopeManager.getValue(part.value, { part });
      if (typeof v === 'string') return v;
      if (v == null) return null;
      if (typeof v === 'number' || typeof v === 'boolean') return String(v);
      return null;
    } catch {
      return null;
    }
  }

  _recordPending(fragment, callIdx) {
    for (const m of fragment.media) this.pending.media.push(m);
    for (const n of fragment.numbering) {
      this.pending.numbering.push({ ...n, callIdx });
    }
    for (const r of fragment.relationships) {
      this.pending.relationships.push({
        callIdx,
        fragmentRelId: r.relationshipId,
        finalRelId: '',
        type: r.type,
        target: r.target,
        targetMode: r.targetMode,
      });
    }
  }
}

const emptyResolved = () => ({
  bodyXml: '',
  inlineXml: '',
  fragment: { bodyXml: '', media: [], relationships: [], numbering: [] },
});

// -----------------------------------------------------------------
// Marker / unmarker for numId + r:id remap
// -----------------------------------------------------------------
//
// htmlToOoxmlFragment emits numIds (1, 2, 3, ...) and rIds (rId6, rId7, ...)
// that are valid only inside its own little world. Once spliced into
// the host document, those IDs may collide with the host's existing
// IDs. We tag them on the way out and replace with the merger's
// assigned IDs in commit().
//
// The marker syntax is intentionally chosen to be a valid XML
// attribute value (no XML-special chars) so docxtemplater treats it
// as opaque text.

const NUM_ID_VAL_RE = /(<w:numId\b[^>]*?\bw:val=")(\d+)(")/g;
// Hyperlinks use r:id, images use r:embed — both reference the same
// rId space in document.xml.rels.
const REL_ID_RE = /(\br:(?:id|embed)=")rId(\d+)(")/g;

// Markers are namespaced by callIdx so two fragments both using
// numId=1 don't collapse into one final numId after merge.
const NUM_MARKER = (callIdx, id) => `__hmod_${callIdx}_num_${id}__`;
const REL_MARKER = (callIdx, id) => `__hmod_${callIdx}_rel_${id}__`;
const NUM_MARKER_RE = /__hmod_(\d+)_num_(\d+)__/g;
const REL_MARKER_RE = /__hmod_(\d+)_rel_(\d+)__/g;

const composedKey = (callIdx, id) => `${callIdx}:${id}`;

function markIds(bodyXml, callIdx) {
  return bodyXml
    .replace(NUM_ID_VAL_RE, (_, pre, id, post) => `${pre}${NUM_MARKER(callIdx, id)}${post}`)
    .replace(REL_ID_RE, (_, pre, id, post) => `${pre}${REL_MARKER(callIdx, id)}${post}`);
}

function unmarkIds(docXml, numIdRemap, relIdRemap) {
  return docXml
    .replace(NUM_MARKER_RE, (_, callIdx, oldId) => {
      const mapped = numIdRemap.get(composedKey(callIdx, oldId));
      return mapped != null ? String(mapped) : oldId;
    })
    .replace(REL_MARKER_RE, (_, callIdx, oldId) => {
      const mapped = relIdRemap.get(composedKey(callIdx, oldId));
      return mapped != null ? mapped : `rId${oldId}`;
    });
}

// -----------------------------------------------------------------
// Namespace-declaration injection
// -----------------------------------------------------------------
//
// Our spliced content uses these prefixes that the host's
// <w:document> typically does not declare:
//   wp  - drawing inline (<wp:inline>, <wp:extent>, ...)
//   a   - drawingml main (<a:graphic>, <a:blip>, ...)
//   pic - drawingml picture (<pic:pic>, <pic:blipFill>, ...)
//   r   - officeDocument relationships (most templates DO declare it,
//         but we add it defensively so {~~html} with a hyperlink works
//         on hand-rolled templates)
//
// This is idempotent: if the namespace is already declared we skip it.

const REQUIRED_NAMESPACES = {
  wp: 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
  a: 'http://schemas.openxmlformats.org/drawingml/2006/main',
  pic: 'http://schemas.openxmlformats.org/drawingml/2006/picture',
  r: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
};

function ensureRequiredNamespaces(docXml) {
  return docXml.replace(/<w:document\b([^>]*)>/, (full, attrs) => {
    let updatedAttrs = attrs;
    for (const [prefix, uri] of Object.entries(REQUIRED_NAMESPACES)) {
      const declRe = new RegExp(`\\bxmlns:${prefix}\\s*=`);
      if (!declRe.test(updatedAttrs)) {
        updatedAttrs += ` xmlns:${prefix}="${uri}"`;
      }
    }
    return `<w:document${updatedAttrs}>`;
  });
}

const wrapError = (message, cause) => {
  const e = new Error(message);
  if (cause instanceof Error) {
    e.cause = cause;
    e.message = `${message}: ${cause.message}`;
  }
  return e;
};

export default HtmlModule;
