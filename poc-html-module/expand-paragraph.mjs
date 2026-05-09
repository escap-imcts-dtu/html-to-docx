// Block-level paragraph expansion for {~html} placeholders.
//
// docxtemplater splits the input XML into discrete parts at every tag
// boundary. Each opening / closing XML tag is its own part of
// `type: 'tag'`, separate from text-content parts. So to find the
// enclosing <w:p>...</w:p>, we walk parts looking for a tag part whose
// value starts with `<w:p` (open) on the left and `</w:p>` on the
// right of our placeholder.
//
// Self-closing <w:p/> never opens a paragraph and is skipped.
//
// If no enclosing <w:p> is found (the template author put {~html}
// somewhere weird, e.g. inside a tableCell tag outside any paragraph),
// we report an error and leave the part untouched.

const PARAGRAPH_OPEN_RE = /^<w:p[\s>]/;
const PARAGRAPH_CLOSE_RE = /^<\/w:p>/;
const PARAGRAPH_SELF_CLOSE_RE = /^<w:p[^>]*\/>$/;

const isTagPart = (p) => p && p.type === 'tag';

const findOpeningParagraphIndex = (parts, startFrom) => {
  for (let i = startFrom; i >= 0; i -= 1) {
    const p = parts[i];
    if (!isTagPart(p)) continue;
    const v = String(p.value);
    if (PARAGRAPH_SELF_CLOSE_RE.test(v)) continue;
    if (PARAGRAPH_OPEN_RE.test(v)) return i;
  }
  return -1;
};

const findClosingParagraphIndex = (parts, startFrom) => {
  for (let i = startFrom; i < parts.length; i += 1) {
    const p = parts[i];
    if (!isTagPart(p)) continue;
    if (PARAGRAPH_CLOSE_RE.test(String(p.value))) return i;
  }
  return -1;
};

/**
 * @param {Array} parts        - docxtemplater postparsed array
 * @param {Object} options
 * @param {string} options.moduleName
 * @param {(p: any) => boolean} [options.shouldExpand]
 * @returns {{ parts: Array, errors: Array<{ message: string, id: string, part: any }> }}
 */
export function expandToParagraph(parts, options) {
  const errors = [];
  const out = [];

  let i = 0;
  while (i < parts.length) {
    const part = parts[i];
    if (!part) {
      i += 1;
      continue;
    }

    const isOurs =
      part.type === 'placeholder' &&
      part.module === options.moduleName &&
      (!options.shouldExpand || options.shouldExpand(part));

    if (!isOurs) {
      out.push(part);
      i += 1;
      continue;
    }

    const openIdx = findOpeningParagraphIndex(parts, i - 1);
    const closeIdx = findClosingParagraphIndex(parts, i + 1);

    if (openIdx < 0 || closeIdx < 0) {
      errors.push({
        message: `HTML tag "${part.value}" is not inside a <w:p> — block-level {~html} must be in a paragraph.`,
        id: 'html_block_not_in_paragraph',
        part,
      });
      out.push(part);
      i += 1;
      continue;
    }

    // Drop any opener-side parts already pushed to `out` — they belong
    // to the paragraph being absorbed into our part.
    while (out.length > openIdx) {
      out.pop();
    }

    const newPart = {
      ...part,
      expanded: parts.slice(openIdx, closeIdx + 1),
    };
    out.push(newPart);

    i = closeIdx + 1;
  }

  return { parts: out, errors };
}
