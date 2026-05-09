// Merge HTML-module list-numbering definitions into the host's
// word/numbering.xml.
//
// What we receive (per pending entry):
//   { numberingId, type: 'ol'|'ul', properties }
// where numberingId is the ID html-to-docx assigned (1-based, starting
// from 1 within a single fragment-render call).
//
// What we have to do
// ===================
// 1. Read the host's word/numbering.xml.
// 2. Find the highest existing abstractNumId and numId so we don't collide.
// 3. Assign each pending def a fresh (abstractNumId, numId) above that.
// 4. Build the <w:abstractNum>...</w:abstractNum> + <w:num>...</w:num>
//    XML for each, using a minimal-but-valid level structure.
// 5. Insert before </w:numbering>.
// 6. Return a mapping { oldNumId -> newNumId } so the caller can
//    remap numId references in word/document.xml.
//
// Notes
// -----
// - We generate 9 levels per abstractNum to match what html-to-docx
//   itself emits — Word complains if a level used by the doc isn't
//   declared.
// - We do NOT try to honour every CSS list-style-type — just decimal
//   for ol and a bullet for ul. The POC explicitly accepts this gap
//   (the paid module also has limited style mapping).

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

const DEFAULT_INDENT_TWIPS = 720;

const buildLevelXml = (ilvl, type) => {
  const numFmt = type === 'ol' ? 'decimal' : 'bullet';
  const lvlText = type === 'ol' ? `%${ilvl + 1}.` : '•'; // '•'
  const indent = (ilvl + 1) * DEFAULT_INDENT_TWIPS;
  const fontXml =
    type === 'ul'
      ? `<w:rPr><w:rFonts w:ascii="Symbol" w:hAnsi="Symbol" w:hint="default"/></w:rPr>`
      : '';
  return (
    `<w:lvl w:ilvl="${ilvl}">` +
    `<w:start w:val="1"/>` +
    `<w:numFmt w:val="${numFmt}"/>` +
    `<w:lvlText w:val="${lvlText}"/>` +
    `<w:lvlJc w:val="left"/>` +
    `<w:pPr>` +
    `<w:tabs><w:tab w:val="num" w:pos="${indent}"/></w:tabs>` +
    `<w:ind w:left="${indent}" w:hanging="360"/>` +
    `</w:pPr>` +
    fontXml +
    `</w:lvl>`
  );
};

const buildAbstractNumXml = (abstractNumId, type) => {
  const levels = Array.from({ length: 9 }, (_, i) => buildLevelXml(i, type)).join('');
  return `<w:abstractNum w:abstractNumId="${abstractNumId}">${levels}</w:abstractNum>`;
};

const buildNumXml = (numId, abstractNumId) =>
  `<w:num w:numId="${numId}">` +
  `<w:abstractNumId w:val="${abstractNumId}"/>` +
  `</w:num>`;

const NUM_ID_RE = /<w:num\s[^>]*w:numId="(\d+)"/g;
const ABS_NUM_ID_RE = /<w:abstractNum\s[^>]*w:abstractNumId="(\d+)"/g;

const findMaxId = (xml, regex) => {
  let max = 0;
  for (const m of xml.matchAll(regex)) {
    const id = parseInt(m[1], 10);
    if (id > max) max = id;
  }
  return max;
};

/**
 * @param {import('pizzip')} zip      live PizZip with the rendered docx
 * @param {Array<{callIdx: number, numberingId: number, type: 'ol'|'ul', properties: any}>} pendingNumbering
 * @returns {{ numIdRemap: Map<string, number>, addedAbstractNums: number, addedNums: number }}
 *   numIdRemap is keyed by `${callIdx}:${numberingId}` so concurrent
 *   fragment renders that all reuse numberingId=1 don't collide.
 */
export function mergeNumbering(zip, pendingNumbering) {
  const result = {
    numIdRemap: new Map(),
    addedAbstractNums: 0,
    addedNums: 0,
  };

  if (!pendingNumbering || pendingNumbering.length === 0) {
    return result;
  }

  const numberingPath = 'word/numbering.xml';
  let numberingXml = zip.file(numberingPath)?.asText();

  // If the template lacks numbering.xml entirely, create a minimal one.
  // (We do NOT register the part in [Content_Types].xml here — caller's
  // template should already have it. The fixture in poc-html-module/
  // does. For real ESCAP templates it will too. If a future template
  // doesn't, we'll fail loudly and the caller adds the override.)
  if (!numberingXml) {
    numberingXml =
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
      `<w:numbering xmlns:w="${W}"></w:numbering>`;
  }

  const maxAbstractId = findMaxId(numberingXml, ABS_NUM_ID_RE);
  const maxNumId = findMaxId(numberingXml, NUM_ID_RE);

  let nextAbstract = maxAbstractId + 1;
  let nextNum = maxNumId + 1;

  const abstractBlocks = [];
  const numBlocks = [];

  for (const def of pendingNumbering) {
    const newAbstractId = nextAbstract;
    const newNumId = nextNum;
    nextAbstract += 1;
    nextNum += 1;

    abstractBlocks.push(buildAbstractNumXml(newAbstractId, def.type));
    numBlocks.push(buildNumXml(newNumId, newAbstractId));
    // Composite key: callIdx isolates per-resolve numbering namespace.
    result.numIdRemap.set(`${def.callIdx}:${def.numberingId}`, newNumId);
    result.addedAbstractNums += 1;
    result.addedNums += 1;
  }

  // Per OOXML spec: <w:abstractNum> elements come BEFORE <w:num> elements
  // inside <w:numbering>. We insert all our abstract defs immediately
  // after the opening <w:numbering> (or its self-close), and our <w:num>
  // entries just before the closing </w:numbering>.
  const insertion = abstractBlocks.join('') + numBlocks.join('');

  let merged;
  if (/<w:numbering[^>]*\/>/.test(numberingXml)) {
    // Self-closed root — replace with proper open/close wrapping our content.
    merged = numberingXml.replace(
      /<w:numbering([^>]*)\/>/,
      `<w:numbering$1>${insertion}</w:numbering>`,
    );
  } else if (/<\/w:numbering>/.test(numberingXml)) {
    merged = numberingXml.replace(
      /<\/w:numbering>/,
      `${insertion}</w:numbering>`,
    );
  } else {
    // Malformed numbering.xml — bail with a no-op.
    return result;
  }

  zip.file(numberingPath, merged);
  return result;
}
