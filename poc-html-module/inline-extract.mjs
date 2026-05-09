// For inline mode {~~html}: extract only the inline run-level OOXML
// from html-to-docx's paragraph-wrapped output.
//
// The fragment API always wraps content in <w:p>...</w:p> (one per
// HTML paragraph). For inline insertion we cannot keep the paragraph
// wrappers — they would create invalid OOXML inside the host's
// enclosing <w:p>. Instead we strip the outer <w:p> of each block
// and concatenate their inner content. Multiple HTML paragraphs
// collapse into a single line of runs separated by soft line breaks
// (<w:br/>).

const PARAGRAPH_BLOCK_RE = /<w:p(?:\s[^>]*)?>([\s\S]*?)<\/w:p>/g;
const TABLE_BLOCK_RE = /<w:tbl(?:\s[^>]*)?>[\s\S]*?<\/w:tbl>/g;
const PPR_BLOCK_RE = /<w:pPr(?:\s[^>]*)?>[\s\S]*?<\/w:pPr>/g;
const SOFT_BREAK = '<w:r><w:br/></w:r>';

/**
 * @param {string} bodyXml
 * @returns {{ inlineXml: string, droppedTables: number }}
 */
export function extractInlineRuns(bodyXml) {
  if (!bodyXml) return { inlineXml: '', droppedTables: 0 };

  const tableMatches = bodyXml.match(/<w:tbl[\s>]/g);
  const droppedTables = tableMatches ? tableMatches.length : 0;

  const noTables = bodyXml.replace(TABLE_BLOCK_RE, '');

  const innerParts = [];
  for (const m of noTables.matchAll(PARAGRAPH_BLOCK_RE)) {
    const inner = m[1] ?? '';
    const withoutPPr = inner.replace(PPR_BLOCK_RE, '');
    if (withoutPPr.trim().length > 0) innerParts.push(withoutPPr);
  }

  if (innerParts.length === 0) return { inlineXml: '', droppedTables };

  return { inlineXml: innerParts.join(SOFT_BREAK), droppedTables };
}
