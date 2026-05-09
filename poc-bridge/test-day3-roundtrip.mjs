// Day 3 verification: full HTML round-trip for all 20 MFA templates.
//
// For each template:
//   1. Load via DocxBag
//   2. Convert body OOXML → HTML (bodyOoxmlToHtml)
//   3. Convert HTML → body OOXML (bodyHtmlToOoxml)
//   4. Replace body in DocxBag with the round-tripped OOXML
//   5. Serialize and reload to verify openability
//
// Semantic equivalence checks (byte-equality is too strict because
// xmldom may re-order attrs / change whitespace; we check the
// invariants that matter):
//   a. Same count of <w:p> in input body vs round-tripped body
//   b. Same count of <w:tbl>
//   c. Same count of <w:sdt>
//   d. Same SET of SDT identifiers (extracted from <w:tag w:val>)
//      OR inner-text identifiers when <w:tag> is missing — see
//      extractSdtIds() below for the rule we use, which mirrors
//      ESCAP's import.ts extractSdtIdentifier().
//   e. The output zip opens via JSZip (basic structural validity)
//
// Pass criterion: every template passes a-d, and at least 18/20 pass e
// (the harness keeps going if a single template trips structural
// equivalence so we get a full report, not a fail-fast).

import { readFileSync, readdirSync, writeFileSync, mkdirSync } from 'node:fs';
import { resolve as resolvePath, basename, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';
import JSZip from 'jszip';

import { DocxBag } from './docx-bag.mjs';
import { bodyOoxmlToHtml } from './body-ooxml-to-html.mjs';
import { bodyHtmlToOoxml } from './body-html-to-ooxml.mjs';

const __dirname = dirname(fileURLToPath(import.meta.url));
const TEMPLATES_DIR =
  '/Users/mario002e/IdeaProjects/bsh-migration/Functional Specs/HCRU/New Templates/MFA';
const ROUND_OUT_DIR = `${__dirname}/out/day3-roundtrip`;
const REPORT_OUT = `${__dirname}/out/day3-roundtrip-report.md`;

const main = async () => {
  const templates = readdirSync(TEMPLATES_DIR)
    .filter((f) => f.toLowerCase().endsWith('.docx') && !f.startsWith('~$'))
    .sort();

  mkdirSync(ROUND_OUT_DIR, { recursive: true });
  mkdirSync(dirname(REPORT_OUT), { recursive: true });

  const results = [];

  for (const name of templates) {
    const path = resolvePath(TEMPLATES_DIR, name);
    const r = await runOne(path);
    results.push({ name, ...r });
    const tag = r.allPass ? 'PASS' : 'FAIL';
    console.log(
      `[${tag}] ${name.padEnd(50)} ` +
        `p=${r.pCountIn}/${r.pCountOut} ` +
        `tbl=${r.tblCountIn}/${r.tblCountOut} ` +
        `sdt=${r.sdtCountIn}/${r.sdtCountOut} ` +
        `body=${r.bodyInLen}→${r.bodyOutLen}c ` +
        (r.allPass ? '' : `failures: ${r.failures.join('; ')}`),
    );
  }

  writeFileSync(REPORT_OUT, formatReport(results));
  const passes = results.filter((r) => r.allPass).length;
  console.log(`\nSummary: ${passes}/${results.length} pass`);
  console.log(`Report: ${REPORT_OUT}`);
  console.log(`Round-tripped .docx files: ${ROUND_OUT_DIR}/`);
  process.exit(results.length === passes ? 0 : 1);
};

const runOne = async (path) => {
  const failures = [];
  let buffer, bag, body, html, newBody, outBuffer, outZip;

  try {
    buffer = readFileSync(path);
    bag = await DocxBag.fromBuffer(buffer);
    body = bag.getBodyContent();
  } catch (e) {
    return blankResult({ failures: [`load: ${e.message}`] });
  }

  const relsXml = bag.getPartText('word/_rels/document.xml.rels') || '';
  const relations = parseRelations(relsXml);

  try {
    html = bodyOoxmlToHtml(body, { relations });
  } catch (e) {
    return blankResult({ failures: [`docx→html: ${e.message}`] });
  }

  try {
    newBody = bodyHtmlToOoxml(html);
  } catch (e) {
    return blankResult({ failures: [`html→docx: ${e.message}`] });
  }

  try {
    bag.replaceBodyContent(newBody);
    outBuffer = await bag.serialize();
    writeFileSync(`${ROUND_OUT_DIR}/${basename(path)}`, outBuffer);
    outZip = await JSZip.loadAsync(outBuffer);
    if (!outZip.file('word/document.xml')) failures.push('output missing word/document.xml');
  } catch (e) {
    failures.push(`serialize: ${e.message}`);
  }

  // Counts
  const pIn = countMatches(body, /<w:p[\s\/>]/g);
  const pOut = countMatches(newBody, /<w:p[\s\/>]/g);
  const tblIn = countMatches(body, /<w:tbl[\s>]/g);
  const tblOut = countMatches(newBody, /<w:tbl[\s>]/g);
  const sdtIn = countMatches(body, /<w:sdt[\s>]/g);
  const sdtOut = countMatches(newBody, /<w:sdt[\s>]/g);

  if (pIn !== pOut) failures.push(`paragraph count: ${pIn} → ${pOut}`);
  if (tblIn !== tblOut) failures.push(`table count: ${tblIn} → ${tblOut}`);
  if (sdtIn !== sdtOut) failures.push(`sdt count: ${sdtIn} → ${sdtOut}`);

  // SDT identifiers
  const idsIn = extractSdtIds(body);
  const idsOut = extractSdtIds(newBody);
  const lostIds = [...idsIn].filter((x) => !idsOut.has(x));
  const addedIds = [...idsOut].filter((x) => !idsIn.has(x));
  if (lostIds.length > 0) failures.push(`SDT ids lost (${lostIds.length}): ${lostIds.slice(0, 3).join(',')}…`);
  if (addedIds.length > 0) failures.push(`SDT ids added (${addedIds.length}): ${addedIds.slice(0, 3).join(',')}…`);

  return {
    allPass: failures.length === 0,
    pCountIn: pIn,
    pCountOut: pOut,
    tblCountIn: tblIn,
    tblCountOut: tblOut,
    sdtCountIn: sdtIn,
    sdtCountOut: sdtOut,
    bodyInLen: body.length,
    bodyOutLen: newBody.length,
    sdtIdsIn: idsIn.size,
    sdtIdsOut: idsOut.size,
    failures,
  };
};

const blankResult = (overrides) => ({
  allPass: false,
  pCountIn: 0, pCountOut: 0,
  tblCountIn: 0, tblCountOut: 0,
  sdtCountIn: 0, sdtCountOut: 0,
  bodyInLen: 0, bodyOutLen: 0,
  sdtIdsIn: 0, sdtIdsOut: 0,
  failures: [],
  ...overrides,
});

const countMatches = (s, re) => {
  let n = 0;
  for (const _ of s.matchAll(re)) n += 1;
  return n;
};

/**
 * Extract SDT identifiers from a body XML chunk. Mirrors the rule
 * from ESCAP's import.ts extractSdtIdentifier(): prefer <w:tag w:val>,
 * fall back to the inner display text of <w:sdtContent>.
 *
 * For round-trip purposes we just want a SET of identifiers per body;
 * exact preservation is the invariant.
 */
function extractSdtIds(bodyXml) {
  const ids = new Set();
  const sdtRe = /<w:sdt(?:\s[^>]*)?>([\s\S]*?)<\/w:sdt>/g;
  for (const m of bodyXml.matchAll(sdtRe)) {
    const inner = m[1] ?? '';
    // 1) explicit tag
    const tagMatch = inner.match(/<w:tag\s+w:val="([^"]*)"\s*\/?>/);
    if (tagMatch && tagMatch[1].trim()) {
      ids.add(tagMatch[1].trim());
      continue;
    }
    // 2) inner display text (strip <w:sdtPr> first)
    const withoutPr = inner.replace(/<w:sdtPr\b[\s\S]*?<\/w:sdtPr>/, '');
    const flat = stripTags(withoutPr).replace(/\s+/g, ' ').trim();
    if (flat) ids.add(flat);
  }
  return ids;
}

const stripTags = (xml) => xml.replace(/<[^>]+>/g, '');

const parseRelations = (xml) => {
  const map = new Map();
  for (const m of xml.matchAll(/<Relationship\s+([^>]+?)\/>/g)) {
    const attrs = m[1];
    const id = (attrs.match(/\bId="([^"]+)"/) || [])[1];
    const target = (attrs.match(/\bTarget="([^"]+)"/) || [])[1];
    if (id && target) map.set(id, target);
  }
  return map;
};

const formatReport = (results) => {
  const passes = results.filter((r) => r.allPass).length;
  const lines = [
    '# Day 3 — HTML round-trip report',
    '',
    `Templates: ${results.length}, passing: ${passes}`,
    '',
    '| Template | <w:p> in/out | <w:tbl> in/out | <w:sdt> in/out | SDT ids in/out | Body chars in/out | Pass | Failures |',
    '|---|---|---|---|---|---|---|---|',
  ];
  for (const r of results) {
    lines.push(
      `| ${r.name} | ${r.pCountIn}/${r.pCountOut} | ${r.tblCountIn}/${r.tblCountOut} | ${r.sdtCountIn}/${r.sdtCountOut} | ${r.sdtIdsIn}/${r.sdtIdsOut} | ${r.bodyInLen}/${r.bodyOutLen} | ${r.allPass ? '✓' : '✗'} | ${r.failures.join('; ')} |`,
    );
  }
  return lines.join('\n') + '\n';
};

main().catch((e) => {
  console.error(e);
  process.exit(1);
});
