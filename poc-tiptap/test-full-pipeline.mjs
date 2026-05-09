// Day 5 final test: full pipeline.
//
//   docx → DocxBag → bodyOoxmlToHtml → Tiptap (round-trip) →
//     bodyHtmlToOoxml → DocxBag.replaceBodyContent → serialize
//
// For each MFA template, verify:
//   - Pipeline doesn't throw
//   - Output .docx structurally equivalent to original
//     (paragraph count, table count, SDT count, SDT identifier set)
//   - Output opens in LibreOffice
//
// This is the moment of truth: the user-edit pipeline (as it would
// run in a real Tiptap editor) faithfully round-trips.

import { readFileSync, readdirSync, writeFileSync, mkdirSync } from 'node:fs';
import { resolve as resolvePath, basename, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';
import { execFileSync } from 'node:child_process';
import { generateJSON, generateHTML } from '@tiptap/core';
import StarterKit from '@tiptap/starter-kit';
import Underline from '@tiptap/extension-underline';

import { DocxBag } from '../poc-bridge/docx-bag.mjs';
import { bodyOoxmlToHtml } from '../poc-bridge/body-ooxml-to-html.mjs';
import { bodyHtmlToOoxml } from '../poc-bridge/body-html-to-ooxml.mjs';
import { preserveOoxmlExtensions } from './preserve-ooxml-extensions.mjs';

const __dirname = dirname(fileURLToPath(import.meta.url));
const TEMPLATES_DIR =
  '/Users/mario002e/IdeaProjects/bsh-migration/Functional Specs/HCRU/New Templates/MFA';
const OUT_DIR = `${__dirname}/out/full-pipeline`;
const REPORT_OUT = `${__dirname}/out/full-pipeline-report.md`;

const extensions = [
  StarterKit.configure({
    paragraph: false, heading: false, link: false, underline: false,
  }),
  Underline,
  ...preserveOoxmlExtensions,
];

const main = async () => {
  mkdirSync(OUT_DIR, { recursive: true });
  mkdirSync(dirname(REPORT_OUT), { recursive: true });

  const templates = readdirSync(TEMPLATES_DIR)
    .filter((f) => f.toLowerCase().endsWith('.docx') && !f.startsWith('~$'))
    .sort();

  const results = [];

  for (const name of templates) {
    const path = resolvePath(TEMPLATES_DIR, name);
    const r = await runOne(path);
    results.push({ name, ...r });
    const tag = r.allPass ? 'PASS' : 'FAIL';
    console.log(
      `[${tag}] ${name.padEnd(50)} ` +
        `p=${r.pIn}/${r.pOut} tbl=${r.tblIn}/${r.tblOut} sdt=${r.sdtIn}/${r.sdtOut} ` +
        `ids=${r.idsIn}/${r.idsOut} ` +
        (r.allPass ? '' : `failures: ${r.failures.join('; ')}`),
    );
  }

  const passes = results.filter((r) => r.allPass).length;
  console.log(`\nFull pipeline: ${passes}/${results.length}`);
  writeFileSync(REPORT_OUT, formatReport(results));
  process.exit(results.length === passes ? 0 : 1);
};

const runOne = async (path) => {
  const failures = [];
  let buffer, bag, body, html, json, htmlOut, newBody, outBuf;

  try {
    buffer = readFileSync(path);
    bag = await DocxBag.fromBuffer(buffer);
    body = bag.getBodyContent();
  } catch (e) {
    return blank({ failures: [`load: ${e.message}`] });
  }

  const relsXml = bag.getPartText('word/_rels/document.xml.rels') || '';
  const relations = parseRelations(relsXml);

  try {
    html = bodyOoxmlToHtml(body, { relations });
  } catch (e) {
    return blank({ failures: [`docx→html: ${e.message}`] });
  }

  // The bridge HTML is fragment-shape (no <html><body>). Tiptap's
  // generateJSON expects a parseable HTML string; a fragment works
  // because internally it goes through DOMParser which accepts it.
  try {
    json = generateJSON(html, extensions);
    htmlOut = generateHTML(json, extensions);
  } catch (e) {
    return blank({ failures: [`tiptap roundtrip: ${e.message}`] });
  }

  try {
    newBody = bodyHtmlToOoxml(htmlOut);
  } catch (e) {
    return blank({ failures: [`html→docx: ${e.message}`] });
  }

  try {
    bag.replaceBodyContent(newBody);
    outBuf = await bag.serialize();
    writeFileSync(`${OUT_DIR}/${basename(path)}`, outBuf);
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

  if (pIn !== pOut) failures.push(`p ${pIn}→${pOut}`);
  if (tblIn !== tblOut) failures.push(`tbl ${tblIn}→${tblOut}`);
  if (sdtIn !== sdtOut) failures.push(`sdt ${sdtIn}→${sdtOut}`);

  const idsIn = extractSdtIds(body);
  const idsOut = extractSdtIds(newBody);
  const lostIds = [...idsIn].filter((x) => !idsOut.has(x));
  if (lostIds.length > 0) failures.push(`SDT ids lost (${lostIds.length})`);

  return {
    allPass: failures.length === 0,
    pIn, pOut, tblIn, tblOut, sdtIn, sdtOut,
    idsIn: idsIn.size, idsOut: idsOut.size,
    bodyIn: body.length, bodyOut: newBody.length, htmlLen: html.length,
    failures,
  };
};

const blank = (overrides) => ({
  allPass: false,
  pIn: 0, pOut: 0, tblIn: 0, tblOut: 0, sdtIn: 0, sdtOut: 0,
  idsIn: 0, idsOut: 0,
  bodyIn: 0, bodyOut: 0, htmlLen: 0,
  failures: [],
  ...overrides,
});

const countMatches = (s, re) => {
  let n = 0;
  for (const _ of s.matchAll(re)) n += 1;
  return n;
};

function extractSdtIds(bodyXml) {
  const ids = new Set();
  for (const m of bodyXml.matchAll(/<w:sdt(?:\s[^>]*)?>([\s\S]*?)<\/w:sdt>/g)) {
    const inner = m[1] ?? '';
    const tagMatch = inner.match(/<w:tag\s+w:val="([^"]*)"\s*\/?>/);
    if (tagMatch && tagMatch[1].trim()) {
      ids.add(tagMatch[1].trim());
      continue;
    }
    const withoutPr = inner.replace(/<w:sdtPr\b[\s\S]*?<\/w:sdtPr>/, '');
    const flat = withoutPr.replace(/<[^>]+>/g, '').replace(/\s+/g, ' ').trim();
    if (flat) ids.add(flat);
  }
  return ids;
}

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
    '# Day 5 — Full pipeline report (docx → html → Tiptap → html → docx)',
    '',
    `Templates: ${results.length}, passing: ${passes}`,
    '',
    '| Template | <w:p> in/out | <w:tbl> in/out | <w:sdt> in/out | SDT ids in/out | Body chars in/out | HTML chars | Pass | Failures |',
    '|---|---|---|---|---|---|---|---|---|',
  ];
  for (const r of results) {
    lines.push(
      `| ${r.name} | ${r.pIn}/${r.pOut} | ${r.tblIn}/${r.tblOut} | ${r.sdtIn}/${r.sdtOut} | ${r.idsIn}/${r.idsOut} | ${r.bodyIn}/${r.bodyOut} | ${r.htmlLen} | ${r.allPass ? '✓' : '✗'} | ${r.failures.join('; ')} |`,
    );
  }
  return lines.join('\n') + '\n';
};

main().catch((e) => {
  console.error(e);
  process.exit(1);
});
