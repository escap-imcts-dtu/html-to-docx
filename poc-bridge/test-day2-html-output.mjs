// Day 2 verification: run all 20 MFA templates through DocxBag +
// bodyOoxmlToHtml. For each:
//   - Compare SDT count: source body has N <w:sdt> blocks → output
//     HTML must have exactly N elements with class="ooxml-sdt".
//   - Confirm no leftover <w: prefix in output (would indicate a node
//     type we didn't dispatch).
//   - Sanity check: no <w:p> / <w:r> survives in HTML.
//   - Save the HTML to ./out/html/<name>.html for visual inspection.
//
// Pass criterion: all 20 templates pass all checks.

import { readFileSync, readdirSync, writeFileSync, mkdirSync } from 'node:fs';
import { resolve as resolvePath, basename, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';

import { DocxBag } from './docx-bag.mjs';
import { bodyOoxmlToHtml } from './body-ooxml-to-html.mjs';

const __dirname = dirname(fileURLToPath(import.meta.url));
const TEMPLATES_DIR =
  '/Users/mario002e/IdeaProjects/bsh-migration/Functional Specs/HCRU/New Templates/MFA';
const HTML_OUT_DIR = `${__dirname}/out/html`;
const REPORT_OUT = `${__dirname}/out/day2-html-report.md`;

const main = async () => {
  const templates = readdirSync(TEMPLATES_DIR)
    .filter((f) => f.toLowerCase().endsWith('.docx'))
    .sort();

  mkdirSync(HTML_OUT_DIR, { recursive: true });
  mkdirSync(dirname(REPORT_OUT), { recursive: true });

  const results = [];

  for (const name of templates) {
    const path = resolvePath(TEMPLATES_DIR, name);
    const r = await runOne(path);
    results.push({ name, ...r });
    const tag = r.allPass ? 'PASS' : 'FAIL';
    console.log(
      `[${tag}] ${name.padEnd(50)} sdt=${r.srcSdtCount}/${r.htmlSdtCount} ` +
        `body=${r.bodyChars}c html=${r.htmlChars}c ` +
        (r.allPass ? '' : `failures: ${r.failures.join('; ')}`),
    );
  }

  writeFileSync(REPORT_OUT, formatReport(results));
  const passes = results.filter((r) => r.allPass).length;
  console.log(`\nSummary: ${passes}/${results.length} pass`);
  console.log(`Report: ${REPORT_OUT}`);
  console.log(`HTML files: ${HTML_OUT_DIR}/`);

  process.exit(results.length === passes ? 0 : 1);
};

const runOne = async (path) => {
  const failures = [];
  let buffer, bag, body, html;

  try {
    buffer = readFileSync(path);
    bag = await DocxBag.fromBuffer(buffer);
    body = bag.getBodyContent();
  } catch (e) {
    return {
      allPass: false,
      srcSdtCount: 0,
      htmlSdtCount: 0,
      bodyChars: 0,
      htmlChars: 0,
      failures: [`load: ${e.message}`],
    };
  }

  // Build relations map for hyperlink/image resolution.
  const relsXml = bag.getPartText('word/_rels/document.xml.rels') || '';
  const relations = parseRelations(relsXml);

  try {
    html = bodyOoxmlToHtml(body, { relations });
  } catch (e) {
    return {
      allPass: false,
      srcSdtCount: countMatches(body, /<w:sdt[\s>]/g),
      htmlSdtCount: 0,
      bodyChars: body.length,
      htmlChars: 0,
      failures: [`convert: ${e.message}`],
    };
  }

  const srcSdtCount = countMatches(body, /<w:sdt[\s>]/g);
  const htmlSdtCount = countMatches(html, /class="ooxml-sdt(?:\s|")/g);
  if (srcSdtCount !== htmlSdtCount) {
    failures.push(`sdt count ${srcSdtCount} → ${htmlSdtCount}`);
  }

  // No leftover w:* tag should remain in HTML (other than inside
  // base64-encoded blobs, which by definition are not tags).
  // Whitelist: nothing.
  const leftover = html.match(/<w:[A-Za-z]/g);
  if (leftover && leftover.length > 0) {
    failures.push(`leftover w:* tags: ${leftover.length}`);
  }

  // Save the HTML for inspection.
  const outPath = `${HTML_OUT_DIR}/${basename(path).replace(/\.docx$/i, '.html')}`;
  writeFileSync(outPath, htmlPreviewWrapper(basename(path), html));

  return {
    allPass: failures.length === 0,
    srcSdtCount,
    htmlSdtCount,
    bodyChars: body.length,
    htmlChars: html.length,
    failures,
  };
};

const countMatches = (s, re) => {
  let n = 0;
  // matchAll instead of String.match() so we get a count for /g-flagged regexes
  // without allocating the full array twice.
  // eslint-disable-next-line no-unused-vars
  for (const _ of s.matchAll(re)) n += 1;
  return n;
};

const parseRelations = (xml) => {
  const map = new Map();
  const re = /<Relationship\s+([^>]+?)\/>/g;
  for (const m of xml.matchAll(re)) {
    const attrs = m[1];
    const id = (attrs.match(/\bId="([^"]+)"/) || [])[1];
    const target = (attrs.match(/\bTarget="([^"]+)"/) || [])[1];
    if (id && target) map.set(id, target);
  }
  return map;
};

const htmlPreviewWrapper = (title, body) =>
  `<!DOCTYPE html>
<html><head>
<meta charset="utf-8">
<title>${title} (lossless preview)</title>
<style>
  body { font-family: Calibri, sans-serif; max-width: 900px; margin: 2em auto; padding: 0 1em; }
  table { border-collapse: collapse; margin: 0.5em 0; }
  td { border: 1px solid #999; padding: 0.3em 0.5em; }
  .ooxml-sdt { background: #fff8c4; padding: 0 2px; border-radius: 2px; }
  .ooxml-sdt.block { display: block; padding: 0.3em; margin: 0.3em 0; border-left: 3px solid #f3b800; }
  .ooxml-blob { color: #888; font-style: italic; font-size: 0.85em; }
  .ooxml-blob::before { content: "[opaque OOXML blob]"; }
  .ooxml-pagebreak::before { content: "—— page break ——"; display: block; text-align: center; color: #888; font-size: 0.85em; margin: 1em 0; }
  .ooxml-tab::before { content: "→ "; color: #ccc; }
  .ooxml-marker { display: none; }
</style>
</head><body>
<h1 style="border-bottom:1px solid #ccc">${title}</h1>
${body}
</body></html>`;

const formatReport = (results) => {
  const passes = results.filter((r) => r.allPass).length;
  const lines = [
    '# Day 2 — body OOXML → lossless HTML report',
    '',
    `Templates: ${results.length}, passing: ${passes}`,
    '',
    '| Template | Body chars | HTML chars | Source SDTs | Output SDTs | Pass | Failures |',
    '|---|---:|---:|---:|---:|---|---|',
  ];
  for (const r of results) {
    lines.push(
      `| ${r.name} | ${r.bodyChars} | ${r.htmlChars} | ${r.srcSdtCount} | ${r.htmlSdtCount} | ${r.allPass ? '✓' : '✗'} | ${r.failures.join('; ')} |`,
    );
  }
  return lines.join('\n') + '\n';
};

main().catch((e) => {
  console.error(e);
  process.exit(1);
});
