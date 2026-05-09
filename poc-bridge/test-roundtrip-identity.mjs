// Day 1 verification harness for DocxBag.
//
// For each of the 20 HCRU MFA templates:
//   1. Load .docx into DocxBag
//   2. Extract body content
//   3. Replace body content with itself (identity)
//   4. Serialize back to .docx
//   5. Reload into a second DocxBag
//   6. Verify part-by-part:
//        a. Same set of parts in both bags (no part lost or added)
//        b. Every part except word/document.xml is byte-identical
//        c. word/document.xml shell (pre, post) is byte-identical
//        d. word/document.xml body content is byte-identical
//
// Pass criterion: all 20 templates pass all 4 checks.

import { readFileSync, readdirSync, writeFileSync, mkdirSync } from 'node:fs';
import { resolve as resolvePath, basename, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';
import { createHash } from 'node:crypto';

import { DocxBag } from './docx-bag.mjs';

const __dirname = dirname(fileURLToPath(import.meta.url));

const TEMPLATES_DIR =
  '/Users/mario002e/IdeaProjects/bsh-migration/Functional Specs/HCRU/New Templates/MFA';
const REPORT_OUT = `${__dirname}/out/day1-roundtrip-report.md`;

const sha256 = (buf) => createHash('sha256').update(buf).digest('hex');

const main = async () => {
  const templateNames = readdirSync(TEMPLATES_DIR)
    .filter((f) => f.toLowerCase().endsWith('.docx'))
    .sort();

  console.log(`Day 1 round-trip identity verification`);
  console.log(`Templates dir: ${TEMPLATES_DIR}`);
  console.log(`Templates found: ${templateNames.length}\n`);

  const results = [];

  for (const name of templateNames) {
    const path = resolvePath(TEMPLATES_DIR, name);
    const result = await checkTemplate(path);
    results.push({ name, ...result });
    const tag = result.allPass ? 'PASS' : 'FAIL';
    console.log(
      `[${tag}] ${name.padEnd(50)} parts=${result.partCount}, ` +
        `body=${result.bodyLen} chars, ` +
        (result.allPass ? 'all checks ok' : `failures: ${result.failures.join(', ')}`),
    );
  }

  const passes = results.filter((r) => r.allPass).length;
  const fails = results.length - passes;
  console.log(`\nSummary: ${passes}/${results.length} pass, ${fails} fail`);

  // Report file with full per-template details.
  mkdirSync(dirname(REPORT_OUT), { recursive: true });
  writeFileSync(REPORT_OUT, formatReport(results));
  console.log(`Report: ${REPORT_OUT}`);

  process.exit(fails > 0 ? 1 : 0);
};

const checkTemplate = async (path) => {
  let buffer, bag, body, out, outBag;
  const failures = [];

  try {
    buffer = readFileSync(path);
  } catch (e) {
    return { allPass: false, partCount: 0, bodyLen: 0, failures: [`read: ${e.message}`] };
  }

  try {
    bag = await DocxBag.fromBuffer(buffer);
  } catch (e) {
    return { allPass: false, partCount: 0, bodyLen: 0, failures: [`load: ${e.message}`] };
  }

  try {
    body = bag.getBodyContent();
    bag.replaceBodyContent(body); // identity round-trip
    out = await bag.serialize();
    outBag = await DocxBag.fromBuffer(out);
  } catch (e) {
    return {
      allPass: false,
      partCount: bag.listParts().length,
      bodyLen: body?.length ?? 0,
      failures: [`roundtrip: ${e.message}`],
    };
  }

  // Check (a): same set of parts.
  const aIn = new Set(bag.listParts());
  const aOut = new Set(outBag.listParts());
  if (aIn.size !== aOut.size || ![...aIn].every((p) => aOut.has(p))) {
    const lost = [...aIn].filter((p) => !aOut.has(p));
    const added = [...aOut].filter((p) => !aIn.has(p));
    failures.push(`parts mismatch (lost=${lost.join(',')}; added=${added.join(',')})`);
  }

  // Check (b): every non-document.xml part byte-identical.
  let nonDocCheckedCount = 0;
  let nonDocFailures = 0;
  for (const path of aIn) {
    if (path === 'word/document.xml') continue;
    if (!aOut.has(path)) continue;
    nonDocCheckedCount += 1;
    const a = bag.getPart(path);
    const b = outBag.getPart(path);
    if (!a || !b || sha256(a) !== sha256(b)) {
      nonDocFailures += 1;
    }
  }
  if (nonDocFailures > 0) {
    failures.push(
      `non-doc bytes differ in ${nonDocFailures}/${nonDocCheckedCount} parts`,
    );
  }

  // Check (c): document.xml shell preserved.
  const inShell = bag.getBodyShell();
  const outShell = outBag.getBodyShell();
  if (inShell.pre !== outShell.pre) {
    failures.push(`document.xml pre-shell differs (in=${inShell.pre.length} chars, out=${outShell.pre.length})`);
  }
  if (inShell.post !== outShell.post) {
    failures.push(
      `document.xml post-shell differs (in=${inShell.post.length} chars, out=${outShell.post.length})`,
    );
  }

  // Check (d): body content preserved.
  const inBody = bag.getBodyContent();
  const outBody = outBag.getBodyContent();
  if (inBody !== outBody) {
    failures.push(`body content differs (in=${inBody.length} chars, out=${outBody.length})`);
  }

  return {
    allPass: failures.length === 0,
    partCount: aIn.size,
    nonDocCheckedCount,
    bodyLen: inBody.length,
    failures,
  };
};

const formatReport = (results) => {
  const passes = results.filter((r) => r.allPass).length;
  const lines = [
    '# Day 1 — DocxBag round-trip identity report',
    '',
    `Templates checked: ${results.length}`,
    `Passes: ${passes}`,
    `Failures: ${results.length - passes}`,
    '',
    '## Per-template detail',
    '',
    '| Template | Parts | Non-doc parts checked | Body chars | Pass | Failures |',
    '|---|---:|---:|---:|---|---|',
  ];
  for (const r of results) {
    lines.push(
      `| ${r.name} | ${r.partCount} | ${r.nonDocCheckedCount ?? '-'} | ${r.bodyLen} | ${r.allPass ? '✓' : '✗'} | ${r.failures.length === 0 ? '' : r.failures.join('; ')} |`,
    );
  }
  return lines.join('\n') + '\n';
};

main().catch((e) => {
  console.error(e);
  process.exit(1);
});
