// Day 4b — visual fidelity diff via PDF.
//
// For each MFA template:
//   1. Convert original .docx → PDF (LibreOffice headless)
//   2. Convert round-tripped .docx → PDF (LibreOffice headless)
//   3. Extract text from each PDF (poppler's pdftotext)
//   4. Compare:
//        - normalized text equality
//        - page count (form-feeds in pdftotext output)
//        - text length delta
//
// Pass criterion: every template has identical normalized text and
// equal page count.

import { readdirSync, writeFileSync, mkdirSync, existsSync } from 'node:fs';
import { resolve as resolvePath, basename, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';
import { execFileSync } from 'node:child_process';

const __dirname = dirname(fileURLToPath(import.meta.url));
const TEMPLATES_DIR =
  '/Users/mario002e/IdeaProjects/bsh-migration/Functional Specs/HCRU/New Templates/MFA';
const ROUND_DIR = `${__dirname}/out/day3-roundtrip`;
const ORIG_PDF_DIR = `${__dirname}/out/day4-pdf-orig`;
const RT_PDF_DIR = `${__dirname}/out/day4-pdf-roundtrip`;
const REPORT_OUT = `${__dirname}/out/day4-fidelity-report.md`;

const main = async () => {
  mkdirSync(ORIG_PDF_DIR, { recursive: true });
  mkdirSync(RT_PDF_DIR, { recursive: true });
  mkdirSync(dirname(REPORT_OUT), { recursive: true });

  const templates = readdirSync(TEMPLATES_DIR)
    .filter((f) => f.toLowerCase().endsWith('.docx') && !f.startsWith('~$'))
    .sort();

  const results = [];

  for (const name of templates) {
    const origDocx = resolvePath(TEMPLATES_DIR, name);
    const rtDocx = resolvePath(ROUND_DIR, name);

    if (!existsSync(rtDocx)) {
      console.log(`[SKIP] ${name} — no round-trip docx`);
      continue;
    }

    const origPdf = `${ORIG_PDF_DIR}/${name.replace(/\.docx$/i, '.pdf')}`;
    const rtPdf = `${RT_PDF_DIR}/${name.replace(/\.docx$/i, '.pdf')}`;

    try {
      ensurePdf(origDocx, ORIG_PDF_DIR);
      ensurePdf(rtDocx, RT_PDF_DIR);
    } catch (e) {
      results.push({ name, allPass: false, failures: [`pdf convert: ${e.message}`] });
      console.log(`[FAIL] ${name} — pdf convert: ${e.message}`);
      continue;
    }

    const origText = pdfText(origPdf);
    const rtText = pdfText(rtPdf);
    const origPages = countPages(origPdf);
    const rtPages = countPages(rtPdf);

    const origNorm = normalize(origText);
    const rtNorm = normalize(rtText);

    const failures = [];
    if (origPages !== rtPages) {
      failures.push(`page count ${origPages} → ${rtPages}`);
    }
    if (origNorm !== rtNorm) {
      const charDelta = rtNorm.length - origNorm.length;
      const wordDelta = rtNorm.split(/\s+/).length - origNorm.split(/\s+/).length;
      const diffPos = firstDiffPos(origNorm, rtNorm);
      const winA = origNorm.slice(Math.max(0, diffPos - 40), diffPos + 80);
      const winB = rtNorm.slice(Math.max(0, diffPos - 40), diffPos + 80);
      failures.push(
        `text differs (Δ chars=${charDelta}, Δ words=${wordDelta}, ` +
          `first diff @${diffPos}: "${winA}" vs "${winB}")`,
      );
    }

    const allPass = failures.length === 0;
    results.push({
      name,
      allPass,
      origPages,
      rtPages,
      origChars: origNorm.length,
      rtChars: rtNorm.length,
      failures,
    });
    const tag = allPass ? 'PASS' : 'FAIL';
    console.log(
      `[${tag}] ${name.padEnd(50)} pages=${origPages}/${rtPages} chars=${origNorm.length}/${rtNorm.length} ` +
        (allPass ? '' : failures.join('; ')),
    );
  }

  writeFileSync(REPORT_OUT, formatReport(results));
  const passes = results.filter((r) => r.allPass).length;
  console.log(`\nVisual fidelity (PDF text+pages): ${passes}/${results.length}`);
  console.log(`Report: ${REPORT_OUT}`);
  process.exit(results.length === passes ? 0 : 1);
};

const ensurePdf = (docx, outDir) => {
  const expected = `${outDir}/${basename(docx).replace(/\.docx$/i, '.pdf')}`;
  if (existsSync(expected)) return;
  execFileSync(
    'soffice',
    ['--headless', '--convert-to', 'pdf', '--outdir', outDir, docx],
    { stdio: 'pipe', timeout: 60000 },
  );
  if (!existsSync(expected)) throw new Error(`PDF not produced at ${expected}`);
};

const pdfText = (pdf) =>
  execFileSync('pdftotext', ['-layout', '-nopgbrk', pdf, '-'], {
    encoding: 'utf8',
    timeout: 30000,
  });

const countPages = (pdf) => {
  const text = execFileSync('pdftotext', ['-layout', pdf, '-'], {
    encoding: 'utf8',
    timeout: 30000,
  });
  let n = 1;
  for (let i = 0; i < text.length; i += 1) if (text.charCodeAt(i) === 12) n += 1; // \f
  // If the PDF has zero text it might still be 1 page; clamp.
  return text.length === 0 ? 0 : n;
};

const normalize = (s) =>
  s
    .replace(/ /g, ' ') // nbsp → space
    .replace(/[ \t]+/g, ' ')
    .replace(/\r?\n[ \t]*/g, '\n')
    .replace(/\n{2,}/g, '\n')
    .trim();

const firstDiffPos = (a, b) => {
  const len = Math.min(a.length, b.length);
  for (let i = 0; i < len; i += 1) if (a.charCodeAt(i) !== b.charCodeAt(i)) return i;
  return len;
};

const formatReport = (results) => {
  const passes = results.filter((r) => r.allPass).length;
  const lines = [
    '# Day 4b — visual fidelity (PDF) report',
    '',
    `Templates: ${results.length}, passing: ${passes}`,
    '',
    '| Template | Pages orig/rt | Text chars orig/rt | Pass | Failures |',
    '|---|---|---|---|---|',
  ];
  for (const r of results) {
    lines.push(
      `| ${r.name} | ${r.origPages ?? '-'}/${r.rtPages ?? '-'} | ${r.origChars ?? '-'}/${r.rtChars ?? '-'} | ${r.allPass ? '✓' : '✗'} | ${r.failures.join('; ')} |`,
    );
  }
  return lines.join('\n') + '\n';
};

main().catch((e) => {
  console.error(e);
  process.exit(1);
});
