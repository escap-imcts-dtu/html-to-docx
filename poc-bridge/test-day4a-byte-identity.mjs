// Day 4a — verify the lossless contract through the FULL HTML round-trip.
//
// Day 1 proved that an identity replacement (replace body with itself)
// preserves every non-body byte. Day 4a proves the same is true after
// the full docx→html→docx pipeline, which is the actual user flow.
//
// For each MFA template:
//   1. Load original via DocxBag
//   2. Load the corresponding round-tripped .docx (Day 3 output)
//   3. For each part path:
//        - if path !== 'word/document.xml': SHA-256 must match
//        - if path === 'word/document.xml': shell must match (pre/post)
//
// Pass criterion: every template upholds the byte-identity contract.

import { readFileSync, readdirSync } from 'node:fs';
import { resolve as resolvePath } from 'node:path';
import { createHash } from 'node:crypto';

import { DocxBag } from './docx-bag.mjs';

const TEMPLATES_DIR =
  '/Users/mario002e/IdeaProjects/bsh-migration/Functional Specs/HCRU/New Templates/MFA';
const ROUND_DIR =
  '/Users/mario002e/IdeaProjects/bsh-migration/html-to-docx/poc-bridge/out/day3-roundtrip';

const sha256 = (buf) => createHash('sha256').update(buf).digest('hex');

const main = async () => {
  const templates = readdirSync(TEMPLATES_DIR)
    .filter((f) => f.toLowerCase().endsWith('.docx'))
    .sort();

  let totalPass = 0;
  let totalFail = 0;
  const allFailures = [];

  for (const name of templates) {
    const origPath = resolvePath(TEMPLATES_DIR, name);
    const rtPath = resolvePath(ROUND_DIR, name);

    const origBuf = readFileSync(origPath);
    let rtBuf;
    try {
      rtBuf = readFileSync(rtPath);
    } catch {
      console.log(`[SKIP] ${name} — no round-trip output (was Day 3 skipped?)`);
      continue;
    }

    const origBag = await DocxBag.fromBuffer(origBuf);
    const rtBag = await DocxBag.fromBuffer(rtBuf);

    const origParts = new Set(origBag.listParts());
    const rtParts = new Set(rtBag.listParts());

    const failures = [];

    // Same set of parts.
    const lost = [...origParts].filter((p) => !rtParts.has(p));
    const added = [...rtParts].filter((p) => !origParts.has(p));
    if (lost.length > 0) failures.push(`lost parts: ${lost.join(', ')}`);
    if (added.length > 0) failures.push(`added parts: ${added.join(', ')}`);

    // Byte-identical for every non-document part.
    let nonDocChecked = 0;
    let nonDocFailed = 0;
    for (const path of origParts) {
      if (path === 'word/document.xml') continue;
      if (!rtParts.has(path)) continue;
      nonDocChecked += 1;
      const a = origBag.getPart(path);
      const b = rtBag.getPart(path);
      if (!a || !b || sha256(a) !== sha256(b)) {
        nonDocFailed += 1;
      }
    }
    if (nonDocFailed > 0) {
      failures.push(`${nonDocFailed}/${nonDocChecked} non-doc parts differ`);
    }

    // document.xml shell preserved.
    const origShell = origBag.getBodyShell();
    const rtShell = rtBag.getBodyShell();
    if (origShell.pre !== rtShell.pre) {
      failures.push(`document.xml shell-pre differs (${origShell.pre.length} → ${rtShell.pre.length} chars)`);
    }
    if (origShell.post !== rtShell.post) {
      failures.push(`document.xml shell-post differs (${origShell.post.length} → ${rtShell.post.length} chars)`);
    }

    if (failures.length === 0) {
      totalPass += 1;
      console.log(`[PASS] ${name.padEnd(50)} ${nonDocChecked} non-doc parts verified`);
    } else {
      totalFail += 1;
      allFailures.push({ name, failures });
      console.log(`[FAIL] ${name.padEnd(50)} ${failures.join('; ')}`);
    }
  }

  console.log(`\nByte-identity through full round-trip: ${totalPass}/${totalPass + totalFail}`);
  if (allFailures.length > 0) {
    console.log('\nFailure details:');
    for (const f of allFailures) {
      console.log(`  ${f.name}:`);
      for (const msg of f.failures) console.log(`    - ${msg}`);
    }
  }
  process.exit(totalFail === 0 ? 0 : 1);
};

main().catch((e) => {
  console.error(e);
  process.exit(1);
});
