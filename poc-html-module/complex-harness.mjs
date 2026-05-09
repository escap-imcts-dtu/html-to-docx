// Complex harness — exercises the HtmlModule against an
// ESCAP-realistic fixture: SDT Content Controls, header/footer,
// applicants loop with rich-text notes per row, signature block,
// existing static table.
//
// Run:
//   node poc-html-module/build-complex-fixture.mjs
//   node poc-html-module/complex-harness.mjs
//
// The trickiest case here is `{~notesHtml}` (block) and `{~~statusHtml}`
// (inline) sitting INSIDE a `{#applicants}...{/applicants}` loop.
// docxtemplater will expand the loop first (duplicating the inner
// template parts per applicant), then apply postparse — meaning our
// expandToParagraph runs against an already-expanded parts array
// where each loop iteration's <w:p> wraps an instance of {~notesHtml}.
// If our paragraph-expansion logic is correct, each instance gets
// independently expanded.

import { readFileSync, writeFileSync } from 'node:fs';
import { dirname } from 'node:path';
import { fileURLToPath } from 'node:url';
import Docxtemplater from 'docxtemplater';
import PizZip from 'pizzip';

import { HtmlModule } from './html-module.mjs';

const __dirname = dirname(fileURLToPath(import.meta.url));
const FIXTURE = `${__dirname}/fixtures/complex-template.docx`;
const OUT = `${__dirname}/out/rendered-complex.docx`;

const applicants = [
  {
    name: 'Aisha Khan',
    ref: 'PTS-2026-0142',
    notesHtml: `
      <p>Application reviewed; <strong>passport scan accepted</strong>. Outstanding items:</p>
      <ul>
        <li>Police clearance (expected by 2026-06-15)</li>
        <li>Updated medical certificate</li>
      </ul>
      <p>Decision is conditional pending the above.</p>
    `,
    statusHtml: `<strong>Conditionally approved</strong>`,
  },
  {
    name: 'Diego Ramírez',
    ref: 'PTS-2026-0143',
    notesHtml: `
      <p>Submission complete. Reviewer comments:</p>
      <ol>
        <li>All documentation verified.</li>
        <li>No further action required.</li>
      </ol>
      <p>See <a href="https://example.org/rules">policy rules</a> for renewal cadence.</p>
    `,
    statusHtml: `<strong style="color: #1a7f37">Approved</strong>`,
  },
  {
    name: 'Mei-Ling Chen',
    ref: 'PTS-2026-0144',
    notesHtml: `<p>Application <em>incomplete</em>. Missing fields highlighted in red on the case file.</p>`,
    statusHtml: `<em>Pending</em>`,
  },
];

const main = async () => {
  const templateBuf = readFileSync(FIXTURE);
  const zip = new PizZip(templateBuf);

  // Build the htmlValues map by walking applicants. The placeholder
  // names inside the loop are 'notesHtml' and 'statusHtml' but
  // docxtemplater dispatches them per-iteration via scope. We don't
  // need an htmlValues map here — letting the scope manager resolve
  // is the realistic path.
  const htmlModule = new HtmlModule({});

  const doc = new Docxtemplater(zip, {
    paragraphLoop: true,
    linebreaks: true,
    modules: [htmlModule],
  });

  await doc.renderAsync({
    recipient: 'Madam Director Singh',
    caseRef: 'CASE-2026-0421',
    applicants,
    signerName: 'Dr. Helena Mwangi',
    signerDate: '2026-05-09',
  });

  const commitResult = htmlModule.commit(doc.getZip());
  console.log(
    `commit: numbering ${commitResult.numbering.addedAbstractNums} abstractNums + ${commitResult.numbering.addedNums} nums; ` +
      `rels ${commitResult.rels.addedRels}; media ${commitResult.rels.addedMedia}`,
  );

  const out = doc.getZip().generate({ type: 'nodebuffer', compression: 'STORE' });
  writeFileSync(OUT, out);
  console.log(`Wrote ${OUT} (${out.length} bytes)`);

  const outZip = new PizZip(out);
  const docXml = outZip.file('word/document.xml')?.asText() ?? '';

  const checks = [
    ['valid PK header', out[0] === 0x50 && out[1] === 0x4b],
    ['recipient resolved (in SDT)', docXml.includes('Madam Director Singh')],
    ['caseRef resolved (in SDT)', docXml.includes('CASE-2026-0421')],
    ['signerName resolved (in SDT)', docXml.includes('Dr. Helena Mwangi')],
    ['signerDate resolved (in SDT)', docXml.includes('2026-05-09')],
    ['applicant 1 name', docXml.includes('Aisha Khan')],
    ['applicant 2 name', docXml.includes('Diego Ramírez')],
    ['applicant 3 name', docXml.includes('Mei-Ling Chen')],
    ['applicant 1 ref', docXml.includes('PTS-2026-0142')],
    ['applicant 2 ref', docXml.includes('PTS-2026-0143')],
    ['applicant 3 ref', docXml.includes('PTS-2026-0144')],
    ['applicant 1 rich notes content (passport scan)', docXml.includes('passport scan accepted')],
    ['applicant 2 rich notes content (verified)', docXml.includes('All documentation verified.')],
    ['applicant 3 rich notes content (incomplete)', docXml.includes('incomplete')],
    ['applicant 2 hyperlink in notes (policy rules)', /policy rules/.test(docXml)],
    ['applicant 1 list rendered (numId)', /<w:numId/.test(docXml)],
    ['applicant 1 status inline (Conditionally approved)', docXml.includes('Conditionally approved')],
    ['applicant 2 status inline (Approved)', docXml.includes('Approved')],
    ['applicant 3 status inline (Pending)', docXml.includes('Pending')],
    ['static table preserved (Code header)', docXml.includes('>Code<')],
    ['static table preserved (Approved row)', docXml.includes('>Approved<')],
    ['no leftover {~notesHtml}', !docXml.includes('{~notesHtml}')],
    ['no leftover {~~statusHtml}', !docXml.includes('{~~statusHtml}')],
    ['no leftover loop tags', !docXml.includes('{#applicants}') && !docXml.includes('{/applicants}')],
    ['no unresolved __hmod_ markers', !/__hmod_/.test(docXml)],
    ['header preserved (UN ESCAP)', outZip.file('word/header1.xml')?.asText().includes('UN ESCAP')],
    ['footer preserved (PAGE field)', outZip.file('word/footer1.xml')?.asText().includes('PAGE')],
  ];

  console.log('\n=== assertions ===');
  let failures = 0;
  for (const [label, ok] of checks) {
    console.log((ok ? 'PASS  ' : 'FAIL  ') + label);
    if (!ok) failures += 1;
  }

  if (failures > 0) {
    console.error(`\n${failures} assertion(s) failed`);
    // Dump some output for debugging
    const aishaPos = docXml.indexOf('Aisha Khan');
    if (aishaPos > 0) {
      console.log('\n=== document.xml around applicant 1 ===');
      console.log(docXml.slice(Math.max(0, aishaPos - 100), aishaPos + 1500));
    }
    process.exit(1);
  }
  console.log('\nAll complex-harness checks passed.');
  console.log(`\nOpen ${OUT} in Word/LibreOffice to verify visual fidelity.`);
};

main().catch((e) => {
  console.error(e);
  process.exit(1);
});
