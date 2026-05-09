// POC harness — loads the fixture template, attaches HtmlModule, and
// renders against rich HTML inputs. Confirms placeholder swap works
// end-to-end and dumps assertions + a snippet of the rendered XML.
//
// Run:
//   node poc-html-module/build-fixture.mjs   # if fixture is missing
//   node poc-html-module/harness.mjs
//
// Debug:
//   HTML_MODULE_DEBUG=1 node poc-html-module/harness.mjs
//
// What this DOES verify (post-resolve render path):
//   - {~html} block tags get the surrounding <w:p> replaced
//   - {~~html} inline tags swap inline runs while leaving the host paragraph intact
//   - Output is a valid .docx zip
//
// What this does NOT yet verify (lands with postrender mergers):
//   - Image media files actually present in word/media/
//   - Image / hyperlink relationships registered in document.xml.rels
//   - List numbering merged into word/numbering.xml so bullets render
// The harness reports the pending mutations so you can see what the
// merger pass will need to handle.

import { readFileSync, writeFileSync } from 'node:fs';
import { dirname } from 'node:path';
import { fileURLToPath } from 'node:url';
import Docxtemplater from 'docxtemplater';
import PizZip from 'pizzip';

import { HtmlModule } from './html-module.mjs';

const __dirname = dirname(fileURLToPath(import.meta.url));
const FIXTURE = `${__dirname}/fixtures/template.docx`;
const OUT = `${__dirname}/out/rendered.docx`;

const blockHtml = `
<h2>Confirmation</h2>
<p>This letter confirms that the application has been <strong>approved</strong>. See <a href="https://example.org/policy">policy doc</a>.</p>
<ol>
  <li>Step one</li>
  <li>Step two
    <ul>
      <li>Sub A</li>
      <li>Sub B</li>
    </ul>
  </li>
</ol>
<p>Tiny logo: <img alt="dot" src="data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNgAAIAAAUAAen63NgAAAAASUVORK5CYII=" /></p>
<p>Signed.</p>
`;

const inlineHtml = `<strong>Mr. <em>John Doe</em></strong>`;

const main = async () => {
  const templateBuf = readFileSync(FIXTURE);
  const zip = new PizZip(templateBuf);

  const htmlModule = new HtmlModule({
    htmlValues: { bodyHtml: blockHtml, inlineHtml },
  });

  const doc = new Docxtemplater(zip, {
    paragraphLoop: true,
    linebreaks: true,
    modules: [htmlModule],
  });

  await doc.renderAsync({ recipient: 'Madam Director' });

  // Apply our zip-level mutations: numbering merge, rels merge,
  // image media, and final marker→ID rewrite in document.xml.
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
  const numXml = outZip.file('word/numbering.xml')?.asText() ?? '';
  const relsXml = outZip.file('word/_rels/document.xml.rels')?.asText() ?? '';
  const mediaFiles = Object.keys(outZip.files).filter((f) =>
    f.startsWith('word/media/'),
  );

  const checks = [
    ['valid PK header', out[0] === 0x50 && out[1] === 0x4b],
    ['contains rendered recipient', docXml.includes('Madam Director')],
    ['no leftover {~bodyHtml}', !docXml.includes('{~bodyHtml}')],
    ['no leftover {~~inlineHtml}', !docXml.includes('{~~inlineHtml}')],
    ['contains a heading2 from html', docXml.includes('Heading2')],
    ['contains list numId from html', /<w:numId/.test(docXml)],
    ['contains bold from inline html', docXml.includes('<w:b/>')],
    ['inline insertion preserves "end of mention"', docXml.includes('end of mention')],
    [
      'block expansion did NOT splice OOXML inside <w:t>',
      !/<w:t[^>]*>[^<]*<w:p[\s>]/.test(docXml),
    ],
    ['no unresolved __hmod_num_ markers in document.xml', !/__hmod_num_/.test(docXml)],
    ['no unresolved __hmod_rel_ markers in document.xml', !/__hmod_rel_/.test(docXml)],
    ['numbering.xml has at least 1 abstractNum from us', /<w:abstractNum/.test(numXml)],
    ['numbering.xml has at least 1 num from us', /<w:num\s/.test(numXml)],
    ['document.xml.rels has hyperlink rel from us', /Type="[^"]*hyperlink"/.test(relsXml)],
    ['document.xml.rels has image rel from us', /Type="[^"]*image"/.test(relsXml)],
    ['word/media/ has at least 1 image file', mediaFiles.length >= 1],
  ];

  console.log('\n=== assertions ===');
  let failures = 0;
  for (const [label, ok] of checks) {
    console.log((ok ? 'PASS  ' : 'FAIL  ') + label);
    if (!ok) failures += 1;
  }

  const headingPos = docXml.indexOf('Heading2');
  if (headingPos > 0) {
    const start = Math.max(0, headingPos - 200);
    const end = Math.min(docXml.length, headingPos + 600);
    console.log('\n=== document.xml around the block insertion ===');
    console.log(docXml.slice(start, end));
  }

  const pending = htmlModule.takePending();
  console.log('\n=== pending zip mutations (for the future merger pass) ===');
  console.log(`media:         ${pending.media.length} file(s)`);
  console.log(`relationships: ${pending.relationships.length}`);
  console.log(`numbering:     ${pending.numbering.length} list def(s)`);

  if (failures > 0) {
    console.error(`\n${failures} assertion(s) failed`);
    process.exit(1);
  }
  console.log('\nAll harness checks passed.');
};

main().catch((e) => {
  console.error(e);
  process.exit(1);
});
