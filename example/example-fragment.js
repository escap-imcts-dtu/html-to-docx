/* eslint-disable */
// Smoke test for the new fragment API. Runs the built ESM bundle.
//
// Usage: node example/example-fragment.js

import { htmlToOoxmlFragment } from '../dist/html-to-docx.esm.js';

const RICH_HTML = `
<h2>Greetings</h2>
<p>Dear <strong>Madam Director</strong>,</p>
<p>This letter confirms that the following actions are <em>complete</em>:</p>
<ol>
  <li>Phase 1 onboarding</li>
  <li>Phase 2 review
    <ul>
      <li>Sub-item A</li>
      <li>Sub-item B</li>
    </ul>
  </li>
</ol>
<p>A summary table:</p>
<table>
  <thead>
    <tr><th>Item</th><th>Status</th></tr>
  </thead>
  <tbody>
    <tr><td>X</td><td>OK</td></tr>
    <tr><td>Y</td><td>Pending</td></tr>
  </tbody>
</table>
<p>With characters needing escape: &amp; &lt; &gt; "quoted" and a <a href="https://example.org">link</a>.</p>
<p>Empty paragraph next:</p>
<p></p>
<p>Inline mix: <strong><em>bold-italic</em></strong> and a small <img alt="tiny" src="data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNgAAIAAAUAAen63NgAAAAASUVORK5CYII=" /> inline image.</p>
`;

const main = async () => {
  console.log('Calling htmlToOoxmlFragment ...');
  const result = await htmlToOoxmlFragment(RICH_HTML);

  console.log('\n=== bodyXml (first 600 chars) ===');
  console.log(result.bodyXml.slice(0, 600));
  console.log('...');
  console.log('=== bodyXml length:', result.bodyXml.length, 'chars ===');

  console.log('\n=== media:', result.media.length, 'file(s) ===');
  for (const m of result.media) {
    console.log('  -', m.nameInMedia, m.contentType, m.data.length, 'bytes');
  }

  console.log('\n=== relationships:', result.relationships.length, '===');
  for (const r of result.relationships) {
    console.log('  - rId' + r.relationshipId, r.type, '->', r.target, '(' + r.targetMode + ')');
  }

  console.log('\n=== numbering:', result.numbering.length, 'list def(s) ===');
  for (const n of result.numbering) {
    console.log('  - id=' + n.numberingId, 'type=' + n.type);
  }

  // Quick structural assertions
  const checks = [
    ['has <w:p>', result.bodyXml.includes('<w:p')],
    ['has <w:tbl>', result.bodyXml.includes('<w:tbl')],
    ['has <w:r>', result.bodyXml.includes('<w:r')],
    ['has at least one image', result.media.length >= 1],
    ['has at least one image rel', result.relationships.some((r) => /image/.test(r.type))],
    ['has list defs', result.numbering.length >= 1],
  ];
  console.log('\n=== assertions ===');
  let failures = 0;
  for (const [label, ok] of checks) {
    console.log((ok ? 'PASS  ' : 'FAIL  ') + label);
    if (!ok) failures += 1;
  }
  if (failures > 0) {
    console.error(`\n${failures} assertion(s) failed`);
    process.exit(1);
  }
  console.log('\nAll smoke checks passed.');
};

main().catch((e) => {
  console.error(e);
  process.exit(1);
});
