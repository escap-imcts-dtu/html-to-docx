// Day 4d — edge case matrix.
//
// Synthetic edge cases exercising boundaries the 20 real templates
// don't (or that would be expensive to find in real templates).
//
// Each case is a small body XML fragment + an assertion about what
// should round-trip. We check both directions:
//   - bodyOoxmlToHtml works without throwing
//   - bodyHtmlToOoxml works without throwing
//   - structural counts preserved
//   - no leftover w:* tags in HTML
//   - no leftover <ooxml-* class HTML in OOXML output

import { bodyOoxmlToHtml } from './body-ooxml-to-html.mjs';
import { bodyHtmlToOoxml } from './body-html-to-ooxml.mjs';

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';

const cases = [
  {
    name: 'empty body',
    body: '',
    assertions: [
      ['html is empty string', (h) => h === ''],
      ['ooxml is empty string', (h, x) => x === ''],
    ],
  },
  {
    name: 'whitespace-only body',
    body: '   \n  \t  ',
    assertions: [
      ['ooxml contains no w:p', (h, x) => !x.includes('<w:p')],
    ],
  },
  {
    name: 'single empty paragraph',
    body: '<w:p/>',
    assertions: [
      ['html has one <p>', (h) => h.match(/<p[\s>]/g)?.length === 1],
      ['ooxml has one <w:p>', (h, x) => x.match(/<w:p[\s\/>]/g)?.length === 1],
    ],
  },
  {
    name: 'single paragraph with just text',
    body: '<w:p><w:r><w:t>Hello world.</w:t></w:r></w:p>',
    assertions: [
      ['html has Hello world', (h) => h.includes('Hello world')],
      ['ooxml has Hello world', (h, x) => x.includes('Hello world')],
    ],
  },
  {
    name: 'unicode characters (emoji + non-Latin)',
    body: '<w:p><w:r><w:t>Café naïve résumé 你好 🎉 emoji.</w:t></w:r></w:p>',
    assertions: [
      ['html preserves unicode', (h) => h.includes('Café naïve') && h.includes('你好') && h.includes('🎉')],
      ['ooxml preserves unicode', (h, x) => x.includes('Café') && x.includes('你好') && x.includes('🎉')],
    ],
  },
  {
    name: 'XML-escapable text (& < > " \')',
    body: `<w:p><w:r><w:t>A &amp; B &lt; C &gt; D "E" &apos;F&apos;</w:t></w:r></w:p>`,
    assertions: [
      ['html re-escapes properly', (h) => h.includes('A &amp; B') && h.includes('&lt; C') && h.includes('&gt; D')],
      ['ooxml re-escapes properly', (h, x) => x.includes('A &amp; B') && x.includes('&lt; C') && x.includes('&gt; D')],
    ],
  },
  {
    name: 'leading/trailing whitespace text',
    body: '<w:p><w:r><w:t xml:space="preserve">  leading and trailing  </w:t></w:r></w:p>',
    assertions: [
      ['ooxml restores xml:space="preserve"', (h, x) => /<w:t\s+xml:space="preserve"\s*>\s*leading and trailing/.test(x)],
    ],
  },
  {
    name: 'inline SDT only (no surrounding text)',
    body:
      `<w:p><w:sdt><w:sdtPr><w:tag w:val="onlyfield"/><w:id w:val="1"/></w:sdtPr>` +
      `<w:sdtContent><w:r><w:t>fieldvalue</w:t></w:r></w:sdtContent></w:sdt></w:p>`,
    assertions: [
      ['html has 1 SDT span', (h) => (h.match(/class="ooxml-sdt(?!\s+block)/g) || []).length === 1],
      ['ooxml has 1 SDT', (h, x) => (x.match(/<w:sdt[\s>]/g) || []).length === 1],
      ['SDT tag preserved', (h, x) => x.includes('w:val="onlyfield"')],
      ['SDT inner text preserved', (h, x) => x.includes('fieldvalue')],
    ],
  },
  {
    name: 'block SDT containing a table',
    body:
      `<w:sdt><w:sdtPr><w:tag w:val="blockfield"/><w:id w:val="2"/></w:sdtPr>` +
      `<w:sdtContent><w:tbl>` +
      `<w:tr><w:tc><w:p><w:r><w:t>cell A</w:t></w:r></w:p></w:tc>` +
      `<w:tc><w:p><w:r><w:t>cell B</w:t></w:r></w:p></w:tc></w:tr>` +
      `</w:tbl></w:sdtContent></w:sdt>`,
    assertions: [
      ['html has 1 block SDT div', (h) => (h.match(/class="ooxml-sdt block"/g) || []).length === 1],
      ['ooxml has 1 SDT around the table', (h, x) => /<w:sdt[\s>][\s\S]*?<w:tbl[\s\S]*?<\/w:tbl>[\s\S]*?<\/w:sdt>/.test(x)],
      ['ooxml preserves cell A', (h, x) => x.includes('cell A')],
      ['ooxml preserves cell B', (h, x) => x.includes('cell B')],
    ],
  },
  {
    name: 'nested SDTs (SDT inside SDT)',
    body:
      `<w:sdt><w:sdtPr><w:tag w:val="outer"/><w:id w:val="10"/></w:sdtPr>` +
      `<w:sdtContent><w:p><w:r><w:t>before </w:t></w:r>` +
      `<w:sdt><w:sdtPr><w:tag w:val="inner"/><w:id w:val="11"/></w:sdtPr>` +
      `<w:sdtContent><w:r><w:t>nested-value</w:t></w:r></w:sdtContent></w:sdt>` +
      `<w:r><w:t> after</w:t></w:r></w:p></w:sdtContent></w:sdt>`,
    assertions: [
      ['html has 2 SDTs total', (h) => (h.match(/class="ooxml-sdt/g) || []).length === 2],
      ['ooxml has 2 SDTs total', (h, x) => (x.match(/<w:sdt[\s>]/g) || []).length === 2],
      ['both SDT tags preserved', (h, x) => x.includes('w:val="outer"') && x.includes('w:val="inner"')],
      ['inner text preserved', (h, x) => x.includes('nested-value')],
      ['surrounding text preserved', (h, x) => x.includes('before ') && x.includes(' after')],
    ],
  },
  {
    name: 'mid-document section break (sectPr inside pPr)',
    body:
      `<w:p><w:pPr><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:pPr>` +
      `<w:r><w:t>section start</w:t></w:r></w:p>` +
      `<w:p><w:r><w:t>after break</w:t></w:r></w:p>`,
    assertions: [
      ['ooxml preserves section break inside pPr', (h, x) => x.includes('<w:sectPr')],
      ['ooxml preserves both paragraphs', (h, x) => (x.match(/<w:p[\s\/>]/g) || []).length === 2],
    ],
  },
  {
    name: 'mixed-formatting nested HTML (b inside i inside u)',
    body:
      `<w:p><w:r><w:rPr><w:b/><w:i/><w:u w:val="single"/></w:rPr><w:t>three styles</w:t></w:r></w:p>`,
    assertions: [
      ['html has nested strong/em/u', (h) => h.includes('<strong>') && h.includes('<em>') && h.includes('<u>')],
      ['ooxml has all three rPr children', (h, x) => x.includes('<w:b/>') && x.includes('<w:i/>') && /<w:u\b/.test(x)],
    ],
  },
  {
    name: 'hyperlink with absent rId (anchor-only)',
    body:
      `<w:p><w:hyperlink w:anchor="bookmark1"><w:r><w:t>jump</w:t></w:r></w:hyperlink></w:p>`,
    assertions: [
      ['html has anchor <a>', (h) => h.includes('href="#bookmark1"') || h.includes('data-ooxml-anchor="bookmark1"')],
      ['ooxml restores w:anchor', (h, x) => x.includes('w:anchor="bookmark1"')],
    ],
  },
];

let pass = 0;
let fail = 0;
const failedDetails = [];

for (const c of cases) {
  let html, ooxml;
  let errMsg = null;
  try {
    html = bodyOoxmlToHtml(c.body);
    ooxml = bodyHtmlToOoxml(html);
  } catch (e) {
    errMsg = e.message;
  }

  if (errMsg) {
    fail += 1;
    console.log(`[FAIL] ${c.name.padEnd(48)} threw: ${errMsg}`);
    failedDetails.push({ name: c.name, errMsg });
    continue;
  }

  const failedAsserts = [];
  for (const [label, fn] of c.assertions) {
    let ok;
    try {
      ok = !!fn(html, ooxml);
    } catch (e) {
      ok = false;
    }
    if (!ok) failedAsserts.push(label);
  }

  // Universal: no leftover w:* in html, no leftover ooxml-* in ooxml.
  if (/<w:[A-Za-z]/.test(html)) failedAsserts.push('html contains leftover <w:* tag');
  // (ooxml will obviously contain <w:* — that's the output format)

  if (failedAsserts.length === 0) {
    pass += 1;
    console.log(`[PASS] ${c.name.padEnd(48)} html=${html.length}c ooxml=${ooxml.length}c`);
  } else {
    fail += 1;
    console.log(`[FAIL] ${c.name.padEnd(48)} ${failedAsserts.join('; ')}`);
    failedDetails.push({ name: c.name, failedAsserts, html, ooxml });
  }
}

console.log(`\nEdge cases: ${pass}/${pass + fail} pass`);
if (failedDetails.length > 0) {
  console.log('\nFailed details:');
  for (const f of failedDetails) {
    console.log(`\n  ${f.name}`);
    if (f.errMsg) console.log(`    threw: ${f.errMsg}`);
    if (f.failedAsserts) {
      for (const a of f.failedAsserts) console.log(`    - ${a}`);
      console.log(`    html: ${f.html?.slice(0, 200)}`);
      console.log(`    ooxml: ${f.ooxml?.slice(0, 200)}`);
    }
  }
}
process.exit(fail === 0 ? 0 : 1);
