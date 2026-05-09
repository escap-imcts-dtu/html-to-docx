// Build a more ESCAP-realistic .docx fixture exercising the cases the
// HCRU MFA pipeline cares about:
//
//   1. Word Content Controls (<w:sdt>) wrapping placeholders — HCRU
//      authors templates this way, not as bare {tag}s.
//   2. A header (with org name) and footer (with page number field).
//   3. A `{#applicants}...{/applicants}` loop containing
//      `{name}`, `{ref}`, AND `{~notesHtml}` (rich-text notes per row).
//   4. A static 2-column table with a bold header row (must survive
//      untouched).
//   5. A signature block at the end with `{signerName}`, `{signerDate}`.
//
// Output: poc-html-module/fixtures/complex-template.docx

import { writeFileSync } from 'node:fs';
import { dirname } from 'node:path';
import { fileURLToPath } from 'node:url';
import PizZip from 'pizzip';

const __dirname = dirname(fileURLToPath(import.meta.url));
const OUT = `${__dirname}/fixtures/complex-template.docx`;

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

// -------------------------------------------------------------------
// document.xml — the meaty bit
// -------------------------------------------------------------------

const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${W}" xmlns:r="${R}">
  <w:body>

    <w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t xml:space="preserve">Notification of decisions</w:t></w:r></w:p>

    <w:p><w:r><w:t xml:space="preserve">Dear </w:t></w:r><w:sdt><w:sdtPr><w:alias w:val="Recipient"/><w:tag w:val="recipient"/><w:id w:val="100001"/></w:sdtPr><w:sdtContent><w:r><w:t xml:space="preserve">{recipient}</w:t></w:r></w:sdtContent></w:sdt><w:r><w:t xml:space="preserve">,</w:t></w:r></w:p>

    <w:p><w:r><w:t xml:space="preserve">Reference: </w:t></w:r><w:sdt><w:sdtPr><w:alias w:val="Case ref"/><w:tag w:val="caseRef"/><w:id w:val="100002"/></w:sdtPr><w:sdtContent><w:r><w:t xml:space="preserve">{caseRef}</w:t></w:r></w:sdtContent></w:sdt></w:p>

    <w:p><w:r><w:t xml:space="preserve">Below is the list of applicants and the decisions concerning them. Per applicant, the &quot;Notes&quot; column may contain rich-text guidance authored by case officers.</w:t></w:r></w:p>

    <w:p><w:pPr><w:pStyle w:val="Heading2"/></w:pPr><w:r><w:t xml:space="preserve">Applicants</w:t></w:r></w:p>

    <w:p><w:r><w:t xml:space="preserve">{#applicants}</w:t></w:r></w:p>

    <w:p><w:pPr><w:pStyle w:val="Heading3"/></w:pPr><w:r><w:t xml:space="preserve">{name}</w:t></w:r><w:r><w:t xml:space="preserve"> (ref. </w:t></w:r><w:r><w:t xml:space="preserve">{ref}</w:t></w:r><w:r><w:t xml:space="preserve">)</w:t></w:r></w:p>

    <w:p><w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">Notes:</w:t></w:r></w:p>

    <w:p><w:r><w:t>{~notesHtml}</w:t></w:r></w:p>

    <w:p><w:r><w:t xml:space="preserve">Inline tag: status is </w:t></w:r><w:r><w:t xml:space="preserve">{~~statusHtml}</w:t></w:r><w:r><w:t xml:space="preserve">.</w:t></w:r></w:p>

    <w:p><w:r><w:t xml:space="preserve">{/applicants}</w:t></w:r></w:p>

    <w:p><w:pPr><w:pStyle w:val="Heading2"/></w:pPr><w:r><w:t xml:space="preserve">Static reference table (must survive untouched)</w:t></w:r></w:p>

    <w:tbl>
      <w:tblPr><w:tblW w:w="5000" w:type="pct"/><w:tblBorders><w:top w:val="single" w:sz="4" w:color="auto"/><w:left w:val="single" w:sz="4" w:color="auto"/><w:bottom w:val="single" w:sz="4" w:color="auto"/><w:right w:val="single" w:sz="4" w:color="auto"/><w:insideH w:val="single" w:sz="4" w:color="auto"/><w:insideV w:val="single" w:sz="4" w:color="auto"/></w:tblBorders></w:tblPr>
      <w:tblGrid><w:gridCol w:w="2500"/><w:gridCol w:w="2500"/></w:tblGrid>
      <w:tr>
        <w:trPr><w:tblHeader/></w:trPr>
        <w:tc><w:tcPr><w:tcW w:w="2500" w:type="dxa"/></w:tcPr><w:p><w:r><w:rPr><w:b/></w:rPr><w:t>Code</w:t></w:r></w:p></w:tc>
        <w:tc><w:tcPr><w:tcW w:w="2500" w:type="dxa"/></w:tcPr><w:p><w:r><w:rPr><w:b/></w:rPr><w:t>Meaning</w:t></w:r></w:p></w:tc>
      </w:tr>
      <w:tr>
        <w:tc><w:tcPr><w:tcW w:w="2500" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>A</w:t></w:r></w:p></w:tc>
        <w:tc><w:tcPr><w:tcW w:w="2500" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>Approved</w:t></w:r></w:p></w:tc>
      </w:tr>
      <w:tr>
        <w:tc><w:tcPr><w:tcW w:w="2500" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>R</w:t></w:r></w:p></w:tc>
        <w:tc><w:tcPr><w:tcW w:w="2500" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>Rejected</w:t></w:r></w:p></w:tc>
      </w:tr>
    </w:tbl>

    <w:p/>

    <w:p><w:r><w:t xml:space="preserve">Signed,</w:t></w:r></w:p>
    <w:p><w:sdt><w:sdtPr><w:alias w:val="Signer name"/><w:tag w:val="signerName"/><w:id w:val="100003"/></w:sdtPr><w:sdtContent><w:r><w:t xml:space="preserve">{signerName}</w:t></w:r></w:sdtContent></w:sdt></w:p>
    <w:p><w:r><w:t xml:space="preserve">Date: </w:t></w:r><w:sdt><w:sdtPr><w:alias w:val="Signer date"/><w:tag w:val="signerDate"/><w:id w:val="100004"/></w:sdtPr><w:sdtContent><w:r><w:t xml:space="preserve">{signerDate}</w:t></w:r></w:sdtContent></w:sdt></w:p>

    <w:sectPr>
      <w:headerReference w:type="default" r:id="rIdHeader1"/>
      <w:footerReference w:type="default" r:id="rIdFooter1"/>
      <w:pgSz w:w="12240" w:h="15840" />
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0" />
    </w:sectPr>
  </w:body>
</w:document>`;

const headerXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:hdr xmlns:w="${W}">
  <w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">UN ESCAP — IMCTS / DTU</w:t></w:r></w:p>
</w:hdr>`;

const footerXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:ftr xmlns:w="${W}">
  <w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:t xml:space="preserve">Page </w:t></w:r><w:fldSimple w:instr="PAGE"><w:r><w:t>1</w:t></w:r></w:fldSimple><w:r><w:t xml:space="preserve"> of </w:t></w:r><w:fldSimple w:instr="NUMPAGES"><w:r><w:t>1</w:t></w:r></w:fldSimple></w:p>
</w:ftr>`;

// -------------------------------------------------------------------
// boilerplate parts (copied wholesale from build-fixture.mjs and
// extended for header/footer + complex content)
// -------------------------------------------------------------------

const contentTypesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="png" ContentType="image/png"/>
  <Default Extension="jpeg" ContentType="image/jpeg"/>
  <Default Extension="gif" ContentType="image/gif"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>
  <Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
  <Override PartName="/word/fontTable.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml"/>
  <Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>
  <Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>`;

const rootRelsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>`;

// IMPORTANT: header/footer rIds use string IDs ("rIdHeader1"/"rIdFooter1")
// referenced from <w:headerReference> / <w:footerReference> in the
// sectPr. Numeric rIds 1..4 are the standard parts.
const docRelsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
  <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable" Target="fontTable.xml"/>
  <Relationship Id="rIdHeader1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>
  <Relationship Id="rIdFooter1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>
</Relationships>`;

const stylesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="${W}">
  <w:docDefaults>
    <w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="22"/></w:rPr></w:rPrDefault>
    <w:pPrDefault><w:pPr><w:spacing w:after="160" w:line="259" w:lineRule="auto"/></w:pPr></w:pPrDefault>
  </w:docDefaults>
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
  <w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:pPr><w:keepNext/><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:b/><w:sz w:val="36"/></w:rPr></w:style>
  <w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:basedOn w:val="Normal"/><w:pPr><w:keepNext/><w:outlineLvl w:val="1"/></w:pPr><w:rPr><w:b/><w:sz w:val="28"/></w:rPr></w:style>
  <w:style w:type="paragraph" w:styleId="Heading3"><w:name w:val="heading 3"/><w:basedOn w:val="Normal"/><w:pPr><w:keepNext/><w:outlineLvl w:val="2"/></w:pPr><w:rPr><w:b/><w:sz w:val="24"/></w:rPr></w:style>
  <w:style w:type="character" w:styleId="Hyperlink"><w:name w:val="Hyperlink"/><w:rPr><w:color w:val="0563C1"/><w:u w:val="single"/></w:rPr></w:style>
</w:styles>`;

const numberingXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:w="${W}"></w:numbering>`;

const settingsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="${W}">
  <w:zoom w:percent="100"/>
  <w:defaultTabStop w:val="720"/>
  <w:characterSpacingControl w:val="doNotCompress"/>
</w:settings>`;

const fontTableXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:fonts xmlns:w="${W}">
  <w:font w:name="Calibri"><w:panose1 w:val="020F0502020204030204"/><w:charset w:val="00"/><w:family w:val="swiss"/><w:pitch w:val="variable"/></w:font>
  <w:font w:name="Times New Roman"><w:panose1 w:val="02020603050405020304"/><w:charset w:val="00"/><w:family w:val="roman"/><w:pitch w:val="variable"/></w:font>
  <w:font w:name="Symbol"><w:panose1 w:val="05050102010706020507"/><w:charset w:val="02"/><w:family w:val="roman"/><w:pitch w:val="variable"/></w:font>
</w:fonts>`;

const corePropsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>ESCAP HtmlModule POC — complex fixture</dc:title>
  <dc:creator>ESCAP IMCTS DTU</dc:creator>
  <cp:lastModifiedBy>ESCAP IMCTS DTU</cp:lastModifiedBy>
  <cp:revision>1</cp:revision>
  <dcterms:created xsi:type="dcterms:W3CDTF">2026-05-09T12:00:00Z</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">2026-05-09T12:00:00Z</dcterms:modified>
</cp:coreProperties>`;

const appPropsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Application>ESCAP HtmlModule POC</Application>
  <DocSecurity>0</DocSecurity>
  <ScaleCrop>false</ScaleCrop>
  <SharedDoc>false</SharedDoc>
  <HyperlinksChanged>false</HyperlinksChanged>
  <AppVersion>16.0000</AppVersion>
</Properties>`;

const zip = new PizZip();
zip.file('[Content_Types].xml', contentTypesXml);
zip.folder('_rels').file('.rels', rootRelsXml);
zip.folder('docProps').file('core.xml', corePropsXml);
zip.folder('docProps').file('app.xml', appPropsXml);
zip.folder('word').file('document.xml', documentXml);
zip.folder('word').file('header1.xml', headerXml);
zip.folder('word').file('footer1.xml', footerXml);
zip.folder('word').file('styles.xml', stylesXml);
zip.folder('word').file('numbering.xml', numberingXml);
zip.folder('word').file('settings.xml', settingsXml);
zip.folder('word').file('fontTable.xml', fontTableXml);
zip.folder('word').folder('_rels').file('document.xml.rels', docRelsXml);

const buffer = zip.generate({ type: 'nodebuffer' });
writeFileSync(OUT, buffer);
console.log(`Wrote ${OUT} (${buffer.length} bytes)`);
