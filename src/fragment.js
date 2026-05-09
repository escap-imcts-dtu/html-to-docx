/**
 * Fragment API — produce just the OOXML body content (no full .docx zip).
 *
 * Companion to the default export (`generateContainer`), which builds a
 * complete `.docx`. The fragment API returns the building blocks needed
 * to splice HTML content into an existing document — the use case being
 * a docxtemplater HTML module that swaps `{~html}` placeholders for the
 * OOXML rendering of an HTML value.
 *
 * Returns an object with:
 *   - bodyXml:       string of <w:p>/<w:tbl>/... elements (no <w:body> wrapper)
 *   - media:         [{ nameInMedia, contentType, data }] for word/media/
 *   - relationships: [{ relationshipId, type, target, targetMode }]
 *   - numbering:     [{ numberingId, type, properties }] list defs
 */

import JSZip from 'jszip';
import { create } from 'xmlbuilder2';
import { decode } from 'html-entities';
import { minify } from 'html-minifier-terser';

import DocxDocument from './docx-document';
import { renderDocumentFile } from './helpers';
import createDocumentOptionsAndMergeWithDefaults from './utils/options-utils';
import namespaces from './namespaces';

const MIME_BY_EXT = {
  png: 'image/png',
  jpg: 'image/jpeg',
  jpeg: 'image/jpeg',
  gif: 'image/gif',
  bmp: 'image/bmp',
  svg: 'image/svg+xml',
  webp: 'image/webp',
};

const minifyHTMLString = async (htmlString) => {
  if (typeof htmlString !== 'string' && !(htmlString instanceof String)) {
    return null;
  }
  try {
    return await minify(htmlString, {
      collapseWhitespace: true,
      removeComments: true,
    });
  } catch (e) {
    return null;
  }
};

const guessContentType = (fileName) => {
  const ext = (fileName.split('.').pop() || '').toLowerCase();
  return MIME_BY_EXT[ext] || 'application/octet-stream';
};

async function htmlToOoxmlFragment(htmlString, documentOptions = {}) {
  if (typeof htmlString !== 'string' || htmlString.length === 0) {
    return {
      bodyXml: '',
      media: [],
      relationships: [],
      numbering: [],
    };
  }

  const normalizedDocumentOptions = createDocumentOptionsAndMergeWithDefaults(documentOptions);

  let contentHTML = htmlString;
  if (!normalizedDocumentOptions.preprocessing.skipHTMLMinify) {
    const minified = await minifyHTMLString(contentHTML);
    if (minified !== null) {
      contentHTML = minified;
    }
  }
  if (normalizedDocumentOptions.decodeUnicode) {
    contentHTML = decode(contentHTML);
  }

  // Throwaway zip — receives image media as a side-effect of buildImage.
  // We harvest it after render and discard the rest.
  const zip = new JSZip();

  const docxDocument = new DocxDocument({
    zip,
    htmlString: contentHTML,
    ...normalizedDocumentOptions,
  });

  const xmlFragment = await renderDocumentFile(docxDocument);

  // Serialize children of the fragment with a parent in scope that
  // declares xmlns:w (and friends). Without a parent declaration,
  // xmlbuilder2 falls back to default-namespace form (<p xmlns="...">)
  // or generated prefixes (ns1:, ns2:), neither of which mixes well
  // with the docxtemplater host document's <w:*> elements.
  //
  // We wrap in a temporary <_h2dx_wrap> element, serialize the whole
  // thing, then strip the wrapper tags from the resulting string.
  const wrapper = create({ encoding: 'UTF-8' }).ele('_h2dx_wrap', {
    [`xmlns:w`]: namespaces.w,
    [`xmlns:r`]: namespaces.r,
    [`xmlns:wp`]: namespaces.wp,
    [`xmlns:a`]: namespaces.a,
    [`xmlns:pic`]: namespaces.pic,
    [`xmlns:o`]: namespaces.o,
    [`xmlns:v`]: namespaces.v,
  });
  wrapper.import(xmlFragment);
  const wrapped = wrapper.end({ headless: true });
  // Strip the wrapper's open and close tags. The wrapper has no
  // attributes other than xmlns:* so it's always self-contained.
  const openEnd = wrapped.indexOf('>');
  const closeStart = wrapped.lastIndexOf('</_h2dx_wrap>');
  const bodyXml =
    openEnd >= 0 && closeStart > openEnd ? wrapped.slice(openEnd + 1, closeStart) : wrapped;

  // Harvest media files written into word/media/ during render.
  const media = [];
  const mediaPromises = [];
  zip.forEach((relativePath, file) => {
    if (!relativePath.startsWith('word/media/') || file.dir) return;
    const nameInMedia = relativePath.slice('word/media/'.length);
    mediaPromises.push(
      file.async('nodebuffer').then((data) => {
        media.push({
          nameInMedia,
          contentType: guessContentType(nameInMedia),
          data,
        });
      })
    );
  });
  await Promise.all(mediaPromises);

  // Relationships: only the document-scoped ones matter for body content.
  // Hyperlinks and images are recorded here by createDocumentRelationships.
  const docRels = docxDocument.relationships.find((r) => r.fileName === 'document') || { rels: [] };

  return {
    bodyXml,
    media,
    relationships: docRels.rels.map((r) => ({ ...r })),
    numbering: docxDocument.numberingObjects.map((n) => ({ ...n })),
  };
}

export default htmlToOoxmlFragment;
