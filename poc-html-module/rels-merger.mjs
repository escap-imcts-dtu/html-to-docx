// Merge HTML-module relationships (hyperlinks, image refs) into the
// host's word/_rels/document.xml.rels — and write image media bytes
// into word/media/.
//
// What we receive (per pending entry):
//   { fragmentRelId, finalRelId, type, target, targetMode }
// fragmentRelId is the numeric ID html-to-docx assigned (rId6, rId7,
// ...) inside its single fragment-render call. We assign a fresh
// finalRelId here and return the mapping so the caller can rewrite
// r:id="rIdX" references in word/document.xml.
//
// Media files come separately (from the same pending payload's
// `media` array). Each is named by its `nameInMedia` (e.g.
// `image-XXX.png`); the rel target should be `media/<nameInMedia>`.

const RELATIONSHIPS_PATH = 'word/_rels/document.xml.rels';
const MEDIA_PATH_PREFIX = 'word/media/';

const REL_ID_RE = /\bId="rId(\d+)"/g;

const findMaxRelId = (relsXml) => {
  let max = 0;
  for (const m of relsXml.matchAll(REL_ID_RE)) {
    const id = parseInt(m[1], 10);
    if (id > max) max = id;
  }
  return max;
};

const escapeXml = (s) =>
  String(s)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');

const buildRelXml = ({ id, type, target, targetMode }) => {
  const tmAttr =
    targetMode && targetMode !== 'Internal' ? ` TargetMode="${targetMode}"` : '';
  return (
    `<Relationship Id="${id}" Type="${escapeXml(type)}" ` +
    `Target="${escapeXml(target)}"${tmAttr}/>`
  );
};

/**
 * @param {import('pizzip')} zip
 * @param {{
 *   relationships: Array<{callIdx: number, fragmentRelId: number, finalRelId: string, type: string, target: string, targetMode: string}>,
 *   media: Array<{nameInMedia: string, contentType: string, data: Buffer}>,
 * }} pending
 * @returns {{ relIdRemap: Map<string, string>, addedRels: number, addedMedia: number }}
 *   relIdRemap is keyed by `${callIdx}:${fragmentRelId}`.
 */
export function mergeRelsAndMedia(zip, pending) {
  const result = {
    relIdRemap: new Map(),
    addedRels: 0,
    addedMedia: 0,
  };

  // 1) Add media files first — the rel targets reference them.
  if (pending.media && pending.media.length > 0) {
    for (const m of pending.media) {
      // PizZip's `file()` overwrites if the path exists; nameInMedia
      // includes nanoid so collisions with the host's existing media
      // are vanishingly unlikely.
      zip.file(`${MEDIA_PATH_PREFIX}${m.nameInMedia}`, m.data);
      result.addedMedia += 1;
    }
  }

  if (!pending.relationships || pending.relationships.length === 0) {
    return result;
  }

  let relsXml = zip.file(RELATIONSHIPS_PATH)?.asText();
  if (!relsXml) {
    // Templates ALWAYS have document.xml.rels in practice, but to be
    // safe — bail out with a clear shape rather than producing bad
    // output.
    return result;
  }

  const maxId = findMaxRelId(relsXml);
  let nextId = maxId + 1;

  const newRelXmls = [];
  for (const r of pending.relationships) {
    const finalId = `rId${nextId}`;
    nextId += 1;

    newRelXmls.push(
      buildRelXml({
        id: finalId,
        type: r.type,
        target: r.target,
        targetMode: r.targetMode,
      }),
    );
    result.relIdRemap.set(`${r.callIdx}:${r.fragmentRelId}`, finalId);
    result.addedRels += 1;
  }

  const insertion = newRelXmls.join('');
  let merged;
  if (/<\/Relationships>/.test(relsXml)) {
    merged = relsXml.replace(
      /<\/Relationships>/,
      `${insertion}</Relationships>`,
    );
  } else if (/<Relationships[^>]*\/>/.test(relsXml)) {
    merged = relsXml.replace(
      /<Relationships([^>]*)\/>/,
      `<Relationships$1>${insertion}</Relationships>`,
    );
  } else {
    return result;
  }

  zip.file(RELATIONSHIPS_PATH, merged);
  return result;
}
