// Fills a one-page Word (.docx) notice template once per participant and
// joins the pages into a single document.
//
// Template conventions (set by the administrator in Word):
//   • Text placeholders: {{編號}}, {{姓名}}, {{通訊地址}}, {{報到時段}} … — typed
//     anywhere in the document, including text boxes.  Word may split a
//     placeholder across several runs; that is handled here.
//   • Image placeholders: a picture whose alt text (描述) is {{報到QR}} or
//     {{導航QR}}.  The picture's size and position are kept; only the image
//     itself is swapped.
//
// Works in the browser (window.NoticeTemplate) and in Node (module.exports).
// The caller passes the JSZip constructor so this file has no dependencies.
(function (root) {
  'use strict';

  const PLACEHOLDER = /\{\{\s*([^{}\u0000]+?)\s*\}\}/g;
  const IMAGE_PLACEHOLDER = /(?:descr|title)="\{\{\s*([^{}"]+?)\s*\}\}"/g;
  const PARAGRAPH_BREAK = '\u0000';

  function decodeXml(value) {
    return value.replace(/&(lt|gt|quot|apos|amp|#(\d+)|#x([0-9a-fA-F]+));/g, (_, name, dec, hex) => {
      if (dec) return String.fromCodePoint(Number(dec));
      if (hex) return String.fromCodePoint(parseInt(hex, 16));
      return { lt: '<', gt: '>', quot: '"', apos: "'", amp: '&' }[name];
    });
  }

  function encodeXml(value) {
    return String(value).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  }

  // Splits the XML into text-run segments and paragraph boundaries so a
  // placeholder can be matched across runs but never across paragraphs.
  function tokenise(xml) {
    const re = /<w:t(?:\s[^>]*)?>([^<]*)<\/w:t>|<w:t\s*\/>|<w:p(?=[\s>\/])|<\/w:p>/g;
    const segments = [];
    let combined = '';
    let match;
    while ((match = re.exec(xml))) {
      if (match[0].startsWith('<w:t')) {
        const text = match[1] === undefined ? '' : decodeXml(match[1]);
        segments.push({ start: match.index, end: re.lastIndex, text, offset: combined.length });
        combined += text;
      } else {
        combined += PARAGRAPH_BREAK;
      }
    }
    return { segments, combined };
  }

  function listTextPlaceholders(xml) {
    const { combined } = tokenise(xml);
    const names = new Set();
    let match;
    PLACEHOLDER.lastIndex = 0;
    while ((match = PLACEHOLDER.exec(combined))) names.add(match[1]);
    return [...names];
  }

  function listImagePlaceholders(xml) {
    const names = new Set();
    let match;
    IMAGE_PLACEHOLDER.lastIndex = 0;
    while ((match = IMAGE_PLACEHOLDER.exec(xml))) names.add(match[1]);
    return [...names];
  }

  // Replaces every match of the global regex `pattern` in the document text,
  // even when Word split the matched text across several runs.  replace(match)
  // returns the new text, or undefined to leave that match untouched.  Line
  // breaks in the new text become Word line breaks.
  function replaceInText(xml, pattern, replace) {
    const { segments, combined } = tokenise(xml);
    if (!segments.length) return xml;
    const deleted = new Uint8Array(combined.length);
    const inserts = new Map(); // combined index -> replacement text
    let changed = false;
    let match;
    pattern.lastIndex = 0;
    while ((match = pattern.exec(combined))) {
      const value = replace(match);
      if (value === undefined || value === null) continue;
      changed = true;
      for (let i = match.index; i < match.index + match[0].length; i += 1) deleted[i] = 1;
      inserts.set(match.index, String(value));
    }
    if (!changed) return xml;

    let out = '';
    let cursor = 0;
    segments.forEach(segment => {
      let text = '';
      let touched = false;
      for (let i = 0; i < segment.text.length; i += 1) {
        const at = segment.offset + i;
        if (inserts.has(at)) { text += inserts.get(at); touched = true; }
        if (deleted[at]) { touched = true; continue; }
        text += segment.text[i];
      }
      if (!touched) return;
      out += xml.slice(cursor, segment.start);
      const parts = text.split(/\r?\n/).map(encodeXml);
      out += `<w:t xml:space="preserve">${parts.join('</w:t><w:br/><w:t xml:space="preserve">')}</w:t>`;
      cursor = segment.end;
    });
    out += xml.slice(cursor);
    return out;
  }

  // resolve(name) returns the value for {{name}}, or undefined to leave it.
  function replaceTextPlaceholders(xml, resolve) {
    return replaceInText(xml, PLACEHOLDER, match => resolve(match[1]));
  }

  // Replaces the image of every picture whose alt text is {{name}} by pointing
  // its blip at `relId`.  The blip follows the picture's docPr/cNvPr element.
  function replaceImagePlaceholder(xml, name, relId) {
    const marker = new RegExp(`(?:descr|title)="\\{\\{\\s*${name.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}\\s*\\}\\}"`, 'g');
    const positions = [];
    let match;
    while ((match = marker.exec(xml))) positions.push(match.index);
    if (!positions.length) return xml;
    const edits = [];
    positions.forEach(position => {
      const embed = /r:embed="([^"]*)"/g;
      embed.lastIndex = position;
      const found = embed.exec(xml);
      if (found && !edits.some(edit => edit.index === found.index)) edits.push({ index: found.index, length: found[0].length });
    });
    edits.sort((a, b) => b.index - a.index);
    let out = xml;
    edits.forEach(edit => { out = out.slice(0, edit.index) + `r:embed="${relId}"` + out.slice(edit.index + edit.length); });
    return out;
  }

  // Finds the start of the last top-level <w:p> in `content`.
  function lastTopLevelParagraph(content) {
    const re = /<w:p(?=[\s>\/])|<\/w:p>/g;
    const tokens = [];
    let match;
    while ((match = re.exec(content))) tokens.push({ index: match.index, close: match[0] === '</w:p>' });
    let depth = 0;
    for (let i = tokens.length - 1; i >= 0; i -= 1) {
      if (tokens[i].close) depth += 1;
      else {
        depth -= 1;
        if (depth === 0) return tokens[i].index;
      }
    }
    return -1;
  }

  function pageSectionProperties(sectPr) {
    // No <w:type> means "next page", which is what separates participants.
    return sectPr.replace(/<w:type\b[^>]*\/>/g, '');
  }

  // Appends a section break to the last paragraph of one participant's page.
  function endPageWithSection(content, sectPr) {
    const trimmed = content.replace(/\s+$/, '');
    const start = trimmed.endsWith('</w:p>') ? lastTopLevelParagraph(trimmed) : -1;
    if (start < 0) return `${content}<w:p><w:pPr>${sectPr}</w:pPr></w:p>`;
    const paragraph = trimmed.slice(start);
    const openEnd = paragraph.indexOf('>') + 1;
    const selfClosing = paragraph[openEnd - 2] === '/';
    if (selfClosing) {
      return trimmed.slice(0, start) + paragraph.slice(0, openEnd - 2) + `><w:pPr>${sectPr}</w:pPr></w:p>` + paragraph.slice(openEnd);
    }
    const afterOpen = paragraph.slice(openEnd);
    if (afterOpen.startsWith('<w:pPr/>')) {
      return trimmed.slice(0, start) + paragraph.slice(0, openEnd) + `<w:pPr>${sectPr}</w:pPr>` + afterOpen.slice('<w:pPr/>'.length);
    }
    if (afterOpen.startsWith('<w:pPr>') || afterOpen.startsWith('<w:pPr ')) {
      const closeAt = afterOpen.indexOf('</w:pPr>');
      const pPr = afterOpen.slice(0, closeAt).replace(/<w:sectPr\b[\s\S]*?<\/w:sectPr>/g, '');
      return trimmed.slice(0, start) + paragraph.slice(0, openEnd) + pPr + sectPr + afterOpen.slice(closeAt);
    }
    return trimmed.slice(0, start) + paragraph.slice(0, openEnd) + `<w:pPr>${sectPr}</w:pPr>` + afterOpen;
  }

  function splitDocument(documentXml) {
    const bodyOpen = documentXml.indexOf('<w:body>');
    const bodyClose = documentXml.lastIndexOf('</w:body>');
    if (bodyOpen < 0 || bodyClose < 0) throw new Error('範本不是有效的 Word 文件。');
    const body = documentXml.slice(bodyOpen + '<w:body>'.length, bodyClose);
    const sectAt = body.lastIndexOf('<w:sectPr');
    const sectPr = sectAt >= 0 ? body.slice(sectAt) : '<w:sectPr/>';
    let content = sectAt >= 0 ? body.slice(0, sectAt) : body;
    // A template saved from a merged file may still carry per-page section
    // breaks; the generator adds its own.
    content = content.replace(/<w:sectPr\b[\s\S]*?<\/w:sectPr>/g, '');
    return {
      head: documentXml.slice(0, bodyOpen + '<w:body>'.length),
      content,
      sectPr: sectPr.trim(),
      tail: documentXml.slice(bodyClose),
    };
  }

  async function inspectTemplate(JSZip, bytes) {
    let zip;
    try { zip = await JSZip.loadAsync(bytes); } catch { throw new Error('檔案不是 Word（.docx）格式。'); }
    const file = zip.file('word/document.xml');
    if (!file) throw new Error('檔案不是 Word（.docx）格式。');
    const xml = await file.async('string');
    const { content } = splitDocument(xml);
    return {
      text: listTextPlaceholders(content),
      images: listImagePlaceholders(content),
      hasMultiplePages: /<w:br\b[^>]*w:type="page"/.test(content),
    };
  }

  function ensureContentType(typesXml, extension, contentType) {
    if (new RegExp(`Extension="${extension}"`, 'i').test(typesXml)) return typesXml;
    return typesXml.replace('</Types>', `<Default Extension="${extension}" ContentType="${contentType}"/></Types>`);
  }

  const IMAGE_TYPES = { gif: 'image/gif', png: 'image/png', jpeg: 'image/jpeg', jpg: 'image/jpeg' };

  // pages: [{ values: { 編號: '…', … }, images: { 報到QR: { bytes, ext } } }]
  // sharedImages: { 導航QR: { bytes, ext } } — the same image on every page.
  async function mergeTemplate(JSZip, templateBytes, pages, { sharedImages = {} } = {}) {
    if (!pages.length) throw new Error('沒有可產出的參與者。');
    const zip = await JSZip.loadAsync(templateBytes);
    const documentXml = await zip.file('word/document.xml').async('string');
    const relsPath = 'word/_rels/document.xml.rels';
    let rels = await zip.file(relsPath).async('string');
    let types = await zip.file('[Content_Types].xml').async('string');
    const { head, content, sectPr, tail } = splitDocument(documentXml);
    const pageSect = pageSectionProperties(sectPr);
    let relCounter = 0;
    const newRels = [];

    function addImage(image, stem) {
      relCounter += 1;
      const ext = (image.ext || 'png').toLowerCase();
      const relId = `rIdNotice${relCounter}`;
      const target = `media/notice_${stem}_${relCounter}.${ext}`;
      zip.file(`word/${target}`, image.bytes);
      types = ensureContentType(types, ext, IMAGE_TYPES[ext] || 'application/octet-stream');
      newRels.push(`<Relationship Id="${relId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="${target}"/>`);
      return relId;
    }

    let base = content;
    Object.entries(sharedImages).forEach(([name, image]) => {
      if (image && image.bytes) base = replaceImagePlaceholder(base, name, addImage(image, 'shared'));
    });

    const parts = [];
    pages.forEach((page, index) => {
      const values = page.values || {};
      let copy = replaceTextPlaceholders(base, name => (Object.prototype.hasOwnProperty.call(values, name) ? values[name] : undefined));
      Object.entries(page.images || {}).forEach(([name, image]) => {
        if (image && image.bytes) copy = replaceImagePlaceholder(copy, name, addImage(image, 'p'));
      });
      parts.push(index < pages.length - 1 ? endPageWithSection(copy, pageSect) : copy);
    });

    let body = parts.join('') + pageSect;
    // Copies of the same page share ids that Word expects to be unique.
    let docPrId = 0;
    body = body.replace(/(<wp:docPr\b[^>]*?\sid=")\d+(")/g, (_, a, b) => `${a}${(docPrId += 1)}${b}`);
    body = body.replace(/\sw14:(?:paraId|textId)="[^"]*"/g, '');
    body = body.replace(/<w:bookmark(?:Start|End)\b[^>]*\/>/g, '');

    zip.file('word/document.xml', head + body + tail);
    zip.file(relsPath, rels.replace('</Relationships>', `${newRels.join('')}</Relationships>`));
    zip.file('[Content_Types].xml', types);
    return zip;
  }

  const api = {
    listTextPlaceholders,
    listImagePlaceholders,
    replaceInText,
    replaceTextPlaceholders,
    replaceImagePlaceholder,
    inspectTemplate,
    mergeTemplate,
  };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else root.NoticeTemplate = api;
})(typeof window !== 'undefined' ? window : globalThis);
