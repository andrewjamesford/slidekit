import JSZip from 'jszip';
import { DOMParser, XMLSerializer } from '@xmldom/xmldom';

/** PptxGenJS can emit one pPr per rich-text run. OOXML allows one, before runs.
 * Normalize our generated package, never the user's original PPTX.
 */
export async function normalizePptxParagraphs(bytes) {
  const zip = await JSZip.loadAsync(bytes);
  for (const file of Object.values(zip.files)) {
    const notes = /^ppt\/notesSlides\/notesSlide[^/]+\.xml$/.test(file.name);
    if (!notes && !/^ppt\/slides\/slide[^/]+\.xml$/.test(file.name)) continue;
    const doc = new DOMParser().parseFromString(await file.async('string'), 'application/xml');
    for (const paragraph of Array.from(doc.getElementsByTagNameNS('http://schemas.openxmlformats.org/drawingml/2006/main', 'p'))) {
      const properties = Array.from(paragraph.childNodes).filter(node => node.localName === 'pPr');
      for (const extra of properties.slice(1)) paragraph.removeChild(extra);
      if (properties[0] && paragraph.firstChild !== properties[0]) {
        paragraph.removeChild(properties[0]);
        paragraph.insertBefore(properties[0], paragraph.firstChild);
      }
    }
    if (notes) {
      // Literal newlines inside a:t survive self-import, but native PowerPoint
      // drops them. Our generated Notes contain one run; use real paragraphs.
      const drawing = 'http://schemas.openxmlformats.org/drawingml/2006/main';
      for (const paragraph of Array.from(doc.getElementsByTagNameNS(drawing, 'p'))) {
        const runs = Array.from(paragraph.childNodes).filter(node => node.localName === 'r');
        if (runs.length !== 1) continue;
        const value = runs[0].getElementsByTagNameNS(drawing, 't')[0];
        if (!value || !value.textContent.includes('\n')) continue;
        for (const line of value.textContent.split('\n')) {
          const replacement = paragraph.cloneNode(true);
          const text = replacement.getElementsByTagNameNS(drawing, 't')[0];
          text.textContent = line;
          text.setAttributeNS('http://www.w3.org/XML/1998/namespace', 'xml:space', 'preserve');
          paragraph.parentNode.insertBefore(replacement, paragraph);
        }
        paragraph.parentNode.removeChild(paragraph);
      }
    }
    zip.file(file.name, new XMLSerializer().serializeToString(doc));
  }
  return zip.generateAsync({type:'uint8array', compression:'DEFLATE'});
}
