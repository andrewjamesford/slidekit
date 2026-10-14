import JSZip from 'jszip';
import { DOMParser, XMLSerializer } from '@xmldom/xmldom';

/** PptxGenJS can emit one pPr per rich-text run. OOXML allows one, before runs.
 * Normalize our generated package, never the user's original PPTX.
 */
export async function normalizePptxParagraphs(bytes) {
  const zip = await JSZip.loadAsync(bytes);
  for (const file of Object.values(zip.files)) {
    if (!/^ppt\/slides\/slide[^/]+\.xml$/.test(file.name)) continue;
    const doc = new DOMParser().parseFromString(await file.async('string'), 'application/xml');
    for (const paragraph of Array.from(doc.getElementsByTagNameNS('http://schemas.openxmlformats.org/drawingml/2006/main', 'p'))) {
      const properties = Array.from(paragraph.childNodes).filter(node => node.localName === 'pPr');
      for (const extra of properties.slice(1)) paragraph.removeChild(extra);
      if (properties[0] && paragraph.firstChild !== properties[0]) {
        paragraph.removeChild(properties[0]);
        paragraph.insertBefore(properties[0], paragraph.firstChild);
      }
    }
    zip.file(file.name, new XMLSerializer().serializeToString(doc));
  }
  return zip.generateAsync({type:'uint8array', compression:'DEFLATE'});
}
