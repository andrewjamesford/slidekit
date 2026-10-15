import { validateImageFrame, validateRetainedText } from './document.js';
import { at, attr, children, EMU_PER_PIXEL } from './pptx-package.js';
const prefix = 'slidekit:migration:1:';
const xml = (doc, name, attrs = {}, content = []) => {
  const node = doc.createElementNS('http://schemas.openxmlformats.org/drawingml/2006/main', 'a:' + name);
  for (const [key, value] of Object.entries(attrs)) node.setAttribute(key, String(value));
  content.forEach(child => node.appendChild(child)); return node;
};
function geometry(doc, frame, box) {
  const w = Math.round(box.width * EMU_PER_PIXEL), h = Math.round(box.height * EMU_PER_PIXEL);
  const r = Math.round(Math.min(frame.radius, box.width / 2, box.height / 2) * EMU_PER_PIXEL);
  const pt = (name, x, y) => xml(doc, name, {}, [xml(doc, 'pt', { x, y })]);
  const arc = angle => xml(doc, 'arcTo', { wR: r, hR: r, stAng: angle * 60000, swAng: 5400000 });
  const path = r === 0 ? [pt('moveTo', 0, 0), pt('lnTo', w, 0), pt('lnTo', w, h), pt('lnTo', 0, h), xml(doc, 'close')] : [pt('moveTo', r, 0), pt('lnTo', w-r, 0), arc(270), pt('lnTo', w, h-r), arc(0), pt('lnTo', r, h), arc(90), pt('lnTo', 0, r), arc(180), xml(doc, 'close')];
  return xml(doc, 'custGeom', {}, [xml(doc, 'avLst'), xml(doc, 'gdLst'), xml(doc, 'ahLst'), xml(doc, 'cxnLst'), xml(doc, 'rect', { l: 0, t: 0, r: w, b: h }), xml(doc, 'pathLst', {}, [xml(doc, 'path', { w, h }, path)])]);
}
const tree = node => node && [node.localName, Array.from(node.attributes).filter(a => !a.name.startsWith('xmlns')).map(a => [a.localName, a.value]).sort(), children(node).map(tree)];
function fillMargins(box, asset) {
  const w = Math.round(box.width * EMU_PER_PIXEL) / EMU_PER_PIXEL, h = Math.round(box.height * EMU_PER_PIXEL) / EMU_PER_PIXEL;
  const scale = Math.min(w / asset.width, h / asset.height);
  const l = Math.round((w - asset.width * scale) / w * 50000), t = Math.round((h - asset.height * scale) / h * 50000);
  return { l, t, r: l, b: t };
}
export function writeMigrationAppearance(node, element, asset) {
  if (!element.frame && !element.retained) return;
  const nonVisual = at(node, node.localName === 'pic' ? 'nvPicPr/cNvPr' : 'nvSpPr/cNvPr');
  nonVisual.setAttribute('descr', prefix + JSON.stringify(element.frame ? { frame: element.frame, fit: element.fit } : { retained: element.retained }));
  if (element.frame) {
    const properties = at(node, 'spPr'), old = at(properties, 'prstGeom') || at(properties, 'custGeom');
    if (!old) throw Error('Framed picture geometry is missing');
    properties.replaceChild(geometry(node.ownerDocument, element.frame, element), old);
    if (element.fit === 'contain') {
      const stretch = at(node, 'blipFill/stretch');
      if (!stretch || !asset) throw Error('Framed picture fill is missing');
      const oldFill = at(stretch, 'fillRect');
      const fill = xml(node.ownerDocument, 'fillRect', fillMargins(element, asset));
      if (oldFill) stretch.replaceChild(fill, oldFill); else stretch.appendChild(fill);
    }
  }
}
export function readMigrationMarker(nonVisual) {
  const value = attr(nonVisual, 'descr') || '';
  if (!value.startsWith(prefix)) return null;
  if (value.length > 4096) throw Error('Migration picture/text metadata exceeds budget');
  const data = JSON.parse(value.slice(prefix.length));
  if (!data || typeof data !== 'object') throw Error('Malformed migration metadata');
  if (Object.keys(data).sort().join() === 'fit,frame' && ['contain', 'cover'].includes(data.fit)) return { frame: validateImageFrame(data.frame), fit: data.fit };
  if (Object.keys(data).join() === 'retained') return { retained: validateRetainedText(data.retained) };
  throw Error('Unsupported migration metadata');
}
export function verifyFrameGeometry(properties, frame, box, shadow, blipFill, asset, fit) {
  if (JSON.stringify(tree(at(properties, 'custGeom'))) !== JSON.stringify(tree(geometry(properties.ownerDocument, frame, box)))) throw Error('Edited or unsupported framed picture geometry requires review');
  if (fit === 'contain') {
    const fill = at(blipFill, 'stretch/fillRect'), expected = fillMargins(box, asset), crop = at(blipFill, 'srcRect');
    if (!fill || Object.entries(expected).some(([key, value]) => Number(attr(fill, key, '0')) !== value) || crop && ['l','t','r','b'].some(key => Number(attr(crop, key, '0')) !== 0)) throw Error('Edited framed picture fill requires review');
  }
  if (!shadow || shadow.color.toLowerCase() !== frame.shadow.color.toLowerCase() || ['opacity','blur','offsetX','offsetY'].some(k => Math.abs(shadow[k] - frame.shadow[k]) > 0.0002) || !shadow.rotateWithShape) throw Error('Framed picture shadow differs from metadata');
}
