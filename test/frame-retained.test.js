import test from 'node:test';
import assert from 'node:assert/strict';
import JSZip from 'jszip';
import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import { createDocument, createElement, paragraph, plainText, validateDocument } from '../src/document.js';
import { DocumentStore } from '../src/store.js';
import { renderElement } from '../src/render.js';
import { exportPptx } from '../src/pptx.js';
import { readPptx } from '../src/pptx-import.js';
globalThis.DOMParser = DOMParser; globalThis.XMLSerializer = XMLSerializer;
const sourceSlideId = '20000000-0000-4000-8000-000000000001';
const frame = { radius: 8, shadow: { color: '#000000', opacity: 0.15, blur: 20, offsetX: 0, offsetY: 4, rotateWithShape: true } };
const retained = { kind: 'legacy-hidden-bullet', sourceSlideId, sourceIndex: 6 };
const png = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+aGZkAAAAASUVORK5CYII=';
function imageDoc() { const d = createDocument(); d.assets.a = { data: png, width: 1, height: 1, name: 'Synthetic' }; d.slides[0].elements.push(createElement('image', { assetId: 'a', width: 100, height: 100, x: -20, frame })); return d; }
test('frame and retained fields validate atomically, without accepting CSS or conflicting paint', () => {
 const d = imageDoc(); assert.deepEqual(validateDocument(d).slides[0].elements[0].frame, frame);
 const store = new DocumentStore(d), id = store.slide.elements[0].id, before = store.getDocument();
 for (const bad of [{ ...frame, radius: '8px' }, { ...frame, radius: 129 }, { ...frame, css: 'position:fixed' }, { ...frame, shadow: { ...frame.shadow, blur: 129 } }, { ...frame, shadow: { ...frame.shadow, rotateWithShape: false } }]) {
  assert.throws(() => store.updateElement(id, { frame: bad })); assert.deepEqual(store.getDocument(), before);
 }
 assert.throws(() => store.updateElement(id, { shadow: frame.shadow }));
 const t = createDocument(); t.slides[0].elements.push(createElement('text', { retained, opacity: 1 })); assert.throws(() => validateDocument(t));
 t.slides[0].elements[0].opacity = 0; assert.doesNotThrow(() => validateDocument(t));
 for (const patch of [{ kind: 'unknown' }, { sourceIndex: -1 }, { sourceSlideId: 'bad' }, { css: true }]) { t.slides[0].elements[0].retained = { ...retained, ...patch }; assert.throws(() => validateDocument(t)); }
});
test('framed image paints its exact box shadow and retained text cannot participate in canvas gestures', () => {
 const old = globalThis.document;
 globalThis.document = { createElement: () => ({ style: {}, dataset: {}, children: [], append(n) { this.children.push(n); } }), createTextNode: text => ({ text }) };
 try {
  const d = imageDoc(), n = renderElement(d.slides[0].elements[0], d.assets);
  assert.equal(n.style.left, '-20px'); assert.equal(n.style.transform, 'none'); assert.equal(n.style.overflow, 'visible');
  assert.equal(n.children[0].style.borderRadius, '8px'); assert.equal(n.children[0].style.boxShadow, '0px 4px 20px rgba(0,0,0,0.15)');
  const t = renderElement(createElement('text', { retained, opacity: 1, content: paragraph('Retained') }), {});
  assert.equal(t.ariaHidden, 'true'); assert.equal(t.inert, true); assert.equal(t.style.opacity, '0'); assert.equal(t.dataset.locked, 'true'); assert.equal(t.style.pointerEvents, 'none');
 } finally { globalThis.document = old; }
});
for (const opacity of [0.15, 0.3]) test('native framed picture geometry and shadow round trip at alpha ' + opacity, async () => {
 const d = imageDoc(); d.slides[0].elements[0].frame.shadow.opacity = opacity;
 const blob = await exportPptx(d), result = await readPptx(blob), e = result.document.slides[0].elements.find(e => e.type === 'image');
 assert.deepEqual(e.frame, d.slides[0].elements[0].frame); assert.equal(e.fit, 'contain'); assert.ok(Math.abs(e.x + 20) < 0.001);
 const z = await JSZip.loadAsync(await blob.arrayBuffer()), xml = await z.file('ppt/slides/slide1.xml').async('string');
 assert.match(xml, /custGeom/); assert.match(xml, /wR="76200"/); assert.match(xml, /outerShdw/);
 z.file('ppt/slides/slide1.xml', xml.replace('wR="76200"', 'wR="76201"'));
 await assert.rejects(readPptx(new Blob([await z.generateAsync({ type: 'uint8array' })])), /geometry requires review/);
});
test('letterboxed framed pictures encode fill margins without stretching or flattening', async () => {
 const d = imageDoc(); d.slides[0].elements[0].height = 80;
 const blob = await exportPptx(d), imported = await readPptx(blob), e = imported.document.slides[0].elements.find(e => e.type === 'image');
 assert.equal(e.fit, 'contain'); assert.equal(e.height, 80); assert.deepEqual(e.frame, frame);
 const z = await JSZip.loadAsync(await blob.arrayBuffer()), p = 'ppt/slides/slide1.xml', xml = await z.file(p).async('string');
 assert.match(xml, /fillRect l="10000" t="0" r="10000" b="0"/);
 z.file(p, xml.replace('l="10000"', 'l="20000"'));
 await assert.rejects(readPptx(new Blob([await z.generateAsync({ type: 'uint8array' })])), /fill requires review/);
});
test('retained text keeps exact whitespace, Notes, history and an enforceable non-rendering tag in PPTX', async () => {
 const d = createDocument(); d.slides[0].notes = 'Only narration'; const e = createElement('text', { content: paragraph('  Retained source  '), retained, opacity: 0 }); d.slides[0].elements.push(e);
 const store = new DocumentStore(d); store.updateElement(e.id, { content: paragraph('Edited retained text') }); store.undo();
 assert.equal(plainText(store.slide.elements[0].content), '  Retained source  '); store.redo();
 assert.equal(store.slide.notes, 'Only narration'); assert.throws(() => store.updateElement(e.id, { opacity: 1 }));
 const blob = await exportPptx(store.getDocument()), result = await readPptx(blob), t = result.document.slides[0].elements.find(e => e.type === 'text');
 assert.deepEqual(t.retained, retained); assert.equal(t.opacity, 0); assert.equal(plainText(t.content), 'Edited retained text'); assert.equal(result.document.slides[0].notes, 'Only narration');
 const z = await JSZip.loadAsync(await blob.arrayBuffer()), p = 'ppt/slides/slide1.xml'; z.file(p, (await z.file(p).async('string')).replace(/<a:alpha val="0"\/>/g, '<a:alpha val="50000"/>'));
 await assert.rejects(readPptx(new Blob([await z.generateAsync({ type: 'uint8array' })])), /Invalid retained/);
});


test('zero radius and disabled frame shadow remain native-round-trip safe', async () => {
 const d = imageDoc(); d.slides[0].elements[0].frame.radius = 0; d.slides[0].elements[0].frame.shadow.opacity = 0;
 const r = await readPptx(await exportPptx(d)); assert.deepEqual(r.document.slides[0].elements[0].frame, d.slides[0].elements[0].frame);
});


test('retained XML-invalid source refuses export instead of replacing bytes', async () => {
 const d = createDocument(); d.slides[0].elements.push(createElement('text', { retained, opacity: 0, content: paragraph('Keep\u0000exact') }));
 await assert.rejects(exportPptx(d), /XML-invalid/); assert.equal(plainText(d.slides[0].elements[0].content), 'Keep\u0000exact');
});
