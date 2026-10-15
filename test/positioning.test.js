import test from 'node:test';
import assert from 'node:assert/strict';
import { createDocument, createElement, validateDocument } from '../src/document.js';
import { DocumentStore } from '../src/store.js';
import { positionElement, renderElement, usesCssPrimitive } from '../src/render.js';

test('fractional unrotated text uses layout coordinates without a raster transform', () => {
  const node = { style: {} };
  positionElement(node, { type: 'text', x: 137.25, y: 81.125, width: 423.5, height: 122.25, rotation: 0, opacity: 1 });
  assert.equal(node.style.left, '137.25px');
  assert.equal(node.style.top, '81.125px');
  assert.equal(node.style.transform, 'none');
  assert.equal(node.style.width, '423.5px');
});

test('text crossing zero rotation clears both previous positioning modes', () => {
  const node = { style: {} };
  const element = { type: 'text', x: 123, y: 87, width: 280, height: 160, opacity: 1 };
  positionElement(node, { ...element, rotation: 0 });
  positionElement(node, { ...element, rotation: 30 });
  assert.equal(node.style.left, '0');
  assert.equal(node.style.top, '0');
  assert.equal(node.style.transform, 'translate(123px, 87px) rotate(30deg)');
  positionElement(node, { ...element, rotation: 360 });
  assert.equal(node.style.left, '123px');
  assert.equal(node.style.top, '87px');
  assert.equal(node.style.transform, 'none');
});

test('non-text objects retain translation and rotation even when unrotated', () => {
  const node = { style: {} };
  positionElement(node, { type: 'shape', x: 62, y: 94, width: 180, height: 120, rotation: 0, opacity: 0.5 });
  assert.equal(node.style.left, '0');
  assert.equal(node.style.top, '0');
  assert.equal(node.style.transform, 'translate(62px, 94px) rotate(0deg)');
  assert.equal(node.style.opacity, '0.5');
});


test('text overflow is opt-in and survives JSON and history without rewriting old documents', () => {
  const old = createDocument(); old.slides[0].elements.push(createElement('text'));
  assert.equal(validateDocument(old).slides[0].elements[0].textOverflow, undefined);
  const store = new DocumentStore(); const id = store.addElement('text', { textOverflow: 'visible' });
  store.updateElement(id, { textOverflow: 'hidden' }); store.undo();
  assert.equal(store.slide.elements[0].textOverflow, 'visible');
  const restored = validateDocument(JSON.parse(JSON.stringify(store.getDocument())));
  assert.equal(restored.slides[0].elements[0].textOverflow, 'visible');
  store.redo(); assert.equal(store.slide.elements[0].textOverflow, 'hidden');
});

test('untrusted overflow CSS and overflow on non-text elements are rejected atomically', () => {
  const store = new DocumentStore(); const id = store.addElement('text'); const before = store.getDocument();
  for (const textOverflow of ['scroll', 'visible;position:fixed', null, true, 4]) {
    assert.throws(() => store.updateElement(id, { textOverflow }));
    assert.deepEqual(store.getDocument(), before);
  }
  const doc = createDocument(); doc.slides[0].elements.push(createElement('shape', { textOverflow: 'visible' }));
  assert.throws(() => validateDocument(doc));
});


test('flat CSS shapes use fractional layout coordinates and paint on the opacity wrapper', () => {
  const previous = globalThis.document;
  globalThis.document = { createElement: () => ({ style: {}, dataset: {}, children: [], append(child) { this.children.push(child); } }) };
  try {
    for (const shape of ['rect', 'ellipse']) {
      const element = { type: 'shape', shape, x: 4.125, y: 8.25, width: 110.5, height: 70.25, rotation: 0, opacity: 0.15, fill: '#123456', strokeWidth: 0 };
      const node = renderElement(element, {});
      assert.equal(node.style.left, '4.125px');
      assert.equal(node.style.top, '8.25px');
      assert.equal(node.style.transform, 'none');
      assert.equal(node.style.backgroundColor, '#123456');
      assert.equal(node.style.borderRadius, shape === 'ellipse' ? '50%' : '0');
      assert.equal(node.style.opacity, '0.15');
      assert.equal(node.children.length, 0);
      positionElement(node, { ...element, rotation: 15 });
      assert.equal(node.style.left, '0');
      assert.equal(node.style.transform, 'translate(4.125px, 8.25px) rotate(15deg)');
    }
  } finally { globalThis.document = previous; }
});

test('complex geometry and appearance keep the SVG renderer', () => {
  const flat = { type: 'shape', shape: 'ellipse', strokeWidth: 0 };
  assert.equal(usesCssPrimitive(flat), true);
  for (const patch of [
    { type: 'image' }, { shape: 'roundRect' }, { shape: 'line' },
    { strokeWidth: 1 }, { strokeWidth: undefined }, { fillOpacity: 0.5 },
    { gradient: { stops: [] } }, { shadow: {} },
  ]) assert.equal(usesCssPrimitive({ ...flat, ...patch }), false);
});
