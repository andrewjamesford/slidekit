import test from 'node:test';
import assert from 'node:assert/strict';
import JSZip from 'jszip';
import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import { createDocument } from '../src/document.js';
import { exportPptx } from '../src/pptx.js';
import { readPptx } from '../src/pptx-import.js';
import { at, attr, children } from '../src/pptx-package.js';
globalThis.DOMParser = DOMParser;
globalThis.XMLSerializer = XMLSerializer;
for (const notes of ['first\nsecond', '\nfirst\n\nlast\n', '  leading\ntrailing  ', 'café Ω 日本語\n& < > "', '', 'single line']) {
  test('native Notes encode real paragraphs and preserve text: ' + JSON.stringify(notes), async () => {
    const doc = createDocument('Synthetic native Notes'); doc.slides[0].notes = notes;
    const blob = await exportPptx(doc), zip = await JSZip.loadAsync(await blob.arrayBuffer());
    const xml = new DOMParser().parseFromString(await zip.file('ppt/notesSlides/notesSlide1.xml').async('string'), 'application/xml');
    const shape = children(at(xml.documentElement, 'cSld/spTree'), 'sp').find(shape => attr(at(shape, 'nvSpPr/nvPr/ph'), 'type') === 'body');
    const paragraphs = children(at(shape, 'txBody'), 'p');
    assert.equal(paragraphs.length, notes.split('\n').length);
    assert.deepEqual(paragraphs.map(p => at(p, 'r/t')?.textContent || ''), notes.split('\n'));
    assert.ok(paragraphs.every(p => !at(p, 'r/t')?.textContent.includes('\n')));
    const imported = await readPptx(blob); assert.equal(imported.document.slides[0].notes, notes);
  });
}
