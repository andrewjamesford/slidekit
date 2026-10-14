import test from 'node:test';
import assert from 'node:assert/strict';
import JSZip from 'jszip';
import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import { createDocument, createElement, paragraph, validateDocument } from '../src/document.js';
import { DocumentStore } from '../src/store.js';
import { exportPptx } from '../src/pptx.js';
import { readPptx } from '../src/pptx-import.js';
globalThis.DOMParser = DOMParser;
globalThis.XMLSerializer = XMLSerializer;

test('schema 1 upgrades without mutation and schema 2 validates Notes', () => {
  const old = createDocument(); old.schemaVersion = 1; delete old.slides[0].notes;
  const upgraded = validateDocument(old);
  assert.equal(upgraded.schemaVersion, 2); assert.equal(upgraded.slides[0].notes, '');
  assert.equal(old.schemaVersion, 1); assert.equal(old.slides[0].notes, undefined);
  upgraded.slides[0].notes = 4; assert.throws(() => validateDocument(upgraded));
});

test('notes stay with IDs through duplication, ordering, history and clearing', () => {
  const store = new DocumentStore(); const original = store.slideId;
  store.updateSlide(original, {notes:'Narration\n\nSecond paragraph'});
  const copy = store.duplicateSlide(original); assert.notEqual(copy,original);
  store.moveSlide(copy,0);
  assert.equal(store.doc.slides[0].notes,store.doc.slides[1].notes);
  store.updateSlide(copy,{notes:''}); store.undo();
  assert.equal(store.doc.slides[0].notes,'Narration\n\nSecond paragraph');
  store.redo(); assert.equal(store.doc.slides[0].notes,'');
  store.deleteSlide(copy); assert.equal(store.doc.slides[0].id,original);
  assert.equal(store.getDocument().slides[0].notes,'Narration\n\nSecond paragraph');
});

test('layer index moves are atomic, undoable and respect locks', () => {
  const store = new DocumentStore();
  const a=store.addElement('shape'),b=store.addElement('text'),c=store.addElement('shape');
  store.moveElement(a,2); assert.deepEqual(store.slide.elements.map(e=>e.id),[b,c,a]);
  store.undo(); assert.deepEqual(store.slide.elements.map(e=>e.id),[a,b,c]);
  store.updateElement(a,{locked:true}); const before=store.getDocument();
  store.moveElement(a,2); assert.deepEqual(store.getDocument(),before);
  assert.throws(()=>store.moveElement(c,99)); assert.deepEqual(store.getDocument(),before);
});

test('PPTX notes preserve unicode, paragraphs, emptiness and relationship targets', async () => {
  const store=new DocumentStore(); const first=store.slideId;
  store.updateSlide(first,{notes:'  A&B <safe> Grüße\n\nLast line  ',name:'First'});
  store.addElement('text',{content:paragraph('First')});
  const second=store.addSlide({name:'Silent',notes:''});
  store.addElement('text',{content:paragraph('Silent')});
  const third=store.duplicateSlide(first); store.moveSlide(third,0);
  const doc=store.getDocument();
  const zip=await JSZip.loadAsync(await (await exportPptx(doc)).arrayBuffer());
  const note=await zip.file('ppt/notesSlides/notesSlide1.xml').async('string');
  zip.remove('ppt/notesSlides/notesSlide1.xml');zip.file('ppt/notesSlides/notesSlide71.xml',note);
  const rel=await zip.file('ppt/slides/_rels/slide1.xml.rels').async('string');
  zip.file('ppt/slides/_rels/slide1.xml.rels',rel.replace('notesSlide1.xml','notesSlide71.xml'));
  const result=await readPptx(new Blob([await zip.generateAsync({type:'uint8array'})]));
  assert.deepEqual(result.document.slides.map(s=>s.notes),doc.slides.map(s=>s.notes));
  assert.equal(result.document.slides.at(-1).notes,'');
  assert.ok(result.warnings.some(w=>w.includes('Notes imported as plain text')));
});


test('generated PPTX has one paragraph property node before rich-text runs', async () => {
  const deck=createDocument();
  deck.slides[0].elements.push(createElement('text',{content:{type:'doc',content:[{type:'paragraph',content:[
    {type:'text',text:'First',marks:[{type:'bold'}]},{type:'text',text:' second',marks:[{type:'italic'}]}
  ]}]}}));
  const zip=await JSZip.loadAsync(await (await exportPptx(deck)).arrayBuffer());
  const xml=new DOMParser().parseFromString(await zip.file('ppt/slides/slide1.xml').async('string'),'application/xml');
  for(const p of Array.from(xml.getElementsByTagNameNS('http://schemas.openxmlformats.org/drawingml/2006/main','p'))) {
    const properties=Array.from(p.childNodes).filter(n=>n.localName==='pPr');
    assert.ok(properties.length<=1);
    if(properties.length) assert.equal(p.firstChild,properties[0]);
  }
});


test('generated Notes relationships and package targets are valid', async () => {
  const deck=createDocument();deck.slides[0].notes='First';
  const store=new DocumentStore(deck);store.addSlide({notes:''});
  const zip=await JSZip.loadAsync(await (await exportPptx(store.getDocument())).arrayBuffer());
  const {posix}=await import('node:path');
  for(const file of Object.values(zip.files).filter(f=>f.name.endsWith('.rels'))) {
    const xml=new DOMParser().parseFromString(await file.async('string'),'application/xml');
    const rels=Array.from(xml.documentElement.childNodes).filter(n=>n.localName==='Relationship');
    assert.equal(new Set(rels.map(r=>r.getAttribute('Id'))).size,rels.length);
    if(/^ppt\/slides\/_rels\//.test(file.name)) assert.equal(rels.filter(r=>r.getAttribute('Type').endsWith('/notesSlide')).length,1);
    const directory=posix.dirname(file.name).replace(/\/?_rels$/,'');
    for(const rel of rels) if(rel.getAttribute('TargetMode')!=='External') {
      const target=posix.normalize(posix.join(directory,rel.getAttribute('Target'))).replace(/^\//,'');
      assert.ok(zip.file(target),`${file.name} refers to missing ${target}`);
    }
  }
});

test('repeated identical Notes input does not consume another undo step', () => {
 const store = new DocumentStore(); const id = store.slideId;
 store.updateSlide(id, {notes: 'First'});
 store.updateSlide(id, {notes: 'Second'});
 store.updateSlide(id, {notes: 'Second'});
 store.undo(); assert.equal(store.slide.notes, 'First');
 store.redo(); assert.equal(store.slide.notes, 'Second');
});

for (const verticalAlign of ['top', 'middle', 'bottom']) test(`text vertical alignment ${verticalAlign} survives validation and PowerPoint round trip`, async () => {
 const deck=createDocument(); const text=createElement('text',{height:300,verticalAlign,content:paragraph('Alignment stays with the text box')});
 deck.slides[0].elements.push(text);
 assert.equal(validateDocument(deck).slides[0].elements[0].verticalAlign,verticalAlign);
 const blob=await exportPptx(deck);const imported=await readPptx(blob);
 const result=imported.document || imported.doc || imported;
 assert.equal(result.slides[0].elements.find(e=>e.type==='text').verticalAlign,verticalAlign);
 assert.ok(!imported.warnings?.some(w=>String(w).includes('文本垂直对齐已按顶部')));
});
test('invalid vertical alignment is rejected and legacy omission stays unchanged', () => {
 const deck=createDocument();deck.slides[0].elements.push(createElement('text'));
 assert.equal(validateDocument(deck).slides[0].elements[0].verticalAlign,undefined);
 deck.slides[0].elements[0].verticalAlign='untrusted';assert.throws(()=>validateDocument(deck));
});
