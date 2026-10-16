import test from 'node:test';
import assert from 'node:assert/strict';
import JSZip from 'jszip';
import {DOMParser, XMLSerializer} from '@xmldom/xmldom';
import {createDocument, createElement} from '../src/document.js';
import {exportPptx} from '../src/pptx.js';
import {readPptx} from '../src/pptx-import.js';
globalThis.DOMParser=DOMParser;globalThis.XMLSerializer=XMLSerializer;
for(const [fontSize,lineHeight]of [[11,1.2],[26,1.2],[28,1.5]]){
 test(`uniform ${fontSize}pt paragraph uses em-based spacing and round-trips ${lineHeight}`,async()=>{
  const d=createDocument('Synthetic point spacing');d.slides[0].elements.push(createElement('text',{fontSize,lineHeight,content:{type:'doc',content:[{type:'paragraph',content:[{type:'text',text:'First\nSecond',marks:[]}]}]}}));
  const bytes=await(await exportPptx(d)).arrayBuffer();const z=await JSZip.loadAsync(bytes);const xml=await z.file('ppt/slides/slide1.xml').async('string');
  assert.match(xml,new RegExp(`<a:spcPts val="${Math.round(fontSize*lineHeight*100)}"`));assert.doesNotMatch(xml,/<a:spcPct/);
  const imported=await readPptx(new Blob([bytes]));const e=imported.document.slides[0].elements.find(e=>e.type==='text');assert.equal(Number(e.content.content[0].attrs.lineHeight),lineHeight);
 });
}
test('mixed-size paragraphs retain the previous percentage spacing behavior',async()=>{
 const d=createDocument('Synthetic mixed spacing');d.slides[0].elements.push(createElement('text',{fontSize:20,lineHeight:1.2,content:{type:'doc',content:[{type:'paragraph',attrs:{lineHeight:'1.5'},content:[{type:'text',text:'Regular '},{type:'text',text:'Large',marks:[{type:'textStyle',attrs:{fontSize:'36pt'}}]}]}]}}));
 const z=await JSZip.loadAsync(await(await exportPptx(d)).arrayBuffer());assert.match(await z.file('ppt/slides/slide1.xml').async('string'),/<a:spcPct val="150000"/);
});
