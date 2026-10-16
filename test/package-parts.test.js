import test from 'node:test';
import assert from 'node:assert/strict';
import JSZip from 'jszip';
import {DOMParser, XMLSerializer} from '@xmldom/xmldom';
import {createDocument, emptySlide} from '../src/document.js';
import {exportPptx} from '../src/pptx.js';
globalThis.DOMParser=DOMParser;globalThis.XMLSerializer=XMLSerializer;
test('multi-slide native export declares only existing package parts',async()=>{
 const d=createDocument('Synthetic package declarations');
 for(let i=1;i<13;i++)d.slides.push(emptySlide());
 d.slides[12].hidden=true;
 const z=await JSZip.loadAsync(await(await exportPptx(d)).arrayBuffer());
 const types=new DOMParser().parseFromString(await z.file('[Content_Types].xml').async('string'),'application/xml');
 const overrides=Array.from(types.getElementsByTagNameNS('http://schemas.openxmlformats.org/package/2006/content-types','Override'));
 assert(overrides.length>13);
 for(const o of overrides)assert(z.file(o.getAttribute('PartName').slice(1)),o.getAttribute('PartName'));
 assert.equal(Object.keys(z.files).filter(n=>/^ppt\/slides\/slide\d+\.xml$/.test(n)).length,13);
});
