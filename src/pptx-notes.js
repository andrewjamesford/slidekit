import { at, attr, children } from './pptx-package.js';

/** Read only speaker body placeholders. Preserve empty paragraphs and soft breaks. */
export function readSpeakerNotes(root) {
  const bodies = children(at(root, 'cSld/spTree'), 'sp').filter(shape => attr(at(shape, 'nvSpPr/nvPr/ph'), 'type') === 'body');
  return bodies.flatMap(shape => children(at(shape, 'txBody'), 'p').map(paragraph => children(paragraph).map(node => {
    if (node.localName === 'br') return '\n';
    if (['r', 'fld'].includes(node.localName)) return at(node, 't')?.textContent || '';
    return '';
  }).join(''))).join('\n');
}
