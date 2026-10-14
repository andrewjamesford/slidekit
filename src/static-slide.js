import { validateDocument } from './document.js';
import { dom, renderSlide } from './render.js';

/** Mount a chrome-free, immutable slide with the same DOM renderer as the editor.
 * @param {HTMLElement} container
 * @param {import('./types.js').Deck} input
 * @param {string} slideId
 * @returns {{element: HTMLElement, ready: Promise<void>, destroy: () => void}}
 */
export function mountStaticSlide(container, input, slideId) {
  const deck = validateDocument(input);
  const slide = deck.slides.find(item => item.id === slideId);
  if (!slide) throw new TypeError('Unknown slide ID');
  const root = dom('div', 'slidekit sk-static');
  root.style.cssText = `width:${deck.width}px;height:${deck.height}px;min-height:0;`;
  const element = renderSlide(slide, deck);
  const world = dom('div', 'sk-world');
  world.style.transform = 'scale(1)';
  world.append(element);
  root.append(world);
  container.replaceChildren(root);
  const ready = Promise.all([
    document.fonts.ready,
    ...Array.from(root.querySelectorAll('img'), image => image.decode()),
  ]).then(() => undefined);
  return { element, ready, destroy: () => root.remove() };
}
