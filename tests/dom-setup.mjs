import { readFileSync } from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { JSDOM } from 'jsdom';

const here = path.dirname(fileURLToPath(import.meta.url));

/**
 * Installs a jsdom document as the global DOM, loaded from a fixture.
 *
 * jsdom has no layout engine, so getBoundingClientRect is stubbed to report a
 * plausible box for anything not display:none. Visibility filtering is
 * therefore exercised, but real geometry is only verifiable on the live page.
 */
export function loadFixture(name, url = 'https://outlook.office.com/mail/') {
  const html = readFileSync(path.join(here, 'fixtures', name), 'utf8');
  const dom = new JSDOM(`<!doctype html><html><body>${html}</body></html>`, {
    url,
    pretendToBeVisual: true,
  });

  const { window } = dom;

  window.Element.prototype.getBoundingClientRect = function getBoundingClientRect() {
    const style = window.getComputedStyle(this);
    const hidden = style.display === 'none' || style.visibility === 'hidden';
    const w = hidden ? 0 : 200;
    const h = hidden ? 0 : 24;
    return { x: 0, y: 0, top: 0, left: 0, right: w, bottom: h, width: w, height: h, toJSON: () => ({}) };
  };

  for (const key of [
    'window',
    'document',
    'location',
    'Element',
    'HTMLElement',
    'HTMLInputElement',
    'HTMLTextAreaElement',
    'Node',
    'NodeFilter',
    'CSS',
    'getComputedStyle',
    'MutationObserver',
  ]) {
    globalThis[key] = window[key] ?? window[key === 'getComputedStyle' ? 'getComputedStyle' : key];
  }
  globalThis.getComputedStyle = window.getComputedStyle.bind(window);

  return dom;
}

export async function lib() {
  return import('./.build/lib.mjs');
}
