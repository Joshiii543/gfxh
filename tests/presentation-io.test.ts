import { describe, it, expect } from 'vitest';
import { buildPptx, buildOdp, parsePptx, parseOdp } from '../presentation-io.js';

const pixel = new Uint8Array([137, 80, 78, 71, 13, 10, 26, 10]);

const deck = [
  {
    bg: '#eeece5',
    layer: { mime: 'image/png', bytes: pixel },
    texts: [
      { text: 'Hallo Welt', x: 0.1, y: 0.2, w: 0.6, h: 0.1, fontPt: 40, color: '#171717', alpha: 1, bold: true, italic: false, family: 'Arial, Helvetica, sans-serif', align: 'l', wrap: false },
      { text: 'Zeile 1\nZeile & 2', x: 0.1, y: 0.5, w: 0.5, h: 0.2, fontPt: 20, color: '#ef3828', alpha: 1, bold: false, italic: false, family: 'Georgia, serif', align: 'ctr', wrap: false },
    ],
  },
  { bg: null, layer: null, texts: [{ text: 'Folie zwei', x: 0.1, y: 0.1, w: 0.8, h: 0.1, fontPt: 30, color: '#000000', alpha: 1, bold: false, italic: false, family: 'monospace', align: 'r', wrap: false }] },
];

describe('PPTX round trip', () => {
  it('keeps slide count, text, position and size', async () => {
    const bytes = await buildPptx(deck);
    const parsed = await parsePptx(bytes.buffer);
    expect(parsed.slides).toHaveLength(2);
    const texts = parsed.slides[0].elements.filter((e) => e.type === 'text');
    expect(texts.map((t) => t.text)).toEqual(['Hallo Welt', 'Zeile 1\nZeile & 2']);
    expect(texts[0].fontPt).toBeCloseTo(40, 1);
    expect(texts[0].x).toBeCloseTo(0.1, 3);
    expect(texts[0].y).toBeCloseTo(0.2, 3);
    expect(texts[0].bold).toBe(true);
    expect(texts[1].align).toBe('ctr');
    expect(texts[1].family).toBe('Georgia, serif');
    expect(parsed.slides[0].elements.some((e) => e.type === 'image')).toBe(true);
    expect(parsed.slides[1].bg).toBeNull();
  });
});

describe('ODP round trip', () => {
  it('keeps slide count, text and style', async () => {
    const bytes = await buildOdp(deck);
    const parsed = await parseOdp(bytes.buffer);
    expect(parsed.slides).toHaveLength(2);
    const texts = parsed.slides[0].elements.filter((e) => e.type === 'text');
    expect(texts.map((t) => t.text)).toEqual(['Hallo Welt', 'Zeile 1\nZeile & 2']);
    expect(texts[0].fontPt).toBeCloseTo(40, 1);
    expect(texts[0].bold).toBe(true);
    expect(parsed.slides[0].bg).toBe('#eeece5');
    expect(parsed.slides[1].elements[0].text).toBe('Folie zwei');
  });
});
