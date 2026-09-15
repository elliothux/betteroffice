import { readFileSync } from 'node:fs';
import { describe, expect, test } from 'bun:test';

import { openPresentation } from './index.js';

const fixture = readFileSync(
  new URL('../../apps/demo/public/betteroffice-demo.pptx', import.meta.url)
);
const font = readFileSync(
  new URL('../../crates/pptx-raster/tests/assets/Carlito-Regular.ttf', import.meta.url)
);

describe('@betteroffice/pptx-node', () => {
  test('opens, inspects, renders, and saves a presentation', async () => {
    const presentation = await openPresentation(fixture);
    const snapshot = presentation.snapshot();

    expect(snapshot.slides.length).toBeGreaterThan(0);
    expect(presentation.slideCount).toBe(snapshot.slides.length);
    expect(presentation.slideIds).toEqual(snapshot.slides.map((slide) => slide.id));
    expect(presentation.slide(0).id).toBe(snapshot.slides[0].id);
    expect(presentation.canUndo).toBe(false);
    presentation.registerFont({ family: 'Arial', data: font });
    presentation.registerFont({ family: 'Calibri', data: font });
    const rendered = await presentation.renderSlide(0);

    expect(rendered.data.subarray(0, 8)).toEqual(Buffer.from([137, 80, 78, 71, 13, 10, 26, 10]));
    expect((await presentation.save()).subarray(0, 2)).toEqual(Buffer.from('PK'));
  });

  test('guards collaboration methods on standalone presentations', async () => {
    const presentation = await openPresentation(fixture);
    expect(() => presentation.encodeStateVector()).toThrow('collaborative presentation');
  });
});
