import { describe, expect, it } from 'bun:test';

import { searchText } from './searchText';
import type { DeckSnapshot, ShapeSnapshot, StorySnapshot } from './types';

const STYLE = {
  bold: null,
  italic: null,
  fontSizePt: null,
  color: null,
  fontFamily: null,
  underline: null,
};

function story(id: string, paragraphs: readonly string[]): StorySnapshot {
  return {
    id,
    length: paragraphs.reduce((length, text) => length + text.length + 1, 0),
    paragraphs: paragraphs.map((text, index) => ({
      id: `${id}-p${index}`,
      alignment: null,
      level: 0,
      bulletJson: null,
      runs: [{ text, style: STYLE }],
    })),
  };
}

function shape(
  id: string,
  textStories: readonly StorySnapshot[] = [],
  children: readonly ShapeSnapshot[] = []
): ShapeSnapshot {
  return {
    id,
    sourceId: 1,
    kind: children.length > 0 ? 'group' : 'shape',
    name: id,
    x: 0,
    y: 0,
    width: 1,
    height: 1,
    rotationDeg: 0,
    flipH: false,
    flipV: false,
    geometry: 'rect',
    adjustValues: {},
    placeholder: null,
    fill: null,
    resolvedFillColor: null,
    outline: null,
    resolvedOutlineColor: null,
    mediaPartPath: null,
    graphic: null,
    textStories: [...textStories],
    children: [...children],
  };
}

describe('PPTX text search', () => {
  it('counts paragraph separators and visits nested group shapes in deck order', () => {
    const deck: DeckSnapshot = {
      widthEmu: 1,
      heightEmu: 1,
      slides: [
        {
          id: 'slide-1',
          sourcePartPath: null,
          layoutPartPath: null,
          name: null,
          shapes: [
            shape('top', [story('top-story', ['Query', 'query'])]),
            shape('group', [], [shape('nested', [story('nested-story', ['before query'])])]),
          ],
        },
      ],
    };

    expect(searchText(deck, 'query')).toEqual([
      {
        slideIndex: 0,
        slideId: 'slide-1',
        shapeId: 'top',
        storyId: 'top-story',
        start: 0,
        end: 5,
        text: 'Query',
      },
      {
        slideIndex: 0,
        slideId: 'slide-1',
        shapeId: 'top',
        storyId: 'top-story',
        start: 6,
        end: 11,
        text: 'query',
      },
      {
        slideIndex: 0,
        slideId: 'slide-1',
        shapeId: 'nested',
        storyId: 'nested-story',
        start: 7,
        end: 12,
        text: 'query',
      },
    ]);
  });
});
