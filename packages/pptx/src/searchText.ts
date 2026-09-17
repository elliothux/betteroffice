import type {
  DeckSnapshot,
  PptxTextMatch,
  PptxTextSearchOptions,
  ShapeSnapshot,
  StorySnapshot,
} from './types';

function patternFor(query: string, caseSensitive: boolean): RegExp {
  const literal = query.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  return new RegExp(literal, caseSensitive ? 'gu' : 'giu');
}

function* descendants(shapes: readonly ShapeSnapshot[]): Generator<ShapeSnapshot> {
  for (const shape of shapes) {
    yield shape;
    yield* descendants(shape.children);
  }
}

function storyText(story: StorySnapshot): string {
  return story.paragraphs
    .map((paragraph) => paragraph.runs.map((run) => run.text).join(''))
    .join('\n');
}

export function searchText(
  deck: DeckSnapshot,
  query: string,
  options: PptxTextSearchOptions = {}
): PptxTextMatch[] {
  if (!query) return [];
  const limit = options.limit ?? Number.POSITIVE_INFINITY;
  if ((!Number.isSafeInteger(limit) && limit !== Number.POSITIVE_INFINITY) || limit < 0) {
    throw new RangeError('search limit must be a non-negative safe integer');
  }
  if (limit === 0) return [];

  const results: PptxTextMatch[] = [];
  const pattern = patternFor(query, options.caseSensitive ?? false);
  for (const [slideIndex, slide] of deck.slides.entries()) {
    for (const shape of descendants(slide.shapes)) {
      for (const story of shape.textStories) {
        const text = storyText(story);
        pattern.lastIndex = 0;
        for (const match of text.matchAll(pattern)) {
          const start = match.index;
          results.push({
            slideIndex,
            slideId: slide.id,
            shapeId: shape.id,
            storyId: story.id,
            start,
            end: start + match[0].length,
            text: match[0],
          });
          if (results.length >= limit) return results;
        }
      }
    }
  }
  return results;
}
