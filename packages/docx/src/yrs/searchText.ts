import type { YrsStorySegment, YrsTextMatch, YrsTextSearchOptions } from './index';

function patternFor(query: string, caseSensitive: boolean): RegExp {
  const literal = query.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  return new RegExp(literal, caseSensitive ? 'gu' : 'giu');
}

export function searchText(
  stories: readonly string[],
  storySegments: (story: string) => readonly YrsStorySegment[],
  query: string,
  options: YrsTextSearchOptions = {}
): YrsTextMatch[] {
  if (!query) return [];
  const limit = options.limit ?? Number.POSITIVE_INFINITY;
  if ((!Number.isSafeInteger(limit) && limit !== Number.POSITIVE_INFINITY) || limit < 0) {
    throw new RangeError('search limit must be a non-negative safe integer');
  }
  if (limit === 0) return [];

  const matches: YrsTextMatch[] = [];
  const pattern = patternFor(query, options.caseSensitive ?? false);
  for (const story of stories) {
    let offset = 0;
    let runStart = 0;
    let run = '';
    const runs: Array<{ start: number; text: string }> = [];
    const finishRun = () => {
      if (run) runs.push({ start: runStart, text: run });
      run = '';
    };

    for (const segment of storySegments(story)) {
      if (segment.kind === 'text') {
        if (!run) runStart = offset;
        run += segment.text;
        offset += segment.text.length;
        continue;
      }
      finishRun();
      if (segment.kind === 'embed') {
        offset += 1;
        continue;
      }
      for (const searchable of runs) {
        pattern.lastIndex = 0;
        for (const match of searchable.text.matchAll(pattern)) {
          const start = searchable.start + match.index;
          matches.push({
            story,
            paraId: segment.paraId,
            start,
            end: start + match[0].length,
            text: match[0],
          });
          if (matches.length >= limit) return matches;
        }
      }
      runs.length = 0;
      offset = 0;
    }
  }
  return matches;
}
