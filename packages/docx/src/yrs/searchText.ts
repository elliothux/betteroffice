import type { YrsParagraph, YrsTextMatch, YrsTextSearchOptions } from './index';

function patternFor(query: string, caseSensitive: boolean): RegExp {
  const literal = query.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  return new RegExp(literal, caseSensitive ? 'gu' : 'giu');
}

export function searchText(
  stories: readonly string[],
  paragraphs: (story: string) => readonly YrsParagraph[],
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
    for (const paragraph of paragraphs(story)) {
      pattern.lastIndex = 0;
      for (const match of paragraph.text.matchAll(pattern)) {
        const start = match.index;
        matches.push({
          story,
          paraId: paragraph.paraId,
          start,
          end: start + match[0].length,
          text: match[0],
        });
        if (matches.length >= limit) return matches;
      }
    }
  }
  return matches;
}
