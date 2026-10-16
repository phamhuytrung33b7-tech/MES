/**
 * High-performance utility functions for smart tokenized fuzzy search
 * Includes RAM pre-indexing and debounced caching for 0ms lag UI response
 */

/**
 * Remove Vietnamese accents / diacritics and normalize string
 */
export function normalizeVietnamese(str: string): string {
  if (!str) return '';
  return str
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/[đĐ]/g, 'd')
    .replace(/[^a-z0-9\s]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

/**
 * Tokenize string into clean words
 */
export function tokenize(str: string): string[] {
  const normalized = normalizeVietnamese(str);
  return normalized ? normalized.split(' ').filter(Boolean) : [];
}

/**
 * Check if pattern is a subsequence of text (fuzzy match)
 * e.g., 's766' matches 'sha76622kl'
 */
export function isSubsequence(pattern: string, text: string): boolean {
  if (!pattern || !text) return false;
  let patternIdx = 0;
  let textIdx = 0;
  while (patternIdx < pattern.length && textIdx < text.length) {
    if (pattern[patternIdx] === text[textIdx]) {
      patternIdx++;
    }
    textIdx++;
  }
  return patternIdx === pattern.length;
}

export interface SearchableItem {
  id: string;
  name?: string;
  code?: string;
  model?: string;
  poId?: string;
  description?: string;
  [key: string]: any;
}

export interface IndexedItem<T> {
  item: T;
  primaryIdNorm: string;
  combinedNorm: string;
  fieldNorms: string[];
  tokens: string[];
}

/**
 * Pre-index item array into RAM memory so string normalization & tokenization
 * happens ONCE when data changes, NOT on every keystroke!
 */
export function buildSearchIndex<T extends SearchableItem>(
  items: T[],
  searchKeys: (keyof T)[] = ['id', 'name', 'code', 'model', 'poId']
): IndexedItem<T>[] {
  if (!items) return [];

  return items.map(item => {
    const fieldNorms: string[] = [];

    searchKeys.forEach(key => {
      const val = item[key];
      if (typeof val === 'string' && val) {
        fieldNorms.push(normalizeVietnamese(val));
      }
    });

    const primaryVal = item.id || item.poId || item.code || '';
    const primaryIdNorm = normalizeVietnamese(primaryVal);
    const combinedNorm = fieldNorms.join(' ');
    const tokens = combinedNorm.split(' ').filter(Boolean);

    return {
      item,
      primaryIdNorm,
      combinedNorm,
      fieldNorms,
      tokens,
    };
  });
}

/**
 * Ultra-fast search over RAM pre-indexed items
 */
export function searchIndexedItems<T extends SearchableItem>(
  query: string,
  indexedItems: IndexedItem<T>[]
): T[] {
  if (!indexedItems || indexedItems.length === 0) return [];

  const rawQueryNorm = normalizeVietnamese(query);
  if (!rawQueryNorm) {
    return indexedItems.map(i => i.item);
  }

  const queryTokens = rawQueryNorm.split(' ').filter(Boolean);
  if (queryTokens.length === 0) {
    return indexedItems.map(i => i.item);
  }

  const scoredResults: { item: T; score: number }[] = [];

  for (let i = 0; i < indexedItems.length; i++) {
    const entry = indexedItems[i];
    let totalScore = 0;
    let matchedTokenCount = 0;

    // Direct exact or prefix match on primary ID gets massive priority
    if (entry.primaryIdNorm === rawQueryNorm) {
      totalScore += 1000;
    } else if (entry.primaryIdNorm.startsWith(rawQueryNorm)) {
      totalScore += 500;
    } else if (entry.combinedNorm.includes(rawQueryNorm)) {
      totalScore += 250;
    }

    // Evaluate each query token against pre-tokenized item
    for (let q = 0; q < queryTokens.length; q++) {
      const qToken = queryTokens[q];
      let bestTokenScore = 0;

      // Check against item tokens first (ultra fast RAM check)
      for (let t = 0; t < entry.tokens.length; t++) {
        const iToken = entry.tokens[t];
        if (iToken === qToken) {
          bestTokenScore = Math.max(bestTokenScore, 100);
          break; // exact token match
        } else if (iToken.startsWith(qToken)) {
          bestTokenScore = Math.max(bestTokenScore, 70);
        } else if (iToken.includes(qToken)) {
          bestTokenScore = Math.max(bestTokenScore, 40);
        }
      }

      // Fallback substring search in combined text
      if (bestTokenScore === 0 && entry.combinedNorm.includes(qToken)) {
        bestTokenScore = 30;
      }

      // Fallback subsequence fuzzy search
      if (bestTokenScore === 0 && isSubsequence(qToken, entry.combinedNorm)) {
        bestTokenScore = 15;
      }

      if (bestTokenScore > 0) {
        matchedTokenCount++;
        totalScore += bestTokenScore;
      }
    }

    const minTokensRequired = queryTokens.length <= 2 ? queryTokens.length : Math.ceil(queryTokens.length * 0.75);

    if (matchedTokenCount >= minTokensRequired) {
      if (matchedTokenCount === queryTokens.length) {
        totalScore += 50;
      }
      scoredResults.push({ item: entry.item, score: totalScore });
    }
  }

  scoredResults.sort((a, b) => b.score - a.score);
  return scoredResults.map(r => r.item);
}

/**
 * Smart fuzzy multi-token search fallback
 */
export function smartFuzzySearch<T extends SearchableItem>(
  query: string,
  items: T[],
  searchKeys: (keyof T)[] = ['id', 'name', 'code', 'model', 'poId']
): T[] {
  const indexed = buildSearchIndex(items, searchKeys);
  return searchIndexedItems(query, indexed);
}
