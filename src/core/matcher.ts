import { normalize } from './normalize'
import type { LintMatch, LintRule } from './types'

/** All left-to-right, non-overlapping offsets of `needle` in `text`. */
export function findRawOffsets(text: string, needle: string): number[] {
  if (!needle) return []
  const offsets: number[] = []
  let fromIndex = 0
  while (true) {
    const start = text.indexOf(needle, fromIndex)
    if (start === -1) break
    offsets.push(start)
    fromIndex = start + needle.length
  }
  return offsets
}

/**
 * True if the match at `matchStart` sits inside one of `exceptions` — a
 * longer, legitimate string that happens to contain `needle`.
 */
export function isExceptionAtOffset(
  text: string,
  matchStart: number,
  needle: string,
  exceptions: string[] | undefined,
): boolean {
  if (!exceptions || exceptions.length === 0) return false

  for (const rawException of exceptions) {
    const exception = normalize(rawException)
    const offsetInException = exception.indexOf(needle)
    if (offsetInException === -1) continue // malformed exception, ignore

    const windowStart = matchStart - offsetInException
    const windowEnd = windowStart + exception.length
    if (windowStart < 0 || windowEnd > text.length) continue

    if (text.slice(windowStart, windowEnd) === exception) return true
  }

  return false
}

/**
 * Finds every occurrence of every rule's `wrong` term in `text`.
 * Term lists here run to at most a few hundred entries, so a plain
 * indexOf loop per rule is simpler than Aho-Corasick and fast enough.
 */
export function findMatches(text: string, rules: LintRule[]): LintMatch[] {
  const normalizedText = normalize(text)
  const matches: LintMatch[] = []

  for (const rule of rules) {
    if (!rule.wrong) continue
    const needle = normalize(rule.wrong)

    for (const start of findRawOffsets(normalizedText, needle)) {
      if (isExceptionAtOffset(normalizedText, start, needle, rule.exceptions)) continue

      matches.push({
        ruleId: rule.id,
        wrong: rule.wrong,
        correct: rule.correct,
        note: rule.note,
        start,
        length: needle.length,
      })
    }
  }

  return matches.sort((a, b) => a.start - b.start)
}
