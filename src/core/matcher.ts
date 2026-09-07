import { normalize } from './normalize'
import type { LintMatch, LintRule } from './types'

/** `text` 内にある `needle` の全出現位置（左から右へ、重複なしで走査）。 */
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
 * `matchStart` の位置のマッチが `exceptions`（`needle` を含む正当な長い語）の
 * いずれかの中に収まっている場合は true を返す。
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
    if (offsetInException === -1) continue // 不正な形式の例外語なので無視

    const windowStart = matchStart - offsetInException
    const windowEnd = windowStart + exception.length
    if (windowStart < 0 || windowEnd > text.length) continue

    if (text.slice(windowStart, windowEnd) === exception) return true
  }

  return false
}

/**
 * `text` の中から、各ルールの `wrong` 語の出現をすべて検出する。
 * 用語リストはせいぜい数百件程度なので、ルールごとに単純なindexOfループを
 * 回すだけでAho-Corasickより実装がシンプルで、速度的にも十分。
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
