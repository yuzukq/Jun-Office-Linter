/** A single "wrong term -> suggested correction" rule the user maintains. */
export interface LintRule {
  id: string
  wrong: string
  correct: string
  note?: string
  /**
   * Longer strings that legitimately contain `wrong` as a substring and
   * should NOT be flagged. Japanese has no word boundaries, so a plain
   * substring match on a 2-char term will false-positive inside compounds.
   */
  exceptions?: string[]
}

/** A match found in a document, still relative to one host-specific text unit. */
export interface LintMatch {
  ruleId: string
  wrong: string
  correct: string
  note?: string
  start: number
  length: number
}
