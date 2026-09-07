import type { LintRule } from '../core/types'

/** One flagged location in the document, ready to be selected or fixed. */
export interface Occurrence {
  id: string
  ruleId: string
  wrong: string
  correct: string
  note?: string
  /** Short surrounding text so the user can judge context without leaving the pane. */
  contextBefore: string
  contextAfter: string
  /** e.g. "スライド 3". Omitted for Word, where paragraph numbers aren't meaningful to the user. */
  location?: string
}

/**
 * Word and PowerPoint expose very different document object models, but both
 * can support the same four operations, so the rest of the app never branches
 * on host.
 */
export interface HostAdapter {
  readonly hostName: string
  scan(rules: LintRule[]): Promise<Occurrence[]>
  selectOccurrence(occurrenceId: string): Promise<void>
  applyOccurrence(occurrenceId: string): Promise<void>
  /**
   * Fixes every occurrence of this rule found in the most recent scan().
   * Replays the locations scan() already stored, rather than re-searching
   * the document, so the fixed set always matches what the pane showed
   * (same exceptions applied, same headers/footers included).
   */
  applyAllForRule(ruleId: string): Promise<number>
  /** Releases tracked ranges held from the last scan(). */
  dispose(): Promise<void>
}
