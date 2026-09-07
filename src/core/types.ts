/** ユーザーが管理する「誤変換語 → 修正候補」のルール1件分。 */
export interface LintRule {
  id: string
  wrong: string
  correct: string
  note?: string
  /**
   * `wrong` を部分文字列として正当に含む、除外すべき長い語。
   * 日本語には単語境界がないため、2文字程度の語をそのまま部分一致させると
   * 複合語の中で誤検出（false positive）が起きてしまう。
   */
  exceptions?: string[]
}

/** ドキュメント内で見つかった1件のマッチ。まだホスト固有のテキスト単位内でのオフセット。 */
export interface LintMatch {
  ruleId: string
  wrong: string
  correct: string
  note?: string
  start: number
  length: number
}
