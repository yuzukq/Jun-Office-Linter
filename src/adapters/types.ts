import type { LintRule } from '../core/types'

/** ドキュメント内で検出された1箇所。選択・修正できる状態にある。 */
export interface Occurrence {
  id: string
  ruleId: string
  wrong: string
  correct: string
  note?: string
  /** ペインを離れなくても文脈を判断できるよう、前後の短いテキストを持たせる。 */
  contextBefore: string
  contextAfter: string
  /** 例: "スライド 3"。Wordでは段落番号がユーザーにとって意味を持たないため省略する。 */
  location?: string
}

/**
 * Word と PowerPoint は文書オブジェクトモデルがまったく異なるが、どちらも
 * 同じ4つの操作をサポートできるようにしておくことで、アプリの他の部分が
 * ホストごとに分岐しなくて済む。
 */
export interface HostAdapter {
  readonly hostName: string
  scan(rules: LintRule[]): Promise<Occurrence[]>
  selectOccurrence(occurrenceId: string): Promise<void>
  applyOccurrence(occurrenceId: string): Promise<void>
  /**
   * 直近の scan() で見つかったこのルールの該当箇所をすべて修正する。
   * ドキュメントを再検索するのではなく scan() が既に保持している位置情報を
   * そのまま使うため、修正される内容は必ずペイン表示と一致する
   * （適用済みの例外リスト・ヘッダー/フッターも含めて同じ）。
   */
  applyAllForRule(ruleId: string): Promise<number>
  /** 直近の scan() で保持していたトラッキング対象のRangeを解放する。 */
  dispose(): Promise<void>
}
