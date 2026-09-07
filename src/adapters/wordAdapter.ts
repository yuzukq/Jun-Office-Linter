import { findRawOffsets, isExceptionAtOffset } from '../core/matcher'
import { normalize } from '../core/normalize'
import type { LintRule } from '../core/types'
import type { HostAdapter, Occurrence } from './types'

const SNIPPET_RADIUS = 15

/**
 * モジュール直下の定数ではなく、必ず関数のままにしておくこと。App.tsx は
 * 両方のアダプタを無条件にimportしているため、トップレベルで
 * `Word.HeaderFooterType.primary` を参照してしまうと、ホストがPowerPointで
 * グローバルな `Word` 名前空間が読み込まれていない場合にimport時点で
 * 例外が発生し、Reactがマウントされる前にバンドル全体がクラッシュしてしまう。
 */
function headerFooterTypes(): Word.HeaderFooterType[] {
  return [Word.HeaderFooterType.primary, Word.HeaderFooterType.firstPage, Word.HeaderFooterType.evenPages]
}

interface TrackedOccurrence {
  range: Word.Range
  correct: string
}

/**
 * 既知の制限: 本文とヘッダー/フッターはスキャンするが、脚注・文末脚注・
 * floatingテキストボックス内のテキストはスキャン対象外。Wordのオブジェクト
 * モデルではこれらが `body.paragraphs` に出てこないため。
 */
export class WordAdapter implements HostAdapter {
  readonly hostName = 'Word'
  private tracked = new Map<string, TrackedOccurrence>()

  async scan(rules: LintRule[]): Promise<Occurrence[]> {
    await this.dispose()

    return Word.run(async (context) => {
      const bodies: Word.Body[] = [context.document.body]

      const sections = context.document.sections
      sections.load('items')
      await context.sync()

      for (const section of sections.items) {
        for (const type of headerFooterTypes()) {
          bodies.push(section.getHeader(type))
          bodies.push(section.getFooter(type))
        }
      }

      const occurrences: Occurrence[] = []
      const countPerRule = new Map<string, number>()

      for (const body of bodies) {
        const paragraphs = body.paragraphs
        paragraphs.load('items')
        await context.sync()

        for (const paragraph of paragraphs.items) {
          paragraph.load('text')
        }
        await context.sync()

        for (const paragraph of paragraphs.items) {
          const text = normalize(paragraph.text)
          if (!text) continue

          for (const rule of rules) {
            if (!rule.wrong) continue
            const needle = normalize(rule.wrong)
            const rawOffsets = findRawOffsets(text, needle)
            if (rawOffsets.length === 0) continue

            const paragraphRange = paragraph.getRange()
            const hits = paragraphRange.search(rule.wrong, { matchCase: true })
            hits.load('items')
            await context.sync()

            if (hits.items.length !== rawOffsets.length) {
              // 本来起こらないはず（どちらも同じテキストを左から右へ重複なく
              // 走査しているだけ）だが、もし発生した場合は誤ったヒットと
              // オフセットを対応付けてしまわないよう、この段落・ルールをスキップする。
              console.warn('Jun Office Linter: search/offset count mismatch, skipping', {
                rule: rule.id,
                hits: hits.items.length,
                offsets: rawOffsets.length,
              })
              continue
            }

            // どちらも同じテキストを左から右へ重複なく走査しているだけなので、
            // rawOffsets[i] は必ず hits.items[i] に対応する。
            hits.items.forEach((hitRange, i) => {
              const start = rawOffsets[i]
              if (isExceptionAtOffset(text, start, needle, rule.exceptions)) return

              const count = countPerRule.get(rule.id) ?? 0
              countPerRule.set(rule.id, count + 1)

              const id = `${rule.id}--${count}`
              hitRange.track()
              this.tracked.set(id, { range: hitRange, correct: rule.correct })

              occurrences.push({
                id,
                ruleId: rule.id,
                wrong: rule.wrong,
                correct: rule.correct,
                note: rule.note,
                ...buildSnippet(text, start, needle.length),
              })
            })
          }
        }
      }

      return occurrences
    })
  }

  async selectOccurrence(occurrenceId: string): Promise<void> {
    const entry = this.tracked.get(occurrenceId)
    if (!entry) return
    await Word.run(entry.range, async (context) => {
      entry.range.select()
      await context.sync()
    })
  }

  async applyOccurrence(occurrenceId: string): Promise<void> {
    const entry = this.tracked.get(occurrenceId)
    if (!entry) return
    await Word.run(entry.range, async (context) => {
      entry.range.insertText(entry.correct, Word.InsertLocation.replace)
      await context.sync()
    })
    this.tracked.delete(occurrenceId)
  }

  async applyAllForRule(ruleId: string): Promise<number> {
    const prefix = `${ruleId}--`
    const entries = [...this.tracked.entries()].filter(([id]) => id.startsWith(prefix))
    if (entries.length === 0) return 0

    const ranges = entries.map(([, entry]) => entry.range)
    await Word.run(ranges, async (context) => {
      for (const [, entry] of entries) {
        entry.range.insertText(entry.correct, Word.InsertLocation.replace)
      }
      await context.sync()
    })

    for (const [id] of entries) this.tracked.delete(id)
    return entries.length
  }

  async dispose(): Promise<void> {
    if (this.tracked.size === 0) return
    const entries = [...this.tracked.values()]
    this.tracked.clear()
    for (const entry of entries) {
      try {
        await Word.run(entry.range, async (context) => {
          entry.range.untrack()
          await context.sync()
        })
      } catch {
        // 元のドキュメントのRangeが既に消えているので、解放するものは何もない
      }
    }
  }
}

function buildSnippet(text: string, start: number, length: number): { contextBefore: string; contextAfter: string } {
  return {
    contextBefore: text.slice(Math.max(0, start - SNIPPET_RADIUS), start),
    contextAfter: text.slice(start + length, start + length + SNIPPET_RADIUS),
  }
}
