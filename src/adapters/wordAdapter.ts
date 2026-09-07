import { findRawOffsets, isExceptionAtOffset } from '../core/matcher'
import { normalize } from '../core/normalize'
import type { LintRule } from '../core/types'
import type { HostAdapter, Occurrence } from './types'

const SNIPPET_RADIUS = 15

/**
 * Must stay a function, not a module-level constant: App.tsx imports both
 * adapters unconditionally, and a top-level `Word.HeaderFooterType.primary`
 * reference would run at import time even when the host is PowerPoint,
 * where the global `Word` namespace isn't loaded — crashing the whole
 * bundle before React ever mounts.
 */
function headerFooterTypes(): Word.HeaderFooterType[] {
  return [Word.HeaderFooterType.primary, Word.HeaderFooterType.firstPage, Word.HeaderFooterType.evenPages]
}

interface TrackedOccurrence {
  range: Word.Range
  correct: string
}

/**
 * Known gap: this scans the main body plus headers/footers, but not
 * footnotes/endnotes or text inside floating text boxes — Word's object
 * model doesn't surface those through `body.paragraphs`.
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
              // Should not happen (both are non-overlapping left-to-right
              // scans of the same text), but if it ever does, skip this rule
              // in this paragraph rather than pairing offsets with the wrong hit.
              console.warn('Jun Office Linter: search/offset count mismatch, skipping', {
                rule: rule.id,
                hits: hits.items.length,
                offsets: rawOffsets.length,
              })
              continue
            }

            // Both scans are non-overlapping and left-to-right over the same
            // text, so rawOffsets[i] always corresponds to hits.items[i].
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
        // underlying document range is gone; nothing to release
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
