import { findMatches } from '../core/matcher'
import { normalize } from '../core/normalize'
import type { LintRule } from '../core/types'
import type { HostAdapter, Occurrence } from './types'

const SNIPPET_RADIUS = 15

/**
 * Shape.textFrame throws InvalidArgument for shape types that can't hold a
 * text frame (pictures, charts, tables, media, ...). Filtering by type
 * before ever touching `.textFrame` keeps one chart on a slide from
 * crashing the whole scan.
 */
const TEXT_CAPABLE_SHAPE_TYPES = new Set<string>([
  'TextBox',
  'GeometricShape',
  'Placeholder',
  'Callout',
  'Freeform',
  'Graphic',
])

/**
 * Unlike Word.Range, PowerPoint.TextRange has no track()/untrack() — it
 * can't survive past the PowerPoint.run() batch that created it. So instead
 * of holding a live range, each occurrence remembers where to find it again
 * (slide + shape id, character offset) and re-derives the TextRange fresh
 * on every select/apply call.
 */
interface OccurrenceLocation {
  ruleId: string
  slideId: string
  shapeId: string
  start: number
  length: number
  correct: string
}

/**
 * Known gaps: speaker notes, grouped shapes nested more than one level, and
 * table cell text are not scanned (tables expose text per-cell, not through
 * a shape-level textFrame).
 */
export class PowerPointAdapter implements HostAdapter {
  readonly hostName = 'PowerPoint'
  private locations = new Map<string, OccurrenceLocation>()

  async scan(rules: LintRule[]): Promise<Occurrence[]> {
    this.locations.clear()

    return PowerPoint.run(async (context) => {
      const slides = context.presentation.slides
      slides.load('items/id')
      await context.sync()

      const occurrences: Occurrence[] = []
      const countPerRule = new Map<string, number>()

      for (let slideIndex = 0; slideIndex < slides.items.length; slideIndex++) {
        const slide = slides.items[slideIndex]
        const shapes = slide.shapes
        shapes.load('items/id,items/type')
        await context.sync()

        const textCapableShapes = shapes.items.filter((shape) => TEXT_CAPABLE_SHAPE_TYPES.has(shape.type))
        for (const shape of textCapableShapes) {
          shape.textFrame.load('hasText')
        }
        await context.sync()

        const shapesWithText = textCapableShapes.filter((shape) => shape.textFrame.hasText)
        for (const shape of shapesWithText) {
          shape.textFrame.textRange.load('text')
        }
        await context.sync()

        for (const shape of shapesWithText) {
          const text = normalize(shape.textFrame.textRange.text)
          if (!text) continue

          for (const match of findMatches(text, rules)) {
            const count = countPerRule.get(match.ruleId) ?? 0
            countPerRule.set(match.ruleId, count + 1)
            const id = `${match.ruleId}--${count}`

            this.locations.set(id, {
              ruleId: match.ruleId,
              slideId: slide.id,
              shapeId: shape.id,
              start: match.start,
              length: match.length,
              correct: match.correct,
            })

            occurrences.push({
              id,
              ruleId: match.ruleId,
              wrong: match.wrong,
              correct: match.correct,
              note: match.note,
              ...buildSnippet(text, match.start, match.length),
              location: `スライド ${slideIndex + 1}`,
            })
          }
        }
      }

      return occurrences
    })
  }

  async selectOccurrence(occurrenceId: string): Promise<void> {
    const loc = this.locations.get(occurrenceId)
    if (!loc) return
    await PowerPoint.run(async (context) => {
      this.resolveRange(context, loc).setSelected()
      await context.sync()
    })
  }

  async applyOccurrence(occurrenceId: string): Promise<void> {
    const loc = this.locations.get(occurrenceId)
    if (!loc) return
    await PowerPoint.run(async (context) => {
      this.resolveRange(context, loc).text = loc.correct
      await context.sync()
    })
    this.locations.delete(occurrenceId)
  }

  async applyAllForRule(ruleId: string): Promise<number> {
    const prefix = `${ruleId}--`
    const entries = [...this.locations.entries()].filter(([id]) => id.startsWith(prefix))
    if (entries.length === 0) return 0

    // Group by shape and apply right-to-left within each shape, so an
    // earlier edit in the same shape never invalidates a later offset.
    const byShape = new Map<string, { slideId: string; shapeId: string; hits: OccurrenceLocation[] }>()
    for (const [, loc] of entries) {
      const key = `${loc.slideId}::${loc.shapeId}`
      const group = byShape.get(key) ?? { slideId: loc.slideId, shapeId: loc.shapeId, hits: [] }
      group.hits.push(loc)
      byShape.set(key, group)
    }

    await PowerPoint.run(async (context) => {
      for (const group of byShape.values()) {
        const textRange = context.presentation.slides.getItem(group.slideId).shapes.getItem(group.shapeId)
          .textFrame.textRange
        const sorted = [...group.hits].sort((a, b) => b.start - a.start)
        for (const hit of sorted) {
          textRange.getSubstring(hit.start, hit.length).text = hit.correct
        }
      }
      await context.sync()
    })

    for (const [id] of entries) this.locations.delete(id)
    return entries.length
  }

  async dispose(): Promise<void> {
    this.locations.clear()
  }

  private resolveRange(context: PowerPoint.RequestContext, loc: OccurrenceLocation): PowerPoint.TextRange {
    const slide = context.presentation.slides.getItem(loc.slideId)
    const shape = slide.shapes.getItem(loc.shapeId)
    return shape.textFrame.textRange.getSubstring(loc.start, loc.length)
  }
}

function buildSnippet(text: string, start: number, length: number): { contextBefore: string; contextAfter: string } {
  return {
    contextBefore: text.slice(Math.max(0, start - SNIPPET_RADIUS), start),
    contextAfter: text.slice(start + length, start + length + SNIPPET_RADIUS),
  }
}
