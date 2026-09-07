import type { LintRule } from '../core/types'
import type { Occurrence } from '../adapters/types'
import { CorrectionMark } from './CorrectionMark'

interface RuleGroup {
  rule: Pick<LintRule, 'id' | 'wrong' | 'correct' | 'note'>
  occurrences: Occurrence[]
}

function groupByRule(occurrences: Occurrence[]): RuleGroup[] {
  const groups = new Map<string, RuleGroup>()
  for (const occ of occurrences) {
    let group = groups.get(occ.ruleId)
    if (!group) {
      group = { rule: { id: occ.ruleId, wrong: occ.wrong, correct: occ.correct, note: occ.note }, occurrences: [] }
      groups.set(occ.ruleId, group)
    }
    group.occurrences.push(occ)
  }
  return [...groups.values()]
}

export function IssueList({
  occurrences,
  onSelect,
  onApplyOne,
  onApplyAll,
}: {
  occurrences: Occurrence[]
  onSelect: (occurrenceId: string) => void
  onApplyOne: (occurrenceId: string) => void
  onApplyAll: (ruleId: string) => void
}) {
  const groups = groupByRule(occurrences)

  if (groups.length === 0) return null

  return (
    <div>
      {groups.map((group) => (
        <div className="rule-group" key={group.rule.id}>
          <div className="rule-group-header">
            <CorrectionMark wrong={group.rule.wrong} correct={group.rule.correct} />
            <button className="fix-all-button" onClick={() => onApplyAll(group.rule.id)}>
              すべて修正 ({group.occurrences.length})
            </button>
          </div>
          {group.rule.note && <div className="rule-group-note">{group.rule.note}</div>}
          <div className="occurrence-list">
            {group.occurrences.map((occ) => (
              <button className="occurrence-row" key={occ.id} onClick={() => onSelect(occ.id)}>
                {occ.location && <span className="occurrence-location">{occ.location}</span>}
                <span className="occurrence-snippet">
                  {occ.contextBefore}
                  <span className="matched">{occ.wrong}</span>
                  {occ.contextAfter}
                </span>
                <span
                  className="occurrence-fix-button"
                  role="button"
                  tabIndex={0}
                  onClick={(e) => {
                    e.stopPropagation()
                    onApplyOne(occ.id)
                  }}
                  onKeyDown={(e) => {
                    if (e.key === 'Enter' || e.key === ' ') {
                      e.stopPropagation()
                      onApplyOne(occ.id)
                    }
                  }}
                >
                  修正
                </span>
              </button>
            ))}
          </div>
        </div>
      ))}
    </div>
  )
}
