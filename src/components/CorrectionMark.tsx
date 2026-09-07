/** The app's signature visual: a proofreading circle around the wrong term, arrow to the fix. */
export function CorrectionMark({ wrong, correct }: { wrong: string; correct: string }) {
  return (
    <span className="correction-mark">
      <span className="wrong">{wrong}</span>
      <span className="arrow" aria-hidden="true">
        →
      </span>
      <span className="correct">{correct}</span>
    </span>
  )
}
