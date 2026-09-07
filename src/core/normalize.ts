/**
 * macOS-origin Japanese text can arrive NFD (decomposed), where a dakuten
 * character is stored as base + combining mark. Normalizing to NFC before
 * matching keeps rule strings and document text comparable.
 */
export function normalize(text: string): string {
  return text.normalize('NFC')
}
