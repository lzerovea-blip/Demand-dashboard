export function splitTextIntoLines(text: string, charactersPerLine = 20): string[] {
  if (charactersPerLine <= 0) return [];
  const characters = Array.from(text);
  const lines: string[] = [];
  for (let index = 0; index < characters.length; index += charactersPerLine) {
    lines.push(characters.slice(index, index + charactersPerLine).join(""));
  }
  return lines;
}
