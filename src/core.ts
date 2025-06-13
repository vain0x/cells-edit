export type SelectedCell = {
  sheet: string
  row: number
  col: number
  value?: string
  formula?: string
}

export type ValidationError = {
  line: number
  lineText: string
  message: string
}

/**
 * Converts a list of selected cells to multiline text.
 *
 * Format:
 *   sheet!(row,col): value
 *   sheet!(row,col): =formula
 */
export function selectedCellsToText(cells: SelectedCell[]): string {
  return cells
    .map(({ sheet, row, col, value, formula }) => {
      const body = formula != null ? `=${formula}` : value ?? ''
      return `${sheet}!(${row},${col}): ${body}`
    })
    .join('\n')
}

/**
 * Parses multiline text into a list of selected cells.
 * Throws an Error if parsing fails.
 */
export function textToSelectedCells(text: string): SelectedCell[] {
  const lines = text.trimEnd().split(/\r?\n/)
  const cells: SelectedCell[] = []

  const pattern = /^([^ !:]+)!\((\d+),(\d+)\):\s*(.*)$/

  for (const [index, line] of lines.entries()) {
    if (!line.trim()) continue

    const match = line.match(pattern)
    if (!match) {
      throw new Error(`Line ${index + 1}: Invalid format ('${line}')`)
    }

    const [, sheet, rowStr, colStr, content] = match
    const row = +rowStr
    const col = +colStr

    if (!row || !col) {
      throw new Error(`Line ${index + 1}: Invalid row or column number ('${line}')`)
    }

    if (content.startsWith('=')) {
      cells.push({ sheet, row, col, formula: content.slice(1) })
    } else {
      cells.push({ sheet, row, col, value: content })
    }
  }

  return cells
}
