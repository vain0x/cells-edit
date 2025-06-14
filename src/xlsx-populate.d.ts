declare global {
  var XlsxPopulate: XlsxPopulate
}

export interface XlsxPopulate {
  fromBlankAsync(): Promise<Workbook>
  fromDataAsync(data: Uint8Array, opts?: unknown): Promise<Workbook>
  // fromFileAsync(filename: string): Promise<Workbook>
}

export declare class Workbook {
  sheets(): Sheet[]
  sheet(sheetNameOrIndex: string | number): Sheet
  addSheet(name: string): Sheet
  // cloneSheet(sheet: Sheet, name: string): Sheet

  // note: return type depends on opts
  outputAsync(opts?: object): Promise<ArrayBuffer>

  // toFileAsync(filename: string): Promise<void>
}

export declare class Sheet {
  usedRange(): Range
  name(): string
  range(address: string): Range
  range(y1: number, x1: number, y2: number, x2: number): Range

  cell(row: number, col: string | number): Cell
  cell(address: string): Cell

  // delete(): void
}

export declare class Cell {
  rowNumber(): number
  columnNumber(): number
  formula(): string | null
  formula(formula: string): void
  value(): CellValue | undefined
  value(v: CellValue | null | undefined): void
  // value(data?: (CellValue | null | undefined)[][]): Range
}

// export declare class Row {
//   rowNumber(): number
// }
// export declare class Column {
//   columnNumber(): number
// }
export declare class Range {
  endCell(): Cell
  // value(): CellValue
  value(v?: CellValue | null | undefined): void
}

export declare class RichText {
}

export type CellValue = string | boolean | number | Date | RichText
