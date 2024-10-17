export type CellFormat = {Type: string | null, Style: string | null}
export type CellObject = {Value?: string, Formula?: string, Format: CellFormat}
export type CellIndex = {RowIndex: number, ColumnIndex: number}