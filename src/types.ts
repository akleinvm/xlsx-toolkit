export type CellFormat = {type: "string" | "number" | null, style: string | null}
export type CellObject = {value?: string, formula?: string, format?: CellFormat}
export type CellIndex = {rowIndex: number, columnIndex: number}