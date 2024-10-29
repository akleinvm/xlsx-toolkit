export type CellFormat = {type: "string" | "number" | null | undefined, style: string | null | undefined}
export type CellObject = {value?: string, formula?: string, format?: CellFormat}
export type CellIndex = {rowIndex: number, columnIndex: number}