export type CellObject = {value?: string | number | boolean, formula?: {value: string, type: 'string' | 'number' | 'boolean'}, style?: string | null}
export type CellIndex = {rowIndex: number, columnIndex: number}

//how will this work????:
// when value is empty use formula???
// or will it the opposite? when formula is empty use value