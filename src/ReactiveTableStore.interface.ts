export type CellType = string | number | boolean

export interface RowData {
  [key: string]: CellType
}

export interface RowNode {
  index: number
  data: RowData
}

export type ExcelRow = unknown[][]
export type RowIndex = number

export interface ColumnDef {
  label: string
  key: string
}

export interface Column {
  label: string
  key: string
  values: CellType[]
}

export interface UseTableParams {
  tableName: string
  sheetName: string
  columnDefs: ColumnDef[]
}
