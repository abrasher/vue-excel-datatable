export type CellType = string | number | boolean

export type RowData = Record<string, CellType>

export interface RowNode<T extends RowData> {
  index: number
  data: T
}

export type ExcelRow = unknown[][]
export type RowIndex = number

export interface ColumnDef<T extends RowData> {
  label: string
  key: keyof T
}

export interface Column {
  label: string
  key: string
  values: CellType[]
}

export interface UseTableParams<T extends RowData> {
  tableName: string
  sheetName: string
  columnDefs: ColumnDef<T>[]
  row?: number
  column?: number
}
