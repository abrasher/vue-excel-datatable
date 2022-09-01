import { ref } from "vue"
import { CellType } from "./TableStore.interface"

type RangeValues<Rows extends number, Columns extends number> = Tuple<Tuple<CellType, Columns>, Rows>

type Tuple<T, N extends number> = N extends N ? (number extends N ? T[] : _TupleOf<T, N, []>) : never
type _TupleOf<T, N extends number, R extends unknown[]> = R["length"] extends N ? R : _TupleOf<T, N, [T, ...R]>

type UseRangeParams<RowLength extends number, ColumnLength extends number> = {
  // Row number of the cell to place in the document (zero-indexed), i.e. A1 = 0,0
  row: number
  // Row column of the cell to place in the document (zero-index), i.e. A1 = 0,0
  column: number
  // Sheet name where the cell is located
  sheetName: string
  // Binding Name
  bindingName: string
  // Number of rows
  numberOfRows: RowLength
  // Number of columns
  numberOfColumns: ColumnLength
}

export const useRangeRef = <R extends number, C extends number>(params: UseRangeParams<R, C>) => {
  const rangeRef = new RangeRef(params)

  rangeRef.init()

  return rangeRef
}

export class RangeRef<R extends number, C extends number> {
  _state = ref<RangeValues<R, C>>()

  readonly row: number
  readonly column: number
  readonly sheetName: string
  readonly bindingName: string
  readonly numRows: number
  readonly numColumns: number

  constructor(data: UseRangeParams<R, C>) {
    this.row = data.row
    this.column = data.column
    this.sheetName = data.sheetName
    this.bindingName = data.bindingName
    this.numRows = data.numberOfRows
    this.numColumns = data.numberOfColumns
  }

  async init() {
    await this.constructSheet()
    await this.loadStateFromExcel()
    await this.addBindingToRange()
  }

  setValue(newValue: RangeValues<R, C>) {
    Excel.run(async (context) => {
      const range = context.workbook.bindings.getItem(this.bindingName).getRange().load({ values: true })

      await context.sync()

      range.values = newValue

      await context.sync()

      this._state.value = newValue
    })
  }

  set value(newValue: RangeValues<R, C>) {
    this.setValue(newValue)
  }

  get value() {
    return this._state.value as RangeValues<R, C>
  }

  // Creates the sheet in Excel if it does not exist
  private constructSheet() {
    return Excel.run(async (context) => {
      let sheet = context.workbook.worksheets.getItemOrNullObject(this.sheetName)
      await context.sync()

      // Create the sheet if it does not exist
      if (sheet.isNullObject) {
        context.workbook.worksheets.add(this.sheetName)
      }
      await context.sync()
    })
  }

  private loadStateFromExcel() {
    return Excel.run(async (context) => {
      const cell = context.workbook.worksheets
        .getItem(this.sheetName)
        .getCell(this.row, this.column)
        .load({ values: true })

      await context.sync()

      this._state.value = cell.values as RangeValues<R, C>
    })
  }

  /**
   * When a row is deleted, inserted, or a cell is edited, update the state of the table.
   * @returns The return value is a promise that resolves when the event handler is registered.
   */
  private async addBindingToRange() {
    return await Excel.run(async (context) => {
      const range = context.workbook.worksheets
        .getItem(this.sheetName)
        .getRangeByIndexes(this.row, this.column, this.numRows, this.numColumns)

      await context.sync()
      const binding = context.workbook.bindings.add(range, "Range", this.bindingName)

      binding.onDataChanged.add(async (event) => {
        const context = event.binding.context
        const eventRange = event.binding.getRange().load({ values: true })

        await context.sync()

        this._state.value = eventRange.values as RangeValues<R, C>
      })
    })
  }
}
