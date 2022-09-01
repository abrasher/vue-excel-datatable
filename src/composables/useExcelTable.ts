import { reactive, readonly } from "vue"
import type { ColumnDef, RowData, UseTableParams, CellType, RowNode } from "./TableStore.interface"
import { numberToLetters, addressToXY } from "../common/utils"

/**
 * It creates a new TableStore instance, initializes it, and returns it.
 * @param {UseTableParams} params
 * @returns {TableStore} A new instance of the TableStore class.
 */
export const useExcelTable = <T extends RowData>(params: UseTableParams<T>) => {
  const table = new TableStore(params)

  table.init()

  return table
}

class TableStore<T extends RowData = any> {
  private tableName: string
  private sheetName: string
  private columnDefs: ColumnDef<T>[]

  private _loading = true
  private _rows = reactive(new Map<number, RowNode<T>>())

  private tablePositionColumn
  private tablePositionRow

  constructor(data: UseTableParams<T>) {
    this.tableName = data.tableName
    this.sheetName = data.sheetName
    this.columnDefs = data.columnDefs
    this.tablePositionRow = data.row ?? 0
    this.tablePositionColumn = data.column ?? 0
  }

  /**
   * The function `init` is an asynchronous function that calls the functions `constructTable`,
   * `loadRowsFromExcel`, and `registerEventHandlers` in sequence, and then returns a promise that
   * resolves to `undefined`.
   */
  async init() {
    await this.constructTable()
    await this.loadRowsFromExcel()
    await this.registerEventHandlers()
    this._loading = false
  }

  /**
   * Getters
   */

  get keys() {
    return this.columnDefs.map((col) => col.key)
  }

  get headers(): string[] {
    return this.columnDefs.map((col) => col.label)
  }

  get numberOfColumns() {
    return this.columnDefs.length
  }

  get loading() {
    return this._loading
  }

  get rows() {
    return readonly(this._rows.values())
  }

  /**
   * Initialization Functions
   */

  // Creates the table in Excel if it does not exist
  private constructTable() {
    return Excel.run(async (context) => {
      const rangeStr = `A1:${numberToLetters(this.numberOfColumns - 1)}1`
      console.log(rangeStr)

      let importSheet = context.workbook.worksheets.getItemOrNullObject(this.sheetName)
      await context.sync()

      // If the sheet doesn't exist, then the table for sure doesn't exist
      if (importSheet.isNullObject) {
        importSheet = context.workbook.worksheets.add(this.sheetName)
      }

      let importTable = importSheet.tables.getItemOrNullObject(this.tableName)
      await context.sync()

      // Create table if it does not exist
      if (importTable.isNullObject) {
        importSheet.getRange(rangeStr).values = [this.headers]

        importTable = importSheet.tables.add(rangeStr, true)
        importTable.name = this.tableName
      }
      await context.sync()
    })
  }

  /**
   * It loads all the rows from the table in the Excel workbook, and then it creates a new row node for
   * each row in the table.
   *
   * The row node is a class that contains the row index and the data for the row.
   *
   */
  private async loadRowsFromExcel() {
    this.runWithTable(async (table, context) => {
      table.load({ rows: { $all: true } })

      await context.sync()

      this._rows.clear()

      table.rows.items.forEach(({ index, values }) => {
        const data = arrayToObject(this.keys, values[0]) as T

        this._rows.set(index, MakeRowNode<T>(index, data))
      })
    })
  }

  /**
   * When a row is deleted, inserted, or a cell is edited, update the state of the table.
   * @returns The return value is a promise that resolves when the event handler is registered.
   */
  private registerEventHandlers() {
    return this.runWithTable(async (table, context) => {
      table.onChanged.add(async (event) => {
        switch (event.changeType) {
          // On row Deleted, refresh the entire state
          case "RowDeleted":
            console.info("Row Deleted")
            this.loadRowsFromExcel()
            break
          // Inserting a row can also immediately trigger a RangeEdited
          case "RowInserted":
            console.info("Row Inserted")
            this.loadRowsFromExcel()
            break
          case "RangeEdited":
            console.info("Range Edited")

            // event details only populated when a single cell is editted
            if (event.details) {
              if (event.details.valueAfter === event.details.valueBefore) return

              const [col, row] = addressToXY(event.address)

              // subtract 1 to account for being 1 indexed
              const rowIndex = row - this.tablePositionRow
              // subtract 1 to account for being 1 indexed (i.e. A=1)
              const columnIndex = col - this.tablePositionColumn - 1

              this.updateRowValue(rowIndex, columnIndex, event.details.valueAfter)
            }

            break
          default:
            break
        }
      })
    })
  }

  /**
   * Internal Helper Functions
   */

  runWithTable(func: (table: Excel.Table, context: Excel.RequestContext) => Promise<void>) {
    return Excel.run(async (context) => {
      try {
        const table = context.workbook.tables.getItem(this.tableName)
        await context.sync()
        return func(table, context)
      } catch (error) {
        console.error("ERROR: Exception occured while running with table \n", error)
      }
    })
  }

  /**
   * Mutation Functions
   */

  /**
   * It adds a row to the table, and then adds the row to the internal _rows map.
   * @param {RowData} row - RowData - The data to add to the table.
   * @param {number} index - The index of the row to add. If omitted, the row will be added to the
   * end of the table.
   */
  async addRow(row: T, index?: number) {
    const rowArray = objectToArray(row)

    await this.runWithTable(async (table, context) => {
      const newRow = table.rows.add(index ?? -1, [rowArray]).load()
      await context.sync()

      this._rows.set(newRow.index, MakeRowNode(newRow.index, row))
    })
  }

  /**
   * It deletes a row from the table, and then deletes the row from the internal _rows collection.
   * @param {number} index - number - The index of the row to delete.
   * @returns The return value is a promise that resolves when the operation is complete.
   */
  deleteRow(index: number) {
    return this.runWithTable(async (table, context) => {
      table.rows.deleteRows([index])

      await context.sync()
      this._rows.delete(index)
    })
  }

  /**
   * It takes an index and a row of data, and updates the row at that index with the new data
   * @param {number} index - number - The index of the row you want to update
   * @param {RowData} data - RowData - The data to update the row with.
   * @returns The return value is the result of the function.
   */
  updateRow(index: number, data: T) {
    const rowNode = this._rows.get(index)

    if (!rowNode) throw new Error(`Row Index: ${index} does not exist`)

    const rowArray = objectToArray(data)
    return this.runWithTable(async (table, context) => {
      const row = table.rows.getItemAt(index)
      row.values = [rowArray]
      await context.sync()

      rowNode.data = data
    })
  }

  updateRowValue(rowIndex: number, columnIndex: number, value: T[keyof T]) {
    const keyName = this.keys[columnIndex] as keyof T

    const colVal = columnIndex + this.tablePositionColumn
    // add 1 to account for header
    const rowVal = rowIndex + this.tablePositionRow + 1

    console.log(`Cell to Update: row ${rowVal}, col ${colVal}`)

    const rowNode = this._rows.get(rowIndex)
    if (rowNode) {
      Excel.run(async (context) => {
        const cell = context.workbook.worksheets.getItem(this.sheetName).getCell(rowVal, rowVal).load({ values: true })

        await context.sync()
        cell.values = [[value]]
        await context.sync()

        rowNode.data[keyName] = value
      })
    }
  }

  get state() {
    return Array.from(this._rows.values())
  }
}

/**
 * `MakeRowNode` takes an index and a row of data and returns a row node.
 * @param {number} index - number - The index of the row in the data array.
 * @param {RowData} data - The data that will be displayed in the row.
 */
const MakeRowNode = <T extends RowData>(index: number, data: T): RowNode<T> => ({
  index,
  data,
})

/**
 * It takes an array of keys and an array of values and returns an object with the keys as the keys and
 * the values as the values
 * @param {string[]} keys - string[] - an array of strings that will be used as the keys for the object
 * @param {CellType[]} arr - The array to be converted to an object
 */
const arrayToObject = (keys: string[], arr: CellType[]) => Object.fromEntries(arr.map((val, idx) => [keys[idx], val]))

/**
 * If keysOrder is defined, return an array of the values of obj in the order of keysOrder, otherwise
 * return an array of the values of obj.
 * @param {RowData} obj - {
 * @param {string[]} [keysOrder] - ['id', 'name', 'age']
 * @returns An array of values from the object.
 */
const objectToArray = (obj: RowData, keysOrder?: string[]) => {
  if (keysOrder) {
    return keysOrder.map((key) => obj[key])
  }

  return Object.values(obj)
}
