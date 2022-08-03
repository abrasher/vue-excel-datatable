import { reactive, readonly } from "vue"
import { ColumnDef, RowData, UseTableParams, CellType, RowNode } from "./ReactiveTableStore.interface"
import { numberToLetters, addressToXY } from "./utils"

class ReactiveTableStore {
  private _loading = true
  private _rows = reactive(new Map<number, RowNode>())

  private tablePositionColumn = 0
  private tablePositionRow = 0

  constructor(private tableName: string, private sheetName: string, private columnDefs: ColumnDef[]) {}

  async init() {
    Excel.run(async (context) => {
      await this.constructTable(context)
      await this.loadRowsFromExcel()
      await this.registerEventHandlers(context)
    })
  }

  /**
   * Getters
   */

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
  private async constructTable(context: Excel.RequestContext) {
    const rangeStr = `A1:${numberToLetters(this.numberOfColumns)}1`

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
  }

  private async loadRowsFromExcel() {
    Excel.run(async (context) => {
      const table = context.workbook.tables.getItem(this.tableName).load({ rows: { $all: true } })
      await context.sync()

      this._rows.clear()

      table.rows.items.forEach(({ index, values }) => {
        const data = arrayToObject(this.headers, values[0])

        this._rows.set(index, MakeRowNode(index, data))
      })
    })
  }

  private registerEventHandlers(context: Excel.RequestContext) {
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

  async addRow(row: RowData, index?: number) {
    const rowArray = objectToArray(row)

    await this.runWithTable(async (table, context) => {
      const newRow = table.rows.add(index ?? -1, [rowArray]).load()
      await context.sync()

      this._rows.set(newRow.index, MakeRowNode(newRow.index, row))
    })
  }

  deleteRow(index: number) {
    return this.runWithTable(async (table, context) => {
      table.rows.deleteRows([index])

      await context.sync()
      this._rows.delete(index)
    })
  }

  updateRow(index: number, data: RowData) {
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

  updateRowValue(rowIndex: number, columnIndex: number, value: CellType) {
    const keyName = this.headers[columnIndex]

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

const MakeRowNode = (index: number, data: RowData): RowNode => ({
  index,
  data,
})

const arrayToObject = (keys: string[], arr: CellType[]) => Object.fromEntries(arr.map((val, idx) => [keys[idx], val]))
const objectToArray = (obj: RowData, keysOrder?: string[]) => {
  if (keysOrder) {
    return keysOrder.map((key) => obj[key])
  }

  return Object.values(obj)
}

export const useTableStore = ({ tableName, sheetName, columnDefs }: UseTableParams) => {
  const table = new ReactiveTableStore(tableName, sheetName, columnDefs)

  table.init()

  return table
}
