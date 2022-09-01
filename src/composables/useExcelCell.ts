import { customRef, ref } from "vue"
import { CellType } from "./TableStore.interface"
import { useRangeRef } from "./useRange"

type UseCellParams = {
  // Row number of the cell to place in the document (zero-indexed), i.e. A1 = 0,0
  row: number
  // Row column of the cell to place in the document (zero-index), i.e. A1 = 0,0
  column: number
  // Sheet name where the cell is located
  sheetName: string
  // Binding Name
  bindingName: string
}

/**
 * It takes a range and returns a cell
 * @param {UseCellParams} params - UseCellParams
 * @returns A function that returns an object with the following properties:
 * - setValue
 * - value
 */
export const useCellRef = <T extends CellType>(params: UseCellParams) => {
  const { setValue, ...range } = useRangeRef({
    bindingName: params.bindingName,
    sheetName: params.sheetName,
    column: params.column,
    row: params.row,
    numberOfColumns: 1,
    numberOfRows: 1,
  })

  return {
    ...range,
    bindingName: range.bindingName,
    set value(data: T) {
      setValue.bind(this)([[data]])
    },
    get value() {
      return this._state.value?.[0][0] as T
    },
  }
}

// const useCellRefCustom = () => customRef((track, trigger) => {
//   return {
//     get() {

//     },
//     set(value) {

//     }
//   }
// })
