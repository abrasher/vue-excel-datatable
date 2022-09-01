import { ref } from "vue"
import { CellType } from "./TableStore.interface"
import { useRangeRef } from "./useRange"

const usePropertyTable = () => {
  const range = useRangeRef({})
}

type UsePropertyParams = {
  label: string
  unit: string
  sheetName: string
  properties: Property[]
  row: number
  column: number
}

type Property = {
  label: string
  unit: string
  key: string
}

export const useProperty = (params: UsePropertyParams) => {}

class PropertyTable<T extends CellType> {
  readonly label: string
  readonly unit: string
  readonly sheetName: string
  readonly propertyName: string
  readonly row: number
  readonly column: number

  _value = ref<T>()

  constructor(data: UsePropertyParams) {
    this.label = data.label
    this.unit = data.unit
    this.sheetName = data.sheetName
    this.propertyName = data.propertyName
    this.row = data.row
    this.column = data.column
  }
}
