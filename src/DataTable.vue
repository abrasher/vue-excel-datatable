<template>
  <div>
    <table id="vuedatatable" class="m-2 table-fixed outline outline-1 outline-solid-gray-600">
      <thead>
        <tr>
          <th v-for="header of table.headers" class="px-2 py-1">{{ header }}</th>
        </tr>
      </thead>
      <tbody>
        <tr v-for="row of table.rows">
          <td v-for="(val, _, index) of row.data" class="x-2 py-1 text-center">
            <input :value="val" @input="(event) => updateValue(row.index, index, event)" />
          </td>
          <td class="x-2 py-1 text-center">
            <button @click="remove(row.index)">Delete</button>
          </td>
          <td class="x-2 py-1 text-center">
            <button @click="update(row.index)">Update</button>
          </td>
        </tr>
      </tbody>
    </table>
    <button @click="add">Add Row</button>

    <br />
    {{ table.state }}
  </div>
</template>

<script setup lang="ts">
import type { ColumnDef } from "../composables/TableStore.interface"
import { useExcelTable } from "../composables"

// Must export props interface to avoid build errors
export interface Props {
  sheetName: string
  tableName: string
  columnDefs: ColumnDef[]
}

const updatePond = () => {}

const props = defineProps<Props>()

const table = useExcelTable({
  sheetName: "TableSheet",
  tableName: "TestTabel",
  columnDefs: [
    {
      key: "name",
      label: "Name",
    },
    {
      key: "age",
      label: "Age",
    },
    {
      key: "location",
      label: "Location",
    },
    {
      key: "gender",
      label: "Gender",
    },
  ],
})

const updateValue = (rowIndex: number, columnIndex: number, event: Event) => {
  if (event.target) {
    table.updateRowValue(rowIndex, columnIndex, (event.target as HTMLInputElement).value)
  }
}

const add = () => {
  table.addRow({
    name: "Samra",
    age: 24,
    location: "Port Credit",
    gender: "female",
  })
}

const remove = (index: number) => {
  table.deleteRow(index)
}

const update = (index: number) => {
  table.updateRow(index, {
    name: "Frank",
    age: Math.random(),
    location: "Cooksville",
    gender: "Male",
  })
}
</script>

<style></style>
