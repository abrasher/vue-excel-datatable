import { defineConfig } from "vite"
import vue from "@vitejs/plugin-vue"
import { readFileSync } from "fs"
import { homedir } from "os"
import { resolve } from "path"
import WindiCSS from "vite-plugin-windicss"

const homeDir = homedir()
// https://vitejs.dev/config/
export default defineConfig({
  plugins: [vue(), WindiCSS()],
  build: {
    lib: {
      entry: resolve(__dirname, "src/index.ts"),
      name: "datatable",
    },
    rollupOptions: {
      external: ["vue"],
      output: {
        globals: {
          vue: "Vue",
        },
      },
    },
  },
})
