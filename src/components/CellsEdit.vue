<script setup lang="ts">
import { computed, ref } from 'vue'
import { Workbook } from '../xlsx-populate'
import { selectedCellsToText, textToSelectedCells, type SelectedCell } from '../core'

// Workbook file contents (binary) and file name
const workbook = ref<Workbook | null>(null)
const workbookBytes = ref<Uint8Array | null>(null)
const filename = ref('')

// Sheet selection state
const selectedSheet = ref('')

// Selected cells (from user interaction or parsed from text)
const selectedCells = ref<SelectedCell[]>([])

// Validation error string (displayed in a banner)
const validationError = ref('')

// added
const visibleRows = ref<any[][]>([])

// ========== COMPUTED ==========

// Sheet list: derived from workbook (to be filled when workbook is loaded)
// const sheetList = computed<string[]>(() => {
//   // ⚠️ You will need to populate this list in the file loading logic
//   // e.g. by using workbook.sheets().map(s => s.name())
//   return [] // placeholder
// })
const sheetList = ref<string[]>([])

// Text representation of selected cells
const text = computed<string>({
  get() {
    return selectedCellsToText(selectedCells.value)
  },
  set(newText: string) {
    try {
      console.log('setText', newText)
      // Clear and re-parse input text
      const parsed = textToSelectedCells(newText)
      selectedCells.value = parsed.sort((a, b) =>
        a.sheet.localeCompare(b.sheet) || a.row - b.row || a.col - b.col
      )
      validationError.value = ''
    } catch (err) {
      validationError.value = String((err as Error).message)
    }
  }
})



// === FILE SELECT ===
function onFileSelected(event: Event) {
  const input = event.target as HTMLInputElement
  const file = input.files?.[0]
  if (!file) return

  filename.value = file.name

  // const reader = new FileReader()
  !(async () => {
    const data = new Uint8Array(await file.arrayBuffer())
    workbookBytes.value = data
    console.log('selected file=', file.name)

    workbook.value = await XlsxPopulate.fromDataAsync(data)
    console.log('workbook=', (window as any).workbook = workbook.value)
    console.log('sheet=', (window as any).sheet = workbook.value.sheets()?.[0])
    const sheets = workbook.value.sheets().map(s => s.name())
    sheetList.value = sheets
    selectedSheet.value = sheets[0] ?? ''
    visibleRows.value = loadVisibleRows(selectedSheet.value)
  })()
}

// === DOWNLOAD ===
function onDownload() {
  if (!workbook.value || !filename.value) return

  // 上書き
  for (const cell of selectedCells.value) {
    const sheet = workbook.value.sheet(cell.sheet)
    const target = sheet.cell(cell.row, cell.col)
    if (cell.formula != null) {
      target.formula(cell.formula)
    } else {
      target.value(cell.value ?? '')
    }
  }

  !(async () => {
    const blob = await workbook.value!.outputAsync()
    console.log('blob=', blob)
    downloadFile(new File([blob], filename.value), filename.value)

    // TODO: reset with the updated workbook
    console.log('downloaded')
  })()
}

// === UTILITY ===
// function downloadFile(blob: Blob, name: string) {
//   const link = document.createElement('a')
//   link.href = URL.createObjectURL(blob)
//   link.download = name
//   link.click()
//   URL.revokeObjectURL(link.href)
// }

// === SHEET NAV ===
function selectSheet(sheet: string) {
  selectedSheet.value = sheet
  visibleRows.value = loadVisibleRows(sheet)
}

// === VISIBLE DATA ===
function loadVisibleRows(sheetName: string): any[][] {
  const sheet = workbook.value?.sheet(sheetName)
  if (!sheet) return []

  const usedRange = sheet.usedRange()
  const rowCount = usedRange?.endCell().rowNumber() ?? 0
  const colCount = usedRange?.endCell().columnNumber() ?? 0

  const result: any[][] = []
  for (let row = 1; row <= rowCount; row++) {
    const rowData: any[] = []
    for (let col = 1; col <= colCount; col++) {
      const cell = sheet.cell(row, col)
      rowData.push(cell)
    }
    result.push(rowData)
  }
  return result
}

// === DISPLAY FORMATTING ===
function formatCell(cell: any): string {
  return cell.formula() != null ? `=${cell.formula()}` : String(cell.value() ?? '')
}

// === CELL SELECTION ===
function isCellSelected(row: number, col: number): boolean {
  return selectedCells.value.some(
    (c) => c.sheet === selectedSheet.value && c.row === row && c.col === col
  )
}

function toggleCellSelection(row: number, col: number) {
  if (!workbook.value) return
  console.log('toggle', row, col)

  const sheet = selectedSheet.value
  const idx = selectedCells.value.findIndex(
    (c) => c.sheet === sheet && c.row === row && c.col === col
  )

  if (idx >= 0) {
    // remove
    selectedCells.value.splice(idx, 1)
  } else {
    const cell = workbook.value.sheet(sheet).cell(row, col)
    const formula = cell.formula()
    const value = cell.value()
    selectedCells.value.push({
      sheet,
      row,
      col,
      formula: formula ?? undefined,
      value: formula == null ? String(value ?? '') : undefined
    })

    // keep sorted
    selectedCells.value.sort((a, b) =>
      a.sheet.localeCompare(b.sheet) || a.row - b.row || a.col - b.col
    )
  }
}

function downloadFile(blob: Blob, filename: string) {
  const url = URL.createObjectURL(blob)
  try {
    const a = document.createElement('a')
    a.href = url
    a.download = filename || 'edited.xlsx'
    document.body.appendChild(a)
    a.click()
    // document.body.removeChild(a)
    a.remove()
  } finally {
    URL.revokeObjectURL(url)
  }
}
</script>

<template>
  <div class="page-root">
    <!-- 🚨 Validation Error Banner -->
    <div v-if="validationError" class="error-banner">
      {{ validationError }}
    </div>

    <!-- 📁 File Controls -->
    <div class="controls">
      <input type="file" accept=".xlsx" @change="onFileSelected" />
      <button :disabled="!workbook" @click="onDownload">Download</button>
    </div>

    <!-- 🗂 Sheet Tabs -->
    <div class="sheet-tabs">
      <button v-for="sheet in sheetList" :key="sheet" :class="['sheet-tab', { active: sheet === selectedSheet }]"
        @click="selectSheet(sheet)">
        {{ sheet }}
      </button>
    </div>

    <!-- 🧾 Sheet Table + Text Area -->
    <div class="sheet-main">
      <!-- 📊 Scrollable Table -->
      <div class="sheet-table-container">
        <table class="sheet-table">
          <tr v-for="(row, rowIndex) in visibleRows" :key="rowIndex">
            <td v-for="(cell, colIndex) in row" :key="colIndex"
              :class="['cell', { selected: isCellSelected(rowIndex + 1, colIndex + 1) }]"
              @contextmenu.prevent.ctrl="toggleCellSelection(rowIndex + 1, colIndex + 1)"
              @click.prevent.ctrl="toggleCellSelection(rowIndex + 1, colIndex + 1)">
              {{ formatCell(cell) }}
            </td>
          </tr>
        </table>
      </div>

      <!-- 📝 Text Area -->
      <textarea class="selection-text" v-model="text" placeholder="Selected cell data appears here..."></textarea>
    </div>
  </div>
</template>

<style>
.page-root {
  max-width: 1024px;
  margin: 0 auto;
  padding: 1rem;
  font-family: sans-serif;
}

.controls {
  margin-bottom: 1rem;
  display: flex;
  gap: 1rem;
}

.sheet-tabs {
  margin-bottom: 0.5rem;
}

.sheet-tab {
  padding: 0.4rem 0.8rem;
  margin-right: 0.4rem;
  border: 1px solid #ccc;
  background-color: #eee;
  cursor: pointer;
}

.sheet-tab.active {
  background-color: #ddd;
  font-weight: bold;
}

.sheet-main {
  display: flex;
  flex-direction: column;
  gap: 1rem;
}

.sheet-table-container {
  max-height: 400px;
  overflow: auto;
  border: 1px solid #ccc;
}

.sheet-table {
  border-collapse: collapse;
  width: 100%;
  table-layout: fixed;
}

.sheet-table td {
  border: 1px solid #ddd;
  padding: 0.3rem;
  overflow: hidden;
  white-space: nowrap;
  text-overflow: ellipsis;
}

.sheet-table .cell.selected {
  background-color: #f0f8ff;
}

.selection-text {
  width: 100%;
  min-height: 200px;
  font-family: monospace;
  padding: 0.5rem;
  border: 1px solid #ccc;
  resize: vertical;
}

.error-banner {
  background-color: #ffe0e0;
  border: 1px solid #f88;
  padding: 0.75rem;
  margin-bottom: 1rem;
  color: #900;
}
</style>
