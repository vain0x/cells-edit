<script setup lang="ts">
import { computed, ref } from 'vue'
import type { Workbook } from '../xlsx-populate'
import { createDummyWorkbook, selectedCellsToText, textToSelectedCells, type SelectedCell } from '../core'

const DEBUG = window.location.hostname === 'localhost'

const workbook = ref<Workbook | null>(null)
const filename = ref('')
const selectedSheet = ref('')
const selectedCells = ref<SelectedCell[]>([])
const validationError = ref('')

const visibleRows = ref<any[][]>([])

const sheetList = computed(() => {
  return workbook.value?.sheets().map(s => s.name()) ?? []
})

const text = computed<string>({
  get() {
    return selectedCellsToText(selectedCells.value)
  },
  set(newText: string) {
    try {
      if (DEBUG) console.log('setText', newText)
      const parsed = textToSelectedCells(newText)
      selectedCells.value = parsed.sort((a, b) =>
        a.sheet.localeCompare(b.sheet) || a.row - b.row || a.col - b.col
      )
      for (const cell of selectedCells.value) {
        if (cell.sheet !== selectedSheet.value) continue
        const row = cell.row
        const col = cell.col
        const target = visibleRows.value.at(row - 1)?.at(col - 1)
        if (target) {
          if (cell.formula != null) {
            target.formula(cell.formula)
          } else {
            target.value(cell.value)
          }
        }
      }
      validationError.value = ''
    } catch (err) {
      validationError.value = (err as Error).message
    }
  }
})

function onFileSelected(ev: Event) {
  const input = ev.target as HTMLInputElement
  const file = input.files?.[0]
  if (!file) return

  filename.value = file.name

  !(async () => {
    const data = new Uint8Array(await file.arrayBuffer())
    workbook.value = await XlsxPopulate.fromDataAsync(data)
    if (DEBUG) console.log('workbook=', (window as any).workbook = workbook.value)
    if (DEBUG) console.log('sheet=', (window as any).sheet = workbook.value.sheets()?.[0])
    const firstSheet = workbook.value.sheets()?.[0]?.name()
    selectedSheet.value = firstSheet ?? ''
    visibleRows.value = loadVisibleRows(selectedSheet.value)
  })()
}

function onDownload() {
  if (!workbook.value || !filename.value) return
  if (DEBUG) console.log('downloading')

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
    downloadFile(new File([blob], filename.value), filename.value)
    if (DEBUG) console.log('download finished')
  })()
}

function downloadFile(blob: Blob, filename: string) {
  const url = URL.createObjectURL(blob)
  try {
    const a = document.createElement('a')
    a.href = url
    a.download = filename || 'edited.xlsx'
    document.body.appendChild(a)
    a.click()
    a.remove()
  } finally {
    URL.revokeObjectURL(url)
  }
}

function selectSheet(sheet: string) {
  selectedSheet.value = sheet
  visibleRows.value = loadVisibleRows(sheet)
}

function loadVisibleRows(sheetName: string): any[][] {
  const sheet = workbook.value?.sheet(sheetName)
  if (!sheet) return []

  const range = sheet.usedRange()
  const rowCount = range?.endCell().rowNumber() ?? 0
  const colCount = range?.endCell().columnNumber() ?? 0

  const output: any[][] = []
  for (let row = 1; row <= rowCount; row++) {
    const rowData: any[] = []
    for (let col = 1; col <= colCount; col++) {
      const cell = sheet.cell(row, col)
      rowData.push(cell)
    }
    output.push(rowData)
  }
  return Object.freeze(output) as any[][]
}

function formatCell(cell: any): string {
  return cell.formula() != null ? `=${cell.formula()}` : String(cell.value() ?? '')
}

function isCellSelected(row: number, col: number): boolean {
  return selectedCells.value.some(
    (c) => c.sheet === selectedSheet.value && c.row === row && c.col === col
  )
}

function toggleCellSelection(row: number, col: number) {
  if (!workbook.value) return
  if (DEBUG) console.log('toggle', row, col)

  const sheet = selectedSheet.value
  const index = selectedCells.value.findIndex(
    (c) => c.sheet === sheet && c.row === row && c.col === col
  )

  if (index >= 0) {
    selectedCells.value.splice(index, 1)
  } else {
    const cell = workbook.value.sheet(sheet).cell(row, col)
    const formula = cell.formula()
    const value = cell.value()
    selectedCells.value.push({
      sheet, row, col,
      formula: formula ?? undefined,
      value: formula == null ? String(value ?? '') : undefined
    })
    selectedCells.value.sort((a, b) =>
      a.sheet.localeCompare(b.sheet) || a.row - b.row || a.col - b.col
    )
  }
}

!(async () => {
  const dummy = await createDummyWorkbook()
  const bytes = await dummy.outputAsync()
  const file = new File([bytes], 'untitled.txt')
  onFileSelected({ target: { files: [file] } } as any)
})()
</script>

<template>
  <div class="page-root">
    <div v-if="validationError" class="error-banner">
      {{ validationError }}
    </div>

    <div class="controls">
      <input type="file" accept=".xlsx" @change="onFileSelected" />
      <button :disabled="!workbook || !filename" @click="onDownload">Download</button>
    </div>

    <div class="sheet-tabs">
      <button v-for="sheet in sheetList" :key="sheet" :class="['sheet-tab', { active: sheet === selectedSheet }]"
        @click="selectSheet(sheet)">
        {{ sheet }}
      </button>
    </div>

    <div class="sheet-main">
      <div class="sheet-table-container">
        <table class="sheet-table">
          <tr v-for="(row, rowIndex) in visibleRows" :key="rowIndex" class="sheet-row">
            <td v-for="(cell, colIndex) in row" :key="colIndex"
              :class="['cell', { selected: isCellSelected(rowIndex + 1, colIndex + 1) }]"
              @click.prevent="toggleCellSelection(rowIndex + 1, colIndex + 1)">
              {{ formatCell(cell) }}
            </td>
          </tr>
        </table>
      </div>

      <textarea class="selection-text" v-model="text"></textarea>
    </div>

    <div class="footer">
      <div>* <kbd>Ctrl+Click</kbd> to select</div>
      <div>* Format: <code>sheet!(row, col): content</code></div>
    </div>
  </div>
</template>

<style>
.page-root {
  font-family: sans-serif;
  display: grid;
  grid-template-rows: auto auto 1fr auto;
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
  max-height: calc(100vh - 425px);
  overflow: auto;
  border: 1px solid #ccc;
}

.sheet-table {
  border-collapse: collapse;
  width: 100%;
  table-layout: fixed;
}


.sheet-row {
  height: 30px;
}

.sheet-table .cell {
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

.footer {
  margin-top: 1rem;
  color: #888888;
}
</style>
