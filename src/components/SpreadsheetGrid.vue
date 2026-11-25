<script setup lang="ts">
import { computed, nextTick, onBeforeUnmount, onMounted, ref } from 'vue'
import { colName, type CellView } from '@/lib/excel'
import { parseTSV, toTSVCell } from '@/lib/tsv'

interface Pos {
  r: number
  c: number
}

// Every row is exactly ROW_H tall, which is what makes row virtualisation cheap.
const ROW_H = 28
const HEADER_H = 28
const ROWNUM_W = 52
const OVERSCAN = 10

const props = defineProps<{
  rowCount: number
  colCount: number
  colWidths: number[]
  frozenRows: number
  frozenCols: number
  /** Bumped by the store whenever cell contents change. */
  revision: number
  view: (r: number, c: number) => CellView | undefined
  editText: (r: number, c: number) => string
  isEdited: (r: number, c: number) => boolean
  resolve: (r: number, c: number) => Pos
}>()

const selected = defineModel<Pos>('selected', { required: true })

const emit = defineEmits<{
  commit: [r: number, c: number, text: string]
  clear: [r: number, c: number]
  paste: [r: number, c: number, rows: string[][]]
  editing: [draft: string | null]
}>()

const scroller = ref<HTMLDivElement>()
// Keyboard focus lives in this hidden textarea rather than the grid itself, so
// text that arrives without a keydown per character (IME, dictation, paste-like
// input) still reaches us through its input event.
const sink = ref<HTMLTextAreaElement>()
const scrollTop = ref(0)
const viewportH = ref(600)
const editing = ref<Pos | null>(null)
const draft = ref('')
let editorEl: HTMLTextAreaElement | null = null
const setEditor = (el: unknown) => (editorEl = el as HTMLTextAreaElement | null)

/** Left edge (px) of column c, at index c. colLeft[colCount + 1] is the total width. */
const colLeft = computed(() => {
  const out = [0, ROWNUM_W]
  for (let c = 1; c <= props.colCount; c++) out.push(out[c] + (props.colWidths[c - 1] ?? 100))
  return out
})
const totalWidth = computed(() => colLeft.value[props.colCount + 1])
const frozenWidth = computed(() => colLeft.value[props.frozenCols + 1])
const frozenHeight = computed(() => HEADER_H + props.frozenRows * ROW_H)

type Item =
  | { kind: 'row'; key: string; r: number; frozenIndex: number; cells: (CellView | undefined)[] }
  | { kind: 'spacer'; key: string; height: number }

/** Frozen rows, then a spacer, the rows in view, and another spacer. */
const items = computed<Item[]>(() => {
  void props.revision
  const row = (r: number, frozenIndex: number): Item => {
    const cells: (CellView | undefined)[] = []
    for (let c = 1; c <= props.colCount; c++) cells.push(props.view(r, c))
    return { kind: 'row', key: `r${r}`, r, frozenIndex, cells }
  }
  const list: Item[] = []
  const frozen = Math.min(props.frozenRows, props.rowCount)
  for (let r = 1; r <= frozen; r++) list.push(row(r, r - 1))

  const first = frozen + 1
  const start = Math.max(first, Math.floor(scrollTop.value / ROW_H) + first - OVERSCAN)
  const end = Math.min(props.rowCount, Math.ceil((scrollTop.value + viewportH.value - HEADER_H) / ROW_H) + OVERSCAN)
  list.push({ kind: 'spacer', key: 'top', height: Math.max(0, start - first) * ROW_H })
  for (let r = start; r <= end; r++) list.push(row(r, -1))
  list.push({ kind: 'spacer', key: 'bottom', height: Math.max(0, props.rowCount - Math.max(end, start - 1)) * ROW_H })
  return list
})

const isSel = (r: number, c: number) => selected.value.r === r && selected.value.c === c
const isEditing = (r: number, c: number) => editing.value?.r === r && editing.value?.c === c

function cellClass(r: number, c: number, cell: CellView | undefined) {
  return [
    cell?.align && `al-${cell.align}`,
    c <= props.frozenCols && 'fz',
    isSel(r, c) && 'sel',
    props.isEdited(r, c) && 'edited',
    isEditing(r, c) && 'editing',
  ]
}

function stickyStyle(frozenIndex: number, c: number) {
  const s: Record<string, string> = {}
  if (c <= props.frozenCols) s.left = `${colLeft.value[c]}px`
  if (frozenIndex >= 0) s.top = `${HEADER_H + frozenIndex * ROW_H}px`
  return s
}

// ---- selection & scrolling -------------------------------------------------

const clamp = (v: number, lo: number, hi: number) => Math.min(hi, Math.max(lo, v))

function select(r: number, c: number) {
  r = clamp(r, 1, props.rowCount)
  c = clamp(c, 1, props.colCount)
  selected.value = { r, c }
  ensureVisible(r, c)
}

function ensureVisible(r: number, c: number) {
  const el = scroller.value
  if (!el) return
  if (r > props.frozenRows) {
    const top = HEADER_H + (r - 1) * ROW_H
    if (top < el.scrollTop + frozenHeight.value) el.scrollTop = top - frozenHeight.value
    else if (top + ROW_H > el.scrollTop + el.clientHeight) el.scrollTop = top + ROW_H - el.clientHeight
  }
  if (c > props.frozenCols) {
    const left = colLeft.value[c]
    const right = colLeft.value[c + 1]
    if (left < el.scrollLeft + frozenWidth.value) el.scrollLeft = left - frozenWidth.value
    else if (right > el.scrollLeft + el.clientWidth) el.scrollLeft = right - el.clientWidth
  }
  // Update now rather than waiting for the scroll event, so the row is rendered on the next tick.
  scrollTop.value = el.scrollTop
}

const onScroll = () => (scrollTop.value = scroller.value?.scrollTop ?? 0)
const pageRows = () => Math.max(1, Math.floor((viewportH.value - frozenHeight.value) / ROW_H) - 1)
const focusGrid = () => sink.value?.focus({ preventScroll: true })

function onSinkInput(e: Event) {
  if ((e as InputEvent).isComposing || !sink.value) return
  const text = sink.value.value
  sink.value.value = ''
  if (text && !editing.value) startEdit(text)
}

// ---- editing ---------------------------------------------------------------

function startEdit(initial?: string) {
  const p = props.resolve(selected.value.r, selected.value.c)
  selected.value = p
  ensureVisible(p.r, p.c)
  editing.value = p
  draft.value = initial ?? props.editText(p.r, p.c)
  emit('editing', draft.value)
  nextTick(() => {
    if (!editorEl) return
    editorEl.focus({ preventScroll: true })
    const n = editorEl.value.length
    editorEl.setSelectionRange(n, n)
  })
}

function commit(move?: [number, number]) {
  const ed = editing.value
  if (!ed) return
  editing.value = null
  emit('editing', null)
  emit('commit', ed.r, ed.c, draft.value)
  if (move) select(ed.r + move[0], ed.c + move[1])
  focusGrid()
}

function cancel() {
  editing.value = null
  emit('editing', null)
  focusGrid()
}

// The editor is removed when we move to another cell; its blur then arrives
// after the next cell is already being edited, so only commit for this cell.
function onEditorBlur(r: number, c: number) {
  if (isEditing(r, c)) commit()
}

function onEditorKeydown(e: KeyboardEvent) {
  if (e.isComposing) return
  // Keys handled here must not bubble to the grid's own handler: by then the edit
  // is committed, so Enter would immediately start editing the next cell.
  if (e.key === 'Enter' || e.key === 'Tab' || e.key === 'Escape') e.stopPropagation()
  if (e.key === 'Enter') {
    e.preventDefault()
    if (e.altKey) {
      // Alt+Enter: new line inside the cell, as in Excel.
      const el = e.target as HTMLTextAreaElement
      el.setRangeText('\n', el.selectionStart, el.selectionEnd, 'end')
      draft.value = el.value
      return
    }
    commit([e.shiftKey ? -1 : 1, 0])
  } else if (e.key === 'Tab') {
    e.preventDefault()
    commit([0, e.shiftKey ? -1 : 1])
  } else if (e.key === 'Escape') {
    e.preventDefault()
    cancel()
  }
}

// ---- mouse & keyboard --------------------------------------------------------

function onCellMouseDown(r: number, c: number, e: MouseEvent) {
  if (editing.value) {
    if (isEditing(r, c)) return
    commit()
  }
  // Right-click selects the cell too, so the context menu acts on it.
  if (e.button === 0 || e.button === 2) selected.value = { r, c }
}

function onHeaderMouseDown(r: number, c: number) {
  if (editing.value) commit()
  select(r, c)
}

function copySelected(cut: boolean) {
  const { r, c } = props.resolve(selected.value.r, selected.value.c)
  void navigator.clipboard?.writeText(toTSVCell(props.view(r, c)?.text ?? ''))
  if (cut) emit('clear', r, c)
}

function onKeydown(e: KeyboardEvent) {
  // Only keys pressed on the grid itself, not ones bubbling up from the cell editor.
  if (e.target !== sink.value && e.target !== scroller.value) return
  if (editing.value || e.isComposing) return
  const { r, c } = selected.value

  if (e.metaKey || e.ctrlKey) {
    const k = e.key.toLowerCase()
    if (k === 'c' || k === 'x') {
      e.preventDefault()
      copySelected(k === 'x')
    }
    return // paste arrives as a 'paste' event; other shortcuts belong to the app
  }
  if (e.altKey) return

  const go = (nr: number, nc: number) => {
    e.preventDefault()
    select(nr, nc)
  }
  switch (e.key) {
    case 'ArrowUp':
      return go(r - 1, c)
    case 'ArrowDown':
      return go(r + 1, c)
    case 'ArrowLeft':
      return go(r, c - 1)
    case 'ArrowRight':
      return go(r, c + 1)
    case 'Tab':
      return go(r, e.shiftKey ? c - 1 : c + 1)
    case 'PageDown':
      return go(r + pageRows(), c)
    case 'PageUp':
      return go(r - pageRows(), c)
    case 'Home':
      return go(r, 1)
    case 'Enter':
      e.preventDefault()
      return e.shiftKey ? select(r - 1, c) : startEdit()
    case 'F2':
      e.preventDefault()
      return startEdit()
    case 'Delete':
    case 'Backspace': {
      e.preventDefault()
      const p = props.resolve(r, c)
      return emit('clear', p.r, p.c)
    }
  }
  // Typing a character starts editing and replaces the cell, as in Excel.
  if (e.key.length === 1) {
    e.preventDefault()
    startEdit(e.key)
  }
}

function onPaste(e: ClipboardEvent) {
  if (editing.value || document.activeElement !== sink.value) return
  const text = e.clipboardData?.getData('text/plain')
  if (!text) return
  e.preventDefault()
  const p = props.resolve(selected.value.r, selected.value.c)
  emit('paste', p.r, p.c, parseTSV(text))
}

let resizeObs: ResizeObserver | undefined
onMounted(() => {
  document.addEventListener('paste', onPaste)
  resizeObs = new ResizeObserver(() => (viewportH.value = scroller.value?.clientHeight ?? 600))
  if (scroller.value) resizeObs.observe(scroller.value)
  focusGrid()
})
onBeforeUnmount(() => {
  document.removeEventListener('paste', onPaste)
  resizeObs?.disconnect()
})

defineExpose({ focus: focusGrid, commit: () => commit(), select, startEdit, copy: () => copySelected(false) })
</script>

<template>
  <div ref="scroller" class="grid-root" tabindex="0" @focus="focusGrid" @keydown="onKeydown" @scroll.passive="onScroll">
    <textarea
      ref="sink"
      class="sink"
      data-grid-sink
      aria-hidden="true"
      tabindex="-1"
      autocomplete="off"
      spellcheck="false"
      @input="onSinkInput"
      @compositionend="onSinkInput"
    />
    <table :style="{ width: `${totalWidth}px` }">
      <colgroup>
        <col :style="{ width: `${ROWNUM_W}px` }" />
        <col v-for="c in colCount" :key="c" :style="{ width: `${colWidths[c - 1]}px` }" />
      </colgroup>
      <thead>
        <tr>
          <th class="corner" />
          <th
            v-for="c in colCount"
            :key="c"
            class="col-head"
            :class="{ active: c === selected.c, fz: c <= frozenCols }"
            :style="c <= frozenCols ? { left: `${colLeft[c]}px` } : undefined"
            @mousedown.prevent="onHeaderMouseDown(selected.r, c)"
          >
            {{ colName(c) }}
          </th>
        </tr>
      </thead>
      <tbody>
        <template v-for="item in items" :key="item.key">
          <tr v-if="item.kind === 'spacer'" class="spacer" :style="{ height: `${item.height}px` }">
            <td :colspan="colCount + 1" />
          </tr>
          <tr v-else :class="{ 'fz-row': item.frozenIndex >= 0 }">
            <th
              class="row-head"
              :class="{ active: item.r === selected.r }"
              :style="stickyStyle(item.frozenIndex, 0)"
              @mousedown.prevent="onHeaderMouseDown(item.r, selected.c)"
            >
              {{ item.r }}
            </th>
            <td
              v-for="(cell, i) in item.cells"
              :key="i"
              :class="cellClass(item.r, i + 1, cell)"
              :style="[cell?.style, stickyStyle(item.frozenIndex, i + 1)]"
              @mousedown="onCellMouseDown(item.r, i + 1, $event)"
              @dblclick="startEdit()"
            >
              <textarea
                v-if="isEditing(item.r, i + 1)"
                :ref="setEditor"
                v-model="draft"
                class="editor"
                rows="1"
                spellcheck="false"
                @input="emit('editing', draft)"
                @keydown="onEditorKeydown"
                @blur="onEditorBlur(item.r, i + 1)"
              />
              <template v-else>{{ cell?.text }}</template>
            </td>
          </tr>
        </template>
      </tbody>
    </table>
  </div>
</template>

<style scoped>
.grid-root {
  --g-surface: var(--background);
  --g-head: var(--muted);
  --g-line: var(--border);
  --g-text: var(--foreground);
  --g-muted: var(--muted-foreground);
  --g-accent: #059669;
  --g-accent-weak: color-mix(in oklab, #10b981 16%, var(--background));
  --g-freeze: color-mix(in oklab, var(--muted-foreground) 55%, transparent);

  position: relative;
  height: 100%;
  overflow: auto;
  outline: none;
  user-select: none;
  background: var(--g-surface);
  font-size: 13px;
}

.sink {
  position: fixed;
  top: 0;
  left: 0;
  width: 1px;
  height: 1px;
  padding: 0;
  border: 0;
  opacity: 0;
  resize: none;
  pointer-events: none;
}

table {
  table-layout: fixed;
  border-collapse: separate;
  border-spacing: 0;
}

tr {
  height: 28px;
}

th,
td {
  box-sizing: border-box;
  height: 28px;
  line-height: 27px;
  padding: 0 6px;
  border-right: 1px solid var(--g-line);
  border-bottom: 1px solid var(--g-line);
  white-space: nowrap;
  overflow: hidden;
  text-overflow: ellipsis;
}

td {
  position: relative;
  background: var(--g-surface);
  color: var(--g-text);
  cursor: cell;
}

th {
  background: var(--g-head);
  color: var(--g-muted);
  font-weight: 500;
  font-size: 11.5px;
  text-align: center;
  font-variant-numeric: tabular-nums;
  cursor: default;
}

th.active {
  background: var(--g-accent-weak);
  color: var(--g-accent);
  font-weight: 650;
}

/* Sticky layers, lowest to highest. */
td.fz {
  position: sticky;
  z-index: 2;
}
.fz-row > td {
  position: sticky;
  z-index: 3;
}
.fz-row > td.fz {
  z-index: 4;
}
.row-head {
  position: sticky;
  left: 0;
  z-index: 5;
}
.fz-row > .row-head {
  z-index: 6;
}
thead th {
  position: sticky;
  top: 0;
  z-index: 7;
}
thead th.fz {
  z-index: 8;
}
thead th.corner {
  left: 0;
  z-index: 9;
}

/* Edge of the frozen panes, like Excel's freeze line. */
.fz-row:nth-last-child(1 of .fz-row) > * {
  border-bottom: 1px solid var(--g-freeze);
}
td.fz:nth-last-child(1 of .fz) {
  border-right: 1px solid var(--g-freeze);
}

.spacer td {
  height: auto;
  padding: 0;
  border: 0;
  background: transparent;
}

.al-right {
  text-align: right;
  font-variant-numeric: tabular-nums;
}
.al-center {
  text-align: center;
}
.al-left {
  text-align: left;
}

td.sel {
  box-shadow:
    inset 0 0 0 2px var(--g-accent),
    inset 0 0 0 3px color-mix(in oklab, var(--background) 80%, transparent);
}

/* Small triangle in the corner of cells changed since the last save. */
td.edited::after {
  content: '';
  position: absolute;
  top: 0;
  right: 0;
  border-style: solid;
  border-width: 0 7px 7px 0;
  border-color: transparent #f59e0b transparent transparent;
}

td.editing {
  overflow: visible;
  padding: 0;
}

.editor {
  position: absolute;
  inset: 0;
  z-index: 10;
  box-sizing: border-box;
  width: 100%;
  height: 100%;
  margin: 0;
  padding: 0 6px;
  border: 0;
  outline: 2px solid var(--g-accent);
  outline-offset: -2px;
  resize: none;
  overflow: hidden;
  white-space: pre;
  font: inherit;
  line-height: 27px;
  color: var(--g-text);
  background: var(--g-surface);
  user-select: text;
}
</style>
