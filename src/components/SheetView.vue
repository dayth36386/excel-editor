<script setup lang="ts">
import { computed, nextTick, ref, watch } from 'vue'
import { useI18n } from 'vue-i18n'
import {
  ArrowDownToLineIcon,
  ArrowLeftToLineIcon,
  ArrowRightToLineIcon,
  ArrowUpToLineIcon,
  BetweenHorizontalStartIcon,
  BetweenVerticalStartIcon,
  CopyIcon,
  EraserIcon,
  Trash2Icon,
} from '@lucide/vue'
import SpreadsheetGrid from '@/components/SpreadsheetGrid.vue'
import { Button } from '@/components/ui/button'
import {
  ContextMenu,
  ContextMenuContent,
  ContextMenuItem,
  ContextMenuSeparator,
  ContextMenuShortcut,
  ContextMenuTrigger,
} from '@/components/ui/context-menu'
import { Tooltip, TooltipContent, TooltipTrigger } from '@/components/ui/tooltip'
import type { WorkbookStore } from '@/composables/useWorkbook'
import { address, colName } from '@/lib/excel'

const props = defineProps<{ book: WorkbookStore }>()
const book = props.book
const { extent, frozen, revision } = book
const { t } = useI18n()

const mod = /Mac|iPhone|iPad/.test(navigator.userAgent) ? '⌘' : 'Ctrl+'
const grid = ref<InstanceType<typeof SpreadsheetGrid>>()
const selected = ref({ r: 1, c: 1 })

// Show empty rows and columns past the data, so there is room to type, as in Excel.
const rowCount = computed(() => Math.max(extent.value.rows + 50, 100))
const colCount = computed(() => Math.max(extent.value.cols + 5, 26))
const widths = computed(() => (void revision.value, book.colWidths(colCount.value)))

// ---- formula bar ---------------------------------------------------------------

const barText = ref('')
const barFocused = ref(false)
const gridDraft = ref<string | null>(null)
const target = () => book.resolve(selected.value.r, selected.value.c)
const selAddress = computed(() => address(selected.value.r, selected.value.c))

watch(
  [selected, revision],
  () => {
    if (!barFocused.value) barText.value = book.getEditText(target().r, target().c)
  },
  { immediate: true },
)

// Clicking another cell changes the selection before the bar loses focus,
// so remember which cell the bar was editing.
let barCell: { r: number; c: number } | null = null

function onBarFocus() {
  barFocused.value = true
  barCell = target()
}

function commitBar() {
  const { r, c } = barCell ?? target()
  book.setCell(r, c, barText.value)
}

function onBarKeydown(e: KeyboardEvent) {
  if (e.isComposing) return
  if (e.key === 'Enter') {
    e.preventDefault()
    commitBar()
    const { r, c } = barCell ?? target()
    grid.value?.focus()
    grid.value?.select(r + 1, c)
  } else if (e.key === 'Escape') {
    e.preventDefault()
    barText.value = book.getEditText(target().r, target().c)
    grid.value?.focus()
  }
}

function onBarBlur() {
  commitBar()
  barFocused.value = false
  barCell = null
  barText.value = book.getEditText(target().r, target().c)
}

// ---- rows, columns, cells ----------------------------------------------------------

function act(fn: () => void) {
  grid.value?.commit()
  fn()
  nextTick(() => grid.value?.focus())
}

const insertRow = (below: boolean) => act(() => book.insertRow(selected.value.r + (below ? 1 : 0)))
const deleteRow = () => act(() => book.deleteRows([selected.value.r]))
const insertCol = (right: boolean) => act(() => book.insertColumn(selected.value.c + (right ? 1 : 0)))
const deleteCol = () => act(() => book.deleteColumn(selected.value.c))
const clearCell = () => act(() => book.setCell(target().r, target().c, ''))
const copyCell = () => grid.value?.copy()

defineExpose({ commit: () => grid.value?.commit() })
</script>

<template>
  <div class="space-y-3">
    <!-- formula bar + row/column actions -->
    <div class="flex flex-wrap items-center gap-2">
      <div class="flex h-9 min-w-72 flex-1 items-center overflow-hidden rounded-lg border bg-background shadow-xs">
        <span class="w-20 shrink-0 border-r text-center font-mono text-xs font-semibold tabular-nums">{{
          selAddress
        }}</span>
        <span class="px-2.5 font-serif text-sm text-muted-foreground italic select-none">fx</span>
        <input
          class="h-full min-w-0 flex-1 bg-transparent pr-3 text-sm outline-none focus:shadow-[inset_0_-2px_0_#059669]"
          :value="gridDraft ?? barText"
          :placeholder="t('sheetView.formulaPlaceholder')"
          spellcheck="false"
          @focus="onBarFocus"
          @input="barText = ($event.target as HTMLInputElement).value"
          @keydown="onBarKeydown"
          @blur="onBarBlur"
        />
      </div>

      <div class="flex items-center gap-1">
        <Tooltip>
          <TooltipTrigger as-child>
            <Button variant="outline" size="sm" class="bg-background" @click="insertRow(false)">
              <BetweenHorizontalStartIcon /><span class="hidden md:inline">{{ t('sheetView.insertRow') }}</span>
            </Button>
          </TooltipTrigger>
          <TooltipContent>{{ t('sheetView.insertRowAbove') }}</TooltipContent>
        </Tooltip>
        <Tooltip>
          <TooltipTrigger as-child>
            <Button variant="outline" size="sm" class="bg-background" @click="insertCol(false)">
              <BetweenVerticalStartIcon /><span class="hidden md:inline">{{ t('sheetView.insertCol') }}</span>
            </Button>
          </TooltipTrigger>
          <TooltipContent>{{ t('sheetView.insertColLeft') }}</TooltipContent>
        </Tooltip>
        <Tooltip>
          <TooltipTrigger as-child>
            <Button
              variant="outline"
              size="icon-sm"
              class="bg-background"
              :aria-label="t('sheetView.deleteRow')"
              @click="deleteRow"
            >
              <Trash2Icon class="rotate-90" />
            </Button>
          </TooltipTrigger>
          <TooltipContent>{{ t('sheetView.deleteRow') }} {{ selected.r }}</TooltipContent>
        </Tooltip>
        <Tooltip>
          <TooltipTrigger as-child>
            <Button
              variant="outline"
              size="icon-sm"
              class="bg-background"
              :aria-label="t('sheetView.deleteCol')"
              @click="deleteCol"
            >
              <Trash2Icon />
            </Button>
          </TooltipTrigger>
          <TooltipContent>{{ t('sheetView.deleteCol') }} {{ colName(selected.c) }}</TooltipContent>
        </Tooltip>
      </div>
    </div>

    <!-- the grid; right-click for the Excel-style context menu -->
    <ContextMenu>
      <ContextMenuTrigger as-child>
        <div class="h-[calc(100svh-16.5rem)] min-h-[420px] overflow-hidden rounded-xl border bg-card shadow-xs">
          <SpreadsheetGrid
            ref="grid"
            v-model:selected="selected"
            :row-count="rowCount"
            :col-count="colCount"
            :col-widths="widths"
            :frozen-rows="frozen.rows"
            :frozen-cols="frozen.cols"
            :revision="revision"
            :view="book.gridView"
            :edit-text="book.getEditText"
            :is-edited="book.isEdited"
            :resolve="book.resolve"
            @commit="book.setCell"
            @clear="(r, c) => book.setCell(r, c, '')"
            @paste="book.pasteBlock"
            @editing="gridDraft = $event"
          />
        </div>
      </ContextMenuTrigger>
      <ContextMenuContent class="w-60">
        <ContextMenuItem @select="copyCell">
          <CopyIcon />{{ t('sheetView.copy') }}<ContextMenuShortcut>{{ mod }}C</ContextMenuShortcut>
        </ContextMenuItem>
        <ContextMenuItem @select="clearCell">
          <EraserIcon />{{ t('sheetView.clear') }}<ContextMenuShortcut>Del</ContextMenuShortcut>
        </ContextMenuItem>
        <ContextMenuSeparator />
        <ContextMenuItem @select="insertRow(false)"
          ><ArrowUpToLineIcon />{{ t('sheetView.insertRowAbove') }}</ContextMenuItem
        >
        <ContextMenuItem @select="insertRow(true)"
          ><ArrowDownToLineIcon />{{ t('sheetView.insertRowBelow') }}</ContextMenuItem
        >
        <ContextMenuItem @select="insertCol(false)"
          ><ArrowLeftToLineIcon />{{ t('sheetView.insertColLeft') }}</ContextMenuItem
        >
        <ContextMenuItem @select="insertCol(true)"
          ><ArrowRightToLineIcon />{{ t('sheetView.insertColRight') }}</ContextMenuItem
        >
        <ContextMenuSeparator />
        <ContextMenuItem variant="destructive" @select="deleteRow"
          ><Trash2Icon />{{ t('sheetView.deleteRow') }}</ContextMenuItem
        >
        <ContextMenuItem variant="destructive" @select="deleteCol"
          ><Trash2Icon />{{ t('sheetView.deleteCol') }}</ContextMenuItem
        >
      </ContextMenuContent>
    </ContextMenu>
  </div>
</template>
