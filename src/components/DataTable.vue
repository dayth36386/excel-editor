<script setup lang="ts">
import { computed, nextTick, ref, watch, type Component } from 'vue'
import { useI18n } from 'vue-i18n'
import {
  getCoreRowModel,
  getFilteredRowModel,
  getPaginationRowModel,
  getSortedRowModel,
  useVueTable,
  type ColumnDef,
  type ColumnFiltersState,
  type PaginationState,
  type RowSelectionState,
  type SortingState,
  type VisibilityState,
} from '@tanstack/vue-table'
import {
  ArrowDownIcon,
  ArrowDownToLineIcon,
  ArrowUpDownIcon,
  ArrowUpIcon,
  ArrowUpToLineIcon,
  CalendarIcon,
  ChevronLeftIcon,
  ChevronRightIcon,
  ChevronsLeftIcon,
  ChevronsRightIcon,
  CircleCheckIcon,
  CirclePlusIcon,
  Columns3Icon,
  EllipsisIcon,
  EyeOffIcon,
  HashIcon,
  PencilIcon,
  PlusIcon,
  SearchIcon,
  TagIcon,
  Trash2Icon,
  TypeIcon,
  XIcon,
} from '@lucide/vue'
import { Button } from '@/components/ui/button'
import { Checkbox } from '@/components/ui/checkbox'
import {
  DropdownMenu,
  DropdownMenuCheckboxItem,
  DropdownMenuContent,
  DropdownMenuItem,
  DropdownMenuLabel,
  DropdownMenuSeparator,
  DropdownMenuTrigger,
} from '@/components/ui/dropdown-menu'
import { Input } from '@/components/ui/input'
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from '@/components/ui/select'
import { Separator } from '@/components/ui/separator'
import { Table, TableBody, TableCell, TableHead, TableHeader, TableRow } from '@/components/ui/table'
import { confirmDialog, promptDialog } from '@/composables/useDialogs'
import type { TableColumn, TableRow as Row, WorkbookStore } from '@/composables/useWorkbook'
import { compareSort } from '@/lib/excel'
import { cn, valueUpdater } from '@/lib/utils'

const props = defineProps<{ book: WorkbookStore }>()
const book = props.book
const { columns, rows, revision } = book
const { t } = useI18n()

const PAGE_SIZES = [10, 20, 50, 100]
const MAX_FACETS = 3
const modKey = /Mac|iPhone|iPad/.test(navigator.userAgent) ? '⌘' : 'Ctrl+'

// Category badge colours, assigned by the value's position in the column's sorted options.
const BADGE_STYLES = [
  ['bg-sky-50 text-sky-700 ring-sky-600/15 dark:bg-sky-400/10 dark:text-sky-300 dark:ring-sky-400/25', 'bg-sky-500'],
  [
    'bg-amber-50 text-amber-700 ring-amber-600/15 dark:bg-amber-400/10 dark:text-amber-300 dark:ring-amber-400/25',
    'bg-amber-500',
  ],
  [
    'bg-violet-50 text-violet-700 ring-violet-600/15 dark:bg-violet-400/10 dark:text-violet-300 dark:ring-violet-400/25',
    'bg-violet-500',
  ],
  [
    'bg-emerald-50 text-emerald-700 ring-emerald-600/15 dark:bg-emerald-400/10 dark:text-emerald-300 dark:ring-emerald-400/25',
    'bg-emerald-500',
  ],
  [
    'bg-rose-50 text-rose-700 ring-rose-600/15 dark:bg-rose-400/10 dark:text-rose-300 dark:ring-rose-400/25',
    'bg-rose-500',
  ],
  [
    'bg-teal-50 text-teal-700 ring-teal-600/15 dark:bg-teal-400/10 dark:text-teal-300 dark:ring-teal-400/25',
    'bg-teal-500',
  ],
  [
    'bg-orange-50 text-orange-700 ring-orange-600/15 dark:bg-orange-400/10 dark:text-orange-300 dark:ring-orange-400/25',
    'bg-orange-500',
  ],
  [
    'bg-slate-100 text-slate-700 ring-slate-600/15 dark:bg-slate-400/10 dark:text-slate-300 dark:ring-slate-400/25',
    'bg-slate-500',
  ],
] as const

function badgeStyle(col: TableColumn, value: string) {
  let i = col.options?.indexOf(value) ?? -1
  if (i < 0) i = [...value].reduce((h, ch) => (h * 31 + ch.charCodeAt(0)) >>> 0, 7)
  return BADGE_STYLES[i % BADGE_STYLES.length]
}

const TYPE_ICONS: Record<string, Component> = {
  text: TypeIcon,
  number: HashIcon,
  date: CalendarIcon,
  boolean: CircleCheckIcon,
  category: TagIcon,
}
const typeOf = (col: TableColumn) => (col.options ? 'category' : col.kind)
const titleOf = (col: TableColumn) => col.title || t('table.untitled', { letter: col.letter })

// ---- table state -------------------------------------------------------------

const sorting = ref<SortingState>([])
const columnVisibility = ref<VisibilityState>({})
const columnFilters = ref<ColumnFiltersState>([])
const rowSelection = ref<RowSelectionState>({})
const globalFilter = ref('')
const pagination = ref<PaginationState>({ pageIndex: 0, pageSize: 20 })

const colByKey = computed(() => new Map(columns.value.map((c) => [c.key, c])))
const categoryColumns = computed(() => columns.value.filter((c) => c.options))
const facetColumns = computed(() => categoryColumns.value.slice(0, MAX_FACETS))

const columnDefs = computed<ColumnDef<Row>[]>(() => [
  { id: 'select', enableSorting: false, enableHiding: false },
  { id: 'row', accessorFn: (r) => r.id, enableHiding: true },
  ...columns.value.map((col): ColumnDef<Row> => ({
    id: col.key,
    accessorFn: (r) => r.cells[col.key]?.sort,
    sortingFn: (a, b, id) => compareSort(a.getValue(id), b.getValue(id)),
    sortUndefined: 'last',
    filterFn: (row, id, values: string[]) => !values?.length || values.includes(row.original.cells[id]?.text ?? ''),
  })),
  { id: 'actions', enableSorting: false, enableHiding: false },
])

const table = useVueTable({
  get data() {
    return rows.value
  },
  get columns() {
    return columnDefs.value
  },
  getRowId: (row) => String(row.id),
  state: {
    get sorting() {
      return sorting.value
    },
    get columnVisibility() {
      return columnVisibility.value
    },
    get columnFilters() {
      return columnFilters.value
    },
    get rowSelection() {
      return rowSelection.value
    },
    get globalFilter() {
      return globalFilter.value
    },
    get pagination() {
      return pagination.value
    },
  },
  onSortingChange: (u) => valueUpdater(u, sorting),
  onColumnVisibilityChange: (u) => valueUpdater(u, columnVisibility),
  onColumnFiltersChange: (u) => valueUpdater(u, columnFilters),
  onRowSelectionChange: (u) => valueUpdater(u, rowSelection),
  onGlobalFilterChange: (u) => valueUpdater(u, globalFilter),
  onPaginationChange: (u) => valueUpdater(u, pagination),
  getColumnCanGlobalFilter: (column) => colByKey.value.has(column.id),
  globalFilterFn: (row, columnId, query: string) => {
    const q = query.trim().toLowerCase()
    return !q || !!row.original.cells[columnId]?.text.toLowerCase().includes(q)
  },
  // Editing a cell replaces the data; stay on the same page when that happens.
  autoResetPageIndex: false,
  getCoreRowModel: getCoreRowModel(),
  getFilteredRowModel: getFilteredRowModel(),
  getSortedRowModel: getSortedRowModel(),
  getPaginationRowModel: getPaginationRowModel(),
})

watch([globalFilter, columnFilters], () => table.setPageIndex(0), { deep: true })

// Inserting or deleting rows renumbers them, so a selection by row number goes stale.
watch(
  () => rows.value.length,
  () => {
    rowSelection.value = {}
    const last = Math.max(0, table.getPageCount() - 1)
    if (pagination.value.pageIndex > last) table.setPageIndex(last)
  },
)

const selectedIds = computed(() =>
  Object.entries(rowSelection.value)
    .filter(([, on]) => on)
    .map(([id]) => Number(id)),
)
const filteredCount = computed(() => table.getFilteredRowModel().rows.length)
const isFiltered = computed(() => !!globalFilter.value || columnFilters.value.length > 0)
const pageAllState = computed(() =>
  table.getIsAllPageRowsSelected() ? true : table.getIsSomePageRowsSelected() ? 'indeterminate' : false,
)
const hideableColumns = computed(() => table.getAllLeafColumns().filter((c) => c.getCanHide()))
const visibleDataColumns = computed(() =>
  table
    .getVisibleLeafColumns()
    .map((c) => colByKey.value.get(c.id))
    .filter((c): c is TableColumn => !!c),
)

/** How many rows hold each value of a category column (for the filter menus). */
const facetCounts = computed(() => {
  const out = new Map<string, Map<string, number>>()
  for (const col of facetColumns.value) {
    const counts = new Map<string, number>()
    for (const row of rows.value) {
      const v = row.cells[col.key]?.text
      if (v) counts.set(v, (counts.get(v) ?? 0) + 1)
    }
    out.set(col.key, counts)
  }
  return out
})

const filterValues = (key: string) => (table.getColumn(key)?.getFilterValue() as string[] | undefined) ?? []

function toggleFilter(key: string, value: string, on: boolean) {
  const next = new Set(filterValues(key))
  if (on) next.add(value)
  else next.delete(value)
  table.getColumn(key)?.setFilterValue(next.size ? [...next] : undefined)
}

function resetFilters() {
  globalFilter.value = ''
  table.resetColumnFilters()
}

/** Re-render edit markers when the store changes (e.g. cleared on save). */
function isEdited(r: number, c: number) {
  void revision.value
  return book.isEdited(r, c)
}

// ---- inline editing ----------------------------------------------------------

const editing = ref<{ r: number; c: number } | null>(null)
const draft = ref('')
let editorEl: HTMLInputElement | null = null
const setEditor = (el: unknown) => (editorEl = el as HTMLInputElement | null)
const isEditing = (r: number, c: number) => editing.value?.r === r && editing.value?.c === c

function startEdit(r: number, c: number) {
  editing.value = { r, c }
  draft.value = book.getEditText(r, c)
  nextTick(() => {
    editorEl?.focus()
    editorEl?.select()
  })
}

function commit(move?: 1 | -1) {
  const ed = editing.value
  if (!ed) return
  editing.value = null
  book.setCell(ed.r, ed.c, draft.value)
  if (!move) return
  const cols = visibleDataColumns.value
  const next = cols[cols.findIndex((col) => col.col === ed.c) + move]
  if (next) startEdit(ed.r, next.col)
}

// Moving to the next cell removes the old input, which fires its blur after the
// new cell is already being edited; only commit if the blur is for that cell.
function onEditorBlur(r: number, c: number) {
  if (isEditing(r, c)) commit()
}

function onEditorKeydown(e: KeyboardEvent) {
  if (e.isComposing) return
  if (e.key === 'Enter') {
    e.preventDefault()
    commit()
  } else if (e.key === 'Tab') {
    e.preventDefault()
    commit(e.shiftKey ? -1 : 1)
  } else if (e.key === 'Escape') {
    e.preventDefault()
    editing.value = null
  }
}

// ---- rows & columns ----------------------------------------------------------

const flashRow = ref(0)

async function addRow(at?: number) {
  resetFilters()
  const r = book.insertRow(at)
  if (!r) return
  await nextTick()
  // Jump to the page that shows the new row and start typing in it.
  const index = table.getPrePaginationRowModel().rows.findIndex((row) => row.original.id === r)
  if (index >= 0) table.setPageIndex(Math.floor(index / pagination.value.pageSize))
  flashRow.value = r
  setTimeout(() => flashRow.value === r && (flashRow.value = 0), 1600)
  const first = visibleDataColumns.value[0]
  if (first) startEdit(r, first.col)
}

async function deleteRows(ids: number[]) {
  if (!ids.length) return
  const ok = await confirmDialog({
    title:
      ids.length === 1 ? t('dialogs.deleteRowTitle', { n: ids[0] }) : t('dialogs.deleteRowsTitle', { n: ids.length }),
    description: t('dialogs.deleteRowsText', { key: modKey }),
    confirmText: t('dialogs.delete'),
    destructive: true,
  })
  if (ok) book.deleteRows(ids)
}

async function addColumn() {
  const title = await promptDialog({
    title: t('dialogs.addColumnTitle'),
    description: t('dialogs.addColumnText'),
    label: t('dialogs.columnName'),
    placeholder: t('dialogs.columnPlaceholder'),
    confirmText: t('dialogs.add'),
  })
  if (title) book.addColumn(title)
}

async function renameColumn(col: TableColumn) {
  const title = await promptDialog({
    title: t('dialogs.renameColumnTitle'),
    label: t('dialogs.columnName'),
    value: col.title,
    confirmText: t('dialogs.save'),
  })
  if (title) book.renameColumn(col.col, title)
}

async function deleteColumn(col: TableColumn) {
  const ok = await confirmDialog({
    title: t('dialogs.deleteColumnTitle', { name: titleOf(col) }),
    description: t('dialogs.deleteColumnText'),
    confirmText: t('dialogs.deleteColumnConfirm'),
    destructive: true,
  })
  if (ok) book.deleteColumn(col.col)
}

defineExpose({ commit: () => commit(), addRow })
</script>

<template>
  <div class="space-y-3">
    <!-- suggestions for category columns while typing -->
    <datalist v-for="col in categoryColumns" :id="`options-${col.key}`" :key="col.key">
      <option v-for="o in col.options" :key="o" :value="o" />
    </datalist>

    <!-- toolbar -->
    <div class="flex flex-wrap items-center gap-2">
      <div class="relative w-full sm:w-64">
        <SearchIcon
          class="pointer-events-none absolute top-1/2 left-2.5 size-4 -translate-y-1/2 text-muted-foreground"
        />
        <Input v-model="globalFilter" :placeholder="t('table.search')" class="h-8 bg-background pr-8 pl-8" />
        <button
          v-if="globalFilter"
          type="button"
          class="absolute top-1/2 right-2 -translate-y-1/2 text-muted-foreground hover:text-foreground"
          :aria-label="t('table.clearSearch')"
          @click="globalFilter = ''"
        >
          <XIcon class="size-4" />
        </button>
      </div>

      <!-- category filters, shadcn "faceted filter" style -->
      <DropdownMenu v-for="col in facetColumns" :key="col.key">
        <DropdownMenuTrigger as-child>
          <Button variant="outline" size="sm" class="border-dashed bg-background">
            <CirclePlusIcon />{{ titleOf(col) }}
            <template v-if="filterValues(col.key).length">
              <Separator orientation="vertical" class="mx-0.5 h-4!" />
              <span v-if="filterValues(col.key).length > 2" class="rounded-sm bg-secondary px-1 text-xs font-normal">
                {{ t('table.filterSelected', { n: filterValues(col.key).length }) }}
              </span>
              <template v-else>
                <span
                  v-for="v in filterValues(col.key)"
                  :key="v"
                  class="rounded-sm bg-secondary px-1 text-xs font-normal"
                >
                  {{ v }}
                </span>
              </template>
            </template>
          </Button>
        </DropdownMenuTrigger>
        <DropdownMenuContent align="start" class="w-56">
          <DropdownMenuLabel class="text-xs font-normal text-muted-foreground">{{ titleOf(col) }}</DropdownMenuLabel>
          <DropdownMenuCheckboxItem
            v-for="o in col.options"
            :key="o"
            :model-value="filterValues(col.key).includes(o)"
            @update:model-value="(v) => toggleFilter(col.key, o, !!v)"
            @select.prevent
          >
            <span :class="cn('size-2 rounded-full', badgeStyle(col, o)[1])" />
            <span class="truncate">{{ o }}</span>
            <span class="ml-auto pl-3 font-mono text-xs text-muted-foreground tabular-nums">
              {{ facetCounts.get(col.key)?.get(o) ?? 0 }}
            </span>
          </DropdownMenuCheckboxItem>
          <template v-if="filterValues(col.key).length">
            <DropdownMenuSeparator />
            <DropdownMenuItem class="justify-center" @select="table.getColumn(col.key)?.setFilterValue(undefined)">
              {{ t('table.clearFilter') }}
            </DropdownMenuItem>
          </template>
        </DropdownMenuContent>
      </DropdownMenu>

      <Button v-if="isFiltered" variant="ghost" size="sm" @click="resetFilters">
        {{ t('table.reset') }}<XIcon />
      </Button>

      <Button v-if="selectedIds.length" variant="destructive" size="sm" @click="deleteRows(selectedIds)">
        <Trash2Icon />{{ t('table.deleteSelected', { n: selectedIds.length }, selectedIds.length) }}
      </Button>

      <div class="ml-auto flex items-center gap-2">
        <DropdownMenu>
          <DropdownMenuTrigger as-child>
            <Button variant="outline" size="sm" class="bg-background">
              <Columns3Icon /><span class="hidden sm:inline">{{ t('table.columns') }}</span>
            </Button>
          </DropdownMenuTrigger>
          <DropdownMenuContent align="end" class="max-h-80 w-56 overflow-y-auto">
            <DropdownMenuLabel>{{ t('table.showColumns') }}</DropdownMenuLabel>
            <DropdownMenuSeparator />
            <DropdownMenuCheckboxItem
              v-for="column in hideableColumns"
              :key="column.id"
              :model-value="column.getIsVisible()"
              @update:model-value="(v) => column.toggleVisibility(!!v)"
              @select.prevent
            >
              <span class="truncate">
                {{ colByKey.get(column.id) ? titleOf(colByKey.get(column.id)!) : t('table.rowNumber') }}
              </span>
            </DropdownMenuCheckboxItem>
          </DropdownMenuContent>
        </DropdownMenu>
        <Button variant="outline" size="sm" class="bg-background" @click="addColumn">
          <PlusIcon />{{ t('table.addColumn') }}
        </Button>
        <Button size="sm" @click="addRow()"><PlusIcon />{{ t('table.addRow') }}</Button>
      </div>
    </div>

    <!-- table -->
    <div class="overflow-hidden rounded-xl border bg-card shadow-xs">
      <Table>
        <TableHeader class="bg-muted/40">
          <TableRow v-for="group in table.getHeaderGroups()" :key="group.id" class="hover:bg-transparent">
            <TableHead
              v-for="header in group.headers"
              :key="header.id"
              :class="
                cn(
                  'h-11',
                  header.column.id === 'select' && 'w-10 pl-4',
                  header.column.id === 'row' && 'w-14',
                  header.column.id === 'actions' && 'w-12',
                  colByKey.get(header.column.id)?.kind === 'number' && 'text-right',
                )
              "
            >
              <Checkbox
                v-if="header.column.id === 'select'"
                :model-value="pageAllState"
                :aria-label="t('table.selectPage')"
                @update:model-value="(v) => table.toggleAllPageRowsSelected(!!v)"
              />

              <button
                v-else-if="header.column.id === 'row'"
                type="button"
                class="inline-flex items-center gap-1 text-xs text-muted-foreground hover:text-foreground"
                :title="t('table.sortByRow')"
                @click="header.column.toggleSorting(header.column.getIsSorted() === 'asc')"
              >
                #
                <ArrowUpIcon v-if="header.column.getIsSorted() === 'asc'" class="size-3" />
                <ArrowDownIcon v-else-if="header.column.getIsSorted() === 'desc'" class="size-3" />
              </button>

              <template v-else-if="colByKey.get(header.column.id)">
                <template v-for="col in [colByKey.get(header.column.id)!]" :key="col.key">
                  <DropdownMenu>
                    <DropdownMenuTrigger as-child>
                      <Button
                        variant="ghost"
                        size="sm"
                        :class="
                          cn(
                            '-ml-2 h-8 max-w-64 gap-1.5 font-medium text-foreground data-[state=open]:bg-accent',
                            col.kind === 'number' && 'mr-[-0.5rem] ml-0',
                          )
                        "
                      >
                        <component
                          :is="TYPE_ICONS[typeOf(col)]"
                          class="size-3.5! text-muted-foreground"
                          :aria-label="t(`table.types.${typeOf(col)}`)"
                        />
                        <span class="truncate">{{ titleOf(col) }}</span>
                        <ArrowUpIcon v-if="header.column.getIsSorted() === 'asc'" class="size-3.5!" />
                        <ArrowDownIcon v-else-if="header.column.getIsSorted() === 'desc'" class="size-3.5!" />
                        <ArrowUpDownIcon v-else class="size-3.5! opacity-30" />
                      </Button>
                    </DropdownMenuTrigger>
                    <DropdownMenuContent align="start" class="w-52">
                      <DropdownMenuLabel class="flex items-center gap-1.5 text-xs font-normal text-muted-foreground">
                        <component :is="TYPE_ICONS[typeOf(col)]" class="size-3.5" />
                        {{ t(`table.types.${typeOf(col)}`) }} · {{ col.letter }}
                      </DropdownMenuLabel>
                      <DropdownMenuItem @select="header.column.toggleSorting(false)">
                        <ArrowUpIcon />{{ t('table.sortAsc') }}
                      </DropdownMenuItem>
                      <DropdownMenuItem @select="header.column.toggleSorting(true)">
                        <ArrowDownIcon />{{ t('table.sortDesc') }}
                      </DropdownMenuItem>
                      <DropdownMenuItem v-if="header.column.getIsSorted()" @select="header.column.clearSorting()">
                        <XIcon />{{ t('table.clearSort') }}
                      </DropdownMenuItem>
                      <DropdownMenuSeparator />
                      <DropdownMenuItem @select="renameColumn(col)"
                        ><PencilIcon />{{ t('table.rename') }}</DropdownMenuItem
                      >
                      <DropdownMenuItem @select="header.column.toggleVisibility(false)">
                        <EyeOffIcon />{{ t('table.hide') }}
                      </DropdownMenuItem>
                      <DropdownMenuSeparator />
                      <DropdownMenuItem variant="destructive" @select="deleteColumn(col)">
                        <Trash2Icon />{{ t('table.deleteColumn') }}
                      </DropdownMenuItem>
                    </DropdownMenuContent>
                  </DropdownMenu>
                </template>
              </template>
            </TableHead>
          </TableRow>
        </TableHeader>

        <TableBody>
          <template v-if="table.getRowModel().rows.length">
            <TableRow
              v-for="row in table.getRowModel().rows"
              :key="row.id"
              :data-state="row.getIsSelected() ? 'selected' : undefined"
              :class="cn('group transition-colors duration-700', flashRow === row.original.id && 'bg-emerald-500/10')"
            >
              <TableCell
                v-for="cell in row.getVisibleCells()"
                :key="cell.id"
                :class="
                  cn('py-1', cell.column.id === 'select' && 'pl-4', cell.column.id === 'actions' && 'pr-3 text-right')
                "
              >
                <Checkbox
                  v-if="cell.column.id === 'select'"
                  :model-value="row.getIsSelected()"
                  :aria-label="t('table.selectRow', { n: row.original.id })"
                  @update:model-value="(v) => row.toggleSelected(!!v)"
                />

                <span
                  v-else-if="cell.column.id === 'row'"
                  class="font-mono text-xs text-muted-foreground/70 tabular-nums"
                >
                  {{ row.original.id }}
                </span>

                <DropdownMenu v-else-if="cell.column.id === 'actions'">
                  <DropdownMenuTrigger as-child>
                    <Button
                      variant="ghost"
                      size="icon-sm"
                      class="text-muted-foreground opacity-50 group-hover:opacity-100 data-[state=open]:opacity-100"
                      :aria-label="t('table.rowActions')"
                    >
                      <EllipsisIcon />
                    </Button>
                  </DropdownMenuTrigger>
                  <DropdownMenuContent align="end" class="w-48">
                    <DropdownMenuLabel class="text-xs font-normal text-muted-foreground">
                      {{ t('table.rowLabel', { n: row.original.id }) }}
                    </DropdownMenuLabel>
                    <DropdownMenuItem @select="addRow(row.original.id)">
                      <ArrowUpToLineIcon />{{ t('table.insertAbove') }}
                    </DropdownMenuItem>
                    <DropdownMenuItem @select="addRow(row.original.id + 1)">
                      <ArrowDownToLineIcon />{{ t('table.insertBelow') }}
                    </DropdownMenuItem>
                    <DropdownMenuSeparator />
                    <DropdownMenuItem variant="destructive" @select="deleteRows([row.original.id])">
                      <Trash2Icon />{{ t('table.deleteRow') }}
                    </DropdownMenuItem>
                  </DropdownMenuContent>
                </DropdownMenu>

                <template v-else-if="colByKey.get(cell.column.id)">
                  <template v-for="col in [colByKey.get(cell.column.id)!]" :key="col.key">
                    <input
                      v-if="isEditing(row.original.id, col.col)"
                      :ref="setEditor"
                      v-model="draft"
                      :list="col.options ? `options-${col.key}` : undefined"
                      :class="
                        cn(
                          'h-8 w-full min-w-32 rounded-md border border-emerald-500 bg-background px-2 text-sm ring-3 ring-emerald-500/20 outline-none',
                          col.kind === 'number' && 'text-right tabular-nums',
                        )
                      "
                      spellcheck="false"
                      @keydown="onEditorKeydown"
                      @blur="onEditorBlur(row.original.id, col.col)"
                    />

                    <div v-else-if="col.kind === 'boolean'" class="flex h-8 items-center px-2">
                      <Checkbox
                        :model-value="row.original.cells[col.key]?.sort === 1"
                        :aria-label="titleOf(col)"
                        @update:model-value="(v) => book.setCell(row.original.id, col.col, v ? 'TRUE' : 'FALSE')"
                      />
                    </div>

                    <button
                      v-else
                      type="button"
                      :title="
                        row.original.cells[col.key]?.formula
                          ? t('table.formula', { formula: book.getEditText(row.original.id, col.col) })
                          : undefined
                      "
                      :class="
                        cn(
                          'relative flex h-8 w-full max-w-80 min-w-16 cursor-text items-center gap-1.5 rounded-md px-2 text-left outline-none hover:bg-muted hover:ring-1 hover:ring-border focus-visible:ring-3 focus-visible:ring-emerald-500/30',
                          (col.kind === 'number' || col.kind === 'date') &&
                            'justify-end font-mono text-[13px] tabular-nums',
                          isEdited(row.original.id, col.col) &&
                            'bg-amber-50 hover:bg-amber-100/70 dark:bg-amber-400/10 dark:hover:bg-amber-400/15',
                        )
                      "
                      @click="startEdit(row.original.id, col.col)"
                    >
                      <template v-if="row.original.cells[col.key]">
                        <span
                          v-if="col.options"
                          :class="
                            cn(
                              'inline-flex max-w-full items-center gap-1.5 truncate rounded-full px-2 py-0.5 text-xs font-medium ring-1 ring-inset',
                              badgeStyle(col, row.original.cells[col.key]!.text)[0],
                            )
                          "
                        >
                          <span
                            :class="
                              cn(
                                'size-1.5 shrink-0 rounded-full',
                                badgeStyle(col, row.original.cells[col.key]!.text)[1],
                              )
                            "
                          />
                          {{ row.original.cells[col.key]!.text }}
                        </span>
                        <span v-else class="truncate">{{ row.original.cells[col.key]!.text }}</span>
                      </template>
                      <span v-else class="text-muted-foreground/0 group-hover:text-muted-foreground/40">—</span>
                      <span
                        v-if="row.original.cells[col.key]?.formula"
                        class="shrink-0 rounded bg-muted px-1 font-mono text-[10px] text-muted-foreground italic"
                      >
                        fx
                      </span>
                      <span
                        v-if="isEdited(row.original.id, col.col)"
                        class="absolute top-1 right-1 size-1.5 rounded-full bg-amber-500"
                      />
                    </button>
                  </template>
                </template>
              </TableCell>
            </TableRow>
          </template>

          <TableRow v-else class="hover:bg-transparent">
            <TableCell :colspan="table.getVisibleLeafColumns().length" class="h-40 text-center text-muted-foreground">
              <template v-if="isFiltered">
                <p>{{ globalFilter ? t('table.noResults', { query: globalFilter }) : t('table.noRows') }}</p>
                <Button variant="link" size="sm" @click="resetFilters">{{ t('table.reset') }}</Button>
              </template>
              <template v-else>
                <p>{{ t('table.noRows') }}</p>
                <Button variant="link" size="sm" @click="addRow()">{{ t('table.addFirstRow') }}</Button>
              </template>
            </TableCell>
          </TableRow>
        </TableBody>
      </Table>
    </div>

    <!-- footer -->
    <div class="flex flex-wrap items-center justify-between gap-4 px-1">
      <p class="text-sm text-muted-foreground">
        <template v-if="selectedIds.length">
          {{ t('table.selected', { n: selectedIds.length, total: filteredCount.toLocaleString() }) }}
        </template>
        <template v-else-if="isFiltered">
          {{ t('table.summaryFiltered', { n: filteredCount.toLocaleString(), total: rows.length.toLocaleString() }) }}
        </template>
        <template v-else>{{ t('table.summary', { n: filteredCount.toLocaleString() }, filteredCount) }}</template>
      </p>

      <div class="flex flex-wrap items-center gap-x-6 gap-y-2">
        <div class="flex items-center gap-2 text-sm">
          <span class="text-muted-foreground">{{ t('table.rowsPerPage') }}</span>
          <Select :model-value="String(pagination.pageSize)" @update:model-value="(v) => table.setPageSize(Number(v))">
            <SelectTrigger size="sm" class="w-18 bg-background"><SelectValue /></SelectTrigger>
            <SelectContent side="top">
              <SelectItem v-for="n in PAGE_SIZES" :key="n" :value="String(n)">{{ n }}</SelectItem>
            </SelectContent>
          </Select>
        </div>
        <span class="text-sm font-medium tabular-nums">
          {{ t('table.page', { page: pagination.pageIndex + 1, pages: Math.max(1, table.getPageCount()) }) }}
        </span>
        <div class="flex items-center gap-1">
          <Button
            variant="outline"
            size="icon-sm"
            class="bg-background"
            :disabled="!table.getCanPreviousPage()"
            :aria-label="t('table.first')"
            @click="table.setPageIndex(0)"
          >
            <ChevronsLeftIcon />
          </Button>
          <Button
            variant="outline"
            size="icon-sm"
            class="bg-background"
            :disabled="!table.getCanPreviousPage()"
            :aria-label="t('table.prev')"
            @click="table.previousPage()"
          >
            <ChevronLeftIcon />
          </Button>
          <Button
            variant="outline"
            size="icon-sm"
            class="bg-background"
            :disabled="!table.getCanNextPage()"
            :aria-label="t('table.next')"
            @click="table.nextPage()"
          >
            <ChevronRightIcon />
          </Button>
          <Button
            variant="outline"
            size="icon-sm"
            class="bg-background"
            :disabled="!table.getCanNextPage()"
            :aria-label="t('table.last')"
            @click="table.setPageIndex(table.getPageCount() - 1)"
          >
            <ChevronsRightIcon />
          </Button>
        </div>
      </div>
    </div>
  </div>
</template>
