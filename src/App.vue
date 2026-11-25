<script setup lang="ts">
import { computed, onBeforeUnmount, onMounted, ref } from 'vue'
import { useI18n } from 'vue-i18n'
import { useColorMode, useLocalStorage } from '@vueuse/core'
import { toast } from 'vue-sonner'
import {
  ChevronDownIcon,
  Columns3Icon,
  DownloadIcon,
  FilePlusIcon,
  FileSpreadsheetIcon,
  FileUpIcon,
  FolderOpenIcon,
  Grid3x3Icon,
  LayersIcon,
  MonitorIcon,
  MoonIcon,
  PaletteIcon,
  PencilLineIcon,
  PlusIcon,
  Redo2Icon,
  Rows3Icon,
  SaveIcon,
  SheetIcon,
  ShieldCheckIcon,
  SparklesIcon,
  SunIcon,
  Table2Icon,
  Undo2Icon,
} from '@lucide/vue'
import DataTable from '@/components/DataTable.vue'
import DialogHost from '@/components/DialogHost.vue'
import SheetView from '@/components/SheetView.vue'
import { Badge } from '@/components/ui/badge'
import { Button } from '@/components/ui/button'
import {
  DropdownMenu,
  DropdownMenuContent,
  DropdownMenuItem,
  DropdownMenuRadioGroup,
  DropdownMenuRadioItem,
  DropdownMenuSeparator,
  DropdownMenuShortcut,
  DropdownMenuTrigger,
} from '@/components/ui/dropdown-menu'
import { Separator } from '@/components/ui/separator'
import { Toaster } from '@/components/ui/sonner'
import { Tabs, TabsList, TabsTrigger } from '@/components/ui/tabs'
import { Tooltip, TooltipContent, TooltipProvider, TooltipTrigger } from '@/components/ui/tooltip'
import { confirmDialog, promptDialog } from '@/composables/useDialogs'
import { SheetNameError, useWorkbook, type SaveResult } from '@/composables/useWorkbook'
import { LOCALES, setLocale, type Locale } from '@/i18n'
import { canSaveAs, canWriteBack, fileFromDrop, pickExcelFile, WritePermissionError } from '@/lib/fileAccess'
import { cn } from '@/lib/utils'

const { t, locale } = useI18n()
const book = useWorkbook()
const {
  workbook,
  fileName,
  fileHandle,
  dirty,
  sheets,
  activeId,
  headerRow,
  columns,
  rows,
  editedCount,
  canUndo,
  canRedo,
} = book

const tableRef = ref<InstanceType<typeof DataTable>>()
const sheetRef = ref<InstanceType<typeof SheetView>>()

/** "table" = shadcn data table, "sheet" = Excel-style grid. Both edit the same workbook. */
type ViewMode = 'table' | 'sheet'
const viewMode = useLocalStorage<ViewMode>('excel-editor.view', 'table')
const views = computed(() => [
  { value: 'table' as const, label: t('view.table'), hint: t('view.tableHint'), icon: Table2Icon },
  { value: 'sheet' as const, label: t('view.sheet'), hint: t('view.sheetHint'), icon: Grid3x3Icon },
])

/** Finish any cell being typed into before something else changes the workbook. */
function commitEdits() {
  tableRef.value?.commit()
  sheetRef.value?.commit()
}

function setView(mode: ViewMode) {
  commitEdits()
  viewMode.value = mode
}
const fileInput = ref<HTMLInputElement>()
const saving = ref(false)
const dragging = ref(false)

const colorMode = useColorMode({ storageKey: 'excel-editor.theme' })
const themeIcon = computed(() =>
  colorMode.value === 'dark' ? MoonIcon : colorMode.value === 'light' ? SunIcon : MonitorIcon,
)

const mod = /Mac|iPhone|iPad/.test(navigator.userAgent) ? '⌘' : 'Ctrl+'
const version = __APP_VERSION__
const sheetName = computed(() => sheets.value.find((s) => s.id === activeId.value)?.name ?? '')
const errorText = (e: unknown) => (e instanceof Error ? e.message : String(e))
const defaultHeaders = () => [1, 2, 3].map((n) => t('newSheet.column', { n }))

const stats = computed(() => [
  {
    label: t('stats.rows'),
    value: rows.value.length.toLocaleString(),
    icon: Rows3Icon,
    tone: 'text-sky-600 bg-sky-500/10',
  },
  {
    label: t('stats.columns'),
    value: columns.value.length.toLocaleString(),
    icon: Columns3Icon,
    tone: 'text-violet-600 bg-violet-500/10',
  },
  {
    label: t('stats.edited'),
    value: editedCount.value.toLocaleString(),
    icon: PencilLineIcon,
    tone: editedCount.value ? 'text-amber-600 bg-amber-500/10' : 'text-muted-foreground bg-muted',
  },
  {
    label: t('stats.sheets'),
    value: sheets.value.length.toLocaleString(),
    icon: LayersIcon,
    tone: 'text-emerald-600 bg-emerald-500/10',
  },
])

const features = computed(() => [
  { icon: ShieldCheckIcon, title: t('hero.privateTitle'), text: t('hero.privateText') },
  { icon: PaletteIcon, title: t('hero.formatTitle'), text: t('hero.formatText') },
  { icon: SaveIcon, title: t('hero.saveTitle'), text: canWriteBack ? t('hero.saveText') : t('hero.saveTextFallback') },
])

// ---- opening -----------------------------------------------------------------

async function confirmDiscard() {
  if (!dirty.value) return true
  return confirmDialog({
    title: t('dialogs.discardTitle'),
    description: t('dialogs.discardText', { name: fileName.value }),
    confirmText: t('dialogs.discard'),
    destructive: true,
  })
}

async function openFrom(file: File, handle: FileSystemFileHandle | null) {
  if (/\.xls$/i.test(file.name)) return toast.error(t('toast.xlsNotSupported'), { description: t('toast.xlsHint') })
  if (!/\.xlsx$/i.test(file.name)) return toast.error(t('toast.xlsxOnly'))
  const id = toast.loading(t('toast.opening', { name: file.name }))
  try {
    await book.load(await file.arrayBuffer(), file.name, handle)
    toast.success(t('toast.opened', { name: file.name }), { id })
  } catch (e) {
    toast.error(t('toast.openFailed'), { id, description: errorText(e) })
  }
}

async function openFile() {
  if (!(await confirmDiscard())) return
  if (!canWriteBack) return fileInput.value?.click()
  try {
    const handle = await pickExcelFile()
    if (handle) await openFrom(await handle.getFile(), handle)
  } catch (e) {
    toast.error(t('toast.openFailed'), { description: errorText(e) })
  }
}

function onFileInput(e: Event) {
  const input = e.target as HTMLInputElement
  const file = input.files?.[0]
  input.value = ''
  if (file) void openFrom(file, null)
}

const isFileDrag = (e: DragEvent) => e.dataTransfer?.types.includes('Files') ?? false

function onDragOver(e: DragEvent) {
  if (!isFileDrag(e)) return
  e.preventDefault()
  dragging.value = true
}

async function onDrop(e: DragEvent) {
  if (!isFileDrag(e)) return
  e.preventDefault()
  dragging.value = false
  const pending = fileFromDrop(e) // must start before anything else awaits
  if (!(await confirmDiscard())) return
  const dropped = await pending
  if (dropped) await openFrom(dropped.file, dropped.handle)
}

async function newFile() {
  if (!(await confirmDiscard())) return
  await book.createBlank(defaultHeaders())
}

async function openSample() {
  if (!(await confirmDiscard())) return
  const id = toast.loading(t('toast.openingSample'))
  try {
    const res = await fetch(`${import.meta.env.BASE_URL}sample.xlsx`)
    if (!res.ok) throw new Error(`${res.status} ${res.statusText}`)
    await book.load(await res.arrayBuffer(), 'sample.xlsx')
    toast.dismiss(id)
  } catch (e) {
    toast.error(t('toast.sampleFailed'), { id, description: errorText(e) })
  }
}

// ---- saving ------------------------------------------------------------------

async function save(kind: 'save' | 'saveAs' | 'download') {
  if (!workbook.value || saving.value) return
  commitEdits()
  saving.value = true
  try {
    const run = kind === 'save' ? book.save : kind === 'saveAs' ? book.saveAs : book.download
    const result: SaveResult = await run()
    if (result === 'written') toast.success(t('toast.saved', { name: fileName.value }))
    else if (result === 'downloaded') toast.success(t('toast.downloaded', { name: fileName.value }))
  } catch (e) {
    toast.error(t('toast.saveFailed'), {
      description: e instanceof WritePermissionError ? t('toast.permissionDenied') : errorText(e),
      action: { label: t('toast.downloadInstead'), onClick: () => void save('download') },
    })
  } finally {
    saving.value = false
  }
}

const saveHint = computed(() =>
  fileHandle.value
    ? t('header.saveHintOverwrite', { name: fileName.value })
    : canSaveAs
      ? t('header.saveHintPick')
      : t('header.saveHintDownload', { name: fileName.value }),
)

// ---- sheets ------------------------------------------------------------------

function switchSheet(id: number) {
  if (id === activeId.value) return
  commitEdits()
  book.selectSheet(id)
}

async function renameSheet(id: number, current: string) {
  const name = await promptDialog({
    title: t('sheet.renameTitle'),
    label: t('sheet.nameLabel'),
    value: current,
    confirmText: t('dialogs.save'),
  })
  if (!name) return
  try {
    book.renameSheet(id, name)
  } catch (e) {
    if (!(e instanceof SheetNameError)) throw e
    toast.error(e.reason === 'duplicate' ? t('sheet.duplicateName', { name }) : t('sheet.invalidName'))
  }
}

// ---- shortcuts ---------------------------------------------------------------

// The Excel grid keeps focus in a hidden textarea (data-grid-sink); shortcuts still apply there.
const isTextTarget = (t: EventTarget | null) =>
  t instanceof HTMLElement &&
  !('gridSink' in t.dataset) &&
  (t.tagName === 'INPUT' || t.tagName === 'TEXTAREA' || t.isContentEditable)

function onGlobalKeydown(e: KeyboardEvent) {
  if (!(e.metaKey || e.ctrlKey) || e.altKey) return
  const k = e.key.toLowerCase()
  if (k === 's') {
    e.preventDefault()
    void save(e.shiftKey ? 'saveAs' : 'save')
  } else if (k === 'o') {
    e.preventDefault()
    void openFile()
  } else if (workbook.value && !isTextTarget(e.target) && !document.querySelector('[role=dialog],[role=alertdialog]')) {
    if (k === 'z') {
      e.preventDefault()
      if (e.shiftKey) book.redo()
      else book.undo()
    } else if (k === 'y') {
      e.preventDefault()
      book.redo()
    }
  }
}

function onBeforeUnload(e: BeforeUnloadEvent) {
  if (dirty.value) e.preventDefault()
}

onMounted(() => {
  window.addEventListener('keydown', onGlobalKeydown)
  window.addEventListener('beforeunload', onBeforeUnload)
})
onBeforeUnmount(() => {
  window.removeEventListener('keydown', onGlobalKeydown)
  window.removeEventListener('beforeunload', onBeforeUnload)
})
</script>

<template>
  <TooltipProvider :delay-duration="300">
    <div class="relative isolate min-h-svh bg-muted/30" @dragover="onDragOver" @drop="onDrop">
      <!-- top bar -->
      <header class="sticky top-0 z-30 border-b bg-background/70 backdrop-blur-xl">
        <div class="mx-auto flex h-14 max-w-[1440px] items-center gap-3 px-4 sm:px-6">
          <div class="flex items-center gap-2.5">
            <span
              class="grid size-8 place-items-center rounded-lg bg-linear-to-br from-emerald-500 to-teal-600 text-white shadow-sm ring-1 shadow-emerald-600/20 ring-white/20"
            >
              <SheetIcon class="size-4.5" />
            </span>
            <span class="hidden font-semibold tracking-tight sm:inline">{{ t('app.name') }}</span>
            <a
              href="https://github.com/dayth36386/excel-editor/releases"
              target="_blank"
              rel="noopener"
              class="hidden rounded-md border px-1.5 py-0.5 font-mono text-[10px] leading-none text-muted-foreground transition-colors hover:text-foreground md:inline"
              :title="t('app.releases')"
            >
              v{{ version }}
            </a>
          </div>

          <template v-if="workbook">
            <Separator orientation="vertical" class="mx-1 h-5!" />
            <div class="flex min-w-0 items-center gap-2">
              <FileSpreadsheetIcon class="size-4 shrink-0 text-emerald-600" />
              <span class="max-w-56 truncate text-sm font-medium" :title="fileName">{{ fileName }}</span>
              <Badge
                v-if="dirty"
                variant="outline"
                class="border-amber-300 bg-amber-50 text-amber-700 dark:border-amber-400/30 dark:bg-amber-400/10 dark:text-amber-300"
              >
                <span class="size-1.5 animate-pulse rounded-full bg-amber-500" />{{ t('header.unsaved') }}
              </Badge>
              <Badge
                v-else
                variant="outline"
                class="border-emerald-200 bg-emerald-50 text-emerald-700 dark:border-emerald-400/30 dark:bg-emerald-400/10 dark:text-emerald-300"
              >
                {{ t('header.saved') }}
              </Badge>
            </div>
          </template>

          <div class="ml-auto flex items-center gap-1.5">
            <template v-if="workbook">
              <Tooltip>
                <TooltipTrigger as-child>
                  <Button
                    variant="ghost"
                    size="icon"
                    :disabled="!canUndo"
                    :aria-label="t('header.undo')"
                    @click="book.undo()"
                  >
                    <Undo2Icon />
                  </Button>
                </TooltipTrigger>
                <TooltipContent>{{ t('header.undo') }} ({{ mod }}Z)</TooltipContent>
              </Tooltip>
              <Tooltip>
                <TooltipTrigger as-child>
                  <Button
                    variant="ghost"
                    size="icon"
                    :disabled="!canRedo"
                    :aria-label="t('header.redo')"
                    @click="book.redo()"
                  >
                    <Redo2Icon />
                  </Button>
                </TooltipTrigger>
                <TooltipContent>{{ t('header.redo') }} ({{ mod }}⇧Z)</TooltipContent>
              </Tooltip>
              <Separator orientation="vertical" class="mx-1 h-5!" />
            </template>

            <!-- language -->
            <div
              class="flex items-center rounded-lg border bg-muted/60 p-0.5"
              role="radiogroup"
              :aria-label="t('header.language')"
            >
              <button
                v-for="l in LOCALES"
                :key="l.code"
                type="button"
                role="radio"
                :aria-checked="locale === l.code"
                :title="l.label"
                :class="
                  cn(
                    'h-7 rounded-md px-2 text-xs font-semibold transition-all',
                    locale === l.code
                      ? 'bg-background text-foreground shadow-sm'
                      : 'text-muted-foreground hover:text-foreground',
                  )
                "
                @click="setLocale(l.code as Locale)"
              >
                {{ l.short }}
              </button>
            </div>

            <!-- theme -->
            <DropdownMenu>
              <DropdownMenuTrigger as-child>
                <Button variant="ghost" size="icon" :aria-label="t('header.theme')">
                  <component :is="themeIcon" />
                </Button>
              </DropdownMenuTrigger>
              <DropdownMenuContent align="end" class="w-40">
                <DropdownMenuRadioGroup v-model="colorMode">
                  <DropdownMenuRadioItem value="light"><SunIcon />{{ t('header.light') }}</DropdownMenuRadioItem>
                  <DropdownMenuRadioItem value="dark"><MoonIcon />{{ t('header.dark') }}</DropdownMenuRadioItem>
                  <DropdownMenuRadioItem value="auto"><MonitorIcon />{{ t('header.system') }}</DropdownMenuRadioItem>
                </DropdownMenuRadioGroup>
              </DropdownMenuContent>
            </DropdownMenu>

            <Button variant="outline" class="bg-background" @click="openFile">
              <FolderOpenIcon /><span class="hidden sm:inline">{{ t('header.open') }}</span>
            </Button>

            <div v-if="workbook" class="flex">
              <Tooltip>
                <TooltipTrigger as-child>
                  <Button class="rounded-r-none" :disabled="saving" @click="save('save')">
                    <SaveIcon /><span class="hidden sm:inline">{{
                      saving ? t('header.saving') : t('header.save')
                    }}</span>
                  </Button>
                </TooltipTrigger>
                <TooltipContent>{{ saveHint }} ({{ mod }}S)</TooltipContent>
              </Tooltip>
              <DropdownMenu>
                <DropdownMenuTrigger as-child>
                  <Button
                    size="icon"
                    class="rounded-l-none border-l border-primary-foreground/20"
                    :disabled="saving"
                    :aria-label="t('header.saveOptions')"
                  >
                    <ChevronDownIcon />
                  </Button>
                </DropdownMenuTrigger>
                <DropdownMenuContent align="end" class="w-60">
                  <DropdownMenuItem @select="save('saveAs')">
                    <SaveIcon />{{ t('header.saveAs') }}
                    <DropdownMenuShortcut>{{ mod }}⇧S</DropdownMenuShortcut>
                  </DropdownMenuItem>
                  <DropdownMenuItem @select="save('download')"
                    ><DownloadIcon />{{ t('header.download') }}</DropdownMenuItem
                  >
                  <DropdownMenuSeparator />
                  <DropdownMenuItem @select="newFile"><FilePlusIcon />{{ t('header.newFile') }}</DropdownMenuItem>
                </DropdownMenuContent>
              </DropdownMenu>
            </div>
          </div>
        </div>
      </header>

      <!-- workbook -->
      <main v-if="workbook" class="mx-auto max-w-[1440px] space-y-6 px-4 py-6 sm:px-6 sm:py-8">
        <div class="flex flex-wrap items-end justify-between gap-4">
          <div class="min-w-0">
            <h1 class="truncate text-2xl font-semibold tracking-tight sm:text-3xl">{{ sheetName }}</h1>
            <p class="mt-1 text-sm text-muted-foreground">{{ t('stats.headerRow', { row: headerRow }) }}</p>
          </div>

          <div class="flex max-w-full flex-wrap items-center gap-2">
            <!-- view switch -->
            <div
              class="flex items-center rounded-lg border bg-muted/70 p-0.5"
              role="radiogroup"
              :aria-label="t('view.label')"
            >
              <Tooltip v-for="v in views" :key="v.value">
                <TooltipTrigger as-child>
                  <button
                    type="button"
                    role="radio"
                    :aria-checked="viewMode === v.value"
                    :class="
                      cn(
                        'flex h-7 items-center gap-1.5 rounded-md px-2.5 text-sm font-medium transition-all [&_svg]:size-4',
                        viewMode === v.value
                          ? 'bg-background text-foreground shadow-sm'
                          : 'text-muted-foreground hover:text-foreground',
                      )
                    "
                    @click="setView(v.value)"
                  >
                    <component :is="v.icon" :class="viewMode === v.value && 'text-emerald-600'" />{{ v.label }}
                  </button>
                </TooltipTrigger>
                <TooltipContent>{{ v.hint }}</TooltipContent>
              </Tooltip>
            </div>

            <div class="flex max-w-full items-center gap-1 overflow-x-auto">
              <Tabs :model-value="String(activeId)" @update:model-value="(v) => switchSheet(Number(v))">
                <TabsList class="bg-muted/70">
                  <TabsTrigger
                    v-for="s in sheets"
                    :key="s.id"
                    :value="String(s.id)"
                    class="px-3"
                    :title="t('sheet.renameHint')"
                    @dblclick="renameSheet(s.id, s.name)"
                  >
                    {{ s.name }}
                  </TabsTrigger>
                </TabsList>
              </Tabs>
              <Tooltip>
                <TooltipTrigger as-child>
                  <Button
                    variant="ghost"
                    size="icon"
                    :aria-label="t('sheet.add')"
                    @click="book.addSheet(defaultHeaders())"
                  >
                    <PlusIcon />
                  </Button>
                </TooltipTrigger>
                <TooltipContent>{{ t('sheet.add') }}</TooltipContent>
              </Tooltip>
            </div>
          </div>
        </div>

        <div v-if="viewMode === 'table'" class="grid grid-cols-2 gap-3 lg:grid-cols-4">
          <div
            v-for="s in stats"
            :key="s.label"
            class="flex items-center gap-3 rounded-xl border bg-card p-4 shadow-xs transition-shadow hover:shadow-sm"
          >
            <span :class="cn('grid size-10 shrink-0 place-items-center rounded-lg', s.tone)">
              <component :is="s.icon" class="size-5" />
            </span>
            <div class="min-w-0">
              <p class="truncate text-xs font-medium text-muted-foreground">{{ s.label }}</p>
              <p class="text-xl font-semibold tracking-tight tabular-nums">{{ s.value }}</p>
            </div>
          </div>
        </div>

        <DataTable v-if="viewMode === 'table'" :key="`t${activeId}`" ref="tableRef" :book="book" />
        <SheetView v-else :key="`s${activeId}`" ref="sheetRef" :book="book" />

        <div
          v-if="viewMode === 'sheet'"
          class="flex flex-wrap items-center justify-center gap-x-5 gap-y-2 pb-2 text-xs text-muted-foreground"
        >
          <span><kbd class="kbd">↑↓←→</kbd> {{ t('hints.move') }}</span>
          <span>{{ t('hints.type') }} · <kbd class="kbd">Enter</kbd> / <kbd class="kbd">F2</kbd></span>
          <span><kbd class="kbd">Alt+Enter</kbd> {{ t('hints.newline') }}</span>
          <span><kbd class="kbd">Del</kbd> {{ t('hints.clear') }}</span>
          <span
            ><kbd class="kbd">{{ mod }}C</kbd> <kbd class="kbd">{{ mod }}V</kbd> {{ t('hints.copyPaste') }}</span
          >
          <span
            ><kbd class="kbd">{{ mod }}Z</kbd> {{ t('header.undo') }}</span
          >
        </div>
        <div
          v-else
          class="flex flex-wrap items-center justify-center gap-x-5 gap-y-2 pb-4 text-xs text-muted-foreground"
        >
          <span>{{ t('hints.edit') }}</span>
          <span><kbd class="kbd">Enter</kbd> {{ t('hints.enter') }}</span>
          <span><kbd class="kbd">Tab</kbd> {{ t('hints.tab') }}</span>
          <span><kbd class="kbd">Esc</kbd> {{ t('hints.esc') }}</span>
          <span
            ><kbd class="kbd">{{ mod }}Z</kbd> {{ t('header.undo') }}</span
          >
          <i18n-t keypath="hints.formula" tag="span">
            <template #example><code class="kbd">=SUM(E2:E10)</code></template>
          </i18n-t>
        </div>
      </main>

      <!-- landing -->
      <main v-else class="relative overflow-hidden">
        <div aria-hidden="true" class="pointer-events-none absolute inset-0 -z-10">
          <div
            class="absolute top-[-14rem] left-1/2 h-[34rem] w-[70rem] -translate-x-1/2 rounded-full bg-linear-to-tr from-emerald-300/40 via-teal-200/30 to-sky-300/30 blur-3xl dark:from-emerald-500/15 dark:via-teal-500/10 dark:to-sky-500/10"
          />
          <div
            class="absolute inset-0 bg-[linear-gradient(to_right,var(--border)_1px,transparent_1px),linear-gradient(to_bottom,var(--border)_1px,transparent_1px)] [mask-image:radial-gradient(ellipse_at_top,black_20%,transparent_65%)] bg-[size:44px_44px] opacity-70"
          />
        </div>

        <section class="mx-auto flex max-w-4xl flex-col items-center px-6 pt-16 pb-20 text-center sm:pt-24">
          <Badge variant="outline" class="h-7 gap-1.5 bg-background/70 px-3 text-xs backdrop-blur">
            <SparklesIcon class="size-3.5! text-emerald-600" />{{ t('hero.badge') }}
          </Badge>
          <h1
            class="mt-6 text-4xl leading-tight font-semibold tracking-tight text-balance sm:text-6xl sm:leading-[1.1]"
          >
            {{ t('hero.title') }}
            <span
              class="block bg-linear-to-r from-emerald-600 via-teal-500 to-sky-500 bg-clip-text pb-1 text-transparent"
            >
              {{ t('hero.titleAccent') }}
            </span>
          </h1>
          <p class="mt-5 max-w-2xl text-base text-pretty text-muted-foreground sm:text-lg">{{ t('hero.subtitle') }}</p>

          <div
            class="group mt-10 w-full max-w-xl rounded-2xl border-2 border-dashed bg-background/80 p-8 shadow-sm backdrop-blur transition-colors hover:border-emerald-400/70 sm:p-10"
          >
            <span
              class="mx-auto grid size-14 place-items-center rounded-2xl bg-emerald-500/10 text-emerald-600 ring-1 ring-emerald-500/20 transition-transform group-hover:-translate-y-0.5"
            >
              <FileUpIcon class="size-6" />
            </span>
            <p class="mt-4 font-medium">{{ t('hero.drop') }}</p>
            <p class="mt-1 text-sm text-muted-foreground">{{ t('hero.dropHint') }}</p>
            <div class="mt-6 flex flex-wrap justify-center gap-2">
              <Button size="lg" class="px-4" @click="openFile"><FolderOpenIcon />{{ t('hero.open') }}</Button>
              <Button size="lg" variant="outline" class="bg-background px-4" @click="newFile">
                <FilePlusIcon />{{ t('hero.new') }}
              </Button>
            </div>
            <Button variant="link" size="sm" class="mt-3 text-muted-foreground" @click="openSample">
              {{ t('hero.sample') }} →
            </Button>
          </div>

          <div class="mt-14 grid w-full gap-4 text-left sm:grid-cols-3">
            <div
              v-for="f in features"
              :key="f.title"
              class="rounded-xl border bg-background/70 p-5 shadow-xs backdrop-blur"
            >
              <span class="grid size-9 place-items-center rounded-lg bg-emerald-500/10 text-emerald-600">
                <component :is="f.icon" class="size-4.5" />
              </span>
              <p class="mt-4 font-medium">{{ f.title }}</p>
              <p class="mt-1 text-sm leading-relaxed text-muted-foreground">{{ f.text }}</p>
            </div>
          </div>
        </section>
      </main>

      <input ref="fileInput" type="file" accept=".xlsx" hidden @change="onFileInput" />

      <Transition
        enter-active-class="transition duration-150"
        enter-from-class="opacity-0"
        leave-active-class="transition duration-100"
        leave-to-class="opacity-0"
      >
        <div
          v-if="dragging"
          class="fixed inset-0 z-50 grid place-items-center bg-emerald-600/10 backdrop-blur-sm"
          @dragleave.self="dragging = false"
        >
          <div
            class="pointer-events-none flex flex-col items-center gap-3 rounded-2xl border-2 border-dashed border-emerald-500 bg-background px-12 py-10 text-lg font-medium text-emerald-700 shadow-xl dark:text-emerald-300"
          >
            <FileUpIcon class="size-8" />{{ t('dropOverlay') }}
          </div>
        </div>
      </Transition>

      <DialogHost />
      <Toaster position="bottom-right" rich-colors :theme="colorMode === 'auto' ? 'system' : colorMode" />
    </div>
  </TooltipProvider>
</template>

<style scoped>
@reference "./style.css";

.kbd {
  @apply rounded border bg-background px-1.5 py-0.5 font-mono text-[11px] text-foreground shadow-xs;
}
</style>
