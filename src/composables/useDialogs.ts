import { reactive } from 'vue'

export interface ConfirmOptions {
  title: string
  description?: string
  confirmText?: string
  destructive?: boolean
}

export interface PromptOptions {
  title: string
  description?: string
  label?: string
  value?: string
  placeholder?: string
  confirmText?: string
}

/** Shared state for the one confirm dialog and one prompt dialog in DialogHost. */
export const dialogs = reactive({
  confirm: {
    open: false,
    options: { title: '' } as ConfirmOptions,
    resolve: null as ((ok: boolean) => void) | null,
  },
  prompt: {
    open: false,
    options: { title: '' } as PromptOptions,
    value: '',
    resolve: null as ((value: string | null) => void) | null,
  },
})

/** Promise-based replacement for window.confirm, rendered with shadcn AlertDialog. */
export function confirmDialog(options: ConfirmOptions): Promise<boolean> {
  dialogs.confirm.resolve?.(false)
  return new Promise((resolve) => {
    dialogs.confirm.options = options
    dialogs.confirm.resolve = resolve
    dialogs.confirm.open = true
  })
}

/** Promise-based replacement for window.prompt, rendered with shadcn Dialog. Resolves null on cancel. */
export function promptDialog(options: PromptOptions): Promise<string | null> {
  dialogs.prompt.resolve?.(null)
  return new Promise((resolve) => {
    dialogs.prompt.options = options
    dialogs.prompt.value = options.value ?? ''
    dialogs.prompt.resolve = resolve
    dialogs.prompt.open = true
  })
}
