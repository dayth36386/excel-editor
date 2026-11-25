<script setup lang="ts">
import { nextTick, ref, watch } from 'vue'
import {
  AlertDialog,
  AlertDialogCancel,
  AlertDialogContent,
  AlertDialogDescription,
  AlertDialogFooter,
  AlertDialogHeader,
  AlertDialogTitle,
} from '@/components/ui/alert-dialog'
import { Button } from '@/components/ui/button'
import {
  Dialog,
  DialogContent,
  DialogDescription,
  DialogFooter,
  DialogHeader,
  DialogTitle,
} from '@/components/ui/dialog'
import { Input } from '@/components/ui/input'
import { Label } from '@/components/ui/label'
import { dialogs } from '@/composables/useDialogs'
import { useI18n } from 'vue-i18n'

const { t } = useI18n()

function finishConfirm(ok: boolean) {
  dialogs.confirm.resolve?.(ok)
  dialogs.confirm.resolve = null
  dialogs.confirm.open = false
}

function finishPrompt(value: string | null) {
  dialogs.prompt.resolve?.(value)
  dialogs.prompt.resolve = null
  dialogs.prompt.open = false
}

const promptInput = ref<HTMLElement | null>(null)
watch(
  () => dialogs.prompt.open,
  (open) => {
    if (!open) return
    nextTick(() => {
      const el = (promptInput.value as unknown as { $el?: HTMLInputElement })?.$el
      el?.focus()
      el?.select()
    })
  },
)
</script>

<template>
  <AlertDialog :open="dialogs.confirm.open" @update:open="(open) => !open && finishConfirm(false)">
    <AlertDialogContent>
      <AlertDialogHeader>
        <AlertDialogTitle>{{ dialogs.confirm.options.title }}</AlertDialogTitle>
        <AlertDialogDescription v-if="dialogs.confirm.options.description">
          {{ dialogs.confirm.options.description }}
        </AlertDialogDescription>
      </AlertDialogHeader>
      <AlertDialogFooter>
        <AlertDialogCancel @click="finishConfirm(false)">{{ t('dialogs.cancel') }}</AlertDialogCancel>
        <!-- A plain Button: AlertDialogAction closes the dialog before our click
             handler runs, which would resolve the promise as "cancelled". -->
        <Button :variant="dialogs.confirm.options.destructive ? 'destructive' : 'default'" @click="finishConfirm(true)">
          {{ dialogs.confirm.options.confirmText ?? t('dialogs.confirm') }}
        </Button>
      </AlertDialogFooter>
    </AlertDialogContent>
  </AlertDialog>

  <Dialog :open="dialogs.prompt.open" @update:open="(open) => !open && finishPrompt(null)">
    <DialogContent class="sm:max-w-md">
      <form class="grid gap-4" @submit.prevent="finishPrompt(dialogs.prompt.value.trim())">
        <DialogHeader>
          <DialogTitle>{{ dialogs.prompt.options.title }}</DialogTitle>
          <DialogDescription v-if="dialogs.prompt.options.description">
            {{ dialogs.prompt.options.description }}
          </DialogDescription>
        </DialogHeader>
        <div class="grid gap-2">
          <Label v-if="dialogs.prompt.options.label" for="prompt-input">{{ dialogs.prompt.options.label }}</Label>
          <Input
            id="prompt-input"
            ref="promptInput"
            v-model="dialogs.prompt.value"
            :placeholder="dialogs.prompt.options.placeholder"
            autocomplete="off"
          />
        </div>
        <DialogFooter>
          <Button type="button" variant="outline" @click="finishPrompt(null)">{{ t('dialogs.cancel') }}</Button>
          <Button type="submit" :disabled="!dialogs.prompt.value.trim()">
            {{ dialogs.prompt.options.confirmText ?? t('dialogs.ok') }}
          </Button>
        </DialogFooter>
      </form>
    </DialogContent>
  </Dialog>
</template>
