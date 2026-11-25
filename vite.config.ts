/// <reference types="vitest/config" />
import { execSync } from 'node:child_process'
import { readFileSync } from 'node:fs'
import { fileURLToPath, URL } from 'node:url'
import { defineConfig } from 'vite'
import vue from '@vitejs/plugin-vue'
import tailwindcss from '@tailwindcss/vite'

/**
 * Version shown in the app: the latest release tag (vX.Y.Z, created by semantic-release),
 * or package.json's version when there is no tag or no git (e.g. a fresh local clone).
 */
function appVersion(): string {
  try {
    const tag = execSync('git describe --tags --abbrev=0 --match "v*"', { stdio: ['ignore', 'pipe', 'ignore'] })
    return tag.toString().trim().replace(/^v/, '')
  } catch {
    const pkg = JSON.parse(readFileSync(new URL('./package.json', import.meta.url), 'utf8')) as { version: string }
    return pkg.version
  }
}
const version = appVersion()

export default defineConfig({
  plugins: [vue(), tailwindcss()],
  // vue-i18n compile-time flags (Composition API only; silences the bundler warning).
  define: {
    __VUE_I18N_FULL_INSTALL__: true,
    __VUE_I18N_LEGACY_API__: false,
    __INTLIFY_PROD_DEVTOOLS__: false,
    __APP_VERSION__: JSON.stringify(version),
  },
  resolve: {
    alias: { '@': fileURLToPath(new URL('./src', import.meta.url)) },
  },
  // ExcelJS alone is ~1 MB; it is split out and loaded when a file is opened.
  build: { chunkSizeWarningLimit: 1200 },
  test: {
    environment: 'node',
  },
})
