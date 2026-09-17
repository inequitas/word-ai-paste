/// <reference types="vitest/config" />
import { defineConfig } from 'vite';
import { resolve } from 'node:path';
import { readFileSync } from 'node:fs';
import { execSync } from 'node:child_process';

const pkg = JSON.parse(readFileSync(resolve(__dirname, 'package.json'), 'utf-8')) as { version: string };

function shortGitHash(): string {
  try {
    return execSync('git rev-parse --short HEAD', { cwd: __dirname }).toString().trim();
  } catch {
    return 'unknown';
  }
}

export default defineConfig({
  base: '/word-ai-paste/',
  root: '.',
  // So Kevin can confirm he's actually running a new build in the task
  // pane's footer — Word's webview can cache taskpane.html for a while on
  // GitHub Pages, so "it looks the same" is otherwise ambiguous.
  define: {
    __APP_VERSION__: JSON.stringify(pkg.version),
    __GIT_HASH__: JSON.stringify(shortGitHash())
  },
  build: {
    outDir: 'dist',
    emptyOutDir: true,
    rollupOptions: {
      input: {
        main: resolve(__dirname, 'index.html'),
        taskpane: resolve(__dirname, 'taskpane.html')
      }
    }
  },
  server: {
    port: 3000
  },
  test: {
    environment: 'happy-dom',
    include: ['tests/**/*.test.ts'],
    globals: false
  }
});
