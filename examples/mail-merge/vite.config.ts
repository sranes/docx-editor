import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';
import tailwindcss from 'tailwindcss';
import autoprefixer from 'autoprefixer';
import { stripUnexpandedTailwind } from '../shared/strip-unexpanded-tailwind';
import path from 'path';

const monorepoRoot = path.resolve(__dirname, '../..');
const core = (entry: string) => path.join(monorepoRoot, 'packages/core/src', entry);

// Resolve the two packages to SOURCE, so the sample runs the working tree, not a stale
// `dist/`. A consumer app imports the published packages and needs no aliases.
export default defineConfig({
  plugins: [react()],
  root: __dirname,
  resolve: {
    alias: [
      {
        find: '@docx-editor.dev/react',
        replacement: path.join(monorepoRoot, 'packages/react/src/index.ts'),
      },
      { find: '@docx-editor.dev/core/editor', replacement: core('editor/index.ts') },
      {
        find: /^@docx-editor\.dev\/core\/(automation|binding|layout|output|store|export|collaboration)$/,
        replacement: core('$1/index.ts'),
      },
      { find: /^@docx-editor\.dev\/core\/contracts\/(.+)$/, replacement: core('contracts/$1.ts') },
      { find: /^@docx-editor\.dev\/core$/, replacement: core('index.ts') },
      {
        find: '@docx-editor.dev/i18n',
        replacement: path.join(monorepoRoot, 'packages/i18n/src/index.ts'),
      },
    ],
  },
  css: {
    postcss: {
      plugins: [
        tailwindcss({ config: path.join(monorepoRoot, 'tailwind.config.js') }),
        autoprefixer(),
        stripUnexpandedTailwind,
      ],
    },
  },
  server: { port: 5181, open: false },
});
