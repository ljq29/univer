import { resolve } from 'node:path';
import { defineConfig } from 'vite';
import cssInjectedByJsPlugin from 'vite-plugin-css-injected-by-js';

export default defineConfig(({ mode }) => {
  const scriptFile = process.env.fileName ?? 'index.ts';
  const format = process.env.format === 'esm' ? 'esm' : 'cjs';

  return {
    base: mode === 'development' ? '/' : '/plugins/univer/',
    publicDir: 'src/univer',
    resolve: {
      alias: {
        '~': resolve(import.meta.dirname, 'src'),
        '@': resolve(import.meta.dirname, 'src'),
        siyuan: 'test',
      },
    },
    define: {
      'process.env.NODE_ENV': JSON.stringify(mode),
      process: { env: { DEV_MODE: mode } },
    },
    plugins: [
      cssInjectedByJsPlugin(),
      {
        name: 'siyuan shim',
        transform(code) {
          if (/import \{(.*)\} from "siyuan"/.test(code)) {
            return {
              code: code.replace(
                /import \{(.*)\} from "siyuan"/,
                'const {$1} = require("siyuan")',
              ),
              map: null,
            };
          }
        },
      },
    ],
    build: {
      lib: {
        entry: `src/univer/${scriptFile}`,
        name: 'univer',
      },
      outDir: 'dist/univer',
      emptyOutDir: process.env.emptyOutDir !== '0',
      sourcemap: false,
      minify: true,
      rollupOptions: {
        output: {
          entryFileNames:
            (format === 'esm' ? 'es/' : '') + scriptFile.replace(/\.tsx?$/, '.js'),
          format,
          assetFileNames: 'asset/[name]-[hash][extname]',
          inlineDynamicImports: true,
        },
        external: ['siyuan', 'process'],
      },
    },
  };
});
