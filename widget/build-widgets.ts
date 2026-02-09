import { buildSync } from 'esbuild';
import { readdirSync, writeFileSync, mkdirSync } from 'fs';
import { resolve, basename } from 'path';

const widgetDir = resolve(__dirname);
const outDir = resolve(__dirname, '..', 'public');
mkdirSync(outDir, { recursive: true });

const entries = readdirSync(widgetDir).filter(f => f.endsWith('.tsx'));

for (const entry of entries) {
  const name = basename(entry, '.tsx');
  const result = buildSync({
    entryPoints: [resolve(widgetDir, entry)],
    bundle: true,
    minify: true,
    format: 'iife',
    target: ['es2020'],
    jsx: 'automatic',
    outdir: outDir,
    loader: {
      '.tsx': 'tsx',
      '.ts': 'ts',
      '.css': 'css',
      '.png': 'dataurl',
      '.svg': 'dataurl',
      '.jpg': 'dataurl',
      '.gif': 'dataurl',
    },
    write: false,
    define: {
      'process.env.NODE_ENV': '"production"',
    },
  });

  // esbuild produces separate JS and CSS output files when CSS is imported
  let js = '';
  let css = '';
  for (const file of result.outputFiles) {
    if (file.path.endsWith('.css')) {
      css = file.text;
    } else {
      js = file.text;
    }
  }

  const html = `<!DOCTYPE html>
<html lang="en">
  <head>
    <meta charset="utf-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1.0" />
    <title>${name}</title>
    <style>
      * { margin: 0; padding: 0; box-sizing: border-box; }
      html, body { width: 100%; min-height: 100%; font-family: system-ui, sans-serif; }
      #root { max-width: 900px; margin: 0 auto; width: 100%; }
    </style>${css ? `\n    <style>${css}</style>` : ''}
  </head>
  <body>
    <div id="root"></div>
    <script>${js}</script>
  </body>
</html>`;

  writeFileSync(resolve(outDir, `${name}-widget.html`), html);
  console.log(`Built public/${name}-widget.html`);
}
