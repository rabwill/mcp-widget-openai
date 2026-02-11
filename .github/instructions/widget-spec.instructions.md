---
applyTo: "public/**,widget/**,src/mcp-server-http.ts"
---

# Widget Specification for Coding Agent

## Context

Build widgets using **@fluentui/react-components v9** installed via npm.
Reference: https://storybooks.fluentui.dev/react/

This repo renders rich UI widgets inside ChatGPT via the **OpenAI Apps SDK** when MCP tools return data. Every widget **must** use Fluent UI React Components v9 for all UI elements, theming, and layout. **All dependencies must be npm packages** — no CDN script tags.

---

## 1. Architecture Overview

```
widget/                       ← Source TSX files (one per widget)
  ├── <name>.tsx              ← React + Fluent UI widget source
  └── build-widgets.ts        ← esbuild script that produces self-contained HTML
public/                       ← Build output (self-contained HTML served by MCP server)
  └── <name>-widget.html      ← Generated — DO NOT hand-edit
src/mcp-server-http.ts        ← Loads public/*.html and serves via resources/read + REST
```

- Widget **source** lives in `widget/<name>.tsx`.
- The build script `widget/build-widgets.ts` uses **esbuild** to bundle each TSX entry into a single self-contained HTML file in `public/`.
- The MCP server (`src/mcp-server-http.ts`) loads the built HTML at startup and serves it via `resources/read` JSON-RPC and `GET /mcp/resources/widget/<name>.html`.
- Tool output reaches the widget through `window.openai.toolOutput` (the `structuredContent` field in the MCP tool response).
- Theme (light/dark) is provided by `window.openai.theme` and the `openai:set_globals` event.

---

## 2. Required npm Dependencies

Install these packages in the project (they are **not** loaded from CDNs):

```bash
# Production dependencies (used at bundle time, inlined into the widget HTML)
npm install react react-dom @fluentui/react-components @fluentui/react-icons

# Dev dependency for the widget bundler
npm install -D esbuild
```

All other third-party libraries (e.g. `leaflet`, `maplibre-gl`, chart libraries) must also be installed via npm and imported in the TSX source.

**Never add `<script src="https://...">` CDN tags.** Everything is bundled by esbuild.

---

## 3. Widget Build System

### 3a. Build script — `widget/build-widgets.ts`

This script discovers every `widget/*.tsx` file, bundles it with esbuild, and writes a self-contained HTML file to `public/<name>-widget.html`.

```typescript
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
```

### 3b. npm scripts

Add to `package.json`:

```jsonc
{
  "scripts": {
    "build:widgets": "ts-node widget/build-widgets.ts",
    "build": "npm run build:widgets && tsc",
    "start:mcp-http": "npm run build && node dist/src/mcp-server-http.js"
  }
}
```

- `npm run build:widgets` — bundles TSX → HTML in `public/`.
- `npm run build` — builds widgets first, then compiles the server TypeScript.
- `npm run start:mcp-http` — full build + start.

---

## 4. Widget Source File Conventions

Each widget is a **single TSX file** in the `widget/` directory.

### File naming

- Source: `widget/<name>.tsx` (e.g. `widget/claim-dashboard.tsx`)
- Output: `public/<name>-widget.html` (e.g. `public/claim-dashboard-widget.html`)

### Required structure of every `widget/<name>.tsx`:

```tsx
import React, { useState, useEffect } from 'react';
import { createRoot } from 'react-dom/client';
import {
  FluentProvider,
  webLightTheme,
  webDarkTheme,
  Card,
  CardHeader,
  Text,
  Badge,
  Button,
  Spinner,
  makeStyles,
  tokens,
  shorthands,
  Title3,
  Subtitle1,
  Body1,
  Caption1,
  MessageBar,
  MessageBarBody,
  // ... add only what you use
} from '@fluentui/react-components';
import {
  ArrowLeft24Regular,
  // ... add only what you use
} from '@fluentui/react-icons';

// ─── Styles ───
const useStyles = makeStyles({
  root: { padding: tokens.spacingHorizontalL },
  centered: {
    display: 'flex',
    flexDirection: 'column',
    alignItems: 'center',
    justifyContent: 'center',
    padding: '60px 20px',
  },
});

// ─── Data helpers ───
const extractData = (raw: any) => {
  if (!raw) return null;
  if (raw.success && raw.data) {
    return Array.isArray(raw.data) ? raw.data : raw.data.items || [raw.data];
  }
  if (raw.claims) return raw.claims;
  if (raw.claim) return [raw.claim];
  if (Array.isArray(raw)) return raw;
  return null;
};

// ─── App ───
const App = () => {
  const styles = useStyles();
  const [isDark, setIsDark] = useState(false);
  const [items, setItems] = useState<any[]>([]);
  const [loading, setLoading] = useState(true);

  useEffect(() => {
    if ((window as any).openai?.theme === 'dark') setIsDark(true);

    const handleSetGlobals = (e: any) => {
      const t = e.detail?.globals?.theme;
      if (t) setIsDark(t === 'dark');
      const d = e.detail?.globals?.toolOutput;
      if (d) {
        const r = extractData(d);
        if (r) { setItems(r); setLoading(false); }
      }
    };
    window.addEventListener('openai:set_globals', handleSetGlobals);

    if (!window.matchMedia) {/* noop */}
    else if (!(window as any).openai?.theme && window.matchMedia('(prefers-color-scheme: dark)').matches) {
      setIsDark(true);
    }

    const load = () => {
      if ((window as any).openai?.toolOutput) {
        const r = extractData((window as any).openai.toolOutput);
        if (r) { setItems(r); setLoading(false); return true; }
      }
      return false;
    };
    if (!load()) {
      const iv = setInterval(() => { if (load()) clearInterval(iv); }, 100);
      setTimeout(() => { clearInterval(iv); setLoading(false); }, 5000);
    }

    return () => window.removeEventListener('openai:set_globals', handleSetGlobals);
  }, []);

  return (
    <FluentProvider theme={isDark ? webDarkTheme : webLightTheme}>
      <div className={styles.root}>
        {loading ? (
          <div className={styles.centered}>
            <Spinner label="Loading..." />
          </div>
        ) : items.length === 0 ? (
          <MessageBar intent="info">
            <MessageBarBody>No data found.</MessageBarBody>
          </MessageBar>
        ) : (
          <div>{/* Render items here using Fluent UI components */}</div>
        )}
      </div>
    </FluentProvider>
  );
};

// ─── Mount ───
createRoot(document.getElementById('root')!).render(<App />);
```

### Key rules

- **Use ES module imports** — `import { ... } from '@fluentui/react-components'`.
- **Never use CDN `<script>` tags** or global UMD references like `FluentUIReactComponents`.
- **Never use `<script type="text/babel">`** — esbuild compiles JSX/TSX natively.
- **No in-browser Babel** — all transformation happens at build time.
- **Import CSS via npm** when a library ships CSS (esbuild's CSS loader inlines it).
- Third-party libs (e.g. `leaflet`) must be `npm install`ed and `import`ed.

---

## 5. Theming — FluentProvider (Required)

**Never use raw CSS custom properties for colors.** Use Fluent UI's `FluentProvider` with `webLightTheme` / `webDarkTheme` and Fluent `tokens` for all styling.

### Theme detection and switching:

```tsx
const App = () => {
  const [isDark, setIsDark] = useState(false);

  useEffect(() => {
    // Read initial theme from OpenAI Apps SDK
    if ((window as any).openai?.theme === 'dark') setIsDark(true);

    const handleSetGlobals = (event: any) => {
      const theme = event.detail?.globals?.theme;
      if (theme) setIsDark(theme === 'dark');
    };
    window.addEventListener('openai:set_globals', handleSetGlobals);

    // Fallback: check system preference
    const mq = window.matchMedia('(prefers-color-scheme: dark)');
    if (!(window as any).openai?.theme && mq.matches) setIsDark(true);

    return () => window.removeEventListener('openai:set_globals', handleSetGlobals);
  }, []);

  return (
    <FluentProvider theme={isDark ? webDarkTheme : webLightTheme}>
      {/* Widget content here */}
    </FluentProvider>
  );
};
```

### Styling with `makeStyles` and `tokens`:

```tsx
const useStyles = makeStyles({
  root: {
    maxWidth: '900px',
    margin: '0 auto',
    padding: tokens.spacingHorizontalL,
  },
  card: {
    ...shorthands.padding(tokens.spacingVerticalM, tokens.spacingHorizontalM),
    marginBottom: tokens.spacingVerticalM,
    cursor: 'pointer',
    ':hover': {
      boxShadow: tokens.shadow8,
    },
  },
  statsGrid: {
    display: 'grid',
    gridTemplateColumns: 'repeat(auto-fit, minmax(140px, 1fr))',
    gap: tokens.spacingHorizontalM,
    marginBottom: tokens.spacingVerticalL,
  },
});
```

**Never hardcode hex colors.** Always use `tokens.colorNeutralBackground1`, `tokens.colorBrandForeground1`, `tokens.colorPaletteGreenForeground1`, etc.

---

## 6. Preferred Fluent UI Components

Use these Fluent UI components instead of hand-rolled HTML/CSS:

| Purpose | Fluent UI Component | Import from |
|---|---|---|
| Layout wrapper | `FluentProvider` | `@fluentui/react-components` |
| Cards | `Card`, `CardHeader`, `CardPreview` | `@fluentui/react-components` |
| Data tables | `DataGrid`, `DataGridHeader`, `DataGridBody`, `DataGridRow`, `DataGridCell` | `@fluentui/react-components` |
| Status indicators | `Badge` (with `appearance` and `color`) | `@fluentui/react-components` |
| Buttons / actions | `Button`, `CompoundButton` | `@fluentui/react-components` |
| Loading states | `Spinner` (with `label` prop) | `@fluentui/react-components` |
| Empty states | `MessageBar` + `MessageBarBody` | `@fluentui/react-components` |
| Typography | `Title3`, `Subtitle1`, `Subtitle2`, `Body1`, `Caption1` | `@fluentui/react-components` |
| Tags / chips | `Badge` or `Tag` | `@fluentui/react-components` |
| Layout dividers | `Divider` | `@fluentui/react-components` |
| Icons | Named icons like `ArrowLeft24Regular` | `@fluentui/react-icons` |

---

## 7. Receiving Data from MCP Tools

The MCP server returns `structuredContent` in tool responses. The OpenAI Apps SDK injects it as `window.openai.toolOutput`.

### Data loading pattern (required):

```tsx
const [data, setData] = useState<any[] | null>(null);
const [loading, setLoading] = useState(true);

useEffect(() => {
  const loadData = () => {
    if ((window as any).openai?.toolOutput) {
      const result = extractData((window as any).openai.toolOutput);
      if (result) { setData(result); setLoading(false); return true; }
    }
    return false;
  };

  const handleSetGlobals = (event: any) => {
    const toolOutput = event.detail?.globals?.toolOutput;
    if (toolOutput) {
      const result = extractData(toolOutput);
      if (result) { setData(result); setLoading(false); }
    }
  };
  window.addEventListener('openai:set_globals', handleSetGlobals);

  if (!loadData()) {
    const interval = setInterval(() => { if (loadData()) clearInterval(interval); }, 100);
    setTimeout(() => { clearInterval(interval); setLoading(false); }, 5000);
  }

  return () => window.removeEventListener('openai:set_globals', handleSetGlobals);
}, []);
```

### Standard server response shape:

```json
{ "success": true, "data": [ ... ], "message": "...", "timestamp": "..." }
```

### Required `extractData` helper:

```tsx
const extractData = (raw: any) => {
  if (!raw) return null;
  if (raw.success && raw.data) {
    return Array.isArray(raw.data) ? raw.data : raw.data.items || [raw.data];
  }
  if (raw.claims) return raw.claims;
  if (raw.claim) return [raw.claim];
  if (Array.isArray(raw)) return raw;
  return null;
};
```

---

## 8. Connecting a Widget to an MCP Tool

In `src/mcp-server-http.ts`, add `_meta` to the tool definition in the `TOOLS` array:

```typescript
{
  name: 'get_claims',
  description: '...',
  inputSchema: { ... },
  _meta: {
    "openai/outputTemplate": "ui://widget/claim-dashboard.html",
    "openai/toolInvocation/invoking": "Fetching claims...",
    "openai/toolInvocation/invoked": "Claims loaded",
  },
},
```

- `openai/outputTemplate` — must match the resource URI in `resources/list`.
- `openai/toolInvocation/invoking` — loading text shown in ChatGPT while the tool runs.
- `openai/toolInvocation/invoked` — text shown after the tool completes.

---

## 9. Registering a Widget Resource

Update **three** places in `src/mcp-server-http.ts`:

### a) Load the built HTML at the top of the file:

```typescript
import { readFileSync } from 'fs';
import { resolve } from 'path';

const widgetHtml = readFileSync(
  resolve(process.cwd(), 'public/<name>-widget.html'),
  'utf8'
);
```

### b) REST endpoint (Express route):

```typescript
app.get('/mcp/resources/widget/<name>.html', (req, res) => {
  res.setHeader('Content-Type', 'text/html+skybridge');
  res.setHeader('X-Widget-Prefers-Border', 'true');
  res.send(widgetHtml);
});
```

### c) JSON-RPC `resources/list` response:

```typescript
{
  uri: 'ui://widget/<name>.html',
  name: '<name>-widget',
  mimeType: 'text/html+skybridge',
  description: '...',
  _meta: {
    'openai/widgetPrefersBorder': true,
    'openai/widgetDescription': '...',
  },
}
```

### d) JSON-RPC `resources/read` handler:

Add a case for the URI returning `{ contents: [{ uri, mimeType: 'text/html+skybridge', text: widgetHtml }] }`.

---

## 10. UI / UX Guidelines

- **Max width** `900px`, centered via `margin: 0 auto`.
- **All colors** via Fluent `tokens` — never hardcode hex values.
- **Cards** use Fluent `Card` component — no custom `.card` CSS classes.
- **Status badges** use `<Badge color="success|warning|danger|informative">` — no custom badge CSS.
- **Loading** state uses `<Spinner label="Loading..." />`.
- **Empty** state uses `<MessageBar intent="info"><MessageBarBody>No data.</MessageBarBody></MessageBar>`.
- **Responsive**: use CSS grid with `repeat(auto-fit, minmax(...))` and a `@media (max-width: 600px)` breakpoint via `makeStyles`.
- Icons from `@fluentui/react-icons` only — no emoji or inline SVG for action icons.

---

## 11. Checklist for Adding a New Widget

1. Create `widget/<name>.tsx` starting from the skeleton in Section 4.
2. `npm install` any third-party packages the widget needs.
3. Use only Fluent UI components and `tokens` for all UI — no CDN scripts.
4. Run `npm run build:widgets` to produce `public/<name>-widget.html`.
5. Load the HTML file at the top of `src/mcp-server-http.ts` with `readFileSync`.
6. Register the resource in the REST endpoint, `resources/list`, and `resources/read` JSON-RPC handlers.
7. Add `_meta` with `openai/outputTemplate` to each tool that should render this widget.
8. Ensure `structuredContent` is returned by `executeTool` (this is the default pattern).
9. Run `npm run start:mcp-http` and test with `npm run inspector`.

---

## 12. Forbidden Patterns

These patterns must **never** appear in widget source:

| Forbidden | Use instead |
|---|---|
| `<script src="https://unpkg.com/...">` | `npm install <pkg>` + `import` |
| `<script src="https://cdn.jsdelivr.net/...">` | `npm install <pkg>` + `import` |
| `<script type="text/babel">` | esbuild compiles TSX at build time |
| `FluentUIReactComponents` global | `import { ... } from '@fluentui/react-components'` |
| `FluentUIReactIcons` global | `import { ... } from '@fluentui/react-icons'` |
| Hardcoded hex/rgb colors | Fluent `tokens.color*` |
| Hand-written CSS class `.card`, `.badge` | Fluent `Card`, `Badge` components |
| In-browser Babel transform | esbuild JSX compilation |
