/**
 * MCP Apps Bridge hooks — JSON-RPC 2.0 over postMessage.
 *
 * These hooks replace the legacy `window.openai` API with the recommended
 * MCP Apps bridge transport described at:
 *   https://developers.openai.com/apps-sdk/build/chatgpt-ui
 *
 * The bridge still falls back to `window.openai` when the host hasn't sent
 * a postMessage notification yet (backwards-compatible with older hosts).
 */

import { useState, useEffect, useCallback, useRef } from 'react';

// ── Types ────────────────────────────────────────────────────────────────────

/** Shape delivered by `ui/notifications/tool-result`. */
interface ToolResultNotification {
  content?: unknown[];
  structuredContent?: unknown;
}

/** Shape delivered by `ui/notifications/tool-input`. */
interface ToolInputNotification {
  name?: string;
  arguments?: Record<string, unknown>;
}

// ── Helpers ──────────────────────────────────────────────────────────────────

/** Type-guard for JSON-RPC 2.0 notifications arriving via postMessage. */
function isJsonRpcNotification(
  data: unknown
): data is { jsonrpc: '2.0'; method: string; params?: Record<string, unknown> } {
  if (typeof data !== 'object' || data === null) return false;
  const msg = data as Record<string, unknown>;
  return msg.jsonrpc === '2.0' && typeof msg.method === 'string';
}

// ── useToolResult ────────────────────────────────────────────────────────────

/**
 * Subscribe to `ui/notifications/tool-result` from the MCP Apps bridge.
 *
 * Returns `structuredContent` from the latest tool-result notification.
 * Falls back to `window.openai.toolOutput` for legacy hosts.
 */
export function useToolResult(): unknown | null {
  const [result, setResult] = useState<unknown | null>(null);

  useEffect(() => {
    // ── 1. Listen for MCP Apps bridge notifications (recommended path) ──
    const onMessage = (event: MessageEvent) => {
      if (event.source !== window.parent) return;
      const data = event.data;
      if (!isJsonRpcNotification(data)) return;

      if (data.method === 'ui/notifications/tool-result') {
        const params = data.params as ToolResultNotification | undefined;
        if (params?.structuredContent != null) {
          setResult(params.structuredContent);
        }
      }
    };

    window.addEventListener('message', onMessage, { passive: true });

    // ── 2. Legacy fallback: `window.openai.toolOutput` ──
    const tryLegacy = (): boolean => {
      const openai = (window as any).openai;
      if (openai?.toolOutput) {
        setResult(openai.toolOutput);
        return true;
      }
      return false;
    };

    // Also listen for the legacy `openai:set_globals` event
    const onSetGlobals = (e: any) => {
      const toolOutput = e.detail?.globals?.toolOutput;
      if (toolOutput != null) {
        setResult(toolOutput);
      }
    };
    window.addEventListener('openai:set_globals', onSetGlobals);

    // Poll briefly for legacy host injection
    if (!tryLegacy()) {
      const iv = setInterval(() => {
        if (tryLegacy()) clearInterval(iv);
      }, 100);
      setTimeout(() => clearInterval(iv), 5000);
    }

    return () => {
      window.removeEventListener('message', onMessage);
      window.removeEventListener('openai:set_globals', onSetGlobals);
    };
  }, []);

  return result;
}

// ── useHostTheme ─────────────────────────────────────────────────────────────

/**
 * Detect the host theme (light / dark).
 *
 * Reads from:
 *   1. `document.documentElement.lang` locale attribute set by the host
 *   2. `ui/notifications/tool-result` or other bridge messages carrying theme
 *   3. Legacy `window.openai.theme` / `openai:set_globals`
 *   4. `prefers-color-scheme` media query as final fallback
 */
export function useHostTheme(): boolean {
  const [isDark, setIsDark] = useState(false);

  useEffect(() => {
    // ── 1. Bridge: theme can arrive as `ui/notifications/set-theme` or via global ──
    const onMessage = (event: MessageEvent) => {
      if (event.source !== window.parent) return;
      const data = event.data;
      if (!isJsonRpcNotification(data)) return;

      // Some hosts send theme information via a dedicated notification
      if (data.method === 'ui/notifications/set-theme') {
        const theme = (data.params as any)?.theme;
        if (theme) setIsDark(theme === 'dark');
      }
    };
    window.addEventListener('message', onMessage, { passive: true });

    // ── 2. Legacy: `window.openai.theme` ──
    const openai = (window as any).openai;
    if (openai?.theme === 'dark') {
      setIsDark(true);
    }

    // ── 3. Legacy event ──
    const onSetGlobals = (e: any) => {
      const theme = e.detail?.globals?.theme;
      if (theme) setIsDark(theme === 'dark');
    };
    window.addEventListener('openai:set_globals', onSetGlobals);

    // ── 4. System preference fallback ──
    if (!openai?.theme && window.matchMedia?.('(prefers-color-scheme: dark)').matches) {
      setIsDark(true);
    }

    return () => {
      window.removeEventListener('message', onMessage);
      window.removeEventListener('openai:set_globals', onSetGlobals);
    };
  }, []);

  return isDark;
}

// ── useSendMessage ───────────────────────────────────────────────────────────

/**
 * Send a follow-up user message to the host via `ui/message`.
 *
 * @example
 *   const send = useSendMessage();
 *   send("Show me the inspections for this claim.");
 */
export function useSendMessage() {
  return useCallback((text: string) => {
    window.parent.postMessage(
      {
        jsonrpc: '2.0',
        method: 'ui/message',
        params: {
          role: 'user',
          content: [{ type: 'text', text }],
        },
      },
      '*'
    );
  }, []);
}

// ── useCallTool ──────────────────────────────────────────────────────────────

let rpcIdCounter = 1;

/**
 * Call an MCP tool from the widget via the MCP Apps bridge (`tools/call`).
 *
 * Returns a function that sends the JSON-RPC request and resolves with the
 * response. Falls back to a direct HTTP call to the server for legacy hosts.
 */
export function useCallTool() {
  const pendingRef = useRef<Map<number, { resolve: (v: any) => void; reject: (e: any) => void }>>(
    new Map()
  );

  useEffect(() => {
    const onMessage = (event: MessageEvent) => {
      if (event.source !== window.parent) return;
      const data = event.data;
      if (typeof data !== 'object' || data === null) return;
      if (data.jsonrpc !== '2.0' || data.id == null) return;

      const pending = pendingRef.current.get(data.id);
      if (pending) {
        pendingRef.current.delete(data.id);
        if (data.error) {
          pending.reject(new Error(data.error.message || 'RPC error'));
        } else {
          pending.resolve(data.result);
        }
      }
    };

    window.addEventListener('message', onMessage, { passive: true });
    return () => window.removeEventListener('message', onMessage);
  }, []);

  return useCallback(
    (toolName: string, args: Record<string, unknown> = {}): Promise<any> => {
      return new Promise((resolve, reject) => {
        const id = rpcIdCounter++;
        pendingRef.current.set(id, { resolve, reject });

        window.parent.postMessage(
          {
            jsonrpc: '2.0',
            id,
            method: 'tools/call',
            params: { name: toolName, arguments: args },
          },
          '*'
        );

        // Timeout — if the bridge doesn't respond within 30s, fall back to HTTP
        setTimeout(() => {
          if (pendingRef.current.has(id)) {
            pendingRef.current.delete(id);
            // Fallback: direct HTTP call
            const baseUrl = (window as any).__MCP_SERVER_URL__ || window.location.origin;
            fetch(`${baseUrl}/mcp/tools/call`, {
              method: 'POST',
              headers: { 'Content-Type': 'application/json' },
              body: JSON.stringify({ name: toolName, arguments: args }),
            })
              .then((r) => r.json())
              .then(resolve)
              .catch(reject);
          }
        }, 30_000);
      });
    },
    []
  );
}

// ── useUpdateModelContext ────────────────────────────────────────────────────

let contextRpcIdCounter = 1000;

/**
 * Notify the host that widget-visible state has changed in a way the model
 * should know about, via `ui/update-model-context`.
 */
export function useUpdateModelContext() {
  const pendingRef = useRef<Map<number, { resolve: (v: any) => void; reject: (e: any) => void }>>(
    new Map()
  );

  useEffect(() => {
    const onMessage = (event: MessageEvent) => {
      if (event.source !== window.parent) return;
      const data = event.data;
      if (typeof data !== 'object' || data === null) return;
      if (data.jsonrpc !== '2.0' || data.id == null) return;

      const pending = pendingRef.current.get(data.id);
      if (pending) {
        pendingRef.current.delete(data.id);
        if (data.error) {
          pending.reject(new Error(data.error.message || 'RPC error'));
        } else {
          pending.resolve(data.result);
        }
      }
    };

    window.addEventListener('message', onMessage, { passive: true });
    return () => window.removeEventListener('message', onMessage);
  }, []);

  return useCallback((text: string) => {
    return new Promise<any>((resolve, reject) => {
      const id = contextRpcIdCounter++;
      pendingRef.current.set(id, { resolve, reject });

      window.parent.postMessage(
        {
          jsonrpc: '2.0',
          id,
          method: 'ui/update-model-context',
          params: {
            content: [{ type: 'text', text }],
          },
        },
        '*'
      );

      // Don't block forever if host doesn't respond
      setTimeout(() => {
        if (pendingRef.current.has(id)) {
          pendingRef.current.delete(id);
          resolve(undefined);
        }
      }, 10_000);
    });
  }, []);
}
