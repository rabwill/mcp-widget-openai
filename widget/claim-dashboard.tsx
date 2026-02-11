import React, { useState, useEffect } from 'react';
import { createRoot } from 'react-dom/client';
import {
  FluentProvider,
  webLightTheme,
  webDarkTheme,
  Card,
  Badge,
  Button,
  Spinner,
  Input,
  Textarea,
  Select,
  Checkbox,
  makeStyles,
  tokens,
  shorthands,
  Title3,
  Subtitle1,
  Subtitle2,
  Body1,
  Body2,
  Caption1,
  MessageBar,
  MessageBarBody,
  Tooltip,
} from '@fluentui/react-components';
import {
  ArrowLeft24Regular,
  Edit24Regular,
  Save24Regular,
  Dismiss24Regular,
  Location24Regular,
  Checkmark16Regular,
  DismissCircle16Regular,
  Mail24Regular,
  ClipboardPaste24Regular,
  CheckboxChecked24Regular,
  CheckboxUnchecked24Regular,
  DismissSquare24Regular,
} from '@fluentui/react-icons';
import { useToolResult, useHostTheme, useCallTool } from './hooks/useMcpBridge';

// ─── Status Options ───
const STATUS_OPTIONS = [
  'Open - Claim is under investigation',
  'Pending Documentation - Awaiting additional information',
  'Under Investigation - Claim is being reviewed by adjusters',
  'In Repair - Repairs in progress',
  'Approved - Claim approved for payment',
  'Closed - Claim settled and closed',
];

// ─── Styles ───
const useStyles = makeStyles({
  root: { padding: tokens.spacingHorizontalM },
  centered: {
    display: 'flex',
    flexDirection: 'column',
    alignItems: 'center',
    justifyContent: 'center',
    padding: '40px 16px',
  },
  header: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'space-between',
    gap: tokens.spacingHorizontalS,
    marginBottom: tokens.spacingVerticalM,
    paddingBottom: tokens.spacingVerticalS,
    borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
  },
  headerLeft: {
    display: 'flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalS,
  },
  logo: {
    width: '36px',
    height: '36px',
    background: `linear-gradient(135deg, ${tokens.colorBrandBackground}, ${tokens.colorBrandBackgroundPressed})`,
    borderRadius: tokens.borderRadiusMedium,
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    color: tokens.colorNeutralForegroundOnBrand,
    fontWeight: '700',
    fontSize: '16px',
    boxShadow: tokens.shadow2,
    flexShrink: 0,
  },
  statsGrid: {
    display: 'grid',
    gridTemplateColumns: 'repeat(auto-fit, minmax(100px, 1fr))',
    gap: tokens.spacingHorizontalS,
    marginBottom: tokens.spacingVerticalM,
  },
  statCard: {
    ...shorthands.padding(tokens.spacingVerticalS, tokens.spacingHorizontalS),
    textAlign: 'center' as const,
  },
  statValue: { fontWeight: '700', fontSize: '1.25rem' },
  statLabel: {
    color: tokens.colorNeutralForeground3,
    fontSize: '0.7rem',
    textTransform: 'uppercase' as const,
    letterSpacing: '0.5px',
    marginTop: '2px',
  },
  claimsList: {
    display: 'flex',
    flexDirection: 'column' as const,
    gap: tokens.spacingVerticalS,
  },
  claimCard: {
    ...shorthands.padding(tokens.spacingVerticalS, tokens.spacingHorizontalM),
    cursor: 'pointer',
    ':hover': { boxShadow: tokens.shadow8 },
  },
  claimCardHeader: {
    display: 'flex',
    alignItems: 'flex-start',
    justifyContent: 'space-between',
    gap: tokens.spacingHorizontalS,
    marginBottom: '4px',
  },
  detailGrid: {
    display: 'grid',
    gridTemplateColumns: 'repeat(3, 1fr)',
    gap: tokens.spacingHorizontalS,
    marginTop: '4px',
  },
  damageTags: {
    display: 'flex',
    flexWrap: 'wrap' as const,
    gap: '4px',
    marginTop: '4px',
  },
  // Detail view — compact
  detailBar: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'space-between',
    gap: tokens.spacingHorizontalS,
    background: `linear-gradient(135deg, ${tokens.colorBrandBackground}, ${tokens.colorBrandBackgroundPressed})`,
    color: tokens.colorNeutralForegroundOnBrand,
    ...shorthands.padding(tokens.spacingVerticalS, tokens.spacingHorizontalM),
    borderRadius: `${tokens.borderRadiusMedium} ${tokens.borderRadiusMedium} 0 0`,
  },
  detailBody: {
    ...shorthands.padding(tokens.spacingVerticalM, tokens.spacingHorizontalM),
  },
  formGrid: {
    display: 'grid',
    gridTemplateColumns: '1fr 1fr',
    gap: `${tokens.spacingVerticalS} ${tokens.spacingHorizontalM}`,
  },
  formField: {
    display: 'flex',
    flexDirection: 'column' as const,
    gap: '2px',
  },
  formFieldFull: {
    display: 'flex',
    flexDirection: 'column' as const,
    gap: '2px',
    gridColumn: '1 / -1',
  },
  fieldLabel: {
    color: tokens.colorNeutralForeground3,
    fontSize: '0.7rem',
    fontWeight: '600',
    textTransform: 'uppercase' as const,
    letterSpacing: '0.4px',
  },
  fieldValue: {
    fontWeight: '500',
    fontSize: '0.85rem',
  },
  fieldValueLarge: {
    fontWeight: '700',
    fontSize: '1.1rem',
    color: tokens.colorPaletteGreenForeground1,
  },
  editableValue: {
    cursor: 'pointer',
    borderRadius: '4px',
    ...shorthands.padding('1px', '4px'),
    marginLeft: '-4px',
    display: 'inline-flex',
    alignItems: 'center',
    gap: '4px',
    ':hover': {
      backgroundColor: tokens.colorNeutralBackground3,
      outline: `1px dashed ${tokens.colorBrandStroke1}`,
    },
  },
  editRow: {
    display: 'flex',
    alignItems: 'center',
    gap: '4px',
  },
  editField: { flex: 1, minWidth: 0 },
  miniBtn: {
    minWidth: 'auto',
    ...shorthands.padding('2px'),
  },
  section: {
    marginTop: tokens.spacingVerticalM,
    paddingTop: tokens.spacingVerticalS,
    borderTop: `1px solid ${tokens.colorNeutralStroke2}`,
  },
  sectionTitle: {
    display: 'flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalXS,
    marginBottom: tokens.spacingVerticalS,
  },
  notesList: {
    display: 'flex',
    flexDirection: 'column' as const,
    gap: '4px',
  },
  noteItem: {
    display: 'flex',
    alignItems: 'flex-start',
    gap: tokens.spacingHorizontalXS,
    backgroundColor: tokens.colorNeutralBackground2,
    borderRadius: tokens.borderRadiusMedium,
    ...shorthands.padding('4px', '8px'),
    fontSize: '0.85rem',
  },
  editActions: {
    display: 'flex',
    gap: '4px',
    marginTop: '4px',
    justifyContent: 'flex-end',
  },
  mapContainer: {
    width: '100%',
    height: '180px',
    borderRadius: tokens.borderRadiusMedium,
    overflow: 'hidden',
    border: `1px solid ${tokens.colorNeutralStroke2}`,
    marginTop: tokens.spacingVerticalXS,
  },
  // Bulk action styles
  bulkBar: {
    display: 'flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalS,
    ...shorthands.padding(tokens.spacingVerticalXS, tokens.spacingHorizontalM),
    marginBottom: tokens.spacingVerticalS,
    backgroundColor: tokens.colorNeutralBackground3,
    borderRadius: tokens.borderRadiusMedium,
    flexWrap: 'wrap' as const,
  },
  bulkBarText: {
    fontWeight: '600',
    fontSize: '0.85rem',
    marginRight: 'auto',
  },
  claimCardSelectable: {
    ...shorthands.padding(tokens.spacingVerticalS, tokens.spacingHorizontalM),
    cursor: 'pointer',
    ':hover': { boxShadow: tokens.shadow8 },
    display: 'flex',
    alignItems: 'flex-start',
    gap: tokens.spacingHorizontalS,
  },
  claimCardContent: {
    flex: 1,
    minWidth: 0,
  },
  cardSelected: {
    outline: `2px solid ${tokens.colorBrandStroke1}`,
    backgroundColor: tokens.colorNeutralBackground1Selected,
  },
});

// ─── Helpers ───
const formatCurrency = (n: number) =>
  new Intl.NumberFormat('en-US', {
    style: 'currency',
    currency: 'USD',
    minimumFractionDigits: 0,
    maximumFractionDigits: 0,
  }).format(n || 0);

const formatDate = (d: string | undefined) => {
  if (!d) return 'N/A';
  return new Date(d).toLocaleDateString('en-US', {
    year: 'numeric',
    month: 'short',
    day: 'numeric',
  });
};

const getShortStatus = (s: string | undefined) => {
  if (!s) return 'Unknown';
  return s.split(' - ')[0];
};

const statusBadgeColor = (
  s: string | undefined
): 'warning' | 'informative' | 'success' | 'danger' | 'subtle' => {
  const l = (s || '').toLowerCase();
  if (l.includes('pending')) return 'warning';
  if (l.includes('investigation') || l.includes('review')) return 'informative';
  if (l.includes('approved')) return 'success';
  if (l.includes('repair')) return 'informative';
  if (l.includes('closed')) return 'subtle';
  if (l.includes('denied') || l.includes('rejected')) return 'danger';
  return 'warning';
};

// ─── Clipboard email helpers (copies rich HTML for pasting into email / Copilot) ───
const statusColor = (s: string): string => {
  const l = (s || '').toLowerCase();
  if (l.includes('pending')) return '#f0ad4e';
  if (l.includes('investigation') || l.includes('review')) return '#5bc0de';
  if (l.includes('approved')) return '#5cb85c';
  if (l.includes('repair')) return '#5bc0de';
  if (l.includes('closed')) return '#999';
  if (l.includes('denied') || l.includes('rejected')) return '#d9534f';
  return '#f0ad4e';
};

const claimToHtmlRow = (c: any): string => `
  <tr>
    <td style="padding:6px 10px;border-bottom:1px solid #eee;font-weight:600">${c.claimNumber || ''}</td>
    <td style="padding:6px 10px;border-bottom:1px solid #eee">${c.policyHolderName || ''}</td>
    <td style="padding:6px 10px;border-bottom:1px solid #eee">
      <span style="background:${statusColor(c.status)};color:#fff;padding:2px 8px;border-radius:10px;font-size:12px">${getShortStatus(c.status)}</span>
    </td>
    <td style="padding:6px 10px;border-bottom:1px solid #eee">${c.property || ''}</td>
    <td style="padding:6px 10px;border-bottom:1px solid #eee;text-align:right;font-weight:600;color:#2e7d32">${formatCurrency(c.estimatedLoss || 0)}</td>
    <td style="padding:6px 10px;border-bottom:1px solid #eee">${c.damageTypes || ''}</td>
    <td style="padding:6px 10px;border-bottom:1px solid #eee">${formatDate(c.dateOfLoss)}</td>
  </tr>`;

const buildBulkHtml = (claims: any[]): string => {
  const total = claims.reduce((s, c) => s + (c.estimatedLoss || 0), 0);
  return `
<div style="font-family:Segoe UI,system-ui,sans-serif;max-width:800px">
  <h2 style="color:#0078d4;margin:0 0 4px">Zava Claims Summary</h2>
  <p style="color:#666;margin:0 0 12px">${claims.length} claim${claims.length !== 1 ? 's' : ''} &bull; Total Est. Loss: <strong style="color:#2e7d32">${formatCurrency(total)}</strong></p>
  <table style="border-collapse:collapse;width:100%;font-size:13px">
    <thead>
      <tr style="background:#f5f5f5">
        <th style="padding:8px 10px;text-align:left;border-bottom:2px solid #ddd">Claim #</th>
        <th style="padding:8px 10px;text-align:left;border-bottom:2px solid #ddd">Policy Holder</th>
        <th style="padding:8px 10px;text-align:left;border-bottom:2px solid #ddd">Status</th>
        <th style="padding:8px 10px;text-align:left;border-bottom:2px solid #ddd">Property</th>
        <th style="padding:8px 10px;text-align:right;border-bottom:2px solid #ddd">Est. Loss</th>
        <th style="padding:8px 10px;text-align:left;border-bottom:2px solid #ddd">Damage</th>
        <th style="padding:8px 10px;text-align:left;border-bottom:2px solid #ddd">Loss Date</th>
      </tr>
    </thead>
    <tbody>${claims.map(claimToHtmlRow).join('')}</tbody>
  </table>
  <p style="color:#999;font-size:11px;margin-top:12px">Generated from Zava Claims Dashboard</p>
</div>`;
};

const buildSingleClaimHtml = (c: any): string => `
<div style="font-family:Segoe UI,system-ui,sans-serif;max-width:600px">
  <h2 style="color:#0078d4;margin:0 0 4px">Claim ${c.claimNumber || ''}</h2>
  <p style="color:#666;margin:0 0 12px">${c.policyHolderName || ''} &bull; Policy ${c.policyNumber || ''}</p>
  <table style="border-collapse:collapse;width:100%;font-size:13px">
    <tr><td style="padding:5px 10px;color:#666;width:130px">Status</td><td style="padding:5px 10px"><span style="background:${statusColor(c.status)};color:#fff;padding:2px 8px;border-radius:10px;font-size:12px">${getShortStatus(c.status)}</span></td></tr>
    <tr><td style="padding:5px 10px;color:#666">Property</td><td style="padding:5px 10px">${c.property || 'N/A'}</td></tr>
    <tr><td style="padding:5px 10px;color:#666">Date of Loss</td><td style="padding:5px 10px">${formatDate(c.dateOfLoss)}</td></tr>
    <tr><td style="padding:5px 10px;color:#666">Estimated Loss</td><td style="padding:5px 10px;font-weight:700;color:#2e7d32">${formatCurrency(c.estimatedLoss || 0)}</td></tr>
    <tr><td style="padding:5px 10px;color:#666">Adjuster</td><td style="padding:5px 10px">${c.adjusterAssigned || 'Unassigned'}</td></tr>
    <tr><td style="padding:5px 10px;color:#666">Damage Types</td><td style="padding:5px 10px">${c.damageTypes || 'None'}</td></tr>
    <tr><td style="padding:5px 10px;color:#666;vertical-align:top">Description</td><td style="padding:5px 10px">${c.description || 'N/A'}</td></tr>
    ${c.notes ? `<tr><td style="padding:5px 10px;color:#666;vertical-align:top">Notes</td><td style="padding:5px 10px">${c.notes}</td></tr>` : ''}
  </table>
  <p style="color:#999;font-size:11px;margin-top:12px">Generated from Zava Claims Dashboard</p>
</div>`;

const copyHtmlToClipboard = async (html: string, _plainText: string): Promise<boolean> => {
  // Strategy 1: Clipboard API (works in non-sandboxed contexts)
  try {
    const blob = new Blob([html], { type: 'text/html' });
    const textBlob = new Blob([_plainText], { type: 'text/plain' });
    await navigator.clipboard.write([
      new ClipboardItem({ 'text/html': blob, 'text/plain': textBlob }),
    ]);
    return true;
  } catch { /* blocked in sandboxed iframe — fall through */ }

  // Strategy 2: execCommand('copy') with a hidden rich-HTML selection
  // This works inside ChatGPT / M365 Copilot sandboxed iframes
  try {
    const el = document.createElement('div');
    el.innerHTML = html;
    Object.assign(el.style, {
      position: 'fixed', left: '-9999px', top: '-9999px',
      opacity: '0', pointerEvents: 'none',
    });
    document.body.appendChild(el);

    const range = document.createRange();
    range.selectNodeContents(el);
    const sel = window.getSelection();
    sel?.removeAllRanges();
    sel?.addRange(range);

    const ok = document.execCommand('copy');
    sel?.removeAllRanges();
    document.body.removeChild(el);
    if (ok) return true;
  } catch { /* fall through */ }

  // Strategy 3: plain-text clipboard fallback
  try {
    await navigator.clipboard.writeText(_plainText);
    return true;
  } catch { /* fall through */ }

  // Strategy 4: execCommand with textarea (last resort, plain text only)
  try {
    const ta = document.createElement('textarea');
    ta.value = _plainText;
    Object.assign(ta.style, { position: 'fixed', left: '-9999px', opacity: '0' });
    document.body.appendChild(ta);
    ta.select();
    const ok = document.execCommand('copy');
    document.body.removeChild(ta);
    return ok;
  } catch { return false; }
};

const buildPlainTextSummary = (claims: any[]): string => {
  const total = claims.reduce((s, c) => s + (c.estimatedLoss || 0), 0);
  const header = `Claims Summary (${claims.length} claims, Total: ${formatCurrency(total)})`;
  const rows = claims.map((c, i) =>
    `${i + 1}. ${c.claimNumber} | ${getShortStatus(c.status)} | ${formatCurrency(c.estimatedLoss || 0)} | ${c.policyHolderName} | ${c.property}`
  );
  return [header, '', ...rows].join('\n');
};

const extractData = (raw: any): any[] | null => {
  if (!raw) return null;
  if (raw.success && raw.data) {
    if (Array.isArray(raw.data)) return raw.data;
    if (raw.data.claims) return raw.data.claims;
    if (raw.data.items) return raw.data.items;
    return [raw.data];
  }
  if (raw.claims) return raw.claims;
  if (raw.claim) return [raw.claim];
  if (Array.isArray(raw)) return raw;
  return null;
};

// ─── PropertyMap (pure SVG — no external requests) ───
const PropertyMap: React.FC<{ address: string }> = ({ address }) => {
  const styles = useStyles();
  return (
    <div className={styles.mapContainer} style={{ position: 'relative', background: '#e8f0e8' }}>
      <svg
        viewBox="0 0 600 180"
        width="100%"
        height="100%"
        preserveAspectRatio="xMidYMid slice"
        style={{ position: 'absolute', inset: 0 }}
      >
        <rect width="600" height="180" fill="#d4e6f1" />
        <path d="M0,30 Q80,10 160,25 T320,20 T480,28 T600,15 L600,180 L0,180Z" fill="#c8dcc8" />
        <path d="M0,90 Q100,70 200,85 T400,75 T600,70 L600,180 L0,180Z" fill="#b8ccb8" />
        <line x1="0" y1="110" x2="600" y2="108" stroke="#999" strokeWidth="2" />
        <line x1="150" y1="40" x2="155" y2="180" stroke="#999" strokeWidth="1.5" />
        <line x1="300" y1="50" x2="305" y2="180" stroke="#999" strokeWidth="1.5" />
        <line x1="450" y1="35" x2="445" y2="180" stroke="#999" strokeWidth="1.5" />
        <rect x="170" y="90" width="16" height="16" rx="2" fill="#ddd" stroke="#bbb" strokeWidth="0.5" />
        <rect x="250" y="120" width="22" height="12" rx="2" fill="#ddd" stroke="#bbb" strokeWidth="0.5" />
        <rect x="350" y="85" width="18" height="18" rx="2" fill="#ddd" stroke="#bbb" strokeWidth="0.5" />
        <ellipse cx="500" cy="130" rx="30" ry="20" fill="#a3c9a3" opacity="0.6" />
        <g transform="translate(300, 65)">
          <ellipse cx="0" cy="22" rx="6" ry="2.5" fill="rgba(0,0,0,0.2)" />
          <path d="M0,-18 C-9,-18 -14,-10 -14,-3 C-14,4 0,22 0,22 C0,22 14,4 14,-3 C14,-10 9,-18 0,-18Z" fill="#0078d4" />
          <circle cx="0" cy="-4" r="5" fill="white" />
        </g>
      </svg>
      <div
        style={{
          position: 'absolute',
          bottom: 6,
          left: 6,
          right: 6,
          backgroundColor: 'rgba(255,255,255,0.92)',
          borderRadius: '4px',
          padding: '4px 8px',
          fontSize: '0.75rem',
          fontWeight: 500,
          color: '#333',
          display: 'flex',
          alignItems: 'center',
          gap: '4px',
          boxShadow: '0 1px 3px rgba(0,0,0,0.12)',
        }}
      >
        <Location24Regular style={{ fontSize: 14, color: '#0078d4', flexShrink: 0 }} />
        <span
          style={{
            overflow: 'hidden',
            textOverflow: 'ellipsis',
            whiteSpace: 'nowrap' as const,
          }}
        >
          {address || 'Property location'}
        </span>
      </div>
    </div>
  );
};

// ─── InlineEdit (text / number / currency) ───
const InlineEdit: React.FC<{
  value: any;
  onSave: (v: any) => void;
  type?: 'text' | 'number' | 'currency';
  large?: boolean;
}> = ({ value, onSave, type = 'text', large = false }) => {
  const styles = useStyles();
  const [editing, setEditing] = useState(false);
  const [draft, setDraft] = useState(String(value ?? ''));

  useEffect(() => {
    setDraft(String(value ?? ''));
  }, [value]);

  const commit = () => {
    const v = type === 'number' || type === 'currency' ? parseFloat(draft) || 0 : draft;
    onSave(v);
    setEditing(false);
  };
  const cancel = () => {
    setDraft(String(value ?? ''));
    setEditing(false);
  };

  if (!editing) {
    const display = type === 'currency' ? formatCurrency(value) : value || '—';
    return (
      <Tooltip content="Click to edit" relationship="description">
        <span
          className={styles.editableValue}
          style={large ? { fontWeight: 700, fontSize: '1.1rem', color: tokens.colorPaletteGreenForeground1 } : undefined}
          onClick={() => setEditing(true)}
        >
          {display}{' '}
          <Edit24Regular style={{ fontSize: 12, opacity: 0.4 }} />
        </span>
      </Tooltip>
    );
  }

  return (
    <div className={styles.editRow}>
      <Input
        className={styles.editField}
        size="small"
        value={draft}
        type={type === 'currency' || type === 'number' ? 'number' : 'text'}
        onChange={(_e, d) => setDraft(d.value)}
        autoFocus
        onKeyDown={(e) => {
          if (e.key === 'Enter') commit();
          if (e.key === 'Escape') cancel();
        }}
      />
      <Button
        className={styles.miniBtn}
        appearance="primary"
        size="small"
        icon={<Checkmark16Regular />}
        onClick={commit}
      />
      <Button
        className={styles.miniBtn}
        appearance="subtle"
        size="small"
        icon={<DismissCircle16Regular />}
        onClick={cancel}
      />
    </div>
  );
};

// ─── InlineMultiline (description) ───
const InlineMultiline: React.FC<{
  value: string;
  onSave: (v: string) => void;
}> = ({ value, onSave }) => {
  const styles = useStyles();
  const [editing, setEditing] = useState(false);
  const [draft, setDraft] = useState(value ?? '');

  useEffect(() => {
    setDraft(value ?? '');
  }, [value]);

  const commit = () => {
    onSave(draft);
    setEditing(false);
  };
  const cancel = () => {
    setDraft(value ?? '');
    setEditing(false);
  };

  if (!editing) {
    return (
      <Tooltip content="Click to edit" relationship="description">
        <div
          className={styles.editableValue}
          style={{ display: 'block', lineHeight: 1.5, fontSize: '0.85rem' }}
          onClick={() => setEditing(true)}
        >
          {value || '—'}{' '}
          <Edit24Regular style={{ fontSize: 12, opacity: 0.4, verticalAlign: 'middle' }} />
        </div>
      </Tooltip>
    );
  }

  return (
    <div>
      <Textarea
        size="small"
        value={draft}
        onChange={(_e, d) => setDraft(d.value)}
        resize="vertical"
        style={{ width: '100%', minHeight: '60px' }}
        autoFocus
      />
      <div className={styles.editActions}>
        <Button appearance="primary" size="small" icon={<Save24Regular />} onClick={commit}>
          Save
        </Button>
        <Button appearance="subtle" size="small" icon={<Dismiss24Regular />} onClick={cancel}>
          Cancel
        </Button>
      </div>
    </div>
  );
};

// ─── StatusDropdown ───
const StatusDropdown: React.FC<{
  value: string;
  onSave: (v: string) => void;
}> = ({ value, onSave }) => {
  const styles = useStyles();
  const [editing, setEditing] = useState(false);
  const [draft, setDraft] = useState(value);

  useEffect(() => {
    setDraft(value);
  }, [value]);

  const commit = () => {
    if (draft !== value) onSave(draft);
    setEditing(false);
  };
  const cancel = () => {
    setDraft(value);
    setEditing(false);
  };

  if (!editing) {
    return (
      <Tooltip content="Click to change status" relationship="description">
        <span className={styles.editableValue} onClick={() => setEditing(true)}>
          <Badge appearance="filled" color={statusBadgeColor(value)} size="medium">
            {getShortStatus(value)}
          </Badge>
          <Edit24Regular style={{ fontSize: 12, opacity: 0.4 }} />
        </span>
      </Tooltip>
    );
  }

  return (
    <div className={styles.editRow}>
      <Select
        className={styles.editField}
        size="small"
        value={draft}
        onChange={(_e, d) => setDraft(d.value)}
        autoFocus
      >
        {STATUS_OPTIONS.map((s) => (
          <option key={s} value={s}>
            {s}
          </option>
        ))}
        {/* include current value if not in the list */}
        {!STATUS_OPTIONS.includes(value) && (
          <option value={value}>{value}</option>
        )}
      </Select>
      <Button
        className={styles.miniBtn}
        appearance="primary"
        size="small"
        icon={<Checkmark16Regular />}
        onClick={commit}
      />
      <Button
        className={styles.miniBtn}
        appearance="subtle"
        size="small"
        icon={<DismissCircle16Regular />}
        onClick={cancel}
      />
    </div>
  );
};

// ─── StatsBar ───
const StatsBar: React.FC<{ claims: any[] }> = ({ claims }) => {
  const styles = useStyles();
  const total = claims.length;
  const pending = claims.filter((c) => (c.status || '').toLowerCase().includes('pending')).length;
  const investigation = claims.filter(
    (c) => (c.status || '').toLowerCase().includes('investigation')
  ).length;
  const totalValue = claims.reduce((s, c) => s + (c.estimatedLoss || 0), 0);

  return (
    <div className={styles.statsGrid}>
      {[
        { label: 'Total', value: total, color: tokens.colorBrandForeground1 },
        { label: 'Pending', value: pending, color: tokens.colorPaletteYellowForeground1 },
        { label: 'Review', value: investigation, color: tokens.colorBrandForeground1 },
        {
          label: 'Value',
          value: formatCurrency(totalValue),
          color: tokens.colorPaletteGreenForeground1,
        },
      ].map((s, i) => (
        <Card key={i} className={styles.statCard}>
          <div className={styles.statValue} style={{ color: s.color }}>
            {s.value}
          </div>
          <div className={styles.statLabel}>{s.label}</div>
        </Card>
      ))}
    </div>
  );
};

// ─── ClaimCard (with optional checkbox for bulk selection) ───
const ClaimCard: React.FC<{
  claim: any;
  onClick: (c: any) => void;
  selectable?: boolean;
  selected?: boolean;
  onToggle?: (id: string) => void;
}> = ({ claim, onClick, selectable, selected, onToggle }) => {
  const styles = useStyles();
  const key = claim.id || claim.claimNumber;

  if (selectable) {
    return (
      <Card
        className={`${styles.claimCardSelectable} ${selected ? styles.cardSelected : ''}`}
        onClick={() => onClick(claim)}
      >
        <div
          onClick={(e) => { e.stopPropagation(); onToggle?.(key); }}
          style={{ paddingTop: '2px' }}
        >
          <Checkbox checked={!!selected} />
        </div>
        <div className={styles.claimCardContent}>
          <div className={styles.claimCardHeader}>
            <div>
              <Subtitle1>{claim.claimNumber}</Subtitle1>
              <Caption1 style={{ display: 'block', color: tokens.colorNeutralForeground3 }}>
                {claim.policyHolderName}
              </Caption1>
            </div>
            <Badge appearance="filled" color={statusBadgeColor(claim.status)} size="small">
              {getShortStatus(claim.status)}
            </Badge>
          </div>
          <div className={styles.detailGrid}>
            <div>
              <Caption1 style={{ color: tokens.colorNeutralForeground3 }}>Property</Caption1>
              <Body2 style={{ display: 'block' }}>{claim.property}</Body2>
            </div>
            <div>
              <Caption1 style={{ color: tokens.colorNeutralForeground3 }}>Loss Date</Caption1>
              <Body2 style={{ display: 'block' }}>{formatDate(claim.dateOfLoss)}</Body2>
            </div>
            <div>
              <Caption1 style={{ color: tokens.colorNeutralForeground3 }}>Est. Loss</Caption1>
              <Body2 style={{ display: 'block', color: tokens.colorPaletteGreenForeground1, fontWeight: '700' }}>
                {formatCurrency(claim.estimatedLoss)}
              </Body2>
            </div>
          </div>
          {claim.damageTypes && (
            <div className={styles.damageTags}>
              {claim.damageTypes.split(',').map((d: string, i: number) => (
                <Badge key={i} appearance="outline" size="small">{d.trim()}</Badge>
              ))}
            </div>
          )}
        </div>
      </Card>
    );
  }

  return (
    <Card className={styles.claimCard} onClick={() => onClick(claim)}>
      <div className={styles.claimCardHeader}>
        <div>
          <Subtitle1>{claim.claimNumber}</Subtitle1>
          <Caption1 style={{ display: 'block', color: tokens.colorNeutralForeground3 }}>
            {claim.policyHolderName}
          </Caption1>
        </div>
        <Badge appearance="filled" color={statusBadgeColor(claim.status)} size="small">
          {getShortStatus(claim.status)}
        </Badge>
      </div>
      <div className={styles.detailGrid}>
        <div>
          <Caption1 style={{ color: tokens.colorNeutralForeground3 }}>Property</Caption1>
          <Body2 style={{ display: 'block' }}>{claim.property}</Body2>
        </div>
        <div>
          <Caption1 style={{ color: tokens.colorNeutralForeground3 }}>Loss Date</Caption1>
          <Body2 style={{ display: 'block' }}>{formatDate(claim.dateOfLoss)}</Body2>
        </div>
        <div>
          <Caption1 style={{ color: tokens.colorNeutralForeground3 }}>Est. Loss</Caption1>
          <Body2 style={{ display: 'block', color: tokens.colorPaletteGreenForeground1, fontWeight: '700' }}>
            {formatCurrency(claim.estimatedLoss)}
          </Body2>
        </div>
      </div>
      {claim.damageTypes && (
        <div className={styles.damageTags}>
          {claim.damageTypes.split(',').map((d: string, i: number) => (
            <Badge key={i} appearance="outline" size="small">{d.trim()}</Badge>
          ))}
        </div>
      )}
    </Card>
  );
};

// ─── BulkActionBar ───
const BulkActionBar: React.FC<{
  items: any[];
  selectedIds: Set<string>;
  onToggleAll: () => void;
  onClear: () => void;
  onBulkStatus: (status: string) => void;
  onCopyForEmail: () => void;
  saving?: boolean;
  copied?: boolean;
}> = ({ items, selectedIds, onToggleAll, onClear, onBulkStatus, onCopyForEmail, saving, copied }) => {
  const styles = useStyles();
  const [statusDraft, setStatusDraft] = useState('');
  const count = selectedIds.size;
  const allSelected = count === items.length && items.length > 0;

  return (
    <div className={styles.bulkBar}>
      <Checkbox
        checked={allSelected ? true : count > 0 ? 'mixed' : false}
        onChange={onToggleAll}
        label={count > 0 ? `${count} selected` : 'Select all'}
      />
      {count > 0 && (
        <>
          <Button appearance="subtle" size="small" icon={<DismissSquare24Regular />} onClick={onClear}>
            Clear
          </Button>
          <Select
            size="small"
            value={statusDraft}
            onChange={(_e, d) => setStatusDraft(d.value)}
            style={{ minWidth: '160px' }}
          >
            <option value="">Change status…</option>
            {STATUS_OPTIONS.map((s) => (
              <option key={s} value={s}>{getShortStatus(s)}</option>
            ))}
          </Select>
          {statusDraft && (
            <Button
              appearance="primary"
              size="small"
              disabled={saving}
              onClick={() => { onBulkStatus(statusDraft); setStatusDraft(''); }}
            >
              {saving ? <Spinner size="tiny" /> : 'Apply'}
            </Button>
          )}
          <Button
            appearance={copied ? 'primary' : 'subtle'}
            size="small"
            icon={<ClipboardPaste24Regular />}
            onClick={onCopyForEmail}
          >
            {copied ? 'Copied!' : 'Copy for Email'}
          </Button>
        </>
      )}
    </div>
  );
};

// ─── ClaimDetailView (compact, single-scroll, no tabs) ───
const ClaimDetailView: React.FC<{
  claim: any;
  onBack: () => void;
  onClaimUpdated?: (updated: any) => void;
  callTool: (toolName: string, args: Record<string, unknown>) => Promise<any>;
}> = ({
  claim: initialClaim,
  onBack,
  onClaimUpdated,
  callTool,
}) => {
  const styles = useStyles();
  const [claim, setClaim] = useState(initialClaim);
  const [saving, setSaving] = useState(false);
  const [saveMsg, setSaveMsg] = useState<{ type: string; text: string } | null>(null);

  const updateField = async (field: string, value: any) => {
    const prev = { ...claim };
    setClaim((p: any) => ({ ...p, [field]: value, updatedAt: new Date().toISOString() }));
    setSaving(true);
    setSaveMsg(null);

    try {
      const json = await callTool('update_claim', { claimId: claim.id, [field]: value });

      if (json.success && json.data) {
        setClaim(json.data);
        onClaimUpdated?.(json.data);
        setSaveMsg({ type: 'success', text: `${field} updated` });
      } else {
        throw new Error(json.error || 'Update failed');
      }
    } catch (err: any) {
      console.error('Update failed:', err);
      setSaveMsg({ type: 'warning', text: `Save failed: ${err.message || 'unknown error'}` });
      setClaim(prev);
    } finally {
      setSaving(false);
      setTimeout(() => setSaveMsg(null), 3000);
    }
  };

  return (
    <div>
      <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: tokens.spacingVerticalXS }}>
        <Button appearance="subtle" size="small" icon={<ArrowLeft24Regular />} onClick={onBack}>
          Back
        </Button>
        <Button
          appearance={saveMsg?.type === 'copied' ? 'primary' : 'subtle'}
          size="small"
          icon={<ClipboardPaste24Regular />}
          onClick={async () => {
            const html = buildSingleClaimHtml(claim);
            const plain = `Claim ${claim.claimNumber} | ${getShortStatus(claim.status)} | ${formatCurrency(claim.estimatedLoss || 0)} | ${claim.policyHolderName}`;
            const ok = await copyHtmlToClipboard(html, plain);
            setSaveMsg(ok ? { type: 'copied', text: 'Copied! Paste into email or tell Copilot to embed it.' } : { type: 'warning', text: 'Copy failed — try again' });
            setTimeout(() => setSaveMsg(null), 4000);
          }}
        >
          {saveMsg?.type === 'copied' ? 'Copied!' : 'Copy for Email'}
        </Button>
      </div>

      {saveMsg && (
        <div style={{ marginBottom: tokens.spacingVerticalXS }}>
          <MessageBar intent={saveMsg.type === 'success' ? 'success' : 'warning'}>
            <MessageBarBody>{saveMsg.text}</MessageBarBody>
          </MessageBar>
        </div>
      )}

      <Card style={{ overflow: 'hidden' }}>
        {/* Compact header bar */}
        <div className={styles.detailBar}>
          <div>
            <Subtitle1 style={{ color: 'inherit' }}>{claim.claimNumber}</Subtitle1>
            <Caption1 style={{ color: 'rgba(255,255,255,0.85)', display: 'block' }}>
              {claim.policyHolderName} &bull; {claim.policyNumber}
            </Caption1>
          </div>
          <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
            {saving && <Spinner size="tiny" />}
          </div>
        </div>

        <div className={styles.detailBody}>
          {/* Status + key fields grid */}
          <div className={styles.formGrid}>
            <div className={styles.formField}>
              <div className={styles.fieldLabel}>Status</div>
              <StatusDropdown value={claim.status} onSave={(v) => updateField('status', v)} />
            </div>
            <div className={styles.formField}>
              <div className={styles.fieldLabel}>Estimated Loss</div>
              <InlineEdit
                value={claim.estimatedLoss}
                type="currency"
                large
                onSave={(v) => updateField('estimatedLoss', v)}
              />
            </div>
            <div className={styles.formField}>
              <div className={styles.fieldLabel}>Adjuster</div>
              <InlineEdit
                value={claim.adjusterAssigned || 'Unassigned'}
                onSave={(v) => updateField('adjusterAssigned', v)}
              />
            </div>
            <div className={styles.formField}>
              <div className={styles.fieldLabel}>Date of Loss</div>
              <div className={styles.fieldValue}>{formatDate(claim.dateOfLoss)}</div>
            </div>
            <div className={styles.formField}>
              <div className={styles.fieldLabel}>Date Reported</div>
              <div className={styles.fieldValue}>{formatDate(claim.dateReported)}</div>
            </div>
            <div className={styles.formField}>
              <div className={styles.fieldLabel}>Last Updated</div>
              <div className={styles.fieldValue}>{formatDate(claim.updatedAt)}</div>
            </div>
          </div>

          {/* Description */}
          <div className={styles.section}>
            <div className={styles.sectionTitle}>
              <Subtitle2>Description</Subtitle2>
            </div>
            <InlineMultiline
              value={claim.description}
              onSave={(v) => updateField('description', v)}
            />
          </div>

          {/* Property + Map */}
          <div className={styles.section}>
            <div className={styles.sectionTitle}>
              <Location24Regular style={{ fontSize: 18 }} />
              <Subtitle2>Property</Subtitle2>
            </div>
            <div className={styles.fieldValue} style={{ marginBottom: '4px' }}>
              {claim.property}
            </div>
            <Caption1 style={{ color: tokens.colorNeutralForeground3 }}>
              {claim.policyHolderEmail}
            </Caption1>
            <PropertyMap address={claim.property} />
          </div>

          {/* Damage Types */}
          {claim.damageTypes && (
            <div className={styles.section}>
              <div className={styles.sectionTitle}>
                <Subtitle2>Damage Types</Subtitle2>
              </div>
              <InlineMultiline
                value={claim.damageTypes}
                onSave={(v) => updateField('damageTypes', v)}
              />
            </div>
          )}

          {/* Notes */}
          <div className={styles.section}>
            <div className={styles.sectionTitle}>
              <Subtitle2>Notes</Subtitle2>
            </div>
            <InlineMultiline
              value={claim.notes || ''}
              onSave={(v) => updateField('notes', v)}
            />
          </div>
        </div>
      </Card>
    </div>
  );
};

// ─── Main App ───
const App = () => {
  const styles = useStyles();
  const isDark = useHostTheme();
  const toolResult = useToolResult();
  const callTool = useCallTool();
  const [items, setItems] = useState<any[]>([]);
  const [loading, setLoading] = useState(true);
  const [selected, setSelected] = useState<any | null>(null);
  const [bulkMode, setBulkMode] = useState(false);
  const [selectedIds, setSelectedIds] = useState<Set<string>>(new Set());
  const [bulkSaving, setBulkSaving] = useState(false);
  const [copied, setCopied] = useState(false);

  const toggleBulkId = (id: string) => {
    setSelectedIds((prev) => {
      const next = new Set(prev);
      if (next.has(id)) next.delete(id); else next.add(id);
      return next;
    });
  };
  const toggleAll = () => {
    if (selectedIds.size === items.length) {
      setSelectedIds(new Set());
    } else {
      setSelectedIds(new Set(items.map((c) => c.id || c.claimNumber)));
    }
  };
  const clearSelection = () => { setSelectedIds(new Set()); setBulkMode(false); };

  const handleBulkStatus = async (status: string) => {
    setBulkSaving(true);
    const ids = Array.from(selectedIds);

    // Fire all update_claim calls in parallel via MCP Apps bridge
    const promises = ids.map((id) =>
      callTool('update_claim', { claimId: id, status })
        .then((json: any) => {
          if (json.success && json.data) return json.data;
          console.error(`Update rejected for ${id}:`, json);
          return null;
        })
        .catch((e: any) => { console.error(`Bulk update failed for ${id}:`, e); return null; })
    );

    const settled = await Promise.allSettled(promises);
    const results = settled
      .filter((r): r is PromiseFulfilledResult<any> => r.status === 'fulfilled' && r.value !== null)
      .map((r) => r.value);

    if (results.length) {
      setItems((prev) => prev.map((c) => {
        const key = c.id || c.claimNumber;
        const updated = results.find((r) => r.id === key || r.claimNumber === key);
        return updated ? { ...c, ...updated } : c;
      }));
    }
    setBulkSaving(false);
    setSelectedIds(new Set());
  };

  const handleBulkEmail = async () => {
    const sel = items.filter((c) => selectedIds.has(c.id || c.claimNumber));
    const html = buildBulkHtml(sel);
    const plain = buildPlainTextSummary(sel);
    const ok = await copyHtmlToClipboard(html, plain);
    if (ok) {
      setCopied(true);
      setTimeout(() => setCopied(false), 3000);
    }
  };

  const handleClaimUpdated = (updated: any) => {
    setSelected(updated);
    setItems(prev => prev.map(c =>
      (c.id && c.id === updated.id) || (c.claimNumber && c.claimNumber === updated.claimNumber)
        ? updated : c
    ));
  };

  // React to tool results from the MCP Apps bridge (or legacy fallback)
  useEffect(() => {
    if (toolResult != null) {
      const r = extractData(toolResult);
      if (r) {
        setItems(r);
        setLoading(false);
        if (r.length === 1) setSelected(r[0]);
      }
    }
  }, [toolResult]);

  // If no data arrives within 5s, stop the spinner
  useEffect(() => {
    const timer = setTimeout(() => setLoading(false), 5000);
    return () => clearTimeout(timer);
  }, []);

  return (
    <FluentProvider theme={isDark ? webDarkTheme : webLightTheme}>
      <div className={styles.root}>
        {loading ? (
          <div className={styles.centered}>
            <Spinner label="Loading claims..." />
          </div>
        ) : items.length === 0 ? (
          <MessageBar intent="info">
            <MessageBarBody>No claims data found.</MessageBarBody>
          </MessageBar>
        ) : selected ? (
          <>
            <div className={styles.header}>
              <div className={styles.headerLeft}>
                <div className={styles.logo}>Z</div>
                <div>
                  <Title3>Claim Dashboard</Title3>
                  <Caption1
                    style={{ display: 'block', color: tokens.colorNeutralForeground3 }}
                  >
                    Editing {selected.claimNumber}
                  </Caption1>
                </div>
              </div>
            </div>
            <ClaimDetailView claim={selected} onBack={() => setSelected(null)} onClaimUpdated={handleClaimUpdated} callTool={callTool} />
          </>
        ) : (
          <>
            <div className={styles.header}>
              <div className={styles.headerLeft}>
                <div className={styles.logo}>Z</div>
                <div>
                  <Title3>Claim Dashboard</Title3>
                  <Caption1
                    style={{ display: 'block', color: tokens.colorNeutralForeground3 }}
                  >
                    {items.length} claim{items.length !== 1 ? 's' : ''} found
                  </Caption1>
                </div>
              </div>
            </div>
            <StatsBar claims={items} />
            <div style={{ display: 'flex', justifyContent: 'flex-end', marginBottom: tokens.spacingVerticalXS }}>
              <Button
                appearance={bulkMode ? 'primary' : 'subtle'}
                size="small"
                icon={bulkMode ? <CheckboxChecked24Regular /> : <CheckboxUnchecked24Regular />}
                onClick={() => { setBulkMode(!bulkMode); if (bulkMode) clearSelection(); }}
              >
                {bulkMode ? 'Exit Bulk' : 'Bulk Edit'}
              </Button>
            </div>
            {bulkMode && (
              <BulkActionBar
                items={items}
                selectedIds={selectedIds}
                onToggleAll={toggleAll}
                onClear={clearSelection}
                onBulkStatus={handleBulkStatus}
                onCopyForEmail={handleBulkEmail}
                saving={bulkSaving}
                copied={copied}
              />
            )}
            <div className={styles.claimsList}>
              {items.map((c) => (
                <ClaimCard
                  key={c.id || c.claimNumber}
                  claim={c}
                  onClick={setSelected}
                  selectable={bulkMode}
                  selected={selectedIds.has(c.id || c.claimNumber)}
                  onToggle={toggleBulkId}
                />
              ))}
            </div>
          </>
        )}
      </div>
    </FluentProvider>
  );
};

// ─── Mount ───
createRoot(document.getElementById('root')!).render(<App />);
