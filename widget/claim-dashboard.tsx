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
} from '@fluentui/react-icons';

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

// ─── EditableList (notes / damageTypes) ───
const EditableList: React.FC<{
  items: string[];
  onSave: (v: string[]) => void;
  label: string;
}> = ({ items, onSave, label }) => {
  const styles = useStyles();
  const [editing, setEditing] = useState(false);
  const [draft, setDraft] = useState((items || []).join('\n'));

  useEffect(() => {
    setDraft((items || []).join('\n'));
  }, [items]);

  const commit = () => {
    const list = draft
      .split('\n')
      .map((s) => s.trim())
      .filter(Boolean);
    onSave(list);
    setEditing(false);
  };
  const cancel = () => {
    setDraft((items || []).join('\n'));
    setEditing(false);
  };

  if (!editing) {
    return (
      <div>
        <div className={styles.notesList}>
          {(items || []).map((item, i) => (
            <div key={i} className={styles.noteItem}>
              • {item}
            </div>
          ))}
          {(!items || items.length === 0) && (
            <Caption1 style={{ color: tokens.colorNeutralForeground4 }}>None</Caption1>
          )}
        </div>
        <Tooltip content={`Edit ${label.toLowerCase()}`} relationship="description">
          <Button
            appearance="subtle"
            size="small"
            icon={<Edit24Regular />}
            onClick={() => setEditing(true)}
            style={{ marginTop: '4px' }}
          >
            Edit
          </Button>
        </Tooltip>
      </div>
    );
  }

  return (
    <div>
      <Textarea
        size="small"
        value={draft}
        onChange={(_e, d) => setDraft(d.value)}
        resize="vertical"
        placeholder={`One ${label.toLowerCase().replace(/s$/, '')} per line`}
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

// ─── ClaimCard ───
const ClaimCard: React.FC<{ claim: any; onClick: (c: any) => void }> = ({ claim, onClick }) => {
  const styles = useStyles();
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
          <Body2
            style={{
              display: 'block',
              color: tokens.colorPaletteGreenForeground1,
              fontWeight: '700',
            }}
          >
            {formatCurrency(claim.estimatedLoss)}
          </Body2>
        </div>
      </div>
      {claim.damageTypes && claim.damageTypes.length > 0 && (
        <div className={styles.damageTags}>
          {claim.damageTypes.map((d: string, i: number) => (
            <Badge key={i} appearance="outline" size="small">
              {d}
            </Badge>
          ))}
        </div>
      )}
    </Card>
  );
};

// ─── ClaimDetailView (compact, single-scroll, no tabs) ───
const ClaimDetailView: React.FC<{ claim: any; onBack: () => void }> = ({
  claim: initialClaim,
  onBack,
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
      const baseUrl = (window as any).__MCP_SERVER_URL__ || window.location.origin;
      const res = await fetch(`${baseUrl}/mcp/tools/call`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          name: 'update_claim',
          arguments: { claimId: claim.id, [field]: value },
        }),
      });

      if (!res.ok) {
        const text = await res.text();
        let msg = `Server returned ${res.status}`;
        try { msg = JSON.parse(text).error || msg; } catch {}
        throw new Error(msg);
      }

      const json = await res.json();

      if (json.success && json.data) {
        setClaim(json.data);
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
      <Button
        appearance="subtle"
        size="small"
        icon={<ArrowLeft24Regular />}
        onClick={onBack}
        style={{ marginBottom: tokens.spacingVerticalXS }}
      >
        Back
      </Button>

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
          {claim.damageTypes && claim.damageTypes.length > 0 && (
            <div className={styles.section}>
              <div className={styles.sectionTitle}>
                <Subtitle2>Damage Types</Subtitle2>
              </div>
              <EditableList
                items={claim.damageTypes}
                label="Damage Types"
                onSave={(v) => updateField('damageTypes', v)}
              />
            </div>
          )}

          {/* Notes */}
          <div className={styles.section}>
            <div className={styles.sectionTitle}>
              <Subtitle2>Notes</Subtitle2>
            </div>
            <EditableList
              items={claim.notes}
              label="Notes"
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
  const [isDark, setIsDark] = useState(false);
  const [items, setItems] = useState<any[]>([]);
  const [loading, setLoading] = useState(true);
  const [selected, setSelected] = useState<any | null>(null);

  useEffect(() => {
    if ((window as any).openai?.theme === 'dark') setIsDark(true);

    const handleSetGlobals = (e: any) => {
      const t = e.detail?.globals?.theme;
      if (t) setIsDark(t === 'dark');
      const d = e.detail?.globals?.toolOutput;
      if (d) {
        const r = extractData(d);
        if (r) {
          setItems(r);
          setLoading(false);
          if (r.length === 1) setSelected(r[0]);
        }
      }
    };
    window.addEventListener('openai:set_globals', handleSetGlobals);

    if (
      !(window as any).openai?.theme &&
      window.matchMedia?.('(prefers-color-scheme: dark)').matches
    ) {
      setIsDark(true);
    }

    const load = () => {
      if ((window as any).openai?.toolOutput) {
        const r = extractData((window as any).openai.toolOutput);
        if (r) {
          setItems(r);
          setLoading(false);
          if (r.length === 1) setSelected(r[0]);
          return true;
        }
      }
      return false;
    };
    if (!load()) {
      const iv = setInterval(() => {
        if (load()) clearInterval(iv);
      }, 100);
      setTimeout(() => {
        clearInterval(iv);
        setLoading(false);
      }, 5000);
    }

    return () => window.removeEventListener('openai:set_globals', handleSetGlobals);
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
            <ClaimDetailView claim={selected} onBack={() => setSelected(null)} />
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
            <div className={styles.claimsList}>
              {items.map((c) => (
                <ClaimCard key={c.id || c.claimNumber} claim={c} onClick={setSelected} />
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
