import React, { useState, useEffect, useRef } from 'react';
import { createRoot } from 'react-dom/client';
import {
  FluentProvider,
  webLightTheme,
  webDarkTheme,
  Card,
  Text,
  Badge,
  Button,
  Spinner,
  Input,
  Textarea,
  Divider,
  makeStyles,
  tokens,
  shorthands,
  Title3,
  Subtitle1,
  Subtitle2,
  Body1,
  Caption1,
  MessageBar,
  MessageBarBody,
  Tooltip,
  TabList,
  Tab,
} from '@fluentui/react-components';
import {
  ArrowLeft24Regular,
  Edit24Regular,
  Save24Regular,
  Dismiss24Regular,
  Location24Regular,
  DocumentText24Regular,
  Note24Regular,
  Warning24Regular,
  Clock24Regular,
} from '@fluentui/react-icons';
import L from 'leaflet';
import 'leaflet/dist/leaflet.css';

// ─── Fix Leaflet default icon paths (broken by bundlers) ───
// Use inline SVG data URIs so there are zero external requests
const markerSvg = `data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='25' height='41' viewBox='0 0 25 41'%3E%3Cpath d='M12.5 0C5.6 0 0 5.6 0 12.5c0 2.4.7 4.7 1.9 6.6L12.5 41l10.6-21.9c1.2-1.9 1.9-4.2 1.9-6.6C25 5.6 19.4 0 12.5 0z' fill='%232D89EF'/%3E%3Ccircle cx='12.5' cy='12.5' r='5' fill='white'/%3E%3C/svg%3E`;
const shadowSvg = `data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='41' height='41' viewBox='0 0 41 41'%3E%3Cellipse cx='13' cy='38' rx='13' ry='3' fill='rgba(0,0,0,0.2)'/%3E%3C/svg%3E`;

delete (L.Icon.Default.prototype as any)._getIconUrl;
L.Icon.Default.mergeOptions({
  iconUrl: markerSvg,
  iconRetinaUrl: markerSvg,
  shadowUrl: shadowSvg,
  iconSize: [25, 41],
  iconAnchor: [12, 41],
  popupAnchor: [1, -34],
  shadowSize: [41, 41],
});

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
  header: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'space-between',
    gap: tokens.spacingHorizontalM,
    marginBottom: tokens.spacingVerticalL,
    paddingBottom: tokens.spacingVerticalM,
    borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
  },
  headerLeft: {
    display: 'flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalM,
  },
  logo: {
    width: '44px',
    height: '44px',
    background: `linear-gradient(135deg, ${tokens.colorBrandBackground}, ${tokens.colorBrandBackgroundPressed})`,
    borderRadius: tokens.borderRadiusMedium,
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    color: tokens.colorNeutralForegroundOnBrand,
    fontWeight: '700',
    fontSize: '18px',
    boxShadow: tokens.shadow4,
  },
  statsGrid: {
    display: 'grid',
    gridTemplateColumns: 'repeat(auto-fit, minmax(140px, 1fr))',
    gap: tokens.spacingHorizontalM,
    marginBottom: tokens.spacingVerticalL,
  },
  statCard: {
    ...shorthands.padding(tokens.spacingVerticalM, tokens.spacingHorizontalM),
    textAlign: 'center' as const,
  },
  statValue: { fontWeight: '700', fontSize: '1.5rem' },
  statLabel: {
    color: tokens.colorNeutralForeground3,
    fontSize: '0.75rem',
    textTransform: 'uppercase' as const,
    letterSpacing: '0.5px',
    marginTop: '4px',
  },
  claimsList: {
    display: 'flex',
    flexDirection: 'column' as const,
    gap: tokens.spacingVerticalM,
  },
  claimCard: {
    ...shorthands.padding(tokens.spacingVerticalM, tokens.spacingHorizontalM),
    cursor: 'pointer',
    ':hover': { boxShadow: tokens.shadow8 },
  },
  claimCardHeader: {
    display: 'flex',
    alignItems: 'flex-start',
    justifyContent: 'space-between',
    gap: tokens.spacingHorizontalM,
    marginBottom: tokens.spacingVerticalS,
  },
  detailGrid: {
    display: 'grid',
    gridTemplateColumns: 'repeat(auto-fit, minmax(200px, 1fr))',
    gap: tokens.spacingHorizontalS,
    marginTop: tokens.spacingVerticalS,
  },
  damageTags: {
    display: 'flex',
    flexWrap: 'wrap' as const,
    gap: '6px',
    marginTop: tokens.spacingVerticalS,
  },
  detailHeader: {
    background: `linear-gradient(135deg, ${tokens.colorBrandBackground}, ${tokens.colorBrandBackgroundPressed})`,
    color: tokens.colorNeutralForegroundOnBrand,
    ...shorthands.padding(tokens.spacingVerticalL, tokens.spacingHorizontalL),
    borderRadius: `${tokens.borderRadiusXLarge} ${tokens.borderRadiusXLarge} 0 0`,
  },
  detailBody: {
    ...shorthands.padding(tokens.spacingVerticalL, tokens.spacingHorizontalL),
  },
  section: { marginBottom: tokens.spacingVerticalL },
  sectionTitle: {
    display: 'flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalS,
    marginBottom: tokens.spacingVerticalM,
    paddingBottom: tokens.spacingVerticalS,
    borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
  },
  infoGrid: {
    display: 'grid',
    gridTemplateColumns: 'repeat(auto-fit, minmax(200px, 1fr))',
    gap: tokens.spacingHorizontalM,
  },
  infoItem: {
    backgroundColor: tokens.colorNeutralBackground2,
    borderRadius: tokens.borderRadiusMedium,
    ...shorthands.padding(tokens.spacingVerticalM, tokens.spacingHorizontalM),
  },
  infoLabel: {
    color: tokens.colorNeutralForeground3,
    fontSize: '0.75rem',
    textTransform: 'uppercase' as const,
    letterSpacing: '0.5px',
    marginBottom: '4px',
  },
  infoValue: { fontWeight: '500' },
  infoValueLarge: {
    fontWeight: '700',
    fontSize: '1.5rem',
    color: tokens.colorPaletteGreenForeground1,
  },
  mapContainer: {
    width: '100%',
    height: '260px',
    borderRadius: tokens.borderRadiusMedium,
    overflow: 'hidden',
    border: `1px solid ${tokens.colorNeutralStroke2}`,
    marginTop: tokens.spacingVerticalS,
  },
  descriptionBox: {
    backgroundColor: tokens.colorNeutralBackground2,
    borderRadius: tokens.borderRadiusMedium,
    ...shorthands.padding(tokens.spacingVerticalM, tokens.spacingHorizontalM),
    lineHeight: '1.6',
  },
  notesList: {
    display: 'flex',
    flexDirection: 'column' as const,
    gap: tokens.spacingVerticalS,
  },
  noteItem: {
    display: 'flex',
    alignItems: 'flex-start',
    gap: tokens.spacingHorizontalS,
    backgroundColor: tokens.colorNeutralBackground2,
    borderRadius: tokens.borderRadiusMedium,
    ...shorthands.padding(tokens.spacingVerticalS, tokens.spacingHorizontalM),
  },
  editRow: {
    display: 'flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalS,
  },
  editField: { flex: 1 },
  editActions: {
    display: 'flex',
    gap: tokens.spacingHorizontalXS,
    marginTop: tokens.spacingVerticalM,
    justifyContent: 'flex-end',
  },
  editableValue: {
    cursor: 'pointer',
    borderRadius: tokens.borderRadiusMedium,
    ...shorthands.padding('2px', '4px'),
    ':hover': {
      backgroundColor: tokens.colorNeutralBackground3,
      outline: `1px dashed ${tokens.colorBrandStroke1}`,
    },
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

const statusBadgeColor = (s: string | undefined): 'warning' | 'informative' | 'success' | 'danger' | 'subtle' => {
  const l = (s || '').toLowerCase();
  if (l.includes('pending')) return 'warning';
  if (l.includes('investigation') || l.includes('review')) return 'informative';
  if (l.includes('approved') || l.includes('settled')) return 'success';
  if (l.includes('denied') || l.includes('rejected')) return 'danger';
  if (l.includes('closed')) return 'subtle';
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

// ─── PropertyMap ───
const PropertyMap: React.FC<{ address: string }> = ({ address }) => {
  const styles = useStyles();
  const mapRef = useRef<HTMLDivElement>(null);
  const leafletMap = useRef<L.Map | null>(null);

  useEffect(() => {
    if (!address || !mapRef.current) return;

    const geocode = async () => {
      try {
        const res = await fetch(
          `https://nominatim.openstreetmap.org/search?format=json&q=${encodeURIComponent(address)}&limit=1`,
          { headers: { 'User-Agent': 'ZavaInsuranceWidget/1.0' } }
        );
        const results = await res.json();
        if (results && results.length > 0) {
          const { lat, lon } = results[0];
          if (leafletMap.current) leafletMap.current.remove();
          const map = L.map(mapRef.current!).setView([parseFloat(lat), parseFloat(lon)], 15);
          L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png', {
            attribution: '&copy; OpenStreetMap contributors',
          }).addTo(map);
          L.marker([parseFloat(lat), parseFloat(lon)])
            .addTo(map)
            .bindPopup(address)
            .openPopup();
          leafletMap.current = map;
        }
      } catch (err) {
        console.error('Geocoding failed:', err);
      }
    };
    geocode();

    return () => {
      if (leafletMap.current) {
        leafletMap.current.remove();
        leafletMap.current = null;
      }
    };
  }, [address]);

  return <div ref={mapRef} className={styles.mapContainer} />;
};

// ─── EditableField ───
const EditableField: React.FC<{
  value: any;
  onSave: (v: any) => void;
  type?: 'text' | 'number' | 'currency';
  multiline?: boolean;
}> = ({ value, onSave, type = 'text', multiline = false }) => {
  const styles = useStyles();
  const [editing, setEditing] = useState(false);
  const [draft, setDraft] = useState(String(value ?? ''));

  useEffect(() => {
    setDraft(String(value ?? ''));
  }, [value]);

  const commit = () => {
    onSave(type === 'number' || type === 'currency' ? parseFloat(draft) || 0 : draft);
    setEditing(false);
  };
  const cancel = () => {
    setDraft(String(value ?? ''));
    setEditing(false);
  };

  if (!editing) {
    return (
      <Tooltip content="Click to edit" relationship="description">
        <span className={styles.editableValue} onClick={() => setEditing(true)}>
          {type === 'currency' ? formatCurrency(value) : value || '—'}{' '}
          <Edit24Regular style={{ fontSize: 14, verticalAlign: 'middle', opacity: 0.5 }} />
        </span>
      </Tooltip>
    );
  }

  if (multiline) {
    return (
      <div>
        <Textarea
          value={draft}
          onChange={(_e, d) => setDraft(d.value)}
          resize="vertical"
          style={{ width: '100%' }}
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
  }

  return (
    <div className={styles.editRow}>
      <Input
        className={styles.editField}
        value={draft}
        type={type === 'currency' || type === 'number' ? 'number' : 'text'}
        onChange={(_e, d) => setDraft(d.value)}
        autoFocus
        onKeyDown={(e) => {
          if (e.key === 'Enter') commit();
          if (e.key === 'Escape') cancel();
        }}
      />
      <Button appearance="primary" size="small" icon={<Save24Regular />} onClick={commit} />
      <Button appearance="subtle" size="small" icon={<Dismiss24Regular />} onClick={cancel} />
    </div>
  );
};

// ─── EditableList ───
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
              <Body1>• {item}</Body1>
            </div>
          ))}
        </div>
        <div style={{ marginTop: tokens.spacingVerticalS }}>
          <Tooltip content="Click to edit" relationship="description">
            <Button
              appearance="subtle"
              size="small"
              icon={<Edit24Regular />}
              onClick={() => setEditing(true)}
            >
              Edit {label}
            </Button>
          </Tooltip>
        </div>
      </div>
    );
  }

  return (
    <div>
      <Textarea
        value={draft}
        onChange={(_e, d) => setDraft(d.value)}
        resize="vertical"
        placeholder={`One ${label.toLowerCase().replace(/s$/, '')} per line`}
        style={{ width: '100%', minHeight: '80px' }}
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
        { label: 'Total Claims', value: total, color: tokens.colorBrandForeground1 },
        { label: 'Pending', value: pending, color: tokens.colorPaletteYellowForeground1 },
        { label: 'Under Review', value: investigation, color: tokens.colorBrandForeground1 },
        {
          label: 'Total Value',
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
            {claim.policyHolderName} &bull; {claim.policyHolderEmail}
          </Caption1>
        </div>
        <Badge appearance="filled" color={statusBadgeColor(claim.status)}>
          {getShortStatus(claim.status)}
        </Badge>
      </div>
      <div className={styles.detailGrid}>
        <div>
          <Caption1 style={{ color: tokens.colorNeutralForeground3 }}>Property</Caption1>
          <Body1 style={{ display: 'block' }}>{claim.property}</Body1>
        </div>
        <div>
          <Caption1 style={{ color: tokens.colorNeutralForeground3 }}>Date of Loss</Caption1>
          <Body1 style={{ display: 'block' }}>{formatDate(claim.dateOfLoss)}</Body1>
        </div>
        <div>
          <Caption1 style={{ color: tokens.colorNeutralForeground3 }}>Estimated Loss</Caption1>
          <Body1
            style={{
              display: 'block',
              color: tokens.colorPaletteGreenForeground1,
              fontWeight: '700',
            }}
          >
            {formatCurrency(claim.estimatedLoss)}
          </Body1>
        </div>
      </div>
      {claim.damageTypes && claim.damageTypes.length > 0 && (
        <div className={styles.damageTags}>
          {claim.damageTypes.map((d: string, i: number) => (
            <Badge key={i} appearance="outline">
              {d}
            </Badge>
          ))}
        </div>
      )}
    </Card>
  );
};

// ─── ClaimDetailView ───
const ClaimDetailView: React.FC<{ claim: any; onBack: () => void }> = ({
  claim: initialClaim,
  onBack,
}) => {
  const styles = useStyles();
  const [claim, setClaim] = useState(initialClaim);
  const [saving, setSaving] = useState(false);
  const [saveMsg, setSaveMsg] = useState<{ type: string; text: string } | null>(null);
  const [activeTab, setActiveTab] = useState('overview');

  const updateField = async (field: string, value: any) => {
    setClaim((prev: any) => ({ ...prev, [field]: value, updatedAt: new Date().toISOString() }));
    setSaving(true);
    setSaveMsg(null);

    try {
      const baseUrl = window.location.origin;
      const res = await fetch(`${baseUrl}/mcp`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          jsonrpc: '2.0',
          id: Date.now(),
          method: 'tools/call',
          params: {
            name: 'update_claim',
            arguments: { claimId: claim.id, [field]: value },
          },
        }),
      });
      const json = await res.json();
      if (json.result?.structuredContent?.success) {
        setClaim(json.result.structuredContent.data);
        setSaveMsg({ type: 'success', text: `${field} updated` });
      } else if (json.result?.content?.[0]?.text) {
        try {
          const parsed = JSON.parse(json.result.content[0].text);
          if (parsed.success && parsed.data) {
            setClaim(parsed.data);
            setSaveMsg({ type: 'success', text: `${field} updated` });
          } else {
            throw new Error('Update failed');
          }
        } catch {
          setSaveMsg({ type: 'success', text: `${field} updated locally` });
        }
      } else {
        setSaveMsg({ type: 'success', text: `${field} updated locally` });
      }
    } catch (err) {
      console.error('Update failed:', err);
      setSaveMsg({ type: 'warning', text: `${field} saved locally (sync pending)` });
    } finally {
      setSaving(false);
      setTimeout(() => setSaveMsg(null), 3000);
    }
  };

  return (
    <div>
      <Button
        appearance="subtle"
        icon={<ArrowLeft24Regular />}
        onClick={onBack}
        style={{ marginBottom: tokens.spacingVerticalM }}
      >
        Back to Claims
      </Button>

      {saveMsg && (
        <div style={{ marginBottom: tokens.spacingVerticalM }}>
          <MessageBar intent={saveMsg.type === 'success' ? 'success' : 'warning'}>
            <MessageBarBody>{saveMsg.text}</MessageBarBody>
          </MessageBar>
        </div>
      )}

      <Card style={{ overflow: 'hidden' }}>
        {/* Header */}
        <div className={styles.detailHeader}>
          <Title3 style={{ color: tokens.colorNeutralForegroundOnBrand }}>
            {claim.claimNumber}
          </Title3>
          <Body1
            style={{
              color: 'rgba(255,255,255,0.9)',
              display: 'block',
              marginTop: '4px',
            }}
          >
            {claim.policyHolderName}
          </Body1>
          <div style={{ marginTop: tokens.spacingVerticalS }}>
            <Badge appearance="filled" color={statusBadgeColor(claim.status)} size="large">
              {getShortStatus(claim.status)}
            </Badge>
            {saving && <Spinner size="tiny" style={{ marginLeft: 8 }} />}
          </div>
        </div>

        {/* Tabs */}
        <div style={{ borderBottom: `1px solid ${tokens.colorNeutralStroke2}` }}>
          <TabList
            selectedValue={activeTab}
            onTabSelect={(_: any, d: any) => setActiveTab(d.value)}
          >
            <Tab value="overview">Overview</Tab>
            <Tab value="property">Property & Map</Tab>
            <Tab value="details">Details & Notes</Tab>
          </TabList>
        </div>

        <div className={styles.detailBody}>
          {/* Overview tab */}
          {activeTab === 'overview' && (
            <>
              <div className={styles.section}>
                <div className={styles.sectionTitle}>
                  <DocumentText24Regular />
                  <Subtitle2>Claim Overview</Subtitle2>
                </div>
                <div className={styles.infoGrid}>
                  <div className={styles.infoItem}>
                    <div className={styles.infoLabel}>Claim ID</div>
                    <div className={styles.infoValue}>{claim.id}</div>
                  </div>
                  <div className={styles.infoItem}>
                    <div className={styles.infoLabel}>Policy Number</div>
                    <div className={styles.infoValue}>{claim.policyNumber}</div>
                  </div>
                  <div className={styles.infoItem}>
                    <div className={styles.infoLabel}>Estimated Loss</div>
                    <div className={styles.infoValueLarge}>
                      <EditableField
                        value={claim.estimatedLoss}
                        type="currency"
                        onSave={(v) => updateField('estimatedLoss', v)}
                      />
                    </div>
                  </div>
                  <div className={styles.infoItem}>
                    <div className={styles.infoLabel}>Adjuster Assigned</div>
                    <div className={styles.infoValue}>
                      <EditableField
                        value={claim.adjusterAssigned || 'Unassigned'}
                        onSave={(v) => updateField('adjusterAssigned', v)}
                      />
                    </div>
                  </div>
                </div>
              </div>

              <div className={styles.section}>
                <div className={styles.sectionTitle}>
                  <Warning24Regular />
                  <Subtitle2>Status</Subtitle2>
                </div>
                <div className={styles.descriptionBox}>
                  <EditableField
                    value={claim.status}
                    onSave={(v) => updateField('status', v)}
                  />
                </div>
              </div>

              <div className={styles.section}>
                <div className={styles.sectionTitle}>
                  <Clock24Regular />
                  <Subtitle2>Timeline</Subtitle2>
                </div>
                <div className={styles.infoGrid}>
                  <div className={styles.infoItem}>
                    <div className={styles.infoLabel}>Date of Loss</div>
                    <div className={styles.infoValue}>{formatDate(claim.dateOfLoss)}</div>
                  </div>
                  <div className={styles.infoItem}>
                    <div className={styles.infoLabel}>Date Reported</div>
                    <div className={styles.infoValue}>{formatDate(claim.dateReported)}</div>
                  </div>
                  <div className={styles.infoItem}>
                    <div className={styles.infoLabel}>Created</div>
                    <div className={styles.infoValue}>{formatDate(claim.createdAt)}</div>
                  </div>
                  <div className={styles.infoItem}>
                    <div className={styles.infoLabel}>Last Updated</div>
                    <div className={styles.infoValue}>{formatDate(claim.updatedAt)}</div>
                  </div>
                </div>
              </div>
            </>
          )}

          {/* Property & Map tab */}
          {activeTab === 'property' && (
            <div className={styles.section}>
              <div className={styles.sectionTitle}>
                <Location24Regular />
                <Subtitle2>Property Location</Subtitle2>
              </div>
              <div className={styles.infoGrid}>
                <div className={styles.infoItem}>
                  <div className={styles.infoLabel}>Property Address</div>
                  <div className={styles.infoValue}>{claim.property}</div>
                </div>
                <div className={styles.infoItem}>
                  <div className={styles.infoLabel}>Policy Holder</div>
                  <div className={styles.infoValue}>{claim.policyHolderName}</div>
                </div>
                <div className={styles.infoItem}>
                  <div className={styles.infoLabel}>Email</div>
                  <div className={styles.infoValue}>{claim.policyHolderEmail}</div>
                </div>
              </div>
              <PropertyMap address={claim.property} />
            </div>
          )}

          {/* Details & Notes tab */}
          {activeTab === 'details' && (
            <>
              <div className={styles.section}>
                <div className={styles.sectionTitle}>
                  <DocumentText24Regular />
                  <Subtitle2>Description</Subtitle2>
                </div>
                <EditableField
                  value={claim.description}
                  multiline
                  onSave={(v) => updateField('description', v)}
                />
              </div>

              {claim.damageTypes && claim.damageTypes.length > 0 && (
                <div className={styles.section}>
                  <div className={styles.sectionTitle}>
                    <Warning24Regular />
                    <Subtitle2>Damage Types</Subtitle2>
                  </div>
                  <EditableList
                    items={claim.damageTypes}
                    label="Damage Types"
                    onSave={(v) => updateField('damageTypes', v)}
                  />
                </div>
              )}

              <div className={styles.section}>
                <div className={styles.sectionTitle}>
                  <Note24Regular />
                  <Subtitle2>Notes</Subtitle2>
                </div>
                <EditableList
                  items={claim.notes}
                  label="Notes"
                  onSave={(v) => updateField('notes', v)}
                />
              </div>
            </>
          )}
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
                    Viewing {selected.claimNumber}
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
