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
  makeStyles,
  tokens,
  shorthands,
  Title3,
  Subtitle1,
  Body1,
  Caption1,
  MessageBar,
  MessageBarBody,
} from '@fluentui/react-components';
import { ArrowLeft24Regular } from '@fluentui/react-icons';
import { useToolResult, useHostTheme } from './hooks/useMcpBridge';

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
  detailView: {
    overflow: 'hidden',
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
    fontSize: '0.875rem',
    fontWeight: '600',
    color: tokens.colorNeutralForeground3,
    textTransform: 'uppercase' as const,
    letterSpacing: '0.5px',
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
      {claim.damageTypes && (
        <div className={styles.damageTags}>
          {claim.damageTypes.split(',').map((d: string, i: number) => (
            <Badge key={i} appearance="outline">
              {d.trim()}
            </Badge>
          ))}
        </div>
      )}
    </Card>
  );
};

// ─── ClaimDetailView ───
const ClaimDetailView: React.FC<{ claim: any; onBack: () => void }> = ({ claim, onBack }) => {
  const styles = useStyles();

  return (
    <div>
      <Button
        appearance="subtle"
        icon={<ArrowLeft24Regular />}
        onClick={onBack}
        style={{ marginBottom: tokens.spacingVerticalM }}
      >
        Back to Claims List
      </Button>

      <Card className={styles.detailView}>
        {/* Header */}
        <div className={styles.detailHeader}>
          <Title3 style={{ color: tokens.colorNeutralForegroundOnBrand }}>
            {claim.claimNumber}
          </Title3>
          <Body1
            style={{ color: 'rgba(255,255,255,0.9)', display: 'block', marginTop: '4px' }}
          >
            {claim.policyHolderName}
          </Body1>
          <div style={{ marginTop: tokens.spacingVerticalS }}>
            <Badge appearance="filled" color={statusBadgeColor(claim.status)} size="large">
              {getShortStatus(claim.status)}
            </Badge>
          </div>
        </div>

        <div className={styles.detailBody}>
          {/* Claim Overview */}
          <div className={styles.section}>
            <div className={styles.sectionTitle}>Claim Overview</div>
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
                <div className={styles.infoValueLarge}>{formatCurrency(claim.estimatedLoss)}</div>
              </div>
              <div className={styles.infoItem}>
                <div className={styles.infoLabel}>Adjuster Assigned</div>
                <div className={styles.infoValue}>
                  {claim.adjusterAssigned || 'Unassigned'}
                </div>
              </div>
            </div>
          </div>

          {/* Property & Contact */}
          <div className={styles.section}>
            <div className={styles.sectionTitle}>Property &amp; Contact</div>
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
          </div>

          {/* Timeline */}
          <div className={styles.section}>
            <div className={styles.sectionTitle}>Timeline</div>
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

          {/* Damage Types */}
          {claim.damageTypes && (
            <div className={styles.section}>
              <div className={styles.sectionTitle}>Damage Types</div>
              <div className={styles.damageTags} style={{ marginTop: 0 }}>
                {claim.damageTypes.split(',').map((d: string, i: number) => (
                  <Badge key={i} appearance="outline">
                    {d.trim()}
                  </Badge>
                ))}
              </div>
            </div>
          )}

          {/* Description */}
          <div className={styles.section}>
            <div className={styles.sectionTitle}>Description</div>
            <div className={styles.descriptionBox}>
              <Body1>{claim.description || 'No description provided.'}</Body1>
            </div>
          </div>

          {/* Status Details */}
          <div className={styles.section}>
            <div className={styles.sectionTitle}>Status Details</div>
            <div className={styles.descriptionBox}>
              <Body1>{claim.status || 'Status not available.'}</Body1>
            </div>
          </div>

          {/* Notes */}
          {claim.notes && (
            <div className={styles.section}>
              <div className={styles.sectionTitle}>Notes</div>
              <div className={styles.descriptionBox}>
                <Body1>{claim.notes}</Body1>
              </div>
            </div>
          )}
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
  const [items, setItems] = useState<any[]>([]);
  const [loading, setLoading] = useState(true);
  const [selected, setSelected] = useState<any | null>(null);

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
            <Spinner label="Loading claims data..." />
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
                  <Title3>Zava Insurance Claims</Title3>
                  <Caption1 style={{ display: 'block', color: tokens.colorNeutralForeground3 }}>
                    Viewing claim {selected.claimNumber}
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
                  <Title3>Zava Insurance Claims</Title3>
                  <Caption1 style={{ display: 'block', color: tokens.colorNeutralForeground3 }}>
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
