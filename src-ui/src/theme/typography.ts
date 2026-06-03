export const fontStack = [
  'Inter',
  '"Microsoft YaHei"',
  '"PingFang SC"',
  '"Noto Sans SC"',
  'system-ui',
  '-apple-system',
  'BlinkMacSystemFont',
  'sans-serif',
].join(', ');

export const typography = {
  display: {
    fontSize: 'var(--text-display)',
    lineHeight: 'var(--lh-display)',
    fontWeight: '600',
  },
  pageTitle: {
    fontSize: 'var(--text-page-title)',
    lineHeight: 'var(--lh-page-title)',
    fontWeight: '600',
  },
  sectionTitle: {
    fontSize: 'var(--text-section-title)',
    lineHeight: 'var(--lh-section-title)',
    fontWeight: '600',
  },
  body: {
    fontSize: 'var(--text-body)',
    lineHeight: 'var(--lh-body)',
    fontWeight: '400',
  },
  bodyStrong: {
    fontSize: 'var(--text-body-strong)',
    lineHeight: 'var(--lh-body-strong)',
    fontWeight: '600',
  },
  label: {
    fontSize: 'var(--text-label)',
    lineHeight: 'var(--lh-label)',
    fontWeight: '500',
  },
  caption: {
    fontSize: 'var(--text-caption)',
    lineHeight: 'var(--lh-caption)',
    fontWeight: '400',
  },
  metric: {
    fontSize: 'var(--text-metric)',
    lineHeight: 'var(--lh-metric)',
    fontWeight: '600',
  },
};
