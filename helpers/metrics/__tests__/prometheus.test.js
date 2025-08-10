const MetricsCollector = require('../../../data/metrics/prometheus');

describe('Prometheus hardening', () => {
  it('does not throw on counters/observe and CPU/disk gauges update', () => {
    const m = new MetricsCollector();
    expect(() => m.recordCommand('t','ok','u1')).not.toThrow();
    expect(() => m.recordCommandDuration('t', 0.12)).not.toThrow();
    expect(() => m.recordApiRequest('google','GET','200')).not.toThrow();
    expect(() => m.recordApiRequestDuration('google', 0.5)).not.toThrow();
    expect(() => m.recordCacheHit('search')).not.toThrow();
    expect(() => m.recordCacheMiss('search')).not.toThrow();
    expect(() => m.recordError('general','cmd')).not.toThrow();
    expect(() => m.updateCpuAndDisk()).not.toThrow();
  });
});

