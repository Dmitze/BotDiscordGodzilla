import ClusterManager from '@/utils/clusterManager';

describe('ClusterManager messaging', () => {
  it('ignores malformed messages', () => {
    const cm: any = new ClusterManager({ workers: 0 });
    const fakeWorker = { id: 1 } as any;
    expect(() => cm['handleWorkerMessage'](fakeWorker, 'bad')).not.toThrow();
  });
});


