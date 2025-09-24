import SchedulerService from '@/services/SchedulerService';

describe('SchedulerService', () => {
  const makeBot = () => ({ getService: () => undefined, client: undefined });

  it('isCronDisabled respects env', () => {
    const svc: any = new SchedulerService(makeBot() as any);
    const prev = process.env['DISABLE_CRON'];
    process.env['DISABLE_CRON'] = 'true';
    // @ts-expect-error private method access via cast
    expect(svc['isCronDisabled']()).toBe(true);
    if (prev === undefined) delete process.env['DISABLE_CRON']; else process.env['DISABLE_CRON'] = prev;
  });

  it('scheduleJob returns null when cron disabled', async () => {
    const svc: any = new SchedulerService(makeBot() as any);
    jest.spyOn(svc, 'isCronDisabled').mockReturnValue(true);
    const job = svc.scheduleJob('t', '* * * * *', () => undefined);
    expect(job).toBeNull();
  });
});


