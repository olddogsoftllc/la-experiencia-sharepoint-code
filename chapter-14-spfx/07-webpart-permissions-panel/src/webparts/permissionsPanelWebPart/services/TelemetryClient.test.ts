import 'jest';
import { TelemetryClient } from './TelemetryClient';

describe('TelemetryClient', () => {
  it('trackSuccess records an ".ok" event with the given props', () => {
    const tc = new TelemetryClient();
    tc.trackSuccess('ListSites', { count: 7 });
    expect(tc.events).toHaveLength(1);
    expect(tc.events[0]).toEqual({ op: 'ListSites.ok', props: { count: 7 } });
  });

  it('trackError records an ".error" event with ctx/code/message', () => {
    const tc = new TelemetryClient();
    tc.trackError('ListSites', 'graph', { statusCode: 403, message: 'forbidden' });
    expect(tc.events).toHaveLength(1);
    expect(tc.events[0].op).toBe('ListSites.error');
    expect(tc.events[0].props).toEqual({ ctx: 'graph', code: 403, message: 'forbidden' });
  });

  it('records multiple events in insertion order', () => {
    const tc = new TelemetryClient();
    tc.trackSuccess('A', {});
    tc.trackError('B', 'ctx', { message: 'boom' });
    tc.trackSuccess('C', {});
    expect(tc.events.map((e) => e.op)).toEqual(['A.ok', 'B.error', 'C.ok']);
  });

  it('accepts an instrumentation key without changing in-memory behavior', () => {
    const tc = new TelemetryClient('00000000-0000-0000-0000-000000000000');
    tc.trackSuccess('Op', { ok: true });
    expect(tc.events).toHaveLength(1);
    expect(tc.events[0].op).toBe('Op.ok');
  });
});