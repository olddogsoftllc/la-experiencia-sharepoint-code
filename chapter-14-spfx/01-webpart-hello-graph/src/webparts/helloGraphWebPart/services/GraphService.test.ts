import 'jest';
import { GraphService } from './GraphService';
import { MSGraphClientV3 } from '@microsoft/sp-http';

// Builds a fake MSGraphClientV3 whose .api().top().select().get() chain resolves
// to { value }. Every chain method returns the chain so order doesn't matter.
function makeFakeClient(value: unknown[]): MSGraphClientV3 {
  const chain: Record<string, unknown> = {
    api: () => chain,
    top: () => chain,
    select: () => chain,
    get: async () => ({ value })
  };
  return chain as unknown as MSGraphClientV3;
}

describe('GraphService.searchSites', () => {
  it('returns [] for an empty query (no Graph call)', async () => {
    const svc = new GraphService(makeFakeClient([]));
    expect(await svc.searchSites('', 10)).toEqual([]);
  });

  it('maps raw Graph sites to ISite with displayName fallback', async () => {
    const raw = [
      { id: 'a', displayName: 'Book Test', name: 'book-test', webUrl: 'https://x/sites/bt', description: 'desc' },
      { id: 'b', displayName: '', name: 'root', webUrl: 'https://x', description: undefined }
    ];
    const svc = new GraphService(makeFakeClient(raw));
    const sites = await svc.searchSites('book', 10);
    expect(sites).toHaveLength(2);
    expect(sites[0]).toEqual({ id: 'a', displayName: 'Book Test', name: 'book-test', url: 'https://x/sites/bt', description: 'desc' });
    // displayName falls back to '' when absent
    expect(sites[1].displayName).toBe('');
  });

  it('returns [] when Graph returns no value', async () => {
    const svc = new GraphService(makeFakeClient(undefined as unknown as never[]));
    expect(await svc.searchSites('zzz', 5)).toEqual([]);
  });
});