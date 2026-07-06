import 'jest';
import { GraphService } from './GraphService';
import { MSGraphClientV3 } from '@microsoft/sp-http';

// Fake MSGraphClientV3 whose .api().select().top().get() chain resolves to { value }.
function makeFakeClient(value: unknown): MSGraphClientV3 {
  const chain: Record<string, unknown> = {
    api: () => chain,
    select: () => chain,
    top: () => chain,
    get: async () => ({ value })
  };
  return chain as unknown as MSGraphClientV3;
}

describe('GraphService.getSiteDrives', () => {
  it('maps raw drives to IDrive with name fallback to ""', async () => {
    const raw = [
      { id: 'd1', name: 'Documents', description: 'desc', webUrl: 'https://x/drives/d1', driveType: 'documentLibrary' },
      { id: 'd2', webUrl: 'https://x/drives/d2' }
    ];
    const svc = new GraphService(makeFakeClient(raw));
    const drives = await svc.getSiteDrives('siteId');
    expect(drives).toHaveLength(2);
    expect(drives[0]).toEqual({
      id: 'd1', name: 'Documents', description: 'desc',
      webUrl: 'https://x/drives/d1', driveType: 'documentLibrary'
    });
    expect(drives[1].name).toBe('');
    expect(drives[1].driveType).toBeUndefined();
  });

  it('returns [] when Graph returns no value', async () => {
    const svc = new GraphService(makeFakeClient(undefined));
    expect(await svc.getSiteDrives('siteId')).toEqual([]);
  });
});

describe('GraphService.getDriveRootItems', () => {
  it('maps raw items, deriving isFolder from the folder child object', async () => {
    const raw = [
      { id: 'i1', name: 'file.txt', webUrl: 'https://x/file.txt', size: 1024 },
      { id: 'i2', name: 'folder', folder: { childCount: 3 }, size: 0 }
    ];
    const svc = new GraphService(makeFakeClient(raw));
    const items = await svc.getDriveRootItems('driveId', 10);
    expect(items).toHaveLength(2);
    expect(items[0]).toEqual({
      id: 'i1', name: 'file.txt', webUrl: 'https://x/file.txt',
      isFolder: false, size: 1024, lastModifiedDateTime: undefined
    });
    expect(items[1].isFolder).toBe(true);
    expect(items[1].name).toBe('folder');
  });

  it('returns [] when Graph returns no value', async () => {
    const svc = new GraphService(makeFakeClient(undefined));
    expect(await svc.getDriveRootItems('driveId', 5)).toEqual([]);
  });
});