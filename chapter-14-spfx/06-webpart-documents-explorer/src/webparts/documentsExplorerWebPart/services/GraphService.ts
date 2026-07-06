import { MSGraphClientV3 } from '@microsoft/sp-http';
import { IDrive, IDriveItem } from '../models/IDrive';

interface IRawDrive {
  id: string;
  name?: string;
  description?: string;
  webUrl: string;
  driveType?: string;
}
interface IRawDriveItem {
  id: string;
  name?: string;
  webUrl?: string;
  folder?: Record<string, unknown>;
  size?: number;
  lastModifiedDateTime?: string;
}
interface ICollection<T> {
  value?: T[];
}

/**
 * Graph wrapper that lists a site's document libraries (drives) and the root
 * children of a drive. Connects to the book's cap4 (Documents).
 */
export class GraphService {
  constructor(private readonly client: MSGraphClientV3) {}

  public async getSiteDrives(siteId: string): Promise<IDrive[]> {
    const response = await this.client
      .api(`/sites/${siteId}/drives`)
      .select('id,name,description,webUrl,driveType')
      .get() as ICollection<IRawDrive>;

    const value: IRawDrive[] = response.value ?? [];
    return value.map((d: IRawDrive): IDrive => ({
      id: d.id,
      name: d.name ?? '',
      description: d.description,
      webUrl: d.webUrl,
      driveType: d.driveType
    }));
  }

  public async getDriveRootItems(driveId: string, limit: number = 10): Promise<IDriveItem[]> {
    const response = await this.client
      .api(`/drives/${driveId}/root/children`)
      .top(limit)
      .select('id,name,webUrl,folder,size,lastModifiedDateTime')
      .get() as ICollection<IRawDriveItem>;

    const value: IRawDriveItem[] = response.value ?? [];
    return value.map((i: IRawDriveItem): IDriveItem => ({
      id: i.id,
      name: i.name ?? '',
      webUrl: i.webUrl ?? '',
      isFolder: !!i.folder,
      size: i.size,
      lastModifiedDateTime: i.lastModifiedDateTime
    }));
  }
}