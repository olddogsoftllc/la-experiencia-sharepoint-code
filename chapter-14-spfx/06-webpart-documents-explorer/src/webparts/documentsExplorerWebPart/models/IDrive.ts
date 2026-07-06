// Models for Graph drive (document library) and drive item (file/folder).
export interface IDrive {
  id: string;
  name: string;
  description?: string;
  webUrl: string;
  driveType?: string;
}

export interface IDriveItem {
  id: string;
  name: string;
  webUrl: string;
  isFolder: boolean;
  size?: number;
  lastModifiedDateTime?: string;
}