// Minimal model for a SharePoint site returned by Graph /sites?search=
// Keep the surface small: only the fields the Web Part renders.
export interface ISite {
  id: string;
  displayName: string;
  name?: string;
  url: string;
  description?: string;
}