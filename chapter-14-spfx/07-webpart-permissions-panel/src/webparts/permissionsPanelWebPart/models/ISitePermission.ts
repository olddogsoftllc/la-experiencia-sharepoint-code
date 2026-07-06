// Model for a permission entry on a Graph site (/sites/{id}/permissions).
export interface ISitePermission {
  id: string;
  roles: string[];
  grantedTo: string;
  isInherited: boolean;
}