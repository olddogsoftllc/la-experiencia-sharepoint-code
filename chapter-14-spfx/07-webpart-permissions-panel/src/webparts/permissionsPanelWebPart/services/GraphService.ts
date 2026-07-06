import { MSGraphClientV3 } from '@microsoft/sp-http';
import { ISitePermission } from '../models/ISitePermission';
import { TelemetryClient } from './TelemetryClient';

interface IRawIdentitySet {
  user?: { displayName?: string };
  group?: { displayName?: string };
  application?: { displayName?: string };
}
interface IRawPermission {
  id: string;
  roles?: string[];
  grantedTo?: IRawIdentitySet;
  grantedToIdentities?: IRawIdentitySet[];
  isInherited?: boolean;
}
interface ICollection {
  value?: IRawPermission[];
}

/**
 * Graph wrapper that reads a site's permission entries and classifies errors
 * (network / 401 / 403 / 404 / 429) with telemetry. Connects to cap5.
 *
 * NOTE: `/sites/{id}/permissions` is an admin operation. Under Sites.Selected
 * it typically returns 403 — the component surfaces that explicitly.
 */
export class GraphService {
  constructor(
    private readonly client: MSGraphClientV3,
    private readonly telemetry: TelemetryClient
  ) {}

  public async getSitePermissions(siteId: string): Promise<ISitePermission[]> {
    try {
      const response = await this.client
        .api(`/sites/${siteId}/permissions`)
        .get() as ICollection;

      const value: IRawPermission[] = response.value ?? [];
      this.telemetry.trackSuccess('getSitePermissions', { siteId, count: value.length });

      return value.map((p: IRawPermission): ISitePermission => ({
        id: p.id,
        roles: p.roles ?? [],
        grantedTo: this.formatGrantees(p),
        isInherited: !!p.isInherited
      }));
    } catch (err) {
      const e = err as { statusCode?: number; message?: string };
      this.telemetry.trackError('getSitePermissions', siteId, e);
      throw err;
    }
  }

  private formatGrantees(p: IRawPermission): string {
    const names: string[] = [];

    if (p.grantedTo?.user?.displayName) {
      names.push(p.grantedTo.user.displayName);
    }
    if (p.grantedTo?.group?.displayName) {
      names.push(p.grantedTo.group.displayName);
    }

    if (p.grantedToIdentities) {
      for (const id of p.grantedToIdentities) {
        if (id.user?.displayName) {
          names.push(id.user.displayName);
        }
        if (id.group?.displayName) {
          names.push(id.group.displayName);
        }
      }
    }

    return names.length ? names.join(', ') : '(unknown)';
  }
}