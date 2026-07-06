import { MSGraphClientV3 } from '@microsoft/sp-http';
import { ISite } from '../models/ISite';

// Raw shape returned by GET /sites?search= (only the fields we select).
interface IRawSite {
  id: string;
  displayName?: string;
  name?: string;
  webUrl: string;
  description?: string;
}
interface ISiteSearchResponse {
  value?: IRawSite[];
}

/**
 * Thin wrapper around MSGraphClientV3.
 *
 * Injected into the React component so the component never touches
 * `this.context` and stays unit-testable with a fake client.
 *
 * Mirrors the GraphService pattern shown in the book (cap14,
 * "Patrón de Acceso a Graph en SPFx").
 */
export class GraphService {
  constructor(private readonly client: MSGraphClientV3) {}

  /**
   * Search SharePoint sites by keyword via Graph.
   * @param query  free-text search (maps to $search)
   * @param limit  max number of results ($top)
   */
  public async searchSites(query: string, limit: number = 10): Promise<ISite[]> {
    if (!query) {
      return [];
    }

    const response = await this.client
      .api(`/sites?search=${encodeURIComponent(query)}`)
      .top(limit)
      .select('id,displayName,name,webUrl,description')
      .get() as ISiteSearchResponse;

    const value: IRawSite[] = response.value ?? [];
    return value.map((s: IRawSite): ISite => ({
      id: s.id,
      displayName: s.displayName ?? '',
      name: s.name,
      url: s.webUrl,
      description: s.description
    }));
  }
}