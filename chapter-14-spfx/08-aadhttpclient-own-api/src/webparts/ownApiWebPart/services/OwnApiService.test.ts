import 'jest';
// Mock @microsoft/sp-http so the test never loads the SPFx runtime; the SUT only
// needs AadHttpClient.configurations.v1 to pass it as the second arg of client.get.
jest.mock('@microsoft/sp-http', () => ({
  AadHttpClient: { configurations: { v1: { __mock: true } } }
}));

import { OwnApiService } from './OwnApiService';
import { AadHttpClient } from '@microsoft/sp-http';

interface FakeResponse {
  ok: boolean;
  status: number;
  statusText: string;
  json: () => Promise<unknown>;
}

function makeFakeClient(response: FakeResponse): AadHttpClient {
  const calls: Array<{ url: string; config: unknown }> = [];
  return {
    get: async (url: string, config: unknown) => {
      calls.push({ url, config });
      return response;
    }
    // expose calls for assertions via the returned object cast
  } as unknown as AadHttpClient & { __calls: typeof calls };
}

describe('OwnApiService.getData', () => {
  it('returns the parsed JSON when the API responds ok', async () => {
    const fake = makeFakeClient({
      ok: true, status: 200, statusText: 'OK',
      json: async () => ({ hello: 'world' })
    });
    const svc = new OwnApiService(fake, 'https://api.example.com/data');
    await expect(svc.getData()).resolves.toEqual({ hello: 'world' });
  });

  it('throws "API returned <status> <statusText>" when the API responds not ok', async () => {
    const fake = makeFakeClient({
      ok: false, status: 500, statusText: 'Internal Server Error',
      json: async () => ({})
    });
    const svc = new OwnApiService(fake, 'https://api.example.com/data');
    await expect(svc.getData()).rejects.toThrow('API returned 500 Internal Server Error');
  });

  it('calls the configured apiUrl with AadHttpClient.configurations.v1', async () => {
    const fake = makeFakeClient({
      ok: true, status: 200, statusText: 'OK', json: async () => ({})
    });
    const svc = new OwnApiService(fake, 'https://api.example.com/data');
    await svc.getData();
    // The fake recorded the call; assert the URL via a second spy to keep it simple.
    const spy = jest.spyOn(fake as unknown as { get: (u: string, c: unknown) => Promise<unknown> }, 'get');
    await svc.getData();
    expect(spy).toHaveBeenCalledWith('https://api.example.com/data', AadHttpClient.configurations.v1);
  });
});