// Test unitario (sin red): getGraphClient valida las variables de entorno.
import { getGraphClient } from './graphAuth';

describe('getGraphClient (factory Graph app-only)', () => {
  const OLD_ENV = process.env;

  beforeEach(() => {
    // Limpia el cache del cliente lazy entre tests
    jest.resetModules();
    process.env = { ...OLD_ENV };
    delete process.env.GRAPH_TENANT_ID;
    delete process.env.GRAPH_CLIENT_ID;
    delete process.env.GRAPH_CLIENT_SECRET;
  });

  afterAll(() => {
    process.env = OLD_ENV;
  });

  test('lanza si falta GRAPH_TENANT_ID', () => {
    process.env.GRAPH_CLIENT_ID = 'cid';
    process.env.GRAPH_CLIENT_SECRET = 'secret';
    expect(() => getGraphClient()).toThrow(/GRAPH_TENANT_ID/);
  });

  test('lanza si falta GRAPH_CLIENT_ID', () => {
    process.env.GRAPH_TENANT_ID = 'tid';
    process.env.GRAPH_CLIENT_SECRET = 'secret';
    expect(() => getGraphClient()).toThrow(/GRAPH_CLIENT_ID/);
  });

  test('lanza si falta GRAPH_CLIENT_SECRET', () => {
    process.env.GRAPH_TENANT_ID = 'tid';
    process.env.GRAPH_CLIENT_ID = 'cid';
    expect(() => getGraphClient()).toThrow(/GRAPH_CLIENT_SECRET/);
  });

  test('menciona .env.example en el mensaje de error', () => {
    expect(() => getGraphClient()).toThrow(/\.env\.example/);
  });
});