import { classifyHttpStatus } from './httpStatus';

describe('classifyHttpStatus', () => {
  it('maps 401 to unauthorized', () => {
    expect(classifyHttpStatus(401)).toBe('unauthorized');
  });

  it('maps 403 to forbidden', () => {
    expect(classifyHttpStatus(403)).toBe('forbidden');
  });

  it('maps 404 to notFound', () => {
    expect(classifyHttpStatus(404)).toBe('notFound');
  });

  it('maps 429 and 503 to throttled', () => {
    expect(classifyHttpStatus(429)).toBe('throttled');
    expect(classifyHttpStatus(503)).toBe('throttled');
  });

  it('maps 5xx (other than 503) to server', () => {
    expect(classifyHttpStatus(500)).toBe('server');
    expect(classifyHttpStatus(502)).toBe('server');
  });

  it('maps undefined to network', () => {
    expect(classifyHttpStatus(undefined)).toBe('network');
  });

  it('maps 2xx/3xx to unknown', () => {
    expect(classifyHttpStatus(200)).toBe('unknown');
    expect(classifyHttpStatus(302)).toBe('unknown');
  });
});