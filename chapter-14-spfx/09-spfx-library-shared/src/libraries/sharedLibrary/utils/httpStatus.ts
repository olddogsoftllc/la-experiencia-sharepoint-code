/**
 * Categories for an HTTP error from Graph / a custom API.
 * Mirrors the classification shown in the book's chapter 14 "Error handling and telemetry".
 */
export type HttpStatusCategory =
  | 'network'
  | 'unauthorized'
  | 'forbidden'
  | 'notFound'
  | 'throttled'
  | 'server'
  | 'unknown';

/**
 * Map an HTTP status code to a coarse error category that the UI can branch on.
 * `undefined` means no status was returned (network/timeout).
 */
export function classifyHttpStatus(statusCode: number | undefined): HttpStatusCategory {
  if (statusCode === 401) { return 'unauthorized'; }
  if (statusCode === 403) { return 'forbidden'; }
  if (statusCode === 404) { return 'notFound'; }
  if (statusCode === 429 || statusCode === 503) { return 'throttled'; }
  if (statusCode === undefined) { return 'network'; }
  if (statusCode >= 500) { return 'server'; }
  return 'unknown';
}