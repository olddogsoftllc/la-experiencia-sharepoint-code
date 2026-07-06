export interface ITelemetryEvent {
  op: string;
  props: Record<string, unknown>;
}

/**
 * Minimal telemetry client stub. Records events in memory for inspection.
 * In production, wire `trackSuccess`/`trackError` to Application Insights
 * (`@microsoft/applicationinsights-web`) when an instrumentation key is set —
 * see the book's chapter 14 "Error handling and telemetry".
 */
export class TelemetryClient {
  private readonly _instrumentationKey: string | undefined;
  private readonly _events: ITelemetryEvent[] = [];

  constructor(instrumentationKey?: string) {
    this._instrumentationKey = instrumentationKey;
  }

  public trackSuccess(op: string, props: Record<string, unknown>): void {
    this._events.push({ op: `${op}.ok`, props });
    if (this._instrumentationKey) {
      // Production: this.ai.trackEvent({ name: `spfx.${op}.ok`, properties: props });
    }
  }

  public trackError(
    op: string,
    ctx: string,
    err: { statusCode?: number; message?: string }
  ): void {
    this._events.push({
      op: `${op}.error`,
      props: { ctx, code: err.statusCode, message: err.message }
    });
    if (this._instrumentationKey) {
      // Production: this.ai.trackException({ exception: err, properties: { op, ctx, code: err.statusCode } });
    }
  }

  public get events(): readonly ITelemetryEvent[] {
    return this._events;
  }
}