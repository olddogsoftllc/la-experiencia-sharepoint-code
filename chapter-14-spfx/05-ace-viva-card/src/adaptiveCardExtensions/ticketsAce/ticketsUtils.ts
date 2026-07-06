/**
 * Pure helpers for the Tickets ACE, extracted so they can be unit-tested
 * without the SharePoint runtime. `import type` keeps ITicket type-only (no
 * runtime load of @microsoft/sp-adaptive-card-extension-base).
 */
import type { ITicket } from './TicketsAceAdaptiveCardExtension';

/** Counts the tickets whose status is "open". */
export function countOpenTickets(tickets: ITicket[]): number {
  return tickets.filter((t) => t.status === 'open').length;
}

/** Builds the dashboard card header: "1 open ticket" / "3 open tickets". */
export function buildOpenTicketsHeader(openCount: number): string {
  return `${openCount} open ticket${openCount === 1 ? '' : 's'}`;
}