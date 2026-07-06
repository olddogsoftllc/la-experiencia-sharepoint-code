import 'jest';
import type { ITicket } from './TicketsAceAdaptiveCardExtension';
import { countOpenTickets, buildOpenTicketsHeader } from './ticketsUtils';

function ticket(id: string, status: ITicket['status']): ITicket {
  return { id, title: `T${id}`, status };
}

describe('ticketsUtils.countOpenTickets', () => {
  it('returns 0 when no tickets are open', () => {
    expect(countOpenTickets([ticket('1', 'inProgress'), ticket('2', 'done')])).toBe(0);
  });
  it('counts only the open ones', () => {
    const tickets = [ticket('1', 'open'), ticket('2', 'done'), ticket('3', 'open')];
    expect(countOpenTickets(tickets)).toBe(2);
  });
  it('returns 0 for an empty list', () => {
    expect(countOpenTickets([])).toBe(0);
  });
});

describe('ticketsUtils.buildOpenTicketsHeader', () => {
  it('uses the singular form for exactly 1', () => {
    expect(buildOpenTicketsHeader(1)).toBe('1 open ticket');
  });
  it('uses the plural form for 0 and >1', () => {
    expect(buildOpenTicketsHeader(0)).toBe('0 open tickets');
    expect(buildOpenTicketsHeader(3)).toBe('3 open tickets');
  });
});