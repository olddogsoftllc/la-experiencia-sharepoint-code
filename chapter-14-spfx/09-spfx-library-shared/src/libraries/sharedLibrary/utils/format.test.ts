import { truncate, formatRelativeDate } from './format';

describe('truncate', () => {
  it('returns the text unchanged when within the limit', () => {
    expect(truncate('hello', 10)).toBe('hello');
  });

  it('returns the text unchanged at exactly the limit', () => {
    expect(truncate('hello', 5)).toBe('hello');
  });

  it('truncates and adds an ellipsis when over the limit', () => {
    expect(truncate('hello world', 8)).toBe('hello w…');
  });

  it('handles max of 1', () => {
    expect(truncate('abc', 1)).toBe('…');
  });
});

describe('formatRelativeDate', () => {
  it('returns "today" for a recent ISO date', () => {
    const iso = new Date().toISOString();
    expect(formatRelativeDate(iso)).toBe('today');
  });

  it('returns "yesterday" for ~1 day ago', () => {
    const d = new Date();
    d.setDate(d.getDate() - 1);
    d.setHours(d.getHours() - 1); // ~25h ago to be safe
    expect(formatRelativeDate(d.toISOString())).toBe('yesterday');
  });

  it('returns "N days ago" for several days ago', () => {
    const d = new Date();
    d.setDate(d.getDate() - 5);
    expect(formatRelativeDate(d.toISOString())).toBe('5 days ago');
  });
});