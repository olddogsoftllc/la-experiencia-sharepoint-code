import 'jest';
import { getPriorityClass, buildBadgeText } from './priorityUtils';

describe('priorityUtils.getPriorityClass', () => {
  it('returns "high" when the value contains "high" (case-insensitive)', () => {
    expect(getPriorityClass('high')).toBe('high');
    expect(getPriorityClass('HIGH')).toBe('high');
    expect(getPriorityClass('Urgent — high priority')).toBe('high');
  });

  it('returns "low" when the value contains "low" but not "high"', () => {
    expect(getPriorityClass('low')).toBe('low');
    expect(getPriorityClass('Low')).toBe('low');
  });

  it('returns "normal" for anything else', () => {
    expect(getPriorityClass('normal')).toBe('normal');
    expect(getPriorityClass('medium')).toBe('normal');
    expect(getPriorityClass('')).toBe('normal');
    expect(getPriorityClass(undefined)).toBe('normal');
  });

  it('prefers "high" when both substrings are present (high is checked first)', () => {
    expect(getPriorityClass('high and low')).toBe('high');
  });
});

describe('priorityUtils.buildBadgeText', () => {
  it('returns the raw value when no prefix is given', () => {
    expect(buildBadgeText('high')).toBe('high');
  });

  it('prepends the prefix with ": " when a prefix is given', () => {
    expect(buildBadgeText('high', 'Prio')).toBe('Prio: high');
  });

  it('treats undefined value as empty string', () => {
    expect(buildBadgeText(undefined, 'Prio')).toBe('Prio: ');
    expect(buildBadgeText(undefined)).toBe('');
  });
});