import 'jest';
import { resolveThemeName, hostRunningSuffix } from './teamsUtils';

describe('teamsUtils.resolveThemeName', () => {
  it('returns "dark" when the theme is inverted', () => {
    expect(resolveThemeName(true)).toBe('dark');
  });
  it('returns "default" when the theme is not inverted', () => {
    expect(resolveThemeName(false)).toBe('default');
  });
});

describe('teamsUtils.hostRunningSuffix', () => {
  it('says "running inside Microsoft Teams" when the Teams SDK is present', () => {
    expect(hostRunningSuffix(true)).toBe(' (running inside Microsoft Teams)');
  });
  it('says "running in SharePoint" when there is no Teams context', () => {
    expect(hostRunningSuffix(false)).toBe(' (running in SharePoint)');
  });
});