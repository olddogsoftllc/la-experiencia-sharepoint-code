import 'jest';
import { shouldShowExport, buildNotifyMessage, buildExportMessage } from './commandUtils';

describe('commandUtils.shouldShowExport', () => {
  it('returns false when nothing is selected', () => {
    expect(shouldShowExport(0)).toBe(false);
  });
  it('returns true when one or more rows are selected', () => {
    expect(shouldShowExport(1)).toBe(true);
    expect(shouldShowExport(42)).toBe(true);
  });
});

describe('commandUtils.buildNotifyMessage', () => {
  it('uses the prefix when provided', () => {
    expect(buildNotifyMessage('Selected', 3)).toBe('Selected: 3 row(s).');
  });
  it('defaults to "Selected" when no prefix is given', () => {
    expect(buildNotifyMessage(undefined, 0)).toBe('Selected: 0 row(s).');
  });
});

describe('commandUtils.buildExportMessage', () => {
  it('includes the selected count', () => {
    expect(buildExportMessage(5)).toBe('Export: 5 row(s) selected (demo).');
    expect(buildExportMessage(0)).toBe('Export: 0 row(s) selected (demo).');
  });
});