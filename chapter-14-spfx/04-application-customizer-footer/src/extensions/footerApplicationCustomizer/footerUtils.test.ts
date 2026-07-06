import 'jest';
import { resolveFooterText, footerBarStyle, DEFAULT_FOOTER_TEXT } from './footerUtils';

describe('footerUtils.resolveFooterText', () => {
  it('returns the given text when non-empty', () => {
    expect(resolveFooterText('My Footer')).toBe('My Footer');
  });
  it('falls back to the default when the text is empty', () => {
    expect(resolveFooterText('')).toBe(DEFAULT_FOOTER_TEXT);
  });
  it('falls back to the default when the text is undefined', () => {
    expect(resolveFooterText(undefined)).toBe(DEFAULT_FOOTER_TEXT);
  });
});

describe('footerUtils.footerBarStyle', () => {
  it('returns the brand blue background and white text', () => {
    const s = footerBarStyle();
    expect(s.backgroundColor).toBe('#0078d4');
    expect(s.color).toBe('#ffffff');
  });
  it('returns padding and fontSize strings', () => {
    const s = footerBarStyle();
    expect(s.padding).toBe('6px 16px');
    expect(s.fontSize).toBe('12px');
  });
});