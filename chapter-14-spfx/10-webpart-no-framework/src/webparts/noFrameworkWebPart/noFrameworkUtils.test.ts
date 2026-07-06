import 'jest';
import { buildMarkup } from './noFrameworkUtils';

describe('noFrameworkUtils.buildMarkup', () => {
  const opts = {
    containerClass: 'wp', teamsClass: 'teams', darkClass: 'dark', btnClass: 'btn'
  };

  it('embeds the heading, site and user', () => {
    const html = buildMarkup('My Title', 'Book Test', 'Efren', opts);
    expect(html).toContain('<h2>My Title</h2>');
    expect(html).toContain('Site: <strong>Book Test</strong>');
    expect(html).toContain('Hello, <strong>Efren</strong>');
  });

  it('uses the Reload button with the btn class', () => {
    const html = buildMarkup('T', 'S', 'U', opts);
    expect(html).toContain('<button class="btn">Reload</button>');
  });

  it('applies the container, teams and dark classes on the section', () => {
    const html = buildMarkup('T', 'S', 'U', opts);
    expect(html).toContain('<section class="wp teams dark">');
  });

  it('leaves the teams/dark classes empty when not set', () => {
    const html = buildMarkup('T', 'S', 'U', { ...opts, teamsClass: '', darkClass: '' });
    expect(html).toContain('<section class="wp  ">');
  });
});