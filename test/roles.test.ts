import { describe, expect, it } from 'vitest';
import { styleForRole, type RoleStyleMap } from '../src/core/roles.js';
import { STYLE_ROLES, type StyleRole } from '../src/core/types.js';

/**
 * The composer and the review preview both resolve roles through this map, so
 * a change here silently changes what the preview promises. These pin the
 * fallbacks that matter.
 */

const empty = (): RoleStyleMap =>
  Object.fromEntries(STYLE_ROLES.map((r) => [r, null])) as RoleStyleMap;

const withStyles = (over: Partial<Record<StyleRole, string>>): RoleStyleMap => ({
  ...empty(),
  ...over,
});

describe('resolving a role to a template style', () => {
  it('uses the dedicated style when the template has one', () => {
    const roles = withStyles({
      body: 'Body',
      chapterTitle: 'ChapTitle',
      copyright: 'CopyPage',
      sceneBreak: 'Scene',
    });
    expect(styleForRole('chapterTitle', roles)).toBe('ChapTitle');
    expect(styleForRole('copyright', roles)).toBe('CopyPage');
    expect(styleForRole('sceneBreak', roles)).toBe('Scene');
  });

  it('sends a chapter subtitle to the chapter title, not to a mid-chapter subhead', () => {
    const roles = withStyles({ body: 'Body', chapterTitle: 'ChapTitle', subheading: 'Subhead' });
    expect(styleForRole('chapterSubtitle', roles)).toBe('ChapTitle');
  });

  it('falls back to the front-matter style for a copyright page', () => {
    const roles = withStyles({ body: 'Body', frontMatter: 'Front' });
    expect(styleForRole('copyright', roles)).toBe('Front');
  });

  it('falls back to body text when the template offers nothing closer', () => {
    const roles = withStyles({ body: 'Body' });
    for (const role of ['sceneBreak', 'blockQuote', 'listItem', 'bodyFirst', 'copyright'] as const) {
      expect(styleForRole(role, roles)).toBe('Body');
    }
  });

  it('returns null rather than guessing when the template defines nothing', () => {
    expect(styleForRole('chapterTitle', empty())).toBeNull();
    expect(styleForRole('body', empty())).toBeNull();
  });
});
