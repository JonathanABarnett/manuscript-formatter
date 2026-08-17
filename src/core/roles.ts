import type { BlockRole, StyleRole } from './types.js';

export type RoleStyleMap = Record<StyleRole, string | null>;

/**
 * Which reference style a role resolves to, walking the fallbacks used when the
 * template has no dedicated style for it.
 *
 * Shared by the composer and the review preview so that what the reviewer sees
 * on screen is what actually gets written into the document.
 */
export function styleForRole(role: BlockRole, roles: RoleStyleMap): string | null {
  switch (role) {
    case 'chapterTitle':
      return roles.chapterTitle;
    case 'partTitle':
      return roles.partTitle ?? roles.chapterTitle;
    case 'chapterSubtitle':
      // A subtitle belongs to the chapter's display head, so without a
      // dedicated style it joins the title rather than dropping to a subhead
      // meant for use mid-chapter.
      return roles.chapterSubtitle ?? roles.chapterTitle ?? roles.subheading;
    case 'subheading':
      return roles.subheading ?? roles.chapterSubtitle;
    case 'frontMatterTitle':
      return roles.frontMatterTitle ?? roles.chapterTitle;
    case 'frontMatter':
      return roles.frontMatter ?? roles.bodyFirst ?? roles.body;
    case 'copyright':
      return roles.copyright ?? roles.frontMatter ?? roles.body;
    case 'bodyFirst':
      return roles.bodyFirst ?? roles.body;
    case 'blockQuote':
      return roles.blockQuote ?? roles.body;
    case 'listItem':
      return roles.listItem ?? roles.body;
    case 'sceneBreak':
      return roles.sceneBreak ?? roles.body;
    case 'body':
    default:
      return roles.body;
  }
}
