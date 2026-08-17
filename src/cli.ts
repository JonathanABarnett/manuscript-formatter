#!/usr/bin/env node
import { parseArgs } from 'node:util';
import { basename, resolve } from 'node:path';
import {
  analyzeDocumentPaths,
  formatManuscript,
  suggestOutputPath,
} from './core/platform/node.js';
import { inchLabel } from './core/analyze/pageSize.js';
import {
  STYLE_ROLES,
  type ChapterStartMode,
  type FormatOptions,
  type StyleRole,
} from './core/types.js';

const twipsLabel = (twips: number): string => `${inchLabel(twips)}"`;

const USAGE = `
Manuscript Formatter — apply a reference document's layout to a manuscript.

  manuscript-formatter <reference.docx> <manuscript.docx> [output.docx] [options]

Arguments
  reference.docx    A .docx already formatted for the intended book.
  manuscript.docx   The .docx holding the book's content. Never modified.
  output.docx       Where to write. Defaults to "<manuscript> (formatted).docx".

Options
  --chapter-start <mode>   newPage (default), oddPage, or continuous
  --scene-break <text>     Replace scene-break ornaments with this text
  --style <role>=<id>      Force a role onto a reference style. Repeatable.
                           Roles: ${STYLE_ROLES.join(', ')}
  --no-front-matter        Leave out everything before the first chapter
  --keep-blanks            Keep blank paragraphs instead of relying on styles
  --keep-indents           Keep tabs and spaces typed as first-line indents
  --no-emphasis            Drop bold and italic from the manuscript text
  --smart-quotes           Convert " to curly quotes, -- to em dashes
  --collapse-spaces        Collapse runs of spaces to one
  --plan                   Analyze and print the plan without writing a file
  --json                   Machine-readable output
  -h, --help               Show this message
`.trim();

interface Cli {
  values: Record<string, unknown>;
  positionals: string[];
}

function parse(): Cli {
  return parseArgs({
    allowPositionals: true,
    options: {
      'chapter-start': { type: 'string' },
      'scene-break': { type: 'string' },
      style: { type: 'string', multiple: true },
      'no-front-matter': { type: 'boolean' },
      'keep-blanks': { type: 'boolean' },
      'keep-indents': { type: 'boolean' },
      'no-emphasis': { type: 'boolean' },
      'smart-quotes': { type: 'boolean' },
      'collapse-spaces': { type: 'boolean' },
      plan: { type: 'boolean' },
      json: { type: 'boolean' },
      help: { type: 'boolean', short: 'h' },
    },
  }) as Cli;
}

function buildOptions(values: Record<string, unknown>): Partial<FormatOptions> {
  const options: Partial<FormatOptions> = {};

  const mode = values['chapter-start'] as string | undefined;
  if (mode !== undefined) {
    if (!['newPage', 'oddPage', 'continuous'].includes(mode)) {
      throw new Error(`--chapter-start must be newPage, oddPage or continuous (got "${mode}")`);
    }
    options.chapterStart = mode as ChapterStartMode;
  }

  if (values['scene-break'] !== undefined) options.sceneBreakText = values['scene-break'] as string;
  if (values['no-front-matter']) options.includeFrontMatter = false;
  if (values['keep-blanks']) options.removeEmptyParagraphs = false;
  if (values['keep-indents']) options.removeManualIndents = false;
  if (values['no-emphasis']) options.keepEmphasis = false;
  if (values['smart-quotes']) options.smartTypography = true;
  if (values['collapse-spaces']) options.collapseMultipleSpaces = true;

  const styles = (values.style as string[] | undefined) ?? [];
  if (styles.length > 0) {
    const roleStyles: Partial<Record<StyleRole, string | null>> = {};
    for (const pair of styles) {
      const eq = pair.indexOf('=');
      if (eq === -1) throw new Error(`--style expects role=styleId (got "${pair}")`);
      const role = pair.slice(0, eq) as StyleRole;
      if (!STYLE_ROLES.includes(role)) {
        throw new Error(`Unknown role "${role}". Roles: ${STYLE_ROLES.join(', ')}`);
      }
      roleStyles[role] = pair.slice(eq + 1) || null;
    }
    options.roleStyles = roleStyles;
  }
  return options;
}

async function main(): Promise<number> {
  const { values, positionals } = parse();
  if (values.help || positionals.length === 0) {
    console.log(USAGE);
    return positionals.length === 0 && !values.help ? 1 : 0;
  }
  if (positionals.length < 2) {
    console.error('Both a reference document and a manuscript are required.\n');
    console.error(USAGE);
    return 1;
  }

  const referencePath = resolve(positionals[0]);
  const manuscriptPath = resolve(positionals[1]);
  const options = buildOptions(values);

  if (values.plan) {
    const { profile, analysis, suggestedOptions } = await analyzeDocumentPaths(
      referencePath,
      manuscriptPath,
    );
    if (values.json) {
      console.log(JSON.stringify({ profile, analysis, suggestedOptions }, null, 2));
      return 0;
    }
    printPlan(profile, analysis, suggestedOptions);
    return 0;
  }

  const outputPath = positionals[2]
    ? resolve(positionals[2])
    : await suggestOutputPath(manuscriptPath);

  const result = await formatManuscript({
    referencePath,
    manuscriptPath,
    outputPath,
    options,
  });

  if (values.json) {
    console.log(JSON.stringify(result, null, 2));
    return 0;
  }

  const s = result.stats;
  console.log(`Wrote ${basename(result.outputPath)}`);
  console.log(`  ${result.outputPath}`);
  console.log(
    `  ${s.paragraphsWritten.toLocaleString()} paragraphs · ${s.chapters} chapters · ` +
      `${s.sceneBreaks} scene breaks · ${s.wordCount.toLocaleString()} words`,
  );
  if (s.imagesCopied > 0) console.log(`  ${s.imagesCopied} images copied`);
  if (s.footnotesCopied > 0) console.log(`  ${s.footnotesCopied} footnotes copied`);
  if (result.warnings.length > 0) {
    console.log('\nWorth checking:');
    for (const warning of result.warnings) console.log(`  - ${warning}`);
  }
  return 0;
}

function printPlan(
  profile: Awaited<ReturnType<typeof analyzeDocumentPaths>>['profile'],
  analysis: Awaited<ReturnType<typeof analyzeDocumentPaths>>['analysis'],
  suggested: FormatOptions,
): void {
  const m = profile.pageSetup.margins;
  console.log(`Reference: ${profile.fileName}`);
  console.log(`  Page       ${profile.pageSizeLabel}`);
  console.log(
    `  Margins    ${twipsLabel(m.top)} top, ${twipsLabel(m.bottom)} bottom, ` +
      `${twipsLabel(m.left)}/${twipsLabel(m.right)} sides, gutter ${twipsLabel(m.gutter)}`,
  );
  console.log(`  Body       ${profile.bodyFontName ?? 'default'} ${profile.bodyFontSizePt ?? ''}pt`);
  console.log(`  Chapters   start ${suggested.chapterStart}`);
  console.log('\nStyle mapping');
  for (const role of STYLE_ROLES) {
    const id = profile.roleStyles[role];
    const why = profile.roleEvidence[role];
    console.log(`  ${role.padEnd(18)} ${(id ?? '(fallback)').padEnd(20)} ${why ?? ''}`);
  }
  console.log(`\nManuscript: ${analysis.fileName}`);
  console.log(
    `  ${analysis.wordCount.toLocaleString()} words · ${analysis.paragraphCount} paragraphs · ` +
      `${analysis.chapterCount} chapters · ${analysis.sceneBreakCount} scene breaks`,
  );
  const warnings = [...profile.warnings, ...analysis.warnings];
  if (warnings.length > 0) {
    console.log('\nWorth checking:');
    for (const warning of warnings) console.log(`  - ${warning}`);
  }
}

main()
  .then((code) => process.exit(code))
  .catch((err: unknown) => {
    console.error(err instanceof Error ? err.message : String(err));
    process.exit(1);
  });
