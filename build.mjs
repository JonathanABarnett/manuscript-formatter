import { cp, mkdir, readFile, rm, writeFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';
import * as esbuild from 'esbuild';

const root = dirname(fileURLToPath(import.meta.url));
const out = join(root, 'dist');
const watch = process.argv.includes('--watch');
/** Vercel only needs the static site, not the Electron or CLI bundles. */
const webOnly = process.argv.includes('--web-only');

const shared = {
  bundle: true,
  sourcemap: true,
  logLevel: 'info',
  define: { 'process.env.NODE_ENV': JSON.stringify(watch ? 'development' : 'production') },
};

/** Main and preload run in Node; Electron itself stays external. */
const nodeBundles = [
  {
    ...shared,
    entryPoints: [join(root, 'src/main/main.ts')],
    outfile: join(out, 'main/main.cjs'),
    platform: 'node',
    format: 'cjs',
    target: 'node20',
    external: ['electron'],
  },
  {
    ...shared,
    entryPoints: [join(root, 'src/main/preload.ts')],
    outfile: join(out, 'main/preload.cjs'),
    platform: 'node',
    format: 'cjs',
    target: 'node20',
    external: ['electron'],
  },
  {
    ...shared,
    entryPoints: [join(root, 'src/cli.ts')],
    outfile: join(out, 'cli.cjs'),
    platform: 'node',
    format: 'cjs',
    target: 'node20',
    // esbuild carries the entry file's own hashbang through; adding a banner
    // here would emit a second one and break the file.
  },
];

/** The renderer is a plain sandboxed page with no Node access. */
const rendererBundle = {
  ...shared,
  entryPoints: [join(root, 'src/renderer/desktop.ts')],
  outfile: join(out, 'renderer/renderer.js'),
  platform: 'browser',
  format: 'iife',
  target: 'chrome120',
};

/**
 * The web build runs the same UI and the same engine entirely in the page —
 * no server, so manuscripts never leave the reader's machine.
 */
const webBundle = {
  ...shared,
  entryPoints: [join(root, 'src/web/main.ts')],
  outfile: join(out, 'web/app.js'),
  platform: 'browser',
  format: 'iife',
  target: 'chrome120',
  minify: !watch,
};

/** Content security policy for the hosted page. */
const WEB_CSP =
  "default-src 'self'; script-src 'self'; style-src 'self'; img-src 'self' data:; " +
  "connect-src 'none'; form-action 'none'; base-uri 'none'";

async function copyStatic() {
  if (!webOnly) {
    await mkdir(join(out, 'renderer'), { recursive: true });
    for (const file of ['index.html', 'styles.css']) {
      await cp(join(root, 'src/renderer', file), join(out, 'renderer', file));
    }
  }

  // The web page is derived from the desktop markup rather than duplicated,
  // so the two shells cannot drift apart.
  await mkdir(join(out, 'web'), { recursive: true });
  const html = await readFile(join(root, 'src/renderer/index.html'), 'utf8');
  await writeFile(
    join(out, 'web/index.html'),
    html
      .replace(/content="default-src[^"]*"/, `content="${WEB_CSP}"`)
      .replace('<script src="renderer.js"></script>', '<script src="app.js"></script>'),
  );
  await cp(join(root, 'src/renderer/styles.css'), join(out, 'web/styles.css'));
  await cp(join(root, 'build/icon.png'), join(out, 'web/icon.png'));
}

async function main() {
  if (!watch) await rm(out, { recursive: true, force: true });
  await mkdir(out, { recursive: true });

  const configs = webOnly ? [webBundle] : [...nodeBundles, rendererBundle, webBundle];
  if (watch) {
    const contexts = await Promise.all(configs.map((c) => esbuild.context(c)));
    await Promise.all(contexts.map((c) => c.watch()));
    await copyStatic();
    console.log('watching for changes...');
  } else {
    await Promise.all(configs.map((c) => esbuild.build(c)));
    await copyStatic();
  }
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
