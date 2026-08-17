# Deploying the web version to Vercel

The web build is a **static site**. The whole formatting engine is compiled into the page and runs
in the browser, so there is no server, no API route, and no upload. Vercel only ever serves files.

That has three consequences worth knowing before you start:

- **Nothing is uploaded.** A manuscript is read, analysed and rebuilt on the reader's own machine.
  Vercel never sees it. For a tool whose input is an unpublished book, that is the point.
- **No limits apply.** There is no 4.5 MB request-body cap and no function timeout, because no
  function runs. A 200 MB manuscript with images works exactly as well as a 40 KB one.
- **It costs nothing at runtime.** Static hosting on the Hobby plan is free.

## What gets deployed

`vercel.json` is already committed and needs no editing:

```json
{
  "buildCommand": "npm run build:web",
  "outputDirectory": "dist/web",
  "framework": null
}
```

`npm run build:web` writes four files into `dist/web`: `index.html`, `app.js`, `styles.css` and
`icon.png`. Nothing else is published — `dist/web` is wiped and rebuilt on every run, so stray
files cannot leak into a deploy.

## Option A — deploy from the CLI

Fastest for a first deploy. From the project root:

```bash
npx vercel
```

It will ask a short series of questions. Accept the detected settings — the build command and output
directory come from `vercel.json`, so **do not** override them when prompted. That produces a
preview URL.

When you are happy with it, publish to production:

```bash
npx vercel --prod
```

## Option B — deploy from GitHub

Better if you want every push built automatically.

1. Push the repo to GitHub. (It has no remote yet: `git remote add origin <url>` then
   `git push -u origin master`.)
2. At [vercel.com/new](https://vercel.com/new), import the repository.
3. Leave every build setting on its default. Vercel reads `vercel.json`; if the form asks anyway,
   the build command is `npm run build:web` and the output directory is `dist/web`.
4. Deploy.

Each push to the default branch then goes to production, and every other branch and pull request
gets its own preview URL.

## Verifying a deploy

Open the URL and check three things:

1. The page loads and shows **Choose your two documents**.
2. Pick a template and a manuscript. The review screen should appear with page previews.
3. Open your browser's network panel and format a document. **No request should be made** — the
   work happens in the page. That is the privacy guarantee, and it is observable.

## Headers

`vercel.json` sets `X-Content-Type-Options`, `Referrer-Policy`, `X-Frame-Options` and a
`Permissions-Policy` that switches off camera, microphone, geolocation and interest-cohort. The page
itself ships a Content-Security-Policy with `connect-src 'none'`, so even a compromised dependency
could not exfiltrate a manuscript — the browser refuses the connection.

If you add anything that genuinely needs the network later, that CSP is the thing to change, in
`src/web/index.html`. Leave it as it is otherwise.

## Custom domain

Project → **Settings** → **Domains** → add the domain and follow the DNS instructions. Nothing in
the app depends on the hostname.

## Troubleshooting

**Build fails with a TypeScript error.** The build runs `tsc` indirectly through esbuild, which does
not typecheck. Run `npm run typecheck` locally — Vercel is failing for the same reason your machine
would.

**Page loads but nothing happens when you pick a file.** Check the browser console. A CSP violation
appears there explicitly; anything else is a genuine bug.

**Deploy succeeds but serves a stale build.** `dist/` is gitignored, so Vercel always builds from
source. A stale page is browser cache — hard-reload with Ctrl+F5.

**"No Output Directory named dist/web".** The build command did not run. Confirm `vercel.json` is at
the repository root and that the Vercel project's root directory is the repository root, not a
subfolder.
