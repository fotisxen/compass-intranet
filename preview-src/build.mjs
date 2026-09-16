import * as esbuild from 'esbuild';
import { fileURLToPath } from 'url';
import path from 'path';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const root = path.resolve(__dirname, '..');

// SPFx's real webpack build scopes every *.module.scss class name via
// sp-css-loader's postcss-modules processing (hashed per-component), so two
// components using the same class name (e.g. both calling it `.card`) never
// collide in the shipped bundle. This preview doesn't run that pipeline, so
// it fakes the same effect: each module's classes are namespaced by the
// file's basename (unique across this project) — must match the prefixing
// done in build-css.mjs when compiling the real CSS.
const scssNamespacedPlugin = {
  name: 'scss-namespaced',
  setup(build) {
    build.onLoad({ filter: /\.module\.scss$/ }, args => {
      const ns = path.basename(args.path, '.module.scss');
      return {
        contents:
          `const ns = ${JSON.stringify(ns)};` +
          'module.exports = new Proxy({}, { get: (_, prop) => (typeof prop === "string" && prop !== "__esModule" ? ns + "__" + prop : undefined) });',
        loader: 'js'
      };
    });
  }
};

// Real SPFx packages like @microsoft/sp-http pull in Microsoft-internal
// dependencies (@msinternal/*, @azure/msal-browser-*-1p) that only resolve
// inside the actual SPFx webpack build (heft build already verifies that
// works). This preview has no real SharePoint context to call anyway, so
// components importing @microsoft/sp-http get a tiny stub instead — just
// enough shape (SPHttpClient.configurations.v1) for the import to resolve;
// the fake spHttpClient passed from entry.tsx handles the actual behavior.
const spHttpStubPlugin = {
  name: 'sp-http-stub',
  setup(build) {
    build.onResolve({ filter: /^@microsoft\/sp-http$/ }, args => ({
      path: args.path,
      namespace: 'sp-http-stub'
    }));
    build.onLoad({ filter: /.*/, namespace: 'sp-http-stub' }, () => ({
      contents: 'exports.SPHttpClient = { configurations: { v1: {} } };',
      loader: 'js'
    }));
  }
};

await esbuild.build({
  absWorkingDir: root,
  entryPoints: [path.join(__dirname, 'entry.tsx')],
  bundle: true,
  outfile: path.join(__dirname, 'dist', 'bundle.js'),
  loader: { '.tsx': 'tsx', '.ts': 'ts', '.png': 'dataurl', '.webp': 'dataurl' },
  plugins: [scssNamespacedPlugin, spHttpStubPlugin],
  define: { 'process.env.NODE_ENV': '"development"' },
  logLevel: 'info'
});

console.log('Build complete.');
