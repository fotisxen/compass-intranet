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

await esbuild.build({
  absWorkingDir: root,
  entryPoints: [path.join(__dirname, 'entry.tsx')],
  bundle: true,
  outfile: path.join(__dirname, 'dist', 'bundle.js'),
  loader: { '.tsx': 'tsx', '.ts': 'ts' },
  plugins: [scssNamespacedPlugin],
  define: { 'process.env.NODE_ENV': '"development"' },
  logLevel: 'info'
});

console.log('Build complete.');
