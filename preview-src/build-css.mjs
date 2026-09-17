import { fileURLToPath } from 'url';
import path from 'path';
import fs from 'fs';
import { execFileSync } from 'child_process';
import { globSync } from 'fs';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const root = path.resolve(__dirname, '..');
const sassBin = path.join(root, 'node_modules', '.bin', 'sass.cmd');

const files = globSync('src/**/*.module.scss', { cwd: root });

let bundle = '';
for (const rel of files) {
  const abs = path.join(root, rel);
  const ns = path.basename(rel, '.module.scss');
  const css = execFileSync(sassBin, [abs, '--no-source-map'], { encoding: 'utf8', shell: true });

  // Namespace every class selector the same way build.mjs's esbuild plugin
  // namespaces the JS-side className lookups, so `styles.card` in two
  // different components never collides in this preview bundle. `:global(...)`
  // wrappers (real CSS Modules syntax, e.g. targeting a Leaflet class) must
  // stay untouched — protect them first, namespace everything else, then
  // restore them verbatim.
  const globals = [];
  const protectedCss = css.replace(/:global\(([^)]*)\)/g, (match, inner) => {
    globals.push(inner);
    return `__GLOBAL_${globals.length - 1}__`;
  });
  const namespacedProtected = protectedCss.replace(/\.([A-Za-z_][\w-]*)/g, (match, cls) => `.${ns}__${cls}`);
  const namespaced = namespacedProtected.replace(/__GLOBAL_(\d+)__/g, (match, i) => globals[Number(i)]);

  bundle += `/* ${rel} */\n${namespaced}\n`;
}

fs.writeFileSync(path.join(__dirname, 'dist', 'bundle.css'), bundle);
console.log(`Compiled ${files.length} scss modules -> preview-src/dist/bundle.css`);
