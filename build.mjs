/* Compila app.jsx (JSX) -> app.js (JS puro que carga la página).
 *
 * Uso:
 *   npm install            (una sola vez: instala typescript, ver package.json)
 *   node build.mjs
 *
 * Por qué: así el portal funciona abriendo index.html directamente (file://)
 * y también servido por HTTP, sin depender de Babel en el navegador.
 */
import ts from 'typescript';
import { readFileSync, writeFileSync } from 'node:fs';
import { createHash } from 'node:crypto';

const src = readFileSync(new URL('./app.jsx', import.meta.url), 'utf8');
const { outputText, diagnostics } = ts.transpileModule(src, {
  compilerOptions: {
    jsx: ts.JsxEmit.React,        // JSX -> React.createElement (React global de la CDN)
    target: ts.ScriptTarget.ES2019,
    module: ts.ModuleKind.None,   // script clásico, sin import/export
    removeComments: false,
  },
  fileName: 'app.jsx',
});
if (diagnostics && diagnostics.length) {
  for (const d of diagnostics) console.warn(ts.flattenDiagnosticMessageText(d.messageText, '\n'));
}
const banner = '/* Generado desde app.jsx (fuente editable). No editar a mano: recompilar con `node build.mjs`. */\n';
const appJs = banner + outputText;
writeFileSync(new URL('./app.js', import.meta.url), appJs);

// Cache-busting: estampa en index.html una versión (?v=hash) para app.js y
// styles.css, así el navegador y la CDN de GitHub Pages sirven siempre la
// última versión tras cada despliegue en lugar de la copia en caché.
const css = readFileSync(new URL('./styles.css', import.meta.url), 'utf8');
const v = createHash('sha1').update(appJs).update(css).digest('hex').slice(0, 10);
const idxUrl = new URL('./index.html', import.meta.url);
const idx = readFileSync(idxUrl, 'utf8')
  .replace(/(href=")styles\.css(?:\?v=[^"]*)?(")/g, `$1styles.css?v=${v}$2`)
  .replace(/(src=")app\.js(?:\?v=[^"]*)?(")/g, `$1app.js?v=${v}$2`);
writeFileSync(idxUrl, idx);
console.log(`app.js generado desde app.jsx; index.html versionado (?v=${v}).`);
