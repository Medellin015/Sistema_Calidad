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
writeFileSync(new URL('./app.js', import.meta.url), banner + outputText);
console.log('app.js generado desde app.jsx.');
