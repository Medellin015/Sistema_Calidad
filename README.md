# Sistema_Calidad

Portal de consulta del **Sistema de Gestión de Calidad (SGC)** de ACTIVA:
un mapa de procesos interactivo (cada proceso abre su página con las
carpetas reales de SharePoint: Procedimientos, Formatos, Manuales,
Instructivos, Plantillas), guías paso a paso con flujograma y el
inventario de documentos oficiales con enlace directo al repositorio.

## Estructura

| Archivo | Qué contiene |
|---|---|
| `index.html` | Esqueleto de la página y carga de librerías (React y Tailwind por CDN). |
| `styles.css` | Estilos propios: fuentes, el "sendero" del flujograma y animaciones. |
| `app.jsx` | **Fuente editable** (JSX): configuración, datos (procesos, documentos, guías) y componentes React. |
| `app.js` | **Generado** a partir de `app.jsx`. Es el archivo que carga la página. No editar a mano. |
| `build.mjs` | Compila `app.jsx` → `app.js`. |

## Cómo verlo

Como `app.js` es JavaScript puro (sin compilación en el navegador), el portal
funciona de dos maneras:

- **Abriendo `index.html` directamente** con doble clic (`file://`), o
- **Servido por HTTP**, por ejemplo con GitHub Pages o `python3 -m http.server`.

> Requiere conexión a internet para cargar React y Tailwind desde su CDN.

## Editar el contenido

Los datos (documentos, guías, procesos) están al inicio de `app.jsx`.
Después de editar ese archivo, regenera `app.js`:

```bash
npm install     # una sola vez: instala TypeScript (ver package.json)
node build.mjs  # compila app.jsx -> app.js
```
