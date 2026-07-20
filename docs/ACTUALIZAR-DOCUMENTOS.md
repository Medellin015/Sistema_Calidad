# Actualizar automáticamente los documentos del buscador

El portal busca sobre el archivo **`documentos.json`** (en la raíz del repositorio).
La app lo carga cada vez que se abre; si no lo encuentra, usa una copia interna
de respaldo. Por eso, para que un documento nuevo aparezca en el buscador, solo
hay que **regenerar `documentos.json`** — no se toca código.

Este documento explica cómo hacerlo **automático** con Power Automate (sin
necesidad de que TI cree nada en Azure), y también cómo hacerlo a mano.

---

## Qué contiene `documentos.json`

Es una **lista** de documentos. Cada documento es un objeto con estos campos
(todos opcionales salvo el enlace y el nombre):

| Campo    | Qué es | Ejemplo |
|----------|--------|---------|
| `codigo` | Código del documento (o nada si no tiene) | `"SBS-F31"` |
| `nombre` | Nombre legible | `"Lista de chequeo lista corta"` |
| `proceso`| Sigla del proceso en el portal | `"BS"` |
| `origen` | Etiqueta cuando no es de un proceso | `"Documentos del SGC"` |
| `path`   | **Ruta del archivo relativa al servidor** (lo más fácil de sacar) | `"/sites/SistemadeGestindeCalidad/Documentos compartidos/Proceso de Contratación/Formatos/SBS-F31 Lista de chequeo lista corta.xlsx"` |
| `url`    | Enlace absoluto ya listo (alternativa a `path`) | `"https://.../SBS-F31%20...xlsx"` |

> **Basta con `nombre` + `path`.** Con la ruta del archivo, la app arma el
> enlace sola (codifica tildes y espacios). Si en cambio pones `url`, debe venir
> ya codificada. El `codigo` y el `proceso` mejoran la búsqueda pero no son
> obligatorios.

Ejemplo mínimo de un `documentos.json`:

```json
[
  { "nombre": "Lista de chequeo lista corta", "codigo": "SBS-F31", "proceso": "BS",
    "path": "/sites/SistemadeGestindeCalidad/Documentos compartidos/Proceso de Contratación/Formatos/SBS-F31 Lista de chequeo lista corta.xlsx" },
  { "nombre": "Acta de inicio", "codigo": "SBS-F08", "proceso": "BS",
    "path": "/sites/SistemadeGestindeCalidad/Documentos compartidos/Proceso de Contratación/Formatos/SBS-F08 Acta de inicio.docx" }
]
```

Siglas de proceso: `OE` (Orientación y Planeación Estratégica), `POL` (Gestión
Comercial), `GOP` (Gestión de Operaciones), `GP` (Gestión de Parques),
`GF` (Financiera), `GA` (Administrativa), `GD` (Documental), `GT` (Talento
Humano), `GTI` (Tecnologías de Información), `BS` (Proceso de Contratación),
`MA` (Mejoramiento Continuo), `CI` (Control Interno), `PD` (Proceso
Disciplinario), `MIPG`.

---

## Opción A — Automático con Power Automate (recomendada)

La idea: un **flujo programado** que cada noche lee la biblioteca del SGC,
arma el JSON y lo sube al repositorio. GitHub Pages republica y los documentos
nuevos aparecen en el buscador.

### Requisitos (una sola vez)
- Un **token de GitHub** (Personal Access Token *fine-grained*) con permiso de
  **Contents: Read and write** solo sobre el repositorio `Medellin015/Sistema_Calidad`.
  (GitHub → Settings → Developer settings → Fine-grained tokens.)
  Guárdalo; se usa en el paso final del flujo.

### Pasos del flujo (Power Automate → *Flujo de nube programado*)

1. **Recurrencia (Recurrence)** — Frecuencia: Día, cada 1 (p. ej. 5:00 a. m.).

2. **Enviar una solicitud HTTP a SharePoint** (*Send an HTTP request to SharePoint*).
   Este es el paso clave: una biblioteca de documentos es, por dentro, una lista
   plana, así que **una sola llamada trae TODOS los archivos de todas las
   subcarpetas**.
   - **Dirección del sitio**: `https://activaparquesyeventos.sharepoint.com/sites/SistemadeGestindeCalidad`
   - **Método**: `GET`
   - **Uri**:
     ```
     _api/web/lists/getbytitle('Documentos')/items?$select=FileLeafRef,FileRef,FSObjType&$top=5000
     ```
     > Si diera error "list not found", el título interno de la biblioteca puede
     > ser `Documents` en vez de `Documentos`; prueba con ese. También puedes
     > mirarlo en *Configuración de la biblioteca* → la URL muestra el nombre.
     > Si tienes **más de 5000 archivos**, hay que paginar con `$skiptoken`
     > (avísame y te dejo la variante con bucle).

3. **Analizar JSON (Parse JSON)** — Contenido: el `body` del paso anterior.
   (Para el *Esquema*, usa "Generar a partir de una muestra" y pega una
   respuesta de ejemplo del paso 2.)

4. **Redactar (Compose)** — construye el arreglo final. Pon esta expresión:
   ```
   @{body('Analizar_JSON')?['d']?['results']}
   ```
   y luego un **Seleccionar (Select)** con:
   - **Desde**: la salida del Compose (los `results`).
   - **Mapa** (modo texto, botón `T`):
     ```json
     {
       "nombre": "@{item()?['FileLeafRef']}",
       "path": "@{item()?['FileRef']}"
     }
     ```
   - Antes, filtra solo archivos (no carpetas) con un **Filtrar arreglo
     (Filter array)**: `item()?['FSObjType']` **es igual a** `0`.

   > Con esto cada documento queda como `{ "nombre": "...", "path": "..." }`.
   > La app deriva el `codigo` del inicio del nombre y arma el enlace. Si
   > quieres además fijar `proceso`, puedes mapear la sigla según la carpeta
   > (aparece dentro de `FileRef`), pero no es imprescindible.

5. **Subir a GitHub** — acción **HTTP** (premium) o **"Enviar solicitud HTTP"**:
   - **Método**: `PUT`
   - **URI**:
     ```
     https://api.github.com/repos/Medellin015/Sistema_Calidad/contents/documentos.json
     ```
   - **Encabezados**:
     - `Authorization`: `Bearer TU_TOKEN`
     - `Accept`: `application/vnd.github+json`
     - `User-Agent`: `activa-sgc-bot`
   - **Cuerpo**:
     ```json
     {
       "message": "Actualizar documentos.json (automático)",
       "content": "@{base64(string(body('Seleccionar')))}",
       "sha": "@{outputs('Obtener_sha_actual')?['body']?['sha']}",
       "branch": "main"
     }
     ```
   - El `sha` es el del `documentos.json` actual; obténlo con un **GET** previo a
     la misma URI (acción HTTP GET → guarda `body.sha`). Si el archivo no existe
     todavía, omite el campo `sha`.

6. Guarda y prueba el flujo (**Test → Manually**). Al terminar, revisa que en el
   repositorio el commit "Actualizar documentos.json (automático)" haya
   cambiado el archivo. En 1–2 minutos, GitHub Pages publica y el buscador ya
   incluye lo nuevo.

> **"Automático" = en la siguiente corrida.** Si subes un documento a mediodía
> y el flujo corre a las 5 a. m., aparecerá al día siguiente. Si lo quieres más
> seguido, cambia la recurrencia a cada hora.

---

## Opción B — A mano (para un cambio puntual y urgente)

1. Descarga `documentos.json` del repositorio.
2. Agrega tu documento al final de la lista, por ejemplo:
   ```json
   { "nombre": "Nombre del documento", "codigo": "XX-F00", "proceso": "BS",
     "path": "/sites/SistemadeGestindeCalidad/Documentos compartidos/Proceso de Contratación/Formatos/XX-F00 Nombre del documento.docx" }
   ```
   (La `path` es la ruta del archivo en SharePoint tras
   `activaparquesyeventos.sharepoint.com`.)
3. Sube el archivo reemplazando el anterior (GitHub → `documentos.json` → editar
   → *Commit changes* a `main`).

---

## Cómo comprobar que quedó

- Abre `https://medellin015.github.io/Sistema_Calidad/#/documentos` y busca el
  código o nombre nuevo.
- El contador del título ("Todos los documentos (N)") debe subir.
- Si algo falla en la generación del JSON, **el buscador no se rompe**: la app
  usa su copia de respaldo interna hasta que el `documentos.json` vuelva a ser
  válido.
