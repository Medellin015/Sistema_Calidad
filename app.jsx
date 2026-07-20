'use strict';
/* Fuente editable (JSX). La página carga app.js (compilado), NO este archivo.
   Tras editar aquí, regenera app.js con:  node build.mjs   (ver README). */
/* ============================================================================
   Portal SGC ACTIVA — app de consulta del Sistema de Gestión de Calidad
   Estructura del archivo:
     1. Configuración (repositorio SharePoint y utilidades)
     2. Datos: procesos, documentos del repositorio y guías paso a paso
     3. Componentes de UI (encabezado, buscador, mapa, flujograma, vistas)
     4. Aplicación y enrutado por hash (#/proceso/GD, #/guia/rp, #/documentos)
   Notas:
     - App informativa: sin backend. Los enlaces abren el repositorio oficial
       en SharePoint y requieren sesión Microsoft de ACTIVA.
     - Fuente del inventario: carpeta "Documentos Word" del sitio SGC
       (corte 11/07/2026).
     - Publicación: GitHub Pages (build automático al fusionar a main).
   ============================================================================ */
const { useState, useMemo, useEffect } = React;

/* ===== 1. Configuración ===== */
const SITIO_SGC =
  'https://activaparquesyeventos.sharepoint.com/sites/SistemadeGestindeCalidad';
const BIBLIOTECA_SGC = SITIO_SGC + '/Documentos compartidos';
const CARPETA_SGC_DEFAULT = BIBLIOTECA_SGC + '/Documentos Word/';

// El porqué: los nombres de archivo traen tildes y espacios; el enlace se
// construye codificando solo el nombre para no romper la ruta de la carpeta.
const enlaceDoc = (archivo) => encodeURI(CARPETA_SGC_DEFAULT) + encodeURIComponent(archivo);

// Enlace a la carpeta de un proceso en SharePoint (o a una de sus subcarpetas).
// encodeURI conserva las barras y codifica tildes y espacios de la ruta.
const enlaceCarpeta = (carpetaProceso, subcarpeta) =>
  encodeURI(BIBLIOTECA_SGC + '/' + carpetaProceso + (subcarpeta ? '/' + subcarpeta : ''));

const ICONO_SECCION = {
  'Procedimientos': '📋', 'Formatos': '📝', 'Manuales': '📘',
  'Instructivos': '🧾', 'Plantillas': '📐', 'Documentos del proceso': '🗃️',
  'Planeación Estratégica': '🗺️',
  'Actas de Comité': '🗒️', 'Autodiagnósticos': '🩺', 'Políticas MIPG': '📜',
  'Certificados curso de integridad': '🎓',
};

const TIPOS = { P: 'Procedimiento', I: 'Instructivo', F: 'Formato', M: 'Manual / Política', C: 'Caracterización' };
const tipoDeCodigo = (codigo) => {
  const m = /^[A-Z]+-\s?([PIFMC])/.exec(codigo || '');
  return m ? TIPOS[m[1]] : 'Documento';
};

/* ===== 2. Datos ===== */

// Mapa de procesos. Cada proceso corresponde 1:1 con una carpeta real de la
// biblioteca "Documentos compartidos" del sitio SGC, y sus "secciones" son
// las subcarpetas que esa carpeta tiene de verdad (inventario del 14/07/2026).
// seccion.carpeta es el nombre EXACTO de la subcarpeta en SharePoint (ojo:
// en Disciplinarios la carpeta se llama "Procedimiento", en singular).
const S = (nombre, carpeta) => ({ nombre, carpeta: carpeta || nombre });
const FRANJAS = [
  { nombre: 'Estratégicos', desc: 'Dan la dirección y la voz de la entidad', procesos: [
    // Orientación y Planeación se gestionan como un solo proceso estratégico.
    // En SharePoint siguen siendo dos carpetas: la sección "Planeación
    // Estratégica" usa `base` para enlazar la carpeta hermana completa.
    { sigla: 'OE', nombre: 'Orientación y Planeación Estratégica',
      carpeta: 'Orientación Estratégica',
      secciones: [S('Procedimientos'), S('Formatos'), S('Manuales'), S('Plantillas'),
        { nombre: 'Planeación Estratégica', carpeta: '', base: 'Planeación Estratégica' }] },
  ]},
  { nombre: 'Misionales', desc: 'La razón de ser: parques y eventos', procesos: [
    { sigla: 'POL', nombre: 'Gestión Comercial',
      carpeta: 'Gestión Comercial',
      secciones: [S('Procedimientos'), S('Formatos'), S('Instructivos')] },
    { sigla: 'GOP', nombre: 'Gestión de Operaciones',
      carpeta: 'Gestión de Operaciones',
      secciones: [S('Procedimientos'), S('Formatos'), S('Instructivos')] },
    // Proceso nuevo: aún no tiene carpeta en SharePoint ni documentos.
    { sigla: 'GP', nombre: 'Gestión de Parques', enConstruccion: true,
      carpeta: '', secciones: [] },
  ]},
  { nombre: 'De apoyo', desc: 'Hacen posible la operación', procesos: [
    { sigla: 'GF', nombre: 'Gestión Financiera',
      carpeta: 'Gestión Financiera',
      secciones: [S('Procedimientos'), S('Formatos'), S('Plantillas')] },
    { sigla: 'GA', nombre: 'Gestión Administrativa',
      carpeta: 'Gestión Administrativa',
      secciones: [S('Procedimientos'), S('Formatos'), S('Instructivos')] },
    { sigla: 'GD', nombre: 'Gestión Documental',
      carpeta: 'Gestión Documental',
      secciones: [S('Procedimientos'), S('Formatos')] },
    { sigla: 'GT', nombre: 'Gestión del Talento Humano',
      carpeta: 'Gestión del Talento Humano',
      secciones: [S('Procedimientos'), S('Formatos'), S('Instructivos')] },
    { sigla: 'GTI', nombre: 'Gestión de Tecnologías de la Información',
      carpeta: 'Gestión de Tecnologías de Información',
      secciones: [S('Instructivos')] },
    { sigla: 'BS', nombre: 'Proceso de Contratación',
      carpeta: 'Proceso de Contratación',
      secciones: [S('Procedimientos'), S('Formatos'), S('Manuales')] },
  ]},
  { nombre: 'Evaluación y mejora', desc: 'Controlan y hacen crecer el sistema', procesos: [
    { sigla: 'MA', nombre: 'Mejoramiento Continuo',
      carpeta: 'Gestión de Procesos y Mejoramiento Continuo',
      secciones: [S('Procedimientos')] },
    { sigla: 'CI', nombre: 'Gestión de Control Interno',
      carpeta: 'Gestión de Control Interno',
      secciones: [S('Procedimientos'), S('Formatos')] },
    { sigla: 'PD', nombre: 'Proceso Disciplinario',
      carpeta: 'Gestión Proceso Disciplinario y juzgamiento',
      secciones: [S('Procedimientos', 'Procedimiento'), S('Formatos')] },
  ]},
];
// MIPG atraviesa todos los procesos: se dibuja como banda transversal del mapa.
const MIPG = {
  sigla: 'MIPG', nombre: 'Modelo Integrado de Planeación y Gestión',
  carpeta: 'Modelo Integrado de Planeación y Gestión', franja: 'Transversal',
  secciones: [
    S('Políticas MIPG', 'Políticas Word MIPG'),
    S('Actas de Comité', 'ACTAS DE COMITÉ'),
    S('Autodiagnósticos', 'AUTODIAGNOSTICOS'),
    S('Certificados curso de integridad', 'Certificados Curso Integridad, transparencia y lucha contra la corrupción'),
  ],
};
const PROCESOS = FRANJAS.flatMap((f) => f.procesos.map((p) => ({ ...p, franja: f.nombre })))
  .concat([MIPG]);

// Inventario real de la carpeta "Documentos Word" del sitio SGC.
// estado: 'vigente' (enlace directo) | 'aprobacion' (borrador nuevo, aún sin publicar).
const DOCUMENTOS = [
  // — Orientación Estratégica
  { codigo: 'OE-F08', nombre: 'Manual de calidad', archivo: 'OE-F08 Manual de calidad.docx', proceso: 'OE', destacado: true },
  { codigo: 'OE-M02', nombre: 'Política de administración y gestión del riesgo', archivo: 'OE-M02 Politica de administracion gestión del riesgo.docx', proceso: 'OE' },
  { codigo: 'OE-M03', nombre: 'Manual de uso vehicular ACTIVA', archivo: 'OE-M03 Manual de Uso Vehicular Activa.docx', proceso: 'OE' },
  { codigo: 'OE-M04', nombre: 'Plan de Manejo Integral de Residuos Sólidos (PMIRS)', archivo: 'OE-M04 Plan de Manejo Integral de Residuos Sólidos (PMIRS).docx', proceso: 'OE' },
  { codigo: 'OE-M05', nombre: 'Manual de manejo de la crisis', archivo: 'OE-M05 Manual Manejo de la Crisis.docx', proceso: 'OE' },
  { codigo: 'OE-M06', nombre: 'Protocolo de vocería institucional', archivo: 'OE-M06 Protocolo vocería institucional.docx', proceso: 'OE' },
  { codigo: 'OE-M08', nombre: 'Procedimiento para publicación en página web', archivo: 'OE-M08 Procedimiento para publicacón en página web.docx', proceso: 'OE' },
  { codigo: 'OE-M09', nombre: 'Política de comunicaciones', archivo: 'OE-M09 Política de Comunicaciones.docx', proceso: 'OE' },
  { codigo: 'OE-P01', nombre: 'Procedimiento de atención vía telefónica', archivo: 'OE-P01 Procedimiento atención  vía telefónica.docx', proceso: 'OE' },
  { codigo: 'OE-P02', nombre: 'Procedimiento de atención a personas con discapacidad', archivo: 'OE-P02 Procedimiento atención  personas con discapacidad.docx', proceso: 'OE' },
  { codigo: '', nombre: 'Caracterización de grupos de valor', archivo: 'CARACTERIZACIÓN GRUPOS DE VALOR.docx', proceso: 'OE' },
  // — Comunicaciones (sin carpeta propia en el sitio: viven con Orientación Estratégica)
  { codigo: 'RC-P01', nombre: 'Procedimiento de comunicaciones', archivo: 'RC-P01 Procedimiento comunicaciones.docx', proceso: 'OE', destacado: true },
  { codigo: '', nombre: 'Términos y condiciones página web ACTIVA', archivo: 'Términos y condiciones página web ACTIVA.docx', proceso: 'OE' },
  // — Gestión Comercial
  { codigo: 'POL-I01', nombre: 'Instructivo Comité de Asignaciones', archivo: 'POL-I01 Comité de Asignaciones.docx', proceso: 'POL' },
  { codigo: 'POL-I02', nombre: 'Políticas comerciales', archivo: 'POL-I02 Políticas comerciales.docx', proceso: 'POL' },
  { codigo: 'POL-P01', nombre: 'Procedimiento de ejecución de contratos de clientes', archivo: 'POL-P01 Procedimiento ejecucion de contratos clientes.docx', proceso: 'POL', destacado: true },
  // — Eventos propios: operación misional. Auditorías: evaluación (EV = Evaluación
  //   independiente en la caracterización EV-C01 de Control Interno).
  { codigo: 'EVEN-P01', nombre: 'Procedimiento de eventos propios', archivo: 'EVEN-P01 Procedimiento Eventos Propios.docx', proceso: 'GOP', destacado: true },
  { codigo: 'EV-P02', nombre: 'Procedimiento de auditorías internas', archivo: 'EV-P02 Procedimiento Auditorías Internas.docx', proceso: 'CI' },
  // — Operaciones
  { codigo: 'GOP-I05', nombre: 'Instructivo Comité de Asignación', archivo: 'GOP-I05 Instructivo Comité de Asignación.docx', proceso: 'GOP' },
  { codigo: 'GOP-P10', nombre: 'Soluciones para el suministro de bienes y servicios', archivo: 'GOP-P10 Soluciones para el suministro de bienes y servicios.docx', proceso: 'GOP' },
  // — Financiera
  { codigo: 'GF-P01', nombre: 'Procedimiento de pagos', archivo: 'GF-P01 Procedimiento de pagos.doc', proceso: 'GF', destacado: true },
  { codigo: 'GF-P02', nombre: 'Certificado de disponibilidad presupuestal (CDP)', archivo: 'GF-P02 Procedimiento para la elaboración del certificado de disponibilidad presupuestal.docx', proceso: 'GF', destacado: true },
  { codigo: 'GF-P03', nombre: 'Registro presupuestal (RP)', archivo: 'GF-P03 Procedimiento para la elaboración del registro presupuestal (RP).docx', proceso: 'GF', destacado: true },
  { codigo: 'GF-P04', nombre: 'Procedimiento de conciliaciones bancarias', archivo: 'GF-P04 Procedimiento conciliaciones bancarias.docx', proceso: 'GF' },
  { codigo: 'GF-P05', nombre: 'Liquidación y pago de impuestos, tasas y estampillas', archivo: 'GF-P05 Procedimiento para liquidación y pago de impuestos, tasas y estampillas.docx', proceso: 'GF' },
  { codigo: 'GF-P06', nombre: 'Facturación y cobro de cartera', archivo: 'GF-P06 Procedimiento de facturacion y cobro de cartera.docx', proceso: 'GF' },
  { codigo: 'GF-P07', nombre: 'Autorización de pagos desde cuenta supervisada', archivo: 'GF-P07 Procedimiento de autorización pagos desde cuenta supervisada.docx', proceso: 'GF' },
  // — Administrativa
  { codigo: 'GA-P02', nombre: 'Viáticos y gastos de viaje (comisiones)', archivo: 'GA-P02 Procedimiento de Viáticos y Gastos de Viaje.docx', proceso: 'GA', destacado: true },
  { codigo: 'GA-P03', nombre: 'Gestión de inventarios', archivo: 'GA-P03 Procedimiento gestión de inventarios.docx', proceso: 'GA' },
  { codigo: 'GA-I01', nombre: 'Instructivo Gestión Transparente — contratos', archivo: 'GA-I01 Instructivo Gestion Transparente Contratos.docx', proceso: 'GA' },
  { codigo: 'GA-I02', nombre: 'Instructivo de reserva de vehículo', archivo: 'GA-I02 Instructivo para realización de reserva de vehículo activa.docx', proceso: 'GA' },
  { codigo: 'GA-F019', nombre: 'Política de tratamiento de datos personales', archivo: 'GA-F019 Política tramite de datos personales.docx', proceso: 'GA' },
  // — Documental
  { codigo: 'GD-P01', nombre: 'Gestión de correspondencia recibida y enviada', archivo: 'GD-P01 Procedimientos para la gestión de correspondencia recibida y enviada.docx', proceso: 'GD' },
  { codigo: 'GD-P02', nombre: 'Gestión de PQRSFD', archivo: 'GD-P02 Procedimiento para la gestión de PQRS.docx', proceso: 'GD', destacado: true },
  { codigo: 'GD-P03', nombre: 'Préstamo de expedientes', archivo: 'GD-P03 Procedimiento Préstamo de expedientes.docx', proceso: 'GD' },
  { codigo: 'GD-P04', nombre: 'Almacenamiento de expedientes', archivo: 'GD-P04 Procedimiento Almacenamiento de expedientes.docx', proceso: 'GD' },
  { codigo: 'GD-P05', nombre: 'Transferencias documentales', archivo: 'GD-P05 Procedimiento Transferencias documental.docx', proceso: 'GD' },
  { codigo: 'GD-P06', nombre: 'Custodia documental', archivo: 'GD-P06 Procedimiento Custodia documental.docx', proceso: 'GD' },
  { codigo: 'GD-P07', nombre: 'Buzón de sugerencias', archivo: 'GD-P07 Procedimiento Buzón de Sugerencias.docx', proceso: 'GD' },
  { codigo: '', nombre: 'Plan Institucional de Archivos (PINAR)', archivo: 'Plan Institucional de Archivos.docx', proceso: 'GD' },
  // — Talento Humano
  { codigo: 'GT-P01', nombre: 'Reclutamiento, selección y vinculación de personal', archivo: 'GT-P01 Procedimiento de reclutamiento, selección y vinculación de personal.docx', proceso: 'GT', destacado: true },
  { codigo: 'GT-P02', nombre: 'Inducción, entrenamiento y reinducción', archivo: 'GT-P02 Procedimiento de inducción, entrenamiento y reinducción.docx', proceso: 'GT' },
  { codigo: 'GT-P03', nombre: 'Desvinculación de personal', archivo: 'GT-P03 Procedimiento de desvinculación de personal.docx', proceso: 'GT' },
  { codigo: 'GT-P04', nombre: 'Liquidación y pago de nómina', archivo: 'GT-P04 Procedimiento de liquidación y pago de nomina.docx', proceso: 'GT' },
  { codigo: 'GT-I01', nombre: 'Concertación de acuerdos de gestión', archivo: 'GT-I01 Instructivo para la concertación de acuerdos de gestión.docx', proceso: 'GT' },
  // — Tecnologías de la Información
  { codigo: 'GTI-P01', nombre: 'Proceso de gestión tecnológica', archivo: 'GTI-P01 Proceso de Gestión Tecnológica.docx', proceso: 'GTI' },
  { codigo: 'GTI-I02', nombre: 'Instructivo Mesa de Ayuda', archivo: 'GTI-I02 Instructivo Mesa de Ayuda.docx', proceso: 'GTI', destacado: true },
  { codigo: 'GTI-I3', nombre: 'Creación de contratos en SAFIX', archivo: 'GTI-I3 Instructivo creación de contratos en SAFIX.docx', proceso: 'GTI' },
  { codigo: 'GTI-I4', nombre: 'Aprobación de informe de evidencias — clientes', archivo: 'GTI-I4 Aprobación informe de Evidencias CLIENTES.docx', proceso: 'GTI' },
  { codigo: 'GTI-I5', nombre: 'Instalación del aplicativo CHIP local', archivo: 'GTI-I5 TUTORIAL INSTALACIÓN APLICATIVO CHIP LOCAL.docx', proceso: 'GTI' },
  { codigo: 'GTI-I6', nombre: 'Agregar un calendario desde el directorio', archivo: 'GTI-I6 INSTRUCTIVO PARA AGREGAR UN CALENDARIO DESDE EL DIRECTORIO.docx', proceso: 'GTI' },
  { codigo: 'GTI-I7', nombre: 'Reserva de carro', archivo: 'GTI-I7 INSTRUCTIVO RESERVA DE CARRO.docx', proceso: 'GTI' },
  { codigo: 'GTI-I8', nombre: 'Reserva de sala de reuniones', archivo: 'GTI-I8 INSTRUCTIVO RESERVA SALA DE REUNIONES.docx', proceso: 'GTI' },
  { codigo: 'GTI-I9', nombre: 'Ingreso a Microsoft 365', archivo: 'GTI-I9 INGRESO AL MICROSOFT 365.docx', proceso: 'GTI' },
  // — Bienes y Servicios
  { codigo: 'SBS-P05', nombre: 'Procedimiento de contratación', archivo: 'SBS-P05 Procedimiento de contratación.docx', proceso: 'BS', destacado: true },
  { codigo: 'SBS-P06', nombre: 'Pólizas de clientes', archivo: 'SBS-P06 Procedimiento pólizas clientes.docx', proceso: 'BS' },
  { codigo: 'SBS-P07', nombre: 'Liquidaciones de contratos', archivo: 'SBS-P07 Procedimiento liquidaciones de contratos.docx', proceso: 'BS' },
  { codigo: 'SBS-P08', nombre: 'Registro de proveedores', archivo: 'SBS-P08 Procedimiento Registro de Proveedores.docx', proceso: 'BS' },
  { codigo: 'SBS-P09', nombre: 'Publicación en SECOP', archivo: 'SBS-P09 Procedimiento Publicación SECOP.docx', proceso: 'BS' },
  // — Mejora Continua
  { codigo: 'MA-P01', nombre: 'Gestión de mejoras', archivo: 'MA-P01 procedimiento de Gestión de mejoras.doc', proceso: 'MA' },
  // — Disciplinarios
  { codigo: 'PD-P01', nombre: 'Instrucción de control interno disciplinario', archivo: 'PD-P01 Procedimiento Instrucción Control Interno Disciplinario.docx', proceso: 'PD' },
];

// Colores de "señal" por responsable dentro del sendero.
const ROLES = {
  solicitante: { nombre: 'Tú (solicitante)', color: '#1E6B47' },
  area:        { nombre: 'Área responsable', color: '#2C5D8A' },
  financiera:  { nombre: 'Gestión Financiera', color: '#8A5A2C' },
  documental:  { nombre: 'Gestión Documental', color: '#5B4A8A' },
  juridica:    { nombre: 'Área Jurídica', color: '#7A2C4E' },
  direccion:   { nombre: 'Dirección / Gerencia', color: '#14231B' },
  ti:          { nombre: 'Gestión TI', color: '#0E6B6B' },
};

// Aplicativo interno de Órdenes de Pedido (crea OP internas y pide CDP/RP).
const APP_OP = 'https://medellin015.github.io/Ordenes-Pedido-Activa';
// Normograma de ACTIVA (matriz de normatividad aplicable), alojado aparte.
const ENLACE_NORMOGRAMA = 'https://medellin015.github.io/Normograma-Activa/';

// Comparación compartida por las guías internas de CDP y RP (se muestra al
// final de ambas).
const COMPARACION_CDP_RP = {
  titulo: 'Diferencia clave CDP vs RP',
  items: [
    'CDP = certifica que hay presupuesto disponible (primero). Pide CPC y Fuente del Presupuesto. Su número lo asigna presupuesto.',
    'RP = compromete ese presupuesto contra un proveedor/contrato (después del CDP). Exige los datos del proveedor y del contratista; su número lo asigna presupuesto y se sincroniza solo a la OP.',
  ],
};

// Guías "¿Cómo pido…?". Los pasos resumen el procedimiento fuente; los plazos
// marcados provienen del documento oficial cuando fue posible verificarlo.
const GUIAS = [
  {
    id: 'op', pregunta: '¿Cómo hago una orden de pedido interna?',
    corta: 'Crear una OP interna (contrato ACTIVA-ACTIVA)',
    proceso: 'GA',
    resumen: 'Las OP internas son las órdenes de pedido de la propia ACTIVA. Se crean dentro del '
      + 'contrato ACTIVA-ACTIVA y su número es un consecutivo automático.',
    instructivo: {
      enlaceApp: APP_OP,
      casos: [
        {
          titulo: 'Paso 1 — Entrar al contrato',
          pasos: [
            'Inicia sesión en el aplicativo de Órdenes de Pedido con el perfil que maneja las OP internas (ACTIVA / OPs Internas) o tu perfil de ejecutivo. (Un perfil Directivo solo ve, no puede crear.)',
            'En el listado de contratos, abre ACTIVA-ACTIVA.',
            'Pulsa el botón verde «Nueva Orden» (arriba a la derecha del listado).',
          ],
        },
        {
          titulo: 'Paso 2 — Diligenciar los datos de la OP',
          intro: 'En el formulario («Nueva Orden de Pedido»), los campos con asterisco rojo * son obligatorios para guardar:',
          tabla: {
            columnas: ['Campo', 'Nota'],
            filas: [
              ['AÑO', 'Vigencia (p. ej. 2026).'],
              ['Fecha orden pedido', 'Viene con la fecha de hoy por defecto; puedes cambiarla.'],
              ['Número orden de pedido', '🔒 Automático — muestra el siguiente consecutivo (candado). El definitivo se asigna al guardar.'],
              ['Nombre/Objeto ⭐', 'Se guarda en MAYÚSCULAS.'],
              ['Fecha inicio evento', 'Fecha de ejecución (dispara la alerta «Sin RP» si está a ≤ 8 días y aún no hay RP).'],
              ['Lugar del evento', '—'],
              ['Área responsable del evento ⭐', 'Al elegirla, se autocompleta el Cargo correspondiente.'],
              ['Contrato proveedor / Nombre proveedor / Nit proveedor', 'El NIT va solo en números (sin dígito de verificación). Si el NIT ya existe, el Nombre proveedor se autocompleta.'],
              ['Valor aprobado', 'Monto en pesos (con separador de miles automático).'],
              ['CDP / RP', 'Déjalos vacíos (ver Paso 4).'],
              ['Descripción/Actividades', '—'],
              ['Cantidad ⭐', '—'],
              ['Nombre del solicitante ⭐', '—'],
              ['Correo del solicitante ⭐', 'Debe terminar en @activa.com.co (la app lo valida).'],
            ],
          },
          pie: '⭐ = obligatorio para guardar.',
        },
        {
          titulo: 'Paso 3 — Datos del contratista (para el RP)',
          intro: 'Más abajo hay un bloque aparte con los datos que alimentan el RP: Supervisor · Cédula del supervisor · Correo del supervisor · Modalidad · Fecha de inicio contrato · Fecha de terminación contrato (y Nombre + Cédula del representante legal solo si la Modalidad no es «Prestación de Servicios - Persona Natural»).',
          nota: 'Estos campos no bloquean guardar la OP: se exigen después, al momento de «Pedir RP». Conviene diligenciarlos ya para no tener que volver.',
        },
        {
          titulo: 'Paso 4 — Dejar CDP y RP vacíos',
          intro: 'El CDP y el RP no se escriben a mano aquí: se piden aparte (botones «Pedir CDP» y «Pedir RP») y el número lo asigna presupuesto en el Aplicativo RP. El RP se llena solo en la columna cuando queda asignado.',
          aviso: 'No pongas el Valor aprobado en CDP/RP. (Si por algún motivo esas columnas quedaran con el mismo valor del «Valor aprobado», la app ahora las limpia sola al abrir el contrato.)',
        },
        {
          titulo: 'Paso 5 — Guardar',
          intro: 'Pulsa «Guardar». La OP queda creada y se le asigna su número definitivo de forma atómica (dos personas creando a la vez reciben números distintos). La verás en el Listado de Órdenes del contrato.',
        },
        {
          titulo: 'Paso 6 — Qué sigue',
          intro: 'Con la OP ya creada, el flujo presupuestal es:',
          pasos: [
            'Pedir CDP (botón «Pedir CDP» → N° de OP + CPC + Fuente del Presupuesto).',
            'Pedir RP (botón «Pedir RP» → N° de OP; ahí se exigen los datos del contratista y, si el NIT es nuevo, el RUT en PDF).',
          ],
          pie: 'Ver los instructivos de «¿Cómo pido un CDP Interno?» y «¿Cómo pido un RP interno?».',
        },
      ],
      resumenFinal: 'ACTIVA-ACTIVA → Nueva Orden → llenar Objeto, Área, Valor, Cantidad, Solicitante + correo @activa.com.co (y datos del contratista) → dejar CDP/RP vacíos → Guardar → luego Pedir CDP y Pedir RP.',
    },
    docs: ['GF-P02', 'GF-P03', 'GF-P01', 'GA-P03'],
  },
  {
    id: 'cdp', pregunta: '¿Cómo pido un CDP Interno?',
    corta: 'Certificado de disponibilidad presupuestal interno (contrato ACTIVA-ACTIVA)',
    proceso: 'GF', fuente: 'GF-P02',
    resumen: 'El CDP interno se pide desde el contrato ACTIVA-ACTIVA. Hay dos caminos según si el '
      + 'gasto ya tiene una Orden de Pedido (OP) creada o no.',
    instructivo: {
      enlaceApp: APP_OP,
      casos: [
        {
          titulo: 'Caso A — CDP CON OP (la OP ya existe)',
          antes: 'La OP debe estar creada y tener diligenciados Nombre/Objeto, Valor aprobado y Correo del solicitante.',
          pasos: [
            { texto: 'Entra al contrato ACTIVA-ACTIVA.' },
            { texto: 'En la barra de acciones (arriba del listado) pulsa «Pedir CDP» (ícono ✔️).' },
            { texto: 'En la ventana escribe:', sub: [
              'Número de la orden de pedido (el N° de la OP).',
              'CPC (código del producto, p. ej. 96511).',
              'Fuente del Presupuesto: VOLCÁN DE LODO o ACTIVA (define el Centro de costos: VOLCÁN DE LODO → CAPITALIZACIÓN; ACTIVA → ACTIVA).',
            ] },
            { texto: 'Pulsa «Ejecutar». Si a la OP le falta Nombre/Objeto, Valor aprobado o Correo del solicitante, la app te lo avisa y no envía nada.' },
            { texto: 'Se abre «Confirmar solicitud de CDP» con el JSON que se enviará. Revísalo.' },
            { texto: 'Pulsa «Aprobar y enviar».' },
          ],
          flujo: 'Se envía al flujo «Solicitud de CDP Plataforma OP» (acción interna Pedir CDP OP interna) y queda un espejo en la colección cdp del Aplicativo RP. El número de CDP lo asigna presupuesto en el Aplicativo RP.',
        },
        {
          titulo: 'Caso B — CDP SIN OP (gasto sin orden de pedido)',
          pasos: [
            { texto: 'Entra al contrato ACTIVA-ACTIVA.' },
            { texto: 'Pulsa «Solicitar CDP sin OP» (ícono ➕).' },
            { texto: 'Diligencia el formulario:', sub: [
              'Nombre/Objeto (obligatorio).',
              'Contrato proveedor (obligatorio; en el JSON viaja como «contrato cliente»).',
              'Valor aprobado (obligatorio; en el JSON viaja como «valor a administrar»).',
              'CPC (obligatorio, p. ej. 96511).',
              'Fuente del Presupuesto (VOLCÁN DE LODO / ACTIVA).',
              'Correo del solicitante (debe terminar en @activa.com.co).',
            ] },
            { texto: 'Pulsa «Continuar» → se abre la confirmación con el JSON.' },
            { texto: 'Pulsa «Aprobar y enviar».' },
          ],
          flujo: 'Mismo flujo de CDP; la acción viaja como «Pedir CDP sin OP» y el espejo queda con op: "SIN OP".',
        },
      ],
      comparacion: COMPARACION_CDP_RP,
    },
    docs: ['GF-P02', 'GF-P03', 'BS-P05'],
  },
  {
    id: 'rp', pregunta: '¿Cómo pido un RP interno?',
    corta: 'Registro presupuestal interno (contrato ACTIVA-ACTIVA)',
    proceso: 'GF', fuente: 'GF-P03',
    resumen: 'El RP interno se pide desde una OP del contrato ACTIVA-ACTIVA. Es el flujo principal; '
      + 'también existe la variante «sin OP».',
    instructivo: {
      enlaceApp: APP_OP,
      casos: [
        {
          titulo: 'Caso A — RP CON OP (lo habitual)',
          antes: 'La OP debe estar completa. Al pedir el RP la app exige:',
          antesSub: [
            'Datos del RP: Valor aprobado, Contrato proveedor, NIT proveedor, Nombre proveedor, Nombre/Objeto (evento), Correo del solicitante, Fecha evento y CDP.',
            'Datos del contratista (bloque «Datos del contratista (para el RP)»): Supervisor, Cédula del supervisor, Correo del supervisor, Modalidad, Fecha de inicio y Fecha de terminación del contrato.',
            'Si la Modalidad ≠ «Prestación de Servicios - Persona Natural», además: Nombre y Cédula del representante legal.',
          ],
          pasos: [
            { texto: 'Entra al contrato ACTIVA-ACTIVA y verifica/completa esos campos en la OP (con «Editar Orden»).' },
            { texto: 'En la barra de acciones pulsa «Pedir RP».' },
            { texto: 'Escribe el número de la orden de pedido y pulsa «Ejecutar». Si falta algún campo obligatorio, la app lista exactamente cuál falta y no envía.' },
            { texto: 'Se abre «Confirmar antes de enviar» con toda la información del RP. Revísala.',
              aviso: 'Si el NIT del proveedor no está registrado, la ventana pide anexar el RUT (PDF) — es obligatorio para continuar (así queda inscrito el proveedor).' },
            { texto: 'Pulsa «Aceptar y enviar».' },
          ],
          flujo: 'Se disparan dos envíos: el RP (flujo «Solicitud de RP Plataforma OP») y, en paralelo, el registro del proveedor/contratista (flujo «Creación de contratos en bases de datos Proveedores», con el RUT si el NIT era nuevo). El número de RP lo asigna presupuesto en el Aplicativo RP y se llena solo en la columna RP de la OP la próxima vez que abras el contrato.',
          nota: 'La alerta roja «Sin RP» aparece cuando una OP tiene la fecha de ejecución en ≤ 8 días y aún no tiene número de RP — es el recordatorio para pedirlo.',
        },
        {
          titulo: 'Caso B — RP SIN OP',
          pasos: [
            { texto: 'Entra al contrato ACTIVA-ACTIVA y pulsa «Solicitar RP sin OP» (ícono ✈️).' },
            { texto: 'Diligencia:', sub: [
              'Descripción (evento).',
              'Contrato proveedor.',
              'Nombre proveedor.',
              'NIT proveedor (solo números, sin dígito de verificación).',
              'Valor aprobado.',
              'CDP.',
              'Fecha evento.',
              'Correo del solicitante (@activa.com.co).',
            ] },
            { texto: 'Pulsa «Continuar» → confirmación → «Aceptar y enviar».' },
          ],
          flujo: 'Viaja como «Pedir RP sin OP» (con numeroOrdenPedido: "SIN OP"). A diferencia del RP con OP, no anexa RUT ni registra proveedor.',
        },
      ],
      comparacion: COMPARACION_CDP_RP,
    },
    docs: ['GF-P03', 'GF-P02', 'GF-P01'],
  },
  {
    id: 'comision', pregunta: '¿Cómo pido una comisión?',
    corta: 'Viáticos y gastos de viaje',
    proceso: 'GA', fuente: 'GA-P02',
    resumen: 'Las comisiones de servicio y sus viáticos se tramitan antes del desplazamiento y '
      + 'se legalizan al regresar. Los valores y plazos exactos están en GA-P02.',
    pasos: [
      { tipo: 'inicio', titulo: 'Surge el desplazamiento', detalle: 'Una actividad institucional fuera de la sede requiere tu presencia. Diligencia el formulario de solicitud de desplazamiento, disponible también en la intranet.', rol: 'solicitante',
        enlace: { texto: 'Formulario de solicitud de desplazamiento (intranet)', url: 'https://forms.cloud.microsoft/pages/responsepage.aspx?id=q-KC6eoWEUGz3_OlN_jXFgwJ-8ds4nhNge4O1yk6VatUNldTM1MyUVdEVkJNNFFHSkpWWFNYUkdERS4u&route=shorturl' } },
      { tipo: 'paso', titulo: 'Solicita la comisión con anticipación', detalle: 'Diligencia la solicitud con destino, fechas, objeto de la comisión y transporte requerido, antes de viajar.', rol: 'solicitante' },
      { tipo: 'paso', titulo: 'Autorización del superior', detalle: 'Tu jefe inmediato o la Dirección avala la pertinencia y la disponibilidad para el desplazamiento.', rol: 'direccion' },
      { tipo: 'paso', titulo: 'Cumple la comisión', detalle: 'Realiza la actividad y conserva soportes: pasabordos, certificaciones de asistencia y demás evidencias.', rol: 'solicitante' },
      { tipo: 'decision', titulo: '¿Legalizaste a tiempo?', detalle: 'Al regresar, presenta el informe y los soportes dentro del plazo definido en GA-P02.', rol: 'solicitante',
        si: 'Comisión cerrada sin novedades', no: 'No podrás tramitar nuevas comisiones hasta legalizar la pendiente' },
      { tipo: 'paso', titulo: 'Comisión legalizada', detalle: 'Con el informe y los soportes aprobados, la comisión queda legalizada en el expediente.', rol: 'financiera' },
      { tipo: 'fin', titulo: 'Liquidación de viáticos', detalle: 'Gestión Administrativa y Financiera liquidan los viáticos y gastos de viaje según la tabla vigente y gestionan el pago, con lo que se cierra la comisión.', rol: 'financiera' },
    ],
    docs: ['GA-P02', 'GF-P01'],
    formatos: [
      { codigo: 'GA-F015', nombre: 'Certificado de cumplimiento de actividades extramural',
        url: enlaceCarpeta('Gestión Administrativa', 'Formatos') + '/' + encodeURIComponent('GA-F015 Certificado de Cumplimiento.docx') },
    ],
  },
  {
    id: 'certificado', pregunta: '¿Cómo solicito un certificado laboral?',
    corta: 'Certificado laboral (formulario en línea)',
    proceso: 'GT',
    resumen: 'Para solicitar tu certificado laboral solo tienes que diligenciar el formulario en '
      + 'línea. No hay que hacer nada más.',
    instructivo: {
      enlace: { texto: 'Diligenciar el formulario de certificado laboral', url: 'https://forms.cloud.microsoft/pages/responsepage.aspx?id=q-KC6eoWEUGz3_OlN_jXFgwJ-8ds4nhNge4O1yk6VatUNTZWTFhNUTFVRDZNMEk5TDlUQUlNMjk0WS4u&route=shorturl' },
    },
  },
];

// Mapa de riesgos. Fuente: "Seguimiento a la Gestión del Riesgo — Vigencia
// 2025" (matriz actualizada al 25/02/2026, Informe CI 2022-2025 Rad.
// 202503000143), elaborado en cumplimiento de la política OE-M02.
// Cada riesgo: [código, proceso, descripción, zona inherente, zona residual,
// tratamiento]. Las zonas salen de P×I (escala de la política OE-M02).
const ZONAS_RIESGO = {
  'BAJO':          { color: '#1E6B47', texto: '#FFFFFF', rango: '1–4' },
  'MEDIO':         { color: '#C9A227', texto: '#14231B', rango: '5–30' },
  'SIGNIFICATIVO': { color: '#E07B2A', texto: '#FFFFFF', rango: '31–60' },
  'ALTO':          { color: '#B5432E', texto: '#FFFFFF', rango: '61–80' },
  'CRÍTICO':       { color: '#7A1F1F', texto: '#FFFFFF', rango: '81–125' },
  'N/D':           { color: '#DCE5DC', texto: '#5b6b5f', rango: '' },
};

const R = (id, proc, desc, zi, zr, trat) => ({ id, proc, desc, zi, zr, trat });
const COMPONENTES_RIESGO = [
  { clave: 'A', nombre: 'Corrupción', fuente: 'Matriz oficial · C01–C10', riesgos: [
    R('C01', 'Gestión Financiera', 'Uso indebido de recursos públicos (sobrecostos, pagos indebidos)', 'SIGNIFICATIVO', 'BAJO', 'Reducir'),
    R('C02', 'Acomp. Jurídico', 'Interpretación subjetiva o manipulada de normas para favorecer intereses particulares', 'MEDIO', 'BAJO', 'Reducir'),
    R('C03', 'Gestión Financiera', 'Evasión fiscal (omisión de retenciones, declaraciones incorrectas)', 'SIGNIFICATIVO', 'BAJO', 'Reducir'),
    R('C04', 'Bienes y servicios', 'Direccionamiento de cotizaciones hacia proveedores específicos', 'ALTO', 'BAJO', 'Reducir'),
    R('C05', 'Bienes y servicios', 'Omisión o manipulación de comités de contratación', 'SIGNIFICATIVO', 'N/D', 'Reducir'),
    R('C06', 'Bienes y servicios', 'Conflictos de interés no declarados en procesos de contratación', 'SIGNIFICATIVO', 'MEDIO', 'Reducir'),
    R('C07', 'Bienes y servicios', 'Manipulación de información presentada a clientes externos en cuadros comparativos', 'MEDIO', 'BAJO', 'Reducir'),
    R('C08', 'Bienes y servicios', 'Validación insuficiente o inexistente de proveedores no marco', 'SIGNIFICATIVO', 'MEDIO', 'Reducir'),
    R('C09', 'Bienes y servicios', 'Favorecimiento indebido mediante omisión de contratos marco', 'MEDIO', 'MEDIO', 'Reducir'),
    R('C10', 'Orient. Estratégica', 'Tráfico de influencias en asignación de contratos o decisiones institucionales', 'SIGNIFICATIVO', 'MEDIO', 'Reducir'),
  ]},
  { clave: 'B', nombre: 'Transparencia', fuente: 'Matriz oficial · T01–T06', riesgos: [
    R('T01', 'Relac. Corporativo', 'Gestión deficiente de datos abiertos (ITA 2024: 6,1%)', 'MEDIO', 'BAJO', 'Reducir'),
    R('T02', 'Relac. Corporativo', 'Opacidad en información de ejecución contractual publicada en SECOP', 'MEDIO', 'BAJO', 'Reducir'),
    R('T03', 'Orient. Estratégica', 'Incumplimiento de obligaciones de transparencia activa (ITA global 59/100)', 'SIGNIFICATIVO', 'BAJO', 'Reducir'),
    R('T04', 'Planeación', 'Ausencia de Plan de Acción institucional con la estructura requerida por ITA', 'SIGNIFICATIVO', 'BAJO', 'Reducir'),
    R('T05', 'Relac. Corporativo', 'Inaccesibilidad del sitio web para personas con discapacidad (WCAG 2.1)', 'BAJO', 'BAJO', 'Reducir'),
    R('T06', 'Relac. Corporativo', 'Publicación deficiente de información de la entidad (estructura, directorio, hojas de vida)', 'BAJO', 'BAJO', 'Reducir'),
  ]},
  { clave: 'C', nombre: 'Supervisión y Ejecución', fuente: 'Matriz oficial · S01–S03', riesgos: [
    R('S01', 'Bienes y servicios', 'Entrega de carpeta de supervisión incompleta o tardía al supervisor designado', 'ALTO', 'MEDIO', 'Reducir'),
    R('S02', 'Bienes y servicios', 'Supervisor designado sin perfil técnico adecuado para el objeto contractual', 'SIGNIFICATIVO', 'MEDIO', 'Reducir'),
    R('S03', 'Bienes y servicios', 'Incumplimiento de obligaciones contractuales no detectado oportunamente', 'ALTO', 'SIGNIFICATIVO', 'Reducir'),
  ]},
  { clave: 'D', nombre: 'Complementarios', fuente: 'Matriz oficial · X01–X06', riesgos: [
    R('X01', 'Gestión Presupuestal', 'CDP emitido antes de la aprobación del cliente externo (compromiso prematuro)', 'SIGNIFICATIVO', 'SIGNIFICATIVO', 'Reducir'),
    R('X02', 'Op. Logística', 'Órdenes internas (POL-F10) sin aprobación formal del director de área', 'BAJO', 'BAJO', 'Reducir'),
    R('X03', 'Exp. del Cliente', 'Contexto estratégico insuficiente en órdenes externas que genera soluciones inadecuadas', 'SIGNIFICATIVO', 'BAJO', 'Reducir'),
    R('X04', 'Exp. del Cliente', 'Insatisfacción del cliente externo por falta de validación post-ejecución', 'BAJO', 'BAJO', 'Reducir'),
    R('X05', 'Bienes y servicios', 'Estudios previos elaborados después de la decisión de asignación', 'BAJO', 'BAJO', 'Reducir'),
    R('X06', 'Bienes y servicios', 'Falta de re-evaluación anual de proveedores marco', 'BAJO', 'BAJO', 'Reducir'),
  ]},
  { clave: 'E', nombre: 'Estratégicos, Financieros y Tecnológicos', fuente: 'GIR · M01–M08', riesgos: [
    R('M01', 'Planeación y Gestión', 'Incumplimiento de objetivos estratégicos del Plan 2026-2029', 'ALTO', 'SIGNIFICATIVO', 'Reducir'),
    R('M02', 'Gestión Financiera', 'Insuficiencia de ingresos propios / EBITDA negativo sostenido (margen neto -28%)', 'ALTO', 'SIGNIFICATIVO', 'Reducir / Evitar'),
    R('M03', 'Talento Humano', 'Pérdida de conocimiento institucional por alta rotación de contratistas', 'MEDIO', 'BAJO', 'Reducir'),
    R('M04', 'Gestión de Eventos', 'Afectación de imagen institucional por fallas operativas en eventos de alto perfil', 'SIGNIFICATIVO', 'MEDIO', 'Reducir'),
    R('M05', 'Tecnología e Info.', 'Vulnerabilidad de seguridad digital: ERP SAFIX, Microsoft 365 y activos críticos', 'SIGNIFICATIVO', 'MEDIO', 'Reducir'),
    R('M06', 'Gestión de Proyectos', 'Desfinanciamiento o sobrecostos en obras del Volcán de Lodo El Totumo', 'ALTO', 'ALTO', 'Reducir / Compartir'),
    R('M07', 'Gestión de Eventos', 'Eventos climáticos extremos que afecten la operación logística al aire libre', 'MEDIO', 'BAJO', 'Asumir / Reducir'),
    R('M08', 'Control Interno', 'Insuficiencia de capacidad técnica en Control Interno (perfil financiero requerido por Contraloría)', 'SIGNIFICATIVO', 'MEDIO', 'Reducir'),
  ]},
];
const TODOS_RIESGOS = COMPONENTES_RIESGO.flatMap((c) => c.riesgos);

// KRI 2026: [código, indicador, verde, amarillo, rojo, frecuencia/responsable]
const KRIS = [
  ['KRI01', 'Margen neto mensual de operaciones propias', '> 0%', '-5% a 0%', '< -5%', 'Mensual · Dir. Administrativo'],
  ['KRI02', 'Cumplimiento de indicadores estratégicos Plan 2026-2029', '> 80%', '60–80%', '< 60%', 'Trimestral · Planeación'],
  ['KRI03', 'Contratos interadministrativos con alertas de incumplimiento', '< 5%', '5–15%', '> 15%', 'Mensual · Dir. Jurídico'],
  ['KRI04', 'Cuentas por pagar (CXP) sin soporte documental completo', '< 5%', '5–20%', '> 20%', 'Mensual · Jefe CI'],
  ['KRI05', 'Hallazgos ITA 2024 cerrados o en plan de mejoramiento', '> 80%', '50–80%', '< 50%', 'Trimestral · Profesional TIC'],
  ['KRI06', 'Incidentes de seguridad digital registrados y reportados', '0', '1–2', '> 2', 'Mensual · Profesional TIC'],
  ['KRI07', 'Rotación de contratistas de honorarios clave', '< 10%', '10–25%', '> 25%', 'Anual · Subg. Comercial'],
  ['KRI08', 'Avance físico obras Volcán de Lodo vs cronograma', '> 90%', '70–90%', '< 70%', 'Mensual · Dir. Administrativo'],
  ['KRI09', 'Respuesta oportuna a PQRSDF dentro de términos legales', '> 95%', '80–95%', '< 80%', 'Mensual · Resp. PQRSDF'],
  ['KRI10', 'Obligaciones de Contraloría cumplidas a tiempo (Gestión Transparente)', '100%', '80–99%', '< 80%', 'Trimestral · Jefe CI'],
  ['KRI11', 'Reuniones del Comité de Saneamiento Contable en el año', '2 o más', '1 reunión', '0 reuniones', 'Semestral · Dir. Administrativo'],
];

// Indicadores del Plan Estratégico 2025-2027, transcritos del documento oficial
// "Indicadores para Plan Estratégico" (sitio Plan Estratégico en SharePoint).
const ENLACE_INDICADORES_PLAN = encodeURI('https://activaparquesyeventos.sharepoint.com/sites/PlanEstratgico/Documentos compartidos/Indicadores para Plan Estratégico.pdf');
const LINEAS_INDICADORES = [
  {
    linea: 'Línea Administración de Parques', lema: 'Espacios que conectan',
    indicadores: [
      ['Número de Parques en Funcionamiento', '1', '2', '3'],
    ],
  },
  {
    linea: 'Línea Eventos Propios', lema: 'Experiencias que trascienden',
    indicadores: [
      ['Realizar Eventos Propios durante la Anualidad', '0', '1', '2'],
    ],
  },
  {
    linea: 'Línea Operación Logística', lema: 'Soluciones integrales para el desarrollo',
    indicadores: [
      ['Nivel de Satisfacción del cliente medida a través de encuesta', '80%', '85%', '90%'],
      ['Cumplimiento del presupuesto de ventas', '100%', '100%', '100%'],
      ['Cumplimiento del presupuesto de honorarios', '100%', '100%', '100%'],
    ],
  },
  {
    linea: 'Línea Transversal', lema: 'Impulsando nuestro crecimiento',
    indicadores: [
      ['Responder el 100% de los requerimientos de Información recibidos en los términos de ley', '100%', '100%', '100%'],
      ['Incremento del EBITDA', '> 0', 'Aumento de 10% con respecto a 2025', 'Aumento de 10% con respecto a 2026'],
      ['Margen Neto de Utilidad', '> 0', 'Aumento de 5% con respecto a 2025', 'Aumento de 10% con respecto a 2026'],
      ['Automatizar los procesos de la empresa', '20% de los procesos automatizados', '30% de los procesos automatizados', '50% de los procesos automatizados'],
      ['Porcentaje de implementación de ERP', '90%', '100%', '100%'],
      ['Actualizar manuales y procedimientos de la Entidad', '30%', '50%', '90%'],
      ['Porcentaje Implementación de Programa de Facturación Electrónica', '100%', '100%', '100%'],
      ['Porcentaje de implementación de la Estrategia Cero Papel', '90%', '100%', '100%'],
      ['Tableros de Power BI Implementados', '3', '3', '4'],
      ['Índice de Desempeño Institucional', '80%', '85%', '90%'],
      ['Ejecución Plan Anual de Auditoría', '100%', '100%', '100%'],
    ],
  },
];

// Informe fuente de esta sección: documento oficial en el sitio SGC (carpeta
// "Documentos del SGC"), no en "Documentos Word", por eso el enlace se arma aparte.
const ENLACE_SEGUIMIENTO_RIESGOS =
  enlaceCarpeta('Documentos del SGC') + '/' + encodeURIComponent('Seguimiento a Mapa de Riesgos.pdf');

// Caracterizaciones de proceso. Las «oficial» están transcritas de los
// archivos *-C01 del repositorio (corte 19/07/2026); las «propuesta» son
// borradores para validación del líder del proceso (GOP, GTI y GP no tienen
// caracterización en el sitio). Cada fila PHVA: [ciclo, proveedor, entrada,
// actividad, salida, receptor].
const CARACTERIZACIONES = {
  OE: {
    codigo: 'OE-C01', version: '01', estado: 'oficial',
    archivo: ['Orientación Estratégica', 'OE-C01 Caracterización Orientación Estratégica.xlsx'],
    proceso: 'Orientación estratégica', lider: 'Gerente General',
    objetivo: 'Asegurar el logro de metas de la compañía logrando el equilibrio financiero, el cumplimiento de requisitos y la satisfacción de los clientes.',
    alcance: 'El proceso va desde la recolección de expectativas de la junta y los clientes hasta la gestión de cambios en la plataforma estratégica y el sistema de gestión.',
    resultado: 'Satisfacción del cliente mediante el logro de metas, equilibrio financiero y el cumplimiento de requisitos — Receptores: todos los procesos, Junta Directiva, proveedores, familia ACTIVA, clientes y cooperadores.',
    phva: [
      ['P', 'Junta Directiva; clientes; proveedores; colaboradores; partners', 'Normatividad aplicable; necesidades y expectativas de partes interesadas; información del contexto y del sector económico', 'Planear estratégicamente: análisis del contexto, definición de la apuesta estratégica, análisis de partes interesadas, alineamiento con procesos, planeación de la implementación y definición de sistemas de medición', 'Estrategia ACTIVA; estructura de procesos', 'Todos los procesos; partes interesadas'],
      ['P', 'Estado colombiano; entes certificadores; todos los procesos', 'Normatividad aplicable; normas técnicas; solicitudes de creación, modificación y eliminación de documentos', 'Documentar políticas y lineamientos: identificación, creación, comunicación, modificación y eliminación de documentos', 'Políticas, instrucciones y formatos', 'Todos los procesos'],
      ['H', 'Estado colombiano; entes certificadores; todos los procesos', 'Normatividad aplicable; normas técnicas; solicitudes de implementación de cambios', 'Gestionar los cambios que surgen: detección, evaluación de impacto/beneficio, acciones de conformidad, recursos, implementación y seguimiento', 'Planes de gestión del cambio cerrados', 'Todos los procesos'],
      ['H', 'Todos los procesos', 'Planes estratégicos o su equivalente; caracterizaciones de los procesos; informes de evaluación y de auditoría', 'Gestionar los riesgos de la organización: identificación, análisis y evaluación de riesgos', 'Matriz de riesgos; mapa de riesgos', 'Todos los procesos'],
      ['V', 'Todos los procesos', 'Informes de gestión de los procesos; registros', 'Hacer seguimiento a planes de ACTIVA', 'Informes de gestión', 'Orientación Estratégica; partes interesadas'],
      ['V', 'Todos los procesos', 'Hojas de vida de indicadores; registros del proceso', 'Medir indicadores', 'Hojas de vida de indicadores; informes de gestión', 'Todos los procesos; partes interesadas'],
      ['A', 'Orientación Estratégica', 'Informes de gestión; hojas de vida de indicadores; matriz de riesgos', 'Gestionar mejoras al proceso y el sistema', 'Planes de mejoramiento', 'Gestión Administrativa; Evaluación Independiente'],
    ],
    indicadores: ['Índice de Desempeño Institucional (Plan Estratégico)', 'KRI02 — Cumplimiento de indicadores estratégicos del Plan 2026-2029'],
    riesgos: ['C10 — Tráfico de influencias en asignación de contratos o decisiones institucionales', 'T01 — Gestión deficiente de datos abiertos', 'T02 — Opacidad en información de ejecución contractual publicada en SECOP', 'T03 — Incumplimiento de obligaciones de transparencia activa', 'T04 — Ausencia de Plan de Acción institucional con estructura ITA', 'T05 — Inaccesibilidad del sitio web para personas con discapacidad (WCAG 2.1)', 'T06 — Publicación deficiente de información de la entidad', 'M01 — Incumplimiento de objetivos estratégicos del Plan 2026-2029'],
    recursos: { humanos: 'Gerente General; Subgerente Comercial; Director Administrativo y Financiero; Jefe de Oficina de Control Interno; Director Jurídico', fisicos: 'Herramientas ofimáticas; espacios de trabajo y equipos de oficina' },
    documentos: 'Listado maestro de documentos y registros',
    requisitosISO: '4 Contexto de la organización · 5 Liderazgo · 6 Planificación · 7.3 Toma de conciencia · 7.4 Comunicación · 7.5 Información documentada · 9.1 Seguimiento, medición, análisis y evaluación · 9.3 Revisión por la dirección',
  },
  POL: {
    codigo: 'POL-C01', version: '01', estado: 'oficial',
    archivo: ['Gestión Comercial', 'POL-C01 Caracterización Gestión Comercial.xls'],
    proceso: 'Proyectos de operación logística', lider: 'Subgerente Comercial',
    objetivo: 'Ejecutar servicios de operación logística satisfaciendo los requisitos acordados con el cliente.',
    alcance: 'Desde la recepción del contrato firmado con el cliente y las órdenes de pedido hasta las actividades de monitoreo y gestión de mejoras.',
    resultado: 'Servicios de operación logística ejecutados satisfactoriamente — Receptor: el cliente.',
    phva: [
      ['P', 'Orientación Estratégica; Desarrollo de la Experiencia del Cliente', 'Estrategia ACTIVA; contrato firmado; acta de inicio firmada; registro de aceptación de contrato; orden de pedido', 'Preparar la operación: comunicación y socialización de servicios contratados, verificación de disponibilidad de insumos, aceptación de cotizaciones por parte del cliente y solicitud de compra', 'Registros de comunicación y socialización; orden de pedido aprobada; solicitud de compra; orden de pedido a proveedores', 'Desarrollo de la Experiencia del Cliente; Suministro de Bienes y Servicios; Gestión Financiera'],
      ['H', 'Proyectos de operación logística; Suministro de Bienes y Servicios', 'Orden de servicio aprobada', 'Ejecutar servicios', 'Evidencias de ejecución de contratos', 'Cliente'],
      ['V', 'Proyectos de operación logística', 'Contratos; evidencias de ejecución de contratos', 'Monitorear la ejecución de los servicios', 'Informes de supervisión con soportes y anexos; informes de gestión', 'Desarrollo de la Experiencia del Cliente; Suministro de Bienes y Servicios; Gestión Financiera; Orientación Estratégica'],
      ['V', 'Proyectos de operación logística', 'Plan de acción de mercadeo y ventas; registros del proceso', 'Medir indicadores', 'Hojas de vida de indicadores', 'Orientación Estratégica; Desarrollo de la Experiencia del Cliente; cliente'],
      ['A', 'Desarrollo de la Experiencia del Cliente; Proyectos de Operación Logística', 'Informes de gestión; hojas de vida de indicadores', 'Gestionar mejoras al proceso', 'Planes de mejoramiento', 'Desarrollo de la Experiencia del Cliente; Proyectos de Operación Logística; Evaluación Independiente'],
    ],
    indicadores: ['Cumplimiento del presupuesto de ventas (Plan Estratégico)', 'Nivel de satisfacción del cliente medido a través de encuesta (Plan Estratégico)'],
    riesgos: ['X03 — Contexto estratégico insuficiente en órdenes externas que genera soluciones inadecuadas', 'X04 — Insatisfacción del cliente externo por falta de validación post-ejecución'],
    recursos: { humanos: 'Subgerente Comercial; Gestor de Mercadeo y Ventas; Ejecutivo Comercial; Profesional de operación logística de eventos; prestadores de servicios asociados', fisicos: 'SIESA; herramientas ofimáticas; espacios de trabajo y equipos de oficina' },
    documentos: 'Listado maestro de documentos y registros',
    requisitosISO: '7.1.5 Recursos de seguimiento y medición · 8.1 Planificación y control operacional · 8.2 Requisitos para los productos y servicios · 8.4 Control de procesos, productos y servicios suministrados externamente · 8.5.1–8.5.4 Producción y provisión del servicio · 8.6 Liberación · 9.1 Seguimiento, medición, análisis y evaluación',
  },
  GOP: {
    codigo: 'GOP-C01', version: '01', estado: 'oficial',
    proceso: 'Gestión de Operaciones', lider: 'Gestor de Operaciones',
    objetivo: 'Planear y ejecutar la operación logística de eventos y los eventos propios institucionales y comerciales, garantizando el cumplimiento de los requisitos y expectativas de clientes, socios, aliados y asistentes.',
    alcance: 'Desde la recepción de la orden de pedido o la aprobación del evento propio hasta la entrega del servicio, el informe de supervisión y el cierre con la validación de satisfacción.',
    resultado: 'Eventos y operaciones logísticas ejecutados a satisfacción — Receptores: clientes, asistentes y todos los procesos.',
    phva: [
      ['P', 'Gestión Comercial; Orientación y Planeación Estratégica', 'Órdenes de pedido aprobadas; contratos; calendario de eventos', 'Planear la operación del evento o servicio: requerimientos, recursos, proveedores y cronograma (GOP-P10)', 'Plan de operación; solicitudes de cotización', 'Gestión de Operaciones; Proceso de Contratación'],
      ['H', 'Proceso de Contratación; proveedores', 'Bienes y servicios contratados; personal asignado', 'Ejecutar la operación logística y los eventos propios (EV-P01, GOP-P10)', 'Evento o servicio ejecutado; evidencias de ejecución', 'Cliente; asistentes'],
      ['V', 'Gestión de Operaciones', 'Informes de supervisión; encuestas post-servicio', 'Supervisar la ejecución y validar la satisfacción del cliente', 'Informe final de supervisión; resultado de satisfacción', 'Gestión Comercial; Gestión Financiera'],
      ['A', 'Gestión de Operaciones', 'Informes de gestión; hojas de vida de indicadores', 'Gestionar mejoras al proceso', 'Planes de mejoramiento (MA-F01)', 'Mejoramiento Continuo; Control Interno'],
    ],
    indicadores: ['Realizar eventos propios durante la anualidad (Plan Estratégico)', 'Nivel de satisfacción del cliente medido a través de encuesta (Plan Estratégico)'],
    riesgos: ['X02 — Órdenes internas (POL-F10) sin aprobación formal del director de área', 'M04 — Afectación de imagen institucional por fallas operativas en eventos de alto perfil', 'M07 — Eventos climáticos extremos que afecten la operación logística al aire libre'],
    recursos: { humanos: 'Gestor de Operaciones; profesionales y contratistas de operación logística', fisicos: 'Aplicativo de Órdenes de Pedido; herramientas ofimáticas; equipos e infraestructura de eventos' },
    documentos: 'GOP-P10 Soluciones para el suministro de bienes y servicios; EV-P01 Procedimiento de eventos propios; GOP-I05 Instructivo Comité de Asignación',
    requisitosISO: '8.1 Planificación y control operacional · 8.2 Requisitos para los productos y servicios · 8.5 Producción y provisión del servicio · 9.1.2 Satisfacción del cliente',
  },
  GP: {
    codigo: 'GP-C01', version: '01', estado: 'oficial',
    proceso: 'Gestión de Parques', lider: 'Subgerente Comercial',
    objetivo: 'Administrar los parques y espacios recreativos asignados a la entidad, asegurando la optimización de los recursos, la implementación de políticas y procedimientos operativos estándar, y la planificación y ejecución de los contratos correspondientes.',
    alcance: 'Desde la recepción o asignación del parque o espacio recreativo hasta su operación, mantenimiento, atención de visitantes y rendición de informes del contrato.',
    resultado: 'Parques en funcionamiento operados de forma segura y sostenible — Receptores: visitantes, entidades contratantes y la comunidad.',
    phva: [
      ['P', 'Orientación y Planeación Estratégica; entidades contratantes', 'Contratos o convenios de administración de parques; presupuesto; normatividad aplicable', 'Planear la operación del parque: modelo de operación, recursos, tarifas y protocolos', 'Plan de operación del parque', 'Gestión de Parques; Gestión Financiera'],
      ['H', 'Proceso de Contratación; Gestión Administrativa', 'Bienes, servicios y personal para la operación', 'Operar y mantener el parque: apertura, atención de visitantes, seguridad, aseo y mantenimiento', 'Parque en funcionamiento; servicios prestados a visitantes', 'Visitantes; entidad contratante'],
      ['V', 'Gestión de Parques', 'Registros de operación; PQRSFD; informes de supervisión', 'Hacer seguimiento a la operación y a la satisfacción de los visitantes', 'Informes de operación y de contrato', 'Entidad contratante; Orientación y Planeación Estratégica'],
      ['A', 'Gestión de Parques', 'Informes de gestión; hojas de vida de indicadores', 'Gestionar mejoras al proceso', 'Planes de mejoramiento (MA-F01)', 'Mejoramiento Continuo; Control Interno'],
    ],
    indicadores: ['Número de parques en funcionamiento (Plan Estratégico)', 'KRI08 — Avance físico de obras Volcán de Lodo vs cronograma'],
    riesgos: ['M06 — Desfinanciamiento o sobrecostos en obras del Volcán de Lodo El Totumo'],
    recursos: { humanos: '[Por definir: equipo de operación del parque]', fisicos: 'Infraestructura de los parques; herramientas ofimáticas' },
    documentos: '[Por definir: procedimientos operativos del parque — el proceso está en estructuración]',
    requisitosISO: '8.1 Planificación y control operacional · 8.5 Producción y provisión del servicio · 9.1.2 Satisfacción del cliente',
  },
  GF: {
    codigo: 'GF-C01', version: '01', estado: 'oficial',
    archivo: ['Gestión Financiera', 'GF-C01 Caracterización Gestión Financiera.xls'],
    proceso: 'Gestión Financiera', lider: 'Director(a) Administrativo(a) y Financiero(a)',
    objetivo: 'Administrar los recursos financieros de ACTIVA garantizando una ejecución presupuestal eficiente, transparente, que presente balances positivos y estados financieros favorables.',
    alcance: 'Desde la planificación de recursos y actividades financieras hasta el saneamiento contable y el mejoramiento continuo del proceso.',
    resultado: 'Ejecución presupuestal eficiente, transparente, con balances positivos y estados financieros favorables — Receptores: todos los procesos.',
    phva: [
      ['P', 'Orientación Estratégica; Junta Directiva', 'Estrategia ACTIVA; normatividad presupuestal y financiera; calendario tributario; presupuesto autorizado; solicitudes de modificación; planes de trabajo; metas financieras', 'Planear la gestión de recursos financieros: actividades de la anualidad, consolidación e incorporación del presupuesto, modificaciones, PAC, cronograma de pagos y calendario de cobros', 'Plan de trabajo anual; presupuesto general versionado; PAC; calendario de pagos y de cobros', 'Gestión Financiera; Gestión Administrativa; líderes de proceso y gestores; proveedores; contratistas; empleados'],
      ['H', 'Todos los procesos; Comité de Contratación', 'Solicitud de CDP; solicitud de RP', 'Generación de disponibilidad presupuestal y de registro presupuestal', 'CDP; RP', 'Gestión de Talentos; Suministro de Bienes y Servicios'],
      ['H', 'Gestión Financiera; proveedores; Proyectos de Operación Logística; Suministro de Bienes y Servicios', 'PAC de gastos; facturas y cuentas por pagar; actas de recibo a satisfacción; certificación de seguridad social y parafiscales; informes de supervisión', 'Pagar', 'Pagos realizados', 'Proveedores de bienes y servicios'],
      ['H', 'Desarrollo de la Experiencia del Cliente; Proyectos de Operación Logística', 'Contrato y acta de inicio; PAC de ingresos; calendario de cobros; cuentas por cobrar; evidencias de ejecución', 'Cobrar', 'Ingresos recaudados', 'Gestión Financiera'],
      ['H', 'DIAN; Gestión Financiera', 'Calendario tributario; información contable', 'Elaborar declaraciones', 'Declaraciones', 'DIAN; Gobernación de Antioquia; Contaduría General'],
      ['V', 'Estado; Gestión Financiera', 'Normatividad; información contable; informes; estados financieros; hechos económicos', 'Elaborar informes financieros', 'Informes; estados financieros', 'DIAN; Junta Directiva; Orientación Estratégica'],
      ['V', 'Gestión Financiera', 'Planes del proceso; registros', 'Hacer seguimiento a planes del proceso', 'Informes de gestión', 'Orientación Estratégica; Gestión Financiera'],
      ['V', 'Gestión Financiera', 'Registros del proceso', 'Medir indicadores', 'Hojas de vida de indicadores', 'Junta Directiva; Orientación Estratégica'],
      ['A', 'Gestión Financiera', 'Informes; estados financieros', 'Hacer saneamiento contable', 'Informe de saneamiento contable', 'Gestión Financiera; Evaluación Independiente'],
      ['A', 'Gestión Financiera', 'Informes de gestión; hojas de vida de indicadores', 'Gestión de mejoras al proceso', 'Planes de mejoramiento', 'Gestión Financiera; Evaluación Independiente'],
    ],
    indicadores: ['Incremento del EBITDA (Plan Estratégico)', 'Margen neto de utilidad (Plan Estratégico)', 'Cumplimiento del presupuesto de honorarios (Plan Estratégico)', 'Implementación del Programa de Facturación Electrónica (Plan Estratégico)', 'KRI01 — Margen neto mensual de operaciones propias', 'KRI04 — Cuentas por pagar sin soporte documental completo', 'KRI11 — Reuniones del Comité de Saneamiento Contable'],
    riesgos: ['C01 — Uso indebido de recursos públicos (sobrecostos, pagos indebidos)', 'C03 — Evasión fiscal (omisión de retenciones, declaraciones incorrectas)', 'X01 — CDP emitido antes de la aprobación del cliente externo', 'M02 — Insuficiencia de ingresos propios / EBITDA negativo sostenido'],
    recursos: { humanos: 'Director Administrativo y Financiero; Gestor Financiero; Gestor Contable; Gestor de Tesorería', fisicos: 'SIESA (gestión financiera); herramientas ofimáticas; espacios de trabajo y equipos de oficina' },
    documentos: 'Listado maestro de documentos',
    requisitosISO: '4.1 Comprensión de la organización y su contexto · 4.2 Necesidades y expectativas de las partes interesadas · 6.1 Acciones para abordar riesgos y oportunidades · 7.1 Recursos',
  },
  GA: {
    codigo: 'GA-C01', version: '01', estado: 'oficial',
    archivo: ['Gestión Administrativa', 'GA-C01 Caracterización Gestión Administrativa.xls'],
    proceso: 'Gestión administrativa', lider: 'Director(a) Administrativo(a) y Financiero(a)',
    objetivo: 'Facilitar la disponibilidad de recursos físicos, tecnológicos, de infraestructura, transporte, documentales y de ambiente de trabajo, idóneos y oportunos.',
    alcance: 'Desde la recepción de la necesidad manifiesta por los diferentes procesos hasta la verificación de la ejecución de planes, la medición de indicadores y las actividades de mejoramiento continuo.',
    resultado: 'Disponibilidad de recursos físicos, tecnológicos, de infraestructura, transporte, documentales y de ambiente de trabajo, idóneos y oportunos — Receptores: todos los procesos.',
    phva: [
      ['P', 'Todos los procesos', 'Estrategia ACTIVA; hojas de vida de bienes y equipos', 'Planear las actividades administrativas: planificación de TI, de la gestión documental y del mantenimiento preventivo y correctivo', 'Plan de desarrollo de TI; PGD y PINAR; plan de mantenimiento preventivo y correctivo', 'Gestión Administrativa; Gestión Financiera; todos los procesos'],
      ['H', 'Gestión Administrativa; todos los procesos', 'Necesidades tecnológicas manifiestas', 'Gestionar las tecnologías de la información: identificación, evaluación y selección; adaptación; capacitación; gestión de licencias y de documentación digital', 'Solicitudes de compra de tecnologías; tecnologías instaladas; licencias actualizadas; recursos funcionales; información digital auténtica, fiable, íntegra y usable', 'Gestión Administrativa; todos los procesos'],
      ['H', 'Gobierno nacional; todos los procesos', 'Directrices del AGN; normatividad documental colombiana; documentación física generada', 'Gestionar los documentos físicos: planeación, producción, gestión y trámite, organización, transferencia, disposición, preservación y valoración documental', 'Documentación física auténtica, fiable, íntegra y usable', 'Todos los procesos'],
      ['H', 'Gestión Administrativa; todos los procesos; proveedores de activos', 'Plan de mantenimiento preventivo; hojas de vida de bienes y equipos; solicitudes de mantenimiento correctivo', 'Gestionar los activos físicos y equipos: recepción, inventario, mantenimiento, administración, reemplazo y disposición', 'Activos físicos y equipos disponibles, actualizados, adecuados y funcionales', 'Todos los procesos'],
      ['H', 'Gestión Administrativa', 'Solicitudes de mantenimiento', 'Apoyar la optimización del ambiente de trabajo: mantenimiento preventivo y correctivo de infraestructura', 'Infraestructura adecuada', 'Todos los procesos'],
      ['V', 'Gestión Administrativa', 'Planes del proceso; registros', 'Hacer seguimiento a planes del proceso', 'Informes de gestión', 'Orientación Estratégica; Gestión Administrativa'],
      ['V', 'Gestión Administrativa', 'Registros del proceso', 'Medir indicadores', 'Hojas de vida de indicadores', 'Orientación Estratégica; Gestión Administrativa'],
      ['A', 'Gestión Administrativa', 'Informes de gestión; hojas de vida de indicadores', 'Gestionar mejoras al proceso', 'Planes de mejoramiento', 'Gestión Administrativa; Evaluación Independiente'],
    ],
    indicadores: ['Sin indicadores asignados a este proceso en los documentos vigentes'],
    riesgos: ['Sin riesgos asignados a este proceso en la matriz vigente (corte 25/02/2026)'],
    recursos: { humanos: 'Director Administrativo y Financiero; Gestor en Tecnología e Información; Gestor de Área Administrativa; Profesional de Suministro de Bienes y Servicios', fisicos: 'SIESA (gestión de activos); herramientas ofimáticas; espacios de trabajo y equipos de oficina' },
    documentos: 'Listado maestro de documentos y registros',
    requisitosISO: '7.1.3 Infraestructura · 7.1.4 Ambiente para la operación de los procesos',
  },
  GD: {
    codigo: 'GD-C01', version: '01', estado: 'oficial',
    archivo: ['Gestión Documental', 'GD-C01 Caracterización Programa de Gestión Documental.xlsx'],
    proceso: 'Gestión Documental', lider: 'Director(a) Administrativo(a) y Financiero(a)',
    objetivo: 'Administrar el flujo documental y establecer la organización de los documentos siguiendo los lineamientos correspondientes del Programa de Gestión Documental en torno a la normativa vigente.',
    alcance: 'Inicia con los requerimientos de gestión documental presentados por usuarios internos y externos, continúa con la radicación y reparto de los mismos y finaliza con la custodia de la documentación teniendo en cuenta las Tablas de Retención Documental y la normatividad archivística vigente.',
    resultado: 'Organización, control, seguridad, accesibilidad y conservación de los expedientes generados en la entidad — Receptores: todos los procesos.',
    phva: [
      ['P', 'Todos los procesos', 'Requerimientos de gestión documental; normatividad archivística', 'Planeación: establecer los criterios para garantizar el ciclo vital de los documentos, su acceso, control, organización y conservación mediante instrumentos archivísticos normalizados', 'Diagnóstico integral de archivos; política general de archivos; PGD; PINAR; indicadores de gestión; sistema integral de conservación; instrumentos archivísticos', 'Todos los procesos; Archivo Central'],
      ['H', 'Todos los procesos; proveedores; clientes; ciudadanía', 'Sistema de PQRS; documentos externos e internos', 'PQRS: determinación de los lineamientos para la gestión de peticiones, quejas, reclamos y sugerencias', 'Procedimiento de radicación; PQRS contestadas; tiempos de respuesta; firmas autorizadas', 'Todos los procesos; proveedores; clientes; ciudadanía'],
      ['H', 'Todas las áreas; proveedores; clientes; ciudadanía', 'Instructivo de radicación; documentos externos e internos', 'Radicación: actividades para el cumplimiento de normas del procedimiento de radicación, con su paso a paso, tiempos y planillas', 'Recepción de documentación; documentos con número de radicación; control de ingreso; planillas; firmas autorizadas', 'Todas las áreas; proveedores; clientes; ciudadanía'],
      ['H', 'Todas las áreas', 'Instructivos de organización documental', 'Almacenamiento: organización de expedientes y cajas de archivo', 'Medios y técnicas de producción; reprografía; inventarios documentales; registro de documentos', 'Archivo General'],
      ['H', 'Todas las áreas', 'Creación de registros, distribución, trámites, acceso y consultas, seguimiento y control', 'Préstamo de expedientes: formatos y soportes para el control, seguridad, distribución y recepción de expedientes', 'Control; consulta; registro de documentos; expediente prestado', 'Todas las áreas'],
      ['H', 'Todos los procesos', 'Instructivos de transferencias documentales', 'Transferencia: remisión de documentos del archivo de gestión al central y de éste al histórico, según TRD y TVD', 'Clasificación, ordenación y descripción documental; inventarios; indicadores y cronograma de transferencias; TVD', 'Todos los procesos'],
      ['V', 'Gestión Documental', 'Plan de acción anual', 'Verificación del seguimiento de acciones a corto, mediano y largo plazo conforme a la organización de los documentos y su ciclo vital', 'Indicadores de desempeño; planes de mejoramiento', 'Todos los procesos'],
      ['A', 'Gestión Documental', 'Indicadores de desempeño', 'Estipulación de planes de acción, identificando factores de riesgo y aplicando los criterios archivísticos', 'Indicadores de desempeño; planes de mejoramiento', 'Todos los procesos'],
    ],
    indicadores: ['Responder el 100% de los requerimientos de información en los términos de ley (Plan Estratégico)', 'Implementación de la Estrategia Cero Papel (Plan Estratégico)', 'KRI09 — Respuesta oportuna a PQRSDF dentro de los términos legales'],
    riesgos: ['Sin riesgos asignados a este proceso en la matriz vigente (corte 25/02/2026)'],
    recursos: { humanos: 'Líder de Archivo; Director Administrativo y Financiero; Gestor en Tecnología e Información; Gestor de Área Administrativa', fisicos: 'Área de archivo; computador' },
    documentos: 'Cronograma de planeación; planes de acción; PGD; Programa de Gestión Documental; PINAR',
    requisitosISO: '7.5.2 Creación y actualización · 7.5.3 Control de la información documentada',
  },
  GT: {
    codigo: 'GT-C01', version: '01', estado: 'oficial',
    archivo: ['Gestión del Talento Humano', 'GT-C01 Caracterización Gestión del Talento Humano.xlsx'],
    proceso: 'Gestión de Talentos', lider: 'Director(a) Administrativo(a) y Financiero(a)',
    objetivo: 'Desarrollar y asegurar PERSONAS con alto desempeño, competentes y satisfechas.',
    alcance: 'Inicia con la selección, sostenimiento y finaliza con el retiro del funcionario.',
    resultado: 'Personas con alto desempeño, competentes y satisfechas — Receptores: todos los procesos.',
    phva: [
      ['P', 'Orientación Estratégica', 'Estrategia ACTIVA; presupuesto', 'Proyección del presupuesto anual de funcionamiento del proceso y de las cuantías para el pago de nómina', 'Presupuesto del proceso', 'Gestión Financiera'],
      ['P', 'Organismos facultados; Orientación Estratégica', 'Normatividad vigente; diagnósticos de capacitación, bienestar y SST', 'Formulación de planes y programas de capacitación, bienestar y SST; gestión de riesgos laborales', 'Plan anual de SST; plan estratégico de talento humano; plan institucional de capacitación; plan social laboral e incentivos', 'Gestión Financiera; Orientación Estratégica'],
      ['P', 'Orientación Estratégica', 'Necesidades de reorganización o ajuste institucional', 'Formulación del manual de funciones con los perfiles requeridos; identificación de vacantes', 'Manual de funciones y perfiles de cargo; caracterización de funcionarios', 'Todos los procesos'],
      ['P', 'Orientación Estratégica', 'Directrices y lineamientos', 'Diseñar y ajustar la evaluación de desempeño', 'Evaluación de desempeño', 'Gestión de Talentos'],
      ['H', 'Todos los procesos', 'Necesidad de pago y novedades de nómina', 'Elaborar la nómina de personal para su trámite de pago; ajustes y reajustes', 'Liquidación de nómina; comprobantes de pago', 'Gestión Financiera; todos los procesos'],
      ['H', 'Gestión de Talentos', 'Plan anual de SST; plan estratégico de talento humano; plan institucional de capacitación', 'Ejecución de planes y programas', 'Planes y cronogramas ejecutados', 'Todos los procesos'],
      ['H', 'Todos los procesos', 'Solicitudes de personal; lineamientos; normatividad; CDP; RP', 'Selección, inducción y reinducción; trámites de ley para el nombramiento de personal', 'Personal seleccionado y entrenado; cargos provistos', 'Todos los procesos'],
      ['H', 'Todos los procesos', 'Notificación de desvinculación (carta de renuncia)', 'Realización de procesos de desvinculación', 'Personal desvinculado', 'Todos los procesos'],
      ['H', 'Gestión de Talentos', 'Evaluación de desempeño', 'Enviar evaluación e instructivo a jefes inmediatos', 'Evaluación e instructivo; resultados (informe); diagnóstico de capacitación', 'Gestión de Talentos; todos los procesos'],
      ['H', 'Todos los procesos', 'Evaluación de desempeño diligenciada', 'Analizar los resultados y definir plan de mejora individual', 'Plan de mejora individual', 'Gestión de Talentos'],
      ['V', 'Gestión de Talentos', 'Planes del proceso; registros', 'Informes y reportes de seguimiento y evaluación de la gestión (clima organizacional, bienestar y SST); planes de mejora individual; situaciones administrativas', 'Informes de gestión; notificaciones de situaciones administrativas', 'Gestión de Talentos; Acompañamiento Jurídico'],
      ['V', 'Gestión de Talentos', 'Registros del proceso', 'Medir indicadores', 'Hojas de vida de indicadores', 'Orientación Estratégica; Gestión de Talentos'],
      ['A', 'Gestión Administrativa', 'Informes de gestión; hojas de vida de indicadores', 'Gestionar mejoras al proceso', 'Planes de mejoramiento', 'Gestión Administrativa; Evaluación Independiente'],
    ],
    indicadores: ['KRI07 — Rotación de contratistas de honorarios clave'],
    riesgos: ['M03 — Pérdida de conocimiento institucional por alta rotación de contratistas'],
    recursos: { humanos: 'Gestora en Talento Humano', fisicos: 'Herramientas ofimáticas; espacios de trabajo' },
    documentos: 'Listado maestro de documentos y registros',
    requisitosISO: '7.1.2 Personas · 7.1.4 Ambiente para la operación de los procesos',
  },
  GTI: {
    codigo: 'GTI-C01', version: '01', estado: 'oficial',
    proceso: 'Gestión de Tecnologías de la Información', lider: 'Profesional TIC',
    objetivo: 'Garantizar la disponibilidad, el soporte y la seguridad de los servicios y activos tecnológicos de la entidad (ERP SAFIX, Microsoft 365 y aplicativos institucionales), gestionando los incidentes y solicitudes a través de la Mesa de Ayuda.',
    alcance: 'Desde la identificación de la necesidad o incidente tecnológico hasta su solución verificada, incluida la administración de activos de información, la seguridad digital y el soporte a los aplicativos institucionales.',
    resultado: 'Servicios tecnológicos disponibles, seguros y con soporte oportuno — Receptores: todos los procesos.',
    phva: [
      ['P', 'Orientación y Planeación Estratégica', 'Plan Estratégico; Política de Seguridad Digital V2; presupuesto TIC', 'Planear la gestión tecnológica: plan de desarrollo de TI, seguridad digital (MSPI) y adquisiciones tecnológicas (GTI-P01)', 'Plan de desarrollo de TI; plan de acción de seguridad digital', 'Todos los procesos'],
      ['H', 'Todos los procesos', 'Incidentes y solicitudes tecnológicas', 'Gestionar la Mesa de Ayuda: registro, clasificación, priorización, solución o escalamiento (GTI-I02)', 'Casos resueltos y documentados', 'Todos los procesos'],
      ['H', 'Gestión de TI', 'Activos de información; licencias; copias de respaldo', 'Administrar los activos de información, los respaldos y los controles de acceso', 'Activos administrados; copias de respaldo; controles implementados', 'Todos los procesos'],
      ['V', 'Gestión de TI', 'Registros de la Mesa de Ayuda; reportes de incidentes', 'Hacer seguimiento a los indicadores del servicio y a los incidentes de seguridad digital (KRI06)', 'Informes de gestión; reporte de incidentes', 'Control Interno; Comité de Gestión y Desempeño'],
      ['A', 'Gestión de TI', 'Informes de gestión; hallazgos ITA', 'Gestionar mejoras al proceso y al MSPI', 'Planes de mejoramiento (MA-F01)', 'Mejoramiento Continuo; Control Interno'],
    ],
    indicadores: ['Automatización de los procesos de la empresa (Plan Estratégico)', 'Implementación de ERP (Plan Estratégico)', 'Tableros de Power BI implementados (Plan Estratégico)', 'KRI05 — Hallazgos ITA 2024 cerrados o en plan de mejoramiento', 'KRI06 — Incidentes de seguridad digital registrados y reportados'],
    riesgos: ['M05 — Vulnerabilidad de seguridad digital: ERP SAFIX, Microsoft 365 y activos de información críticos'],
    recursos: { humanos: 'Profesional TIC', fisicos: 'ERP SAFIX; Microsoft 365; aplicativos institucionales; infraestructura tecnológica' },
    documentos: 'GTI-P01 Proceso de Gestión Tecnológica; GTI-I02 Instructivo Mesa de Ayuda; Política de Uso Aceptable de Recursos Tecnológicos; Política de Seguridad Digital V2',
    requisitosISO: '7.1.3 Infraestructura · 7.5 Información documentada · 8.1 Planificación y control operacional',
  },
  BS: {
    codigo: 'SBS-C01', version: '01', estado: 'oficial',
    archivo: ['Proceso de Contratación', 'SBS- C01 Caracterización Proceso de Contratación.xls'],
    proceso: 'Proceso de Contratación', lider: 'Director(a) Jurídico(a)',
    objetivo: 'Adquirir los bienes y servicios requeridos para la operación de los procesos de ACTIVA de forma oportuna, eficiente e idónea.',
    alcance: 'Desde la recepción de la necesidad de adquisición de bienes y servicios manifiesta desde los diferentes procesos hasta la verificación de la ejecución del plan anual de adquisiciones, evaluación de proveedores, la medición de indicadores y las actividades de mejoramiento continuo.',
    resultado: 'Bienes y servicios comprados de forma oportuna, eficiente e idónea — Receptores: todos los procesos.',
    phva: [
      ['P', 'Orientación Estratégica; Gestión Administrativa; Acompañamiento Jurídico', 'Estrategia ACTIVA; presupuesto general versionado; fichas técnicas de bienes y equipos; manual de contratación', 'Planear la contratación de compra de bienes y servicios', 'Plan Anual de Adquisiciones (versionado)', 'Gestión Administrativa; Gestión Financiera; todos los procesos'],
      ['P', 'Proveedores', 'Registros realizados en aplicativo', 'Crear y actualizar bases de datos de proveedores', 'Base de datos de proveedores inscritos', 'Suministro de Bienes y Servicios'],
      ['P', 'Suministro de Bienes y Servicios; Proyectos de Operación Logística; todos los procesos', 'Plan Anual de Adquisiciones; registros de comunicación y socialización; orden de pedido aprobada; solicitud de compra', 'Recibir la necesidad', 'Estudio de corredor de seguros; borradores de orden de compra y estudios previos; documentos de soporte', 'Suministro de Bienes y Servicios'],
      ['H', 'Suministro de Bienes y Servicios; Gestión Financiera', 'Borradores y documentos de soporte; CDP; acuerdos marco (cuando aplica)', 'Viabilizar la compra: estudio de mercado, topes de precio, aprobación del cliente (cuando aplica), comité de contratación y correcciones', 'Estudio de corredor aprobado; orden de compra final; estudio previo final; certificación del comité de contratación', 'Suministro de Bienes y Servicios; todos los procesos'],
      ['H', 'Suministro de Bienes y Servicios; Gestión Financiera', 'Estudio de corredor aprobado; orden de compra y estudio previo finales; certificación del comité', 'Concluir la estructuración: expedición de póliza, publicación de procesos (cuando aplica), evaluación de proponentes y adjudicación', 'RP; póliza firmada; orden de compra firmada; estudio previo firmado; expediente precontractual', 'Suministro de Bienes y Servicios; Gestión Financiera'],
      ['H', 'Suministro de Bienes y Servicios', 'RP; póliza, orden de compra y estudio previo firmados; expediente precontractual', 'Perfeccionamiento de la compra: firmas de orden, contrato, acta de inicio y póliza; publicación en SECOP; notificación de supervisión', 'Orden de compra y contrato firmados; acta de inicio; póliza; SECOP actualizado; comunicación de supervisión', 'Suministro de Bienes y Servicios; todos los procesos'],
      ['H', 'Proveedores', 'Contrato y orden de compra firmados; comunicación de supervisión; orden de pedido a proveedores', 'Recepción de bienes y servicios y supervisión', 'Bienes y servicios recibidos; evidencias de ejecución; informe de supervisión con soportes; acta de recibo a satisfacción', 'Todos los procesos; Gestión Financiera; Suministro de Bienes y Servicios'],
      ['V', 'Suministro de Bienes y Servicios; Proyectos de Operación Logística; todos los procesos', 'Bienes y servicios recibidos; evidencias; informes de supervisión; actas de recibo', 'Evaluación y reevaluación de proveedores', 'Registros de evaluación y reevaluación; planes de mejoramiento', 'Proveedor; Suministro de Bienes y Servicios'],
      ['V', 'Suministro de Bienes y Servicios', 'Planes del proceso; registros', 'Hacer seguimiento a planes del proceso', 'Informes de gestión', 'Orientación Estratégica; Suministro de Bienes y Servicios'],
      ['V', 'Suministro de Bienes y Servicios', 'Registros del proceso', 'Medir indicadores', 'Hojas de vida de indicadores', 'Orientación Estratégica; Suministro de Bienes y Servicios'],
      ['A', 'Suministro de Bienes y Servicios', 'Informes de gestión; hojas de vida de indicadores', 'Gestionar mejoras al proceso', 'Planes de mejoramiento', 'Suministro de Bienes y Servicios; Evaluación Independiente'],
    ],
    indicadores: ['KRI03 — Contratos interadministrativos con alertas de incumplimiento'],
    riesgos: ['C02 — Interpretación subjetiva o manipulada de normas para favorecer intereses particulares', 'C04 — Direccionamiento de cotizaciones hacia proveedores específicos', 'C05 — Omisión o manipulación de comités de contratación', 'C06 — Conflictos de interés no declarados en procesos de contratación', 'C07 — Manipulación de información presentada a clientes externos en cuadros comparativos', 'C08 — Validación insuficiente o inexistente de proveedores no marco', 'C09 — Favorecimiento indebido mediante omisión de contratos marco', 'S01 — Entrega de carpeta de supervisión incompleta o tardía', 'S02 — Supervisor designado sin perfil técnico adecuado', 'S03 — Incumplimiento de obligaciones contractuales no detectado oportunamente', 'X05 — Estudios previos elaborados después de la decisión de asignación', 'X06 — Falta de re-evaluación anual de proveedores marco'],
    recursos: { humanos: 'Director Administrativo y Financiero; Director Jurídico; Profesional de Suministro de Bienes y Servicios; Gestor de Área Administrativa; Gestor Jurídico; interesado en la compra', fisicos: 'Herramientas ofimáticas; espacios de trabajo y equipos de oficina' },
    documentos: 'Listado maestro de documentos',
    requisitosISO: '6.1 Riesgos · 7.1.6 Conocimiento · 7.5 Información documentada · 8.1 Planificación · 8.4 Control de productos y servicios externos · 8.5.3 Propiedad de proveedores · 9.1.3 Análisis y evaluación · 10 Mejora',
  },
  MA: {
    codigo: 'MA-C01', version: '01', estado: 'oficial',
    archivo: ['Gestión de Procesos y Mejoramiento Continuo', 'MA-C01 Caracterizacion Mejoramiento Activo.xlsx'],
    proceso: 'Mejoramiento Activo', lider: 'Director(a) Administrativo(a) y Financiero(a)',
    objetivo: 'Implementar acciones encaminadas a la medición, análisis, evaluación y mejora del Sistema de Gestión mediante el diseño y la aplicación de lineamientos y metodologías que permitan la conformidad de productos y servicios, optimizar la gestión institucional y la satisfacción de las partes interesadas y grupos de valor.',
    alcance: 'Inicia con la definición e implementación de lineamientos y metodologías, continúa con la formulación de acciones de mejora, y finaliza con la optimización de los procesos y la satisfacción de las partes interesadas y grupos de valor.',
    resultado: 'Ambiente de mejoramiento continuo con seguimiento a planes de mejora, cambios y evaluaciones, generando mejora a los procesos internos del SGC — Receptores: todos los procesos.',
    phva: [
      ['P', 'Todos los procesos', 'Acciones de mejora', 'Planear las actividades de mejoramiento continuo: plan de auditoría anual, informe anual de acciones de mejora e informe de evaluación de satisfacción', 'Plan anual de auditoría', 'Todos los procesos; líderes de proceso; líderes de los sistemas de gestión; Comité Institucional de Gestión y Desempeño; Oficina de Control Interno'],
      ['H', 'Líderes de los sistemas de gestión; líderes de políticas de gestión y desempeño; Comité de Coordinación de Control Interno', 'Programa de auditorías de los sistemas del SIG', 'Auditoría interna del SGC: selección de auditores, plan de auditoría, ejecución, informe, hallazgos y evaluación del auditor', 'Informe de auditoría interna de los sistemas', 'Líderes de proceso'],
      ['H', 'Todos los procesos', 'Solicitudes de gestión de cambio', 'Gestión del cambio: detección, evaluación de impacto/beneficio, acciones de conformidad, recursos, implementación y seguimiento', 'Registro y control de la gestión de cambios; informe de seguimiento al desempeño del cambio', 'Líderes de proceso; líder del SGC; Comité Institucional de Gestión y Desempeño'],
      ['H', 'Todos los procesos', 'Hallazgos de todas las fuentes contempladas', 'Gestionar las mejoras: identificación de hallazgos, acciones correctivas y oportunidades de mejora, gestión, implementación y seguimiento', 'Plan de mejoramiento por proceso; sistema de administración de acciones; informes de seguimiento', 'Líderes de proceso; líderes de los sistemas; Comité Institucional de Gestión y Desempeño; Oficina de Control Interno'],
      ['H', 'Desarrollo de la Experiencia del Cliente; Proyectos de Operación Logística', 'Evaluación de satisfacción de clientes', 'Medición de la satisfacción del cliente: elaboración de la evaluación, consolidación, informe definitivo, acciones de mejora y presentación a Gerencia', 'Informe de satisfacción del cliente', 'Orientación Estratégica; Mejoramiento Activo; Proyectos de Operación Logística; Desarrollo de la Experiencia del Cliente'],
      ['V', 'Todos los procesos', 'Planes del proceso; registros', 'Hacer seguimiento a planes del proceso', 'Informes de gestión', 'Orientación Estratégica; Mejoramiento Activo'],
      ['V', 'Todos los procesos', 'Registros del proceso', 'Medir indicadores', 'Hojas de vida de indicadores', 'Orientación Estratégica; Mejoramiento Activo'],
      ['A', 'Todos los procesos', 'Informes de gestión; hojas de vida de indicadores', 'Gestionar mejoras al proceso', 'Planes de mejoramiento', 'Mejoramiento Activo; Evaluación Independiente'],
    ],
    indicadores: ['Actualización de manuales y procedimientos de la Entidad (Plan Estratégico)'],
    riesgos: ['Sin riesgos asignados a este proceso en la matriz vigente (corte 25/02/2026)'],
    recursos: { humanos: 'Jefe de Planeación o quien haga sus veces (Director Administrativo y Financiero); apoyo del SGC', fisicos: 'Herramientas ofimáticas; espacios de trabajo y equipos de oficina' },
    documentos: 'Listado maestro de documentos y registros',
    requisitosISO: '4.1 Conocimiento de la organización y su contexto · 6.1 Acciones para abordar riesgos y oportunidades · 6.3 Planificación de los cambios · 7.1.1 Generalidades · 7.1.2 Personas',
  },
  CI: {
    codigo: 'EV-C01', version: '01', estado: 'oficial',
    archivo: ['Gestión de Control Interno', 'EV-C01 Caracterización evaluación independiente.xls'],
    proceso: 'Evaluación independiente', lider: 'Jefe de Oficina de Control Interno',
    objetivo: 'Evaluar la gestión mediante la ejecución de auditorías, seguimientos y elaboración de informes a los diferentes procesos, procedimientos, dependencias, proyectos y contratos de la entidad, según requisitos legales y reglamentarios y generando recomendaciones que contribuyan al mejoramiento continuo y cumplimiento.',
    alcance: 'Inicia con la elaboración del Plan Anual de Auditoría y culmina con su seguimiento.',
    resultado: 'Evaluación de la gestión mediante auditorías, seguimientos e informes con recomendaciones para el mejoramiento continuo — Receptores: todos los procesos.',
    phva: [
      ['P', 'Congreso de la República; organismos de control; procesos del sistema', 'Normas constitucionales y legales vigentes; necesidades de auditoría', 'Elaboración del Plan Anual de Auditorías, Seguimientos e Informes de Control Interno', 'Plan Anual de Auditorías, Seguimientos e Informes', 'Evaluación Independiente; Orientación Estratégica'],
      ['H', 'Todos los procesos', 'Información del proceso, dependencia, proyecto, contrato o tema a auditar', 'Realizar las auditorías, seguimientos e informes de Control Interno; elaborar, presentar y publicar los informes', 'Registros de auditoría (listas de chequeo, notas); informes de auditoría; informes a entes de control', 'Evaluación Independiente; entes de control; todos los procesos'],
      ['V', 'Todos los procesos', 'Riesgos y oportunidades reportados', 'Seguimiento a los controles de los riesgos', 'Reportes de materialización del riesgo; matriz y mapa de riesgos actualizado', 'Todos los procesos'],
      ['V', 'Todos los procesos', 'Planes de mejoramiento', 'Evaluación y seguimiento de los planes de mejoramiento', 'Resultados del seguimiento de los planes de mejoramiento', 'Todos los procesos'],
      ['V', 'Evaluación Independiente', 'Plan Anual de Auditorías', 'Seguimiento y evaluación de las actividades del Plan Anual de Auditorías', 'Informe de seguimiento del Plan Anual de Auditorías', 'Evaluación Independiente'],
      ['A', 'Evaluación Independiente', 'Informes de gestión; hojas de vida de indicadores', 'Gestionar mejoras al proceso', 'Planes de mejoramiento', 'Todos los procesos; partes interesadas'],
      ['A', 'Todos los procesos', 'Registros del proceso', 'Medir indicadores', 'Hojas de vida de indicadores; informes de gestión', 'Todos los procesos'],
    ],
    indicadores: ['Ejecución del Plan Anual de Auditoría (Plan Estratégico)', 'KRI10 — Obligaciones de Contraloría cumplidas a tiempo (Gestión Transparente)'],
    riesgos: ['M08 — Insuficiencia de capacidad técnica en Control Interno (perfil financiero requerido por Contraloría)'],
    recursos: { humanos: 'Jefe de Oficina de Control Interno', fisicos: 'Espacios de trabajo y equipos de oficina; herramientas ofimáticas (Microsoft Office)' },
    documentos: 'Procedimiento de control interno',
    requisitosISO: '4.1 Contexto · 4.2 Partes interesadas · 6.1 Riesgos y oportunidades · 6.3 Cambios · 7.3 Toma de conciencia · 7.5.2–7.5.3 Información documentada · 9.1 Seguimiento y medición · 9.2 Auditoría interna · 10.1–10.3 Mejora',
  },
  PD: {
    codigo: 'DISC-C01', version: '01', estado: 'oficial',
    archivo: ['Gestión Proceso Disciplinario y juzgamiento', 'DISC-C01 Caracterización Proceso Disciplinario.xlsx'],
    proceso: 'Procesos Disciplinarios en Etapa de Instrucción', lider: 'Jefe de Instrucción de Control Interno Disciplinario',
    objetivo: 'Adelantar el proceso disciplinario en etapa de instrucción, en aras de determinar la necesidad de formular cargos a algún servidor público adscrito a la Empresa de Parques y Eventos de Antioquia - ACTIVA, por la ocurrencia de una falta disciplinaria.',
    alcance: 'Inicia con la recepción de la queja, informe o de oficio para determinar la pertinencia y competencia de iniciar la acción disciplinaria y termina con un auto inhibitorio, auto de terminación de proceso disciplinario - archivo definitivo o remisión para juzgamiento, posterior a formulación de pliego de cargos.',
    resultado: 'Recepción y radicación de quejas e informes; impulso procesal de los expedientes; servidores públicos capacitados — Receptores: servidores públicos de ACTIVA y entes de control.',
    phva: [
      ['P', 'Oficina de Control Interno Disciplinario; servidores públicos; entes de control', 'Queja, solicitud o informe respecto a una conducta de un servidor público disciplinable', 'Formular el plan de acción; recibir y radicar la queja; identificar de oficio conductas relevantes; valorar y repartir según su contenido', 'Plan de acción aprobado; traslado a otra dependencia; archivo; reparto interno; auto inhibitorio', 'Todos los procesos del sistema; entidades externas; servidores públicos'],
      ['H', 'Proceso Control Disciplinario; usuarios internos y externos; ramas del poder público; entidades gubernamentales', 'Acciones del proceso; plan de capacitación en normativa disciplinaria; informes y quejas; políticas de operación y conceptos jurídicos', 'Ejecutar las acciones del proceso en sede de instrucción: evaluación de quejas, auto inhibitorio, indagación previa, investigación, pliego de cargos y remisión a juzgamiento, terminación o archivo, caducidad y prescripción; ejecutar el plan de capacitación', 'Impulsos procesales; expedientes tramitados; servidores capacitados', 'Usuario interno o externo; Proceso de Instrucción de Control Interno Disciplinario'],
      ['V', 'Proceso Control Disciplinario; Atención a la Ciudadanía; Evaluación Independiente; Procuraduría', 'Resultados de gestión; PQRSD; informes de auditoría', 'Gestionar PQRSD; seguimiento a partes interesadas; seguimiento y medición; seguimiento a riesgos; análisis de hallazgos; verificación de eficacia de acciones', 'PQRSD gestionadas; partes interesadas actualizadas; indicadores analizados; riesgos gestionados; hallazgos analizados; eficacia verificada', 'Instrucción de Control Interno Disciplinario; Mejoramiento Continuo; Atención a la Ciudadanía'],
      ['A', 'Mejoramiento Continuo', 'Indicadores y riesgos analizados; PQRSD gestionadas; hallazgos de auditoría; procedimiento de acciones correctivas', 'Tomar acciones para la mejora (con base en indicadores, riesgos y PQRSD)', 'Planes de mejoramiento implementados', 'Instrucción de Control Interno Disciplinario; Evaluación Independiente; Mejoramiento Continuo; Procuraduría'],
    ],
    indicadores: ['Sin indicadores asignados a este proceso en los documentos vigentes'],
    riesgos: ['Sin riesgos asignados a este proceso en la matriz vigente (corte 25/02/2026)'],
    recursos: { humanos: 'Jefe de Oficina de Control Interno Disciplinario', fisicos: 'Espacios de trabajo y equipos de oficina; instalaciones, hardware, software e internet' },
    documentos: 'Procedimiento de Instrucción de Control Interno Disciplinario y sus formatos; Constitución Política; Código Disciplinario Único; CPACA; códigos Penal, General del Proceso, Sustantivo y Procesal del Trabajo, y de Comercio',
    requisitosISO: '4 Contexto · 5.3 Roles y responsabilidades · 6 Planificación · 7.1.2 Personas · 7.2 Competencias · 7.3 Toma de conciencia · 7.4 Comunicación · 7.5 Información documentada · 9 Evaluación del desempeño · 10 Mejora',
  },
};

// Inventario COMPLETO del repositorio SharePoint (rastreo del 20/07/2026):
// todas las carpetas de proceso con sus subcarpetas, Documentos del SGC,
// Diagramas y Documentos Word. {carpeta: {subruta: [archivos...]}}.
const REPO_ARCHIVOS = {
  "Diagramas": {
    "": [
      "EV-P01 Procedimiento Auditorías Internas.vsdx",
      "GD-P03 Procedimiento Préstamo de expedientes.vsdx",
      "GD-P04 Procedimiento Almacenamiento de documentos.vsdx",
      "GD-P05 Procedimiento_Transferencias documental.vsdx",
      "GD-P06 Procedimiento Custodia documental.vsdx",
      "GF-P01 Procedimiento de pagos.vsdx",
      "GF-P02 Procedimiento para la elaboración del certificado de disponibilidad presupuestal.vsdx",
      "GF-P03 Procedimiento para la elaboración del registro presupuestal (RP).vsdx",
      "GF-P05 Procedimiento para liquidación y pago de impuestos, tasas y estampillas.vsdx",
      "GF-P05 Procedimiento para liquidación y pago de impuestos, tasas y estampillas2.vsdx",
      "GF-P05 Procedimiento para liquidación y pago de impuestos, tasas y estampillas3.vsdx",
      "GF-P06 Procedimiento de facturacion y cobro de cartera.vsdx",
      "GF-P07 Procedimiento de autorización pagos desde cuenta supervisada.vsdx",
      "Gestión de Inventarios - copia.vsdx",
      "Gestión de Inventarios_.vsdx",
      "RC-P01 Procedimiento comunicaciones.vsdx",
      "Registro de proveedores.vsdx",
      "SBS-P06 Procedimiento polizas clientes1.vsdx",
      "SBS-P07 Procedimiento liquidaciones de contratos 2.vsdx",
      "SBS-P07 Procedimiento liquidaciones de contratos 3.vsdx",
      "SBS-P07 Procedimiento liquidaciones de contratos.vsdx",
      "SBS-P09 Procedimiento Publicación SECOP.vsdx",
    ],
  },
  "Documentos Word": {
    "": [
      "CARACTERIZACIÓN GRUPOS DE VALOR.docx",
      "EV-P02 Procedimiento Auditorías Internas.docx",
      "EVEN-P01 Procedimiento Eventos Propios.docx",
      "GA-F019 Política tramite de datos personales.docx",
      "GA-I01 Instructivo Gestion Transparente Contratos.docx",
      "GA-I02 Instructivo para realización de reserva de vehículo activa.docx",
      "GA-P02 Procedimiento de Viáticos y Gastos de Viaje.docx",
      "GA-P03 Procedimiento gestión de inventarios.docx",
      "GD-P01 Procedimientos para la gestión de correspondencia recibida y enviada.docx",
      "GD-P02 Procedimiento para la gestión de PQRS.docx",
      "GD-P03 Procedimiento Préstamo de expedientes.docx",
      "GD-P04 Procedimiento Almacenamiento de expedientes.docx",
      "GD-P05 Procedimiento Transferencias documental.docx",
      "GD-P06 Procedimiento Custodia documental.docx",
      "GD-P07 Procedimiento Buzón de Sugerencias.docx",
      "GF-P01 Procedimiento de pagos.doc",
      "GF-P02 Procedimiento para la elaboración del certificado de disponibilidad presupuestal.docx",
      "GF-P03 Procedimiento para la elaboración del registro presupuestal (RP).docx",
      "GF-P04 Procedimiento conciliaciones bancarias.docx",
      "GF-P05 Procedimiento para liquidación y pago de impuestos, tasas y estampillas.docx",
      "GF-P06 Procedimiento de facturacion y cobro de cartera.docx",
      "GF-P07 Procedimiento de autorización pagos desde cuenta supervisada.docx",
      "GOP-I05 Instructivo Comité de Asignación.docx",
      "GOP-P10 Soluciones para el suministro de bienes y servicios.docx",
      "GT-I01 Instructivo para la concertación de acuerdos de gestión.docx",
      "GT-P01 Procedimiento de reclutamiento, selección y vinculación de personal.docx",
      "GT-P02 Procedimiento de inducción, entrenamiento y reinducción.docx",
      "GT-P03 Procedimiento de desvinculación de personal.docx",
      "GT-P04 Procedimiento de liquidación y pago de nomina.docx",
      "GTI-I02 Instructivo Mesa de Ayuda.docx",
      "GTI-I3 Instructivo creación de contratos en SAFIX.docx",
      "GTI-I4 Aprobación informe de Evidencias CLIENTES.docx",
      "GTI-I5 TUTORIAL INSTALACIÓN APLICATIVO CHIP LOCAL.docx",
      "GTI-I6 INSTRUCTIVO PARA AGREGAR UN CALENDARIO DESDE EL DIRECTORIO.docx",
      "GTI-I7 INSTRUCTIVO RESERVA DE CARRO.docx",
      "GTI-I8 INSTRUCTIVO RESERVA SALA DE REUNIONES.docx",
      "GTI-I9 INGRESO AL MICROSOFT 365.docx",
      "GTI-P01 Proceso de Gestión Tecnológica.docx",
      "Informe sobre el estado de formulación, validación y supervisión del PTEP.docx",
      "MA-P01 procedimiento de Gestión de mejoras.doc",
      "Manual versión Gerente 12072023.docx",
      "OE-F08 Manual de calidad.docx",
      "OE-M02 Politica de administracion gestión del riesgo.docx",
      "OE-M03 Manual de Uso Vehicular Activa.docx",
      "OE-M04 Plan de Manejo Integral de Residuos Sólidos (PMIRS).docx",
      "OE-M05 Manual Manejo de la Crisis.docx",
      "OE-M06 Protocolo vocería institucional.docx",
      "OE-M08 Procedimiento para publicacón en página web.docx",
      "OE-M09 Política de Comunicaciones.docx",
      "OE-P01 Procedimiento atención  vía telefónica.docx",
      "OE-P02 Procedimiento atención  personas con discapacidad.docx",
      "PD-P01 Procedimiento Instrucción Control Interno Disciplinario.docx",
      "POL-I01 Comité de Asignaciones.docx",
      "POL-I02 Políticas comerciales.docx",
      "POL-P01 Procedimiento ejecucion de contratos clientes.docx",
      "Plan Institucional de Archivos.docx",
      "Plan Institucional de Archivos.pdf",
      "Procedimiento para gestión comercial.docx",
      "RC-P01 Procedimiento comunicaciones.docx",
      "SBS-P05 Procedimiento de contratación.docx",
      "SBS-P06 Procedimiento pólizas clientes.docx",
      "SBS-P07 Procedimiento liquidaciones de contratos.docx",
      "SBS-P08 Procedimiento Registro de Proveedores.docx",
      "SBS-P09 Procedimiento Publicación SECOP.docx",
      "Términos y condiciones página web ACTIVA.docx",
    ],
  },
  "Documentos del SGC": {
    "": [
      "ACUERDO Nro. 002 - POR EL CUAL SE ADOPTAN LOS ESTATUTOS DE LA EMPRESA DE PARQUES Y EVENTOS DE ANTIOQUIA – ACTIVA-F.pdf",
      "CARACTERIZACIÓN GRUPOS DE VALOR.pdf",
      "Circular 014 de 2023. Rendición Secop.pdf",
      "Decreto Ordenanzal Activa.pdf",
      "Formato Políticas.docx",
      "Informe sobre el estado de formulación, validación y supervisión del PTEP.pdf",
      "MAPA DE SITIO WEB ACTIVA.pdf",
      "Matriz de Riesgos Activa.xlsx",
      "Normograma.xlsx",
      "OE-F08 Manual de calidad.pdf",
      "OE-F09 Objetivos_de_Calidad.xlsx",
      "OE-M02 Politica de administracion gestión del riesgo.pdf",
      "Objetivos de los Procesos.xlsx",
      "Organigrama Interactivo - ACTIVA.pdf",
      "Organigrama.png",
      "Plantilla Presentación ACTIVA.pptx",
      "Plantilla Procedimiento.doc",
      "RESOLUCION-032-DE-2024.pdf",
      "RESOLUCIÓN 006 DE 2025 - 202503000011 MANUAL DE FUNCIONES.pdf",
      "RESOLUCIÓN 012 DE 2023.pdf",
      "Reglamentos, manuales y Políticas.docx",
      "Seguimiento a Mapa de Riesgos.pdf",
      "Términos y condiciones página web ACTIVA.pdf",
    ],
  },
  "Gestión Administrativa": {
    "": [
      "GA-C01 Caracterización Gestión Administrativa.xls",
      "GA-F019 Política tramite de datos personales.pdf",
      "Plan de Mantenimiento 2022.xlsx",
    ],
    "Formatos": [
      "GA-F01 Acta de entrega muebles y equipos Funcionarios.docx",
      "GA-F012 Formato Registro de Inventario Activa.xlsx",
      "GA-F013 Formato Bienes de Consumo Activa.xlsx",
      "GA-F014 Aval de pago.docx",
      "GA-F015 Certificado de Cumplimiento.docx",
      "GA-F026 Formato Solicitud Baja de Bienes.xlsx",
      "GA-F08 Autorización comisión de servicios, otorgamiento de viáticos, gastos de transporte y desplazamiento.xlsx",
      "GA-F09 Liquidación del reconociendo de viáticos.xlsx",
      "GA-F16 Liquidación del reconociendo de gastos de viaje.xlsx",
      "GA-F17 Solicitud comisión de servicios.xlsx",
      "GA-F20 Aviso de Pricacidad.docx",
      "GA-F21 Tratamiento de datos en el proceso de selección.docx",
      "GA-F22 Registro de incidentes.docx",
      "GA-F23 Formato Hoja de Control.xlsx",
      "GA-F24 Acta de entrega muebles y equipos Contratistas.docx",
      "GA-F25 Formato Plan de Acción.xlsx",
      "GA-F27 Lista de chequeo Viaticos.xlsx",
    ],
    "Instructivos": [
      "GA-I01 Instructivo Gestion Transparente Contratos.pdf",
      "GA-I02 Instructivo para realización de reserva de vehículo activa.pdf",
    ],
    "Procedimientos": [
      "GA-P02 Procedimiento de viáticos y gastos de viaje.pdf",
      "GA-P03 Procedimiento gestión de inventarios.pdf",
    ],
  },
  "Gestión Comercial": {
    "": [
      "POL-C01 Caracterización Gestión Comercial.xls",
    ],
    "Formatos": [
      "DEC-F01 Formato cotización OP operador.xlsx",
      "DEC-F02 Formato cotización OP a cliente.xlsx",
      "DEC-F03 Modelo propuesta.docx",
      "DEC-F04 Lista de chequeo Clientes.xlsx",
      "POL-F01 Acta comité asignación.docx",
      "POL-F03 Formato seguimiento a requerimientos.xlsx",
      "POL-F04 Seguimiento a contratos marco.xlsx",
      "POL-F05 Informe de evidencias.pptx",
      "POL-F06 Informe técnico.docx",
      "POL-F08 Formato remision.xlsx",
      "POL-F09 Solicitud de pedido.docx",
      "POL-F11 Acta de Entrega Servicio de Alimentación.docx",
      "POL-F12 Planilla Paz y Salvo Transporte Implementación Deportiva.docx",
      "POL-F13 Formato Liquidación de Eventos ACTIVA.xlsx",
      "POL-F14 Planilla Paz y Salvo Transporte Personal .docx",
      "POL-F15 Remisión Servicios Adicionales.xlsx",
      "POL-F16 Formato seguimiento a requerimientos comunicaciones.xlsx",
      "POL-F17 FORMATO PRODUCTO, DESCRIPCIÓN, CANTIDAD PARA OP.xlsx",
      "POL-F18 FORMATO REMISIÓN IMPREVISTOS.xlsx",
      "POL-F19 Planilla de Alojamiento.xlsx",
      "POL-F20 Formato Registro Hotelero.docx",
      "POL-F21 Planilla de Auxilio Económico.xlsx",
      "POL-F22 Planilla Paz y Salvo Transporte.docx",
      "POL-F23 Acta de entrega material educativo.docx",
    ],
    "Instructivos": [
      "POL-I01 Comité de Asignaciones.pdf",
      "POL-I02 Políticas comerciales.pdf",
    ],
    "Procedimientos": [
      "DEC-P01 procedimiento Gestión comercial.pdf",
      "POL-P01 Procedimiento ejecucion de contratos clientes.pdf",
    ],
  },
  "Gestión Documental": {
    "": [
      "GD-C01 Caracterización Programa de Gestión Documental.xlsx",
      "PROGRAMA DE GESTIÓN DOCUMENTAL 2022-2024.docx",
      "Plan Institucional de Archivos 2022-2024.docx",
    ],
    "Formatos": [
      "GD-F04 Formato Planilla de correspondencia para mensajería.xlsx",
      "GD-F06 Formato informe mensual PQRSFD.docx",
      "GD-F07 Formato único inventario documental.xlsx",
      "GD-F08 Formato inventario de existencia expedientes contractuales.xlsx",
      "GD-F09 Formato control préstamo documental.xlsx",
      "GD-F10 Formato rotúlo carpetas de archivo.xlsx",
      "GD-F11 Formato rótulo identificación cajas de archivo.xlsx",
      "GD-F12 Formato inventario topográfico documental.xlsx",
      "GD-F13 Formato hoja de control.xlsx",
      "GD-F14 Formato tablas de retención documental.xlsx",
      "GD-F15 Formato cuadros de clasificación documental.xlsx",
      "GD-F16 Formato inventario de eliminación documental.xlsx",
      "GD-F17 Formato inventario general de liquidación anual.xlsx",
      "GD-F18 Formato cronograma transferencias documentales.xlsx",
      "GD-F19 Formato acta autorización salida de expedientes.docx",
      "GD-F20 Formato acta ingreso de expedientes.docx",
      "GD-F21 Planilla transferencias documentales.xlsx",
      "GD-F22 Comunicación Interna Activa.docx",
      "GD-F23 Comunicación Externa Activa.docx",
      "GD-F24 Acta de instalación y formalización.docx",
      "GD-F25 Acta Parcial Sesiones de Empalme.docx",
      "GD-F26 Registro de Entrega de Insumos de Papelería.xlsx",
      "GD-F27 Acta de Apertura.docx",
      "GD-F29 Formato Registro de PQRSFD.xlsx",
      "GD-F30 Formato Referencia Cruzada.xlsx",
      "GD-F31 Formato Rótulo de Identificación de Estantería de Archivos.xlsx",
    ],
    "Formatos/Obsoletos": [
      "GD-F01 Formato Registro de correspondencia recibida.xlsx",
      "GD-F02  Formato Registro de correspondencia enviada.xlsx",
      "GD-F03 Formato Planilla de registro de facturas.xlsx",
      "GD-F05 Formato Distribución de correspondencia recibida.xlsx",
    ],
    "Procedimientos": [
      "GD-P01 Procedimiento para la gestión de correspondencia recibida y enviada.pdf",
      "GD-P02 Procedimiento para la gestión de PQRS.pdf",
      "GD-P03 Procedimiento Préstamo de expedientes.pdf",
      "GD-P04 Procedimiento Almacenamiento de documentos.pdf",
      "GD-P05 Procedimiento_Transferencias documental.pdf",
      "GD-P06 Procedimiento Custodia documental.pdf",
      "GD-P07 Procedimiento Buzón de Sugerencias.pdf",
    ],
  },
  "Gestión Financiera": {
    "": [
      "GF-C01 Caracterización Gestión Financiera.xls",
    ],
    "Formatos": [
      "GF-F01 Certificado de disponibilidad presupuestal.xlsx",
      "GF-F02 Certificado de registro presupuestal.xlsx",
      "GF-F04 Solicitud de CDP.docx",
      "GF-F05 Solicitud de RP.docx",
      "GF-F06 Reintegro de recursos contrato marco.docx",
      "GF-F07 Solicitud de liberación de CDP y RP.docx",
      "GF-F08 Boletín diario de bancos.xlsx",
      "GF-F09 Recibo pago anticipado.docx",
      "GF-F10 Viabilidad presupuestal.docx",
      "GF-F11 Planilla de pago contratistas.xlsx",
      "GF-F12 Formato solicitud caja menor.xlsx",
      "GF-F13 Formato Reembolso de caja menor.xlsx",
      "GF-F14 Arqueo de caja menor.xlsx",
      "GF-F15 Documento equivalente no responsables de IVA.xlsx",
      "GF-F16 Formato para Anticipo o Desembolso.docx",
      "GF-F17 Informe de ejecución recursos administración delegada o mandato.xlsx",
      "GF-F18  Lista de Verfiicación para Autorización de desembolso en cuenta supervisada.xlsx",
      "GF-F19 Formato para solicitar autorización desembolso desde cuenta supervisada.docx",
      "GF-F20 Formato para autorizar desembolso desde cuenta supervisada.docx",
      "GF-F21 Formato consolidado de ejecución final.xlsx",
      "GF-F22 Informe de mandato.xlsx",
      "GF-F23 Certificación pago a terceros.docx",
      "GF-F24 Formato requisitos de pago a terceros.docx",
      "GF-F25 Lista de Chequeo Certificación Facturas.docx",
      "GF-F26 Certificación Pago Directo.docx",
      "GF-F27 Estimación de Ingresos en Servicios.xlsx",
      "GF-F28 Estimación de Costo en Servicios.xlsx",
      "GF-F29 Certificado de liberación presupuestal.xlsx",
      "GF-F30 SARLAFT Declaracion PERSONA NATURAL.docx",
      "GF-F31 SARLAFT Declaracion PERSONA JURIDICA.docx",
    ],
    "Plantillas": [
      "Palntilla Resolución xxx - Traslado Presupuestal.docx",
      "Plantilla Certificado Juramentado de Dependientes 2021.docx",
    ],
    "Procedimientos": [
      "GF-P01 Procedimiento de pagos.pdf",
      "GF-P02 Procedimiento para la elaboración del certificado de disponibilidad presupuestal.pdf",
      "GF-P03 Procedimiento para la elaboración del registro presupuestal (RP).pdf",
      "GF-P04 Procedimiento conciliaciones bancarias.pdf",
      "GF-P05 Procedimiento para liquidación y pago de impuestos, tasas y estampillas.pdf",
      "GF-P06 Procedimiento de facturacion y cobro de cartera.pdf",
      "GF-P07 Procedimiento de autorización pagos desde cuenta supervisada.pdf",
    ],
  },
  "Gestión Proceso Disciplinario y juzgamiento": {
    "": [
      "DISC-C01 Caracterización Proceso Disciplinario.xlsx",
    ],
    "Formatos": [
      "PD-F01 Acta de Confesión y Aceptación de Responsabilidad.docx",
      "PD-F02 Ampliación y Ratificación de Queja.docx",
      "PD-F03 Auto Accede a Solicitud de Copias.docx",
      "PD-F04 Auto Apertura de Apertura de Investigación Disciplinaria.docx",
      "PD-F05 Auto de Cierre de Investigación Disciplinaria y Traslado para Alegatos Precalificatorios.docx",
      "PD-F06 Auto de Remisión.docx",
      "PD-F07 Auto de Vinculación de Sujeto a una Investigación Disciplinaria.docx",
      "PD-F08 Auto Declara Impedimento.docx",
      "PD-F09 Auto Despacho Comisorio.docx",
      "PD-F10 Auto Declara la Prescripción.docx",
      "PD-F11 Auto Declara una Nulidad.docx",
      "PD-F12 Auto Decreto de Pruebas.docx",
      "PD-F13 Auto Formulación Pliego de Cargos.docx",
      "PD-F14 Auto General.docx",
      "PD-F15 Auto Indagación Previa.docx",
      "PD-F16 Auto Inhibitorio.docx",
      "PD-F17 Auto Posesión Defensor Público o Estudiante de Consultorio Jurídico.docx",
      "PD-F18 Auto que Ordena Acumulación Procesal.docx",
      "PD-F19 Auto Reconoce Personería Jurídica.docx",
      "PD-F20 Auto Resuelve Procedencia de Recurso de Apelación.docx",
      "PD-F21 Auto Resuelve Recurso de Queja.docx",
      "PD-F22 Auto Resuelve Recurso de Reposición.docx",
      "PD-F23 Auto Terminación de Proceso – Archivo Definitivo.docx",
      "PD-F24 Constancia de Archivo Físico.docx",
      "PD-F25 Constancia de Ejecutoria.docx",
      "PD-F26 Constancia Secretarial.docx",
      "PD-F27 Declaración Bajo la Gravedad de Juramento.docx",
      "PD-F28 Formato de Registro para Control de Procesos Disciplinarios.xlsx",
      "PD-F29 Notificación Personal.docx",
      "PD-F30 Notificación por Edicto.docx",
      "PD-F31Notificación por Estados.docx",
      "PD-F32 Queja Disciplinaria.docx",
      "PD-F33 Versión Libre y Espontánea.docx",
    ],
    "Procedimiento": [
      "PD-P01 Procedimiento Instrucción Control Interno Disciplinario.pdf",
    ],
  },
  "Gestión de Control Interno": {
    "": [
      "Codigo de etica del auditor.pdf",
      "EV-C01 Caracterización evaluación independiente.xls",
      "Estatuto de Control Interno_Activa _ Res 079 de 16_12_2022.pdf",
    ],
    "Formatos": [
      "EV-F01 Formato gestión mejoras.xlsx",
      "EV-F02 Plan Anual de Auditoria CI.xlsx",
      "EV-F03 Plan de Mejoramiento.xlsx",
    ],
    "Procedimientos": [
      "EV-P02 Procedimiento Auditorías Internas.pdf",
    ],
  },
  "Gestión de Operaciones": {
    "Formatos": [
      "GOP-F01 Solicitud de Bienes y Servicios.xlsx",
      "GOP-F02 Adición Solicitud de Bienes y Servicios.xlsx.docx",
      "GOP-F03 Formato de Ampliación Plazo SBS.docx",
    ],
    "Instructivos": [
      "GOP-I05 Instructivo Comité de Asignación.pdf",
    ],
    "Procedimientos": [
      "GOP-P10 Soluciones para el suministro de bienes y servicios.pdf",
    ],
  },
  "Gestión de Procesos y Mejoramiento Continuo": {
    "": [
      "MA-C01 Caracterizacion Mejoramiento Activo.xlsx",
    ],
    "Procedimientos": [
      "MA-P01 procedimiento de Gestión de mejoras.pdf",
    ],
  },
  "Gestión de Tecnologías de Información": {
    "": [
      "GTI-P01 Proceso de Gestión Tecnológica.pdf",
      "Política de Uso Aceptable de Recursos Tecnológicos ACTIVA.pdf",
    ],
    "Instructivos": [
      "GA-I01 Instructivo Gestion Transparente Contratos.pdf",
      "GTI-I02 Instructivo Mesa de Ayuda.pdf",
      "GTI-I3 Instructivo creación de contratos en SAFIX.pdf",
      "GTI-I4 Aprobación informe de Evidencias CLIENTES.pdf",
      "GTI-I5 TUTORIAL INSTALACIÓN APLICATIVO CHIP LOCAL.pdf",
      "GTI-I6 INSTRUCTIVO PARA AGREGAR UN CALENDARIO DESDE EL DIRECTORIO.pdf",
      "GTI-I7 INSTRUCTIVO RESERVA DE CARRO.pdf",
      "GTI-I8 INSTRUCTIVO RESERVA SALA DE REUNIONES.pdf",
      "GTI-I9 INGRESO AL MICROSOFT 365.pdf",
      "Instructivo_Pedir_RP_OP_Internas.pdf",
    ],
  },
  "Gestión del Talento Humano": {
    "": [
      "CODIGO DE INTEGRIDAD ACTIVA.pdf",
      "GT-C01 Caracterización Gestión del Talento Humano.xlsx",
      "Presentación Codigo de Integridad.pdf",
    ],
    "Formatos": [
      "GT-F01 Lista de chequeo vinculación de personal.docx",
      "GT-F02 Solicitud disminución base gravable.docx",
      "GT-F026 Paz y Salvo para Retiro de Funcionarios.xlsx",
      "GT-F03 Programa de inducción.xlsx",
      "GT-F05 Lista de asistencia.xlsx",
      "GT-F06 Lista de chequeo retiro de personal_.docx",
      "GT-F07 Encuesta de retiro de servidores públicos y trabajadores oficiales_.xlsx",
      "GT-F08 Control de Permisos.xlsx",
      "GT-F09 Acuerdo de confidencialidad.docx",
      "GT-F10 Contrato de trabajo término indefinido trabajadores oficiales.docx",
      "GT-F11 Autorización descuento por nómina.docx",
      "GT-F12 Acta entrega del cargo.docx",
      "GT-F13 Constancia de horario flexible.docx",
      "GT-F14 Informe de Entrega del Cargo.docx",
      "GT-F15 Acta informe de Gestión.docx",
      "GT-F16 Evaluacion_de_la_Induccion.docx",
      "GT-F17 Evaluación de la entrevista .xlsx",
      "GT-F18 Formato retiro de Cesantias.xlsx",
      "GT-F19 Solicitud de vacaciones o licencias no remuneradas.xlsx",
      "GT-F20 Evaluación de desempeño.xlsx",
      "GT-F23 Autorización de Tratamiento de Datos Personales.docx",
      "GT-F24 Autorización para la publicacion de imagenes.docx",
      "GT-F25 Formato comision cumplido.docx",
      "GT-F27 Formato Horas Extras y Recargos.docx",
      "GT-F28 FORMATO -SST-Inspeccion locativa.xlsx",
      "GT-F29 FORMATO-SST  Inspección de botiquín para primeros auxilios vehículo.xlsx",
      "GT-F30 FORMATO-SST  Inspección de botiquin para primeros auxilios.xlsx",
      "GT-F31 FORMATO-SST Inspección de extintores.xlsx",
      "GT-F32 Matriz Identificación de Peligros.xlsx",
      "GT-F33 FORMATO-SST Inspección de extintores.xlsx",
    ],
    "Instructivos": [
      "GT-I01 Instructivo para la concertación de acuerdos de gestión.pdf",
    ],
    "Procedimientos": [
      "GT-P01 Procedimiento de reclutamiento, selección y vinculación de personal.pdf",
      "GT-P02 Procedimiento de inducción, entrenamiento y reinducción.pdf",
      "GT-P03 Procedimiento de desvinculación de personal.pdf",
      "GT-P04 Procedimiento de liquidación y pago de nomina.pdf",
    ],
  },
  "Modelo Integrado de Planeación y Gestión": {
    "": [
      "DIMENSIONES.png",
      "MIPG.png",
      "Plan de Acción Programa de Transparencia y Ética Pública.xlsx",
      "Plantilla Presentación ACTIVA.pptx",
      "Politica de Fortalecimiento Institucional.pdf",
      "Política Cero Papel Empresa de Parques y Eventos de Antioquia ACTIVA.pdf",
      "Política Gestión del Talento Humano.pdf",
      "Política de Archivos y Gestión Documental.pdf",
      "Política de Compras y Contratación.pdf",
      "Política de Control Interno.pdf",
      "Política de Gestión Estadística.pdf",
      "Política de Gestión Presupuestal.pdf",
      "Política de Gestión del Conocimiento.pdf",
      "Política de Gobierno Digital.pdf",
      "Política de Integridad.pdf",
      "Política de Mejora Normativa.pdf",
      "Política de Participación Ciudadana.pdf",
      "Política de Planeación Institucional.pdf",
      "Política de Prevención del Daño Antijurídico.pdf",
      "Política de Racionalización de Trámites.pdf",
      "Política de Seguimiento y Evaluación.pdf",
      "Política de Seguridad Digital.pdf",
      "Política de Servicio al Ciudadano.pdf",
      "Política de Transparencia.pdf",
      "Políticas MIPG.jpg",
      "RESOLUCIÓN Nro. 015 de 2022 Comité Institucional de Gestión y Desempeño.pdf",
      "Seguimiento al Cumplimiento de la Política de Servicio al Ciudadano.docx",
      "Seguimiento al Cumplimiento de la Política de Servicio al Ciudadano.pdf",
    ],
    "ACTAS DE COMITÉ": [
      "OE-F01_Acta de Comité No. 2.pdf",
      "OE-F01_Acta de Comité No. 3.pdf",
      "OE-F01_Formato_de_acta[1].pdf",
      "Presentación Para comité 2.pptx",
      "Presentación Programa de Transparencia y Ética Pública.pptx",
      "Presentación para comité 1.pptx",
    ],
    "AUTODIAGNOSTICOS": [
      "01 Autodiagnóstico de Gestión del talento humano.xlsx",
      "02 Autodiagnóstico de Integridad.xlsx",
      "03 Autodiagnóstico para la Gestión de Conflictos de Intereses.xlsx",
      "04 Autodiagnóstico de Transparencia y Acceso a la Información.xlsx",
      "05 Autodiagnóstico de Plan Anticorrupción.xlsx",
      "06 Autodiagnóstico de Gobierno Digital.xlsx",
      "07 Autodiagnóstico de Defensa Jurídica (Territorio).xlsm",
      "08 Autodiagnóstico de Servicio al Ciudadano.xlsx",
      "09 Autodiagnóstico de Trámites.xlsx",
      "10 Autodiagnóstico de Participación Ciudadana.xlsx",
      "11 Autodiagnóstico de Rendición de Cuentas.xlsx",
      "14 Autodiagnóstico Gestión Documental.xlsx",
      "15 Autodiagnóstico de Gestión del Conocimiento y la Innovación.xlsx",
      "16 Autodiagnóstico de Control Interno - Entidades Pequeñas.xlsx",
      "17 Autodiagnóstico  de Gestión de la Información Estadística (Territorio) - Nueva.xlsx",
      "Autodiagnostico MIPG V21_02_2023 .pdf",
      "RESUMEN.xlsx",
    ],
    "Certificados Curso Integridad, transparencia y lucha contra la corrupción": [
      "CERTIFICADO - CATALINA VALENCIA VANEGAS.pdf",
      "CERTIFICADO FUNCION PUBLICA 20 HORAS, INTEGRIDAD, TRANSPARENCIA JHONATAN VALENCIA GOMEZ.pdf",
      "CERTIFICADO MARIA ISABEL CASTAÑO G.pdf",
      "CERTIFICADO MARIA ISABEL SANCHEZ.pdf",
      "CURSO INTEGRIDAD, TRANSPARENCIA Y LUCHA DARLYN CHAVES CARO.pdf",
      "Certificado 20 hrs William Alejandro Riaño Florez.pdf",
      "Certificado Integridad Caterine Zapata Estrada.pdf",
      "Certificado Maria Patricia Muñoz Cardona.pdf",
      "Certificado _ Vanessa Urrea Gallo.pdf",
      "Certificado curso EVA_DRIANA.pdf",
      "Certificado de integridad y transparencia Daniela Jiménez Bedoya.pdf",
      "Certificado integridad, transparencia y lucha contra la Corrupción Karen Valencia.pdf",
      "Certificado_ Cristian Murillo.pdf",
      "Certificado_ Esteban García.pdf",
      "Certificado_ Stefany Pajón Ortiz.pdf",
      "Certificado_Víctor Daniel Lezcano Quintero.pdf",
      "Certificado_Yuri Maria Gallego Alzate.pdf",
      "HERNAN CASTAÑO_Curso Integridad, Transparencia y Lucha contra la Corrupción.pdf",
      "JUAN JOSÉ BARRIENTOS RESTREPO - CERTIFICADO FUNCIÓN PÚBLICA - INTEGRIDAD, TRANSPARENCIA Y LUCHA CONTRA LA CORRUPCIÓN.pdf",
      "ceretificado lucha contra la corrupción.pdf",
      "integridad,transparencia y lucha contra la corrupcion.pdf",
    ],
    "Políticas Word MIPG": [
      "Análisis Resultados FURAG.docx",
      "Análisis Resultados FURAG.pdf",
      "DISTRIBUCIÓN.xlsx",
      "Politica de Fortalecimiento Institucional.docx",
      "Política Gestión del Talento Humano.docx",
      "Política de Archivos y Gestión Documental.docx",
      "Política de Compras y Contratación.docx",
      "Política de Control Interno.docx",
      "Política de Gestión Estadística.docx",
      "Política de Gestión Presupuestal.docx",
      "Política de Gestión del Conocimiento.docx",
      "Política de Gobierno Digital.docx",
      "Política de Integridad.docx",
      "Política de Mejora Normativa.docx",
      "Política de Participación Ciudadana.docx",
      "Política de Planeación Institucional.docx",
      "Política de Prevención del Daño Antijurídico.docx",
      "Política de Racionalización de Trámites.docx",
      "Política de Seguimiento y Evaluación.docx",
      "Política de Seguridad Digital.docx",
      "Política de Servicio al Ciudadano.docx",
      "Política de Transparencia.docx",
    ],
  },
  "Orientación Estratégica": {
    "": [
      "OE-C01 Caracterización Orientación Estratégica.xlsx",
    ],
    "Formatos": [
      "OE-F01 Formato de acta.docx",
      "OE-F02 Formato de acta comite de gerencia-contratacion.docx",
      "OE-F03 Formato plan de acción por dirección.xlsx",
      "OE-F04 Formato Mapa de riesgos entidad.xlsx",
      "OE-F05 Listado maestro de documentos y registros.xlsx",
      "OE-F06 Formato plan estrategico.xlsx",
      "OE-F07 PMO- Project Management Office.xlsx",
      "OE-F09 Objetivos_de_Calidad.xlsx",
      "OE-F10 Formulario llamada telefónica.docx",
      "OE-F11 Formato seguimiento via telefónica.docx",
      "OE-F12 Formato Atención Discapacidad.docx",
    ],
    "Manuales": [
      "OE-M01 Manual para la información documentada.pdf",
      "OE-M03 Manual de Uso Vehicular Activa.pdf",
      "OE-M04 Plan de Manejo Integral de Residuos Sólidos (PMIRS).pdf",
      "OE-M05 Manual Manejo de la Crisis.pdf",
      "OE-M06 Protocolo vocería institucional.pdf",
      "OE-M08 Procedimiento para publicacón en página web.pdf",
      "OE-M09 Política de Comunicaciones.pdf",
    ],
    "Plantillas": [
      "Plantilla Manuales.doc",
      "Plantilla Procedimiento.doc",
      "Plantilla caracterización.xlsx",
      "Plantilla instructivo.docx",
    ],
    "Procedimientos": [
      "OE-P01 Procedimiento atención  vía telefónica.pdf",
    ],
  },
  "Planeación Estratégica": {
    "": [
      "CÓDIGO DE INTEGRIDAD ACTUALIZADO.docx",
    ],
  },
  "Proceso de Contratación": {
    "": [
      "SBS- C01 Caracterización Proceso de Contratación.xls",
    ],
    "Formatos": [
      "POL-F10 Formato Orden de Pedido OP Interna.xlsx",
      "SBS-F01 Formato solicitud de cotización de bienes y servicios.docx",
      "SBS-F02 Orden de compra de bienes y servicios.docx",
      "SBS-F03 Estudios previos.docx",
      "SBS-F04 Invitación a contratar.docx",
      "SBS-F05 Carta de presentación de la oferta.docx",
      "SBS-F06 Minuta del contrato.docx",
      "SBS-F07 Designación de supervisión.docx",
      "SBS-F08 Acta de inicio.docx",
      "SBS-F09 Aprobación de póliza.docx",
      "SBS-F10 Certificado Inhabilidades Incompatibles Contratos.docx",
      "SBS-F11 Oferta económica.xlsx",
      "SBS-F12 Formato Otrosí.docx",
      "SBS-F13 Acuerdo de confidencialidad y no divulgación.docx",
      "SBS-F14 Respuesta a observaciones.docx",
      "SBS-F15 Solicitud de modificación de contrato.docx",
      "SBS-F16 Adenda.docx",
      "SBS-F17 Conformación y designación del comité asesor y evaluador.docx",
      "SBS-F18 Informe preliminar de evaluación de proceso de selección.docx",
      "SBS-F19 Informe final de evaluación de proceso de selección.docx",
      "SBS-F20 Informe de ejecución mensual de contrato.docx",
      "SBS-F21 Informe supervisión.docx",
      "SBS-F22 Acta de recibo y terminación.docx",
      "SBS-F23 Acta de liquidacion.docx",
      "SBS-F24 Acta recibo a satisfacción.docx",
      "SBS-F25 Contrato prestación de servicios.docx",
      "SBS-F26 Certificado experiencia.docx",
      "SBS-F27 Acta comité ejecución de contratos marco.docx",
      "SBS-F28 Estudio previo prestación de servicios.docx",
      "SBS-F29 Lista de chequeo orden de compra.xlsx",
      "SBS-F30 Lista de chequeo contratación directa.xlsx",
      "SBS-F31 Lista de chequeo lista corta.xlsx",
      "SBS-F32 Lista de chequeo prestación de servicios profesionales.xlsx",
      "SBS-F33 Lista de chequeo invitación abierta.xlsx",
      "SBS-F34 Minuta del contrato prestación de servicios.docx",
      "SBS-F35 Informe de evaluación y análisis de la oferta económica y técnica.docx",
      "SBS-F36 Matriz de riesgos contratacion - Contratación Directa.xlsx",
      "SBS-F36 Matriz de riesgos contratacion - Lista corta - Invitacion abierta.xlsx",
      "SBS-F36 Matriz de riesgos contratacion - Orden de compra.xlsx",
      "SBS-F36 Matriz de riesgos contratacion.xlsx",
      "SBS-F37 Certificado comité de contratación.docx",
      "SBS-F37 Lista de chequeo Alianza Comercial.xlsx",
      "SBS-F38 Formato acta recepción de ofertas.docx",
      "SBS-F39 Lista chequeo Contratos Marco IA 001.xlsx",
      "SBS-F40 Formato de necesidad, conveniencia y oportunidad de la evaluación de oferta.docx",
      "SBS-F41 Constancia Cierre Proceso Contractual.docx",
      "SBS-F42 Acta de suspension y reinicio.docx",
      "SBS-F43 Orden de compra de bienes y servicios 3 SMMLV.docx",
      "SBS-F44 Formato Designación Comité Asesor y Evaluador.docx",
      "SBS-F45 ANEXO 2- carta de presentacion de la propuesta.docx",
      "SBS-F46 ANEXO 3 - Pago seguridad social y parafiscales.docx",
      "SBS-F47 ANEXO 4-  Compromiso anticorrupción.docx",
      "SBS-F48 ANEXO 5 - Certificado de inhabilidades e incompatibilidades.docx",
      "SBS-F49 ANEXO 6 - Acuerdo de niveles de servicio - ANS.docx",
      "SBS-F50 ANEXO 7 - Acuerdo de confidencialidad.docx",
      "SBS-F51 ANEXO 8 -  Aceptación Gestión Comercial.docx",
      "SBS-F52 ANEXO 9 - Cumplimiento y aceptación tarifario.docx",
      "SBS-F53 ANEXO 10 - Modelo acuerdo de consorcio.docx",
      "SBS-F54 ANEXO 11 - Modelo acuerdo union temporal.docx",
      "SBS-F55 ANEXO 12-  Capacidad financiera.docx",
      "SBS-F56 ANEXO 13 -  Experiencia del proponente .docx",
      "SBS-F57 ANEXO 14. Carta de compromiso del personal minimo requerido.docx",
      "SBS-F58 ANEXO 24- Oferta de porcentaje por pago a terceros_.docx",
      "SBS-F59 ANEXO 25 Experiencia adicional del proponente.docx",
      "SBS-F60 ANEXO 26. Carta de compromiso personal adicional.docx",
      "SBS-F61 Formato Orden de Servicio.docx",
      "SBS-F61 Formato Orden de Servicio.xlsx",
      "SBS-F62 Contrato Vinculación.docx",
      "SBS-F63 Autorización para la celebración de contratos y órdenes de servicio que superen los 500 SMLMV.docx",
      "SBS-F64 Acta de Recibo a Satisfaccion_Indeportes.xlsx",
      "SBS-F65 Nota Aclaratoria.docx",
      "SBS-F66 Acta Terminación ODS para liberación.docx",
      "SBS-F67 Formato Seguimiento a Contrato.xlsx",
      "SBS-F68 Lista de Chequeo Modificación Contractual.xlsx",
    ],
    "Manuales": [
      "Acuerdo No. 001 del 13 de julio de 2023 Manual de contratación.pdf",
      "SBS-M01 Manual de Supervisión.pdf",
    ],
    "Procedimientos": [
      "SBS-P01 Procedimiento de compra por órdenes de compra.pdf",
      "SBS-P02 Procedimiento de compra por contratación directa.pdf",
      "SBS-P03 Procedimiento de compra por invitación con lista corta.pdf",
      "SBS-P04 Procedimiento de compra por invitación abierta.pdf",
      "SBS-P06 Procedimiento pólizas clientes.pdf",
      "SBS-P07 Procedimiento liquidaciones de contratos.pdf",
      "SBS-P08 Procedimiento Registro de Proveedores.pdf",
      "SBS-P09 Procedimiento Publicación SECOP.pdf",
    ],
  },
};

// Carpeta del repositorio → sigla del proceso en el portal.
const CARPETA_A_SIGLA = {
  'Orientación Estratégica': 'OE', 'Planeación Estratégica': 'OE',
  'Gestión Comercial': 'POL', 'Gestión de Operaciones': 'GOP',
  'Gestión Financiera': 'GF', 'Gestión Administrativa': 'GA',
  'Gestión Documental': 'GD', 'Gestión del Talento Humano': 'GT',
  'Gestión de Tecnologías de Información': 'GTI', 'Proceso de Contratación': 'BS',
  'Gestión de Procesos y Mejoramiento Continuo': 'MA', 'Gestión de Control Interno': 'CI',
  'Gestión Proceso Disciplinario y juzgamiento': 'PD', 'Modelo Integrado de Planeación y Gestión': 'MIPG',
};

// Convierte el inventario del repositorio en entradas de búsqueda: extrae el
// código del nombre del archivo (si lo tiene) y arma el enlace directo.
const entradasRepositorio = () => {
  const out = [];
  for (const [carpeta, rutas] of Object.entries(REPO_ARCHIVOS)) {
    for (const [ruta, nombres] of Object.entries(rutas)) {
      for (const archivo of nombres) {
        const base = archivo.replace(/\.[A-Za-z0-9]+$/, '');
        const m = /^([A-Z]{2,5}- ?[A-Z]{0,4}\d+[A-Z]?)[\s.-]*(.*)$/.exec(base);
        out.push({
          codigo: m ? m[1].replace(/\s+/g, '') : null,
          nombre: m && m[2] ? m[2] : base,
          proceso: CARPETA_A_SIGLA[carpeta],
          origen: CARPETA_A_SIGLA[carpeta] ? undefined : carpeta,
          url: enlaceCarpeta(carpeta, ruta || undefined) + '/' + encodeURIComponent(archivo),
        });
      }
    }
  }
  return out;
};

// Índice de búsqueda: absolutamente todos los documentos del sistema.
// Une el inventario curado de "Documentos Word", las caracterizaciones, los
// informes de referencia y el rastreo completo del repositorio; se eliminan
// duplicados por URL (gana la entrada curada, que tiene mejor nombre).
const DOCS_TODOS = (() => {
  const candidatos = [
    ...DOCUMENTOS.map((d) => ({ ...d, url: enlaceDoc(d.archivo) })),
    ...Object.entries(CARACTERIZACIONES).map(([sigla, c]) => ({
      codigo: c.codigo,
      nombre: `Caracterización ${PROCESOS.find((p) => p.sigla === sigla)?.nombre || sigla}`,
      proceso: sigla,
      url: c.archivo ? enlaceCarpeta(c.archivo[0]) + '/' + encodeURIComponent(c.archivo[1]) : null,
      ruta: `proceso/${sigla}`,
    })),
    { codigo: 'GA-F015', nombre: 'Certificado de cumplimiento de actividades extramural', proceso: 'GA',
      url: enlaceCarpeta('Gestión Administrativa', 'Formatos') + '/' + encodeURIComponent('GA-F015 Certificado de Cumplimiento.docx') },
    { codigo: null, nombre: 'Seguimiento a la Gestión del Riesgo — Mapa de Riesgos (vigencia 2025)', origen: 'Documentos del SGC', url: ENLACE_SEGUIMIENTO_RIESGOS, ruta: 'riesgos' },
    { codigo: null, nombre: 'Indicadores para Plan Estratégico', origen: 'Sitio Plan Estratégico', url: ENLACE_INDICADORES_PLAN, ruta: 'indicadores' },
    { codigo: null, nombre: 'Listado maestro de documentos', origen: 'Repositorio SGC',
      url: encodeURI(BIBLIOTECA_SGC + '/') + encodeURIComponent('LISTADO MAESTRO DE DOCUMENTOS.xlsx') },
    { codigo: null, nombre: 'Membrete ACTIVA', origen: 'Repositorio SGC',
      url: encodeURI(BIBLIOTECA_SGC + '/') + encodeURIComponent('Membrete ACTIVA.docx') },
    ...entradasRepositorio(),
  ];
  const vistos = new Set();
  return candidatos.filter((d) => {
    const clave = d.url || 'ruta:' + d.ruta + ':' + d.codigo;
    if (vistos.has(clave)) return false;
    vistos.add(clave);
    return true;
  });
})();

// Texto en el que busca cada buscador (código + nombre + proceso/origen).
const textoBuscable = (d) => [d.codigo, d.nombre, PROCESOS.find((p) => p.sigla === d.proceso)?.nombre, d.origen]
  .filter(Boolean).join(' ').toLowerCase();

/* ===== 3. Componentes ===== */

const Codigo = ({ children }) => children
  ? <span className="f-mono text-xs font-bold px-1.5 py-0.5 rounded bg-[#14231B] text-[#B5E048]">{children}</span>
  : <span className="f-mono text-xs px-1.5 py-0.5 rounded bg-[#DCE5DC] text-[#5b6b5f]">sin código</span>;

const FilaDocumento = ({ doc }) => {
  // Destino: archivo en SharePoint si existe; si no, la vista del portal.
  const url = doc.url || (doc.archivo && !Array.isArray(doc.archivo) ? enlaceDoc(doc.archivo) : null);
  return (
    <div className="tarjeta flex items-center gap-3 bg-white rounded-xl border border-[#DCE5DC] px-4 py-3">
      <Codigo>{doc.codigo}</Codigo>
      <div className="flex-1 min-w-0">
        <p className="font-medium leading-snug break-words">{doc.nombre}</p>
        <p className="text-xs text-[#5b6b5f]">{[tipoDeCodigo(doc.codigo), PROCESOS.find((p) => p.sigla === doc.proceso)?.nombre || doc.origen].filter(Boolean).join(' · ')}</p>
      </div>
      {url
        ? <a href={url} target="_blank" rel="noopener"
             className="text-sm font-semibold text-[#1E6B47] hover:text-[#144D33] whitespace-nowrap">Abrir ↗</a>
        : <a href={'#/' + doc.ruta}
             className="text-sm font-semibold text-[#1E6B47] hover:text-[#144D33] whitespace-nowrap">Ver ↗</a>}
    </div>
  );
};

const Buscador = ({ irA }) => {
  const [q, setQ] = useState('');
  // Busca sobre TODOS los documentos del sistema (DOCS_TODOS), sin tope de
  // resultados: si hay muchas coincidencias, la lista se desplaza.
  const resultados = useMemo(() => {
    const t = q.trim().toLowerCase();
    if (t.length < 2) return [];
    return DOCS_TODOS.filter((d) => textoBuscable(d).includes(t));
  }, [q]);
  return (
    <div className="relative">
      <input value={q} onChange={(e) => setQ(e.target.value)}
        placeholder="Buscar por código o nombre…"
        aria-label="Buscar documento del SGC"
        className="w-full rounded-2xl border-2 border-[#14231B] bg-white px-5 py-4 text-base shadow-[4px_4px_0_#14231B] focus:outline-none focus:border-[#1E6B47]" />
      {resultados.length > 0 && (
        <div className="absolute z-20 mt-2 w-full bg-white border border-[#DCE5DC] rounded-2xl shadow-xl overflow-hidden">
          <div className="max-h-80 overflow-y-auto">
            {resultados.map((d) => d.url ? (
              <a key={d.url} href={d.url} target="_blank" rel="noopener" onClick={() => setQ('')}
                className="w-full text-left px-4 py-3 hover:bg-[#F7F8F4] flex items-center gap-3 border-b border-[#F0F3EE] last:border-0">
                <Codigo>{d.codigo}</Codigo>
                <span className="text-sm flex-1 min-w-0 break-words">{d.nombre}</span>
                <span className="text-sm font-semibold text-[#1E6B47] whitespace-nowrap">Abrir ↗</span>
              </a>
            ) : (
              <a key={'#/' + d.ruta} href={'#/' + d.ruta} onClick={() => setQ('')}
                className="w-full text-left px-4 py-3 hover:bg-[#F7F8F4] flex items-center gap-3 border-b border-[#F0F3EE] last:border-0">
                <Codigo>{d.codigo}</Codigo>
                <span className="text-sm flex-1 min-w-0 break-words">{d.nombre}</span>
                <span className="text-sm font-semibold text-[#1E6B47] whitespace-nowrap">Ver ↗</span>
              </a>
            ))}
          </div>
          <p className="text-[11px] text-[#5b6b5f] px-4 py-1.5 border-t border-[#F0F3EE] bg-[#F7F8F4]">{resultados.length} de {DOCS_TODOS.length} documentos del sistema</p>
        </div>
      )}
    </div>
  );
};

// Envuelve una tabla ancha con desplazamiento horizontal y, en móvil, una
// sombra en el borde derecho + micro-etiqueta que avisan que se puede deslizar.
const TablaScroll = ({ className = '', children }) => (
  <div className="tabla-scroll">
    <div className={`tabla-scroll-vp overflow-x-auto ${className}`}>{children}</div>
    <span className="tabla-hint" aria-hidden="true">Desliza para ver todo →</span>
  </div>
);

// El mapa de procesos como diagrama clásico y clicable: franjas horizontales,
// flechas laterales de entrada (necesidades) y salida (satisfacción), y la
// banda transversal del MIPG. Cada nodo navega a la vista del proceso.
const NodoProceso = ({ p, irA }) => (
  <button onClick={() => irA(`proceso/${p.sigla}`)}
    aria-label={`Entrar al proceso ${p.nombre}`}
    className="nodo-proceso">
    <span className="flex-1 text-left leading-tight">{p.nombre}</span>
    <span className="f-mono text-[10px] font-bold opacity-70">{p.sigla}</span>
  </button>
);

const FranjaMapa = ({ franja, clase, irA }) => (
  <div className={`franja-mapa ${clase}`}>
    <p className="franja-titulo">{franja.nombre}</p>
    <div className="franja-nodos">
      {franja.procesos.map((p) => <NodoProceso key={p.sigla} p={p} irA={irA} />)}
    </div>
  </div>
);

const MapaProcesos = ({ irA }) => (
  <nav className="mapa-marco" aria-label="Mapa de procesos del SGC">
    <FranjaMapa franja={FRANJAS[0]} clase="f-estrategicos" irA={irA} />
    <div className="mapa-centro">
      <div className="flecha-lateral flecha-entrada" aria-hidden="true">
        <span>Necesidades y expectativas</span>
      </div>
      <div className="flex-1 min-w-0 space-y-2">
        <FranjaMapa franja={FRANJAS[1]} clase="f-misionales" irA={irA} />
        <FranjaMapa franja={FRANJAS[2]} clase="f-apoyo" irA={irA} />
        <FranjaMapa franja={FRANJAS[3]} clase="f-evaluacion" irA={irA} />
      </div>
      <div className="flecha-lateral flecha-salida" aria-hidden="true">
        <span>Satisfacción de los grupos de valor</span>
      </div>
    </div>
    <button onClick={() => irA('proceso/MIPG')} className="banda-mipg"
      aria-label="Entrar al Modelo Integrado de Planeación y Gestión (MIPG)">
      Modelo Integrado de Planeación y Gestión — MIPG
      <span className="font-normal opacity-80"> · transversal a todos los procesos</span>
    </button>
    <p className="text-xs text-[#5b6b5f] text-center mt-2 no-print">
      Haz clic en un proceso para ver sus procedimientos, formatos y manuales.
    </p>
  </nav>
);

const Flujograma = ({ pasos }) => (
  <div className="sendero">
    {pasos.map((paso, i) => {
      const rol = ROLES[paso.rol];
      const clase = paso.tipo === 'decision' ? 'decision'
        : paso.tipo === 'inicio' ? 'extremo' : paso.tipo === 'fin' ? 'extremo fin' : '';
      return (
        <div key={i} className={`nodo ${clase} pb-6`} style={{ '--i': i }}>
          <div className="marca text-sm" style={{ color: paso.tipo === 'inicio' || paso.tipo === 'fin' ? undefined : rol.color }}>
            <span>{paso.tipo === 'decision' ? '?' : paso.tipo === 'fin' ? '✓' : i + 1}</span>
          </div>
          <div className="bg-white rounded-2xl border border-[#DCE5DC] p-4 shadow-sm">
            <div className="flex flex-wrap items-center gap-2 mb-1">
              <h4 className="f-display font-semibold text-base leading-snug flex-1">{paso.titulo}</h4>
              {paso.plazo && (
                <span className="text-xs font-bold text-[#14231B] bg-[#F4A93C] rounded-full px-2.5 py-0.5">⏱ {paso.plazo}</span>
              )}
            </div>
            <p className="text-sm text-[#3c4a40] leading-relaxed">{paso.detalle}</p>
            {paso.enlace && (
              <a href={paso.enlace.url} target="_blank" rel="noopener"
                 className="mt-2 inline-flex items-center gap-1.5 text-sm font-semibold text-[#1E6B47] bg-[#B5E048]/25 border border-[#1E6B47]/40 rounded-lg px-3 py-1.5 hover:bg-[#B5E048]/45">
                🔗 {paso.enlace.texto} ↗
              </a>
            )}
            <div className="mt-2 flex flex-wrap items-center gap-2">
              <span className="inline-flex items-center gap-1.5 text-xs font-semibold rounded-full px-2.5 py-1 text-white" style={{ background: rol.color }}>
                <span className="w-1.5 h-1.5 rounded-full bg-white/80"></span>{rol.nombre}
              </span>
            </div>
            {paso.tipo === 'decision' && (
              <div className="mt-3 grid grid-cols-1 sm:grid-cols-2 gap-2">
                <div className="rounded-xl border-2 border-[#1E6B47] bg-[#1E6B47]/5 px-3 py-2 text-sm">
                  <span className="font-bold text-[#1E6B47]">Sí →</span> {paso.si}
                </div>
                <div className="rounded-xl border-2 border-[#F4A93C] bg-[#F4A93C]/10 px-3 py-2 text-sm">
                  <span className="font-bold text-[#8A5A2C]">No →</span> {paso.no}
                </div>
              </div>
            )}
          </div>
        </div>
      );
    })}
  </div>
);

// Instructivo paso a paso con casos (A/B), requisitos previos, avisos y
// comparación final. Se usa en las guías que no son un flujograma lineal.
const Instructivo = ({ data }) => (
  <div className="space-y-5">
    {data.enlaceApp && (
      <a href={data.enlaceApp} target="_blank" rel="noopener"
         className="inline-flex items-center gap-2 rounded-xl bg-[#1E6B47] text-white font-semibold px-4 py-2.5 hover:bg-[#144D33] shadow-[3px_3px_0_#14231B]">
        🔗 Abrir el aplicativo de Órdenes de Pedido ↗
      </a>
    )}
    {data.enlace && (
      <a href={data.enlace.url} target="_blank" rel="noopener"
         className="inline-flex items-center gap-2 rounded-xl bg-[#1E6B47] text-white font-semibold px-4 py-2.5 hover:bg-[#144D33] shadow-[3px_3px_0_#14231B]">
        🔗 {data.enlace.texto} ↗
      </a>
    )}
    {(data.casos || []).map((c, i) => (
      <div key={i} className="bg-white rounded-2xl border border-[#DCE5DC] p-4 sm:p-5 shadow-sm">
        <h3 className="f-display text-lg font-bold text-[#14231B] mb-3">{c.titulo}</h3>
        {c.intro && <p className="text-sm text-[#3c4a40] leading-relaxed mb-2">{c.intro}</p>}
        {c.antes && (
          <div className="mb-3 rounded-xl bg-[#DCE5DC]/40 border border-[#DCE5DC] px-3 py-2">
            <p className="text-sm font-semibold text-[#14231B] mb-0.5">Antes de empezar</p>
            <p className="text-sm text-[#3c4a40] leading-relaxed">{c.antes}</p>
            {c.antesSub && (
              <ul className="mt-1.5 list-disc pl-5 text-sm text-[#3c4a40] space-y-1">
                {c.antesSub.map((s, j) => <li key={j}>{s}</li>)}
              </ul>
            )}
          </div>
        )}
        {c.pasos && (
          <ol className="list-decimal pl-5 space-y-2 text-sm text-[#3c4a40] marker:font-bold marker:text-[#1E6B47]">
            {c.pasos.map((p, j) => {
              const paso = typeof p === 'string' ? { texto: p } : p;
              return (
                <li key={j} className="leading-relaxed">
                  {paso.texto}
                  {paso.sub && (
                    <ul className="mt-1 list-disc pl-5 space-y-0.5 text-[#3c4a40]">
                      {paso.sub.map((s, k) => <li key={k}>{s}</li>)}
                    </ul>
                  )}
                  {paso.aviso && (
                    <p className="mt-1.5 rounded-lg bg-[#F4A93C]/15 border border-[#F4A93C]/50 px-2.5 py-1.5 text-[#8A5A2C]">⚠️ {paso.aviso}</p>
                  )}
                </li>
              );
            })}
          </ol>
        )}
        {c.tabla && (
          <TablaScroll className="mt-2">
            <table className="w-full text-sm border-collapse">
              <thead>
                <tr>
                  {c.tabla.columnas.map((col, k) => (
                    <th key={k} className="text-left font-semibold text-[#14231B] bg-[#DCE5DC]/50 border border-[#DCE5DC] px-2.5 py-1.5">{col}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {c.tabla.filas.map((fila, fi) => (
                  <tr key={fi}>
                    {fila.map((cel, ci) => (
                      <td key={ci} className={`align-top border border-[#DCE5DC] px-2.5 py-1.5 ${ci === 0 ? 'font-medium text-[#14231B]' : 'text-[#3c4a40]'}`}>{cel}</td>
                    ))}
                  </tr>
                ))}
              </tbody>
            </table>
          </TablaScroll>
        )}
        {c.pie && <p className="mt-1.5 text-xs text-[#5b6b5f]">{c.pie}</p>}
        {c.aviso && (
          <p className="mt-2 rounded-lg bg-[#F4A93C]/15 border border-[#F4A93C]/50 px-2.5 py-1.5 text-sm text-[#8A5A2C]">⚠️ {c.aviso}</p>
        )}
        {c.flujo && (
          <p className="mt-3 rounded-xl bg-[#1E6B47]/10 border border-[#1E6B47]/25 px-3 py-2 text-sm text-[#14231B] leading-relaxed">➡️ {c.flujo}</p>
        )}
        {c.nota && (
          <p className="mt-2 rounded-xl bg-[#B5E048]/25 border border-[#B5E048]/60 px-3 py-2 text-sm text-[#14231B] leading-relaxed">💡 {c.nota}</p>
        )}
      </div>
    ))}
    {data.resumenFinal && (
      <div className="rounded-2xl bg-[#B5E048]/25 border-2 border-[#1E6B47]/30 p-4">
        <p className="f-mono text-xs font-bold text-[#1E6B47] uppercase tracking-widest mb-1">Resumen rápido</p>
        <p className="text-sm text-[#14231B] leading-relaxed">{data.resumenFinal}</p>
      </div>
    )}
    {data.comparacion && (
      <div className="bg-[#14231B] text-[#F7F8F4] rounded-2xl p-5">
        <p className="f-display font-semibold mb-2">{data.comparacion.titulo}</p>
        <ul className="space-y-1.5 text-sm text-white/90 list-disc pl-5">
          {data.comparacion.items.map((it, i) => <li key={i}>{it}</li>)}
        </ul>
      </div>
    )}
  </div>
);

const VistaGuia = ({ guia, irA }) => {
  const docs = (guia.docs || []).map((c) => DOCUMENTOS.find((d) => d.codigo === c)).filter(Boolean);
  const formatos = guia.formatos || [];
  const proceso = PROCESOS.find((p) => p.sigla === guia.proceso);
  return (
    <div className="max-w-3xl mx-auto">
      <button onClick={() => irA('')} className="no-print text-sm font-semibold text-[#1E6B47] mb-4">← Volver al inicio</button>
      <p className="f-mono text-xs font-bold text-[#1E6B47] uppercase tracking-widest mb-1">Guía{guia.fuente ? ` · fuente ${guia.fuente}` : ''}</p>
      <h2 className="f-display text-3xl sm:text-4xl font-extrabold leading-tight mb-2">{guia.pregunta}</h2>
      <p className="text-[#3c4a40] mb-6 leading-relaxed">{guia.resumen}</p>
      {guia.instructivo ? <Instructivo data={guia.instructivo} /> : <Flujograma pasos={guia.pasos} />}
      {(docs.length > 0 || formatos.length > 0) && (
        <div className="mt-4 bg-[#14231B] text-[#F7F8F4] rounded-2xl p-5">
          <p className="f-display font-semibold mb-3">Documentos oficiales de esta guía</p>
          <div className="space-y-2">
            {docs.map((d) => (
              <div key={d.codigo} className="flex items-center gap-3 text-sm">
                <span className="f-mono text-xs font-bold text-[#B5E048]">{d.codigo}</span>
                <span className="flex-1 text-white/90">{d.nombre}</span>
                {d.estado === 'aprobacion'
                  ? <span className="text-xs text-[#F4A93C]">En aprobación</span>
                  : <a className="font-semibold text-[#B5E048] hover:underline" href={enlaceDoc(d.archivo)} target="_blank" rel="noopener">Abrir ↗</a>}
              </div>
            ))}
            {formatos.map((f) => (
              <div key={f.codigo || f.nombre} className="flex items-center gap-3 text-sm">
                <span className="f-mono text-xs font-bold text-[#B5E048]">{f.codigo || 'Formato'}</span>
                <span className="flex-1 text-white/90">{f.nombre}</span>
                <a className="font-semibold text-[#B5E048] hover:underline" href={f.url} target="_blank" rel="noopener">Abrir ↗</a>
              </div>
            ))}
          </div>
          <p className="text-xs text-white/60 mt-3">Esta guía es un resumen orientativo; ante cualquier diferencia, manda el procedimiento oficial. Proceso dueño: {proceso?.nombre}.</p>
        </div>
      )}
    </div>
  );
};

// Tarjetas de sección: una por subcarpeta real del proceso en SharePoint.
// Si el proceso no tiene subcarpetas (Planeación Estratégica), se ofrece la
// carpeta raíz como única sección.
const SeccionesProceso = ({ proceso }) => {
  const secciones = proceso.secciones.length
    ? proceso.secciones
    : [{ nombre: 'Documentos del proceso', carpeta: '' }];
  return (
    <div className="mt-5">
      <div className="flex items-baseline justify-between gap-3 mb-2">
        <h3 className="f-display text-lg font-semibold">Carpetas del proceso en el repositorio</h3>
        <a href={enlaceCarpeta(proceso.carpeta)} target="_blank" rel="noopener"
           className="text-sm font-semibold text-[#1E6B47] hover:underline whitespace-nowrap">Abrir carpeta ↗</a>
      </div>
      <div className="grid grid-cols-2 sm:grid-cols-3 gap-3">
        {secciones.map((s) => (
          <a key={s.nombre} target="_blank" rel="noopener"
             href={enlaceCarpeta(s.base || proceso.carpeta, s.carpeta || undefined)}
             className="tarjeta bg-white rounded-2xl border-2 border-[#DCE5DC] hover:border-[#1E6B47] p-4 flex flex-col items-start gap-1">
            <span className="text-2xl" aria-hidden="true">{ICONO_SECCION[s.nombre] || '📁'}</span>
            <span className="font-semibold leading-snug break-words hyphens-auto w-full">{s.nombre}</span>
            <span className="text-xs text-[#5b6b5f]">Abrir en SharePoint ↗</span>
          </a>
        ))}
      </div>
    </div>
  );
};

// Primera página de cada proceso: su caracterización.
const VistaCaracterizacion = ({ proceso, c }) => (
  <div>
    <div className="flex flex-wrap items-center gap-2 mb-4">
      <span className="f-mono text-xs font-bold px-2 py-1 rounded bg-[#14231B] text-[#B5E048]">{c.codigo}</span>
      <span className="text-xs font-semibold text-[#5b6b5f]">Versión: {c.version}</span>
      <span className="text-xs text-[#5b6b5f]">· Líder del proceso: {c.lider}</span>
    </div>
    <div className="space-y-3">
      <div className="bg-white rounded-2xl border border-[#DCE5DC] p-4">
        <p className="f-mono text-[10px] font-bold text-[#1E6B47] uppercase tracking-widest mb-1">Objetivo</p>
        <p className="text-sm leading-relaxed">{c.objetivo}</p>
        <p className="f-mono text-[10px] font-bold text-[#1E6B47] uppercase tracking-widest mt-3 mb-1">Alcance</p>
        <p className="text-sm leading-relaxed">{c.alcance}</p>
        <p className="f-mono text-[10px] font-bold text-[#1E6B47] uppercase tracking-widest mt-3 mb-1">Resultado relevante</p>
        <p className="text-sm leading-relaxed">{c.resultado}</p>
      </div>
      <div className="bg-white rounded-2xl border border-[#DCE5DC] p-4">
        <p className="f-mono text-[10px] font-bold text-[#1E6B47] uppercase tracking-widest mb-2">Requisitos ISO 9001:2015</p>
        <p className="text-sm leading-relaxed text-[#3c4a40]">{c.requisitosISO}</p>
      </div>
      <div className="bg-white rounded-2xl border border-[#DCE5DC] p-4">
        <p className="f-mono text-[10px] font-bold text-[#1E6B47] uppercase tracking-widest mb-2">Ciclo PHVA</p>
        <TablaScroll>
          <table className="w-full text-xs border-collapse min-w-[52rem]">
            <thead>
              <tr className="bg-[#DCE5DC]/50 text-left">
                <th className="px-2 py-1.5 font-bold w-10">PHVA</th>
                <th className="px-2 py-1.5 font-semibold">Proveedor</th>
                <th className="px-2 py-1.5 font-semibold">Entrada / recurso</th>
                <th className="px-2 py-1.5 font-semibold">Actividades</th>
                <th className="px-2 py-1.5 font-semibold">Producto / salida</th>
                <th className="px-2 py-1.5 font-semibold">Receptor</th>
              </tr>
            </thead>
            <tbody>
              {c.phva.map((f, i) => (
                <tr key={i} className="border-t border-[#F0F3EE] align-top">
                  <td className="px-2 py-1.5"><span className="f-mono font-bold text-white bg-[#1E6B47] rounded px-1.5">{f[0]}</span></td>
                  <td className="px-2 py-1.5 text-[#3c4a40]">{f[1]}</td>
                  <td className="px-2 py-1.5 text-[#3c4a40]">{f[2]}</td>
                  <td className="px-2 py-1.5 font-medium">{f[3]}</td>
                  <td className="px-2 py-1.5 text-[#3c4a40]">{f[4]}</td>
                  <td className="px-2 py-1.5 text-[#3c4a40]">{f[5]}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </TablaScroll>
      </div>
      <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
        <div className="bg-white rounded-2xl border border-[#DCE5DC] p-4">
          <p className="f-mono text-[10px] font-bold text-[#1E6B47] uppercase tracking-widest mb-1.5">Indicadores</p>
          <ul className="text-sm space-y-1 list-disc pl-4 text-[#3c4a40]">{c.indicadores.map((x, i) => <li key={i}>{x}</li>)}</ul>
        </div>
        <div className="bg-white rounded-2xl border border-[#DCE5DC] p-4">
          <p className="f-mono text-[10px] font-bold text-[#B5432E] uppercase tracking-widest mb-1.5">Riesgos</p>
          <ul className="text-sm space-y-1 list-disc pl-4 text-[#3c4a40]">{c.riesgos.map((x, i) => <li key={i}>{x}</li>)}</ul>
        </div>
      </div>
      <div className="bg-white rounded-2xl border border-[#DCE5DC] p-4 text-sm space-y-2">
        <p><span className="font-semibold">Recursos humanos:</span> <span className="text-[#3c4a40]">{c.recursos.humanos}</span></p>
        <p><span className="font-semibold">Recursos físicos:</span> <span className="text-[#3c4a40]">{c.recursos.fisicos}</span></p>
        <p><span className="font-semibold">Documentos asociados:</span> <span className="text-[#3c4a40]">{c.documentos}</span></p>
      </div>
    </div>
  </div>
);

const VistaProceso = ({ sigla, irA }) => {
  const proceso = PROCESOS.find((p) => p.sigla === sigla);
  const docs = DOCUMENTOS.filter((d) => d.proceso === sigla);
  const guias = GUIAS.filter((g) => g.proceso === sigla);
  const caract = CARACTERIZACIONES[sigla];
  // La caracterización es la primera página del proceso; los documentos, la segunda.
  const [pestana, setPestana] = useState('caracterizacion');
  if (!proceso) return <p className="text-center py-10">Proceso no encontrado. <button className="text-[#1E6B47] font-semibold" onClick={() => irA('')}>Volver</button></p>;
  const vistaDocumentos = (
    <div>
      {proceso.enConstruccion ? (
        <div className="text-center py-8">
          <p className="text-5xl mb-3" aria-hidden="true">🚧</p>
          <p className="text-lg font-semibold text-[#8A5A2C]">En construcción</p>
          <p className="text-sm text-[#5b6b5f] mt-1 max-w-md mx-auto">Este proceso está en estructuración: pronto tendrá su carpeta en el repositorio y sus documentos oficiales.</p>
        </div>
      ) : (
        <SeccionesProceso proceso={proceso} />
      )}
      {guias.length > 0 && (
        <div className="mt-5 flex flex-wrap gap-2">
          {guias.map((g) => (
            <button key={g.id} onClick={() => irA(`guia/${g.id}`)}
              className="text-sm font-semibold bg-[#B5E048] text-[#14231B] rounded-full px-4 py-2 hover:bg-[#a3d13a]">
              {g.pregunta}
            </button>
          ))}
        </div>
      )}
      {docs.length > 0 && (
        <div className="mt-6">
          <h3 className="f-display text-lg font-semibold mb-2">Documentos del inventario (Word)</h3>
          <div className="space-y-2">
            {docs.map((d) => <FilaDocumento key={d.codigo + d.nombre} doc={d} />)}
          </div>
        </div>
      )}
    </div>
  );
  return (
    <div className="max-w-4xl mx-auto">
      <button onClick={() => irA('')} className="no-print text-sm font-semibold text-[#1E6B47] mb-4">← Volver al inicio</button>
      <div className="mb-3">
        <p className="f-mono text-xs font-bold text-[#1E6B47] uppercase tracking-widest">{proceso.franja} · {proceso.sigla}</p>
        <h2 className="f-display text-3xl font-extrabold leading-tight">{proceso.nombre}</h2>
      </div>
      {caract ? (
        <div>
          <div className="no-print flex gap-2 mb-4 border-b-2 border-[#DCE5DC]">
            {[['caracterizacion', 'Caracterización'], ['documentos', 'Documentos y carpetas']].map(([id, nombre]) => (
              <button key={id} onClick={() => setPestana(id)}
                className={`text-sm font-bold px-4 py-2 rounded-t-xl ${pestana === id ? 'bg-[#1E6B47] text-white' : 'text-[#5b6b5f] hover:text-[#1E6B47]'}`}>
                {nombre}
              </button>
            ))}
          </div>
          {pestana === 'caracterizacion' ? <VistaCaracterizacion proceso={proceso} c={caract} /> : vistaDocumentos}
        </div>
      ) : vistaDocumentos}
    </div>
  );
};

const ZonaBadge = ({ zona }) => {
  const z = ZONAS_RIESGO[zona] || ZONAS_RIESGO['N/D'];
  return (
    <span className="inline-block text-[10px] font-bold rounded-full px-2 py-0.5 whitespace-nowrap"
      style={{ background: z.color, color: z.texto }}>{zona}</span>
  );
};

// Tabla de KRI compartida entre la vista de riesgos y la de indicadores.
const TablaKris = () => (
  <TablaScroll className="rounded-2xl border border-[#DCE5DC] bg-white">
    <table className="w-full text-sm border-collapse min-w-[38rem]">
      <thead>
        <tr className="bg-[#DCE5DC]/50 text-left">
          <th className="px-3 py-2 font-semibold">KRI</th>
          <th className="px-3 py-2 font-semibold">Indicador</th>
          <th className="px-3 py-2 font-semibold text-[#1E6B47]">Verde</th>
          <th className="px-3 py-2 font-semibold text-[#8A5A2C]">Amarillo</th>
          <th className="px-3 py-2 font-semibold text-[#B5432E]">Rojo</th>
          <th className="px-3 py-2 font-semibold">Frecuencia · Responsable</th>
        </tr>
      </thead>
      <tbody>
        {KRIS.map((k) => (
          <tr key={k[0]} className="border-t border-[#F0F3EE]">
            <td className="px-3 py-2 f-mono text-xs font-bold text-[#1E6B47]">{k[0]}</td>
            <td className="px-3 py-2">{k[1]}</td>
            <td className="px-3 py-2 whitespace-nowrap font-semibold text-[#1E6B47]">{k[2]}</td>
            <td className="px-3 py-2 whitespace-nowrap text-[#8A5A2C]">{k[3]}</td>
            <td className="px-3 py-2 whitespace-nowrap font-semibold text-[#B5432E]">{k[4]}</td>
            <td className="px-3 py-2 text-xs text-[#5b6b5f]">{k[5]}</td>
          </tr>
        ))}
      </tbody>
    </table>
  </TablaScroll>
);

const VistaRiesgos = ({ irA }) => {
  // Conteo por zona inherente para las fichas del resumen
  const porZona = Object.keys(ZONAS_RIESGO).filter((z) => z !== 'N/D')
    .map((z) => ({ zona: z, n: TODOS_RIESGOS.filter((r) => r.zi === z).length }));
  return (
    <div className="max-w-4xl mx-auto">
      <button onClick={() => irA('')} className="no-print text-sm font-semibold text-[#1E6B47] mb-4">← Volver al inicio</button>
      <p className="f-mono text-xs font-bold text-[#1E6B47] uppercase tracking-widest mb-1">Gestión del riesgo · Vigencia 2025</p>
      <h2 className="f-display text-3xl sm:text-4xl font-extrabold leading-tight mb-2">Mapa de riesgos</h2>
      <p className="text-[#3c4a40] leading-relaxed max-w-2xl">
        Seguimiento a la gestión del riesgo con la matriz actualizada al 25/02/2026
        (Informe CI 2022-2025, Rad. 202503000143), en cumplimiento de la política OE-M02.
        {' '}{TODOS_RIESGOS.length} riesgos en 5 componentes.
      </p>
      <a href={ENLACE_SEGUIMIENTO_RIESGOS} target="_blank" rel="noopener"
         className="mt-3 inline-flex items-center gap-2 rounded-xl bg-[#1E6B47] text-white font-semibold px-4 py-2.5 hover:bg-[#144D33] shadow-[3px_3px_0_#14231B]">
        📄 Ver el informe completo de Seguimiento a Riesgos ↗
      </a>

      {/* Resumen por zona inherente */}
      <div className="mt-5 grid grid-cols-2 sm:grid-cols-5 gap-3">
        {porZona.map(({ zona, n }) => {
          const z = ZONAS_RIESGO[zona];
          return (
            <div key={zona} className="rounded-2xl border-2 p-3 text-center" style={{ borderColor: z.color }}>
              <p className="f-display text-3xl font-extrabold" style={{ color: z.color }}>{n}</p>
              <p className="text-[11px] font-bold uppercase tracking-wide" style={{ color: z.color }}>{zona}</p>
              <p className="f-mono text-[10px] text-[#5b6b5f]">P×I {z.rango}</p>
            </div>
          );
        })}
      </div>
      <p className="text-xs text-[#5b6b5f] mt-2">
        Zona inherente (antes de controles). La escala P×I es la de la política OE-M02:
        probabilidad 1–5 e impacto 1–25. Los riesgos en zona Significativo, Alto o Crítico
        exigen plan de acción en el formato MA-F01.
      </p>

      {/* Componentes A–E */}
      {COMPONENTES_RIESGO.map((c) => (
        <section key={c.clave} className="mt-8">
          <div className="flex items-baseline gap-3 mb-2">
            <span className="f-mono text-xs font-bold bg-[#14231B] text-[#B5E048] rounded px-1.5 py-0.5">{c.clave}</span>
            <h3 className="f-display text-xl font-bold">{c.nombre}</h3>
            <span className="text-xs text-[#5b6b5f]">{c.fuente}</span>
          </div>
          <div className="space-y-2">
            {c.riesgos.map((r) => (
              <div key={r.id} className="tarjeta bg-white rounded-xl border border-[#DCE5DC] px-4 py-3 flex flex-wrap items-center gap-x-3 gap-y-1.5">
                <span className="f-mono text-xs font-bold text-[#1E6B47] w-9">{r.id}</span>
                <div className="flex-1 min-w-[14rem]">
                  <p className="text-sm font-medium leading-snug">{r.desc}</p>
                  <p className="text-xs text-[#5b6b5f]">{r.proc} · Tratamiento: {r.trat}</p>
                </div>
                <span className="flex items-center gap-1.5 text-xs text-[#5b6b5f]">
                  <ZonaBadge zona={r.zi} /> → <ZonaBadge zona={r.zr} />
                </span>
              </div>
            ))}
          </div>
        </section>
      ))}
      <p className="text-xs text-[#5b6b5f] mt-2">Inherente → Residual (después de controles). N/D: sin calificación residual en la matriz.</p>

      {/* Indicadores clave de riesgo */}
      <section className="mt-8">
        <h3 className="f-display text-xl font-bold mb-1">Indicadores clave de riesgo (KRI) 2026</h3>
        <p className="text-sm text-[#3c4a40] mb-3 max-w-2xl">
          Alertas tempranas: si un indicador entra en amarillo, el líder notifica a Control
          Interno en máximo 5 días hábiles; en rojo, escala al Comité Institucional de
          Gestión y Desempeño.
        </p>
        <TablaKris />
      </section>

      <div className="mt-6 bg-[#14231B] text-[#F7F8F4] rounded-2xl p-5">
        <p className="f-display font-semibold mb-2">Documentos de referencia</p>
        <div className="space-y-2">
          <div className="flex items-center gap-3 text-sm">
            <span className="f-mono text-xs font-bold text-[#B5E048] whitespace-nowrap">Informe</span>
            <span className="flex-1 text-white/90">Seguimiento a la Gestión del Riesgo — Mapa de Riesgos (vigencia 2025)</span>
            <a className="font-semibold text-[#B5E048] hover:underline" href={ENLACE_SEGUIMIENTO_RIESGOS} target="_blank" rel="noopener">Abrir ↗</a>
          </div>
          <div className="flex items-center gap-3 text-sm">
            <span className="f-mono text-xs font-bold text-[#B5E048] whitespace-nowrap">OE-M02</span>
            <span className="flex-1 text-white/90">Política de administración y gestión del riesgo</span>
            <a className="font-semibold text-[#B5E048] hover:underline" href={enlaceDoc('OE-M02 Politica de administracion gestión del riesgo.docx')} target="_blank" rel="noopener">Abrir ↗</a>
          </div>
        </div>
        <p className="text-xs text-white/60 mt-3">Resumen orientativo del informe de seguimiento; ante cualquier diferencia, manda el documento oficial aprobado por la Dirección.</p>
      </div>
    </div>
  );
};

const VistaDocumentos = ({ irA }) => {
  const [q, setQ] = useState('');
  const t = q.trim().toLowerCase();
  // Lista y busca sobre TODOS los documentos del sistema (mismo índice que
  // el buscador del inicio): inventario + caracterizaciones + informes.
  const docs = DOCS_TODOS.filter((d) => t === '' || textoBuscable(d).includes(t))
    .slice().sort((a, b) => (a.codigo || 'ZZ').localeCompare(b.codigo || 'ZZ'));
  return (
    <div className="max-w-3xl mx-auto">
      <button onClick={() => irA('')} className="text-sm font-semibold text-[#1E6B47] mb-4">← Volver al inicio</button>
      <h2 className="f-display text-3xl font-extrabold mb-4">Todos los documentos <span className="text-[#5b6b5f] text-xl font-semibold">({docs.length})</span></h2>
      <div className="relative mb-4">
        <input value={q} onChange={(e) => setQ(e.target.value)}
          placeholder="Buscar por código o nombre…"
          aria-label="Buscar en el listado de documentos"
          className="w-full rounded-2xl border-2 border-[#14231B] bg-white px-5 py-3 text-base shadow-[3px_3px_0_#14231B] focus:outline-none focus:border-[#1E6B47]" />
        {q && (
          <button onClick={() => setQ('')} aria-label="Limpiar búsqueda"
            className="absolute right-3 top-1/2 -translate-y-1/2 text-[#5b6b5f] hover:text-[#14231B] text-lg font-bold">×</button>
        )}
      </div>
      {docs.length > 0 ? (
        <div className="space-y-2">{docs.map((d) => <FilaDocumento key={d.url || d.ruta} doc={d} />)}</div>
      ) : (
        <p className="text-sm text-[#5b6b5f] py-8 text-center">Sin resultados para «{q}». Prueba con otro código o nombre.</p>
      )}
    </div>
  );
};

const Inicio = ({ irA }) => (
  <div className="max-w-4xl mx-auto">
    {/* El mapa de procesos es la puerta de entrada: va arriba de todo */}
    <div className="pt-6 pb-4 text-center">
      <p className="f-mono text-xs font-bold text-[#1E6B47] uppercase tracking-[.2em] mb-1">Sistema de Gestión de Calidad · ACTIVA</p>
      <h1 className="f-display text-3xl sm:text-4xl font-extrabold leading-tight">Mapa de procesos</h1>
    </div>
    <MapaProcesos irA={irA} />

    <h2 className="f-display text-2xl sm:text-3xl font-extrabold leading-tight mt-10 mb-4">
      ¿Cómo pido <span className="text-[#1E6B47]">lo que necesito</span>?
    </h2>
    <div className="flex flex-wrap gap-2 mb-6">
      {GUIAS.map((g) => (
        <button key={g.id} onClick={() => irA(`guia/${g.id}`)}
          className="tarjeta text-sm font-semibold bg-[#B5E048] text-[#14231B] rounded-full px-4 py-2">
          {g.pregunta}
        </button>
      ))}
    </div>
    <Buscador irA={irA} />
    <div className="mt-3 flex justify-end gap-5">
      <button onClick={() => irA('riesgos')} className="text-sm font-semibold text-[#1E6B47] hover:underline">
        Mapa de riesgos →
      </button>
      <button onClick={() => irA('documentos')} className="text-sm font-semibold text-[#1E6B47] hover:underline">
        Ver todos los documentos →
      </button>
    </div>
  </div>
);

// Indicadores del Plan Estratégico 2025-2027 + KRI, solo de los documentos oficiales.
const VistaIndicadores = ({ irA }) => (
  <div className="max-w-4xl mx-auto">
    <button onClick={() => irA('')} className="no-print text-sm font-semibold text-[#1E6B47] mb-4">← Volver al inicio</button>
    <div className="mb-4">
      <p className="f-mono text-xs font-bold text-[#1E6B47] uppercase tracking-widest">Seguimiento y medición</p>
      <h2 className="f-display text-3xl font-extrabold leading-tight">Indicadores</h2>
      <p className="text-sm text-[#3c4a40] mt-2 max-w-2xl">
        Metas 2025–2027 del Plan Estratégico de ACTIVA, por línea estratégica, e indicadores
        clave de riesgo (KRI) del seguimiento al mapa de riesgos.
      </p>
    </div>
    {LINEAS_INDICADORES.map((l) => (
      <section key={l.linea} className="mb-6">
        <h3 className="f-display text-xl font-bold">{l.linea}</h3>
        <p className="text-sm text-[#5b6b5f] italic mb-3">{l.lema}</p>
        <TablaScroll className="rounded-2xl border border-[#DCE5DC] bg-white">
          <table className="w-full text-sm border-collapse min-w-[36rem]">
            <thead>
              <tr className="bg-[#DCE5DC]/50 text-left">
                <th className="px-3 py-2 font-semibold min-w-[13rem]">Indicador</th>
                <th className="px-3 py-2 font-semibold whitespace-nowrap">Meta 2025</th>
                <th className="px-3 py-2 font-semibold whitespace-nowrap">Meta 2026</th>
                <th className="px-3 py-2 font-semibold whitespace-nowrap">Meta 2027</th>
              </tr>
            </thead>
            <tbody>
              {l.indicadores.map((i) => (
                <tr key={i[0]} className="border-t border-[#F0F3EE]">
                  <td className="px-3 py-2 min-w-[13rem]">{i[0]}</td>
                  <td className="px-3 py-2 font-semibold text-[#1E6B47]">{i[1]}</td>
                  <td className="px-3 py-2 font-semibold text-[#1E6B47]">{i[2]}</td>
                  <td className="px-3 py-2 font-semibold text-[#1E6B47]">{i[3]}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </TablaScroll>
      </section>
    ))}
    <section className="mt-8">
      <h3 className="f-display text-xl font-bold mb-1">Indicadores clave de riesgo (KRI) 2026</h3>
      <p className="text-sm text-[#3c4a40] mb-3 max-w-2xl">
        Alertas tempranas del seguimiento al mapa de riesgos. El detalle de los riesgos que
        vigilan está en la <button onClick={() => irA('riesgos')} className="font-semibold text-[#1E6B47] hover:underline">sección de riesgos</button>.
      </p>
      <TablaKris />
    </section>
    <div className="mt-6 bg-[#14231B] text-[#F7F8F4] rounded-2xl p-5">
      <p className="f-display font-semibold mb-2">Documentos de referencia</p>
      <div className="space-y-2">
        <div className="flex items-center gap-3 text-sm">
          <span className="f-mono text-xs font-bold text-[#B5E048] whitespace-nowrap">Plan</span>
          <span className="flex-1 text-white/90">Indicadores para Plan Estratégico (sitio Plan Estratégico)</span>
          <a className="font-semibold text-[#B5E048] hover:underline" href={ENLACE_INDICADORES_PLAN} target="_blank" rel="noopener">Abrir ↗</a>
        </div>
        <div className="flex items-center gap-3 text-sm">
          <span className="f-mono text-xs font-bold text-[#B5E048] whitespace-nowrap">Informe</span>
          <span className="flex-1 text-white/90">Seguimiento a la Gestión del Riesgo — Mapa de Riesgos (vigencia 2025)</span>
          <a className="font-semibold text-[#B5E048] hover:underline" href={ENLACE_SEGUIMIENTO_RIESGOS} target="_blank" rel="noopener">Abrir ↗</a>
        </div>
      </div>
      <p className="text-xs text-white/60 mt-3">Transcripción de los documentos oficiales; ante cualquier diferencia, mandan los documentos aprobados por la Dirección.</p>
    </div>
  </div>
);

// Organigrama oficial de ACTIVA: se muestra la imagen institucional tal cual.
const VistaOrganigrama = ({ irA }) => (
  <div className="max-w-4xl mx-auto">
    <button onClick={() => irA('')} className="no-print text-sm font-semibold text-[#1E6B47] mb-4">← Volver al inicio</button>
    <div className="mb-4">
      <p className="f-mono text-xs font-bold text-[#1E6B47] uppercase tracking-widest">Estructura organizacional</p>
      <h2 className="f-display text-3xl font-extrabold leading-tight">Organigrama</h2>
    </div>
    <div className="bg-white rounded-2xl border border-[#DCE5DC] p-3 sm:p-5">
      <a href="organigrama.png" target="_blank" rel="noopener" title="Abrir el organigrama en tamaño completo">
        <img src="organigrama.png" alt="Organigrama de ACTIVA: Junta Directiva, Gerente General y sus equipos (Subgerencia Comercial, Dirección Jurídica, Dirección Administrativa y Financiera, Oficina de Control Interno e Instrucción CID)"
          className="w-full h-auto rounded-xl" />
      </a>
      <div className="mt-3 flex flex-wrap items-center justify-between gap-2">
        <p className="text-xs text-[#5b6b5f]">En el celular, ábrelo en pantalla completa para leer los cargos.</p>
        <a href="organigrama.png" target="_blank" rel="noopener"
          className="text-sm font-semibold text-white bg-[#1E6B47] rounded-full px-4 py-2 hover:bg-[#144D33] whitespace-nowrap">
          Ver en pantalla completa ↗
        </a>
      </div>
    </div>
  </div>
);

/* ===== 4. Aplicación y enrutado por hash ===== */
const rutaActual = () => decodeURIComponent((window.location.hash || '#/').replace(/^#\/?/, ''));

const App = () => {
  const [ruta, setRuta] = useState(rutaActual());
  useEffect(() => {
    const alCambiar = () => { setRuta(rutaActual()); window.scrollTo(0, 0); };
    window.addEventListener('hashchange', alCambiar);
    return () => window.removeEventListener('hashchange', alCambiar);
  }, []);
  // El porqué del hash: permite compartir enlaces directos (#/guia/rp) sin
  // servidor ni router; funciona abriendo el archivo o desde GitHub Pages.
  const irA = (destino) => { window.location.hash = '/' + destino; };

  const [seccion, parametro] = ruta.split('/');
  let contenido;
  if (seccion === 'guia' && GUIAS.some((g) => g.id === parametro)) {
    contenido = <VistaGuia guia={GUIAS.find((g) => g.id === parametro)} irA={irA} />;
  } else if (seccion === 'proceso' && parametro) {
    // PE se fusionó con OE; los enlaces viejos siguen funcionando.
    contenido = <VistaProceso sigla={parametro === 'PE' ? 'OE' : parametro} irA={irA} />;
  } else if (seccion === 'documentos') {
    contenido = <VistaDocumentos irA={irA} />;
  } else if (seccion === 'riesgos') {
    contenido = <VistaRiesgos irA={irA} />;
  } else if (seccion === 'organigrama') {
    contenido = <VistaOrganigrama irA={irA} />;
  } else if (seccion === 'indicadores') {
    contenido = <VistaIndicadores irA={irA} />;
  } else {
    contenido = <Inicio irA={irA} />;
  }

  return (
    <div className="min-h-screen flex flex-col">
      <header className="sticky top-0 z-30 bg-[#F7F8F4]/90 backdrop-blur border-b border-[#DCE5DC]">
        <div className="max-w-4xl mx-auto px-4 py-3 flex flex-wrap items-center gap-x-3 gap-y-1">
          <button onClick={() => irA('')} className="f-display font-extrabold tracking-tight flex items-center gap-2 shrink-0 text-left">
            <span className="w-7 h-7 rounded-lg bg-[#1E6B47] text-[#B5E048] flex items-center justify-center text-sm shrink-0" aria-hidden="true">A</span>
            <span className="text-[12px] sm:text-[13px] leading-tight max-w-[14rem] sm:max-w-[22rem]">Sistema de Gestión de Calidad - Empresa de Parques y Eventos de Antioquia ACTIVA</span>
          </button>
          <nav className="w-full sm:w-auto sm:ml-auto flex flex-wrap items-center gap-x-4 gap-y-1 text-sm font-semibold">
            <button onClick={() => irA('')} className="hover:text-[#1E6B47]">Inicio</button>
            <button onClick={() => irA('organigrama')} className="hover:text-[#1E6B47]">Organigrama</button>
            <button onClick={() => irA('indicadores')} className="hover:text-[#1E6B47]">Indicadores</button>
            <button onClick={() => irA('riesgos')} className="hover:text-[#1E6B47]">Riesgos</button>
            <button onClick={() => irA('documentos')} className="hover:text-[#1E6B47]">Documentos</button>
            <a href={ENLACE_NORMOGRAMA} target="_blank" rel="noopener" className="hover:text-[#1E6B47] inline-flex items-center gap-1">Normograma <span aria-hidden="true">↗</span></a>
          </nav>
        </div>
      </header>
      <main className="flex-1 px-4 py-6">{contenido}</main>
      <footer className="border-t border-[#DCE5DC] bg-white">
        <div className="max-w-4xl mx-auto px-4 py-5 text-sm text-[#5b6b5f] space-y-1">
          <p>Empresa de Parques y Eventos de Antioquia — ACTIVA · Portal de consulta del SGC.</p>
          <p>Los enlaces abren el <a className="font-semibold text-[#1E6B47] hover:underline" href={CARPETA_SGC_DEFAULT} target="_blank" rel="noopener">repositorio oficial en SharePoint</a> y requieren sesión Microsoft institucional.</p>
        </div>
      </footer>
    </div>
  );
};

ReactDOM.createRoot(document.getElementById('root')).render(<App />);
