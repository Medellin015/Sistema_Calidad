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
   ============================================================================ */
const { useState, useMemo, useEffect } = React;

/* ===== 1. Configuración ===== */
const SITIO_SGC =
  'https://activaparquesyeventos.sharepoint.com/sites/SistemadeGestindeCalidad';
const BIBLIOTECA_SGC = SITIO_SGC + '/Documentos compartidos';
const CARPETA_SGC_DEFAULT = BIBLIOTECA_SGC + '/Documentos Word/';
const FECHA_CORTE_DEFAULT = '11/07/2026';

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
  'Actas de Comité': '🗒️', 'Autodiagnósticos': '🩺', 'Políticas MIPG': '📜',
  'Certificados curso de integridad': '🎓',
};

const TIPOS = { P: 'Procedimiento', I: 'Instructivo', F: 'Formato', M: 'Manual / Política' };
const tipoDeCodigo = (codigo) => {
  const m = /^[A-Z]+-([PIFM])/.exec(codigo || '');
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
    { sigla: 'OE', nombre: 'Orientación Estratégica',
      carpeta: 'Orientación Estratégica',
      secciones: [S('Procedimientos'), S('Formatos'), S('Manuales'), S('Plantillas')] },
    { sigla: 'PE', nombre: 'Planeación Estratégica',
      carpeta: 'Planeación Estratégica',
      secciones: [] }, // los documentos (plan estratégico, PTEP…) están sueltos en la carpeta
  ]},
  { nombre: 'Misionales', desc: 'La razón de ser: parques y eventos', procesos: [
    { sigla: 'POL', nombre: 'Gestión Comercial',
      carpeta: 'Gestión Comercial',
      secciones: [S('Procedimientos'), S('Formatos'), S('Instructivos')] },
    { sigla: 'GOP', nombre: 'Gestión de Operaciones',
      carpeta: 'Gestión de Operaciones',
      secciones: [S('Procedimientos'), S('Formatos'), S('Instructivos')] },
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
  { codigo: 'POL-P02', nombre: 'Procedimiento de gestión comercial', archivo: '', proceso: 'POL', estado: 'aprobacion' },
  // — Eventos propios: operación misional. Auditorías: evaluación (EV = Evaluación
  //   independiente en la caracterización EV-C01 de Control Interno).
  { codigo: 'EV-P01', nombre: 'Procedimiento de eventos propios', archivo: 'EV-P01 Procedimiento Eventos Propios.docx', proceso: 'GOP', destacado: true },
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
  { codigo: 'GA-P01', nombre: 'Procedimiento de órdenes de pedido (OP)', archivo: '', proceso: 'GA', estado: 'aprobacion', destacado: true },
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
  { codigo: 'GTI-P02', nombre: 'Copias de respaldo y restauración de información', archivo: '', proceso: 'GTI', estado: 'aprobacion' },
  { codigo: 'GTI-I02', nombre: 'Instructivo Mesa de Ayuda', archivo: 'GTI-I02 Instructivo Mesa de Ayuda.docx', proceso: 'GTI', destacado: true },
  { codigo: 'GTI-I3', nombre: 'Creación de contratos en SAFIX', archivo: 'GTI-I3 Instructivo creación de contratos en SAFIX.docx', proceso: 'GTI' },
  { codigo: 'GTI-I4', nombre: 'Aprobación de informe de evidencias — clientes', archivo: 'GTI-I4 Aprobación informe de Evidencias CLIENTES.docx', proceso: 'GTI' },
  { codigo: 'GTI-I5', nombre: 'Instalación del aplicativo CHIP local', archivo: 'GTI-I5 TUTORIAL INSTALACIÓN APLICATIVO CHIP LOCAL.docx', proceso: 'GTI' },
  { codigo: 'GTI-I6', nombre: 'Agregar un calendario desde el directorio', archivo: 'GTI-I6 INSTRUCTIVO PARA AGREGAR UN CALENDARIO DESDE EL DIRECTORIO.docx', proceso: 'GTI' },
  { codigo: 'GTI-I7', nombre: 'Reserva de carro', archivo: 'GTI-I7 INSTRUCTIVO RESERVA DE CARRO.docx', proceso: 'GTI' },
  { codigo: 'GTI-I8', nombre: 'Reserva de sala de reuniones', archivo: 'GTI-I8 INSTRUCTIVO RESERVA SALA DE REUNIONES.docx', proceso: 'GTI' },
  { codigo: 'GTI-I9', nombre: 'Ingreso a Microsoft 365', archivo: 'GTI-I9 INGRESO AL MICROSOFT 365.docx', proceso: 'GTI' },
  // — Bienes y Servicios
  { codigo: 'BS-P05', nombre: 'Procedimiento de contratación', archivo: 'BS-P05 Procedimiento de contratación.docx', proceso: 'BS', destacado: true },
  { codigo: 'BS-P06', nombre: 'Pólizas de clientes', archivo: 'BS-P06 Procedimiento pólizas clientes.docx', proceso: 'BS' },
  { codigo: 'BS-P07', nombre: 'Liquidaciones de contratos', archivo: 'BS-P07 Procedimiento liquidaciones de contratos.docx', proceso: 'BS' },
  { codigo: 'BS-P08', nombre: 'Registro de proveedores', archivo: 'BS-P08 Procedimiento Registro de Proveedores.docx', proceso: 'BS' },
  { codigo: 'BS-P09', nombre: 'Publicación en SECOP', archivo: 'BS-P09 Procedimiento Publicación SECOP.docx', proceso: 'BS' },
  // — Mejora Continua
  { codigo: 'MA-P01', nombre: 'Gestión de mejoras', archivo: 'MA-P01 procedimiento de Gestión de mejoras.doc', proceso: 'MA' },
  { codigo: 'MA-P02', nombre: 'Control de documentos del SGC', archivo: '', proceso: 'MA', estado: 'aprobacion', destacado: true },
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
    proceso: 'GA', fuente: 'GA-P01',
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
    docs: ['GA-P01', 'GF-P02', 'GF-P03', 'GF-P01', 'GA-P03'],
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

/* ===== 3. Componentes ===== */

const Codigo = ({ children }) => children
  ? <span className="f-mono text-xs font-bold px-1.5 py-0.5 rounded bg-[#14231B] text-[#B5E048]">{children}</span>
  : <span className="f-mono text-xs px-1.5 py-0.5 rounded bg-[#DCE5DC] text-[#5b6b5f]">sin código</span>;

const FilaDocumento = ({ doc }) => (
  <div className="tarjeta flex items-center gap-3 bg-white rounded-xl border border-[#DCE5DC] px-4 py-3">
    <Codigo>{doc.codigo}</Codigo>
    <div className="flex-1 min-w-0">
      <p className="font-medium leading-snug">{doc.nombre}</p>
      <p className="text-xs text-[#5b6b5f]">{tipoDeCodigo(doc.codigo)} · {PROCESOS.find((p) => p.sigla === doc.proceso)?.nombre}</p>
    </div>
    {doc.estado === 'aprobacion'
      ? <span className="text-xs font-semibold text-[#8A5A2C] bg-[#F4A93C]/20 border border-[#F4A93C]/50 rounded-full px-3 py-1 whitespace-nowrap">En aprobación</span>
      : <a href={enlaceDoc(doc.archivo)} target="_blank" rel="noopener"
           className="text-sm font-semibold text-[#1E6B47] hover:text-[#144D33] whitespace-nowrap">Abrir ↗</a>}
  </div>
);

const Buscador = ({ irA }) => {
  const [q, setQ] = useState('');
  const resultados = useMemo(() => {
    const t = q.trim().toLowerCase();
    if (t.length < 2) return [];
    return DOCUMENTOS.filter((d) =>
      d.codigo.toLowerCase().includes(t) || d.nombre.toLowerCase().includes(t)).slice(0, 8);
  }, [q]);
  return (
    <div className="relative">
      <input value={q} onChange={(e) => setQ(e.target.value)}
        placeholder="Busca por código o nombre: GF-P03, viáticos, PQRS…"
        aria-label="Buscar documento del SGC"
        className="w-full rounded-2xl border-2 border-[#14231B] bg-white px-5 py-4 text-base shadow-[4px_4px_0_#14231B] focus:outline-none focus:border-[#1E6B47]" />
      {resultados.length > 0 && (
        <div className="absolute z-20 mt-2 w-full bg-white border border-[#DCE5DC] rounded-2xl shadow-xl overflow-hidden">
          {resultados.map((d) => (
            <button key={d.codigo + d.nombre} onClick={() => { setQ(''); irA(`proceso/${d.proceso}`); }}
              className="w-full text-left px-4 py-3 hover:bg-[#F7F8F4] flex items-center gap-3 border-b border-[#F0F3EE] last:border-0">
              <Codigo>{d.codigo}</Codigo><span className="text-sm">{d.nombre}</span>
            </button>
          ))}
        </div>
      )}
    </div>
  );
};

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
          <div className="mt-2 overflow-x-auto">
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
          </div>
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
             href={s.carpeta ? enlaceCarpeta(proceso.carpeta, s.carpeta) : enlaceCarpeta(proceso.carpeta)}
             className="tarjeta bg-white rounded-2xl border-2 border-[#DCE5DC] hover:border-[#1E6B47] p-4 flex flex-col items-start gap-1">
            <span className="text-2xl" aria-hidden="true">{ICONO_SECCION[s.nombre] || '📁'}</span>
            <span className="font-semibold leading-snug">{s.nombre}</span>
            <span className="text-xs text-[#5b6b5f]">Abrir en SharePoint ↗</span>
          </a>
        ))}
      </div>
    </div>
  );
};

const VistaProceso = ({ sigla, irA }) => {
  const proceso = PROCESOS.find((p) => p.sigla === sigla);
  const docs = DOCUMENTOS.filter((d) => d.proceso === sigla);
  const guias = GUIAS.filter((g) => g.proceso === sigla);
  if (!proceso) return <p className="text-center py-10">Proceso no encontrado. <button className="text-[#1E6B47] font-semibold" onClick={() => irA('')}>Volver</button></p>;
  return (
    <div className="max-w-3xl mx-auto">
      <button onClick={() => irA('')} className="text-sm font-semibold text-[#1E6B47] mb-4">← Volver al inicio</button>
      <div className="mb-1">
        <p className="f-mono text-xs font-bold text-[#1E6B47] uppercase tracking-widest">{proceso.franja} · {proceso.sigla}</p>
        <h2 className="f-display text-3xl font-extrabold leading-tight">{proceso.nombre}</h2>
      </div>
      <SeccionesProceso proceso={proceso} />
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
};

const ZonaBadge = ({ zona }) => {
  const z = ZONAS_RIESGO[zona] || ZONAS_RIESGO['N/D'];
  return (
    <span className="inline-block text-[10px] font-bold rounded-full px-2 py-0.5 whitespace-nowrap"
      style={{ background: z.color, color: z.texto }}>{zona}</span>
  );
};

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
        <div className="overflow-x-auto rounded-2xl border border-[#DCE5DC] bg-white">
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
        </div>
      </section>

      <div className="mt-6 bg-[#14231B] text-[#F7F8F4] rounded-2xl p-5">
        <p className="f-display font-semibold mb-2">Documento marco</p>
        <div className="flex items-center gap-3 text-sm">
          <span className="f-mono text-xs font-bold text-[#B5E048]">OE-M02</span>
          <span className="flex-1 text-white/90">Política de administración y gestión del riesgo</span>
          <a className="font-semibold text-[#B5E048] hover:underline" href={enlaceDoc('OE-M02 Politica de administracion gestión del riesgo.docx')} target="_blank" rel="noopener">Abrir ↗</a>
        </div>
        <p className="text-xs text-white/60 mt-3">Resumen orientativo del informe de seguimiento; ante cualquier diferencia, manda el documento oficial aprobado por la Dirección.</p>
      </div>
    </div>
  );
};

const VistaDocumentos = ({ irA }) => {
  const [filtro, setFiltro] = useState('todos');
  const docs = DOCUMENTOS.filter((d) => filtro === 'todos' || tipoDeCodigo(d.codigo) === filtro)
    .slice().sort((a, b) => (a.codigo || 'ZZ').localeCompare(b.codigo || 'ZZ'));
  return (
    <div className="max-w-3xl mx-auto">
      <button onClick={() => irA('')} className="text-sm font-semibold text-[#1E6B47] mb-4">← Volver al inicio</button>
      <h2 className="f-display text-3xl font-extrabold mb-4">Todos los documentos <span className="text-[#5b6b5f] text-xl font-semibold">({docs.length})</span></h2>
      <div className="flex flex-wrap gap-2 mb-4">
        {['todos', ...Object.values(TIPOS)].map((f) => (
          <button key={f} onClick={() => setFiltro(f)}
            className={`text-sm font-semibold rounded-full px-4 py-1.5 border-2 ${filtro === f ? 'bg-[#14231B] text-[#B5E048] border-[#14231B]' : 'bg-white border-[#DCE5DC] hover:border-[#1E6B47]'}`}>
            {f === 'todos' ? 'Todos' : f}
          </button>
        ))}
      </div>
      <div className="space-y-2">{docs.map((d) => <FilaDocumento key={d.codigo + d.nombre} doc={d} />)}</div>
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

    <h2 className="f-display text-2xl sm:text-3xl font-extrabold leading-tight mt-10">
      ¿Cómo pido <span className="text-[#1E6B47]">lo que necesito</span>?
    </h2>
    <p className="mt-2 mb-4 text-[#3c4a40] max-w-xl leading-relaxed">
      Guías paso a paso con su flujograma y los {DOCUMENTOS.length} documentos
      oficiales del SGC, con enlace directo al repositorio.
    </p>
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
    contenido = <VistaProceso sigla={parametro} irA={irA} />;
  } else if (seccion === 'documentos') {
    contenido = <VistaDocumentos irA={irA} />;
  } else if (seccion === 'riesgos') {
    contenido = <VistaRiesgos irA={irA} />;
  } else {
    contenido = <Inicio irA={irA} />;
  }

  return (
    <div className="min-h-screen flex flex-col">
      <header className="sticky top-0 z-30 bg-[#F7F8F4]/90 backdrop-blur border-b border-[#DCE5DC]">
        <div className="max-w-4xl mx-auto px-4 py-3 flex items-center gap-3">
          <button onClick={() => irA('')} className="f-display font-extrabold text-lg tracking-tight flex items-center gap-2">
            <span className="w-7 h-7 rounded-lg bg-[#1E6B47] text-[#B5E048] flex items-center justify-center text-sm">A</span>
            SGC ACTIVA
          </button>
          <nav className="ml-auto flex gap-4 text-sm font-semibold">
            <button onClick={() => irA('')} className="hover:text-[#1E6B47]">Inicio</button>
            <button onClick={() => irA('riesgos')} className="hover:text-[#1E6B47]">Riesgos</button>
            <button onClick={() => irA('documentos')} className="hover:text-[#1E6B47]">Documentos</button>
          </nav>
        </div>
      </header>
      <main className="flex-1 px-4 py-6">{contenido}</main>
      <footer className="border-t border-[#DCE5DC] bg-white">
        <div className="max-w-4xl mx-auto px-4 py-5 text-sm text-[#5b6b5f] space-y-1">
          <p>Empresa de Parques y Eventos de Antioquia — ACTIVA · Portal de consulta del SGC.</p>
          <p>Los enlaces abren el <a className="font-semibold text-[#1E6B47] hover:underline" href={CARPETA_SGC_DEFAULT} target="_blank" rel="noopener">repositorio oficial en SharePoint</a> y requieren sesión Microsoft institucional.</p>
          <p className="f-mono text-xs">Inventario al corte {FECHA_CORTE_DEFAULT} · Los flujogramas resumen el procedimiento fuente.</p>
        </div>
      </footer>
    </div>
  );
};

ReactDOM.createRoot(document.getElementById('root')).render(<App />);
