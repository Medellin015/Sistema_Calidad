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

// Guías "¿Cómo pido…?". Los pasos resumen el procedimiento fuente; los plazos
// marcados provienen del documento oficial cuando fue posible verificarlo.
const GUIAS = [
  {
    id: 'rp', pregunta: '¿Cómo pido un RP?',
    corta: 'Registro presupuestal para comprometer recursos',
    proceso: 'GF', fuente: 'GF-P03',
    resumen: 'El RP (registro presupuestal) afecta de forma definitiva la apropiación y respalda '
      + 'el compromiso adquirido. Siempre llega después del CDP y del compromiso perfeccionado.',
    pasos: [
      { tipo: 'inicio', titulo: 'Necesitas comprometer recursos', detalle: 'Hay un contrato u obligación por perfeccionar con cargo al presupuesto.', rol: 'solicitante' },
      { tipo: 'decision', titulo: '¿Ya existe el CDP?', detalle: 'Todo RP nace de un CDP vigente que garantiza la apropiación.', rol: 'financiera',
        si: 'Continúa con la solicitud del RP', no: 'Primero tramita el CDP (guía «¿Cómo pido un CDP?», GF-P02)' },
      { tipo: 'paso', titulo: 'Reúne el soporte del compromiso', detalle: 'Contrato suscrito, acto administrativo u orden con valor, objeto y beneficiario definidos, citando el número del CDP.', rol: 'solicitante' },
      { tipo: 'paso', titulo: 'Solicita el RP a Gestión Financiera', detalle: 'Remite la solicitud con los soportes por el canal institucional para su verificación.', rol: 'solicitante' },
      { tipo: 'paso', titulo: 'Verificación y expedición', detalle: 'Financiera valida saldo del CDP, rubro y datos del beneficiario, y expide el RP con su número consecutivo.', rol: 'financiera' },
      { tipo: 'fin', titulo: 'RP expedido', detalle: 'Guarda el número: lo necesitarás en las órdenes de pedido y en el trámite de pagos (GF-P01).', rol: 'financiera' },
    ],
    docs: ['GF-P03', 'GF-P02', 'GF-P01'],
  },
  {
    id: 'cdp', pregunta: '¿Cómo pido un CDP?',
    corta: 'Certificado de disponibilidad presupuestal',
    proceso: 'GF', fuente: 'GF-P02',
    resumen: 'El CDP certifica que existe apropiación presupuestal libre para iniciar un '
      + 'compromiso. Es el primer paso presupuestal de cualquier contratación o gasto.',
    pasos: [
      { tipo: 'inicio', titulo: 'Identificas una necesidad de gasto', detalle: 'Una contratación, compra o compromiso que afectará el presupuesto de la vigencia.', rol: 'solicitante' },
      { tipo: 'paso', titulo: 'Define objeto y valor estimado', detalle: 'Describe el objeto del gasto y su valor con la mayor precisión posible; identifica el rubro si lo conoces.', rol: 'solicitante' },
      { tipo: 'paso', titulo: 'Radica la solicitud de CDP', detalle: 'Envía la solicitud a Gestión Financiera por el canal institucional con la justificación de la necesidad.', rol: 'solicitante' },
      { tipo: 'decision', titulo: '¿Hay apropiación disponible?', detalle: 'Financiera verifica el saldo libre del rubro en el presupuesto vigente.', rol: 'financiera',
        si: 'Se expide el CDP', no: 'Se requiere traslado o adición presupuestal antes de continuar' },
      { tipo: 'fin', titulo: 'CDP expedido', detalle: 'Con el número de CDP puedes adelantar el proceso de contratación (BS-P05) y luego solicitar el RP.', rol: 'financiera' },
    ],
    docs: ['GF-P02', 'GF-P03', 'BS-P05'],
  },
  {
    id: 'comision', pregunta: '¿Cómo pido una comisión?',
    corta: 'Viáticos y gastos de viaje',
    proceso: 'GA', fuente: 'GA-P02',
    resumen: 'Las comisiones de servicio y sus viáticos se tramitan antes del desplazamiento y '
      + 'se legalizan al regresar. Los valores y plazos exactos están en GA-P02.',
    pasos: [
      { tipo: 'inicio', titulo: 'Surge el desplazamiento', detalle: 'Una actividad institucional fuera de la sede requiere tu presencia.', rol: 'solicitante' },
      { tipo: 'paso', titulo: 'Solicita la comisión con anticipación', detalle: 'Diligencia la solicitud con destino, fechas, objeto de la comisión y transporte requerido, antes de viajar.', rol: 'solicitante' },
      { tipo: 'paso', titulo: 'Autorización del superior', detalle: 'Tu jefe inmediato o la Dirección avala la pertinencia y la disponibilidad para el desplazamiento.', rol: 'direccion' },
      { tipo: 'paso', titulo: 'Liquidación de viáticos', detalle: 'Gestión Administrativa y Financiera liquidan los viáticos y gastos de viaje según la tabla vigente y gestionan el pago.', rol: 'financiera' },
      { tipo: 'paso', titulo: 'Cumple la comisión', detalle: 'Realiza la actividad y conserva soportes: pasabordos, certificaciones de asistencia y demás evidencias.', rol: 'solicitante' },
      { tipo: 'decision', titulo: '¿Legalizaste a tiempo?', detalle: 'Al regresar, presenta el informe y los soportes dentro del plazo definido en GA-P02.', rol: 'solicitante',
        si: 'Comisión cerrada sin novedades', no: 'No podrás tramitar nuevas comisiones hasta legalizar la pendiente' },
      { tipo: 'fin', titulo: 'Comisión legalizada', detalle: 'El expediente queda completo y habilitas futuros desplazamientos.', rol: 'financiera' },
    ],
    docs: ['GA-P02', 'GF-P01'],
  },
  {
    id: 'op', pregunta: '¿Cómo hago una orden de pedido?',
    corta: 'Bienes y servicios con cargo a contratos vigentes',
    proceso: 'GA', fuente: 'GA-P01',
    resumen: 'Las OP se registran únicamente en el aplicativo institucional de Órdenes de '
      + 'Pedido y se aprueban solo con saldo presupuestal y contractual verificado.',
    pasos: [
      { tipo: 'inicio', titulo: 'Necesitas un bien o servicio', detalle: 'La necesidad está amparada por un contrato de suministro vigente.', rol: 'solicitante' },
      { tipo: 'paso', titulo: 'Registra la OP en el aplicativo', detalle: 'Indica bienes o servicios, cantidades, contrato imputado, justificación y fecha requerida. El consecutivo es automático.', rol: 'solicitante' },
      { tipo: 'decision', titulo: '¿Hay saldo presupuestal y de contrato?', detalle: 'El módulo CDP & CRP verifica el CDP, el RP y el saldo disponible del contrato.', rol: 'financiera',
        si: 'La OP pasa a aprobación', no: 'La OP se devuelve; gestiona adición o tramita por BS-P05' },
      { tipo: 'paso', titulo: 'Aprobación', detalle: 'El aprobador autoriza la OP en el aplicativo; quien registra nunca aprueba (segregación de funciones).', rol: 'direccion' },
      { tipo: 'paso', titulo: 'Envío al proveedor y recepción', detalle: 'El supervisor remite la OP al proveedor; al recibir, verifica cantidades y calidad y registra la recepción con su soporte.', rol: 'area' },
      { tipo: 'fin', titulo: 'Cierre y pago', detalle: 'Con factura y soportes cargados, la OP se cierra y pasa al procedimiento de pagos (GF-P01).', rol: 'financiera' },
    ],
    docs: ['GA-P01', 'GF-P02', 'GF-P03', 'GF-P01', 'GA-P03'],
  },
  {
    id: 'pqrsfd', pregunta: '¿Cómo se gestiona una PQRSFD?',
    corta: 'Peticiones, quejas, reclamos, sugerencias, felicitaciones y denuncias',
    proceso: 'GD', fuente: 'GD-P02',
    resumen: 'Toda PQRSFD se radica en QFDocument y se responde dentro de los términos de ley: '
      + '15 días hábiles la mayoría, 10 para documentos e información, 30 para consultas.',
    pasos: [
      { tipo: 'inicio', titulo: 'Llega la PQRSFD', detalle: 'Por taquilla, buzón físico, teléfono, correo info@activa.com.co, formulario web o redes sociales.', rol: 'documental' },
      { tipo: 'paso', titulo: 'Radicación en QFDocument', detalle: 'Se radica y digitaliza con número consecutivo (ej. 202603000000), que oficializa el trámite y activa los plazos.', rol: 'documental' },
      { tipo: 'paso', titulo: 'Distribución al área competente', detalle: 'A más tardar el día hábil siguiente, la solicitud se remite al área responsable de proyectar la respuesta.', rol: 'documental', plazo: '≤ 1 día hábil' },
      { tipo: 'paso', titulo: 'Proyección de la respuesta', detalle: 'El área responde de fondo, de forma clara, precisa y congruente, dentro del término según el tipo de petición.', rol: 'area', plazo: '15 / 10 / 30 días hábiles' },
      { tipo: 'decision', titulo: '¿Requiere visto bueno jurídico?', detalle: 'Si el caso lo amerita, se solicita revisión con 3 a 5 días hábiles de antelación al vencimiento.', rol: 'juridica',
        si: 'Jurídica revisa y ajusta', no: 'Pasa directo a firma' },
      { tipo: 'paso', titulo: 'Firma de Gerencia', detalle: 'Cuando corresponde, Gerencia revisa y firma la respuesta con la documentación completa y a tiempo.', rol: 'direccion' },
      { tipo: 'paso', titulo: 'Envío al peticionario', detalle: 'La respuesta firmada se envía de inmediato por correo o mensajería, registrando el envío.', rol: 'documental' },
      { tipo: 'fin', titulo: 'Seguimiento mensual', detalle: 'Dentro de los primeros 10 días calendario del mes siguiente se reporta el informe de PQRSFD a Control Interno.', rol: 'documental', plazo: '≤ 10 días calendario' },
    ],
    docs: ['GD-P02', 'GD-P07', 'GD-P01'],
  },
  {
    id: 'docsgc', pregunta: '¿Cómo creo o actualizo un documento del SGC?',
    corta: 'Códigos, versiones y publicación oficial',
    proceso: 'MA', fuente: 'MA-P02',
    resumen: 'Todo documento oficial nace de la plantilla, recibe un código {SIGLA}-{TIPO}{NN} '
      + 'y solo existe en el repositorio de SharePoint tras revisión y aprobación.',
    pasos: [
      { tipo: 'inicio', titulo: 'Identificas la necesidad', detalle: 'Un cambio normativo, operativo o una acción de mejora exige crear o actualizar un documento.', rol: 'area' },
      { tipo: 'paso', titulo: 'Elabora el borrador en la plantilla', detalle: 'Usa la plantilla oficial y diligencia las nueve secciones aplicables del documento.', rol: 'area' },
      { tipo: 'paso', titulo: 'Solicita el código', detalle: 'El apoyo al SGC verifica el consecutivo en el listado maestro y asigna el código {SIGLA}-{TIPO}{NN}.', rol: 'documental' },
      { tipo: 'paso', titulo: 'Revisión del gestor', detalle: 'El gestor del proceso valida pertinencia técnica y coherencia con la operación real; ajusta las veces necesarias.', rol: 'area' },
      { tipo: 'decision', titulo: '¿Aprobado por la Dirección?', detalle: 'El Director Administrativo y Financiero —o Gerencia si es institucional— aprueba la versión.', rol: 'direccion',
        si: 'Se registra en control de versiones', no: 'Vuelve a ajuste con las observaciones' },
      { tipo: 'paso', titulo: 'Publicación en SharePoint', detalle: 'Se publica la versión vigente (Word + PDF) y la anterior pasa al histórico marcada como obsoleta.', rol: 'documental' },
      { tipo: 'fin', titulo: 'Divulgación', detalle: 'El líder del proceso socializa el documento y queda evidencia; el listado maestro se actualiza.', rol: 'area' },
    ],
    docs: ['MA-P02', 'RC-P01'],
  },
  {
    id: 'soporte', pregunta: '¿Cómo pido soporte de TI?',
    corta: 'Fallas, accesos y solicitudes tecnológicas',
    proceso: 'GTI', fuente: 'GTI-I02',
    resumen: 'Los incidentes y solicitudes tecnológicas se gestionan por la Mesa de Ayuda, que '
      + 'registra, prioriza y da trazabilidad a cada caso.',
    pasos: [
      { tipo: 'inicio', titulo: 'Detectas una falla o necesidad', detalle: 'Un equipo, aplicativo, acceso o servicio tecnológico no funciona o requieres uno nuevo.', rol: 'solicitante' },
      { tipo: 'paso', titulo: 'Reporta a la Mesa de Ayuda', detalle: 'Registra el caso por el canal definido en GTI-I02 describiendo qué pasa, desde cuándo y qué afecta.', rol: 'solicitante' },
      { tipo: 'paso', titulo: 'Registro y priorización', detalle: 'TI clasifica el caso (incidente o solicitud) y lo prioriza según el impacto en la operación.', rol: 'ti' },
      { tipo: 'decision', titulo: '¿Se resuelve en primer nivel?', detalle: 'Muchos casos se solucionan de inmediato; otros requieren escalamiento al proveedor o especialista.', rol: 'ti',
        si: 'Solución aplicada y verificada contigo', no: 'Se escala y se te informa el estado' },
      { tipo: 'fin', titulo: 'Caso cerrado', detalle: 'Confirmas que todo funciona; la solución queda documentada para casos futuros.', rol: 'ti' },
    ],
    docs: ['GTI-I02', 'GTI-I9', 'GTI-P02'],
  },
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

const VistaGuia = ({ guia, irA }) => {
  const docs = guia.docs.map((c) => DOCUMENTOS.find((d) => d.codigo === c)).filter(Boolean);
  const proceso = PROCESOS.find((p) => p.sigla === guia.proceso);
  return (
    <div className="max-w-3xl mx-auto">
      <button onClick={() => irA('')} className="no-print text-sm font-semibold text-[#1E6B47] mb-4">← Volver al inicio</button>
      <p className="f-mono text-xs font-bold text-[#1E6B47] uppercase tracking-widest mb-1">Guía · fuente {guia.fuente}</p>
      <h2 className="f-display text-3xl sm:text-4xl font-extrabold leading-tight mb-2">{guia.pregunta}</h2>
      <p className="text-[#3c4a40] mb-6 leading-relaxed">{guia.resumen}</p>
      <Flujograma pasos={guia.pasos} />
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
        </div>
        <p className="text-xs text-white/60 mt-3">Esta guía es un resumen orientativo; ante cualquier diferencia, manda el procedimiento oficial. Proceso dueño: {proceso?.nombre}.</p>
      </div>
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
    <div className="mt-3 text-right">
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
