/* Generado desde app.jsx (fuente editable). No editar a mano: recompilar con `node build.mjs`. */
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
const SITIO_SGC = 'https://activaparquesyeventos.sharepoint.com/sites/SistemadeGestindeCalidad';
const BIBLIOTECA_SGC = SITIO_SGC + '/Documentos compartidos';
const CARPETA_SGC_DEFAULT = BIBLIOTECA_SGC + '/Documentos Word/';
const FECHA_CORTE_DEFAULT = '11/07/2026';
// El porqué: los nombres de archivo traen tildes y espacios; el enlace se
// construye codificando solo el nombre para no romper la ruta de la carpeta.
const enlaceDoc = (archivo) => encodeURI(CARPETA_SGC_DEFAULT) + encodeURIComponent(archivo);
// Enlace a la carpeta de un proceso en SharePoint (o a una de sus subcarpetas).
// encodeURI conserva las barras y codifica tildes y espacios de la ruta.
const enlaceCarpeta = (carpetaProceso, subcarpeta) => encodeURI(BIBLIOTECA_SGC + '/' + carpetaProceso + (subcarpeta ? '/' + subcarpeta : ''));
const ICONO_SECCION = {
    'Procedimientos': '📋', 'Formatos': '📝', 'Manuales': '📘',
    'Instructivos': '🧾', 'Plantillas': '📐', 'Documentos del proceso': '🗃️',
    'Planeación Estratégica': '🗺️',
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
            // Orientación y Planeación se gestionan como un solo proceso estratégico.
            // En SharePoint siguen siendo dos carpetas: la sección "Planeación
            // Estratégica" usa `base` para enlazar la carpeta hermana completa.
            { sigla: 'OE', nombre: 'Orientación y Planeación Estratégica',
                carpeta: 'Orientación Estratégica',
                secciones: [S('Procedimientos'), S('Formatos'), S('Manuales'), S('Plantillas'),
                    { nombre: 'Planeación Estratégica', carpeta: '', base: 'Planeación Estratégica' }] },
        ] },
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
        ] },
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
        ] },
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
        ] },
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
    area: { nombre: 'Área responsable', color: '#2C5D8A' },
    financiera: { nombre: 'Gestión Financiera', color: '#8A5A2C' },
    documental: { nombre: 'Gestión Documental', color: '#5B4A8A' },
    juridica: { nombre: 'Área Jurídica', color: '#7A2C4E' },
    direccion: { nombre: 'Dirección / Gerencia', color: '#14231B' },
    ti: { nombre: 'Gestión TI', color: '#0E6B6B' },
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
    'BAJO': { color: '#1E6B47', texto: '#FFFFFF', rango: '1–4' },
    'MEDIO': { color: '#C9A227', texto: '#14231B', rango: '5–30' },
    'SIGNIFICATIVO': { color: '#E07B2A', texto: '#FFFFFF', rango: '31–60' },
    'ALTO': { color: '#B5432E', texto: '#FFFFFF', rango: '61–80' },
    'CRÍTICO': { color: '#7A1F1F', texto: '#FFFFFF', rango: '81–125' },
    'N/D': { color: '#DCE5DC', texto: '#5b6b5f', rango: '' },
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
        ] },
    { clave: 'B', nombre: 'Transparencia', fuente: 'Matriz oficial · T01–T06', riesgos: [
            R('T01', 'Relac. Corporativo', 'Gestión deficiente de datos abiertos (ITA 2024: 6,1%)', 'MEDIO', 'BAJO', 'Reducir'),
            R('T02', 'Relac. Corporativo', 'Opacidad en información de ejecución contractual publicada en SECOP', 'MEDIO', 'BAJO', 'Reducir'),
            R('T03', 'Orient. Estratégica', 'Incumplimiento de obligaciones de transparencia activa (ITA global 59/100)', 'SIGNIFICATIVO', 'BAJO', 'Reducir'),
            R('T04', 'Planeación', 'Ausencia de Plan de Acción institucional con la estructura requerida por ITA', 'SIGNIFICATIVO', 'BAJO', 'Reducir'),
            R('T05', 'Relac. Corporativo', 'Inaccesibilidad del sitio web para personas con discapacidad (WCAG 2.1)', 'BAJO', 'BAJO', 'Reducir'),
            R('T06', 'Relac. Corporativo', 'Publicación deficiente de información de la entidad (estructura, directorio, hojas de vida)', 'BAJO', 'BAJO', 'Reducir'),
        ] },
    { clave: 'C', nombre: 'Supervisión y Ejecución', fuente: 'Matriz oficial · S01–S03', riesgos: [
            R('S01', 'Bienes y servicios', 'Entrega de carpeta de supervisión incompleta o tardía al supervisor designado', 'ALTO', 'MEDIO', 'Reducir'),
            R('S02', 'Bienes y servicios', 'Supervisor designado sin perfil técnico adecuado para el objeto contractual', 'SIGNIFICATIVO', 'MEDIO', 'Reducir'),
            R('S03', 'Bienes y servicios', 'Incumplimiento de obligaciones contractuales no detectado oportunamente', 'ALTO', 'SIGNIFICATIVO', 'Reducir'),
        ] },
    { clave: 'D', nombre: 'Complementarios', fuente: 'Matriz oficial · X01–X06', riesgos: [
            R('X01', 'Gestión Presupuestal', 'CDP emitido antes de la aprobación del cliente externo (compromiso prematuro)', 'SIGNIFICATIVO', 'SIGNIFICATIVO', 'Reducir'),
            R('X02', 'Op. Logística', 'Órdenes internas (POL-F10) sin aprobación formal del director de área', 'BAJO', 'BAJO', 'Reducir'),
            R('X03', 'Exp. del Cliente', 'Contexto estratégico insuficiente en órdenes externas que genera soluciones inadecuadas', 'SIGNIFICATIVO', 'BAJO', 'Reducir'),
            R('X04', 'Exp. del Cliente', 'Insatisfacción del cliente externo por falta de validación post-ejecución', 'BAJO', 'BAJO', 'Reducir'),
            R('X05', 'Bienes y servicios', 'Estudios previos elaborados después de la decisión de asignación', 'BAJO', 'BAJO', 'Reducir'),
            R('X06', 'Bienes y servicios', 'Falta de re-evaluación anual de proveedores marco', 'BAJO', 'BAJO', 'Reducir'),
        ] },
    { clave: 'E', nombre: 'Estratégicos, Financieros y Tecnológicos', fuente: 'GIR · M01–M08', riesgos: [
            R('M01', 'Planeación y Gestión', 'Incumplimiento de objetivos estratégicos del Plan 2026-2029', 'ALTO', 'SIGNIFICATIVO', 'Reducir'),
            R('M02', 'Gestión Financiera', 'Insuficiencia de ingresos propios / EBITDA negativo sostenido (margen neto -28%)', 'ALTO', 'SIGNIFICATIVO', 'Reducir / Evitar'),
            R('M03', 'Talento Humano', 'Pérdida de conocimiento institucional por alta rotación de contratistas', 'MEDIO', 'BAJO', 'Reducir'),
            R('M04', 'Gestión de Eventos', 'Afectación de imagen institucional por fallas operativas en eventos de alto perfil', 'SIGNIFICATIVO', 'MEDIO', 'Reducir'),
            R('M05', 'Tecnología e Info.', 'Vulnerabilidad de seguridad digital: ERP SAFIX, Microsoft 365 y activos críticos', 'SIGNIFICATIVO', 'MEDIO', 'Reducir'),
            R('M06', 'Gestión de Proyectos', 'Desfinanciamiento o sobrecostos en obras del Volcán de Lodo El Totumo', 'ALTO', 'ALTO', 'Reducir / Compartir'),
            R('M07', 'Gestión de Eventos', 'Eventos climáticos extremos que afecten la operación logística al aire libre', 'MEDIO', 'BAJO', 'Asumir / Reducir'),
            R('M08', 'Control Interno', 'Insuficiencia de capacidad técnica en Control Interno (perfil financiero requerido por Contraloría)', 'SIGNIFICATIVO', 'MEDIO', 'Reducir'),
        ] },
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
// Informe fuente de esta sección: documento oficial en el sitio SGC (carpeta
// "Documentos del SGC"), no en "Documentos Word", por eso el enlace se arma aparte.
const ENLACE_SEGUIMIENTO_RIESGOS = enlaceCarpeta('Documentos del SGC') + '/' + encodeURIComponent('Seguimiento a Mapa de Riesgos.pdf');
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
        acuerdos: [
            ['Satisfacción de clientes', 'Clientes satisfechos en un 80%'],
            ['Equilibrio financiero', 'Punto de equilibrio financiero en $0'],
            ['Cumplimiento de requisitos', 'Ejecución del presupuesto inicial de ingresos al 100%'],
        ],
        phva: [
            ['P', 'Junta Directiva; clientes; proveedores; colaboradores; partners', 'Normatividad aplicable; necesidades y expectativas de partes interesadas; información del contexto y del sector económico', 'Planear estratégicamente: análisis del contexto, definición de la apuesta estratégica, análisis de partes interesadas, alineamiento con procesos, planeación de la implementación y definición de sistemas de medición', 'Estrategia ACTIVA; estructura de procesos', 'Todos los procesos; partes interesadas'],
            ['P', 'Estado colombiano; entes certificadores; todos los procesos', 'Normatividad aplicable; normas técnicas; solicitudes de creación, modificación y eliminación de documentos', 'Documentar políticas y lineamientos: identificación, creación, comunicación, modificación y eliminación de documentos', 'Políticas, instrucciones y formatos', 'Todos los procesos'],
            ['H', 'Estado colombiano; entes certificadores; todos los procesos', 'Normatividad aplicable; normas técnicas; solicitudes de implementación de cambios', 'Gestionar los cambios que surgen: detección, evaluación de impacto/beneficio, acciones de conformidad, recursos, implementación y seguimiento', 'Planes de gestión del cambio cerrados', 'Todos los procesos'],
            ['H', 'Todos los procesos', 'Planes estratégicos o su equivalente; caracterizaciones de los procesos; informes de evaluación y de auditoría', 'Gestionar los riesgos de la organización: identificación, análisis y evaluación de riesgos', 'Matriz de riesgos; mapa de riesgos', 'Todos los procesos'],
            ['V', 'Todos los procesos', 'Informes de gestión de los procesos; registros', 'Hacer seguimiento a planes de ACTIVA', 'Informes de gestión', 'Orientación Estratégica; partes interesadas'],
            ['V', 'Todos los procesos', 'Hojas de vida de indicadores; registros del proceso', 'Medir indicadores', 'Hojas de vida de indicadores; informes de gestión', 'Todos los procesos; partes interesadas'],
            ['A', 'Orientación Estratégica', 'Informes de gestión; hojas de vida de indicadores; matriz de riesgos', 'Gestionar mejoras al proceso y el sistema', 'Planes de mejoramiento', 'Gestión Administrativa; Evaluación Independiente'],
        ],
        indicadores: ['Ejecución presupuestal', 'Punto de equilibrio', 'Satisfacción del cliente'],
        riesgos: ['Reducción extrema de los márgenes del sector', 'Cambios en la tecnología', 'Competidor fuera de serie', 'Cambio de prioridades del cliente', 'Fracaso del nuevo proyecto', 'Estancamiento de mercado', 'Cambios en condiciones tributarias', 'Cambios normativos', 'Recortes presupuestales'],
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
        acuerdos: [
            ['Clientes satisfechos', 'Calificación de satisfacción superior al 80%'],
            ['Proveedores idóneos', 'Calificación del nivel de servicio de proveedores superior al 80%'],
        ],
        phva: [
            ['P', 'Orientación Estratégica; Desarrollo de la Experiencia del Cliente', 'Estrategia ACTIVA; contrato firmado; acta de inicio firmada; registro de aceptación de contrato; orden de pedido', 'Preparar la operación: comunicación y socialización de servicios contratados, verificación de disponibilidad de insumos, aceptación de cotizaciones por parte del cliente y solicitud de compra', 'Registros de comunicación y socialización; orden de pedido aprobada; solicitud de compra; orden de pedido a proveedores', 'Desarrollo de la Experiencia del Cliente; Suministro de Bienes y Servicios; Gestión Financiera'],
            ['H', 'Proyectos de operación logística; Suministro de Bienes y Servicios', 'Orden de servicio aprobada', 'Ejecutar servicios', 'Evidencias de ejecución de contratos', 'Cliente'],
            ['V', 'Proyectos de operación logística', 'Contratos; evidencias de ejecución de contratos', 'Monitorear la ejecución de los servicios', 'Informes de supervisión con soportes y anexos; informes de gestión', 'Desarrollo de la Experiencia del Cliente; Suministro de Bienes y Servicios; Gestión Financiera; Orientación Estratégica'],
            ['V', 'Proyectos de operación logística', 'Plan de acción de mercadeo y ventas; registros del proceso', 'Medir indicadores', 'Hojas de vida de indicadores', 'Orientación Estratégica; Desarrollo de la Experiencia del Cliente; cliente'],
            ['A', 'Desarrollo de la Experiencia del Cliente; Proyectos de Operación Logística', 'Informes de gestión; hojas de vida de indicadores', 'Gestionar mejoras al proceso', 'Planes de mejoramiento', 'Desarrollo de la Experiencia del Cliente; Proyectos de Operación Logística; Evaluación Independiente'],
        ],
        indicadores: ['Satisfacción de clientes', 'Satisfacción con proveedores de insumos'],
        riesgos: ['Insumos inoportunos', 'Insumos no disponibles', 'Insumos inadecuados', 'Servicios que no cumplen características contratadas', 'Solicitudes urgentes por parte de los clientes', 'Proveedores deficientes', 'Productos que perjudican a los clientes'],
        recursos: { humanos: 'Subgerente Comercial; Gestor de Mercadeo y Ventas; Ejecutivo Comercial; Profesional de operación logística de eventos; prestadores de servicios asociados', fisicos: 'SIESA; herramientas ofimáticas; espacios de trabajo y equipos de oficina' },
        documentos: 'Listado maestro de documentos y registros',
        requisitosISO: '7.1.5 Recursos de seguimiento y medición · 8.1 Planificación y control operacional · 8.2 Requisitos para los productos y servicios · 8.4 Control de procesos, productos y servicios suministrados externamente · 8.5.1–8.5.4 Producción y provisión del servicio · 8.6 Liberación · 9.1 Seguimiento, medición, análisis y evaluación',
    },
    GOP: {
        codigo: 'GOP-C01', version: 'Propuesta', estado: 'propuesta',
        proceso: 'Gestión de Operaciones', lider: 'Subgerente Comercial / Gestor de Operaciones',
        objetivo: 'Planear y ejecutar la operación logística de eventos y los eventos propios institucionales y comerciales, garantizando el cumplimiento de los requisitos y expectativas de clientes, socios, aliados y asistentes.',
        alcance: 'Desde la recepción de la orden de pedido o la aprobación del evento propio hasta la entrega del servicio, el informe de supervisión y el cierre con la validación de satisfacción.',
        resultado: 'Eventos y operaciones logísticas ejecutados a satisfacción — Receptores: clientes, asistentes y todos los procesos.',
        acuerdos: [
            ['Satisfacción del cliente', 'Encuesta post-servicio igual o superior al 85% (meta 2026 del Plan Estratégico)'],
            ['Eventos propios', 'Realizar los eventos propios de la anualidad según el Plan Estratégico (1 en 2026, 2 en 2027)'],
        ],
        phva: [
            ['P', 'Gestión Comercial; Orientación y Planeación Estratégica', 'Órdenes de pedido aprobadas; contratos; calendario de eventos', 'Planear la operación del evento o servicio: requerimientos, recursos, proveedores y cronograma (GOP-P10)', 'Plan de operación; solicitudes de cotización', 'Gestión de Operaciones; Proceso de Contratación'],
            ['H', 'Proceso de Contratación; proveedores', 'Bienes y servicios contratados; personal asignado', 'Ejecutar la operación logística y los eventos propios (EV-P01, GOP-P10)', 'Evento o servicio ejecutado; evidencias de ejecución', 'Cliente; asistentes'],
            ['V', 'Gestión de Operaciones', 'Informes de supervisión; encuestas post-servicio', 'Supervisar la ejecución y validar la satisfacción del cliente', 'Informe final de supervisión; resultado de satisfacción', 'Gestión Comercial; Gestión Financiera'],
            ['A', 'Gestión de Operaciones', 'Informes de gestión; hojas de vida de indicadores', 'Gestionar mejoras al proceso', 'Planes de mejoramiento (MA-F01)', 'Mejoramiento Continuo; Control Interno'],
        ],
        indicadores: ['Nivel de satisfacción del cliente medido por encuesta (Plan Estratégico)', 'Eventos propios realizados en la anualidad (Plan Estratégico)'],
        riesgos: ['M04 — Afectación de imagen por fallas operativas en eventos de alto perfil', 'M07 — Eventos climáticos extremos que afecten la operación al aire libre', 'X03 — Contexto estratégico insuficiente en órdenes externas'],
        recursos: { humanos: 'Gestor de Operaciones; profesionales y contratistas de operación logística', fisicos: 'Aplicativo de Órdenes de Pedido; herramientas ofimáticas; equipos e infraestructura de eventos' },
        documentos: 'GOP-P10 Soluciones para el suministro de bienes y servicios; EV-P01 Procedimiento de eventos propios; GOP-I05 Instructivo Comité de Asignación',
        requisitosISO: '8.1 Planificación y control operacional · 8.2 Requisitos para los productos y servicios · 8.5 Producción y provisión del servicio · 9.1.2 Satisfacción del cliente',
    },
    GP: {
        codigo: 'GP-C01', version: 'Propuesta', estado: 'propuesta',
        proceso: 'Gestión de Parques', lider: '[Por definir por la Dirección]',
        objetivo: 'Administrar los parques y espacios recreativos asignados a la entidad, asegurando la optimización de los recursos, la implementación de políticas y procedimientos operativos estándar, y la planificación y ejecución de los contratos correspondientes.',
        alcance: 'Desde la recepción o asignación del parque o espacio recreativo hasta su operación, mantenimiento, atención de visitantes y rendición de informes del contrato.',
        resultado: 'Parques en funcionamiento operados de forma segura y sostenible — Receptores: visitantes, entidades contratantes y la comunidad.',
        acuerdos: [
            ['Parques en funcionamiento', 'Cumplir la meta del Plan Estratégico (2 parques en 2026, 3 en 2027)'],
            ['Atención al visitante', 'PQRSFD de visitantes respondidas dentro de los términos de ley'],
        ],
        phva: [
            ['P', 'Orientación y Planeación Estratégica; entidades contratantes', 'Contratos o convenios de administración de parques; presupuesto; normatividad aplicable', 'Planear la operación del parque: modelo de operación, recursos, tarifas y protocolos', 'Plan de operación del parque', 'Gestión de Parques; Gestión Financiera'],
            ['H', 'Proceso de Contratación; Gestión Administrativa', 'Bienes, servicios y personal para la operación', 'Operar y mantener el parque: apertura, atención de visitantes, seguridad, aseo y mantenimiento', 'Parque en funcionamiento; servicios prestados a visitantes', 'Visitantes; entidad contratante'],
            ['V', 'Gestión de Parques', 'Registros de operación; PQRSFD; informes de supervisión', 'Hacer seguimiento a la operación y a la satisfacción de los visitantes', 'Informes de operación y de contrato', 'Entidad contratante; Orientación y Planeación Estratégica'],
            ['A', 'Gestión de Parques', 'Informes de gestión; hojas de vida de indicadores', 'Gestionar mejoras al proceso', 'Planes de mejoramiento (MA-F01)', 'Mejoramiento Continuo; Control Interno'],
        ],
        indicadores: ['Número de parques en funcionamiento (Plan Estratégico: 2 en 2026, 3 en 2027)'],
        riesgos: ['M06 — Desfinanciamiento o sobrecostos en obras del Volcán de Lodo El Totumo', 'M07 — Eventos climáticos extremos'],
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
        acuerdos: [
            ['Eficiencia', 'Ejecución presupuestal del 100% al final de la vigencia'],
            ['Eficacia', 'Margen de utilidades superior al 5%'],
            ['Transparencia', '0 eventos de corrupción identificados en la vigencia'],
        ],
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
        indicadores: ['Margen EBITDA compuesto', 'Margen neto', '% cumplimiento del presupuesto de ventas (Comercial)', '% cartera vencida', '% ejecución financiera cliente', '% ejecución financiera proveedores', '% oportunidad en el pago a proveedores (inferior a 30 días calendario)', '# de eventos de corrupción identificados'],
        riesgos: ['Ver matriz de riesgos del proceso (C01, C03, M02, X01, entre otros)'],
        recursos: { humanos: 'Director Administrativo y Financiero; Gestor Financiero; Gestor Contable; Gestor de Tesorería', fisicos: 'SIESA (gestión financiera); herramientas ofimáticas; espacios de trabajo y equipos de oficina' },
        documentos: 'Listado maestro de documentos',
        requisitosISO: '4.1 Comprensión de la organización y su contexto · 4.2 Necesidades y expectativas de las partes interesadas · 6.1 Acciones para abordar riesgos y oportunidades · 7.1 Recursos',
    },
    GA: {
        codigo: 'GA-C01', version: '01', estado: 'oficial',
        archivo: ['Gestión Administrativa', 'GA-C01 Caracterización Gestión Administrativa.xls'],
        proceso: 'Gestión administrativa', lider: 'Director(a) Administrativo(a) y Gestor TI',
        objetivo: 'Facilitar la disponibilidad de recursos físicos, tecnológicos, de infraestructura, transporte, documentales y de ambiente de trabajo, idóneos y oportunos.',
        alcance: 'Desde la recepción de la necesidad manifiesta por los diferentes procesos hasta la verificación de la ejecución de planes, la medición de indicadores y las actividades de mejoramiento continuo.',
        resultado: 'Disponibilidad de recursos físicos, tecnológicos, de infraestructura, transporte, documentales y de ambiente de trabajo, idóneos y oportunos — Receptores: todos los procesos.',
        acuerdos: [
            ['Recursos oportunos', '80% de los acuerdos de servicio cumplidos en términos de tiempo'],
            ['Recursos idóneos', '0 eventos de pérdida o daño de la información digital y física'],
            ['Ambiente de trabajo', 'Evaluación de satisfacción de bienestar por encima de 3'],
        ],
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
        indicadores: ['Oportunidad en la resolución de tickets o casos', 'Seguridad de la información', 'Evaluación de bienestar'],
        riesgos: ['Espacios de trabajo deficientes', 'Fallas tecnológicas', 'Pérdida de información', 'Robo de información', 'Daño de información', 'Fallas de infraestructura', 'Demoras'],
        recursos: { humanos: 'Director Administrativo y Financiero; Gestor en Tecnología e Información; Gestor de Área Administrativa; Profesional de Suministro de Bienes y Servicios', fisicos: 'SIESA (gestión de activos); herramientas ofimáticas; espacios de trabajo y equipos de oficina' },
        documentos: 'Listado maestro de documentos y registros',
        requisitosISO: '7.1.3 Infraestructura · 7.1.4 Ambiente para la operación de los procesos',
    },
    GD: {
        codigo: 'GD-C01', version: '01', estado: 'oficial',
        archivo: ['Gestión Documental', 'GD-C01 Caracterización Programa de Gestión Documental.xlsx'],
        proceso: 'Gestión Documental', lider: 'Gestor administrativo',
        objetivo: 'Administrar el flujo documental y establecer la organización de los documentos siguiendo los lineamientos correspondientes del Programa de Gestión Documental en torno a la normativa vigente.',
        alcance: 'Inicia con los requerimientos de gestión documental presentados por usuarios internos y externos, continúa con la radicación y reparto de los mismos y finaliza con la custodia de la documentación teniendo en cuenta las Tablas de Retención Documental y la normatividad archivística vigente.',
        resultado: 'Organización, control, seguridad, accesibilidad y conservación de los expedientes generados en la entidad — Receptores: todos los procesos.',
        acuerdos: [
            ['Respuestas oportunas', '100% de PQRS contestadas de acuerdo con los términos de ley'],
            ['Documentos distribuidos correctamente', 'Total de comunicaciones distribuidas correctamente / total recibidas por ventanilla'],
            ['Transferencias ejecutadas', '100% de transferencias realizadas vs. programadas'],
        ],
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
        indicadores: ['[Pendiente en el documento oficial]'],
        riesgos: ['[Pendiente en el documento oficial]'],
        recursos: { humanos: 'Líder de Archivo; Director Administrativo y Financiero; Gestor en Tecnología e Información; Gestor de Área Administrativa', fisicos: 'Área de archivo; computador' },
        documentos: 'Cronograma de planeación; planes de acción; PGD; Programa de Gestión Documental; PINAR',
        requisitosISO: '7.5.2 Creación y actualización · 7.5.3 Control de la información documentada',
    },
    GT: {
        codigo: 'GT-C01', version: '01', estado: 'oficial',
        archivo: ['Gestión del Talento Humano', 'GT-C01 Caracterización Gestión del Talento Humano.xlsx'],
        proceso: 'Gestión de Talentos', lider: 'Gestora en Talento Humano',
        objetivo: 'Desarrollar y asegurar PERSONAS con alto desempeño, competentes y satisfechas.',
        alcance: 'Inicia con la selección, sostenimiento y finaliza con el retiro del funcionario.',
        resultado: 'Personas con alto desempeño, competentes y satisfechas — Receptores: todos los procesos.',
        acuerdos: [
            ['Alto desempeño', 'Evaluación de desempeño por encima del 80% (trabajadores oficiales)'],
            ['Competentes', 'Acuerdos de gestión para cargos directivos de libre nombramiento (80%); calificación de entrevista mayor o igual a 20'],
            ['Satisfechas', 'Evaluación de satisfacción de bienestar por encima de 3; clima laboral 60%'],
        ],
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
        indicadores: ['Evaluación del desempeño (alto desempeño)', 'Nivel de acuerdos de gestión (competentes)'],
        riesgos: ['Contratación de personal no idóneo para el cargo', 'Errores en la liquidación de nóminas y pago de seguridad social fuera del tiempo legal (controlado con el programa SIESA)'],
        recursos: { humanos: 'Gestora en Talento Humano', fisicos: 'Herramientas ofimáticas; espacios de trabajo' },
        documentos: 'Listado maestro de documentos y registros',
        requisitosISO: '7.1.2 Personas · 7.1.4 Ambiente para la operación de los procesos',
    },
    GTI: {
        codigo: 'GTI-C01', version: 'Propuesta', estado: 'propuesta',
        proceso: 'Gestión de Tecnologías de la Información', lider: 'Profesional TIC',
        objetivo: 'Garantizar la disponibilidad, el soporte y la seguridad de los servicios y activos tecnológicos de la entidad (ERP SAFIX, Microsoft 365 y aplicativos institucionales), gestionando los incidentes y solicitudes a través de la Mesa de Ayuda.',
        alcance: 'Desde la identificación de la necesidad o incidente tecnológico hasta su solución verificada, incluida la administración de activos de información, la seguridad digital y el soporte a los aplicativos institucionales.',
        resultado: 'Servicios tecnológicos disponibles, seguros y con soporte oportuno — Receptores: todos los procesos.',
        acuerdos: [
            ['Soporte oportuno', 'Oportunidad en la resolución de tickets o casos de la Mesa de Ayuda'],
            ['Seguridad digital', '0 incidentes de seguridad sin gestionar (KRI06: verde = 0 incidentes en el período)'],
        ],
        phva: [
            ['P', 'Orientación y Planeación Estratégica', 'Plan Estratégico; Política de Seguridad Digital V2; presupuesto TIC', 'Planear la gestión tecnológica: plan de desarrollo de TI, seguridad digital (MSPI) y adquisiciones tecnológicas (GTI-P01)', 'Plan de desarrollo de TI; plan de acción de seguridad digital', 'Todos los procesos'],
            ['H', 'Todos los procesos', 'Incidentes y solicitudes tecnológicas', 'Gestionar la Mesa de Ayuda: registro, clasificación, priorización, solución o escalamiento (GTI-I02)', 'Casos resueltos y documentados', 'Todos los procesos'],
            ['H', 'Gestión de TI', 'Activos de información; licencias; copias de respaldo', 'Administrar los activos de información, los respaldos y los controles de acceso', 'Activos administrados; copias de respaldo; controles implementados', 'Todos los procesos'],
            ['V', 'Gestión de TI', 'Registros de la Mesa de Ayuda; reportes de incidentes', 'Hacer seguimiento a los indicadores del servicio y a los incidentes de seguridad digital (KRI06)', 'Informes de gestión; reporte de incidentes', 'Control Interno; Comité de Gestión y Desempeño'],
            ['A', 'Gestión de TI', 'Informes de gestión; hallazgos ITA', 'Gestionar mejoras al proceso y al MSPI', 'Planes de mejoramiento (MA-F01)', 'Mejoramiento Continuo; Control Interno'],
        ],
        indicadores: ['Oportunidad en la resolución de tickets (Mesa de Ayuda)', 'KRI06 — Incidentes de seguridad digital registrados y reportados', '% de implementación de ERP (Plan Estratégico)'],
        riesgos: ['M05 — Vulnerabilidad de seguridad digital: ERP SAFIX, Microsoft 365 y activos críticos'],
        recursos: { humanos: 'Profesional TIC', fisicos: 'ERP SAFIX; Microsoft 365; aplicativos institucionales; infraestructura tecnológica' },
        documentos: 'GTI-P01 Proceso de Gestión Tecnológica; GTI-I02 Instructivo Mesa de Ayuda; Política de Uso Aceptable de Recursos Tecnológicos; Política de Seguridad Digital V2',
        requisitosISO: '7.1.3 Infraestructura · 7.5 Información documentada · 8.1 Planificación y control operacional',
    },
    BS: {
        codigo: 'SBS-C01', version: '01', estado: 'oficial',
        archivo: ['Proceso de Contratación', 'SBS- C01 Caracterización Proceso de Contratación.xls'],
        proceso: 'Proceso de Contratación', lider: 'Director(a) Administrativo(a) y Financiero(a) y Director(a) Jurídico(a)',
        objetivo: 'Adquirir los bienes y servicios requeridos para la operación de los procesos de ACTIVA de forma oportuna, eficiente e idónea.',
        alcance: 'Desde la recepción de la necesidad de adquisición de bienes y servicios manifiesta desde los diferentes procesos hasta la verificación de la ejecución del plan anual de adquisiciones, evaluación de proveedores, la medición de indicadores y las actividades de mejoramiento continuo.',
        resultado: 'Bienes y servicios comprados de forma oportuna, eficiente e idónea — Receptores: todos los procesos.',
        acuerdos: [
            ['Oportunidad', '80% de compras realizadas oportunamente (según tiempos de invitación abierta, directa, órdenes de compra y lista corta)'],
            ['Eficiencia', 'Valor de compras inferior al 100% del presupuesto'],
            ['Idoneidad', '100% de bienes y servicios adquiridos con percepción positiva del nivel de servicio'],
        ],
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
        indicadores: ['Oportunidad en compras', 'Eficiencia en las compras', 'Recibo a satisfacción de bienes y servicios comprados'],
        riesgos: ['Ver matriz de riesgos del proceso (C04–C09, S01–S03, X05, X06, entre otros)'],
        recursos: { humanos: 'Director Administrativo y Financiero; Director Jurídico; Profesional de Suministro de Bienes y Servicios; Gestor de Área Administrativa; Gestor Jurídico; interesado en la compra', fisicos: 'Herramientas ofimáticas; espacios de trabajo y equipos de oficina' },
        documentos: 'Listado maestro de documentos',
        requisitosISO: '6.1 Riesgos · 7.1.6 Conocimiento · 7.5 Información documentada · 8.1 Planificación · 8.4 Control de productos y servicios externos · 8.5.3 Propiedad de proveedores · 9.1.3 Análisis y evaluación · 10 Mejora',
    },
    MA: {
        codigo: 'MA-C01', version: '01', estado: 'oficial',
        archivo: ['Gestión de Procesos y Mejoramiento Continuo', 'MA-C01 Caracterizacion Mejoramiento Activo.xlsx'],
        proceso: 'Mejoramiento Activo', lider: 'Jefe de Oficina de Planeación o quien haga sus veces',
        objetivo: 'Implementar acciones encaminadas a la medición, análisis, evaluación y mejora del Sistema de Gestión mediante el diseño y la aplicación de lineamientos y metodologías que permitan la conformidad de productos y servicios, optimizar la gestión institucional y la satisfacción de las partes interesadas y grupos de valor.',
        alcance: 'Inicia con la definición e implementación de lineamientos y metodologías, continúa con la formulación de acciones de mejora, y finaliza con la optimización de los procesos y la satisfacción de las partes interesadas y grupos de valor.',
        resultado: 'Ambiente de mejoramiento continuo con seguimiento a planes de mejora, cambios y evaluaciones, generando mejora a los procesos internos del SGC — Receptores: todos los procesos.',
        acuerdos: [
            ['Cumplimiento de la norma', 'Ejecución de auditorías internas de los procesos al 100%'],
            ['Acciones de mejora', 'Acciones de mejora abiertas entre 0 y 10%'],
            ['Satisfacción al cliente', 'Medición de satisfacción del cliente 80%'],
        ],
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
        indicadores: ['Número de auditorías internas realizadas', '% de acciones de mejora cerradas', 'Evaluación de la satisfacción del cliente'],
        riesgos: ['Auditorías mal ejecutadas', 'Acciones de mejora no efectivas', 'Clientes insatisfechos', 'Cambios que afecten la misión de la entidad'],
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
        acuerdos: [
            ['Cumplimiento', 'Cumplimiento del plan anual de auditorías, seguimientos e informes de Control Interno en un 90%'],
            ['Eficacia', 'Implementación y cierre del 85% de las acciones derivadas de auditorías, seguimientos e informes'],
        ],
        phva: [
            ['P', 'Congreso de la República; organismos de control; procesos del sistema', 'Normas constitucionales y legales vigentes; necesidades de auditoría', 'Elaboración del Plan Anual de Auditorías, Seguimientos e Informes de Control Interno', 'Plan Anual de Auditorías, Seguimientos e Informes', 'Evaluación Independiente; Orientación Estratégica'],
            ['H', 'Todos los procesos', 'Información del proceso, dependencia, proyecto, contrato o tema a auditar', 'Realizar las auditorías, seguimientos e informes de Control Interno; elaborar, presentar y publicar los informes', 'Registros de auditoría (listas de chequeo, notas); informes de auditoría; informes a entes de control', 'Evaluación Independiente; entes de control; todos los procesos'],
            ['V', 'Todos los procesos', 'Riesgos y oportunidades reportados', 'Seguimiento a los controles de los riesgos', 'Reportes de materialización del riesgo; matriz y mapa de riesgos actualizado', 'Todos los procesos'],
            ['V', 'Todos los procesos', 'Planes de mejoramiento', 'Evaluación y seguimiento de los planes de mejoramiento', 'Resultados del seguimiento de los planes de mejoramiento', 'Todos los procesos'],
            ['V', 'Evaluación Independiente', 'Plan Anual de Auditorías', 'Seguimiento y evaluación de las actividades del Plan Anual de Auditorías', 'Informe de seguimiento del Plan Anual de Auditorías', 'Evaluación Independiente'],
            ['A', 'Evaluación Independiente', 'Informes de gestión; hojas de vida de indicadores', 'Gestionar mejoras al proceso', 'Planes de mejoramiento', 'Todos los procesos; partes interesadas'],
            ['A', 'Todos los procesos', 'Registros del proceso', 'Medir indicadores', 'Hojas de vida de indicadores; informes de gestión', 'Todos los procesos'],
        ],
        indicadores: ['Eficacia en acciones de mejoramiento', 'Cumplimiento del Plan Anual de Auditorías, Seguimientos e Informes de Control Interno'],
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
        acuerdos: [
            ['Impulso procesal', 'Impulso procesal de los expedientes en un porcentaje mínimo del 85%'],
            ['Capacitación', 'Capacitación al 90% de los servidores públicos de la entidad'],
            ['Valores', 'Apropiación de valores y principios normativos'],
        ],
        phva: [
            ['P', 'Oficina de Control Interno Disciplinario; servidores públicos; entes de control', 'Queja, solicitud o informe respecto a una conducta de un servidor público disciplinable', 'Formular el plan de acción; recibir y radicar la queja; identificar de oficio conductas relevantes; valorar y repartir según su contenido', 'Plan de acción aprobado; traslado a otra dependencia; archivo; reparto interno; auto inhibitorio', 'Todos los procesos del sistema; entidades externas; servidores públicos'],
            ['H', 'Proceso Control Disciplinario; usuarios internos y externos; ramas del poder público; entidades gubernamentales', 'Acciones del proceso; plan de capacitación en normativa disciplinaria; informes y quejas; políticas de operación y conceptos jurídicos', 'Ejecutar las acciones del proceso en sede de instrucción: evaluación de quejas, auto inhibitorio, indagación previa, investigación, pliego de cargos y remisión a juzgamiento, terminación o archivo, caducidad y prescripción; ejecutar el plan de capacitación', 'Impulsos procesales; expedientes tramitados; servidores capacitados', 'Usuario interno o externo; Proceso de Instrucción de Control Interno Disciplinario'],
            ['V', 'Proceso Control Disciplinario; Atención a la Ciudadanía; Evaluación Independiente; Procuraduría', 'Resultados de gestión; PQRSD; informes de auditoría', 'Gestionar PQRSD; seguimiento a partes interesadas; seguimiento y medición; seguimiento a riesgos; análisis de hallazgos; verificación de eficacia de acciones', 'PQRSD gestionadas; partes interesadas actualizadas; indicadores analizados; riesgos gestionados; hallazgos analizados; eficacia verificada', 'Instrucción de Control Interno Disciplinario; Mejoramiento Continuo; Atención a la Ciudadanía'],
            ['A', 'Mejoramiento Continuo', 'Indicadores y riesgos analizados; PQRSD gestionadas; hallazgos de auditoría; procedimiento de acciones correctivas', 'Tomar acciones para la mejora (con base en indicadores, riesgos y PQRSD)', 'Planes de mejoramiento implementados', 'Instrucción de Control Interno Disciplinario; Evaluación Independiente; Mejoramiento Continuo; Procuraduría'],
        ],
        indicadores: ['Impulso procesal de expedientes', 'Capacitación de servidores públicos', 'Recepción y radicación de quejas e informes'],
        riesgos: ['Vencimiento de términos que genere prescripción', 'Pérdida de documentación', 'Hechos de favorecimiento dentro de los procesos'],
        recursos: { humanos: 'Jefe de Oficina de Control Interno Disciplinario', fisicos: 'Espacios de trabajo y equipos de oficina; instalaciones, hardware, software e internet' },
        documentos: 'Procedimiento de Instrucción de Control Interno Disciplinario y sus formatos; Constitución Política; Código Disciplinario Único; CPACA; códigos Penal, General del Proceso, Sustantivo y Procesal del Trabajo, y de Comercio',
        requisitosISO: '4 Contexto · 5.3 Roles y responsabilidades · 6 Planificación · 7.1.2 Personas · 7.2 Competencias · 7.3 Toma de conciencia · 7.4 Comunicación · 7.5 Información documentada · 9 Evaluación del desempeño · 10 Mejora',
    },
};
/* ===== 3. Componentes ===== */
const Codigo = ({ children }) => children
    ? React.createElement("span", { className: "f-mono text-xs font-bold px-1.5 py-0.5 rounded bg-[#14231B] text-[#B5E048]" }, children)
    : React.createElement("span", { className: "f-mono text-xs px-1.5 py-0.5 rounded bg-[#DCE5DC] text-[#5b6b5f]" }, "sin c\u00F3digo");
const FilaDocumento = ({ doc }) => {
    var _a;
    return (React.createElement("div", { className: "tarjeta flex items-center gap-3 bg-white rounded-xl border border-[#DCE5DC] px-4 py-3" },
        React.createElement(Codigo, null, doc.codigo),
        React.createElement("div", { className: "flex-1 min-w-0" },
            React.createElement("p", { className: "font-medium leading-snug" }, doc.nombre),
            React.createElement("p", { className: "text-xs text-[#5b6b5f]" },
                tipoDeCodigo(doc.codigo),
                " \u00B7 ", (_a = PROCESOS.find((p) => p.sigla === doc.proceso)) === null || _a === void 0 ? void 0 :
                _a.nombre)),
        doc.estado === 'aprobacion'
            ? React.createElement("span", { className: "text-xs font-semibold text-[#8A5A2C] bg-[#F4A93C]/20 border border-[#F4A93C]/50 rounded-full px-3 py-1 whitespace-nowrap" }, "En aprobaci\u00F3n")
            : React.createElement("a", { href: enlaceDoc(doc.archivo), target: "_blank", rel: "noopener", className: "text-sm font-semibold text-[#1E6B47] hover:text-[#144D33] whitespace-nowrap" }, "Abrir \u2197")));
};
const Buscador = ({ irA }) => {
    const [q, setQ] = useState('');
    const resultados = useMemo(() => {
        const t = q.trim().toLowerCase();
        if (t.length < 2)
            return [];
        return DOCUMENTOS.filter((d) => d.codigo.toLowerCase().includes(t) || d.nombre.toLowerCase().includes(t)).slice(0, 8);
    }, [q]);
    return (React.createElement("div", { className: "relative" },
        React.createElement("input", { value: q, onChange: (e) => setQ(e.target.value), placeholder: "Busca por c\u00F3digo o nombre: GF-P03, vi\u00E1ticos, PQRS\u2026", "aria-label": "Buscar documento del SGC", className: "w-full rounded-2xl border-2 border-[#14231B] bg-white px-5 py-4 text-base shadow-[4px_4px_0_#14231B] focus:outline-none focus:border-[#1E6B47]" }),
        resultados.length > 0 && (React.createElement("div", { className: "absolute z-20 mt-2 w-full bg-white border border-[#DCE5DC] rounded-2xl shadow-xl overflow-hidden" }, resultados.map((d) => (React.createElement("button", { key: d.codigo + d.nombre, onClick: () => { setQ(''); irA(`proceso/${d.proceso}`); }, className: "w-full text-left px-4 py-3 hover:bg-[#F7F8F4] flex items-center gap-3 border-b border-[#F0F3EE] last:border-0" },
            React.createElement(Codigo, null, d.codigo),
            React.createElement("span", { className: "text-sm" }, d.nombre))))))));
};
// El mapa de procesos como diagrama clásico y clicable: franjas horizontales,
// flechas laterales de entrada (necesidades) y salida (satisfacción), y la
// banda transversal del MIPG. Cada nodo navega a la vista del proceso.
const NodoProceso = ({ p, irA }) => (React.createElement("button", { onClick: () => irA(`proceso/${p.sigla}`), "aria-label": `Entrar al proceso ${p.nombre}`, className: "nodo-proceso" },
    React.createElement("span", { className: "flex-1 text-left leading-tight" }, p.nombre),
    React.createElement("span", { className: "f-mono text-[10px] font-bold opacity-70" }, p.sigla)));
const FranjaMapa = ({ franja, clase, irA }) => (React.createElement("div", { className: `franja-mapa ${clase}` },
    React.createElement("p", { className: "franja-titulo" }, franja.nombre),
    React.createElement("div", { className: "franja-nodos" }, franja.procesos.map((p) => React.createElement(NodoProceso, { key: p.sigla, p: p, irA: irA })))));
const MapaProcesos = ({ irA }) => (React.createElement("nav", { className: "mapa-marco", "aria-label": "Mapa de procesos del SGC" },
    React.createElement(FranjaMapa, { franja: FRANJAS[0], clase: "f-estrategicos", irA: irA }),
    React.createElement("div", { className: "mapa-centro" },
        React.createElement("div", { className: "flecha-lateral flecha-entrada", "aria-hidden": "true" },
            React.createElement("span", null, "Necesidades y expectativas")),
        React.createElement("div", { className: "flex-1 min-w-0 space-y-2" },
            React.createElement(FranjaMapa, { franja: FRANJAS[1], clase: "f-misionales", irA: irA }),
            React.createElement(FranjaMapa, { franja: FRANJAS[2], clase: "f-apoyo", irA: irA }),
            React.createElement(FranjaMapa, { franja: FRANJAS[3], clase: "f-evaluacion", irA: irA })),
        React.createElement("div", { className: "flecha-lateral flecha-salida", "aria-hidden": "true" },
            React.createElement("span", null, "Satisfacci\u00F3n de los grupos de valor"))),
    React.createElement("button", { onClick: () => irA('proceso/MIPG'), className: "banda-mipg", "aria-label": "Entrar al Modelo Integrado de Planeaci\u00F3n y Gesti\u00F3n (MIPG)" },
        "Modelo Integrado de Planeaci\u00F3n y Gesti\u00F3n \u2014 MIPG",
        React.createElement("span", { className: "font-normal opacity-80" }, " \u00B7 transversal a todos los procesos")),
    React.createElement("p", { className: "text-xs text-[#5b6b5f] text-center mt-2 no-print" }, "Haz clic en un proceso para ver sus procedimientos, formatos y manuales.")));
const Flujograma = ({ pasos }) => (React.createElement("div", { className: "sendero" }, pasos.map((paso, i) => {
    const rol = ROLES[paso.rol];
    const clase = paso.tipo === 'decision' ? 'decision'
        : paso.tipo === 'inicio' ? 'extremo' : paso.tipo === 'fin' ? 'extremo fin' : '';
    return (React.createElement("div", { key: i, className: `nodo ${clase} pb-6`, style: { '--i': i } },
        React.createElement("div", { className: "marca text-sm", style: { color: paso.tipo === 'inicio' || paso.tipo === 'fin' ? undefined : rol.color } },
            React.createElement("span", null, paso.tipo === 'decision' ? '?' : paso.tipo === 'fin' ? '✓' : i + 1)),
        React.createElement("div", { className: "bg-white rounded-2xl border border-[#DCE5DC] p-4 shadow-sm" },
            React.createElement("div", { className: "flex flex-wrap items-center gap-2 mb-1" },
                React.createElement("h4", { className: "f-display font-semibold text-base leading-snug flex-1" }, paso.titulo),
                paso.plazo && (React.createElement("span", { className: "text-xs font-bold text-[#14231B] bg-[#F4A93C] rounded-full px-2.5 py-0.5" },
                    "\u23F1 ",
                    paso.plazo))),
            React.createElement("p", { className: "text-sm text-[#3c4a40] leading-relaxed" }, paso.detalle),
            paso.enlace && (React.createElement("a", { href: paso.enlace.url, target: "_blank", rel: "noopener", className: "mt-2 inline-flex items-center gap-1.5 text-sm font-semibold text-[#1E6B47] bg-[#B5E048]/25 border border-[#1E6B47]/40 rounded-lg px-3 py-1.5 hover:bg-[#B5E048]/45" },
                "\uD83D\uDD17 ",
                paso.enlace.texto,
                " \u2197")),
            React.createElement("div", { className: "mt-2 flex flex-wrap items-center gap-2" },
                React.createElement("span", { className: "inline-flex items-center gap-1.5 text-xs font-semibold rounded-full px-2.5 py-1 text-white", style: { background: rol.color } },
                    React.createElement("span", { className: "w-1.5 h-1.5 rounded-full bg-white/80" }),
                    rol.nombre)),
            paso.tipo === 'decision' && (React.createElement("div", { className: "mt-3 grid grid-cols-1 sm:grid-cols-2 gap-2" },
                React.createElement("div", { className: "rounded-xl border-2 border-[#1E6B47] bg-[#1E6B47]/5 px-3 py-2 text-sm" },
                    React.createElement("span", { className: "font-bold text-[#1E6B47]" }, "S\u00ED \u2192"),
                    " ",
                    paso.si),
                React.createElement("div", { className: "rounded-xl border-2 border-[#F4A93C] bg-[#F4A93C]/10 px-3 py-2 text-sm" },
                    React.createElement("span", { className: "font-bold text-[#8A5A2C]" }, "No \u2192"),
                    " ",
                    paso.no))))));
})));
// Instructivo paso a paso con casos (A/B), requisitos previos, avisos y
// comparación final. Se usa en las guías que no son un flujograma lineal.
const Instructivo = ({ data }) => (React.createElement("div", { className: "space-y-5" },
    data.enlaceApp && (React.createElement("a", { href: data.enlaceApp, target: "_blank", rel: "noopener", className: "inline-flex items-center gap-2 rounded-xl bg-[#1E6B47] text-white font-semibold px-4 py-2.5 hover:bg-[#144D33] shadow-[3px_3px_0_#14231B]" }, "\uD83D\uDD17 Abrir el aplicativo de \u00D3rdenes de Pedido \u2197")),
    data.enlace && (React.createElement("a", { href: data.enlace.url, target: "_blank", rel: "noopener", className: "inline-flex items-center gap-2 rounded-xl bg-[#1E6B47] text-white font-semibold px-4 py-2.5 hover:bg-[#144D33] shadow-[3px_3px_0_#14231B]" },
        "\uD83D\uDD17 ",
        data.enlace.texto,
        " \u2197")),
    (data.casos || []).map((c, i) => (React.createElement("div", { key: i, className: "bg-white rounded-2xl border border-[#DCE5DC] p-4 sm:p-5 shadow-sm" },
        React.createElement("h3", { className: "f-display text-lg font-bold text-[#14231B] mb-3" }, c.titulo),
        c.intro && React.createElement("p", { className: "text-sm text-[#3c4a40] leading-relaxed mb-2" }, c.intro),
        c.antes && (React.createElement("div", { className: "mb-3 rounded-xl bg-[#DCE5DC]/40 border border-[#DCE5DC] px-3 py-2" },
            React.createElement("p", { className: "text-sm font-semibold text-[#14231B] mb-0.5" }, "Antes de empezar"),
            React.createElement("p", { className: "text-sm text-[#3c4a40] leading-relaxed" }, c.antes),
            c.antesSub && (React.createElement("ul", { className: "mt-1.5 list-disc pl-5 text-sm text-[#3c4a40] space-y-1" }, c.antesSub.map((s, j) => React.createElement("li", { key: j }, s)))))),
        c.pasos && (React.createElement("ol", { className: "list-decimal pl-5 space-y-2 text-sm text-[#3c4a40] marker:font-bold marker:text-[#1E6B47]" }, c.pasos.map((p, j) => {
            const paso = typeof p === 'string' ? { texto: p } : p;
            return (React.createElement("li", { key: j, className: "leading-relaxed" },
                paso.texto,
                paso.sub && (React.createElement("ul", { className: "mt-1 list-disc pl-5 space-y-0.5 text-[#3c4a40]" }, paso.sub.map((s, k) => React.createElement("li", { key: k }, s)))),
                paso.aviso && (React.createElement("p", { className: "mt-1.5 rounded-lg bg-[#F4A93C]/15 border border-[#F4A93C]/50 px-2.5 py-1.5 text-[#8A5A2C]" },
                    "\u26A0\uFE0F ",
                    paso.aviso))));
        }))),
        c.tabla && (React.createElement("div", { className: "mt-2 overflow-x-auto" },
            React.createElement("table", { className: "w-full text-sm border-collapse" },
                React.createElement("thead", null,
                    React.createElement("tr", null, c.tabla.columnas.map((col, k) => (React.createElement("th", { key: k, className: "text-left font-semibold text-[#14231B] bg-[#DCE5DC]/50 border border-[#DCE5DC] px-2.5 py-1.5" }, col))))),
                React.createElement("tbody", null, c.tabla.filas.map((fila, fi) => (React.createElement("tr", { key: fi }, fila.map((cel, ci) => (React.createElement("td", { key: ci, className: `align-top border border-[#DCE5DC] px-2.5 py-1.5 ${ci === 0 ? 'font-medium text-[#14231B]' : 'text-[#3c4a40]'}` }, cel)))))))))),
        c.pie && React.createElement("p", { className: "mt-1.5 text-xs text-[#5b6b5f]" }, c.pie),
        c.aviso && (React.createElement("p", { className: "mt-2 rounded-lg bg-[#F4A93C]/15 border border-[#F4A93C]/50 px-2.5 py-1.5 text-sm text-[#8A5A2C]" },
            "\u26A0\uFE0F ",
            c.aviso)),
        c.flujo && (React.createElement("p", { className: "mt-3 rounded-xl bg-[#1E6B47]/10 border border-[#1E6B47]/25 px-3 py-2 text-sm text-[#14231B] leading-relaxed" },
            "\u27A1\uFE0F ",
            c.flujo)),
        c.nota && (React.createElement("p", { className: "mt-2 rounded-xl bg-[#B5E048]/25 border border-[#B5E048]/60 px-3 py-2 text-sm text-[#14231B] leading-relaxed" },
            "\uD83D\uDCA1 ",
            c.nota))))),
    data.resumenFinal && (React.createElement("div", { className: "rounded-2xl bg-[#B5E048]/25 border-2 border-[#1E6B47]/30 p-4" },
        React.createElement("p", { className: "f-mono text-xs font-bold text-[#1E6B47] uppercase tracking-widest mb-1" }, "Resumen r\u00E1pido"),
        React.createElement("p", { className: "text-sm text-[#14231B] leading-relaxed" }, data.resumenFinal))),
    data.comparacion && (React.createElement("div", { className: "bg-[#14231B] text-[#F7F8F4] rounded-2xl p-5" },
        React.createElement("p", { className: "f-display font-semibold mb-2" }, data.comparacion.titulo),
        React.createElement("ul", { className: "space-y-1.5 text-sm text-white/90 list-disc pl-5" }, data.comparacion.items.map((it, i) => React.createElement("li", { key: i }, it)))))));
const VistaGuia = ({ guia, irA }) => {
    const docs = (guia.docs || []).map((c) => DOCUMENTOS.find((d) => d.codigo === c)).filter(Boolean);
    const formatos = guia.formatos || [];
    const proceso = PROCESOS.find((p) => p.sigla === guia.proceso);
    return (React.createElement("div", { className: "max-w-3xl mx-auto" },
        React.createElement("button", { onClick: () => irA(''), className: "no-print text-sm font-semibold text-[#1E6B47] mb-4" }, "\u2190 Volver al inicio"),
        React.createElement("p", { className: "f-mono text-xs font-bold text-[#1E6B47] uppercase tracking-widest mb-1" },
            "Gu\u00EDa",
            guia.fuente ? ` · fuente ${guia.fuente}` : ''),
        React.createElement("h2", { className: "f-display text-3xl sm:text-4xl font-extrabold leading-tight mb-2" }, guia.pregunta),
        React.createElement("p", { className: "text-[#3c4a40] mb-6 leading-relaxed" }, guia.resumen),
        guia.instructivo ? React.createElement(Instructivo, { data: guia.instructivo }) : React.createElement(Flujograma, { pasos: guia.pasos }),
        (docs.length > 0 || formatos.length > 0) && (React.createElement("div", { className: "mt-4 bg-[#14231B] text-[#F7F8F4] rounded-2xl p-5" },
            React.createElement("p", { className: "f-display font-semibold mb-3" }, "Documentos oficiales de esta gu\u00EDa"),
            React.createElement("div", { className: "space-y-2" },
                docs.map((d) => (React.createElement("div", { key: d.codigo, className: "flex items-center gap-3 text-sm" },
                    React.createElement("span", { className: "f-mono text-xs font-bold text-[#B5E048]" }, d.codigo),
                    React.createElement("span", { className: "flex-1 text-white/90" }, d.nombre),
                    d.estado === 'aprobacion'
                        ? React.createElement("span", { className: "text-xs text-[#F4A93C]" }, "En aprobaci\u00F3n")
                        : React.createElement("a", { className: "font-semibold text-[#B5E048] hover:underline", href: enlaceDoc(d.archivo), target: "_blank", rel: "noopener" }, "Abrir \u2197")))),
                formatos.map((f) => (React.createElement("div", { key: f.codigo || f.nombre, className: "flex items-center gap-3 text-sm" },
                    React.createElement("span", { className: "f-mono text-xs font-bold text-[#B5E048]" }, f.codigo || 'Formato'),
                    React.createElement("span", { className: "flex-1 text-white/90" }, f.nombre),
                    React.createElement("a", { className: "font-semibold text-[#B5E048] hover:underline", href: f.url, target: "_blank", rel: "noopener" }, "Abrir \u2197"))))),
            React.createElement("p", { className: "text-xs text-white/60 mt-3" },
                "Esta gu\u00EDa es un resumen orientativo; ante cualquier diferencia, manda el procedimiento oficial. Proceso due\u00F1o: ", proceso === null || proceso === void 0 ? void 0 :
                proceso.nombre,
                ".")))));
};
// Tarjetas de sección: una por subcarpeta real del proceso en SharePoint.
// Si el proceso no tiene subcarpetas (Planeación Estratégica), se ofrece la
// carpeta raíz como única sección.
const SeccionesProceso = ({ proceso }) => {
    const secciones = proceso.secciones.length
        ? proceso.secciones
        : [{ nombre: 'Documentos del proceso', carpeta: '' }];
    return (React.createElement("div", { className: "mt-5" },
        React.createElement("div", { className: "flex items-baseline justify-between gap-3 mb-2" },
            React.createElement("h3", { className: "f-display text-lg font-semibold" }, "Carpetas del proceso en el repositorio"),
            React.createElement("a", { href: enlaceCarpeta(proceso.carpeta), target: "_blank", rel: "noopener", className: "text-sm font-semibold text-[#1E6B47] hover:underline whitespace-nowrap" }, "Abrir carpeta \u2197")),
        React.createElement("div", { className: "grid grid-cols-2 sm:grid-cols-3 gap-3" }, secciones.map((s) => (React.createElement("a", { key: s.nombre, target: "_blank", rel: "noopener", href: enlaceCarpeta(s.base || proceso.carpeta, s.carpeta || undefined), className: "tarjeta bg-white rounded-2xl border-2 border-[#DCE5DC] hover:border-[#1E6B47] p-4 flex flex-col items-start gap-1" },
            React.createElement("span", { className: "text-2xl", "aria-hidden": "true" }, ICONO_SECCION[s.nombre] || '📁'),
            React.createElement("span", { className: "font-semibold leading-snug" }, s.nombre),
            React.createElement("span", { className: "text-xs text-[#5b6b5f]" }, "Abrir en SharePoint \u2197")))))));
};
// Exporta la caracterización a Excel (xlsx-js-style por CDN); si la librería
// no cargó, cae a CSV para no dejar al usuario sin descarga.
const exportarCaracterizacion = (proceso, c) => {
    const filas = [
        ['CARACTERIZACIÓN DE PROCESO — ' + proceso.nombre.toUpperCase()],
        ['Código', c.codigo, '', 'Versión', c.version],
        ['Proceso/Subproceso', c.proceso],
        ['Cargo líder del proceso', c.lider],
        ['Estado', c.estado === 'oficial' ? 'Oficial (transcrita del repositorio)' : 'PROPUESTA — pendiente de validación'],
        [],
        ['OBJETIVO', c.objetivo],
        ['ALCANCE', c.alcance],
        ['RESULTADO RELEVANTE', c.resultado],
        [],
        ['ACUERDOS DE SERVICIO (Atributo)', 'CRITERIO DE ACEPTABILIDAD'],
        ...c.acuerdos.map((a) => [a[0], a[1]]),
        [],
        ['PHVA', 'Proveedor (¿quién suministra?)', 'Entrada / recurso', 'Actividades', 'Producto / salida', 'Receptor (¿quién recibe?)'],
        ...c.phva.map((f) => [...f]),
        [],
        ['INDICADORES', c.indicadores.join(' · ')],
        ['RIESGOS', c.riesgos.join(' · ')],
        ['RECURSOS HUMANOS', c.recursos.humanos],
        ['RECURSOS FÍSICOS', c.recursos.fisicos],
        ['DOCUMENTOS ASOCIADOS', c.documentos],
        ['REQUISITOS ISO 9001:2015', c.requisitosISO],
    ];
    const nombre = `${c.codigo} Caracterización ${proceso.nombre}`;
    if (typeof XLSX !== 'undefined') {
        const ws = XLSX.utils.aoa_to_sheet(filas);
        ws['!cols'] = [{ wch: 10 }, { wch: 40 }, { wch: 45 }, { wch: 55 }, { wch: 45 }, { wch: 40 }];
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, c.codigo);
        XLSX.writeFile(wb, nombre + '.xlsx');
    }
    else {
        const csv = filas.map((f) => f.map((x) => `"${String(x == null ? '' : x).replace(/"/g, '""')}"`).join(';')).join('\n');
        const enlace = document.createElement('a');
        enlace.href = URL.createObjectURL(new Blob(['﻿' + csv], { type: 'text/csv;charset=utf-8' }));
        enlace.download = nombre + '.csv';
        enlace.click();
    }
};
// Primera página de cada proceso: su caracterización (oficial o propuesta).
const VistaCaracterizacion = ({ proceso, c }) => (React.createElement("div", null,
    c.estado === 'propuesta' && (React.createElement("div", { className: "mb-4 rounded-xl bg-[#F4A93C]/15 border-2 border-[#F4A93C] px-4 py-3" },
        React.createElement("p", { className: "font-bold text-[#8A5A2C]" }, "\u26A0\uFE0F Propuesta de caracterizaci\u00F3n"),
        React.createElement("p", { className: "text-sm text-[#8A5A2C]" }, "Este proceso a\u00FAn no tiene caracterizaci\u00F3n en el repositorio. Este es un borrador para validaci\u00F3n del l\u00EDder del proceso y la Oficina de Calidad."))),
    React.createElement("div", { className: "flex flex-wrap items-center gap-2 mb-4" },
        React.createElement("span", { className: "f-mono text-xs font-bold px-2 py-1 rounded bg-[#14231B] text-[#B5E048]" }, c.codigo),
        React.createElement("span", { className: "text-xs font-semibold text-[#5b6b5f]" },
            "Versi\u00F3n: ",
            c.version),
        React.createElement("span", { className: "text-xs text-[#5b6b5f]" },
            "\u00B7 L\u00EDder: ",
            c.lider),
        React.createElement("span", { className: "flex-1" }),
        React.createElement("button", { onClick: () => exportarCaracterizacion(proceso, c), className: "no-print text-sm font-bold bg-[#1E6B47] text-white rounded-xl px-4 py-2 hover:bg-[#144D33] shadow-[3px_3px_0_#14231B]" }, "\u2B07\uFE0F Descargar Excel"),
        c.archivo && (React.createElement("a", { href: enlaceCarpeta(c.archivo[0]) + '/' + encodeURIComponent(c.archivo[1]), target: "_blank", rel: "noopener", className: "no-print text-sm font-semibold text-[#1E6B47] hover:underline" }, "Original \u2197"))),
    React.createElement("div", { className: "space-y-3" },
        React.createElement("div", { className: "bg-white rounded-2xl border border-[#DCE5DC] p-4" },
            React.createElement("p", { className: "f-mono text-[10px] font-bold text-[#1E6B47] uppercase tracking-widest mb-1" }, "Objetivo"),
            React.createElement("p", { className: "text-sm leading-relaxed" }, c.objetivo),
            React.createElement("p", { className: "f-mono text-[10px] font-bold text-[#1E6B47] uppercase tracking-widest mt-3 mb-1" }, "Alcance"),
            React.createElement("p", { className: "text-sm leading-relaxed" }, c.alcance),
            React.createElement("p", { className: "f-mono text-[10px] font-bold text-[#1E6B47] uppercase tracking-widest mt-3 mb-1" }, "Resultado relevante"),
            React.createElement("p", { className: "text-sm leading-relaxed" }, c.resultado)),
        React.createElement("div", { className: "bg-white rounded-2xl border border-[#DCE5DC] p-4" },
            React.createElement("p", { className: "f-mono text-[10px] font-bold text-[#1E6B47] uppercase tracking-widest mb-2" }, "Acuerdos de servicio"),
            React.createElement("div", { className: "space-y-1.5" }, c.acuerdos.map((a, i) => (React.createElement("div", { key: i, className: "flex flex-wrap gap-x-3 text-sm" },
                React.createElement("span", { className: "font-semibold min-w-[11rem]" }, a[0]),
                React.createElement("span", { className: "flex-1 text-[#3c4a40]" }, a[1])))))),
        React.createElement("div", { className: "bg-white rounded-2xl border border-[#DCE5DC] p-4" },
            React.createElement("p", { className: "f-mono text-[10px] font-bold text-[#1E6B47] uppercase tracking-widest mb-2" }, "Ciclo PHVA"),
            React.createElement("div", { className: "overflow-x-auto" },
                React.createElement("table", { className: "w-full text-xs border-collapse min-w-[52rem]" },
                    React.createElement("thead", null,
                        React.createElement("tr", { className: "bg-[#DCE5DC]/50 text-left" },
                            React.createElement("th", { className: "px-2 py-1.5 font-bold w-10" }, "PHVA"),
                            React.createElement("th", { className: "px-2 py-1.5 font-semibold" }, "Proveedor"),
                            React.createElement("th", { className: "px-2 py-1.5 font-semibold" }, "Entrada / recurso"),
                            React.createElement("th", { className: "px-2 py-1.5 font-semibold" }, "Actividades"),
                            React.createElement("th", { className: "px-2 py-1.5 font-semibold" }, "Producto / salida"),
                            React.createElement("th", { className: "px-2 py-1.5 font-semibold" }, "Receptor"))),
                    React.createElement("tbody", null, c.phva.map((f, i) => (React.createElement("tr", { key: i, className: "border-t border-[#F0F3EE] align-top" },
                        React.createElement("td", { className: "px-2 py-1.5" },
                            React.createElement("span", { className: "f-mono font-bold text-white bg-[#1E6B47] rounded px-1.5" }, f[0])),
                        React.createElement("td", { className: "px-2 py-1.5 text-[#3c4a40]" }, f[1]),
                        React.createElement("td", { className: "px-2 py-1.5 text-[#3c4a40]" }, f[2]),
                        React.createElement("td", { className: "px-2 py-1.5 font-medium" }, f[3]),
                        React.createElement("td", { className: "px-2 py-1.5 text-[#3c4a40]" }, f[4]),
                        React.createElement("td", { className: "px-2 py-1.5 text-[#3c4a40]" }, f[5])))))))),
        React.createElement("div", { className: "grid grid-cols-1 sm:grid-cols-2 gap-3" },
            React.createElement("div", { className: "bg-white rounded-2xl border border-[#DCE5DC] p-4" },
                React.createElement("p", { className: "f-mono text-[10px] font-bold text-[#1E6B47] uppercase tracking-widest mb-1.5" }, "Indicadores"),
                React.createElement("ul", { className: "text-sm space-y-1 list-disc pl-4 text-[#3c4a40]" }, c.indicadores.map((x, i) => React.createElement("li", { key: i }, x)))),
            React.createElement("div", { className: "bg-white rounded-2xl border border-[#DCE5DC] p-4" },
                React.createElement("p", { className: "f-mono text-[10px] font-bold text-[#B5432E] uppercase tracking-widest mb-1.5" }, "Riesgos"),
                React.createElement("ul", { className: "text-sm space-y-1 list-disc pl-4 text-[#3c4a40]" }, c.riesgos.map((x, i) => React.createElement("li", { key: i }, x))))),
        React.createElement("div", { className: "bg-white rounded-2xl border border-[#DCE5DC] p-4 text-sm space-y-2" },
            React.createElement("p", null,
                React.createElement("span", { className: "font-semibold" }, "Recursos humanos:"),
                " ",
                React.createElement("span", { className: "text-[#3c4a40]" }, c.recursos.humanos)),
            React.createElement("p", null,
                React.createElement("span", { className: "font-semibold" }, "Recursos f\u00EDsicos:"),
                " ",
                React.createElement("span", { className: "text-[#3c4a40]" }, c.recursos.fisicos)),
            React.createElement("p", null,
                React.createElement("span", { className: "font-semibold" }, "Documentos asociados:"),
                " ",
                React.createElement("span", { className: "text-[#3c4a40]" }, c.documentos)),
            React.createElement("p", null,
                React.createElement("span", { className: "font-semibold" }, "Requisitos ISO 9001:2015:"),
                " ",
                React.createElement("span", { className: "text-[#3c4a40]" }, c.requisitosISO))))));
const VistaProceso = ({ sigla, irA }) => {
    const proceso = PROCESOS.find((p) => p.sigla === sigla);
    const docs = DOCUMENTOS.filter((d) => d.proceso === sigla);
    const guias = GUIAS.filter((g) => g.proceso === sigla);
    const caract = CARACTERIZACIONES[sigla];
    // La caracterización es la primera página del proceso; los documentos, la segunda.
    const [pestana, setPestana] = useState('caracterizacion');
    if (!proceso)
        return React.createElement("p", { className: "text-center py-10" },
            "Proceso no encontrado. ",
            React.createElement("button", { className: "text-[#1E6B47] font-semibold", onClick: () => irA('') }, "Volver"));
    const vistaDocumentos = (React.createElement("div", null,
        proceso.enConstruccion ? (React.createElement("div", { className: "text-center py-8" },
            React.createElement("p", { className: "text-5xl mb-3", "aria-hidden": "true" }, "\uD83D\uDEA7"),
            React.createElement("p", { className: "text-lg font-semibold text-[#8A5A2C]" }, "En construcci\u00F3n"),
            React.createElement("p", { className: "text-sm text-[#5b6b5f] mt-1 max-w-md mx-auto" }, "Este proceso est\u00E1 en estructuraci\u00F3n: pronto tendr\u00E1 su carpeta en el repositorio y sus documentos oficiales."))) : (React.createElement(SeccionesProceso, { proceso: proceso })),
        guias.length > 0 && (React.createElement("div", { className: "mt-5 flex flex-wrap gap-2" }, guias.map((g) => (React.createElement("button", { key: g.id, onClick: () => irA(`guia/${g.id}`), className: "text-sm font-semibold bg-[#B5E048] text-[#14231B] rounded-full px-4 py-2 hover:bg-[#a3d13a]" }, g.pregunta))))),
        docs.length > 0 && (React.createElement("div", { className: "mt-6" },
            React.createElement("h3", { className: "f-display text-lg font-semibold mb-2" }, "Documentos del inventario (Word)"),
            React.createElement("div", { className: "space-y-2" }, docs.map((d) => React.createElement(FilaDocumento, { key: d.codigo + d.nombre, doc: d })))))));
    return (React.createElement("div", { className: "max-w-4xl mx-auto" },
        React.createElement("button", { onClick: () => irA(''), className: "no-print text-sm font-semibold text-[#1E6B47] mb-4" }, "\u2190 Volver al inicio"),
        React.createElement("div", { className: "mb-3" },
            React.createElement("p", { className: "f-mono text-xs font-bold text-[#1E6B47] uppercase tracking-widest" },
                proceso.franja,
                " \u00B7 ",
                proceso.sigla),
            React.createElement("h2", { className: "f-display text-3xl font-extrabold leading-tight" }, proceso.nombre)),
        caract ? (React.createElement("div", null,
            React.createElement("div", { className: "no-print flex gap-2 mb-4 border-b-2 border-[#DCE5DC]" }, [['caracterizacion', 'Caracterización'], ['documentos', 'Documentos y carpetas']].map(([id, nombre]) => (React.createElement("button", { key: id, onClick: () => setPestana(id), className: `text-sm font-bold px-4 py-2 rounded-t-xl ${pestana === id ? 'bg-[#1E6B47] text-white' : 'text-[#5b6b5f] hover:text-[#1E6B47]'}` }, nombre)))),
            pestana === 'caracterizacion' ? React.createElement(VistaCaracterizacion, { proceso: proceso, c: caract }) : vistaDocumentos)) : vistaDocumentos));
};
const ZonaBadge = ({ zona }) => {
    const z = ZONAS_RIESGO[zona] || ZONAS_RIESGO['N/D'];
    return (React.createElement("span", { className: "inline-block text-[10px] font-bold rounded-full px-2 py-0.5 whitespace-nowrap", style: { background: z.color, color: z.texto } }, zona));
};
const VistaRiesgos = ({ irA }) => {
    // Conteo por zona inherente para las fichas del resumen
    const porZona = Object.keys(ZONAS_RIESGO).filter((z) => z !== 'N/D')
        .map((z) => ({ zona: z, n: TODOS_RIESGOS.filter((r) => r.zi === z).length }));
    return (React.createElement("div", { className: "max-w-4xl mx-auto" },
        React.createElement("button", { onClick: () => irA(''), className: "no-print text-sm font-semibold text-[#1E6B47] mb-4" }, "\u2190 Volver al inicio"),
        React.createElement("p", { className: "f-mono text-xs font-bold text-[#1E6B47] uppercase tracking-widest mb-1" }, "Gesti\u00F3n del riesgo \u00B7 Vigencia 2025"),
        React.createElement("h2", { className: "f-display text-3xl sm:text-4xl font-extrabold leading-tight mb-2" }, "Mapa de riesgos"),
        React.createElement("p", { className: "text-[#3c4a40] leading-relaxed max-w-2xl" },
            "Seguimiento a la gesti\u00F3n del riesgo con la matriz actualizada al 25/02/2026 (Informe CI 2022-2025, Rad. 202503000143), en cumplimiento de la pol\u00EDtica OE-M02.",
            ' ',
            TODOS_RIESGOS.length,
            " riesgos en 5 componentes."),
        React.createElement("a", { href: ENLACE_SEGUIMIENTO_RIESGOS, target: "_blank", rel: "noopener", className: "mt-3 inline-flex items-center gap-2 rounded-xl bg-[#1E6B47] text-white font-semibold px-4 py-2.5 hover:bg-[#144D33] shadow-[3px_3px_0_#14231B]" }, "\uD83D\uDCC4 Ver el informe completo de Seguimiento a Riesgos \u2197"),
        React.createElement("div", { className: "mt-5 grid grid-cols-2 sm:grid-cols-5 gap-3" }, porZona.map(({ zona, n }) => {
            const z = ZONAS_RIESGO[zona];
            return (React.createElement("div", { key: zona, className: "rounded-2xl border-2 p-3 text-center", style: { borderColor: z.color } },
                React.createElement("p", { className: "f-display text-3xl font-extrabold", style: { color: z.color } }, n),
                React.createElement("p", { className: "text-[11px] font-bold uppercase tracking-wide", style: { color: z.color } }, zona),
                React.createElement("p", { className: "f-mono text-[10px] text-[#5b6b5f]" },
                    "P\u00D7I ",
                    z.rango)));
        })),
        React.createElement("p", { className: "text-xs text-[#5b6b5f] mt-2" }, "Zona inherente (antes de controles). La escala P\u00D7I es la de la pol\u00EDtica OE-M02: probabilidad 1\u20135 e impacto 1\u201325. Los riesgos en zona Significativo, Alto o Cr\u00EDtico exigen plan de acci\u00F3n en el formato MA-F01."),
        COMPONENTES_RIESGO.map((c) => (React.createElement("section", { key: c.clave, className: "mt-8" },
            React.createElement("div", { className: "flex items-baseline gap-3 mb-2" },
                React.createElement("span", { className: "f-mono text-xs font-bold bg-[#14231B] text-[#B5E048] rounded px-1.5 py-0.5" }, c.clave),
                React.createElement("h3", { className: "f-display text-xl font-bold" }, c.nombre),
                React.createElement("span", { className: "text-xs text-[#5b6b5f]" }, c.fuente)),
            React.createElement("div", { className: "space-y-2" }, c.riesgos.map((r) => (React.createElement("div", { key: r.id, className: "tarjeta bg-white rounded-xl border border-[#DCE5DC] px-4 py-3 flex flex-wrap items-center gap-x-3 gap-y-1.5" },
                React.createElement("span", { className: "f-mono text-xs font-bold text-[#1E6B47] w-9" }, r.id),
                React.createElement("div", { className: "flex-1 min-w-[14rem]" },
                    React.createElement("p", { className: "text-sm font-medium leading-snug" }, r.desc),
                    React.createElement("p", { className: "text-xs text-[#5b6b5f]" },
                        r.proc,
                        " \u00B7 Tratamiento: ",
                        r.trat)),
                React.createElement("span", { className: "flex items-center gap-1.5 text-xs text-[#5b6b5f]" },
                    React.createElement(ZonaBadge, { zona: r.zi }),
                    " \u2192 ",
                    React.createElement(ZonaBadge, { zona: r.zr }))))))))),
        React.createElement("p", { className: "text-xs text-[#5b6b5f] mt-2" }, "Inherente \u2192 Residual (despu\u00E9s de controles). N/D: sin calificaci\u00F3n residual en la matriz."),
        React.createElement("section", { className: "mt-8" },
            React.createElement("h3", { className: "f-display text-xl font-bold mb-1" }, "Indicadores clave de riesgo (KRI) 2026"),
            React.createElement("p", { className: "text-sm text-[#3c4a40] mb-3 max-w-2xl" }, "Alertas tempranas: si un indicador entra en amarillo, el l\u00EDder notifica a Control Interno en m\u00E1ximo 5 d\u00EDas h\u00E1biles; en rojo, escala al Comit\u00E9 Institucional de Gesti\u00F3n y Desempe\u00F1o."),
            React.createElement("div", { className: "overflow-x-auto rounded-2xl border border-[#DCE5DC] bg-white" },
                React.createElement("table", { className: "w-full text-sm border-collapse min-w-[38rem]" },
                    React.createElement("thead", null,
                        React.createElement("tr", { className: "bg-[#DCE5DC]/50 text-left" },
                            React.createElement("th", { className: "px-3 py-2 font-semibold" }, "KRI"),
                            React.createElement("th", { className: "px-3 py-2 font-semibold" }, "Indicador"),
                            React.createElement("th", { className: "px-3 py-2 font-semibold text-[#1E6B47]" }, "Verde"),
                            React.createElement("th", { className: "px-3 py-2 font-semibold text-[#8A5A2C]" }, "Amarillo"),
                            React.createElement("th", { className: "px-3 py-2 font-semibold text-[#B5432E]" }, "Rojo"),
                            React.createElement("th", { className: "px-3 py-2 font-semibold" }, "Frecuencia \u00B7 Responsable"))),
                    React.createElement("tbody", null, KRIS.map((k) => (React.createElement("tr", { key: k[0], className: "border-t border-[#F0F3EE]" },
                        React.createElement("td", { className: "px-3 py-2 f-mono text-xs font-bold text-[#1E6B47]" }, k[0]),
                        React.createElement("td", { className: "px-3 py-2" }, k[1]),
                        React.createElement("td", { className: "px-3 py-2 whitespace-nowrap font-semibold text-[#1E6B47]" }, k[2]),
                        React.createElement("td", { className: "px-3 py-2 whitespace-nowrap text-[#8A5A2C]" }, k[3]),
                        React.createElement("td", { className: "px-3 py-2 whitespace-nowrap font-semibold text-[#B5432E]" }, k[4]),
                        React.createElement("td", { className: "px-3 py-2 text-xs text-[#5b6b5f]" }, k[5])))))))),
        React.createElement("div", { className: "mt-6 bg-[#14231B] text-[#F7F8F4] rounded-2xl p-5" },
            React.createElement("p", { className: "f-display font-semibold mb-2" }, "Documentos de referencia"),
            React.createElement("div", { className: "space-y-2" },
                React.createElement("div", { className: "flex items-center gap-3 text-sm" },
                    React.createElement("span", { className: "f-mono text-xs font-bold text-[#B5E048] whitespace-nowrap" }, "Informe"),
                    React.createElement("span", { className: "flex-1 text-white/90" }, "Seguimiento a la Gesti\u00F3n del Riesgo \u2014 Mapa de Riesgos (vigencia 2025)"),
                    React.createElement("a", { className: "font-semibold text-[#B5E048] hover:underline", href: ENLACE_SEGUIMIENTO_RIESGOS, target: "_blank", rel: "noopener" }, "Abrir \u2197")),
                React.createElement("div", { className: "flex items-center gap-3 text-sm" },
                    React.createElement("span", { className: "f-mono text-xs font-bold text-[#B5E048] whitespace-nowrap" }, "OE-M02"),
                    React.createElement("span", { className: "flex-1 text-white/90" }, "Pol\u00EDtica de administraci\u00F3n y gesti\u00F3n del riesgo"),
                    React.createElement("a", { className: "font-semibold text-[#B5E048] hover:underline", href: enlaceDoc('OE-M02 Politica de administracion gestión del riesgo.docx'), target: "_blank", rel: "noopener" }, "Abrir \u2197"))),
            React.createElement("p", { className: "text-xs text-white/60 mt-3" }, "Resumen orientativo del informe de seguimiento; ante cualquier diferencia, manda el documento oficial aprobado por la Direcci\u00F3n."))));
};
const VistaDocumentos = ({ irA }) => {
    const [filtro, setFiltro] = useState('todos');
    const docs = DOCUMENTOS.filter((d) => filtro === 'todos' || tipoDeCodigo(d.codigo) === filtro)
        .slice().sort((a, b) => (a.codigo || 'ZZ').localeCompare(b.codigo || 'ZZ'));
    return (React.createElement("div", { className: "max-w-3xl mx-auto" },
        React.createElement("button", { onClick: () => irA(''), className: "text-sm font-semibold text-[#1E6B47] mb-4" }, "\u2190 Volver al inicio"),
        React.createElement("h2", { className: "f-display text-3xl font-extrabold mb-4" },
            "Todos los documentos ",
            React.createElement("span", { className: "text-[#5b6b5f] text-xl font-semibold" },
                "(",
                docs.length,
                ")")),
        React.createElement("div", { className: "flex flex-wrap gap-2 mb-4" }, ['todos', ...Object.values(TIPOS)].map((f) => (React.createElement("button", { key: f, onClick: () => setFiltro(f), className: `text-sm font-semibold rounded-full px-4 py-1.5 border-2 ${filtro === f ? 'bg-[#14231B] text-[#B5E048] border-[#14231B]' : 'bg-white border-[#DCE5DC] hover:border-[#1E6B47]'}` }, f === 'todos' ? 'Todos' : f)))),
        React.createElement("div", { className: "space-y-2" }, docs.map((d) => React.createElement(FilaDocumento, { key: d.codigo + d.nombre, doc: d })))));
};
const Inicio = ({ irA }) => (React.createElement("div", { className: "max-w-4xl mx-auto" },
    React.createElement("div", { className: "pt-6 pb-4 text-center" },
        React.createElement("p", { className: "f-mono text-xs font-bold text-[#1E6B47] uppercase tracking-[.2em] mb-1" }, "Sistema de Gesti\u00F3n de Calidad \u00B7 ACTIVA"),
        React.createElement("h1", { className: "f-display text-3xl sm:text-4xl font-extrabold leading-tight" }, "Mapa de procesos")),
    React.createElement(MapaProcesos, { irA: irA }),
    React.createElement("h2", { className: "f-display text-2xl sm:text-3xl font-extrabold leading-tight mt-10 mb-4" },
        "\u00BFC\u00F3mo pido ",
        React.createElement("span", { className: "text-[#1E6B47]" }, "lo que necesito"),
        "?"),
    React.createElement("div", { className: "flex flex-wrap gap-2 mb-6" }, GUIAS.map((g) => (React.createElement("button", { key: g.id, onClick: () => irA(`guia/${g.id}`), className: "tarjeta text-sm font-semibold bg-[#B5E048] text-[#14231B] rounded-full px-4 py-2" }, g.pregunta)))),
    React.createElement(Buscador, { irA: irA }),
    React.createElement("div", { className: "mt-3 flex justify-end gap-5" },
        React.createElement("button", { onClick: () => irA('riesgos'), className: "text-sm font-semibold text-[#1E6B47] hover:underline" }, "Mapa de riesgos \u2192"),
        React.createElement("button", { onClick: () => irA('documentos'), className: "text-sm font-semibold text-[#1E6B47] hover:underline" }, "Ver todos los documentos \u2192"))));
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
        contenido = React.createElement(VistaGuia, { guia: GUIAS.find((g) => g.id === parametro), irA: irA });
    }
    else if (seccion === 'proceso' && parametro) {
        // PE se fusionó con OE; los enlaces viejos siguen funcionando.
        contenido = React.createElement(VistaProceso, { sigla: parametro === 'PE' ? 'OE' : parametro, irA: irA });
    }
    else if (seccion === 'documentos') {
        contenido = React.createElement(VistaDocumentos, { irA: irA });
    }
    else if (seccion === 'riesgos') {
        contenido = React.createElement(VistaRiesgos, { irA: irA });
    }
    else {
        contenido = React.createElement(Inicio, { irA: irA });
    }
    return (React.createElement("div", { className: "min-h-screen flex flex-col" },
        React.createElement("header", { className: "sticky top-0 z-30 bg-[#F7F8F4]/90 backdrop-blur border-b border-[#DCE5DC]" },
            React.createElement("div", { className: "max-w-4xl mx-auto px-4 py-3 flex items-center gap-3" },
                React.createElement("button", { onClick: () => irA(''), className: "f-display font-extrabold text-lg tracking-tight flex items-center gap-2" },
                    React.createElement("span", { className: "w-7 h-7 rounded-lg bg-[#1E6B47] text-[#B5E048] flex items-center justify-center text-sm" }, "A"),
                    "SGC ACTIVA"),
                React.createElement("nav", { className: "ml-auto flex gap-4 text-sm font-semibold" },
                    React.createElement("button", { onClick: () => irA(''), className: "hover:text-[#1E6B47]" }, "Inicio"),
                    React.createElement("button", { onClick: () => irA('riesgos'), className: "hover:text-[#1E6B47]" }, "Riesgos"),
                    React.createElement("button", { onClick: () => irA('documentos'), className: "hover:text-[#1E6B47]" }, "Documentos")))),
        React.createElement("main", { className: "flex-1 px-4 py-6" }, contenido),
        React.createElement("footer", { className: "border-t border-[#DCE5DC] bg-white" },
            React.createElement("div", { className: "max-w-4xl mx-auto px-4 py-5 text-sm text-[#5b6b5f] space-y-1" },
                React.createElement("p", null, "Empresa de Parques y Eventos de Antioquia \u2014 ACTIVA \u00B7 Portal de consulta del SGC."),
                React.createElement("p", null,
                    "Los enlaces abren el ",
                    React.createElement("a", { className: "font-semibold text-[#1E6B47] hover:underline", href: CARPETA_SGC_DEFAULT, target: "_blank", rel: "noopener" }, "repositorio oficial en SharePoint"),
                    " y requieren sesi\u00F3n Microsoft institucional."),
                React.createElement("p", { className: "f-mono text-xs" },
                    "Inventario al corte ",
                    FECHA_CORTE_DEFAULT,
                    " \u00B7 Los flujogramas resumen el procedimiento fuente.")))));
};
ReactDOM.createRoot(document.getElementById('root')).render(React.createElement(App, null));
