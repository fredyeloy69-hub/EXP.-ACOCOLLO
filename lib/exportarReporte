import jsPDF from "jspdf";
import autoTable from "jspdf-autotable";
import { LOGO_PUNO_BASE64 } from "./logoPuno";

const ESTADO_LABEL = {
  completa: "COMPLETA",
  incompleta: "INCOMPLETA",
  vacia: "VACÍA",
};

const ESTADO_RGB = {
  completa: [46, 204, 113],
  incompleta: [243, 156, 18],
  vacia: [231, 76, 60],
};

const NARANJA_HEADER = [230, 126, 34];
const TEAL = [23, 163, 152];
const TEAL_CLARO = [45, 212, 191];
const VERDE = [46, 204, 113];
const NARANJA = [243, 156, 18];
const ROJO = [231, 76, 60];
const AZUL_OSCURO = [20, 30, 60];

export function generarReportePorArea(areaNombre, carpetasDelArea, opts = {}) {
  const {
    logoBase64 = LOGO_PUNO_BASE64,
    proyectoNombre =
      '"MEJORAMIENTO DEL SERVICIO DE ATENCION DE SALUD BASICOS EN CHIJNAYA DISTRITO DE PUCARA DE LA PROVINCIA DE LAMPA DEL DEPARTAMENTO DE PUNO"',
  } = opts;

  const doc = new jsPDF({ orientation: "portrait", unit: "mm", format: "a4" });
  const pageWidth = doc.internal.pageSize.getWidth();
  const pageHeight = doc.internal.pageSize.getHeight();
  const marginX = 14;

  const grupos = {};
  const ordenGrupos = [];
  for (const c of carpetasDelArea) {
    const partes = (c.ruta || c.nombre || "").split(" / ").filter(Boolean);
    const especialidad = partes.length > 1 ? partes[1] : "(raíz)";
    if (!grupos[especialidad]) {
      grupos[especialidad] = [];
      ordenGrupos.push(especialidad);
    }
    grupos[especialidad].push(c);
  }
  ordenGrupos.sort((a, b) => a.localeCompare(b, undefined, { numeric: true, sensitivity: "base" }));

  const total = carpetasDelArea.length;
  const completas = carpetasDelArea.filter((c) => c.estado === "completa").length;
  const incompletas = carpetasDelArea.filter((c) => c.estado === "incompleta").length;
  const vacias = carpetasDelArea.filter((c) => c.estado === "vacia").length;
  let necesarios = 0;
  let completados = 0;
  for (const c of carpetasDelArea) {
    necesarios += c.archivosNecesarios || 0;
    completados += c.archivosCompletados || 0;
  }
  const pctArchivos = necesarios > 0 ? Math.round((completados / necesarios) * 100) : 0;

  const body = [];
  let contadorFila = 1;

  // Compara nombres respetando numeración real (01,02,10 / Anexo1,Anexo2,Anexo10 / etc.)
  function comparaNatural(a, b) {
    return (a || "").localeCompare(b || "", undefined, { numeric: true, sensitivity: "base" });
  }

  // Mezcla un color con blanco según un factor (0 = color puro, 1 = blanco puro) —
  // para crear el efecto "cada vez más pálido/transparente" según el nivel.
  function mezclarConBlanco(rgb, factor) {
    return rgb.map((c) => Math.round(c + (255 - c) * factor));
  }

  // Colores de fondo para las carpetas FINALES (las filas con COMPLETA/INCOMPLETA/VACÍA),
  // no solo para los encabezados de subcarpeta — mientras más profundo el nivel, más pálido.
  const FACTORES_HOJA = [0.72, 0.80, 0.86, 0.90, 0.93];
  function colorHoja(nivelVisual) {
    const factor = FACTORES_HOJA[Math.min(nivelVisual - 1, FACTORES_HOJA.length - 1)];
    return mezclarConBlanco(NARANJA_HEADER, factor);
  }

  function empujarFila(c, nivelVisual) {
    const estado = c.estado || "incompleta";
    const nombreMostrado = c.nombre || (c.ruta || "").split(" / ").pop() || "-";
    const sangriaIzq = 2 + Math.max(0, nivelVisual - 1) * 6; // sangría real en mm, no espacios de texto
    const fondo = colorHoja(nivelVisual);
    body.push([
      { content: String(contadorFila++), styles: { fillColor: fondo } },
      { content: nombreMostrado, styles: { fillColor: fondo, cellPadding: { top: 2, right: 2, bottom: 2, left: sangriaIzq } } },
      {
        content: ESTADO_LABEL[estado] || estado.toUpperCase(),
        styles: { fillColor: fondo, textColor: ESTADO_RGB[estado] || [100, 100, 100], fontStyle: "bold", halign: "center" },
      },
      { content: c.detalle || "-", styles: { fillColor: fondo } },
    ]);
  }

  function empujarSubEncabezado(nombre, nivel) {
    // Mismo tono naranja de la carpeta madre, pero cada vez más pálido mientras
    // más profundo el nivel (2, 3, 4) — la carpeta madre es la más nítida.
    const factoresPorNivel = [0.45, 0.65, 0.82]; // nivel 2, 3, 4
    const factor = factoresPorNivel[Math.min(nivel - 2, factoresPorNivel.length - 1)];
    const color = mezclarConBlanco(NARANJA_HEADER, factor);
    const sangriaIzq = 2 + (nivel - 1) * 6; // sangría real en mm — cada nivel más adentro que el anterior, tipo escalera
    body.push([
      {
        content: `➤  ${nombre}`,
        colSpan: 4,
        styles: {
          fillColor: color,
          textColor: [70, 45, 15],
          fontStyle: "bold",
          fontSize: 9.5,
          halign: "left",
          cellPadding: { top: 3, right: 2, bottom: 3, left: sangriaIzq },
        },
      },
    ]);
  }

  // Agrupa recursivamente por cada segmento de ruta, hasta NIVEL_MAX_INTERMEDIO
  // niveles intermedios (especialidad=nivel1, hasta 3 intermedios más = nivel4,
  // y el ítem final = nivel5 → 5 niveles totales, como se pidió). TODO se
  // intercala en un solo orden natural (01,02,03... o Anexo1,Anexo2...) — un
  // ítem suelto (ej. "2.6") y una subcarpeta con más adentro (ej. "2.7") se
  // ordenan juntos por su propio nombre, no separados en dos grupos distintos
  // (eso era el bug: antes todos los "sueltos" salían primero, sin importar numeración).
  const NIVEL_MAX_INTERMEDIO = 4; // índice de partes[] hasta donde se agrupa (2,3,4)
  function agruparRecursivo(items, nivelIdx, nivelVisual) {
    if (nivelIdx > NIVEL_MAX_INTERMEDIO) {
      const ordenado = [...items].sort((a, b) => comparaNatural(a.nombre, b.nombre));
      ordenado.forEach((c) => empujarFila(c, nivelVisual));
      return;
    }

    const subgrupos = {};
    let hayNivelMasProfundo = false;
    for (const c of items) {
      const partes = (c.ruta || c.nombre || "").split(" / ").filter(Boolean);
      // partes[nivelIdx] existe y NO es el último segmento (el nombre final) → hay un nivel más
      const clave = partes.length > nivelIdx + 1 ? partes[nivelIdx] : null;
      if (clave) hayNivelMasProfundo = true;
      const key = clave || `__directo__${c.nombre || c.id}`; // clave única por ítem suelto
      if (!subgrupos[key]) subgrupos[key] = [];
      subgrupos[key].push(c);
    }

    if (!hayNivelMasProfundo) {
      const ordenado = [...items].sort((a, b) => comparaNatural(a.nombre, b.nombre));
      ordenado.forEach((c) => empujarFila(c, nivelVisual));
      return;
    }

    // Cada entrada (sea ítem suelto o subgrupo) se ordena por SU PROPIO nombre,
    // todas juntas — así queda exactamente el orden real del Drive.
    const entradas = Object.keys(subgrupos).map((key) => ({
      key,
      nombreOrden: key.startsWith("__directo__") ? subgrupos[key][0].nombre || "" : key,
      esGrupo: !key.startsWith("__directo__"),
    }));
    entradas.sort((a, b) => comparaNatural(a.nombreOrden, b.nombreOrden));

    for (const entrada of entradas) {
      if (entrada.esGrupo) {
        empujarSubEncabezado(entrada.key, nivelIdx);
        agruparRecursivo(subgrupos[entrada.key], nivelIdx + 1, nivelVisual + 1);
      } else {
        const ordenado = [...subgrupos[entrada.key]].sort((a, b) => comparaNatural(a.nombre, b.nombre));
        ordenado.forEach((c) => empujarFila(c, nivelVisual));
      }
    }
  }

  for (const especialidad of ordenGrupos) {
    body.push([
      { content: especialidad.toUpperCase(), colSpan: 4, styles: { fillColor: NARANJA_HEADER, textColor: [255, 255, 255], fontStyle: "bold", fontSize: 11, halign: "left", cellPadding: 3 } },
    ]);
    agruparRecursivo(grupos[especialidad], 2, 1);
  }

  function dibujarBloqueInstitucional() {
    let y = 12;
    if (logoBase64) {
      const anchoLogo = 13;
      const altoLogo = 14.6;
      doc.addImage(logoBase64, "PNG", marginX, y - 3, anchoLogo, altoLogo);
    }
    doc.setFont("helvetica", "bold");
    doc.setFontSize(10);
    doc.text("GOBIERNO REGIONAL DE PUNO GERENCIA REGIONAL DE INFRAESTRUCTURA", pageWidth / 2, y, { align: "center" });
    y += 5;
    doc.text("SUB GERENCIA DE ESTUDIOS DEFINITIVOS", pageWidth / 2, y, { align: "center" });
    y += 6;
    doc.setFontSize(8.5);
    const lineasProyecto = doc.splitTextToSize(proyectoNombre, pageWidth - marginX * 2 - 10);
    doc.text(lineasProyecto, pageWidth / 2, y, { align: "center" });
    y += lineasProyecto.length * 4 + 4;
    return y;
  }

  function dibujarEncabezado() {
    let y = dibujarBloqueInstitucional();

    doc.setFillColor(...AZUL_OSCURO);
    doc.rect(marginX, y, pageWidth - marginX * 2, 9, "F");
    doc.setTextColor(255, 255, 255);
    doc.setFontSize(11);
    doc.text(`REPORTE DE AVANCE — ${areaNombre.toUpperCase()}`, pageWidth / 2, y + 6.2, { align: "center" });
    doc.setTextColor(0, 0, 0);
    y += 16;

    const fecha = new Date().toLocaleString("es-PE");
    doc.setFont("helvetica", "normal");
    doc.setFontSize(8);
    doc.text(`Generado: ${fecha}`, marginX, y);
    doc.text(`Total: ${total}  ·  Completas: ${completas}  ·  Pendientes: ${total - completas}`, pageWidth - marginX, y, { align: "right" });
    return y + 10;
  }

  function dibujarBarra(x, y, ancho, alto, pct, colorRgb, colorFondoRgb) {
    const fondo = colorFondoRgb || [20, 37, 41];
    doc.setFillColor(...fondo);
    doc.roundedRect(x, y, ancho, alto, alto / 2, alto / 2, "F");
    const anchoLleno = Math.max(alto, (ancho * Math.min(100, Math.max(0, pct))) / 100);
    doc.setFillColor(...colorRgb);
    doc.roundedRect(x, y, anchoLleno, alto, alto / 2, alto / 2, "F");
  }

  function dibujarTarjeta(x, y, ancho, alto, valor, etiqueta, colorRgb) {
    doc.setDrawColor(...colorRgb);
    doc.setLineWidth(0.6);
    doc.roundedRect(x, y, ancho, alto, 2, 2, "S");
    doc.setFont("helvetica", "bold");
    doc.setFontSize(20);
    doc.setTextColor(...colorRgb);
    doc.text(String(valor), x + ancho / 2, y + alto / 2 - 1, { align: "center" });
    doc.setFont("helvetica", "normal");
    doc.setFontSize(8);
    doc.setTextColor(90, 90, 90);
    doc.text(etiqueta, x + ancho / 2, y + alto - 5, { align: "center" });
    doc.setTextColor(0, 0, 0);
  }

  // ============================================================
  // PORTADA — resumen gráfico, como el modo presentación del dashboard
  // ============================================================
  let y = dibujarBloqueInstitucional();

  doc.setFillColor(...AZUL_OSCURO);
  doc.rect(marginX, y, pageWidth - marginX * 2, 11, "F");
  doc.setTextColor(255, 255, 255);
  doc.setFont("helvetica", "bold");
  doc.setFontSize(13);
  doc.text(`RESUMEN DE AVANCE — ${areaNombre.toUpperCase()}`, pageWidth / 2, y + 7.5, { align: "center" });
  doc.setTextColor(0, 0, 0);
  y += 20;

  const fechaGenerado = new Date().toLocaleString("es-PE");
  doc.setFont("helvetica", "normal");
  doc.setFontSize(8.5);
  doc.setTextColor(100, 100, 100);
  doc.text(`Generado: ${fechaGenerado}`, pageWidth / 2, y, { align: "center" });
  doc.setTextColor(0, 0, 0);
  y += 12;

  const pctCarpetas = total > 0 ? Math.round((completas / total) * 100) : 0;

  doc.setFont("helvetica", "bold");
  doc.setFontSize(10);
  doc.text("AVANCE POR CARPETAS", marginX, y);
  doc.setFontSize(16);
  doc.setTextColor(...TEAL);
  doc.text(`${pctCarpetas}%`, pageWidth - marginX, y, { align: "right" });
  doc.setTextColor(0, 0, 0);
  y += 4;
  dibujarBarra(marginX, y, pageWidth - marginX * 2, 6, pctCarpetas, TEAL);
  y += 14;

  doc.setFont("helvetica", "bold");
  doc.setFontSize(10);
  doc.text("AVANCE POR ARCHIVOS (más preciso)", marginX, y);
  doc.setFontSize(16);
  doc.setTextColor(...TEAL_CLARO);
  doc.text(`${pctArchivos}%`, pageWidth - marginX, y, { align: "right" });
  doc.setTextColor(0, 0, 0);
  y += 4;
  dibujarBarra(marginX, y, pageWidth - marginX * 2, 6, pctArchivos, TEAL_CLARO);
  y += 18;

  const anchoTarjeta = (pageWidth - marginX * 2 - 3 * 4) / 4;
  const altoTarjeta = 22;
  dibujarTarjeta(marginX, y, anchoTarjeta, altoTarjeta, total, "CARPETAS", TEAL);
  dibujarTarjeta(marginX + (anchoTarjeta + 4) * 1, y, anchoTarjeta, altoTarjeta, completas, "COMPLETAS", VERDE);
  dibujarTarjeta(marginX + (anchoTarjeta + 4) * 2, y, anchoTarjeta, altoTarjeta, incompletas, "INCOMPLETAS", NARANJA);
  dibujarTarjeta(marginX + (anchoTarjeta + 4) * 3, y, anchoTarjeta, altoTarjeta, vacias, "VACÍAS", ROJO);
  y += altoTarjeta + 14;

  doc.setFont("helvetica", "bold");
  doc.setFontSize(11);
  doc.text(`AVANCE POR ESPECIALIDAD (${ordenGrupos.length})`, marginX, y);
  y += 8;

  doc.setFont("helvetica", "normal");
  doc.setFontSize(8);
  for (const especialidad of ordenGrupos) {
    if (y > pageHeight - 20) {
      doc.addPage();
      y = 18;
    }
    const items = grupos[especialidad];
    let espNecesarios = 0;
    let espCompletados = 0;
    for (const c of items) {
      espNecesarios += c.archivosNecesarios || 0;
      espCompletados += c.archivosCompletados || 0;
    }
    const pctEsp = espNecesarios > 0 ? Math.round((espCompletados / espNecesarios) * 100) : 0;
    const colorEsp = pctEsp >= 100 ? VERDE : pctEsp >= 50 ? TEAL_CLARO : NARANJA;

    const etiqueta = doc.splitTextToSize(especialidad, 70)[0];
    doc.setTextColor(60, 60, 60);
    doc.text(etiqueta, marginX, y + 3);
    doc.setTextColor(0, 0, 0);

    const barraX = marginX + 74;
    const barraAncho = pageWidth - marginX - 22 - barraX;
    dibujarBarra(barraX, y, barraAncho, 4.5, pctEsp, colorEsp);

    doc.setFont("helvetica", "bold");
    doc.setTextColor(...colorEsp);
    doc.text(`${pctEsp}%`, pageWidth - marginX, y + 3.3, { align: "right" });
    doc.setTextColor(0, 0, 0);
    doc.setFont("helvetica", "normal");

    y += 8.5;
  }

  // ============================================================
  // DETALLE — tabla carpeta por carpeta (empieza en página nueva)
  // ============================================================
  doc.addPage();
  const primeraPaginaTabla = doc.internal.getCurrentPageInfo().pageNumber;
  const startY = dibujarEncabezado();

  autoTable(doc, {
    startY,
    margin: { left: marginX, right: marginX, top: startY },
    head: [["N°", "DESCRIPCIÓN", "ESTADO", "DETALLE"]],
    body,
    styles: { fontSize: 8.5, cellPadding: 2, valign: "middle" },
    headStyles: { fillColor: AZUL_OSCURO, textColor: 255, fontStyle: "bold" },
    columnStyles: {
      0: { cellWidth: 12, halign: "center" },
      1: { cellWidth: "auto" },
      2: { cellWidth: 26 },
      3: { cellWidth: 55 },
    },
    didDrawPage: () => {
      if (doc.internal.getCurrentPageInfo().pageNumber > primeraPaginaTabla) {
        dibujarEncabezado();
      }
    },
  });

  const totalPaginas = doc.internal.getNumberOfPages();
  for (let p = 1; p <= totalPaginas; p++) {
    doc.setPage(p);
    doc.setFontSize(7.5);
    doc.setTextColor(120, 120, 120);
    doc.text(`Página ${p} de ${totalPaginas}`, pageWidth - marginX, pageHeight - 8, { align: "right" });
    doc.setTextColor(0, 0, 0);
  }

  const nombreArchivo = `Reporte_${areaNombre.replace(/[^a-zA-Z0-9]+/g, "_")}_${new Date()
    .toISOString()
    .slice(0, 10)}.pdf`;
  doc.save(nombreArchivo);
}
