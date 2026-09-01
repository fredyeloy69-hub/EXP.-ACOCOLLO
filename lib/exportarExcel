import ExcelJS from "exceljs";

const ESTADO_LABEL = {
  completa: "COMPLETA",
  incompleta: "INCOMPLETA",
  vacia: "VACÍA",
};

const ESTADO_COLOR = {
  completa: "FF2ECC71",
  incompleta: "FFF39C12",
  vacia: "FFE74C3C",
};

const AZUL_OSCURO = "FF141E3C";
const NARANJA = "FFE67E22";

/**
 * Genera y descarga un Excel (.xlsx) con formato real — encabezado institucional,
 * jerarquía de hasta 5 niveles de carpetas (igual que el PDF), colores por estado,
 * bordes y columnas ajustadas. Reemplaza la versión anterior que solo llenaba
 * celdas sin ningún formato (la librería "xlsx" no soporta estilos al escribir).
 *
 * @param {string} areaNombre - nombre del área a exportar
 * @param {Array} carpetasDelArea - carpetas (ya filtradas a esa área)
 */
export async function generarReporteExcelPorArea(areaNombre, carpetasDelArea) {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = "Visor Chijnaya";
  workbook.created = new Date();

  const nombreHoja = areaNombre.replace(/[\\/*?:[\]]/g, "").slice(0, 31) || "Reporte";
  const sheet = workbook.addWorksheet(nombreHoja, {
    pageSetup: { orientation: "portrait", fitToPage: true, fitToWidth: 1 },
  });

  sheet.columns = [
    { width: 6 },
    { width: 55 },
    { width: 16 },
    { width: 55 },
  ];

  // --- Encabezado institucional ---
  sheet.mergeCells("A1:D1");
  sheet.getCell("A1").value = "GOBIERNO REGIONAL DE PUNO — GERENCIA REGIONAL DE INFRAESTRUCTURA";
  sheet.getCell("A1").font = { bold: true, size: 12 };
  sheet.getCell("A1").alignment = { horizontal: "center", vertical: "middle" };
  sheet.getRow(1).height = 20;

  sheet.mergeCells("A2:D2");
  sheet.getCell("A2").value = "SUB GERENCIA DE ESTUDIOS DEFINITIVOS";
  sheet.getCell("A2").font = { bold: true, size: 10 };
  sheet.getCell("A2").alignment = { horizontal: "center", vertical: "middle" };

  sheet.mergeCells("A3:D3");
  const celdaTitulo = sheet.getCell("A3");
  celdaTitulo.value = `REPORTE DE AVANCE — ${areaNombre.toUpperCase()}`;
  celdaTitulo.font = { bold: true, size: 13, color: { argb: "FFFFFFFF" } };
  celdaTitulo.alignment = { horizontal: "center", vertical: "middle" };
  celdaTitulo.fill = { type: "pattern", pattern: "solid", fgColor: { argb: AZUL_OSCURO } };
  sheet.getRow(3).height = 24;

  const total = carpetasDelArea.length;
  const completas = carpetasDelArea.filter((c) => c.estado === "completa").length;
  const incompletas = carpetasDelArea.filter((c) => c.estado === "incompleta").length;
  const vacias = carpetasDelArea.filter((c) => c.estado === "vacia").length;

  sheet.mergeCells("A4:B4");
  sheet.getCell("A4").value = `Generado: ${new Date().toLocaleString("es-PE")}`;
  sheet.getCell("A4").font = { italic: true, size: 9, color: { argb: "FF666666" } };

  sheet.mergeCells("C4:D4");
  sheet.getCell("C4").value = `Total: ${total}  ·  Completas: ${completas}  ·  Incompletas: ${incompletas}  ·  Vacías: ${vacias}`;
  sheet.getCell("C4").font = { italic: true, size: 9, color: { argb: "FF666666" } };
  sheet.getCell("C4").alignment = { horizontal: "right" };

  sheet.addRow([]); // fila 5 en blanco, de separación

  // --- Encabezado de columnas de la tabla (fila 6) ---
  const FILA_ENCABEZADO = 6;
  const filaEncabezado = sheet.getRow(FILA_ENCABEZADO);
  filaEncabezado.values = ["N°", "DESCRIPCIÓN", "ESTADO", "DETALLE"];
  filaEncabezado.eachCell((cell) => {
    cell.font = { bold: true, color: { argb: "FFFFFFFF" } };
    cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: AZUL_OSCURO } };
    cell.alignment = { vertical: "middle", horizontal: "center" };
    cell.border = { bottom: { style: "medium", color: { argb: "FF000000" } } };
  });
  filaEncabezado.height = 18;

  // --- Agrupar por especialidad, igual que en el PDF ---
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

  function comparaNatural(a, b) {
    return (a || "").localeCompare(b || "", undefined, { numeric: true, sensitivity: "base" });
  }
  ordenGrupos.sort(comparaNatural);

  let contadorFila = 1;

  // Mezcla el naranja con blanco según un factor (0 = naranja puro, 1 = blanco puro) —
  // mismo criterio de degradé que el PDF, para que la carpeta madre sea la más
  // nítida y cada nivel más profundo se vea más pálido.
  function mezclarConBlanco(hexRgb, factor) {
    const r = parseInt(hexRgb.slice(2, 4), 16);
    const g = parseInt(hexRgb.slice(4, 6), 16);
    const b = parseInt(hexRgb.slice(6, 8), 16);
    const mezcla = (c) => Math.round(c + (255 - c) * factor).toString(16).padStart(2, "0").toUpperCase();
    return `FF${mezcla(r)}${mezcla(g)}${mezcla(b)}`;
  }

  // Colores de fondo para las carpetas FINALES (las filas con COMPLETA/INCOMPLETA/VACÍA),
  // no solo los encabezados de subcarpeta — mientras más profundo el nivel, más pálido.
  const FACTORES_HOJA = [0.72, 0.80, 0.86, 0.90, 0.93];
  function colorHoja(nivelVisual) {
    const factor = FACTORES_HOJA[Math.min((nivelVisual || 1) - 1, FACTORES_HOJA.length - 1)];
    return mezclarConBlanco(NARANJA, factor);
  }

  function agregarSubEncabezado(nombre, nivel) {
    const fila = sheet.addRow([`➤  ${nombre}`]);
    sheet.mergeCells(`A${fila.number}:D${fila.number}`);
    const celda = fila.getCell(1);
    const esNivel1 = nivel === 1;
    const factoresPorNivel = [0, 0.45, 0.65, 0.82]; // nivel 1 = naranja puro, luego cada vez más pálido
    const factor = factoresPorNivel[Math.min(nivel - 1, factoresPorNivel.length - 1)];
    celda.font = { bold: true, size: esNivel1 ? 13 : 11, color: { argb: esNivel1 ? "FFFFFFFF" : "FF4A2D0F" } };
    celda.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: mezclarConBlanco(NARANJA, factor) },
    };
    // Sangría NATIVA de Excel (no espacios de texto) — cada nivel más adentro que el anterior, tipo escalera
    celda.alignment = { horizontal: "left", vertical: "middle", indent: Math.max(0, nivel - 1) * 2 };
    fila.height = esNivel1 ? 20 : 17;
  }

  function agregarFila(c, nivelVisual) {
    const estado = c.estado || "incompleta";
    const nombreMostrado = c.nombre || (c.ruta || "").split(" / ").pop() || "-";
    const fila = sheet.addRow([contadorFila++, nombreMostrado, ESTADO_LABEL[estado] || estado.toUpperCase(), c.detalle || "-"]);
    const fondo = colorHoja(nivelVisual);

    fila.getCell(1).alignment = { horizontal: "center", vertical: "top" };
    // Sangría nativa también en las filas finales, según qué tan adentro estén en el ramal
    fila.getCell(2).alignment = { wrapText: true, vertical: "top", indent: Math.max(0, (nivelVisual || 1) - 1) * 2 };
    fila.getCell(3).font = { bold: true, color: { argb: ESTADO_COLOR[estado] || "FF666666" } };
    fila.getCell(3).alignment = { horizontal: "center", vertical: "top" };
    fila.getCell(4).alignment = { wrapText: true, vertical: "top" };

    fila.eachCell((cell) => {
      cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: fondo } };
      cell.border = {
        top: { style: "hair", color: { argb: "FFDDDDDD" } },
        bottom: { style: "hair", color: { argb: "FFDDDDDD" } },
        left: { style: "hair", color: { argb: "FFDDDDDD" } },
        right: { style: "hair", color: { argb: "FFDDDDDD" } },
      };
    });
  }

  // Agrupación recursiva hasta 5 niveles totales (especialidad + 3 intermedios + ítem final),
  // ordenada numéricamente en cada nivel — igual criterio que en el PDF.
  const NIVEL_MAX_INTERMEDIO = 4;
  function agruparRecursivo(items, nivelIdxRuta, nivelVisual) {
    if (nivelIdxRuta > NIVEL_MAX_INTERMEDIO) {
      const ordenado = [...items].sort((a, b) => comparaNatural(a.nombre, b.nombre));
      ordenado.forEach((c) => agregarFila(c, nivelVisual));
      return;
    }

    const subgrupos = {};
    let hayNivelMasProfundo = false;
    for (const c of items) {
      const partes = (c.ruta || c.nombre || "").split(" / ").filter(Boolean);
      const clave = partes.length > nivelIdxRuta + 1 ? partes[nivelIdxRuta] : null;
      if (clave) hayNivelMasProfundo = true;
      const key = clave || `__directo__${c.nombre || c.id}`; // clave única por ítem suelto
      if (!subgrupos[key]) subgrupos[key] = [];
      subgrupos[key].push(c);
    }

    if (!hayNivelMasProfundo) {
      const ordenado = [...items].sort((a, b) => comparaNatural(a.nombre, b.nombre));
      ordenado.forEach((c) => agregarFila(c, nivelVisual));
      return;
    }

    // Cada entrada (ítem suelto o subgrupo) se ordena por SU PROPIO nombre, todas
    // juntas — antes los ítems sueltos salían todos primero sin importar numeración.
    const entradas = Object.keys(subgrupos).map((key) => ({
      key,
      nombreOrden: key.startsWith("__directo__") ? subgrupos[key][0].nombre || "" : key,
      esGrupo: !key.startsWith("__directo__"),
    }));
    entradas.sort((a, b) => comparaNatural(a.nombreOrden, b.nombreOrden));

    for (const entrada of entradas) {
      if (entrada.esGrupo) {
        agregarSubEncabezado(entrada.key, nivelVisual);
        agruparRecursivo(subgrupos[entrada.key], nivelIdxRuta + 1, nivelVisual + 1);
      } else {
        const ordenado = [...subgrupos[entrada.key]].sort((a, b) => comparaNatural(a.nombre, b.nombre));
        ordenado.forEach((c) => agregarFila(c, nivelVisual));
      }
    }
  }

  for (const especialidad of ordenGrupos) {
    agregarSubEncabezado(especialidad.toUpperCase(), 1);
    agruparRecursivo(grupos[especialidad], 2, 2);
  }

  // Fija el encabezado de columnas al hacer scroll
  sheet.views = [{ state: "frozen", ySplit: FILA_ENCABEZADO }];

  // --- Descargar ---
  const buffer = await workbook.xlsx.writeBuffer();
  const blob = new Blob([buffer], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = `Reporte_${areaNombre.replace(/[^a-zA-Z0-9]+/g, "_")}_${new Date()
    .toISOString()
    .slice(0, 10)}.xlsx`;
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  URL.revokeObjectURL(url);
}
