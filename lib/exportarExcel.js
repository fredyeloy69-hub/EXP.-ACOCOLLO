import ExcelJS from "exceljs";
import { LOGO_PUNO_BASE64 } from "@/lib/logoPuno";

const ESTADO_LABEL = {
  completa: "COMPLETA",
  incompleta: "INCOMPLETA",
  vacia: "VACÍA",
};

const ESTADO_COLOR = {
  completa: "FF2A9D8F",
  incompleta: "FFF39C12",
  vacia: "FFC0392B",
};

const AZUL_PETROLEO = "FF4A0E17";
const VERDE_AZULADO = "FF7A1F2B";

const CODIGO_PROYECTO = "ACASO";

function sufijoFiltroArchivo(tipoFiltro) {
  const mapa = {
    todas: "Todo",
    completas: "Solo_Completas",
    incompletas: "Solo_Incompletas",
    vacias: "Solo_Vacias",
    incompletas_vacias: "Incompletas_Vacias",
  };
  return mapa[tipoFiltro] || "Todo";
}

/**
 * Genera y descarga un Excel (.xlsx) con formato real — encabezado institucional,
 * logo, nombre del proyecto, jerarquía de carpetas, colores por estado, bordes y notas al pie visibles.
 *
 * @param {string} areaNombre - nombre del área a exportar
 * @param {Array} carpetasDelArea - carpetas (ya filtradas a esa área)
 * @param {string} tipoFiltro - filtro aplicado ("todas" | "completas" | "incompletas" | "vacias" | "incompletas_vacias"), solo afecta el nombre del archivo
 */
export async function generarReporteExcelPorArea(areaNombre, carpetasDelArea, tipoFiltro = "todas") {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = "Visor Acocollo I-2";
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

  // --- Insertar logo institucional ---
  if (LOGO_PUNO_BASE64) {
    try {
      const base64Data = LOGO_PUNO_BASE64.includes("base64,")
        ? LOGO_PUNO_BASE64.split("base64,")[1]
        : LOGO_PUNO_BASE64;
      const imageId = workbook.addImage({
        base64: base64Data,
        extension: "png",
      });
      sheet.addImage(imageId, {
        tl: { col: 0.2, row: 0.2 },
        ext: { width: 45, height: 50 },
      });
    } catch (e) {
      console.error("No se pudo cargar el logo en el Excel:", e);
    }
  }

  // --- Encabezado institucional ---
  sheet.mergeCells("A1:D1");
  sheet.getCell("A1").value = "GOBIERNO REGIONAL DE PUNO — GERENCIA REGIONAL DE INFRAESTRUCTURA";
  sheet.getCell("A1").font = { bold: true, size: 11 };
  sheet.getCell("A1").alignment = { horizontal: "center", vertical: "middle" };
  sheet.getRow(1).height = 18;

  sheet.mergeCells("A2:D2");
  sheet.getCell("A2").value = "SUB GERENCIA DE ESTUDIOS DEFINITIVOS";
  sheet.getCell("A2").font = { bold: true, size: 10 };
  sheet.getCell("A2").alignment = { horizontal: "center", vertical: "middle" };
  sheet.getRow(2).height = 16;

  sheet.mergeCells("A3:D3");
  sheet.getCell("A3").value = "PROYECTO: MEJORAMIENTO DEL SERVICIO DE ATENCION DE SALUD BASICOS EN ACOCOLLO DISTRITO DE HUANCANE DE LA PROVINCIA DE HUANCANE DEL DEPARTAMENTO DE PUNO";
  sheet.getCell("A3").font = { bold: true, size: 10, color: { argb: "FF7A1F2B" } };
  sheet.getCell("A3").alignment = { horizontal: "center", vertical: "middle", wrapText: true };
  sheet.getRow(3).height = 30;

  sheet.mergeCells("A4:D4");
  const celdaTitulo = sheet.getCell("A4");
  celdaTitulo.value = `REPORTE DE AVANCE — ${areaNombre.toUpperCase()}`;
  celdaTitulo.font = { bold: true, size: 12, color: { argb: "FFFFFFFF" } };
  celdaTitulo.alignment = { horizontal: "center", vertical: "middle" };
  celdaTitulo.fill = { type: "pattern", pattern: "solid", fgColor: { argb: AZUL_PETROLEO } };
  sheet.getRow(4).height = 22;

  const total = carpetasDelArea.length;
  const completas = carpetasDelArea.filter((c) => c.estado === "completa").length;
  const incompletas = carpetasDelArea.filter((c) => c.estado === "incompleta").length;
  const vacias = carpetasDelArea.filter((c) => c.estado === "vacia").length;

  sheet.mergeCells("A5:B5");
  sheet.getCell("A5").value = `Generado: ${new Date().toLocaleString("es-PE")}`;
  sheet.getCell("A5").font = { italic: true, size: 9, color: { argb: "FF666666" } };

  sheet.mergeCells("C5:D5");
  sheet.getCell("C5").value = `Total: ${total}  ·  Completas: ${completas}  ·  Incompletas: ${incompletas}  ·  Vacías: ${vacias}`;
  sheet.getCell("C5").font = { italic: true, size: 9, color: { argb: "FF666666" } };
  sheet.getCell("C5").alignment = { horizontal: "right" };
  sheet.getRow(5).height = 16;

  sheet.addRow([]); // fila 6 en blanco, de separación

  // --- Encabezado de columnas de la tabla (fila 7) ---
  const FILA_ENCABEZADO = 7;
  const filaEncabezado = sheet.getRow(FILA_ENCABEZADO);
  filaEncabezado.values = ["N°", "DESCRIPCIÓN", "ESTADO", "DETALLE"];
  filaEncabezado.eachCell((cell) => {
    cell.font = { bold: true, color: { argb: "FFFFFFFF" } };
    cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: AZUL_PETROLEO } };
    cell.alignment = { vertical: "middle", horizontal: "center" };
    cell.border = { bottom: { style: "medium", color: { argb: "FF000000" } } };
  });
  filaEncabezado.height = 18;

  // --- Agrupar por especialidad ---
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

  function mezclarConBlanco(hexRgb, factor) {
    const r = parseInt(hexRgb.slice(2, 4), 16);
    const g = parseInt(hexRgb.slice(4, 6), 16);
    const b = parseInt(hexRgb.slice(6, 8), 16);
    const mezcla = (c) => Math.round(c + (255 - c) * factor).toString(16).padStart(2, "0").toUpperCase();
    return `FF${mezcla(r)}${mezcla(g)}${mezcla(b)}`;
  }

  const FACTORES_HOJA = [0.72, 0.80, 0.86, 0.90, 0.93];
  function colorHoja(nivelVisual) {
    const factor = FACTORES_HOJA[Math.min((nivelVisual || 1) - 1, FACTORES_HOJA.length - 1)];
    return mezclarConBlanco(VERDE_AZULADO, factor);
  }

  function contarEstados(items) {
    let completas = 0, incompletas = 0, vacias = 0;
    for (const c of items) {
      if (c.estado === "completa") completas++;
      else if (c.estado === "incompleta") incompletas++;
      else if (c.estado === "vacia") vacias++;
    }
    return { completas, incompletas, vacias };
  }

  function textoResumen(items) {
    const { completas, incompletas, vacias } = contarEstados(items);
    const total = items.length;
    return `  (Total: ${total} — ${completas} completas · ${incompletas} incompletas · ${vacias} vacías)`;
  }

  function agregarSubEncabezado(nombre, nivel, items) {
    // El resumen numérico (completas/incompletas/vacías) solo se muestra en
    // los títulos de nivel 1 y 2 (carpetas mayores) — de ahí para abajo se
    // omite para no saturar los sub-encabezados.
    const resumen = items && nivel <= 2 ? textoResumen(items) : "";
    const fila = sheet.addRow([`➤  ${nombre}${resumen}`]);
    sheet.mergeCells(`A${fila.number}:D${fila.number}`);
    const celda = fila.getCell(1);
    const esNivel1 = nivel === 1;
    const factoresPorNivel = [0, 0.45, 0.65, 0.82];
    const factor = factoresPorNivel[Math.min(nivel - 1, factoresPorNivel.length - 1)];
    celda.font = { bold: true, size: esNivel1 ? 13 : 11, color: { argb: esNivel1 ? "FFFFFFFF" : "FF1A3840" } };
    celda.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: mezclarConBlanco(VERDE_AZULADO, factor) },
    };
    celda.alignment = { horizontal: "left", vertical: "middle", indent: Math.max(0, nivel - 1) * 2 };
    fila.height = esNivel1 ? 20 : 17;
  }

  function agregarFila(c, nivelVisual) {
    const estado = c.estado || "incompleta";
    const nombreMostrado = c.nombre || (c.ruta || "").split(" / ").pop() || "-";
    const fila = sheet.addRow([contadorFila++, nombreMostrado, ESTADO_LABEL[estado] || estado.toUpperCase(), c.detalle || "-"]);
    const fondo = colorHoja(nivelVisual);

    fila.getCell(1).alignment = { horizontal: "center", vertical: "top" };
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
      const key = clave || `__directo__${c.nombre || c.id}`;
      if (!subgrupos[key]) subgrupos[key] = [];
      subgrupos[key].push(c);
    }

    if (!hayNivelMasProfundo) {
      const ordenado = [...items].sort((a, b) => comparaNatural(a.nombre, b.nombre));
      ordenado.forEach((c) => agregarFila(c, nivelVisual));
      return;
    }

    const entradas = Object.keys(subgrupos).map((key) => ({
      key,
      nombreOrden: key.startsWith("__directo__") ? subgrupos[key][0].nombre || "" : key,
      esGrupo: !key.startsWith("__directo__"),
    }));
    entradas.sort((a, b) => comparaNatural(a.nombreOrden, b.nombreOrden));

    for (const entrada of entradas) {
      if (entrada.esGrupo) {
        agregarSubEncabezado(entrada.key, nivelVisual, subgrupos[entrada.key]);
        agruparRecursivo(subgrupos[entrada.key], nivelIdxRuta + 1, nivelVisual + 1);
      } else {
        const ordenado = [...subgrupos[entrada.key]].sort((a, b) => comparaNatural(a.nombre, b.nombre));
        ordenado.forEach((c) => agregarFila(c, nivelVisual));
      }
    }
  }

  for (const especialidad of ordenGrupos) {
    agregarSubEncabezado(especialidad.toUpperCase(), 1, grupos[especialidad]);
    agruparRecursivo(grupos[especialidad], 2, 1);
  }

  // --- PIE DE PÁGINA VISIBLE AL FINAL DE LA TABLA (Celdas reales con diseño) ---
  sheet.addRow([]); // Espacio de separación

  const filaPie1 = sheet.addRow(["Documento generado automáticamente desde el Sistema de Control de Proyectos - Acocollo I-2"]);
  sheet.mergeCells(`A${filaPie1.number}:D${filaPie1.number}`);
  filaPie1.getCell(1).font = { italic: true, size: 9, color: { argb: "FF555555" }, bold: true };
  filaPie1.getCell(1).alignment = { horizontal: "center", vertical: "middle" };
  filaPie1.getCell(1).fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFEAEDED" } };
  filaPie1.height = 18;

  const filaPie2 = sheet.addRow([`Fecha y hora de emisión: ${new Date().toLocaleDateString("es-PE")} ${new Date().toLocaleTimeString("es-PE")}`]);
  sheet.mergeCells(`A${filaPie2.number}:D${filaPie2.number}`);
  filaPie2.getCell(1).font = { italic: true, size: 8.5, color: { argb: "FF7F8C8D" } };
  filaPie2.getCell(1).alignment = { horizontal: "center", vertical: "middle" };
  filaPie2.getCell(1).fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFEAEDED" } };
  filaPie2.height = 16;

  // Aplicar bordes finos a las celdas del pie para que formen un bloque estético
  [filaPie1, filaPie2].forEach(f => {
    f.eachCell(cell => {
      cell.border = {
        top: { style: "thin", color: { argb: "FFBDC3C7" } },
        bottom: { style: "thin", color: { argb: "FFBDC3C7" } },
        left: { style: "thin", color: { argb: "FFBDC3C7" } },
        right: { style: "thin", color: { argb: "FFBDC3C7" } }
      };
    });
  });

  sheet.views = [{ state: "frozen", ySplit: FILA_ENCABEZADO }];

  // --- PIE DE PÁGINA NATIVO DE EXCEL (Para impresión) ---
  sheet.headerFooter.oddFooter = "&L&I[Sistema Acocollo I-2] Reporte institucional&R&IPágina &P de &N";

  const buffer = await workbook.xlsx.writeBuffer();
  const blob = new Blob([buffer], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = `Reporte_${CODIGO_PROYECTO}_${sufijoFiltroArchivo(tipoFiltro)}_${areaNombre.replace(/[^a-zA-Z0-9]+/g, "_")}_${new Date()
    .toISOString()
    .slice(0, 10)}.xlsx`;
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  URL.revokeObjectURL(url);
}

// --- Palabras que indican una carpeta "contenedor de tipo de archivo" (PDF,
// Editable, CAD, Excel, etc.) — cuando el último nivel de una carpeta es uno
// de estos, la lista de separadores usa el penúltimo nivel en su lugar,
// porque el nombre del contenedor no aporta información real a un separador.
const PALABRAS_CONTENEDOR = [
  "PDF", "EDITABLE", "EDITABLES", "EXCEL", "PLANILLA", "PLANILLAS",
  "CAD", "DWG", "WORD", "DOC", "DOCX", "RVT", "REVIT",
];

function esCarpetaContenedor(nombre) {
  const upper = (nombre || "").toUpperCase();
  return PALABRAS_CONTENEDOR.some((palabra) => upper.includes(palabra));
}

function comparaNaturalSep(a, b) {
  return (a || "").localeCompare(b || "", undefined, { numeric: true, sensitivity: "base" });
}

// Para cada carpeta, calcula los segmentos de ruta "efectivos" para el árbol
// de separadores: se descarta el nombre del área (primer segmento) y, si el
// último nivel es un contenedor de tipo de archivo (PDF/EDITABLE/CAD/etc.),
// se usa el penúltimo nivel en su lugar.
function rutaEfectivaSeparador(carpeta) {
  const partes = (carpeta.ruta || carpeta.nombre || "").split(" / ").filter(Boolean);
  let segmentos = partes.slice(1); // quita el nombre del área
  if (segmentos.length === 0) return [];
  const ultimo = segmentos[segmentos.length - 1];
  if (segmentos.length > 1 && esCarpetaContenedor(ultimo)) {
    segmentos = segmentos.slice(0, -1);
  }
  return segmentos;
}

// Construye un árbol anidado a partir de las rutas efectivas de todas las
// carpetas de un área. Carpetas distintas que colapsan a la misma ruta
// efectiva (ej. "PDF" y "EDITABLE" bajo la misma carpeta madre) se
// deduplican automáticamente al insertarse en el mismo nodo del árbol.
function construirArbolSeparadores(carpetasDelArea) {
  const raiz = {};
  for (const c of carpetasDelArea) {
    const segmentos = rutaEfectivaSeparador(c);
    let nodo = raiz;
    for (const seg of segmentos) {
      if (!nodo[seg]) nodo[seg] = {};
      nodo = nodo[seg];
    }
  }
  return raiz;
}

// Recorre el árbol en el mismo orden en que aparece en Drive (orden
// natural, respetando números), asignando numeración correlativa jerárquica
// (1, 1.1, 1.2, 1.3, 1.3.1, ... 2, 2.1, ...) a cada nodo, sea rama o final.
function numerarArbolSeparadores(nodo, prefijo, resultado) {
  const claves = Object.keys(nodo).sort(comparaNaturalSep);
  let contador = 1;
  for (const clave of claves) {
    const numero = prefijo ? `${prefijo}.${contador}` : `${contador}`;
    resultado.push({ numero, nombre: clave });
    numerarArbolSeparadores(nodo[clave], numero, resultado);
    contador++;
  }
}

const DORADO_FASE2 = "FFB8860B"; // acento distinto (dorado), exclusivo de este documento "fase 2"

/**
 * Genera y descarga un Excel simple de dos columnas (N° / Descripción),
 * pensado para imprimir los separadores físicos de un expediente técnico —
 * NO es un reporte de avance, es un índice numerado plano de las secciones
 * reales del área, en el mismo orden en que están en Drive.
 *
 * @param {string} areaNombre - nombre del área (carpeta madre) a exportar
 * @param {Array} carpetasDelArea - carpetas (ya filtradas a esa área)
 */
export async function generarListaSeparadoresExcel(areaNombre, carpetasDelArea) {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = "Visor ACASO I-2";
  workbook.created = new Date();

  const nombreHoja = areaNombre.replace(/[\\/*?:[\]]/g, "").slice(0, 31) || "Separadores";
  const sheet = workbook.addWorksheet(nombreHoja, {
    pageSetup: { orientation: "portrait", fitToPage: true, fitToWidth: 1 },
  });

  sheet.columns = [
    { width: 14 },
    { width: 75 },
  ];

  if (LOGO_PUNO_BASE64) {
    try {
      const base64Data = LOGO_PUNO_BASE64.includes("base64,")
        ? LOGO_PUNO_BASE64.split("base64,")[1]
        : LOGO_PUNO_BASE64;
      const imageId = workbook.addImage({ base64: base64Data, extension: "png" });
      sheet.addImage(imageId, { tl: { col: 0.2, row: 0.2 }, ext: { width: 45, height: 50 } });
    } catch (e) {
      console.error("No se pudo cargar el logo en el Excel:", e);
    }
  }

  sheet.mergeCells("A1:B1");
  sheet.getCell("A1").value = "GOBIERNO REGIONAL DE PUNO — GERENCIA REGIONAL DE INFRAESTRUCTURA";
  sheet.getCell("A1").font = { bold: true, size: 11 };
  sheet.getCell("A1").alignment = { horizontal: "center", vertical: "middle" };
  sheet.getRow(1).height = 18;

  sheet.mergeCells("A2:B2");
  sheet.getCell("A2").value = "SUB GERENCIA DE ESTUDIOS DEFINITIVOS";
  sheet.getCell("A2").font = { bold: true, size: 10 };
  sheet.getCell("A2").alignment = { horizontal: "center", vertical: "middle" };
  sheet.getRow(2).height = 16;

  sheet.mergeCells("A3:B3");
  const celdaTitulo = sheet.getCell("A3");
  celdaTitulo.value = `ÍNDICE DE SEPARADORES — ${areaNombre.toUpperCase()}`;
  celdaTitulo.font = { bold: true, size: 12, color: { argb: "FFFFFFFF" } };
  celdaTitulo.alignment = { horizontal: "center", vertical: "middle" };
  celdaTitulo.fill = { type: "pattern", pattern: "solid", fgColor: { argb: DORADO_FASE2 } };
  sheet.getRow(3).height = 22;

  sheet.mergeCells("A4:B4");
  sheet.getCell("A4").value = `Generado: ${new Date().toLocaleString("es-PE")}  ·  Lista plana para impresión de separadores físicos — no refleja estado de avance`;
  sheet.getCell("A4").font = { italic: true, size: 8.5, color: { argb: "FF666666" } };
  sheet.getCell("A4").alignment = { horizontal: "center" };
  sheet.getRow(4).height = 14;

  sheet.addRow([]); // fila 5 en blanco

  const FILA_ENCABEZADO = 6;
  const filaEncabezado = sheet.getRow(FILA_ENCABEZADO);
  filaEncabezado.values = ["N°", "DESCRIPCIÓN"];
  filaEncabezado.eachCell((cell) => {
    cell.font = { bold: true, color: { argb: "FFFFFFFF" } };
    cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: DORADO_FASE2 } };
    cell.alignment = { vertical: "middle", horizontal: "center" };
    cell.border = { bottom: { style: "medium", color: { argb: "FF000000" } } };
  });
  filaEncabezado.height = 18;

  // --- Construir y recorrer el árbol jerárquico ---
  const arbol = construirArbolSeparadores(carpetasDelArea);
  const items = [];
  numerarArbolSeparadores(arbol, "", items);

  for (const item of items) {
    const fila = sheet.addRow([item.numero, item.nombre]);
    fila.getCell(1).alignment = { horizontal: "left", vertical: "middle" };
    fila.getCell(1).numFmt = "@"; // forzar texto, para que "1.10" no se lea como número
    fila.getCell(2).alignment = { horizontal: "left", vertical: "middle", wrapText: true };
    fila.eachCell((cell) => {
      cell.border = {
        top: { style: "hair", color: { argb: "FFDDDDDD" } },
        bottom: { style: "hair", color: { argb: "FFDDDDDD" } },
        left: { style: "hair", color: { argb: "FFDDDDDD" } },
        right: { style: "hair", color: { argb: "FFDDDDDD" } },
      };
    });
  }

  sheet.views = [{ state: "frozen", ySplit: FILA_ENCABEZADO }];
  sheet.headerFooter.oddFooter = "&L&I[Sistema ACASO I-2] Índice de separadores&R&IPágina &P de &N";

  const buffer = await workbook.xlsx.writeBuffer();
  const blob = new Blob([buffer], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = `Separadores_${CODIGO_PROYECTO}_${areaNombre.replace(/[^a-zA-Z0-9]+/g, "_")}_${new Date()
    .toISOString()
    .slice(0, 10)}.xlsx`;
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  URL.revokeObjectURL(url);
}

