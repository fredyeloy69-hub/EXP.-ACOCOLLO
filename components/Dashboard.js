"use client";

import { useEffect, useState, useRef } from "react";
import { db } from "@/lib/firebaseClient";
import {
  collection,
  doc,
  onSnapshot,
  orderBy,
  query,
  limit,
  where,
  getDocs,
} from "firebase/firestore";
import { generarReportePorArea, generarReporteConsolidadoGlobal } from "@/lib/exportarReporte";
import { generarReporteExcelPorArea } from "@/lib/exportarExcel";
import { LOGO_PUNO_BASE64 } from "@/lib/logoPuno";
import {
  getAuth,
  GoogleAuthProvider,
  signInWithPopup,
  onAuthStateChanged,
} from "firebase/auth";

const auth = getAuth(db.app);

const COLLAPSE_STORAGE_KEY = "acocollo_i2_grupos_colapsados";

const ESTADO_COLOR = {
  completa: "#f39c12",
  incompleta: "#d35400",
  vacia: "#c0392b",
};

const AREA_COLORS = [
  "#2a9d8f",
  "#264653",
  "#f4a261",
  "#e76f51",
  "#457b9d",
  "#1d3557",
  "#f1c40f",
  "#2a9d8f",
];

function colorForArea(area) {
  let hash = 0;
  for (let i = 0; i < area.length; i++) hash = area.charCodeAt(i) + ((hash << 5) - hash);
  return AREA_COLORS[Math.abs(hash) % AREA_COLORS.length];
}

const ESTADO_OPTIONS = [
  { value: "pendientes", label: "Pendientes", color: "#f39c12" },
  { value: "incompleta", label: "Incompletas", color: "#d35400" },
  { value: "vacia", label: "Vacías", color: "#c0392b" },
  { value: "completa", label: "Completas", color: "#f1c40f" },
  { value: "todas", label: "Todas", color: "#2a9d8f" },
];

const EVENTO_LABEL = {
  archivo_subido: "subió",
  archivo_reemplazado: "reemplazó",
  archivo_borrado: "borró",
  carpeta_creada: "creó la carpeta",
  carpeta_borrada: "borró la carpeta",
  carpeta_movida: "movió la carpeta",
  carpeta_marcada_completa: "marcó como completa",
  carpeta_desmarcada: "desmarcó",
};

const EVENTO_COLOR = {
  archivo_subido: "#2a9d8f",
  archivo_reemplazado: "#f39c12",
  archivo_borrado: "#c0392b",
  carpeta_creada: "#264653",
  carpeta_borrada: "#c0392b",
  carpeta_movida: "#d35400",
  carpeta_marcada_completa: "#2a9d8f",
  carpeta_desmarcada: "#f39c12",
};

const EVENTO_ICONO = {
  archivo_subido: "↑",
  archivo_reemplazado: "⟲",
  archivo_borrado: "✕",
  carpeta_creada: "+",
  carpeta_borrada: "✕",
  carpeta_movida: "⇄",
  carpeta_marcada_completa: "✓",
  carpeta_desmarcada: "↺",
};

function fechaLimaISO(fecha) {
  const partes = new Intl.DateTimeFormat("en-CA", {
    timeZone: "America/Lima",
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
  }).formatToParts(fecha);
  const obj = {};
  for (const p of partes) obj[p.type] = p.value;
  return `${obj.year}-${obj.month}-${obj.day}`;
}

function formatearFechaLarga(fechaEntrada) {
  const fecha = typeof fechaEntrada === "string" ? new Date(fechaEntrada + "T12:00:00") : fechaEntrada;
  if (!fecha || isNaN(fecha.getTime())) return String(fechaEntrada);
  const texto = fecha.toLocaleDateString("es-PE", { weekday: "long", day: "numeric", month: "long", year: "numeric" });
  return texto.charAt(0).toUpperCase() + texto.slice(1);
}

function tiempoRelativo(date) {
  if (!date) return "...";
  const diffMs = Date.now() - date.getTime();
  const diffMin = Math.floor(diffMs / 60000);
  if (diffMin < 1) return "justo ahora";
  if (diffMin < 60) return `hace ${diffMin} min`;
  const diffH = Math.floor(diffMin / 60);
  if (diffH < 24) return `hace ${diffH} h`;
  const diffD = Math.floor(diffH / 24);
  if (diffD < 7) return `hace ${diffD} d`;
  return date.toLocaleDateString("es-PE");
}

export default function Dashboard() {
  const [resumen, setResumen] = useState(null);
  const [carpetas, setCarpetas] = useState([]);
  const [eventos, setEventos] = useState([]);
  const [filtroArea, setFiltroArea] = useState("Todas");
  const [filtroEstado, setFiltroEstado] = useState("pendientes");
  const [sincronizando, setSincronizando] = useState(false);
  const [mensajeSync, setMensajeSync] = useState(null);
  const [busqueda, setBusqueda] = useState("");
  const [colapsados, setColapsados] = useState({});
  const [colapsoListo, setColapsoListo] = useState(false);
  const [exportandoArea, setExportandoArea] = useState(null);
  const [exportandoExcelArea, setExportandoExcelArea] = useState(null);
  const [exportandoGlobal, setExportandoGlobal] = useState(false);
  const [modoPresentacion, setModoPresentacion] = useState(false);
  const [historial, setHistorial] = useState([]);
  const [actividadPorDia, setActividadPorDia] = useState({});
  const [marcandoId, setMarcandoId] = useState(null);
  const [mostrarMarcadas, setMostrarMarcadas] = useState(false);
  const [usuarioGoogle, setUsuarioGoogle] = useState(null);
  const [rangoDiasHeatmap, setRangoDiasHeatmap] = useState(84);

  useEffect(() => {
    const unsub = onAuthStateChanged(auth, (user) => {
      setUsuarioGoogle(user ? { email: user.email, displayName: user.displayName } : null);
    });
    return () => unsub();
  }, []);

  useEffect(() => {
    try {
      const guardado = localStorage.getItem(COLLAPSE_STORAGE_KEY);
      if (guardado) setColapsados(JSON.parse(guardado));
    } catch {}
    setColapsoListo(true);
  }, []);

  useEffect(() => {
    if (!colapsoListo) return;
    try {
      localStorage.setItem(COLLAPSE_STORAGE_KEY, JSON.stringify(colapsados));
    } catch {}
  }, [colapsados, colapsoListo]);

  function toggleGrupo(key) {
    setColapsados((prev) => ({ ...prev, [key]: !prev[key] }));
  }

  function handleExportarArea(areaNombre, carpetasDelArea) {
    setExportandoArea(areaNombre);
    try {
      const usuarioFirma = usuarioGoogle?.email || usuarioGoogle?.displayName || "Sistema Acocollo I-2";
      generarReportePorArea(areaNombre, carpetasDelArea, { usuarioFirma });
    } finally {
      setExportandoArea(null);
    }
  }

  async function handleExportarGlobal() {
    setExportandoGlobal(true);
    try {
      const usuarioFirma = usuarioGoogle?.email || usuarioGoogle?.displayName || "Sistema Acocollo I-2";
      generarReporteConsolidadoGlobal(carpetas, { usuarioFirma });
    } catch (err) {
      alert(`No se pudo generar el reporte consolidado: ${err.message}`);
    } finally {
      setExportandoGlobal(false);
    }
  }

  async function handleExportarExcelArea(areaNombre, carpetasDelArea) {
    setExportandoExcelArea(areaNombre);
    try {
      const timeoutPromise = new Promise((_, reject) =>
        setTimeout(() => reject(new Error("Tiempo de espera agotado al generar el Excel")), 10000)
      );
      await Promise.race([
        generarReporteExcelPorArea(areaNombre, carpetasDelArea),
        timeoutPromise,
      ]);
    } catch (err) {
      alert(`No se pudo generar el Excel: ${err.message}`);
    } finally {
      setExportandoExcelArea(null);
    }
  }

  async function handleMarcarCompleta(folderId, forzada, folderName, folderRuta) {
    let user = auth.currentUser;
    if (!user) {
      try {
        const cred = await signInWithPopup(auth, new GoogleAuthProvider());
        user
