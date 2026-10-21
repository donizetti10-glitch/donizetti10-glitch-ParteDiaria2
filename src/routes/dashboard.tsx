import { createFileRoute, Link } from "@tanstack/react-router";
import { useEffect, useMemo, useRef, useState } from "react";
import {
  emptyHeader,
  formatDate,
  formatDuration,
  formatDurationHM,
  formatTime,
  totals,
  type DailyHeader,
  type WorkEvent,
  sortAndNormalizeEvents,
} from "@/lib/parte-diaria";
import {
  subscribeCloudReports,
  type SavedCloudReport,
  cloudReportToHeaderAndEvents,
} from "@/lib/cloud-reports";
import {
  Cloud,
  Calendar,
  User,
  ArrowLeft,
  Search,
  Filter,
  RotateCcw,
  TrendingUp,
  AlertTriangle,
  Clock,
  MapPin,
  Activity,
  FileSpreadsheet,
  BarChart3,
  Layers,
  ChevronRight,
  Eye,
  CheckCircle2,
  ClipboardCheck,
  Target,
} from "lucide-react";
import { exportAllReportsToExcel, exportSingleReportToExcel } from "@/lib/excel-export";
import { ThemeToggle } from "@/components/ThemeToggle";
import { GdaAnalyticsDashboard } from "@/components/GdaAnalyticsDashboard";
import { classifyGdaStatus, type GdaStatusKey, GDA_STATUS_CONFIG } from "@/lib/gda-analytics";

export const Route = createFileRoute("/dashboard")({
  component: DashboardPage,
  head: () => ({
    meta: [
      { title: "Dashboard & Panorama Geral • Inspeções na Nuvem" },
      {
        name: "description",
        content: "Panorama geral com filtros interativos de todos os relatórios de inspeção.",
      },
    ],
  }),
});

const STORAGE_KEY = "parte-diaria-v1";

type State = { header: DailyHeader; events: WorkEvent[]; roster: string[] };

export function DashboardPage() {
  const [state, setState] = useState<State>({ header: emptyHeader(), events: [], roster: [] });
  const [cloudReports, setCloudReports] = useState<SavedCloudReport[]>([]);
  const [viewMode, setViewMode] = useState<"panorama" | "single">("panorama");
  const [selectedReportId, setSelectedReportId] = useState<string>("all_cloud");

  // Filter states
  const [searchTerm, setSearchTerm] = useState("");
  const [periodFilter, setPeriodFilter] = useState<
    "todos" | "hoje" | "7dias" | "30dias" | "mes_atual" | "custom"
  >("todos");
  const [startDate, setStartDate] = useState("");
  const [endDate, setEndDate] = useState("");
  const [selectedLocal, setSelectedLocal] = useState("todos");
  const [selectedAtividade, setSelectedAtividade] = useState("todos");
  const [selectedResponsavel, setSelectedResponsavel] = useState("todos");
  const [selectedGdaStatus, setSelectedGdaStatus] = useState<GdaStatusKey | "todos">("todos");

  const canvasRef = useRef<HTMLCanvasElement | null>(null);
  const panoramaCanvasRef = useRef<HTMLCanvasElement | null>(null);

  // Subscribe to cloud reports in real time
  useEffect(() => {
    const unsub = subscribeCloudReports((reports) => {
      setCloudReports(reports);
    });
    return unsub;
  }, []);

  // Load local storage draft
  useEffect(() => {
    try {
      const raw = localStorage.getItem(STORAGE_KEY);
      if (raw) {
        const parsed = JSON.parse(raw);
        setState({
          header: { ...emptyHeader(), ...(parsed.header ?? {}) },
          events: parsed.events ?? [],
          roster: parsed.roster ?? [],
        });
      }
    } catch (e) {
      console.debug("Error loading local draft:", e);
    }
  }, []);

  // Unique options for dynamic dropdown filters
  const filterOptions = useMemo(() => {
    const locais = new Set<string>();
    const atividades = new Set<string>();
    const responsaveis = new Set<string>();

    cloudReports.forEach((r) => {
      if (r.local?.trim()) locais.add(r.local.trim());
      if (r.atividade?.trim()) atividades.add(r.atividade.trim());
      if (r.nome?.trim()) responsaveis.add(r.nome.trim());
    });

    return {
      locais: Array.from(locais).sort(),
      atividades: Array.from(atividades).sort(),
      responsaveis: Array.from(responsaveis).sort(),
    };
  }, [cloudReports]);

  // Filtered cloud reports based on user selections
  const filteredReports = useMemo(() => {
    return cloudReports.filter((r) => {
      // Search term
      if (searchTerm.trim()) {
        const term = searchTerm.toLowerCase();
        const matchesSearch =
          (r.atividade || "").toLowerCase().includes(term) ||
          (r.nome || "").toLowerCase().includes(term) ||
          (r.local || "").toLowerCase().includes(term) ||
          (r.om || "").toLowerCase().includes(term) ||
          (r.via || "").toLowerCase().includes(term) ||
          (r.descritivo || "").toLowerCase().includes(term) ||
          (r.observacaoGeral || "").toLowerCase().includes(term) ||
          (r.id || "").toLowerCase().includes(term);
        if (!matchesSearch) return false;
      }

      // Local filter
      if (selectedLocal !== "todos" && r.local !== selectedLocal) {
        return false;
      }

      // Atividade filter
      if (selectedAtividade !== "todos" && r.atividade !== selectedAtividade) {
        return false;
      }

      // Responsavel filter
      if (selectedResponsavel !== "todos" && r.nome !== selectedResponsavel) {
        return false;
      }

      // GDA Status filter
      if (selectedGdaStatus !== "todos") {
        const classified = classifyGdaStatus(r);
        if (classified.status !== selectedGdaStatus) {
          return false;
        }
      }

      // Date / Period filter
      if (r.data) {
        const reportDateStr = r.data; // YYYY-MM-DD
        const todayStr = new Date().toISOString().slice(0, 10);

        if (periodFilter === "hoje") {
          if (reportDateStr !== todayStr) return false;
        } else if (periodFilter === "7dias") {
          const sevenDaysAgo = new Date();
          sevenDaysAgo.setDate(sevenDaysAgo.getDate() - 7);
          const sevenDaysStr = sevenDaysAgo.toISOString().slice(0, 10);
          if (reportDateStr < sevenDaysStr) return false;
        } else if (periodFilter === "30dias") {
          const thirtyDaysAgo = new Date();
          thirtyDaysAgo.setDate(thirtyDaysAgo.getDate() - 30);
          const thirtyDaysStr = thirtyDaysAgo.toISOString().slice(0, 10);
          if (reportDateStr < thirtyDaysStr) return false;
        } else if (periodFilter === "mes_atual") {
          const currentYearMonth = new Date().toISOString().slice(0, 7); // YYYY-MM
          if (!reportDateStr.startsWith(currentYearMonth)) return false;
        } else if (periodFilter === "custom") {
          if (startDate && reportDateStr < startDate) return false;
          if (endDate && reportDateStr > endDate) return false;
        }
      }

      return true;
    });
  }, [
    cloudReports,
    searchTerm,
    periodFilter,
    startDate,
    endDate,
    selectedLocal,
    selectedAtividade,
    selectedResponsavel,
    selectedGdaStatus,
  ]);

  // Count active filters
  const activeFiltersCount = useMemo(() => {
    let count = 0;
    if (searchTerm.trim()) count++;
    if (periodFilter !== "todos") count++;
    if (selectedLocal !== "todos") count++;
    if (selectedAtividade !== "todos") count++;
    if (selectedResponsavel !== "todos") count++;
    if (selectedGdaStatus !== "todos") count++;
    return count;
  }, [
    searchTerm,
    periodFilter,
    selectedLocal,
    selectedAtividade,
    selectedResponsavel,
    selectedGdaStatus,
  ]);

  const clearAllFilters = () => {
    setSearchTerm("");
    setPeriodFilter("todos");
    setStartDate("");
    setEndDate("");
    setSelectedLocal("todos");
    setSelectedAtividade("todos");
    setSelectedResponsavel("todos");
    setSelectedGdaStatus("todos");
  };

  // Panorama KPI metrics across filtered reports
  const panoramaMetrics = useMemo(() => {
    let totalProducao = 0;
    let totalDefeitos = 0;
    let totalAtividadeMs = 0;
    let totalParadaMs = 0;
    let totalJornadaMs = 0;

    const stopReasonsMap: Record<string, { count: number; durMs: number }> = {};
    const activityMap: Record<string, { count: number; producao: number }> = {};
    const localMap: Record<string, { count: number; producao: number }> = {};

    filteredReports.forEach((r) => {
      const prod = r.totals?.producaoTotal ?? 0;
      const def = r.totals?.defeitosTotal ?? 0;
      const ativ = r.totals?.atividade ?? 0;
      const par = r.totals?.parada ?? 0;
      const jor = r.totals?.jornada ?? 0;

      totalProducao += prod;
      totalDefeitos += def;
      totalAtividadeMs += ativ;
      totalParadaMs += par;
      totalJornadaMs += jor;

      // Group by activity
      const ativName = r.atividade?.trim() || "Não Informada";
      if (!activityMap[ativName]) activityMap[ativName] = { count: 0, producao: 0 };
      activityMap[ativName].count += 1;
      activityMap[ativName].producao += prod;

      // Group by location
      const locName = r.local?.trim() || "Não Informado";
      if (!localMap[locName]) localMap[locName] = { count: 0, producao: 0 };
      localMap[locName].count += 1;
      localMap[locName].producao += prod;

      // Group stops
      const sortedEvs = sortAndNormalizeEvents(r.events || []);
      for (let i = 0; i < sortedEvs.length; i++) {
        if (sortedEvs[i].kind === "parada") {
          const next = sortedEvs[i + 1];
          const durMs = next ? Math.max(0, next.timestamp - sortedEvs[i].timestamp) : 0;
          const reason = sortedEvs[i].descricao?.trim() || "Parada Sem Motivo";
          if (!stopReasonsMap[reason]) stopReasonsMap[reason] = { count: 0, durMs: 0 };
          stopReasonsMap[reason].count += 1;
          stopReasonsMap[reason].durMs += durMs;
        }
      }
    });

    const totalReportsCount = filteredReports.length;
    const avgProducao = totalReportsCount > 0 ? totalProducao / totalReportsCount : 0;
    const efficiency =
      totalJornadaMs > 0 ? Math.min(100, Math.round((totalAtividadeMs / totalJornadaMs) * 100)) : 0;

    const topStops = Object.entries(stopReasonsMap)
      .map(([reason, val]) => ({ reason, count: val.count, durMs: val.durMs }))
      .sort((a, b) => b.durMs - a.durMs)
      .slice(0, 5);

    const topActivities = Object.entries(activityMap)
      .map(([name, val]) => ({ name, count: val.count, producao: val.producao }))
      .sort((a, b) => b.producao - a.producao);

    const topLocais = Object.entries(localMap)
      .map(([name, val]) => ({ name, count: val.count, producao: val.producao }))
      .sort((a, b) => b.producao - a.producao);

    return {
      totalReportsCount,
      totalProducao,
      totalDefeitos,
      totalAtividadeMs,
      totalParadaMs,
      totalJornadaMs,
      avgProducao,
      efficiency,
      topStops,
      topActivities,
      topLocais,
    };
  }, [filteredReports]);

  // Draw Panorama Canvas Chart (Chronological Production & Defects across filtered reports)
  useEffect(() => {
    if (viewMode !== "panorama") return;
    const canvas = panoramaCanvasRef.current;
    if (!canvas) return;

    const dpr = window.devicePixelRatio || 1;
    const W = 920;
    const H = 340;
    canvas.width = W * dpr;
    canvas.height = H * dpr;
    canvas.style.width = "100%";
    canvas.style.maxWidth = W + "px";
    canvas.style.height = "auto";
    const ctx = canvas.getContext("2d");
    if (!ctx) return;
    ctx.scale(dpr, dpr);

    // Background
    ctx.fillStyle = "#ffffff";
    ctx.fillRect(0, 0, W, H);

    if (filteredReports.length === 0) {
      ctx.fillStyle = "#64748b";
      ctx.font = "14px system-ui, sans-serif";
      ctx.textAlign = "center";
      ctx.fillText("Nenhum relatório encontrado com os filtros selecionados.", W / 2, H / 2);
      return;
    }

    // Sort reports chronologically
    const sortedRep = [...filteredReports].sort((a, b) =>
      (a.data || "").localeCompare(b.data || ""),
    );

    const padL = 60,
      padR = 20,
      padT = 40,
      padB = 50;
    const innerW = W - padL - padR;
    const innerH = H - padT - padB;

    // Header Title
    ctx.fillStyle = "#0f172a";
    ctx.font = "bold 15px system-ui, sans-serif";
    ctx.textAlign = "left";
    ctx.fillText("Produção (m) e Defeitos por Relatório de Inspeção", padL, 24);

    const maxProd = Math.max(...sortedRep.map((r) => r.totals?.producaoTotal || 0), 10);
    const maxDef = Math.max(...sortedRep.map((r) => r.totals?.defeitosTotal || 0), 5);

    // Y Axis grid lines (4 steps)
    ctx.strokeStyle = "#e2e8f0";
    ctx.fillStyle = "#64748b";
    ctx.font = "10px system-ui, sans-serif";
    ctx.lineWidth = 1;

    for (let i = 0; i <= 4; i++) {
      const gy = padT + innerH - (i / 4) * innerH;
      ctx.beginPath();
      ctx.moveTo(padL, gy);
      ctx.lineTo(padL + innerW, gy);
      ctx.stroke();

      const val = Math.round((i / 4) * maxProd);
      ctx.textAlign = "right";
      ctx.fillText(`${val}m`, padL - 8, gy + 3);
    }

    // Draw Bars for each report
    const n = sortedRep.length;
    const barGroupW = innerW / n;
    const barWidth = Math.min(28, Math.max(8, barGroupW * 0.45));

    sortedRep.forEach((r, idx) => {
      const groupCenterX = padL + idx * barGroupW + barGroupW / 2;
      const prodVal = r.totals?.producaoTotal || 0;
      const defVal = r.totals?.defeitosTotal || 0;

      const prodH = (prodVal / maxProd) * innerH;
      const defH = (defVal / maxDef) * (innerH * 0.6); // scale defect indicator

      // Production bar (Green)
      const bx = groupCenterX - barWidth / 2;
      const by = padT + innerH - prodH;
      ctx.fillStyle = "#16a34a";
      ctx.fillRect(bx, by, barWidth, prodH);

      // Value on top of bar
      if (prodVal > 0) {
        ctx.fillStyle = "#15803d";
        ctx.font = "bold 10px system-ui, sans-serif";
        ctx.textAlign = "center";
        ctx.fillText(`${prodVal}m`, groupCenterX, by - 4);
      }

      // Defect badge dot (Red) if defects exist
      if (defVal > 0) {
        ctx.fillStyle = "#dc2626";
        ctx.beginPath();
        ctx.arc(groupCenterX, padT + innerH - defH, 5, 0, Math.PI * 2);
        ctx.fill();
        ctx.fillStyle = "#ffffff";
        ctx.font = "bold 8px system-ui, sans-serif";
        ctx.fillText(String(defVal), groupCenterX, padT + innerH - defH + 3);
      }

      // Date Label below
      ctx.fillStyle = "#334155";
      ctx.font = "10px system-ui, sans-serif";
      ctx.textAlign = "center";
      const dateLabel = formatDate(r.data).slice(0, 5); // DD/MM
      ctx.fillText(dateLabel, groupCenterX, padT + innerH + 16);

      // Local / Activity shortened label
      ctx.fillStyle = "#64748b";
      ctx.font = "9px system-ui, sans-serif";
      const subLabel = (r.local || r.atividade || "").slice(0, 8);
      ctx.fillText(subLabel, groupCenterX, padT + innerH + 28);
    });

    // Legend
    ctx.textAlign = "right";
    ctx.fillStyle = "#16a34a";
    ctx.fillRect(W - 200, 14, 12, 12);
    ctx.fillStyle = "#334155";
    ctx.font = "11px system-ui, sans-serif";
    ctx.fillText("Produção (m)", W - 120, 24);

    ctx.fillStyle = "#dc2626";
    ctx.beginPath();
    ctx.arc(W - 100, 20, 5, 0, Math.PI * 2);
    ctx.fill();
    ctx.fillStyle = "#334155";
    ctx.fillText("Defeitos", W - 40, 24);
  }, [filteredReports, viewMode]);

  // Active dataset depending on selection when in single mode
  const activeReportData = useMemo(() => {
    if (selectedReportId === "local_draft") {
      return {
        title: "Relatório Atual (Rascunho Local)",
        header: state.header,
        events: state.events,
      };
    }

    if (selectedReportId === "all_cloud") {
      const allEvents: WorkEvent[] = [];
      filteredReports.forEach((r) => {
        if (r.events) allEvents.push(...r.events);
      });
      return {
        title: `Consolidado (${filteredReports.length} relatórios filtrados)`,
        header: {
          ...emptyHeader(),
          atividade: `Consolidado Nuvem (${filteredReports.length} relatórios)`,
          nome: "Todos os Responsáveis",
        },
        events: allEvents,
      };
    }

    const found = cloudReports.find((r) => r.id === selectedReportId);
    if (found) {
      const { header, events } = cloudReportToHeaderAndEvents(found);
      return {
        title: `Relatório: ${found.atividade || "Inspeção"} (${formatDate(found.data)}) - ${found.nome || "Sem resp."}`,
        header,
        events,
      };
    }

    return {
      title: "Relatório Selecionado",
      header: state.header,
      events: state.events,
    };
  }, [selectedReportId, state, cloudReports, filteredReports]);

  const currentHeader = activeReportData.header || emptyHeader();
  const currentEvents = useMemo(() => activeReportData.events || [], [activeReportData.events]);

  const sortedSingle = useMemo(() => sortAndNormalizeEvents(currentEvents), [currentEvents]);
  const tSingle = useMemo(() => totals(currentEvents), [currentEvents]);

  // Series for Single mode line charts
  const singleSeries = useMemo(() => {
    let prodAcc = 0;
    let defAcc = 0;
    const points: { ts: number; prod: number; def: number; prodInc: number; defInc: number }[] = [];
    for (const e of sortedSingle) {
      if (e.kind === "fim_atividade") {
        const prodInc = Number(e.producao) || 0;
        const defInc = Number(e.defeitos) || 0;
        prodAcc += prodInc;
        defAcc += defInc;
        points.push({ ts: e.timestamp, prod: prodAcc, def: defAcc, prodInc, defInc });
      }
    }
    return points;
  }, [sortedSingle]);

  // Single mode paradas
  const singleParadas = useMemo(() => {
    const items: { ts: number; durMin: number; motivo?: string }[] = [];
    for (let i = 0; i < sortedSingle.length; i++) {
      if (sortedSingle[i].kind === "parada") {
        const next = sortedSingle[i + 1];
        if (next) {
          items.push({
            ts: sortedSingle[i].timestamp,
            durMin: (next.timestamp - sortedSingle[i].timestamp) / 60000,
            motivo: sortedSingle[i].descricao,
          });
        }
      }
    }
    return items;
  }, [sortedSingle]);

  const singleTempoDist = useMemo(() => {
    const outros = Math.max(
      0,
      tSingle.jornada - tSingle.atividade - tSingle.preparacao - tSingle.parada,
    );
    return [
      { label: "Atividade", ms: tSingle.atividade, color: "#16a34a" },
      { label: "Preparação", ms: tSingle.preparacao, color: "#3b82f6" },
      { label: "Paradas", ms: tSingle.parada, color: "#f59e0b" },
      { label: "Outros", ms: outros, color: "#94a3b8" },
    ];
  }, [tSingle]);

  // Draw Single Mode Canvas Chart
  useEffect(() => {
    if (viewMode !== "single") return;
    const canvas = canvasRef.current;
    if (!canvas) return;
    const dpr = window.devicePixelRatio || 1;
    const W = 900;
    const H = 720;
    canvas.width = W * dpr;
    canvas.height = H * dpr;
    canvas.style.width = "100%";
    canvas.style.maxWidth = W + "px";
    canvas.style.height = "auto";
    const ctx = canvas.getContext("2d");
    if (!ctx) return;
    ctx.scale(dpr, dpr);

    // Canvas Background
    ctx.fillStyle = "#ffffff";
    ctx.fillRect(0, 0, W, H);

    ctx.fillStyle = "#0f172a";
    ctx.font = "bold 20px system-ui, sans-serif";
    ctx.fillText("Parte Diária — Produção e Perdas", 24, 32);
    ctx.font = "13px system-ui, sans-serif";
    ctx.fillStyle = "#475569";
    const dateStr = formatDate(currentHeader.data) || "";
    ctx.fillText(
      `${currentHeader.atividade || activeReportData.title}${dateStr ? " • " + dateStr : ""}${
        currentHeader.nome ? " • Resp: " + currentHeader.nome : ""
      }`,
      24,
      52,
    );

    drawLineChart(ctx, {
      x: 24,
      y: 80,
      w: 560,
      h: 240,
      title: "Produção acumulada (m)",
      color: "#16a34a",
      points: singleSeries.map((p) => ({ x: p.ts, y: p.prod })),
      xLabel: (v) => formatTime(v).slice(0, 5),
      yLabel: (v) => Math.round(v).toString(),
    });

    drawLineChart(ctx, {
      x: 24,
      y: 340,
      w: 560,
      h: 200,
      title: "Defeitos acumulados",
      color: "#dc2626",
      points: singleSeries.map((p) => ({ x: p.ts, y: p.def })),
      xLabel: (v) => formatTime(v).slice(0, 5),
      yLabel: (v) => Math.round(v).toString(),
    });

    drawBarChart(ctx, {
      x: 24,
      y: 560,
      w: 560,
      h: 140,
      title: "Duração das paradas (min)",
      color: "#f59e0b",
      bars: singleParadas.map((p) => ({ label: formatTime(p.ts).slice(0, 5), value: p.durMin })),
    });

    drawDonut(ctx, {
      cx: 740,
      cy: 220,
      r: 90,
      inner: 55,
      title: "Distribuição do tempo",
      slices: singleTempoDist,
    });

    const boxX = 610,
      boxY = 340,
      boxW = 265,
      boxH = 360;
    ctx.strokeStyle = "#e2e8f0";
    ctx.lineWidth = 1;
    ctx.strokeRect(boxX, boxY, boxW, boxH);
    ctx.fillStyle = "#0f172a";
    ctx.font = "bold 14px system-ui, sans-serif";
    ctx.fillText("Resumo da Inspeção", boxX + 12, boxY + 24);
    ctx.font = "13px system-ui, sans-serif";
    const rows: [string, string][] = [
      ["Jornada", formatDuration(tSingle.jornada)],
      ["Atividade", formatDuration(tSingle.atividade)],
      ["Preparação", formatDuration(tSingle.preparacao)],
      ["Paradas", formatDuration(tSingle.parada)],
      ["Produção (m)", String(tSingle.producaoTotal || 0)],
      ["Defeitos", String(tSingle.defeitosTotal || 0)],
      ["Nº paradas", String(singleParadas.length)],
      ["Nº atividades", String(singleSeries.length)],
    ];
    rows.forEach(([k, v], i) => {
      const yy = boxY + 52 + i * 24;
      ctx.fillStyle = "#64748b";
      ctx.fillText(k, boxX + 12, yy);
      ctx.fillStyle = "#0f172a";
      ctx.font = "bold 13px system-ui, sans-serif";
      const tw = ctx.measureText(v).width;
      ctx.fillText(v, boxX + boxW - 12 - tw, yy);
      ctx.font = "13px system-ui, sans-serif";
    });
  }, [
    singleSeries,
    singleParadas,
    singleTempoDist,
    tSingle,
    currentHeader,
    activeReportData,
    viewMode,
  ]);

  const shareWhatsApp = async () => {
    const canvas = canvasRef.current;
    if (!canvas) return;
    const blob: Blob | null = await new Promise((resolve) =>
      canvas.toBlob((b) => resolve(b), "image/png"),
    );
    if (!blob) return;
    const filename = `dashboard-${currentHeader.data || "inspecao"}.png`;
    const file = new File([blob], filename, { type: "image/png" });
    const text = `Dashboard Registro de Atividade Diária — ${formatDate(currentHeader.data)}\nProdução: ${tSingle.producaoTotal || 0} m • Defeitos: ${tSingle.defeitosTotal || 0} • Paradas: ${formatDurationHM(tSingle.parada)}`;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const nav = navigator as any;
    if (nav.canShare && nav.canShare({ files: [file] })) {
      try {
        await nav.share({ files: [file], title: "Dashboard", text });
        return;
      } catch (e) {
        console.debug("Share cancelled or failed:", e);
      }
    }
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = filename;
    a.click();
    URL.revokeObjectURL(url);
    const phone = (currentHeader.telefone || "").replace(/\D/g, "");
    const wa = phone
      ? `https://wa.me/${phone}?text=${encodeURIComponent(text + "\n\n(Imagem baixada — anexe no WhatsApp)")}`
      : `https://wa.me/?text=${encodeURIComponent(text + "\n\n(Imagem baixada — anexe no WhatsApp)")}`;
    window.open(wa, "_blank");
  };

  const downloadPng = async () => {
    const canvas = canvasRef.current;
    if (!canvas) return;
    const blob: Blob | null = await new Promise((r) => canvas.toBlob((b) => r(b), "image/png"));
    if (!blob) return;
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `dashboard-${currentHeader.data || "inspecao"}.png`;
    a.click();
    URL.revokeObjectURL(url);
  };

  return (
    <main className="min-h-screen bg-background text-foreground pb-16">
      {/* Page Header */}
      <header className="border-b border-border bg-card shadow-xs">
        <div className="mx-auto flex max-w-7xl flex-col sm:flex-row sm:items-center justify-between gap-4 px-4 py-6">
          <div>
            <div className="flex items-center gap-2">
              <h1 className="text-2xl font-bold tracking-tight">
                Dashboard & Panorama de Inspeções
              </h1>
              <span className="inline-flex items-center gap-1.5 rounded-full bg-emerald-500/10 px-2.5 py-0.5 text-xs font-semibold text-emerald-600 dark:text-emerald-400 border border-emerald-500/20">
                <Cloud className="h-3.5 w-3.5" />
                <span>Nuvem ({cloudReports.length})</span>
              </span>
            </div>
            <p className="mt-1 text-xs sm:text-sm text-muted-foreground">
              Análise consolidada e individual de todos os relatórios salvos no banco de dados.
            </p>
          </div>

          <div className="flex items-center gap-2">
            <Link
              to="/checklist"
              className="btn-secondary inline-flex items-center gap-1.5 text-xs py-2 px-3 border border-primary/30 text-primary hover:bg-primary/10"
              title="Acessar Checklist de Máquinas e Equipamentos"
            >
              <ClipboardCheck className="h-4 w-4" />
              <span>Checklist de Equipamento</span>
            </Link>

            <button
              onClick={() => exportAllReportsToExcel(filteredReports)}
              disabled={filteredReports.length === 0}
              className="btn-secondary inline-flex items-center gap-1.5 text-xs py-2 px-3"
              title="Exportar dados filtrados para planilha Excel"
            >
              <FileSpreadsheet className="h-4 w-4 text-emerald-600 dark:text-emerald-400" />
              <span>Exportar Excel ({filteredReports.length})</span>
            </button>

            <ThemeToggle className="py-2 px-3" />

            <Link to="/" className="btn-primary inline-flex items-center gap-1.5 text-xs py-2 px-3">
              <ArrowLeft className="h-4 w-4" />
              <span>Voltar ao Registro</span>
            </Link>
          </div>
        </div>
      </header>

      <div className="mx-auto max-w-7xl space-y-6 px-4 py-6">
        {/* Navigation Tabs (Panorama vs Relatório Individual) */}
        <div className="flex flex-wrap items-center justify-between gap-3 border-b border-border pb-3">
          <div className="flex items-center gap-2 bg-secondary/60 p-1 rounded-lg border border-border">
            <button
              onClick={() => setViewMode("panorama")}
              className={`inline-flex items-center gap-2 px-4 py-2 text-xs font-semibold rounded-md transition-all ${
                viewMode === "panorama"
                  ? "bg-card text-foreground shadow-xs border border-border"
                  : "text-muted-foreground hover:text-foreground"
              }`}
            >
              <BarChart3 className="h-4 w-4 text-primary" />
              <span>Panorama Geral ({filteredReports.length})</span>
            </button>

            <button
              onClick={() => setViewMode("single")}
              className={`inline-flex items-center gap-2 px-4 py-2 text-xs font-semibold rounded-md transition-all ${
                viewMode === "single"
                  ? "bg-card text-foreground shadow-xs border border-border"
                  : "text-muted-foreground hover:text-foreground"
              }`}
            >
              <Layers className="h-4 w-4 text-primary" />
              <span>Relatório Individual / Gráfico</span>
            </button>
          </div>

          {activeFiltersCount > 0 && (
            <div className="flex items-center gap-2 text-xs text-amber-600 dark:text-amber-400 font-medium">
              <Filter className="h-3.5 w-3.5" />
              <span>{activeFiltersCount} filtro(s) ativo(s)</span>
              <button
                onClick={clearAllFilters}
                className="ml-1 underline text-muted-foreground hover:text-foreground inline-flex items-center gap-1"
              >
                <RotateCcw className="h-3 w-3" />
                <span>Limpar</span>
              </button>
            </div>
          )}
        </div>

        {/* INTERACTIVE FILTERS BAR */}
        <section className="rounded-xl border border-border bg-card p-4 shadow-sm space-y-4">
          <div className="flex items-center justify-between border-b border-border/50 pb-2">
            <div className="flex items-center gap-2">
              <Filter className="h-4 w-4 text-primary" />
              <h2 className="text-xs font-bold uppercase tracking-wider text-foreground">
                Filtros Interativos da Nuvem
              </h2>
            </div>
            {activeFiltersCount > 0 && (
              <button
                onClick={clearAllFilters}
                className="text-xs text-muted-foreground hover:text-destructive inline-flex items-center gap-1"
              >
                <RotateCcw className="h-3 w-3" />
                <span>Resetar Filtros</span>
              </button>
            )}
          </div>

          <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-3 lg:grid-cols-6 gap-3">
            {/* Search Input */}
            <div className="space-y-1 sm:col-span-2 lg:col-span-1">
              <label className="text-[11px] font-semibold text-muted-foreground flex items-center gap-1">
                <Search className="h-3 w-3" />
                <span>Busca Geral</span>
              </label>
              <div className="relative">
                <input
                  type="text"
                  placeholder="Pesquisar termo..."
                  value={searchTerm}
                  onChange={(e) => setSearchTerm(e.target.value)}
                  className="input pl-8 text-xs py-1.5"
                />
                <Search className="absolute left-2.5 top-2.5 h-3.5 w-3.5 text-muted-foreground" />
              </div>
            </div>

            {/* Período Preset */}
            <div className="space-y-1">
              <label className="text-[11px] font-semibold text-muted-foreground flex items-center gap-1">
                <Calendar className="h-3 w-3" />
                <span>Período</span>
              </label>
              <select
                value={periodFilter}
                onChange={(e) => setPeriodFilter(e.target.value as typeof periodFilter)}
                className="input text-xs py-1.5"
              >
                <option value="todos">Todos os Períodos</option>
                <option value="hoje">Hoje</option>
                <option value="7dias">Últimos 7 dias</option>
                <option value="30dias">Últimos 30 dias</option>
                <option value="mes_atual">Este Mês</option>
                <option value="custom">Personalizado...</option>
              </select>
            </div>

            {/* Status da GDA */}
            <div className="space-y-1">
              <label className="text-[11px] font-semibold text-muted-foreground flex items-center gap-1">
                <Target className="h-3 w-3" />
                <span>Status GDA</span>
              </label>
              <select
                value={selectedGdaStatus}
                onChange={(e) => setSelectedGdaStatus(e.target.value as GdaStatusKey | "todos")}
                className="input text-xs py-1.5 font-medium"
              >
                <option value="todos">Todos os Status</option>
                <option value="executada">✓ GDA Executada</option>
                <option value="cancelada">✕ GDA Cancelada</option>
                <option value="sem_gda">? Sem GDA</option>
                <option value="reprovada">⚠ GDA Reprovada</option>
                <option value="parcial">⚡ Execução Parcial</option>
                <option value="planejada">○ Apenas Planejada</option>
              </select>
            </div>

            {/* Local / Trecho */}
            <div className="space-y-1">
              <label className="text-[11px] font-semibold text-muted-foreground flex items-center gap-1">
                <MapPin className="h-3 w-3" />
                <span>Local / Trecho</span>
              </label>
              <select
                value={selectedLocal}
                onChange={(e) => setSelectedLocal(e.target.value)}
                className="input text-xs py-1.5"
              >
                <option value="todos">Todos os Locais ({filterOptions.locais.length})</option>
                {filterOptions.locais.map((loc) => (
                  <option key={loc} value={loc}>
                    {loc}
                  </option>
                ))}
              </select>
            </div>

            {/* Atividade */}
            <div className="space-y-1">
              <label className="text-[11px] font-semibold text-muted-foreground flex items-center gap-1">
                <Activity className="h-3 w-3" />
                <span>Atividade</span>
              </label>
              <select
                value={selectedAtividade}
                onChange={(e) => setSelectedAtividade(e.target.value)}
                className="input text-xs py-1.5"
              >
                <option value="todos">
                  Todas as Atividades ({filterOptions.atividades.length})
                </option>
                {filterOptions.atividades.map((ativ) => (
                  <option key={ativ} value={ativ}>
                    {ativ}
                  </option>
                ))}
              </select>
            </div>

            {/* Responsável */}
            <div className="space-y-1">
              <label className="text-[11px] font-semibold text-muted-foreground flex items-center gap-1">
                <User className="h-3 w-3" />
                <span>Responsável</span>
              </label>
              <select
                value={selectedResponsavel}
                onChange={(e) => setSelectedResponsavel(e.target.value)}
                className="input text-xs py-1.5"
              >
                <option value="todos">Todos ({filterOptions.responsaveis.length})</option>
                {filterOptions.responsaveis.map((resp) => (
                  <option key={resp} value={resp}>
                    {resp}
                  </option>
                ))}
              </select>
            </div>
          </div>

          {/* Custom Date Pickers when periodFilter === 'custom' */}
          {periodFilter === "custom" && (
            <div className="grid grid-cols-1 sm:grid-cols-2 gap-3 pt-2 border-t border-border/40">
              <div className="space-y-1">
                <label className="text-[11px] font-medium text-muted-foreground">
                  Data Inicial
                </label>
                <input
                  type="date"
                  value={startDate}
                  onChange={(e) => setStartDate(e.target.value)}
                  className="input text-xs py-1.5"
                />
              </div>
              <div className="space-y-1">
                <label className="text-[11px] font-medium text-muted-foreground">Data Final</label>
                <input
                  type="date"
                  value={endDate}
                  onChange={(e) => setEndDate(e.target.value)}
                  className="input text-xs py-1.5"
                />
              </div>
            </div>
          )}
        </section>

        {/* VIEW MODE 1: PANORAMA GERAL */}
        {viewMode === "panorama" && (
          <div className="space-y-6">
            {/* KPI Cards Grid */}
            <div className="grid grid-cols-2 sm:grid-cols-3 lg:grid-cols-6 gap-3">
              <KpiCard
                icon={<Layers className="h-4 w-4 text-blue-500" />}
                label="Total Relatórios"
                value={String(panoramaMetrics.totalReportsCount)}
                subtitle="Inspeções salvas"
              />
              <KpiCard
                icon={<TrendingUp className="h-4 w-4 text-emerald-500" />}
                label="Produção Total"
                value={`${panoramaMetrics.totalProducao} m`}
                subtitle={`Média: ${Math.round(panoramaMetrics.avgProducao)}m / rel.`}
              />
              <KpiCard
                icon={<AlertTriangle className="h-4 w-4 text-rose-500" />}
                label="Defeitos Totais"
                value={String(panoramaMetrics.totalDefeitos)}
                subtitle="Ocorrências"
              />
              <KpiCard
                icon={<Clock className="h-4 w-4 text-emerald-600" />}
                label="Tempo Atividade"
                value={formatDuration(panoramaMetrics.totalAtividadeMs)}
                subtitle="Trabalho efetivo"
              />
              <KpiCard
                icon={<Clock className="h-4 w-4 text-amber-500" />}
                label="Tempo Paradas"
                value={formatDuration(panoramaMetrics.totalParadaMs)}
                subtitle="Tempo improdutivo"
              />
              <KpiCard
                icon={<CheckCircle2 className="h-4 w-4 text-indigo-500" />}
                label="Eficiência"
                value={`${panoramaMetrics.efficiency}%`}
                subtitle="Atividade / Jornada"
              />
            </div>

            {/* GDA Analytics Dashboard & Planned vs Executed Charts */}
            <GdaAnalyticsDashboard
              reports={filteredReports}
              selectedGdaFilter={selectedGdaStatus}
              onSelectGdaFilter={setSelectedGdaStatus}
            />

            {/* Panorama Graphic Canvas */}
            <div className="rounded-xl border border-border bg-card p-4 shadow-sm space-y-3">
              <div className="flex flex-wrap items-center justify-between gap-2 border-b border-border/50 pb-2">
                <div>
                  <h3 className="text-sm font-bold uppercase tracking-wider text-foreground">
                    Gráfico Comparativo de Produção e Defeitos
                  </h3>
                  <p className="text-xs text-muted-foreground">
                    Evolução cronológica dos relatórios salvos no banco de dados.
                  </p>
                </div>
              </div>

              <div className="overflow-x-auto">
                <canvas
                  ref={panoramaCanvasRef}
                  className="rounded-lg border border-border bg-white"
                />
              </div>
            </div>

            {/* Secondary Analytics: Top Paradas & Top Atividades */}
            <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
              {/* Top Paradas / Paralisações */}
              <div className="rounded-xl border border-border bg-card p-4 shadow-sm space-y-3">
                <div className="flex items-center gap-2 border-b border-border/50 pb-2">
                  <AlertTriangle className="h-4 w-4 text-amber-500" />
                  <h3 className="text-xs font-bold uppercase tracking-wider text-foreground">
                    Maiores Causas de Paradas / Perdas de Tempo
                  </h3>
                </div>

                {panoramaMetrics.topStops.length === 0 ? (
                  <p className="text-xs text-muted-foreground py-4 text-center">
                    Nenhuma parada registrada nos relatórios selecionados.
                  </p>
                ) : (
                  <ul className="space-y-2 text-xs">
                    {panoramaMetrics.topStops.map((stop, i) => (
                      <li
                        key={i}
                        className="flex items-center justify-between p-2 rounded-lg bg-secondary/40 border border-border/40"
                      >
                        <div className="flex items-center gap-2 min-w-0 pr-2">
                          <span className="font-bold text-muted-foreground text-[10px] w-4">
                            #{i + 1}
                          </span>
                          <span className="font-medium text-foreground truncate">
                            {stop.reason}
                          </span>
                        </div>
                        <div className="flex items-center gap-3 shrink-0 text-right">
                          <span className="text-muted-foreground font-mono">{stop.count}x</span>
                          <span className="font-mono font-bold text-amber-600 dark:text-amber-400">
                            {formatDuration(stop.durMs)}
                          </span>
                        </div>
                      </li>
                    ))}
                  </ul>
                )}
              </div>

              {/* Produção por Atividade */}
              <div className="rounded-xl border border-border bg-card p-4 shadow-sm space-y-3">
                <div className="flex items-center gap-2 border-b border-border/50 pb-2">
                  <Activity className="h-4 w-4 text-emerald-500" />
                  <h3 className="text-xs font-bold uppercase tracking-wider text-foreground">
                    Produção por Tipo de Atividade
                  </h3>
                </div>

                {panoramaMetrics.topActivities.length === 0 ? (
                  <p className="text-xs text-muted-foreground py-4 text-center">
                    Nenhuma atividade cadastrada nos relatórios.
                  </p>
                ) : (
                  <ul className="space-y-2 text-xs">
                    {panoramaMetrics.topActivities.slice(0, 5).map((ativ, i) => (
                      <li
                        key={i}
                        className="flex items-center justify-between p-2 rounded-lg bg-secondary/40 border border-border/40"
                      >
                        <div className="flex items-center gap-2 min-w-0 pr-2">
                          <span className="font-medium text-foreground truncate">{ativ.name}</span>
                          <span className="text-[10px] text-muted-foreground">
                            ({ativ.count} rel.)
                          </span>
                        </div>
                        <span className="font-mono font-bold text-emerald-600 dark:text-emerald-400">
                          {ativ.producao} m
                        </span>
                      </li>
                    ))}
                  </ul>
                )}
              </div>
            </div>

            {/* TABELA DETALHADA DOS RELATÓRIOS FILTRADOS */}
            <div className="rounded-xl border border-border bg-card p-4 shadow-sm space-y-3">
              <div className="flex flex-wrap items-center justify-between gap-2 border-b border-border/50 pb-2">
                <div className="flex items-center gap-2">
                  <Layers className="h-4 w-4 text-primary" />
                  <h3 className="text-xs font-bold uppercase tracking-wider text-foreground">
                    Relatórios Filtrados ({filteredReports.length})
                  </h3>
                </div>
                <span className="text-xs text-muted-foreground">
                  Clique em um relatório para visualizar o gráfico detalhado.
                </span>
              </div>

              {filteredReports.length === 0 ? (
                <div className="py-8 text-center text-xs text-muted-foreground">
                  Nenhum relatório encontrado para os filtros selecionados.
                </div>
              ) : (
                <div className="overflow-x-auto">
                  <table className="w-full text-left text-xs">
                    <thead className="bg-secondary/70 text-muted-foreground text-[10px] uppercase font-semibold">
                      <tr>
                        <th className="py-2.5 px-3">Data</th>
                        <th className="py-2.5 px-3">Status GDA / Nº</th>
                        <th className="py-2.5 px-3">Atividade / Descrição</th>
                        <th className="py-2.5 px-3">Prog. vs. Exec.</th>
                        <th className="py-2.5 px-3">Responsável</th>
                        <th className="py-2.5 px-3">Local / Trecho</th>
                        <th className="py-2.5 px-3">Produção</th>
                        <th className="py-2.5 px-3">Defeitos</th>
                        <th className="py-2.5 px-3">Tempo Parado</th>
                        <th className="py-2.5 px-3 text-right">Ação</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-border">
                      {filteredReports.map((r) => {
                        const gdaInfo = classifyGdaStatus(r);
                        return (
                          <tr
                            key={r.id}
                            className="hover:bg-accent/40 transition-colors cursor-pointer"
                            onClick={() => {
                              setSelectedReportId(r.id);
                              setViewMode("single");
                            }}
                          >
                            <td className="py-2.5 px-3 font-mono font-semibold text-foreground whitespace-nowrap">
                              {formatDate(r.data)}
                            </td>
                            <td className="py-2.5 px-3 whitespace-nowrap">
                              <div className="flex flex-col items-start gap-0.5">
                                <span
                                  className={`inline-flex items-center px-2 py-0.5 rounded text-[10px] font-semibold border ${gdaInfo.config.bgColor} ${gdaInfo.config.textColor} ${gdaInfo.config.borderColor}`}
                                >
                                  {gdaInfo.config.label}
                                </span>
                                <span className="font-mono text-[10px] text-muted-foreground">
                                  {r.gda && r.gda.trim() && r.gda.trim() !== "-"
                                    ? `GDA: ${r.gda}`
                                    : "Sem GDA"}
                                </span>
                              </div>
                            </td>
                            <td className="py-2.5 px-3 font-medium text-foreground">
                              <div>{r.atividade || "Sem Atividade"}</div>
                              {r.om && (
                                <div className="text-[10px] text-muted-foreground font-mono">
                                  OM: {r.om}
                                </div>
                              )}
                            </td>
                            <td className="py-2.5 px-3 font-mono text-[11px] whitespace-nowrap">
                              <div className="text-muted-foreground">
                                Prog: {r.kmProgramado || "-"}
                              </div>
                              <div
                                className={
                                  r.executado && r.executado.trim() !== "-"
                                    ? "text-foreground font-semibold"
                                    : "text-muted-foreground"
                                }
                              >
                                Exec: {r.executado || "-"}
                              </div>
                            </td>
                            <td className="py-2.5 px-3 text-muted-foreground whitespace-nowrap">
                              {r.nome || "-"}
                            </td>
                            <td className="py-2.5 px-3 text-muted-foreground whitespace-nowrap">
                              {r.local || "-"}
                            </td>
                            <td className="py-2.5 px-3 font-mono font-bold text-emerald-600 dark:text-emerald-400 whitespace-nowrap">
                              {r.totals?.producaoTotal || 0} m
                            </td>
                            <td className="py-2.5 px-3 font-mono font-bold text-rose-600 dark:text-rose-400 whitespace-nowrap">
                              {r.totals?.defeitosTotal || 0}
                            </td>
                            <td className="py-2.5 px-3 font-mono text-amber-600 dark:text-amber-400 whitespace-nowrap">
                              {formatDuration(r.totals?.parada || 0)}
                            </td>
                            <td className="py-2.5 px-3 text-right whitespace-nowrap">
                              <button
                                onClick={(e) => {
                                  e.stopPropagation();
                                  setSelectedReportId(r.id);
                                  setViewMode("single");
                                }}
                                className="btn-secondary text-[10px] py-1 px-2 inline-flex items-center gap-1"
                              >
                                <Eye className="h-3 w-3" />
                                <span>Ver Gráfico</span>
                              </button>
                            </td>
                          </tr>
                        );
                      })}
                    </tbody>
                  </table>
                </div>
              )}
            </div>
          </div>
        )}

        {/* VIEW MODE 2: RELATÓRIO INDIVIDUAL / GRÁFICO CONSOLIDADO */}
        {viewMode === "single" && (
          <div className="space-y-6">
            {/* Report Selector Bar */}
            <div className="rounded-xl border border-border bg-card p-4 shadow-sm flex flex-col sm:flex-row items-start sm:items-center justify-between gap-3">
              <div className="flex items-center gap-2">
                <Cloud className="h-5 w-5 text-primary" />
                <span className="text-xs font-semibold uppercase tracking-wider text-foreground">
                  Selecione o Relatório:
                </span>
              </div>

              <select
                value={selectedReportId}
                onChange={(e) => setSelectedReportId(e.target.value)}
                className="input text-xs w-full sm:w-auto font-medium"
              >
                <option value="local_draft">Relatório Atual (Rascunho Local)</option>
                {filteredReports.length > 0 && (
                  <option value="all_cloud">
                    Consolidado ({filteredReports.length} relatórios filtrados)
                  </option>
                )}
                {filteredReports.map((r) => (
                  <option key={r.id} value={r.id}>
                    {formatDate(r.data)} - {r.atividade || "Sem nome"} ({r.nome || "Sem resp."})
                  </option>
                ))}
              </select>
            </div>

            {/* If a single cloud report is selected, show GDA & Planned vs Executed badge header */}
            {selectedReportId !== "local_draft" &&
              selectedReportId !== "all_cloud" &&
              (() => {
                const r = cloudReports.find((item) => item.id === selectedReportId);
                if (!r) return null;
                const gda = classifyGdaStatus(r);
                return (
                  <div className="rounded-xl border border-border bg-card p-3.5 shadow-xs flex flex-wrap items-center justify-between gap-3">
                    <div className="flex flex-wrap items-center gap-2.5">
                      <span
                        className={`px-2.5 py-1 rounded-md text-xs font-semibold border ${gda.config.bgColor} ${gda.config.textColor} ${gda.config.borderColor}`}
                      >
                        {gda.config.label}
                      </span>
                      <span className="font-mono text-xs font-bold text-foreground">
                        {r.gda && r.gda.trim() && r.gda.trim() !== "-"
                          ? `GDA: ${r.gda}`
                          : "Sem GDA"}
                      </span>
                      {r.om && (
                        <span className="font-mono text-xs text-muted-foreground border-l border-border pl-2.5">
                          OM: {r.om}
                        </span>
                      )}
                    </div>
                    <div className="flex flex-wrap items-center gap-3 text-xs font-mono">
                      <span className="text-muted-foreground">
                        Programado:{" "}
                        <strong className="text-foreground">{r.kmProgramado || "-"}</strong>
                      </span>
                      <span className="text-muted-foreground">
                        Executado: <strong className="text-foreground">{r.executado || "-"}</strong>
                      </span>
                    </div>
                  </div>
                );
              })()}

            {/* Single Report Canvas */}
            <div className="rounded-xl border border-border bg-card p-4 shadow-sm">
              <div className="mb-3 flex flex-wrap items-center justify-between gap-2 border-b border-border/50 pb-2">
                <h2 className="text-xs font-bold uppercase tracking-wider text-foreground">
                  Gráfico Consolidado — {activeReportData.title}
                </h2>
                <div className="flex flex-wrap gap-2">
                  <button onClick={downloadPng} className="btn-secondary text-xs py-1.5 px-3">
                    Baixar PNG
                  </button>
                  <button onClick={shareWhatsApp} className="btn-primary text-xs py-1.5 px-3">
                    Compartilhar no WhatsApp
                  </button>
                </div>
              </div>

              <div className="overflow-x-auto">
                <canvas ref={canvasRef} className="rounded-lg border border-border bg-white" />
              </div>
            </div>

            {/* Summary Mini Cards */}
            <div className="grid gap-4 sm:grid-cols-2 lg:grid-cols-4">
              <MiniCard label="Produção total" value={`${tSingle.producaoTotal || 0} m`} />
              <MiniCard label="Defeitos" value={String(tSingle.defeitosTotal || 0)} />
              <MiniCard label="Tempo de atividade" value={formatDuration(tSingle.atividade)} />
              <MiniCard label="Tempo parado" value={formatDuration(tSingle.parada)} />
            </div>
          </div>
        )}
      </div>
    </main>
  );
}

function MiniCard({ label, value }: { label: string; value: string }) {
  return (
    <div className="rounded-lg border border-border bg-card px-4 py-3 shadow-xs">
      <div className="text-[10px] font-semibold uppercase tracking-widest text-muted-foreground">
        {label}
      </div>
      <div className="mt-1 font-mono text-2xl tabular-nums font-bold text-foreground">{value}</div>
    </div>
  );
}

function KpiCard({
  icon,
  label,
  value,
  subtitle,
}: {
  icon: React.ReactNode;
  label: string;
  value: string;
  subtitle: string;
}) {
  return (
    <div className="rounded-xl border border-border bg-card p-3 shadow-xs space-y-1">
      <div className="flex items-center justify-between">
        <span className="text-[10px] font-semibold uppercase tracking-wider text-muted-foreground">
          {label}
        </span>
        {icon}
      </div>
      <div className="font-mono text-xl font-bold text-foreground tabular-nums">{value}</div>
      <div className="text-[10px] text-muted-foreground truncate">{subtitle}</div>
    </div>
  );
}

// ============ Helpers de desenho no Canvas ============

function drawAxes(
  ctx: CanvasRenderingContext2D,
  x: number,
  y: number,
  w: number,
  h: number,
  title: string,
) {
  ctx.fillStyle = "#0f172a";
  ctx.font = "bold 13px system-ui, sans-serif";
  ctx.fillText(title, x, y - 8);
  ctx.strokeStyle = "#e2e8f0";
  ctx.lineWidth = 1;
  ctx.strokeRect(x, y, w, h);
}

function drawLineChart(
  ctx: CanvasRenderingContext2D,
  opts: {
    x: number;
    y: number;
    w: number;
    h: number;
    title: string;
    color: string;
    points: { x: number; y: number }[];
    xLabel: (v: number) => string;
    yLabel: (v: number) => string;
  },
) {
  const { x, y, w, h, title, color, points, xLabel, yLabel } = opts;
  const padL = 44,
    padB = 22,
    padT = 6,
    padR = 8;
  drawAxes(ctx, x, y, w, h, title);
  const innerX = x + padL;
  const innerY = y + padT;
  const innerW = w - padL - padR;
  const innerH = h - padT - padB;

  if (points.length === 0) {
    ctx.fillStyle = "#94a3b8";
    ctx.font = "12px system-ui, sans-serif";
    ctx.fillText("Sem dados", innerX + innerW / 2 - 30, innerY + innerH / 2);
    return;
  }

  const xs = points.map((p) => p.x);
  const ys = points.map((p) => p.y);
  const xMin = Math.min(...xs);
  const xMax = Math.max(...xs);
  const yMin = 0;
  const yMax = Math.max(...ys, 1);
  const xR = xMax - xMin || 1;
  const yR = yMax - yMin || 1;

  ctx.strokeStyle = "#f1f5f9";
  ctx.fillStyle = "#64748b";
  ctx.font = "10px system-ui, sans-serif";
  for (let i = 0; i <= 4; i++) {
    const gy = innerY + innerH - (i / 4) * innerH;
    ctx.beginPath();
    ctx.moveTo(innerX, gy);
    ctx.lineTo(innerX + innerW, gy);
    ctx.stroke();
    const v = yMin + (i / 4) * yR;
    ctx.fillText(yLabel(v), x + 4, gy + 3);
  }

  const nLabels = Math.min(5, points.length);
  for (let i = 0; i < nLabels; i++) {
    const idx = Math.round((i / (nLabels - 1 || 1)) * (points.length - 1));
    const p = points[idx];
    const px = innerX + ((p.x - xMin) / xR) * innerW;
    ctx.fillText(xLabel(p.x), px - 14, y + h - 6);
  }

  ctx.strokeStyle = color;
  ctx.lineWidth = 2;
  ctx.beginPath();
  points.forEach((p, i) => {
    const px = innerX + ((p.x - xMin) / xR) * innerW;
    const py = innerY + innerH - ((p.y - yMin) / yR) * innerH;
    if (i === 0) ctx.moveTo(px, py);
    else ctx.lineTo(px, py);
  });
  ctx.stroke();

  ctx.fillStyle = color + "22";
  ctx.beginPath();
  points.forEach((p, i) => {
    const px = innerX + ((p.x - xMin) / xR) * innerW;
    const py = innerY + innerH - ((p.y - yMin) / yR) * innerH;
    if (i === 0) ctx.moveTo(px, py);
    else ctx.lineTo(px, py);
  });
  const lastX = innerX + ((points[points.length - 1].x - xMin) / xR) * innerW;
  const firstX = innerX + ((points[0].x - xMin) / xR) * innerW;
  ctx.lineTo(lastX, innerY + innerH);
  ctx.lineTo(firstX, innerY + innerH);
  ctx.closePath();
  ctx.fill();

  ctx.fillStyle = color;
  points.forEach((p) => {
    const px = innerX + ((p.x - xMin) / xR) * innerW;
    const py = innerY + innerH - ((p.y - yMin) / yR) * innerH;
    ctx.beginPath();
    ctx.arc(px, py, 3, 0, Math.PI * 2);
    ctx.fill();
  });
}

function drawBarChart(
  ctx: CanvasRenderingContext2D,
  opts: {
    x: number;
    y: number;
    w: number;
    h: number;
    title: string;
    color: string;
    bars: { label: string; value: number }[];
  },
) {
  const { x, y, w, h, title, color, bars } = opts;
  drawAxes(ctx, x, y, w, h, title);
  const padL = 44,
    padB = 22,
    padT = 6,
    padR = 8;
  const innerX = x + padL;
  const innerY = y + padT;
  const innerW = w - padL - padR;
  const innerH = h - padT - padB;

  if (bars.length === 0) {
    ctx.fillStyle = "#94a3b8";
    ctx.font = "12px system-ui, sans-serif";
    ctx.fillText("Sem paradas registradas", innerX + innerW / 2 - 60, innerY + innerH / 2);
    return;
  }

  const yMax = Math.max(...bars.map((b) => b.value), 1);
  ctx.strokeStyle = "#f1f5f9";
  ctx.fillStyle = "#64748b";
  ctx.font = "10px system-ui, sans-serif";
  for (let i = 0; i <= 4; i++) {
    const gy = innerY + innerH - (i / 4) * innerH;
    ctx.beginPath();
    ctx.moveTo(innerX, gy);
    ctx.lineTo(innerX + innerW, gy);
    ctx.stroke();
    ctx.fillText(Math.round((i / 4) * yMax).toString(), x + 4, gy + 3);
  }

  const bw = innerW / bars.length;
  bars.forEach((b, i) => {
    const bh = (b.value / yMax) * innerH;
    const bx = innerX + i * bw + 4;
    const by = innerY + innerH - bh;
    ctx.fillStyle = color;
    ctx.fillRect(bx, by, Math.max(6, bw - 8), bh);
    ctx.fillStyle = "#64748b";
    ctx.font = "10px system-ui, sans-serif";
    ctx.fillText(b.label, bx, y + h - 6);
    ctx.fillStyle = "#0f172a";
    ctx.font = "bold 10px system-ui, sans-serif";
    const val = b.value < 10 ? b.value.toFixed(1) : Math.round(b.value).toString();
    ctx.fillText(val, bx, by - 3);
  });
}

function drawDonut(
  ctx: CanvasRenderingContext2D,
  opts: {
    cx: number;
    cy: number;
    r: number;
    inner: number;
    title: string;
    slices: { label: string; ms: number; color: string }[];
  },
) {
  const { cx, cy, r, inner, title, slices } = opts;
  ctx.fillStyle = "#0f172a";
  ctx.font = "bold 13px system-ui, sans-serif";
  ctx.fillText(title, cx - r, cy - r - 16);

  const total = slices.reduce((a, s) => a + s.ms, 0);
  if (total <= 0) {
    ctx.strokeStyle = "#e2e8f0";
    ctx.lineWidth = r - inner;
    ctx.beginPath();
    ctx.arc(cx, cy, (r + inner) / 2, 0, Math.PI * 2);
    ctx.stroke();
    ctx.fillStyle = "#94a3b8";
    ctx.font = "12px system-ui, sans-serif";
    ctx.fillText("Sem dados", cx - 30, cy);
    return;
  }
  let start = -Math.PI / 2;
  slices.forEach((s) => {
    if (s.ms <= 0) return;
    const ang = (s.ms / total) * Math.PI * 2;
    ctx.fillStyle = s.color;
    ctx.beginPath();
    ctx.moveTo(cx, cy);
    ctx.arc(cx, cy, r, start, start + ang);
    ctx.closePath();
    ctx.fill();
    start += ang;
  });

  ctx.fillStyle = "#ffffff";
  ctx.beginPath();
  ctx.arc(cx, cy, inner, 0, Math.PI * 2);
  ctx.fill();

  ctx.font = "11px system-ui, sans-serif";
  slices.forEach((s, i) => {
    const ly = cy + r + 18 + i * 16;
    ctx.fillStyle = s.color;
    ctx.fillRect(cx - r, ly - 9, 10, 10);
    ctx.fillStyle = "#0f172a";
    const pct = total > 0 ? Math.round((s.ms / total) * 100) : 0;
    ctx.fillText(`${s.label} — ${pct}%`, cx - r + 16, ly);
  });
}
