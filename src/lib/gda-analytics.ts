import { type SavedCloudReport } from "./cloud-reports";
import { formatDate } from "./parte-diaria";

export type GdaStatusKey =
  "executada" | "cancelada" | "sem_gda" | "reprovada" | "parcial" | "planejada";

export interface GdaStatusConfig {
  key: GdaStatusKey;
  label: string;
  shortLabel: string;
  color: string;
  bgColor: string;
  textColor: string;
  borderColor: string;
  description: string;
}

export const GDA_STATUS_CONFIG: Record<GdaStatusKey, GdaStatusConfig> = {
  executada: {
    key: "executada",
    label: "GDA Executada",
    shortLabel: "Executada",
    color: "#16a34a", // Emerald 600
    bgColor: "bg-emerald-500/10 dark:bg-emerald-500/20",
    textColor: "text-emerald-700 dark:text-emerald-300",
    borderColor: "border-emerald-500/30",
    description: "Atividade com GDA realizada com sucesso e produção concluída.",
  },
  cancelada: {
    key: "cancelada",
    label: "GDA Cancelada",
    shortLabel: "Cancelada",
    color: "#dc2626", // Red 600
    bgColor: "bg-rose-500/10 dark:bg-rose-500/20",
    textColor: "text-rose-700 dark:text-rose-300",
    borderColor: "border-rose-500/30",
    description: "Atividade/GDA suspensa ou cancelada por intercorrências.",
  },
  sem_gda: {
    key: "sem_gda",
    label: "Sem GDA",
    shortLabel: "Sem GDA",
    color: "#64748b", // Slate 500
    bgColor: "bg-slate-500/10 dark:bg-slate-500/20",
    textColor: "text-slate-700 dark:text-slate-300",
    borderColor: "border-slate-500/30",
    description: "Atividade realizada ou programada sem registro formal de número de GDA.",
  },
  reprovada: {
    key: "reprovada",
    label: "GDA Reprovada",
    shortLabel: "Reprovada",
    color: "#ea580c", // Orange 600
    bgColor: "bg-amber-500/10 dark:bg-amber-500/20",
    textColor: "text-amber-800 dark:text-amber-300",
    borderColor: "border-amber-500/30",
    description: "Atividade com defeitos críticos, não conformidade ou reprovação técnica.",
  },
  parcial: {
    key: "parcial",
    label: "Execução Parcial",
    shortLabel: "Parcial",
    color: "#0284c7", // Sky 600
    bgColor: "bg-sky-500/10 dark:bg-sky-500/20",
    textColor: "text-sky-700 dark:text-sky-300",
    borderColor: "border-sky-500/30",
    description: "Produção abaixo do planejado ou com interrupções consideráveis.",
  },
  planejada: {
    key: "planejada",
    label: "Apenas Planejada",
    shortLabel: "Planejada",
    color: "#7c3aed", // Violet 600
    bgColor: "bg-violet-500/10 dark:bg-violet-500/20",
    textColor: "text-violet-700 dark:text-violet-300",
    borderColor: "border-violet-500/30",
    description: "Programação de atividade cadastrada ainda sem registro de execução.",
  },
};

/**
 * Converte notações de km/metragem como "25/21-26/21", "20/300-20/800" ou números diretos em metros.
 */
export function parseKmIntervalToMeters(str?: string): number | null {
  if (!str || typeof str !== "string") return null;
  const clean = str.trim();
  if (!clean || clean === "-") return null;

  // Formato direto numérico (ex: "500", "500m", "1500 m")
  const directNum = clean.replace(/m$/i, "").replace(/\./g, "").trim();
  if (/^\d+$/.test(directNum)) {
    const n = Number(directNum);
    return isNaN(n) ? null : n;
  }

  // Formato de trecho: "KmIni - KmFim" (ex: "25/21-26/21" ou "20/300 - 20/800")
  if (clean.includes("-")) {
    const parts = clean.split("-").map((p) => p.trim());
    if (parts.length === 2) {
      const p1 = parseMilestoneToMeters(parts[0]);
      const p2 = parseMilestoneToMeters(parts[1]);
      if (p1 !== null && p2 !== null) {
        return Math.abs(p2 - p1);
      }
    }
  }

  return null;
}

function parseMilestoneToMeters(milestone: string): number | null {
  const m = milestone.trim();
  if (!m) return null;

  // "25/21" -> km 25 + 21 (se 2 dígitos, consideramos 210 metros; se 3 dígitos, 210m; se 1 dígito, 100m)
  if (m.includes("/")) {
    const [kmPart, decPart] = m.split("/").map((s) => s.trim());
    const km = Number(kmPart);
    if (isNaN(km)) return null;

    let meters = 0;
    if (decPart) {
      // Normaliza para metros (3 casas decimais de km)
      const padded = decPart.padEnd(3, "0").slice(0, 3);
      meters = Number(padded) || 0;
    }
    return km * 1000 + meters;
  }

  // Apenas número
  const val = Number(m.replace(/\D/g, ""));
  return isNaN(val) ? null : val;
}

export interface ClassifiedGda {
  status: GdaStatusKey;
  config: GdaStatusConfig;
  isPlanned: boolean;
  isExecuted: boolean;
  isCancelled: boolean;
  isReproved: boolean;
  gdaCode: string;
  plannedMeters: number;
  executedMeters: number;
  motivoCancelamento?: string;
}

/**
 * Classifica o status da GDA de um relatório com base nas regras operacionais ferroviárias.
 */
export function classifyGdaStatus(report: SavedCloudReport): ClassifiedGda {
  const rawGda = (report.gda || "").trim();
  const rawObs = (report.observacaoGeral || "").toLowerCase();
  const rawDesc = (report.descritivo || "").toLowerCase();
  const rawAtiv = (report.atividade || "").toLowerCase();

  // Verifica eventos do relatório
  const events = report.events || [];
  const cancelEvent = events.find((e) => e.kind === "cancelado");
  const hasCancelEvent = Boolean(cancelEvent);
  const cancelTextMatch =
    rawObs.includes("cancelad") ||
    rawDesc.includes("cancelad") ||
    rawAtiv.includes("cancelad") ||
    rawGda.toLowerCase().includes("canc");

  const isCancelled = hasCancelEvent || cancelTextMatch;
  const motivoCancelamento =
    cancelEvent?.descricao ||
    (cancelTextMatch ? report.observacaoGeral || report.descritivo : undefined);

  // Verificação de Sem GDA
  const hasNoGda =
    !rawGda ||
    rawGda === "-" ||
    rawGda.toLowerCase() === "s/gda" ||
    rawGda.toLowerCase() === "sem gda" ||
    rawGda.toLowerCase() === "não informado";

  // Verificação de Defeitos / Reprovação
  const defeitosCount = report.totals?.defeitosTotal ?? 0;
  const hasDefectsInEvents = events.some((e) => Number(e.defeitos) > 0);
  const reprovaTextMatch =
    rawObs.includes("reprovad") ||
    rawDesc.includes("reprovad") ||
    rawObs.includes("não conforme") ||
    rawObs.includes("inapto");

  const isReproved = defeitosCount > 0 || hasDefectsInEvents || reprovaTextMatch;

  // Metragens planejada vs executada
  const parsedProgMeters = parseKmIntervalToMeters(report.kmProgramado);
  const parsedExecMeters = parseKmIntervalToMeters(report.executado);
  const totalProd = report.totals?.producaoTotal ?? 0;

  const plannedMeters = parsedProgMeters ?? (totalProd > 0 ? totalProd : 0);
  const executedMeters = parsedExecMeters ?? totalProd;

  // Verifica execução
  const hasFinishedEvent = events.some((e) => e.kind === "fim_atividade");
  const hasExecField = Boolean(
    report.executado && report.executado.trim() && report.executado !== "-",
  );
  const hasProduction = totalProd > 0 || executedMeters > 0;
  const hasWorkDone = hasFinishedEvent || hasExecField || hasProduction;

  let status: GdaStatusKey = "executada";

  if (isCancelled) {
    status = "cancelada";
  } else if (hasNoGda) {
    status = "sem_gda";
  } else if (isReproved) {
    status = "reprovada";
  } else if (!hasWorkDone) {
    status = "planejada";
  } else if (plannedMeters > 0 && executedMeters > 0 && executedMeters < plannedMeters * 0.9) {
    status = "parcial";
  } else {
    status = "executada";
  }

  const isPlanned = true; // Todo relatório registrado representa uma atividade programada
  const isExecuted =
    !isCancelled &&
    (status === "executada" || status === "parcial" || (status === "sem_gda" && hasWorkDone));

  return {
    status,
    config: GDA_STATUS_CONFIG[status],
    isPlanned,
    isExecuted,
    isCancelled,
    isReproved,
    gdaCode: hasNoGda ? "Sem GDA" : rawGda,
    plannedMeters,
    executedMeters: isCancelled ? 0 : executedMeters,
    motivoCancelamento,
  };
}

export interface GdaDailyTimelineItem {
  date: string;
  dateLabel: string;
  planejadas: number;
  executadas: number;
  canceladas: number;
  semGda: number;
  reprovadas: number;
  parciais: number;
  taxaAderencia: number;
  programadoMeters: number;
  executadoMeters: number;
}

export interface GdaResponsibleItem {
  nome: string;
  planejadas: number;
  executadas: number;
  canceladas: number;
  semGda: number;
  reprovadas: number;
  taxaAderencia: number;
}

export interface GdaAnalyticsSummary {
  totalPlanned: number;
  totalExecuted: number;
  totalCancelled: number;
  totalSemGda: number;
  totalReproved: number;
  totalPartial: number;
  totalWithoutExecution: number;
  adherenceRate: number; // % (executadas / planejadas * 100)
  totalPlannedMeters: number;
  totalExecutedMeters: number;
  metersAdherenceRate: number; // % (executadoMeters / plannedMeters * 100)
  statusDistribution: {
    key: GdaStatusKey;
    label: string;
    value: number;
    color: string;
    percentage: number;
  }[];
  dailyTimeline: GdaDailyTimelineItem[];
  responsibleStats: GdaResponsibleItem[];
  lineStats: {
    linha: string;
    planejadas: number;
    executadas: number;
    canceladas: number;
  }[];
}

/**
 * Calcula todas as métricas analíticas e estatísticas de GDA e Planejado vs Executado.
 */
export function calculateGdaAnalytics(reports: SavedCloudReport[]): GdaAnalyticsSummary {
  let totalPlanned = 0;
  let totalExecuted = 0;
  let totalCancelled = 0;
  let totalSemGda = 0;
  let totalReproved = 0;
  let totalPartial = 0;
  let totalWithoutExecution = 0;

  let totalPlannedMeters = 0;
  let totalExecutedMeters = 0;

  const statusCounts: Record<GdaStatusKey, number> = {
    executada: 0,
    cancelada: 0,
    sem_gda: 0,
    reprovada: 0,
    parcial: 0,
    planejada: 0,
  };

  const dailyMap: Record<
    string,
    {
      date: string;
      planejadas: number;
      executadas: number;
      canceladas: number;
      semGda: number;
      reprovadas: number;
      parciais: number;
      progMeters: number;
      execMeters: number;
    }
  > = {};

  const respMap: Record<
    string,
    {
      planejadas: number;
      executadas: number;
      canceladas: number;
      semGda: number;
      reprovadas: number;
    }
  > = {};

  const lineMap: Record<
    string,
    {
      planejadas: number;
      executadas: number;
      canceladas: number;
    }
  > = {};

  reports.forEach((report) => {
    totalPlanned++;
    const classified = classifyGdaStatus(report);

    statusCounts[classified.status]++;

    if (classified.isExecuted) totalExecuted++;
    if (classified.isCancelled) totalCancelled++;
    if (classified.status === "sem_gda") totalSemGda++;
    if (classified.isReproved) totalReproved++;
    if (classified.status === "parcial") totalPartial++;
    if (classified.status === "planejada") totalWithoutExecution++;

    totalPlannedMeters += classified.plannedMeters;
    totalExecutedMeters += classified.executedMeters;

    // Agrupamento por dia (Data)
    const dateKey = report.data || "Sem Data";
    if (!dailyMap[dateKey]) {
      dailyMap[dateKey] = {
        date: dateKey,
        planejadas: 0,
        executadas: 0,
        canceladas: 0,
        semGda: 0,
        reprovadas: 0,
        parciais: 0,
        progMeters: 0,
        execMeters: 0,
      };
    }
    dailyMap[dateKey].planejadas++;
    if (classified.isExecuted) dailyMap[dateKey].executadas++;
    if (classified.isCancelled) dailyMap[dateKey].canceladas++;
    if (classified.status === "sem_gda") dailyMap[dateKey].semGda++;
    if (classified.isReproved) dailyMap[dateKey].reprovadas++;
    if (classified.status === "parcial") dailyMap[dateKey].parciais++;
    dailyMap[dateKey].progMeters += classified.plannedMeters;
    dailyMap[dateKey].execMeters += classified.executedMeters;

    // Agrupamento por Responsável
    const respName = report.nome?.trim() || "Não Informado";
    if (!respMap[respName]) {
      respMap[respName] = { planejadas: 0, executadas: 0, canceladas: 0, semGda: 0, reprovadas: 0 };
    }
    respMap[respName].planejadas++;
    if (classified.isExecuted) respMap[respName].executadas++;
    if (classified.isCancelled) respMap[respName].canceladas++;
    if (classified.status === "sem_gda") respMap[respName].semGda++;
    if (classified.isReproved) respMap[respName].reprovadas++;

    // Agrupamento por Linha
    const lineKey = report.linha?.trim() ? `Linha ${report.linha.trim()}` : "Linha N/I";
    if (!lineMap[lineKey]) {
      lineMap[lineKey] = { planejadas: 0, executadas: 0, canceladas: 0 };
    }
    lineMap[lineKey].planejadas++;
    if (classified.isExecuted) lineMap[lineKey].executadas++;
    if (classified.isCancelled) lineMap[lineKey].canceladas++;
  });

  const adherenceRate = totalPlanned > 0 ? Math.round((totalExecuted / totalPlanned) * 100) : 0;

  const metersAdherenceRate =
    totalPlannedMeters > 0
      ? Math.min(150, Math.round((totalExecutedMeters / totalPlannedMeters) * 100))
      : 0;

  // Status Distribution Pie Data
  const statusDistribution = (Object.keys(GDA_STATUS_CONFIG) as GdaStatusKey[])
    .map((key) => {
      const val = statusCounts[key];
      const cfg = GDA_STATUS_CONFIG[key];
      const pct = totalPlanned > 0 ? Math.round((val / totalPlanned) * 100) : 0;
      return {
        key,
        label: cfg.label,
        value: val,
        color: cfg.color,
        percentage: pct,
      };
    })
    .filter((item) => item.value > 0);

  // Daily Timeline Data ordenada cronologicamente
  const dailyTimeline: GdaDailyTimelineItem[] = Object.keys(dailyMap)
    .sort()
    .map((date) => {
      const d = dailyMap[date];
      const taxa = d.planejadas > 0 ? Math.round((d.executadas / d.planejadas) * 100) : 0;
      return {
        date,
        dateLabel: formatDate(date).slice(0, 5), // DD/MM
        planejadas: d.planejadas,
        executadas: d.executadas,
        canceladas: d.canceladas,
        semGda: d.semGda,
        reprovadas: d.reprovadas,
        parciais: d.parciais,
        taxaAderencia: taxa,
        programadoMeters: d.progMeters,
        executadoMeters: d.execMeters,
      };
    });

  // Responsible stats
  const responsibleStats: GdaResponsibleItem[] = Object.entries(respMap)
    .map(([nome, v]) => ({
      nome,
      planejadas: v.planejadas,
      executadas: v.executadas,
      canceladas: v.canceladas,
      semGda: v.semGda,
      reprovadas: v.reprovadas,
      taxaAderencia: v.planejadas > 0 ? Math.round((v.executadas / v.planejadas) * 100) : 0,
    }))
    .sort((a, b) => b.planejadas - a.planejadas);

  // Line stats
  const lineStats = Object.entries(lineMap)
    .map(([linha, v]) => ({
      linha,
      planejadas: v.planejadas,
      executadas: v.executadas,
      canceladas: v.canceladas,
    }))
    .sort((a, b) => b.planejadas - a.planejadas);

  return {
    totalPlanned,
    totalExecuted,
    totalCancelled,
    totalSemGda,
    totalReproved,
    totalPartial,
    totalWithoutExecution,
    adherenceRate,
    totalPlannedMeters,
    totalExecutedMeters,
    metersAdherenceRate,
    statusDistribution,
    dailyTimeline,
    responsibleStats,
    lineStats,
  };
}
