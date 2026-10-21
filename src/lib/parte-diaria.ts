export type EventKind =
  | "inicio_jornada"
  | "preparacao"
  | "inicio_atividade"
  | "fim_atividade"
  | "parada"
  | "deslocamento_ferroviario"
  | "deslocamento_rodoviario"
  | "pronto_cco"
  | "pronto_patio"
  | "pedido_cco"
  | "liberacao_acesso"
  | "chegada_local"
  | "retorno_base"
  | "dimensionamento_defeito"
  | "refeicao"
  | "cancelado"
  | "outros"
  | "fim_jornada";

export const EVENT_LABEL: Record<EventKind, string> = {
  inicio_jornada: "Início da jornada",
  preparacao: "Preparação/Manutenção",
  inicio_atividade: "Início da atividade",
  fim_atividade: "Fim da atividade",
  parada: "Parada",
  deslocamento_ferroviario: "Deslocamento Ferroviário",
  deslocamento_rodoviario: "Deslocamento Rodoviário",
  pronto_cco: "Pronto pra sair ao CCO",
  pronto_patio: "Pronto pra sair no pátio",
  pedido_cco: "Pedido ao CCO",
  liberacao_acesso: "Liberação de acesso a via",
  chegada_local: "Chegada ao Local",
  retorno_base: "Retorno a base",
  dimensionamento_defeito: "Dimensionamento Defeito",
  refeicao: "Janta/Almoço",
  cancelado: "Cancelado",
  outros: "Outros",
  fim_jornada: "Fim da jornada",
};

export interface WorkEvent {
  id: string;
  kind: EventKind;
  timestamp: number;
  local?: string;
  kmInicial?: string;
  kmFinal?: string;
  producao?: string;
  defeitos?: string;
  ultimoDefeito?: string;
  observacoes?: string;
  comentarios?: string;
  descricao?: string;
}

export interface DailyHeader {
  data: string;
  local: string;
  sinais: string;
  linha: string;
  via: string;
  gda: string;
  om: string;
  empresa: string;
  atividade: string;
  nome: string;
  colaboradores: string;
  descritivo: string;
  kmProgramado: string;
  executado: string;
  observacaoGeral: string;
  telefone: string;
}

export interface ReportPhoto {
  id: string;
  dataUrl: string;
  caption?: string;
}

export interface DailyReport {
  header: DailyHeader;
  events: WorkEvent[];
  photos?: ReportPhoto[];
}

export const emptyHeader = (): DailyHeader => ({
  data: new Date().toISOString().slice(0, 10),
  local: "",
  sinais: "",
  linha: "7",
  via: "",
  gda: "",
  om: "",
  empresa: "",
  atividade: "",
  nome: "",
  colaboradores: "",
  descritivo: "",
  kmProgramado: "",
  executado: "",
  observacaoGeral: "",
  telefone: "",
});

export const formatTime = (ts: number) =>
  new Date(ts).toLocaleTimeString("pt-BR", {
    hour: "2-digit",
    minute: "2-digit",
    second: "2-digit",
  });

export const formatTimeShort = (ts: number) =>
  new Date(ts).toLocaleTimeString("pt-BR", { hour: "2-digit", minute: "2-digit" });

export const formatDate = (iso: string) => {
  if (!iso) return "";
  const [y, m, d] = iso.split("-");
  return `${d}/${m}/${y}`;
};

export const formatPdfDate = (isoStr?: string): string => {
  if (!isoStr) {
    const now = new Date();
    const d = String(now.getDate()).padStart(2, "0");
    const m = String(now.getMonth() + 1).padStart(2, "0");
    const y = now.getFullYear();
    return `${d}.${m}.${y}`;
  }
  const parts = isoStr.split("-");
  if (parts.length === 3) {
    const [y, m, d] = parts;
    return `${d.padStart(2, "0")}.${m.padStart(2, "0")}.${y}`;
  }
  return isoStr.replace(/\//g, ".");
};

export const generatePdfFilename = (header: DailyHeader, events?: WorkEvent[]): string => {
  const dateStr = formatPdfDate(header.data);
  const isCancelado = Boolean(events && events.some((e) => e.kind === "cancelado"));

  const parts: string[] = ["RDA", dateStr];

  if (isCancelado) {
    parts.push("Cancelamento");
  }

  if (header.atividade && header.atividade.trim()) {
    parts.push(header.atividade.trim());
  }

  if (header.local && header.local.trim()) {
    parts.push(header.local.trim());
  }

  if (header.gda && header.gda.trim()) {
    parts.push(header.gda.trim());
  } else if (header.nome && header.nome.trim()) {
    parts.push(header.nome.trim());
  }

  const name = parts.filter(Boolean).join(" - ");
  const safeName = name.replace(/[/\\]/g, "_");
  return `${safeName}.pdf`;
};

export const maskSinais = (value: string): string => {
  const upper = value.toUpperCase();
  const raw = upper.replace(/[^A-Z0-9]/g, "").slice(0, 10);
  if (raw.length <= 5) {
    return raw;
  }
  return `${raw.slice(0, 5)}-${raw.slice(5)}`;
};

export const maskGda = (value: string): string => {
  const upper = value.toUpperCase();
  const raw = upper.replace(/[^A-Z0-9]/g, "").slice(0, 10);
  if (raw.length <= 4) {
    return raw;
  }
  return `${raw.slice(0, 4)}-${raw.slice(4)}`;
};

export const maskProgramadoExecutado = (value: string): string => {
  const digits = value.replace(/\D/g, "").slice(0, 8);
  if (digits.length <= 2) {
    return digits;
  }
  if (digits.length <= 4) {
    return `${digits.slice(0, 2)}/${digits.slice(2)}`;
  }
  if (digits.length <= 6) {
    return `${digits.slice(0, 2)}/${digits.slice(2, 4)}-${digits.slice(4)}`;
  }
  return `${digits.slice(0, 2)}/${digits.slice(2, 4)}-${digits.slice(4, 6)}/${digits.slice(6)}`;
};

export const formatDuration = (ms: number) => {
  if (ms < 0) ms = 0;
  const s = Math.floor(ms / 1000);
  const h = Math.floor(s / 3600);
  const m = Math.floor((s % 3600) / 60);
  const sec = s % 60;
  const pad = (n: number) => n.toString().padStart(2, "0");
  return `${pad(h)}:${pad(m)}:${pad(sec)}`;
};

export const formatDurationHM = (ms: number) => {
  if (ms < 0) ms = 0;
  const s = Math.floor(ms / 1000);
  const h = Math.floor(s / 3600);
  const m = Math.floor((s % 3600) / 60);
  const pad = (n: number) => n.toString().padStart(2, "0");
  return `${pad(h)}:${pad(m)}`;
};

export const sortAndNormalizeEvents = (events: WorkEvent[]): WorkEvent[] => {
  if (!events || events.length === 0) return [];

  // Sort strictly by timestamp ascending
  const sorted = [...events].sort((a, b) => {
    if (a.timestamp !== b.timestamp) return a.timestamp - b.timestamp;
    return a.id.localeCompare(b.id);
  });

  // Ensure timestamps are strictly increasing with no overlaps or duplicate timestamps
  const normalized: WorkEvent[] = [];
  for (let i = 0; i < sorted.length; i++) {
    const ev = { ...sorted[i] };
    if (i > 0) {
      const prevTs = normalized[i - 1].timestamp;
      if (ev.timestamp <= prevTs) {
        ev.timestamp = prevTs + 1000; // 1 second gap for identical or overlapping timestamps
      }
    }
    normalized.push(ev);
  }

  return normalized;
};

export const totals = (events: WorkEvent[]) => {
  const sorted = sortAndNormalizeEvents(events);
  const first = sorted[0]?.timestamp;
  const last = sorted[sorted.length - 1]?.timestamp;
  const jornada = first && last ? Math.max(0, last - first) : 0;

  let atividade = 0;
  let openStart: number | null = null;
  for (const e of sorted) {
    if (e.kind === "inicio_atividade") openStart = e.timestamp;
    if (e.kind === "fim_atividade" && openStart !== null) {
      atividade += Math.max(0, e.timestamp - openStart);
      openStart = null;
    }
  }

  let preparacao = 0;
  for (let i = 0; i < sorted.length; i++) {
    if (sorted[i].kind === "preparacao") {
      const next = sorted[i + 1];
      if (next) preparacao += Math.max(0, next.timestamp - sorted[i].timestamp);
    }
  }

  let parada = 0;
  for (let i = 0; i < sorted.length; i++) {
    if (sorted[i].kind === "parada") {
      const next = sorted[i + 1];
      if (next) parada += Math.max(0, next.timestamp - sorted[i].timestamp);
    }
  }

  const producaoTotal = sorted
    .filter((e) => e.kind === "fim_atividade")
    .reduce((acc, e) => acc + (Number(e.producao) || 0), 0);

  const defeitosTotal = sorted
    .filter((e) => e.kind === "fim_atividade")
    .reduce((acc, e) => acc + (Number(e.defeitos) || 0), 0);

  return { jornada, atividade, preparacao, parada, producaoTotal, defeitosTotal };
};

export const buildWhatsAppMessage = (report: DailyReport) => {
  const { header, events } = report;
  const sorted = sortAndNormalizeEvents(events);
  const t = totals(sorted);
  const lines: string[] = [];
  lines.push(`*REGISTRO DE ATIVIDADE DIÁRIA*`);
  lines.push(`Data: ${formatDate(header.data)}`);
  if (header.atividade) lines.push(`Atividade: ${header.atividade}`);
  if (header.descritivo) lines.push(`Descritivo da atividade: ${header.descritivo}`);
  if (header.nome) lines.push(`Responsável: ${header.nome}`);
  if (header.linha || header.via)
    lines.push(`Linha/Via: ${header.linha}${header.via ? " / " + header.via : ""}`);
  if (header.colaboradores) lines.push(`Equipe: ${header.colaboradores}`);
  if (header.kmProgramado) lines.push(`Km programado: ${header.kmProgramado}`);
  if (header.gda) lines.push(`GDA: ${header.gda}`);
  if (header.om) lines.push(`OM: ${header.om}`);
  lines.push("");
  lines.push(`*Eventos:*`);
  for (const e of sorted) {
    let extra = "";
    if (e.kind === "inicio_atividade" && e.kmInicial) extra = ` — km ${e.kmInicial}`;
    if (e.local) extra = ` — Local: ${e.local}`;
    if (e.kind === "fim_atividade") {
      const parts: string[] = [];
      if (e.kmFinal) parts.push(`km ${e.kmFinal}`);
      if (e.producao) parts.push(`${e.producao} m`);
      if (e.defeitos) parts.push(`${e.defeitos} defeito(s)`);
      if (parts.length) extra = ` — ${parts.join(", ")}`;
    }
    lines.push(`• ${formatTimeShort(e.timestamp)} ${EVENT_LABEL[e.kind]}${extra}`);
    if (e.descricao) lines.push(`   ${e.descricao}`);
    if (e.observacoes) lines.push(`   obs: ${e.observacoes}`);
  }
  lines.push("");
  lines.push(`*Totais:*`);
  lines.push(`Jornada: ${formatDurationHM(t.jornada)}`);
  lines.push(`Atividade: ${formatDurationHM(t.atividade)}`);
  if (t.preparacao) lines.push(`Preparação: ${formatDurationHM(t.preparacao)}`);
  if (t.parada) lines.push(`Paradas: ${formatDurationHM(t.parada)}`);
  if (t.producaoTotal) lines.push(`Produção total: ${t.producaoTotal} m`);
  if (t.defeitosTotal) lines.push(`Defeitos: ${t.defeitosTotal}`);
  return lines.join("\n");
};

export const buildCSV = (report: DailyReport) => {
  const { header, events } = report;
  const sorted = sortAndNormalizeEvents(events);
  const rows: string[] = [];
  const esc = (v: unknown) => {
    const s = v == null ? "" : String(v);
    return /[",;\n]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s;
  };
  rows.push(
    ["Data", "Nome", "Linha", "Via", "Colaboradores", "Km programado", "GDA", "OM"]
      .map(esc)
      .join(";"),
  );
  rows.push(
    [
      formatDate(header.data),
      header.nome,
      header.linha,
      header.via,
      header.colaboradores,
      header.kmProgramado,
      header.gda,
      header.om,
    ]
      .map(esc)
      .join(";"),
  );
  rows.push("");
  rows.push(
    [
      "Hora",
      "Evento",
      "Km inicial",
      "Km final",
      "Produção (m)",
      "Defeitos",
      "Último defeito",
      "Observações",
    ]
      .map(esc)
      .join(";"),
  );
  for (const e of sorted) {
    rows.push(
      [
        formatTime(e.timestamp),
        EVENT_LABEL[e.kind],
        e.kmInicial ?? "",
        e.kmFinal ?? "",
        e.producao ?? "",
        e.defeitos ?? "",
        e.ultimoDefeito ?? "",
        e.observacoes ?? "",
      ]
        .map(esc)
        .join(";"),
    );
  }
  return "\uFEFF" + rows.join("\n");
};
