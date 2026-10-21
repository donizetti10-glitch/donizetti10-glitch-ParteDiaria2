import * as XLSX from "xlsx";
import { type SavedCloudReport } from "./cloud-reports";
import {
  EVENT_LABEL,
  formatDate,
  formatDuration,
  formatTime,
  sortAndNormalizeEvents,
} from "./parte-diaria";

export function buildSingleReportExcelBlob(report: SavedCloudReport): {
  blob: Blob;
  filename: string;
} {
  const wb = XLSX.utils.book_new();

  // 1. Aba Resumo e Cabeçalho
  const summaryRows = [
    ["REGISTRO DE ATIVIDADE DIÁRIA - RELATÓRIO INDIVIDUAL"],
    ["Código:", report.id, "Data de Envio:", new Date(report.createdAt).toLocaleString("pt-BR")],
    [""],
    ["1. CABEÇALHO DA ATIVIDADE"],
    ["Data do Relatório", formatDate(report.data)],
    ["Responsável", report.nome],
    ["Local", report.local],
    ["Sinais", report.sinais],
    ["Linha", report.linha],
    ["Via", report.via],
    ["GDA", report.gda],
    ["OM", report.om],
    ["Empresa", report.empresa],
    ["Nome da Atividade", report.atividade],
    ["Colaboradores / Equipe", report.colaboradores],
    ["Descritivo da Atividade", report.descritivo],
    ["Programado", report.kmProgramado],
    ["Executado", report.executado],
    ["Observação Geral", report.observacaoGeral],
    [""],
    ["2. RESUMO DE TEMPOS E PRODUÇÃO"],
    ["Tempo de Jornada", formatDuration(report.totals?.jornada ?? 0)],
    ["Tempo de Atividade", formatDuration(report.totals?.atividade ?? 0)],
    ["Tempo de Preparação", formatDuration(report.totals?.preparacao ?? 0)],
    ["Tempo de Paradas", formatDuration(report.totals?.parada ?? 0)],
    ["Produção Total (m)", report.totals?.producaoTotal ?? 0],
    ["Defeitos Totais", report.totals?.defeitosTotal ?? 0],
    [""],
    ["3. LINHA DO TEMPO DE EVENTOS"],
    [
      "Hora Início",
      "Hora Fim",
      "Duração",
      "Tipo de Evento",
      "Km Inicial",
      "Km Final",
      "Produção (m)",
      "Defeitos",
      "Último Defeito",
      "Descrição",
      "Observações",
    ],
  ];

  const sortedEvents = sortAndNormalizeEvents(report.events || []);

  for (let i = 0; i < sortedEvents.length; i++) {
    const e = sortedEvents[i];
    const next = sortedEvents[i + 1];

    const horaInicio = `${formatDate(report.data)} ${formatTime(e.timestamp)}`;
    let horaFim = "";
    let duracao = "";

    if (next) {
      horaFim = `${formatDate(report.data)} ${formatTime(next.timestamp)}`;
      duracao = formatDuration(Math.max(0, next.timestamp - e.timestamp));
    } else if (e.kind === "fim_jornada") {
      horaFim = `${formatDate(report.data)} ${formatTime(e.timestamp)}`;
      duracao = "00:00:00";
    } else {
      horaFim = "Em andamento";
      duracao = "(em andamento)";
    }

    summaryRows.push([
      horaInicio,
      horaFim,
      duracao,
      EVENT_LABEL[e.kind] || e.kind,
      e.kmInicial || "",
      e.kmFinal || "",
      e.producao || "",
      e.defeitos || "",
      e.ultimoDefeito || "",
      e.descricao || "",
      e.observacoes || "",
    ]);
  }

  const ws = XLSX.utils.aoa_to_sheet(summaryRows);

  ws["!cols"] = [
    { wch: 22 },
    { wch: 22 },
    { wch: 16 },
    { wch: 28 },
    { wch: 12 },
    { wch: 12 },
    { wch: 14 },
    { wch: 12 },
    { wch: 16 },
    { wch: 30 },
    { wch: 25 },
  ];

  XLSX.utils.book_append_sheet(wb, ws, "Relatório Diário");

  const filename = `relatorio-${report.data || "diario"}-${(report.nome || "atividade").replace(/\s+/g, "_")}.xlsx`;
  const wbout = XLSX.write(wb, { bookType: "xlsx", type: "array" });
  const blob = new Blob([wbout], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
  return { blob, filename };
}

export function exportSingleReportToExcel(report: SavedCloudReport) {
  const { blob, filename } = buildSingleReportExcelBlob(report);
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  a.click();
  URL.revokeObjectURL(url);
}

export function exportAllReportsToExcel(reports: SavedCloudReport[], customFilename?: string) {
  const wb = XLSX.utils.book_new();

  // Aba 1: Banco de Dados Geral (uma linha por relatório)
  const masterRows = [
    [
      "ID Relatório",
      "Data de Envio",
      "Data do Relatório",
      "Responsável",
      "Local",
      "Sinais",
      "Linha",
      "Via",
      "GDA",
      "OM",
      "Empresa",
      "Nome da Atividade",
      "Colaboradores / Equipe",
      "Descritivo da Atividade",
      "Programado",
      "Executado",
      "Produção Total (m)",
      "Defeitos Totais",
      "Tempo Jornada",
      "Tempo Atividade",
      "Tempo Preparação",
      "Tempo Paradas",
      "Qtd Eventos",
      "Observações Gerais",
    ],
  ];

  // Aba 2: Eventos Detalhados (uma linha por evento)
  const detailRows = [
    [
      "ID Relatório",
      "Data do Relatório",
      "Responsável",
      "Local",
      "Linha",
      "Via",
      "Atividade",
      "Empresa",
      "Hora Início",
      "Hora Fim",
      "Duração",
      "Tipo de Evento",
      "Km Inicial",
      "Km Final",
      "Produção (m)",
      "Defeitos",
      "Último Defeito",
      "Descrição do Evento",
      "Observações",
    ],
  ];

  for (const r of reports) {
    masterRows.push([
      r.id,
      new Date(r.createdAt).toLocaleString("pt-BR"),
      formatDate(r.data),
      r.nome || "",
      r.local || "",
      r.sinais || "",
      r.linha || "",
      r.via || "",
      r.gda || "",
      r.om || "",
      r.empresa || "",
      r.atividade || "",
      r.colaboradores || "",
      r.descritivo || "",
      r.kmProgramado || "",
      r.executado || "",
      String(r.totals?.producaoTotal ?? 0),
      String(r.totals?.defeitosTotal ?? 0),
      formatDuration(r.totals?.jornada ?? 0),
      formatDuration(r.totals?.atividade ?? 0),
      formatDuration(r.totals?.preparacao ?? 0),
      formatDuration(r.totals?.parada ?? 0),
      String(r.events?.length ?? 0),
      r.observacaoGeral || "",
    ]);

    const sortedEvents = sortAndNormalizeEvents(r.events || []);
    for (let i = 0; i < sortedEvents.length; i++) {
      const e = sortedEvents[i];
      const next = sortedEvents[i + 1];

      const horaInicio = `${formatDate(r.data)} ${formatTime(e.timestamp)}`;
      let horaFim = "";
      let duracao = "";

      if (next) {
        horaFim = `${formatDate(r.data)} ${formatTime(next.timestamp)}`;
        duracao = formatDuration(Math.max(0, next.timestamp - e.timestamp));
      } else if (e.kind === "fim_jornada") {
        horaFim = `${formatDate(r.data)} ${formatTime(e.timestamp)}`;
        duracao = "00:00:00";
      } else {
        horaFim = "Em andamento";
        duracao = "(em andamento)";
      }

      detailRows.push([
        r.id,
        formatDate(r.data),
        r.nome || "",
        r.local || "",
        r.linha || "",
        r.via || "",
        r.atividade || "",
        r.empresa || "",
        horaInicio,
        horaFim,
        duracao,
        EVENT_LABEL[e.kind] || e.kind,
        e.kmInicial || "",
        e.kmFinal || "",
        e.producao || "",
        e.defeitos || "",
        e.ultimoDefeito || "",
        e.descricao || "",
        e.observacoes || "",
      ]);
    }
  }

  const wsMaster = XLSX.utils.aoa_to_sheet(masterRows);
  wsMaster["!cols"] = [
    { wch: 16 }, // ID
    { wch: 18 }, // Envio
    { wch: 14 }, // Data
    { wch: 22 }, // Resp
    { wch: 16 }, // Local
    { wch: 12 }, // Sinais
    { wch: 10 }, // Linha
    { wch: 10 }, // Via
    { wch: 10 }, // GDA
    { wch: 10 }, // OM
    { wch: 18 }, // Empresa
    { wch: 25 }, // Atividade
    { wch: 30 }, // Equipe
    { wch: 35 }, // Descritivo
    { wch: 14 }, // Programado
    { wch: 14 }, // Executado
    { wch: 16 }, // Produção
    { wch: 14 }, // Defeitos
    { wch: 14 }, // Jornada
    { wch: 14 }, // Atividade
    { wch: 14 }, // Prep
    { wch: 14 }, // Parada
    { wch: 12 }, // Qtd Eventos
    { wch: 30 }, // Obs
  ];

  const wsDetail = XLSX.utils.aoa_to_sheet(detailRows);
  wsDetail["!cols"] = [
    { wch: 16 },
    { wch: 14 },
    { wch: 20 },
    { wch: 16 },
    { wch: 10 },
    { wch: 10 },
    { wch: 22 },
    { wch: 16 },
    { wch: 20 }, // Hora Início
    { wch: 20 }, // Hora Fim
    { wch: 16 }, // Duração
    { wch: 22 }, // Tipo
    { wch: 12 },
    { wch: 12 },
    { wch: 14 },
    { wch: 12 },
    { wch: 15 },
    { wch: 30 },
    { wch: 25 },
  ];

  XLSX.utils.book_append_sheet(wb, wsMaster, "Banco de Dados - Relatórios");
  XLSX.utils.book_append_sheet(wb, wsDetail, "Detalhamento de Eventos");

  const filename =
    customFilename || `banco-de-dados-relatorios-${new Date().toISOString().slice(0, 10)}.xlsx`;
  XLSX.writeFile(wb, filename);
}
