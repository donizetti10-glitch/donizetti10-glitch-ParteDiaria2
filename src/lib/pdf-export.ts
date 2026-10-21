import jsPDF from "jspdf";
import autoTable from "jspdf-autotable";
import { TIC_TRENS_LOGO_DATA_URL } from "./tic-trens-logo-data";
import {
  type DailyHeader,
  type WorkEvent,
  type ReportPhoto,
  EVENT_LABEL,
  formatDate,
  generatePdfFilename,
  totals,
} from "./parte-diaria";

const pad2 = (n: number) => n.toString().padStart(2, "0");

function getImageDimensions(dataUrl: string): Promise<{ width: number; height: number }> {
  return new Promise((resolve) => {
    const img = new Image();
    img.onload = () => {
      resolve({ width: img.width || 800, height: img.height || 600 });
    };
    img.onerror = () => {
      resolve({ width: 800, height: 600 });
    };
    img.src = dataUrl;
  });
}

export async function buildPdfBlobForReport(
  header: DailyHeader,
  events: WorkEvent[],
  photos?: ReportPhoto[],
): Promise<Blob> {
  const doc = new jsPDF({ unit: "mm", format: "a4" });
  const h = header;
  const sortedEvents = [...(events || [])].sort((a, b) => a.timestamp - b.timestamp);
  const t = totals(sortedEvents);

  const pageW = doc.internal.pageSize.getWidth();
  const pageH = doc.internal.pageSize.getHeight();
  const marginL = 12;
  const marginR = 12;
  const contentW = pageW - marginL - marginR;
  const boxTop = 10;

  let y = 14;

  // Logo TIC TRENS
  try {
    doc.addImage(TIC_TRENS_LOGO_DATA_URL, "JPEG", marginL + 2, y, 48, 15);
    y += 18;
  } catch {
    y += 2;
  }

  // Título do Relatório
  doc.setFont("helvetica", "bold");
  doc.setFontSize(13);
  doc.text("Relatório de Atividade", marginL + 2, y);
  y += 5;

  if (h.atividade) {
    doc.setFont("helvetica", "bold");
    doc.setFontSize(11);
    const wrapped = doc.splitTextToSize(h.atividade, contentW - 4);
    doc.text(wrapped, marginL + 2, y);
    y += wrapped.length * 5 + 1;
  } else {
    y += 1;
  }

  // Campos do cabeçalho
  doc.setFontSize(9.5);
  const labelX = marginL + 2;
  const valueX = marginL + 36;

  const drawKV = (label: string, value: string) => {
    doc.setFont("helvetica", "bold");
    doc.text(label, labelX, y);
    doc.setFont("helvetica", "normal");
    doc.text(value || "", valueX, y);
    y += 4.5;
  };

  drawKV("DATA:", formatDate(h.data));
  drawKV("RESPONSÁVEL:", h.nome);
  drawKV("LOCAL:", h.local);
  drawKV("SINAIS:", h.sinais);
  drawKV("LINHA:", h.linha);
  drawKV("VIA:", h.via);
  drawKV("GDA:", h.gda);
  if (h.om) drawKV("OM:", h.om);
  y += 2;

  // PARTICIPANTES
  doc.setFont("helvetica", "bold");
  doc.setFontSize(10);
  doc.text("PARTICIPANTES:", labelX, y);
  y += 3;

  autoTable(doc, {
    startY: y,
    margin: { left: marginL, right: marginR },
    head: [["Empresa", "Colaboradores"]],
    body: [[h.empresa || "TIC Trens", h.colaboradores || ""]],
    styles: { fontSize: 8.5, cellPadding: 2, lineColor: [0, 0, 0], lineWidth: 0.2 },
    headStyles: { fillColor: [245, 245, 245], textColor: 0, fontStyle: "bold", lineWidth: 0.2 },
    columnStyles: { 0: { cellWidth: 45 } },
  });
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  y = ((doc as any).lastAutoTable?.finalY ?? y) + 5;

  // Descritivo da Atividade
  doc.setFont("helvetica", "bold");
  doc.setFontSize(10);
  doc.text("Descritivo da Atividade:", labelX, y);
  y += 4;

  if (h.descritivo) {
    doc.setFont("helvetica", "normal");
    doc.setFontSize(8.5);
    const wrapped = doc.splitTextToSize(h.descritivo, contentW - 4);
    doc.text(wrapped, labelX, y);
    y += wrapped.length * 4.5 + 2;
  }

  // Tabela cronológica
  const tableBody = sortedEvents.map((e) => {
    const d = new Date(e.timestamp);
    const dd = `${pad2(d.getDate())}/${pad2(d.getMonth() + 1)}/${d.getFullYear()}`;
    const hh = `${pad2(d.getHours())}:${pad2(d.getMinutes())}:${pad2(d.getSeconds())}`;
    const dtStr = `${dd}  ${hh}`;

    const label = EVENT_LABEL[e.kind] || e.kind;
    let desc = "";

    if (e.kind === "inicio_atividade") {
      const kms = [e.kmInicial, e.kmFinal].filter(Boolean).join("/");
      const kmStr = kms ? ` (km ${kms})` : "";
      const detail = e.descricao ? ` — ${e.descricao}` : "";
      desc = `${label}${kmStr}${detail}`;
    } else if (e.kind === "fim_atividade") {
      const kms = [e.kmInicial, e.kmFinal].filter(Boolean).join("/");
      const parts: string[] = [];
      if (kms) parts.push(`km ${kms}`);
      if (e.producao) parts.push(`${e.producao} m`);
      if (e.defeitos !== undefined && e.defeitos !== null && e.defeitos !== "")
        parts.push(`${e.defeitos} def.`);
      const extraStr = parts.length ? ` (${parts.join(", ")})` : "";
      const detail = e.descricao ? ` — ${e.descricao}` : "";
      desc = `${label}${extraStr}${detail}`;
    } else {
      const parts = [e.descricao, e.observacoes].filter(Boolean);
      if (parts.length > 0) {
        desc = `${label} — ${parts.join(" — ")}`;
      } else {
        desc = label;
      }
    }

    return [dtStr, desc];
  });

  autoTable(doc, {
    startY: y,
    margin: { left: marginL, right: marginR },
    showHead: "never",
    body: tableBody,
    styles: {
      fontSize: 8.5,
      cellPadding: 1.8,
      lineColor: [255, 255, 255],
      lineWidth: 0,
      textColor: [20, 20, 20],
    },
    alternateRowStyles: {
      fillColor: [248, 249, 250],
    },
    bodyStyles: {
      fillColor: [255, 255, 255],
    },
    columnStyles: {
      0: { cellWidth: 40 },
      1: { cellWidth: contentW - 40 },
    },
  });
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  y = ((doc as any).lastAutoTable?.finalY ?? y) + 5;

  // Resumo do Rodapé
  doc.setFontSize(9.5);
  const drawSummary = (label: string, value: string) => {
    doc.setFont("helvetica", "bold");
    doc.text(label, labelX, y);
    doc.setFont("helvetica", "normal");
    doc.text(value || "", valueX, y);
    y += 4.5;
  };

  drawSummary("Produção", String(t.producaoTotal || ""));
  drawSummary("Defeitos", String(t.defeitosTotal ?? 0));
  drawSummary("Programado", h.kmProgramado || "");
  drawSummary("Executado:", h.executado || "");

  if (h.observacaoGeral) {
    doc.setFont("helvetica", "bold");
    doc.text("Observação:", labelX, y);
    doc.setFont("helvetica", "normal");
    const wrapped = doc.splitTextToSize(h.observacaoGeral, contentW - 38);
    doc.text(wrapped, valueX, y);
    y += wrapped.length * 4.5 + 2;
  }

  // Borda externa página 1
  doc.setDrawColor(0);
  doc.setLineWidth(0.3);
  const finalBoxHeight = Math.max(y - boxTop + 3, pageH - 2 * boxTop);
  doc.rect(marginL, boxTop, contentW, finalBoxHeight);

  // ANEXO FOTOGRÁFICO DE INSPEÇÃO (Opcional - até 12 fotos, máximo 6 por página)
  if (photos && photos.length > 0) {
    const targetPhotos = photos.slice(0, 12);
    const photosPerPage = 6;
    const totalPhotoPages = Math.ceil(targetPhotos.length / photosPerPage);

    for (let pageIdx = 0; pageIdx < totalPhotoPages; pageIdx++) {
      doc.addPage("a4", "p");

      const photoPageTop = 10;
      let py = photoPageTop + 4;

      // Header do Anexo
      try {
        doc.addImage(TIC_TRENS_LOGO_DATA_URL, "JPEG", marginL + 2, py, 40, 12.5);
      } catch {
        // ignore logo load error
      }

      doc.setFont("helvetica", "bold");
      doc.setFontSize(12);
      doc.setTextColor(0, 0, 0);
      const pageSuffix = totalPhotoPages > 1 ? ` (${pageIdx + 1}/${totalPhotoPages})` : "";
      doc.text(`ANEXO FOTOGRÁFICO DE INSPEÇÃO${pageSuffix}`, marginL + 46, py + 5);

      doc.setFont("helvetica", "normal");
      doc.setFontSize(8.5);
      doc.setTextColor(60, 60, 60);
      const subHeader = `Atividade: ${h.atividade || "Relatório"}  |  Data: ${formatDate(h.data)}  |  Local: ${h.local || "-"}`;
      doc.text(doc.splitTextToSize(subHeader, contentW - 48), marginL + 46, py + 9);

      py += 15;
      doc.setDrawColor(180, 185, 190);
      doc.setLineWidth(0.3);
      doc.line(marginL, py, marginL + contentW, py);
      py += 5;

      // Fotos desta página (até 6 fotos por página)
      const pagePhotos = targetPhotos.slice(pageIdx * photosPerPage, (pageIdx + 1) * photosPerPage);

      // Grid de 2 colunas x até 3 linhas (para até 6 fotos por página)
      const colW = 88; // mm
      const gapX = 10; // mm
      const rowH = 72; // mm
      const gapY = 6; // mm

      for (let i = 0; i < pagePhotos.length; i++) {
        const globalIdx = pageIdx * photosPerPage + i;
        const p = pagePhotos[i];
        const col = i % 2;
        const row = Math.floor(i / 2);

        const cellX = marginL + col * (colW + gapX);
        const cellY = py + row * (rowH + gapY);

        // Moldura do card da foto
        doc.setDrawColor(200, 205, 210);
        doc.setFillColor(252, 253, 254);
        doc.rect(cellX, cellY, colW, rowH, "FD");

        // Banner superior da foto (FOTO 01, FOTO 02...)
        doc.setFillColor(240, 242, 245);
        doc.rect(cellX, cellY, colW, 6, "F");
        doc.setFont("helvetica", "bold");
        doc.setFontSize(8);
        doc.setTextColor(40, 40, 40);
        doc.text(`FOTO ${pad2(globalIdx + 1)}`, cellX + 3, cellY + 4.2);

        // Bounding box para a imagem sem distorção
        const boxX = cellX + 2;
        const boxY = cellY + 7;
        const boxW = colW - 4; // 84mm
        const boxH = rowH - 15; // 57mm

        try {
          const dims = await getImageDimensions(p.dataUrl);
          const imgRatio = dims.width / dims.height;
          const boxRatio = boxW / boxH;

          let drawW = boxW;
          let drawH = boxH;

          if (imgRatio > boxRatio) {
            drawW = boxW;
            drawH = boxW / imgRatio;
          } else {
            drawH = boxH;
            drawW = boxH * imgRatio;
          }

          const drawX = boxX + (boxW - drawW) / 2;
          const drawY = boxY + (boxH - drawH) / 2;

          const imgFormat = p.dataUrl.startsWith("data:image/png") ? "PNG" : "JPEG";
          doc.addImage(p.dataUrl, imgFormat, drawX, drawY, drawW, drawH);
        } catch (e) {
          console.warn("Erro ao renderizar foto no PDF:", e);
        }

        // Legenda da foto (se houver)
        doc.setFont("helvetica", "normal");
        doc.setFontSize(7.5);
        doc.setTextColor(30, 30, 30);
        const captionText = p.caption?.trim() || `Registro fotográfico #${globalIdx + 1}`;
        const wrappedCaption = doc.splitTextToSize(captionText, colW - 4);
        doc.text(wrappedCaption, cellX + 2, cellY + rowH - 3);
      }

      // Moldura da página do anexo
      const photosRows = Math.ceil(pagePhotos.length / 2);
      const photosBoxH = photosRows * (rowH + gapY) + 22;
      doc.setDrawColor(0);
      doc.setLineWidth(0.3);
      doc.rect(marginL, photoPageTop, contentW, Math.max(photosBoxH, pageH - 2 * photoPageTop));
    }
  }

  return doc.output("blob");
}

export async function downloadReportPdf(
  header: DailyHeader,
  events: WorkEvent[],
  photos?: ReportPhoto[],
  filename?: string,
) {
  const blob = await buildPdfBlobForReport(header, events, photos);
  const name = filename || generatePdfFilename(header, events);
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = name;
  a.click();
  URL.revokeObjectURL(url);
}
