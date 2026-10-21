import jsPDF from "jspdf";
import autoTable from "jspdf-autotable";
import { TIC_TRENS_LOGO_DATA_URL } from "./tic-trens-logo-data";
import { type EquipmentChecklist } from "./checklist";

function getImageDimensions(dataUrl: string): Promise<{ width: number; height: number }> {
  return new Promise((resolve) => {
    if (typeof window === "undefined") {
      resolve({ width: 4, height: 3 });
      return;
    }
    const img = new Image();
    img.onload = () => {
      resolve({
        width: img.naturalWidth || img.width || 4,
        height: img.naturalHeight || img.height || 3,
      });
    };
    img.onerror = () => {
      resolve({ width: 4, height: 3 });
    };
    img.src = dataUrl;
  });
}

export async function generateChecklistPdf(checklist: EquipmentChecklist): Promise<void> {
  const doc = new jsPDF({ unit: "mm", format: "a4" });
  const pageW = doc.internal.pageSize.getWidth();
  const pageH = doc.internal.pageSize.getHeight();
  const marginL = 12;
  const marginR = 12;
  const contentW = pageW - marginL - marginR;
  const boxTop = 10;

  // Outer border
  doc.setDrawColor(30, 41, 59);
  doc.setLineWidth(0.3);

  // Header Box
  const headerH = 26;
  doc.rect(marginL, boxTop, contentW, headerH);

  // Logo on Left
  const logoW = 38;
  const logoH = 18;
  try {
    doc.addImage(TIC_TRENS_LOGO_DATA_URL, "PNG", marginL + 3, boxTop + 4, logoW, logoH);
  } catch (err) {
    console.warn("Logo warning", err);
  }

  // Title in Center
  doc.setFont("helvetica", "bold");
  doc.setFontSize(12);
  doc.setTextColor(20, 30, 50);
  doc.text("VRF-SM PH-1016 • INSPEÇÃO PRÉ-OPERAÇÃO", marginL + logoW + 6, boxTop + 9);
  doc.setFontSize(8);
  doc.setFont("helvetica", "normal");
  doc.setTextColor(70, 80, 95);
  doc.text(
    "Checklist diário obrigatório de componentes antes do início da operação",
    marginL + logoW + 6,
    boxTop + 15,
  );
  doc.text(
    "TIC Trens • Operação de Veículos Rodoferroviários & Carro Ultrassom",
    marginL + logoW + 6,
    boxTop + 20,
  );

  // Status Badge on Right
  const statusColor =
    checklist.statusGeral === "apto"
      ? [22, 163, 74] // Green
      : checklist.statusGeral === "atencao"
        ? [217, 119, 6] // Amber
        : [220, 38, 38]; // Red

  const statusLabel =
    checklist.statusGeral === "apto"
      ? "APTO P/ OPERAÇÃO"
      : checklist.statusGeral === "atencao"
        ? "APTO C/ RESTRIÇÃO"
        : "IMPEDIDO / INAPTO";

  doc.setFillColor(statusColor[0], statusColor[1], statusColor[2]);
  const badgeW = 42;
  const badgeH = 9;
  const badgeX = pageW - marginR - badgeW - 3;
  const badgeY = boxTop + 8;
  doc.roundedRect(badgeX, badgeY, badgeW, badgeH, 2, 2, "F");
  doc.setTextColor(255, 255, 255);
  doc.setFont("helvetica", "bold");
  doc.setFontSize(8);
  doc.text(statusLabel, badgeX + badgeW / 2, badgeY + 5.8, { align: "center" });

  // Metadata Box
  const metaY = boxTop + headerH + 3;
  const metaH = 26;
  doc.rect(marginL, metaY, contentW, metaH);
  doc.setFillColor(248, 250, 252);
  doc.rect(marginL, metaY, contentW, 6, "F");
  doc.setFontSize(8);
  doc.setFont("helvetica", "bold");
  doc.setTextColor(30, 41, 59);
  doc.text("DADOS DO EQUIPAMENTO E OPERAÇÃO", marginL + 3, metaY + 4.2);

  // Metadata fields in 3 rows
  doc.setFont("helvetica", "normal");
  doc.setFontSize(7.5);
  doc.setTextColor(80, 90, 105);

  const row1Y = metaY + 10.5;
  const row2Y = metaY + 16;
  const row3Y = metaY + 21.5;

  const col1X = marginL + 3;
  const col2X = marginL + 48;
  const col3X = marginL + 96;
  const col4X = marginL + 142;

  // ROW 1
  doc.text(`Equipamento: `, col1X, row1Y);
  doc.setFont("helvetica", "bold");
  doc.setTextColor(15, 23, 42);
  doc.text(checklist.equipamento || "—", col1X + 19, row1Y);

  doc.setFont("helvetica", "normal");
  doc.setTextColor(80, 90, 105);
  doc.text(`Prefixo / Placa: `, col2X, row1Y);
  doc.setFont("helvetica", "bold");
  doc.setTextColor(15, 23, 42);
  doc.text(checklist.prefixo || "—", col2X + 21, row1Y);

  doc.setFont("helvetica", "normal");
  doc.setTextColor(80, 90, 105);
  doc.text(`Data: `, col3X, row1Y);
  doc.setFont("helvetica", "bold");
  doc.setTextColor(15, 23, 42);
  const dataFmt = checklist.data ? checklist.data.split("-").reverse().join("/") : "—";
  doc.text(`${dataFmt} às ${checklist.hora || ""}`, col3X + 9, row1Y);

  doc.setFont("helvetica", "normal");
  doc.setTextColor(80, 90, 105);
  doc.text(`Turno: `, col4X, row1Y);
  doc.setFont("helvetica", "bold");
  doc.setTextColor(15, 23, 42);
  doc.text(checklist.turno || "—", col4X + 10, row1Y);

  // ROW 2: Horímetros e Quilometragem
  doc.setFont("helvetica", "normal");
  doc.setTextColor(80, 90, 105);
  doc.text(`Horímetro Motor: `, col1X, row2Y);
  doc.setFont("helvetica", "bold");
  doc.setTextColor(15, 23, 42);
  doc.text(checklist.horimetroMotor || checklist.horimetroOuKm || "—", col1X + 24, row2Y);

  doc.setFont("helvetica", "normal");
  doc.setTextColor(80, 90, 105);
  doc.text(`H. Impl. Ferroviário: `, col2X, row2Y);
  doc.setFont("helvetica", "bold");
  doc.setTextColor(15, 23, 42);
  doc.text(checklist.horimetroImplementoFerroviario || "—", col2X + 27, row2Y);

  doc.setFont("helvetica", "normal");
  doc.setTextColor(80, 90, 105);
  doc.text(`H. Gerador: `, col3X, row2Y);
  doc.setFont("helvetica", "bold");
  doc.setTextColor(15, 23, 42);
  doc.text(checklist.horimetroGerador || "—", col3X + 17, row2Y);

  doc.setFont("helvetica", "normal");
  doc.setTextColor(80, 90, 105);
  doc.text(`KM Caminhão: `, col4X, row2Y);
  doc.setFont("helvetica", "bold");
  doc.setTextColor(15, 23, 42);
  doc.text(checklist.quilometragemCaminhao || "—", col4X + 20, row2Y);

  // ROW 3: Operador e Local
  doc.setFont("helvetica", "normal");
  doc.setTextColor(80, 90, 105);
  doc.text(`Operador / Inspetor: `, col1X, row3Y);
  doc.setFont("helvetica", "bold");
  doc.setTextColor(15, 23, 42);
  doc.text(checklist.operador || "—", col1X + 27, row3Y);

  doc.setFont("helvetica", "normal");
  doc.setTextColor(80, 90, 105);
  doc.text(`Local / Frente: `, col3X, row3Y);
  doc.setFont("helvetica", "bold");
  doc.setTextColor(15, 23, 42);
  doc.text(checklist.local || "—", col3X + 20, row3Y);

  // Table of items
  const tableData = (checklist.items || []).map((item, idx) => {
    let statusText = "Pendente";
    if (item.status === "conforme") statusText = "OK";
    else if (item.status === "nao_conforme") statusText = "NOK";
    else if (item.status === "nao_se_aplica") statusText = "N/A";

    let commentText = item.notes || "—";
    if (item.acaoSolucao && item.acaoSolucao.trim()) {
      commentText = item.notes
        ? `${item.notes} | Ação: ${item.acaoSolucao.trim()}`
        : `Ação: ${item.acaoSolucao.trim()}`;
    }

    return [String(idx + 1).padStart(2, "0"), item.name, statusText, commentText];
  });

  const tableStartY = metaY + metaH + 4;

  autoTable(doc, {
    startY: tableStartY,
    margin: { left: marginL, right: marginR },
    head: [["#", "Componentes Inspecionados (VRF-SM PH-1016)", "Status", "Comentários"]],
    body: tableData,
    theme: "grid",
    headStyles: {
      fillColor: [30, 41, 59],
      textColor: [255, 255, 255],
      fontSize: 7.5,
      fontStyle: "bold",
      halign: "left",
      cellPadding: 2,
    },
    bodyStyles: {
      fontSize: 7.5,
      textColor: [30, 41, 59],
      cellPadding: 1.8,
    },
    columnStyles: {
      0: { cellWidth: 8, halign: "center" },
      1: { cellWidth: 105 },
      2: { cellWidth: 20, halign: "center", fontStyle: "bold" },
      3: { cellWidth: "auto" },
    },
    didParseCell: (data) => {
      if (data.section === "body" && data.column.index === 2) {
        const val = String(data.cell.raw || "");
        if (val === "OK") {
          data.cell.styles.textColor = [22, 163, 74];
        } else if (val === "NOK") {
          data.cell.styles.textColor = [220, 38, 38];
        } else {
          data.cell.styles.textColor = [100, 116, 139];
        }
      }
    },
  });

  // Bottom Observation & Signatures
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const lastY = (doc as any).lastAutoTable?.finalY || 200;
  const remainingSpace = pageH - lastY;

  let currentY = lastY + 4;
  if (remainingSpace < 45) {
    doc.addPage();
    currentY = 15;
  }

  // General observations box
  doc.setDrawColor(180, 190, 205);
  doc.setFillColor(250, 250, 250);
  doc.rect(marginL, currentY, contentW, 16);
  doc.setFont("helvetica", "bold");
  doc.setFontSize(7.5);
  doc.setTextColor(50, 60, 75);
  doc.text("OBSERVAÇÕES GERAIS / RECOMENDAÇÕES:", marginL + 2, currentY + 4.5);
  doc.setFont("helvetica", "normal");
  doc.setTextColor(30, 41, 59);
  const splitObs = doc.splitTextToSize(
    checklist.observacoesGerais || "Nenhuma observação registrada.",
    contentW - 4,
  );
  doc.text(splitObs, marginL + 2, currentY + 9);

  // Signatures
  const sigY = currentY + 22;
  const sigW = 75;

  // Operator signature
  doc.line(marginL + 10, sigY + 12, marginL + 10 + sigW, sigY + 12);
  doc.setFontSize(7.5);
  doc.setFont("helvetica", "bold");
  doc.text("Assinatura do Operador / Inspetor", marginL + 10 + sigW / 2, sigY + 16, {
    align: "center",
  });
  doc.setFont("helvetica", "normal");
  doc.setFontSize(7);
  doc.text(checklist.operador || "Responsável pela inspeção", marginL + 10 + sigW / 2, sigY + 20, {
    align: "center",
  });

  // Supervisor signature
  const sig2X = pageW - marginR - sigW - 10;
  doc.line(sig2X, sigY + 12, sig2X + sigW, sigY + 12);
  doc.setFontSize(7.5);
  doc.setFont("helvetica", "bold");
  doc.text("Assinatura da Supervisão / Manutenção", sig2X + sigW / 2, sigY + 16, {
    align: "center",
  });
  doc.setFont("helvetica", "normal");
  doc.setFontSize(7);
  doc.text("Liberação operacional", sig2X + sigW / 2, sigY + 20, { align: "center" });

  // Anexo Fotográfico (se houver fotos anexadas)
  if (Array.isArray(checklist.fotos) && checklist.fotos.length > 0) {
    const photos = checklist.fotos;
    const photosPerPage = 4;
    const totalPhotoPages = Math.ceil(photos.length / photosPerPage);

    for (let pageIdx = 0; pageIdx < totalPhotoPages; pageIdx++) {
      doc.addPage();
      const photoPageTop = 10;
      let py = photoPageTop + 4;

      // Header do Anexo
      try {
        doc.addImage(TIC_TRENS_LOGO_DATA_URL, "PNG", marginL + 2, py, 38, 15);
      } catch {
        // ignore logo load error
      }

      doc.setFont("helvetica", "bold");
      doc.setFontSize(11);
      doc.setTextColor(20, 30, 50);
      const pageSuffix = totalPhotoPages > 1 ? ` (${pageIdx + 1}/${totalPhotoPages})` : "";
      doc.text(`ANEXO FOTOGRÁFICO - EVIDÊNCIAS DE INSPEÇÃO${pageSuffix}`, marginL + 46, py + 6);

      doc.setFont("helvetica", "normal");
      doc.setFontSize(7.5);
      doc.setTextColor(80, 90, 105);
      const subHeader = `Equipamento: ${checklist.equipamento || "—"}  |  Prefixo: ${checklist.prefixo || "—"}  |  Data: ${checklist.data || "—"}  |  Operador: ${checklist.operador || "—"}`;
      doc.text(subHeader, marginL + 46, py + 12);

      py += 18;
      doc.setDrawColor(200, 210, 225);
      doc.setLineWidth(0.3);
      doc.line(marginL, py, marginL + contentW, py);
      py += 6;

      const pagePhotos = photos.slice(pageIdx * photosPerPage, (pageIdx + 1) * photosPerPage);
      const colW = (contentW - 6) / 2;
      const rowH = 105;
      const gapX = 6;
      const gapY = 6;

      for (let i = 0; i < pagePhotos.length; i++) {
        const globalIdx = pageIdx * photosPerPage + i;
        const p = pagePhotos[i];
        const col = i % 2;
        const row = Math.floor(i / 2);

        const cellX = marginL + col * (colW + gapX);
        const cellY = py + row * (rowH + gapY);

        // Moldura do card
        doc.setDrawColor(200, 210, 225);
        doc.setFillColor(252, 253, 254);
        doc.rect(cellX, cellY, colW, rowH, "FD");

        // Banner superior
        doc.setFillColor(241, 245, 249);
        doc.rect(cellX, cellY, colW, 6.5, "F");
        doc.setFont("helvetica", "bold");
        doc.setFontSize(8);
        doc.setTextColor(30, 41, 59);
        doc.text(`EVIDÊNCIA #${String(globalIdx + 1).padStart(2, "0")}`, cellX + 3, cellY + 4.5);

        if (p.itemName) {
          doc.setFont("helvetica", "normal");
          doc.setFontSize(6.5);
          doc.setTextColor(71, 85, 105);
          const truncatedItem =
            p.itemName.length > 28 ? p.itemName.slice(0, 28) + "..." : p.itemName;
          doc.text(truncatedItem, cellX + colW - 3, cellY + 4.5, { align: "right" });
        }

        // Bounding box para a imagem sem distorção
        const boxX = cellX + 2;
        const boxY = cellY + 8;
        const boxW = colW - 4;
        const boxH = rowH - 22;

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
          console.warn("Erro ao renderizar foto no PDF do checklist:", e);
        }

        // Área da legenda
        const capY = cellY + rowH - 12;
        doc.setFillColor(248, 250, 252);
        doc.rect(cellX, capY, colW, 12, "F");
        doc.setFont("helvetica", "normal");
        doc.setFontSize(7);
        doc.setTextColor(51, 65, 85);
        const captionText =
          p.caption || "Registro fotográfico capturado durante inspeção pré-operação.";
        const splitCap = doc.splitTextToSize(captionText, colW - 4);
        doc.text(splitCap, cellX + 2, capY + 4.5);
      }
    }
  }

  const filename =
    `Checklist_${checklist.prefixo || checklist.equipamento || "Equipamento"}_${checklist.data || "Inspecao"}.pdf`.replace(
      /[^a-zA-Z0-9_.-]/g,
      "_",
    );
  doc.save(filename);
}
