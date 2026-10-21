import { Workbook, VbaProject } from "@node-projects/excelforge";
import { type WorkEvent, sortAndNormalizeEvents } from "./parte-diaria";

export interface SapCollaborator {
  nome: string;
  matricula: string;
  selecionado: boolean;
}

export const OFFICIAL_SAP_COLLABORATORS: Omit<SapCollaborator, "selecionado">[] = [
  { nome: "Alessandro Pires Costa", matricula: "10056322" },
  { nome: "Cássio Messias da Silva", matricula: "10056574" },
  { nome: "Elicarlos Oliveira Irmão", matricula: "10056353" },
  { nome: "Kayque Freitas Fagundes", matricula: "10056812" },
  { nome: "Leandro Donizete Fontana", matricula: "10059616" },
  { nome: "Moacir Borges Junior", matricula: "10056600" },
];

export interface SapInterval {
  id: string;
  atividade: string;
  dataInicio: string; // DD.MM.YYYY
  horaInicio: string; // HH:MM:SS
  dataFim: string; // DD.MM.YYYY
  horaFim: string; // HH:MM:SS
  eventoDesvio?: string;
}

/**
 * Retorna a data no formato DD-MM-YYYY (com traço), usado no nome do arquivo e na aba.
 */
export function formatSapDateDashed(dateOrIso?: string | number | Date): string {
  if (!dateOrIso) {
    const now = new Date();
    const d = String(now.getDate()).padStart(2, "0");
    const m = String(now.getMonth() + 1).padStart(2, "0");
    const y = now.getFullYear();
    return `${d}-${m}-${y}`;
  }
  if (typeof dateOrIso === "number") {
    const dObj = new Date(dateOrIso);
    const d = String(dObj.getDate()).padStart(2, "0");
    const m = String(dObj.getMonth() + 1).padStart(2, "0");
    const y = dObj.getFullYear();
    return `${d}-${m}-${y}`;
  }
  if (dateOrIso instanceof Date) {
    const d = String(dateOrIso.getDate()).padStart(2, "0");
    const m = String(dateOrIso.getMonth() + 1).padStart(2, "0");
    const y = dateOrIso.getFullYear();
    return `${d}-${m}-${y}`;
  }

  const clean = String(dateOrIso).trim();
  if (clean.includes("-")) {
    const parts = clean.split("T")[0].split("-");
    if (parts.length === 3) {
      if (parts[0].length === 4) {
        // YYYY-MM-DD -> DD-MM-YYYY
        const [y, m, d] = parts;
        return `${d.padStart(2, "0")}-${m.padStart(2, "0")}-${y}`;
      }
      // DD-MM-YYYY
      return clean.split("T")[0];
    }
  }
  if (clean.includes("/")) {
    const parts = clean.split("/");
    if (parts.length === 3) {
      const [d, m, y] = parts;
      return `${d.padStart(2, "0")}-${m.padStart(2, "0")}-${y}`;
    }
  }
  if (clean.includes(".")) {
    const parts = clean.split(".");
    if (parts.length === 3) {
      const [d, m, y] = parts;
      return `${d.padStart(2, "0")}-${m.padStart(2, "0")}-${y}`;
    }
  }
  return clean;
}

export function formatSapDate(dateOrIso?: string | number | Date): string {
  if (!dateOrIso) {
    const now = new Date();
    const d = String(now.getDate()).padStart(2, "0");
    const m = String(now.getMonth() + 1).padStart(2, "0");
    const y = now.getFullYear();
    return `${d}.${m}.${y}`;
  }
  if (typeof dateOrIso === "number") {
    const dObj = new Date(dateOrIso);
    const d = String(dObj.getDate()).padStart(2, "0");
    const m = String(dObj.getMonth() + 1).padStart(2, "0");
    const y = dObj.getFullYear();
    return `${d}.${m}.${y}`;
  }
  if (dateOrIso instanceof Date) {
    const d = String(dateOrIso.getDate()).padStart(2, "0");
    const m = String(dateOrIso.getMonth() + 1).padStart(2, "0");
    const y = dateOrIso.getFullYear();
    return `${d}.${m}.${y}`;
  }

  const clean = String(dateOrIso).trim();
  if (clean.includes("-")) {
    const parts = clean.split("T")[0].split("-");
    if (parts.length === 3) {
      if (parts[0].length === 4) {
        const [y, m, d] = parts;
        return `${d.padStart(2, "0")}.${m.padStart(2, "0")}.${y}`;
      }
      const [d, m, y] = parts;
      return `${d.padStart(2, "0")}.${m.padStart(2, "0")}.${y}`;
    }
  }
  if (clean.includes("/")) {
    const parts = clean.split("/");
    if (parts.length === 3) {
      const [d, m, y] = parts;
      return `${d.padStart(2, "0")}.${m.padStart(2, "0")}.${y}`;
    }
  }
  return clean;
}

export function formatSapTime(tsOrDate: number | Date): string {
  const d = typeof tsOrDate === "number" ? new Date(tsOrDate) : tsOrDate;
  const h = String(d.getHours()).padStart(2, "0");
  const m = String(d.getMinutes()).padStart(2, "0");
  const s = String(d.getSeconds()).padStart(2, "0");
  return `${h}:${m}:${s}`;
}

export function mapEventToSapActivity(kind: string, descricao?: string): string {
  const text = (descricao || "").toLowerCase();
  if (text.includes("prepar") || text.includes("manuten")) return "Preparação";
  if (text.includes("desloc") || text.includes("viagem")) return "deslocamento";
  if (text.includes("acesso") || text.includes("libera") || text.includes("cco"))
    return "acesso a via";
  if (text.includes("execut") || text.includes("inspe") || text.includes("ensaio"))
    return "executando atividade";
  if (text.includes("desmob") || text.includes("fim ativid")) return "desmobilização";
  if (text.includes("retorno") || text.includes("base")) return "retorno a base";

  switch (kind) {
    case "inicio_jornada":
    case "preparacao":
      return "Preparação";
    case "deslocamento_rodoviario":
    case "deslocamento_ferroviario":
    case "pronto_cco":
    case "pronto_patio":
      return "deslocamento";
    case "pedido_cco":
    case "liberacao_acesso":
    case "chegada_local":
      return "acesso a via";
    case "inicio_atividade":
      return "executando atividade";
    case "fim_atividade":
      return "desmobilização";
    case "retorno_base":
      return "retorno a base";
    case "refeicao":
      return "refeição";
    case "parada":
      return "parada";
    case "dimensionamento_defeito":
      return "dimensionamento defeito";
    default:
      return descricao || kind;
  }
}

/**
 * Converte eventos registrados em intervalos operacionais para apropriação IW41.
 * Se um intervalo cruza a meia-noite (23:59:59 -> 00:00:00), ele é dividido
 * automaticamente para garantir compatibilidade com o formato esperado pelo SAP.
 */
export function buildSapIntervalsFromEvents(
  events: WorkEvent[],
  defaultDateIso?: string,
): SapInterval[] {
  const sorted = sortAndNormalizeEvents(events || []);
  if (sorted.length === 0) {
    return createDefaultSapIntervals(defaultDateIso);
  }

  const intervals: SapInterval[] = [];

  for (let i = 0; i < sorted.length; i++) {
    const ev = sorted[i];
    const next = sorted[i + 1];

    if (ev.kind === "fim_jornada" && !next) {
      break;
    }

    const startTs = ev.timestamp;
    let endTs: number;

    if (next) {
      endTs = next.timestamp;
    } else {
      endTs = startTs + 30 * 60 * 1000;
    }

    const dStart = new Date(startTs);
    const dEnd = new Date(endTs);

    const atividade = mapEventToSapActivity(ev.kind, ev.descricao);
    const isDesvio = ev.kind === "parada" || ev.kind === "cancelado";
    const eventoDesvio = isDesvio ? ev.descricao || ev.observacoes || "Parada operacional" : "";

    const isSameDay =
      dStart.getFullYear() === dEnd.getFullYear() &&
      dStart.getMonth() === dEnd.getMonth() &&
      dStart.getDate() === dEnd.getDate();

    if (isSameDay) {
      intervals.push({
        id: `intv-${i}`,
        atividade,
        dataInicio: formatSapDate(dStart),
        horaInicio: formatSapTime(dStart),
        dataFim: formatSapDate(dEnd),
        horaFim: formatSapTime(dEnd),
        eventoDesvio,
      });
    } else {
      intervals.push({
        id: `intv-${i}-part1`,
        atividade,
        dataInicio: formatSapDate(dStart),
        horaInicio: formatSapTime(dStart),
        dataFim: formatSapDate(dStart),
        horaFim: "23:59:59",
        eventoDesvio,
      });
      intervals.push({
        id: `intv-${i}-part2`,
        atividade,
        dataInicio: formatSapDate(dEnd),
        horaInicio: "00:00:00",
        dataFim: formatSapDate(dEnd),
        horaFim: formatSapTime(dEnd),
        eventoDesvio,
      });
    }
  }

  return intervals.length > 0 ? intervals : createDefaultSapIntervals(defaultDateIso);
}

/**
 * Cria os 6 intervalos operacionais padrão demonstrados no modelo oficial:
 * Preparação, deslocamento, acesso a via, executando atividade, desmobilização, retorno a base.
 */
export function createDefaultSapIntervals(defaultDateIso?: string): SapInterval[] {
  const refDateStr = formatSapDate(defaultDateIso);

  let prevDateStr = refDateStr;
  if (defaultDateIso && defaultDateIso.includes("-")) {
    const parts = defaultDateIso.split("T")[0].split("-").map(Number);
    if (parts.length === 3) {
      const [y, m, d] = parts;
      const prev = new Date(y, m - 1, d - 1);
      prevDateStr = formatSapDate(prev);
    }
  } else {
    const today = new Date();
    const prev = new Date(today.getTime() - 24 * 60 * 60 * 1000);
    prevDateStr = formatSapDate(prev);
  }

  return [
    {
      id: "def-1",
      atividade: "Preparação",
      dataInicio: prevDateStr,
      horaInicio: "21:00:00",
      dataFim: prevDateStr,
      horaFim: "22:00:00",
    },
    {
      id: "def-2",
      atividade: "deslocamento",
      dataInicio: prevDateStr,
      horaInicio: "23:00:00",
      dataFim: prevDateStr,
      horaFim: "23:59:59",
    },
    {
      id: "def-3",
      atividade: "acesso a via",
      dataInicio: refDateStr,
      horaInicio: "00:00:00",
      dataFim: refDateStr,
      horaFim: "01:20:00",
    },
    {
      id: "def-4",
      atividade: "executando atividade",
      dataInicio: refDateStr,
      horaInicio: "01:20:00",
      dataFim: refDateStr,
      horaFim: "02:17:00",
    },
    {
      id: "def-5",
      atividade: "desmobilização",
      dataInicio: refDateStr,
      horaInicio: "02:17:00",
      dataFim: refDateStr,
      horaFim: "03:11:00",
    },
    {
      id: "def-6",
      atividade: "retorno a base",
      dataInicio: refDateStr,
      horaInicio: "03:11:00",
      dataFim: refDateStr,
      horaFim: "03:42:00",
    },
  ];
}

export interface GenerateSapIw41Options {
  om?: string;
  op?: string;
  reportDate?: string;
  collaborators: { nome: string; matricula: string }[];
  intervals: SapInterval[];
}

/**
 * Código VBA exato fornecido pelo usuário para automação SAP IW41.
 */
export const SAP_IW41_VBA_MACRO = [
  "Sub Macro_IW44_TIC()",
  "",
  "     Dim Appl, SapGuiAuto, Connection, session, WScript",
  "       ",
  "     Dim linha",
  "     linha = 2",
  "     ",
  '         Do Until Cells(linha, 2) = ""',
  "                ",
  "        If Not IsObject(Appl) Then",
  '           Set SapGuiAuto = GetObject("SAPGUI")',
  "           Set Appl = SapGuiAuto.GetScriptingEngine",
  "        End If",
  "        If Not IsObject(Connection) Then",
  "           Set Connection = Appl.Children(0)",
  "        End If",
  "        If Not IsObject(session) Then",
  "           Set session = Connection.Children(0)",
  "        End If",
  "        If IsObject(WScript) Then",
  '           WScript.ConnectObject session, "on"',
  '           WScript.ConnectObject Appl, "on"',
  "    ",
  '       WScript.ConnectObject Application, "on"',
  "       ",
  "    End If",
  "",
  'session.findById("wnd[0]").maximize',
  'session.findById("wnd[0]/tbar[0]/okcd").Text = "iw41"',
  'session.findById("wnd[0]").sendVKey 0',
  'session.findById("wnd[0]/usr/ctxtCORUF-AUFNR").Text = ActiveSheet.Range("b2")',
  'session.findById("wnd[0]/usr/txtCORUF-VORNR").Text = ActiveSheet.Range("c2")',
  'session.findById("wnd[0]/usr/txtCORUF-VORNR").SetFocus',
  'session.findById("wnd[0]/usr/txtCORUF-VORNR").caretPosition = 2',
  'session.findById("wnd[0]").sendVKey 0',
  'session.findById("wnd[0]/usr/chkAFRUD-AUERU").Selected = False',
  'session.findById("wnd[0]/usr/chkAFRUD-AUSOR").Selected = False',
  'session.findById("wnd[0]/usr/chkAFRUD-LEKNW").Selected = False',
  'session.findById("wnd[0]/usr/ctxtAFRUD-PERNR").Text = ActiveSheet.Range("d2")',
  'session.findById("wnd[0]/usr/ctxtAFRUD-ISDD").Text = ActiveSheet.Range("e2")',
  'session.findById("wnd[0]/usr/ctxtAFRUD-ISDZ").Text = ActiveSheet.Range("f2")',
  'session.findById("wnd[0]/usr/ctxtAFRUD-IEDD").Text = ActiveSheet.Range("g2")',
  'session.findById("wnd[0]/usr/ctxtAFRUD-IEDZ").Text = ActiveSheet.Range("h2")',
  'session.findById("wnd[0]/usr/ctxtAFRUD-IEDZ").SetFocus',
  'session.findById("wnd[0]/usr/ctxtAFRUD-IEDZ").caretPosition = 8',
  'session.findById("wnd[0]").sendVKey 0',
  'session.findById("wnd[0]/usr/chkAFRUD-AUERU").Selected = ActiveSheet.Range("z2")',
  'session.findById("wnd[0]/tbar[0]/btn[11]").press',
  'session.findById("wnd[0]/tbar[0]/btn[15]").press',
  "    ",
  '    Rows("2:2").Select',
  "    Selection.Delete Shift:=xlUp",
  "     Loop",
  "     ",
  '     MsgBox "HH LANÇADO COM SUCESSO  !!", vbInformation, "SAP IW41 - TIC Trens"',
  "     ",
  "End Sub",
  "",
  "Sub Macro_IW41_TIC()",
  "     Call Macro_IW44_TIC",
  "End Sub",
].join("\r\n");

/**
 * Código para ThisWorkbook que inicializa atalhos de teclado (Ctrl + Shift + R e Ctrl + Shift + S).
 */
export const THIS_WORKBOOK_VBA = [
  "Private Sub Workbook_Open()",
  "    On Error Resume Next",
  '    Application.OnKey "^+r", "Macro_IW44_TIC"',
  '    Application.OnKey "^+R", "Macro_IW44_TIC"',
  '    Application.OnKey "^+s", "Macro_IW44_TIC"',
  '    Application.OnKey "^+S", "Macro_IW44_TIC"',
  "End Sub",
].join("\r\n");

/**
 * Constrói a planilha XLSM habilitada para macro contendo os dados de apropriação,
 * o botão na planilha e a macro VBA embutida.
 */
export async function buildSapIw41Xlsm(options: GenerateSapIw41Options): Promise<{
  data: Uint8Array;
  filename: string;
}> {
  const { om = "", op = "10", reportDate, collaborators, intervals } = options;

  const dateDashed = formatSapDateDashed(reportDate);
  const dateHeader = formatSapDate(reportDate);
  const cleanOm = (om || "").trim();

  const wb = new Workbook();
  // Nome da aba no Excel conforme a data da inspeção separada por traço
  const ws = wb.addSheet(dateDashed);

  // Configuração das larguras das colunas
  ws.setColumnWidth(1, 24); // Atividade / Data
  ws.setColumnWidth(2, 18); // Ordem
  ws.setColumnWidth(3, 8); // OP
  ws.setColumnWidth(4, 14); // Matricula
  ws.setColumnWidth(5, 14); // DATA INICIO
  ws.setColumnWidth(6, 14); // HORA INICIO
  ws.setColumnWidth(7, 14); // DATA FIM
  ws.setColumnWidth(8, 14); // HORA FIM
  ws.setColumnWidth(9, 18); // Evento Desvio
  ws.setColumnWidth(10, 22); // Confirmação Final ?
  ws.setColumnWidth(11, 4); // Espaço separador
  ws.setColumnWidth(12, 34); // Painel da Macro / Atalho
  ws.setColumnWidth(13, 30);
  ws.setColumnWidth(26, 12); // Coluna Z (AUERU / Selected)

  // Cabeçalhos principais (Linha 1, 1-indexed)
  ws.setValue(1, 1, dateHeader);
  ws.setValue(1, 2, "Ordem");
  ws.setValue(1, 3, "OP");
  ws.setValue(1, 4, "Matricula");
  ws.setValue(1, 5, "DATA INICIO");
  ws.setValue(1, 6, "HORA INICIO");
  ws.setValue(1, 7, "DATA FIM");
  ws.setValue(1, 8, "HORA FIM");
  ws.setValue(1, 9, "Evento Desvio");
  ws.setValue(1, 10, "Confirmação Final ?");
  ws.setValue(1, 26, "AUERU");

  // Informações de instrução no topo direito (Colunas L a N)
  ws.setValue(1, 12, "AUTOMAÇÃO SAP IW41 (TIC TRENS):");
  ws.setValue(4, 12, "• Atalho de teclado configurado: Ctrl + Shift + R");
  ws.setValue(5, 12, "• Ou pressione Alt + F8 e execute 'Macro_IW44_TIC'");
  ws.setValue(6, 12, "• Certifique-se de que o SAP GUI está aberto e conectado");
  ws.setValue(7, 12, "• A macro consome as linhas automaticamente até a conclusão");

  // Inserção do Botão de Formulário para disparar a Macro diretamente no Excel
  ws.addFormControl({
    type: "button",
    text: "▶ EXECUTAR NO SAP (IW41) [Ctrl+Shift+R]",
    macro: "Macro_IW44_TIC",
    from: { col: 11, row: 1 }, // Coluna L, Linha 2 (0-indexed)
    width: 290,
    height: 40,
  });

  // Preenche as linhas de dados (iniciando na linha 2)
  let currentRow = 2;
  collaborators.forEach((colab, cIdx) => {
    const isLastCollaborator = cIdx === collaborators.length - 1;

    intervals.forEach((intv, iIdx) => {
      const isLastInterval = iIdx === intervals.length - 1;
      const isFinal = isLastCollaborator && isLastInterval;
      const confirmacaoFinal = isFinal ? "x" : "";

      ws.setValue(currentRow, 1, intv.atividade);
      // Se não tiver OM, deixa a coluna em branco para o usuário digitar na planilha
      ws.setValue(currentRow, 2, cleanOm || "");
      ws.setValue(currentRow, 3, op || "10");
      ws.setValue(currentRow, 4, colab.matricula);
      ws.setValue(currentRow, 5, intv.dataInicio);
      ws.setValue(currentRow, 6, intv.horaInicio);
      ws.setValue(currentRow, 7, intv.dataFim);
      ws.setValue(currentRow, 8, intv.horaFim);
      ws.setValue(currentRow, 9, intv.eventoDesvio || "");
      ws.setValue(currentRow, 10, confirmacaoFinal);
      // Coluna Z: valor booleano lido diretamente por ActiveSheet.Range("z2") no script SAP
      ws.setValue(currentRow, 26, isFinal);

      currentRow++;
    });
  });

  // Criação do Projeto VBA embutido
  const vba = new VbaProject();

  // Módulo padrão com a Macro
  vba.addModule({
    name: "Module1",
    type: "standard",
    code: SAP_IW41_VBA_MACRO,
  });

  // Módulo ThisWorkbook com ativação automática do atalho de teclado
  vba.addModule({
    name: "ThisWorkbook",
    type: "document",
    code: THIS_WORKBOOK_VBA,
  });

  wb.vbaProject = vba;

  // Compila o arquivo .xlsm em formato binário OpenXML habilitado para macro
  const data = await wb.build();
  const filename = `${dateDashed}.xlsm`;

  return { data, filename };
}

/**
 * Gera e dispara o download do arquivo .xlsm diretamente no navegador do usuário.
 */
export async function exportSapIw41ToExcel(options: GenerateSapIw41Options): Promise<string> {
  const { data, filename } = await buildSapIw41Xlsm(options);

  const blob = new Blob([data as unknown as BlobPart], {
    type: "application/vnd.ms-excel.sheet.macroEnabled.12",
  });

  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);

  return filename;
}
