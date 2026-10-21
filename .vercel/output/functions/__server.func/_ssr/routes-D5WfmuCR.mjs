import { o as __toESM } from "../_runtime.mjs";
import { t as supabase } from "./client-6krmjhG-.mjs";
import { u as require_react } from "../_libs/@floating-ui/react-dom+[...].mjs";
import { t as require_jsx_dev_runtime } from "../_libs/react.mjs";
import {
  a as setDoc,
  i as onSnapshot,
  n as getDoc,
  s as doc,
} from "../_libs/@firebase/firestore+[...].mjs";
import {
  A as LoaderCircle,
  B as FileCheck,
  C as Play,
  G as Database,
  H as ExternalLink,
  I as Funnel,
  J as CloudUpload,
  K as Copy,
  L as FileText,
  O as Maximize2,
  P as ImagePlus,
  Q as Circle,
  R as FileSpreadsheet,
  S as Plus,
  U as EllipsisVertical,
  V as Eye,
  W as Download,
  Y as Clock,
  Z as ClipboardCheck,
  _ as Search,
  a as WifiOff,
  at as Camera,
  ct as Briefcase,
  et as CircleCheck,
  h as ShieldCheck,
  i as Wifi,
  k as MapPin,
  m as Sparkles,
  n as X,
  nt as ChevronRight,
  o as Users,
  ot as Calendar,
  q as Cloud,
  rt as Check,
  s as User,
  st as Building2,
  t as Zap,
  tt as CircleAlert,
  u as Trash2,
  ut as ArrowRight,
  w as Pencil,
  x as RefreshCw,
} from "../_libs/lucide-react.mjs";
import {
  A as maskSinais,
  B as syncCloudReportsAcrossDevices,
  E as getLocalChecklists,
  I as saveReportToCloud,
  L as sortAndNormalizeEvents,
  O as maskGda,
  S as generatePdfFilename,
  T as getLocalCache,
  V as totals,
  _ as exportSingleReportToExcel,
  a as buildSingleReportExcelBlob,
  d as db,
  f as deduplicateCloudReports,
  g as exportAllReportsToExcel,
  h as emptyHeader,
  i as ThemeToggle,
  j as mergeReports,
  k as maskProgramadoExecutado,
  l as cloudReportToHeaderAndEvents,
  m as deleteCloudReport,
  n as EVENT_LABEL,
  o as buildWhatsAppMessage,
  r as TIC_TRENS_LOGO_DATA_URL,
  v as formatDate,
  x as formatTime,
  y as formatDuration,
  z as subscribeCloudReports,
} from "./excel-export-D9mJ2uBj.mjs";
import { h as Link } from "../_libs/@tanstack/react-router+[...].mjs";
import { t as require_jspdf_node_min } from "../_libs/jspdf.mjs";
import { t as autoTable } from "../_libs/jspdf-autotable.mjs";
import { t as clsx } from "../_libs/clsx.mjs";
import { n as Workbook, t as VbaProject } from "../_libs/node-projects__excelforge.mjs";
import {
  a as Label2,
  c as Root2,
  d as SubTrigger2,
  f as Trigger,
  i as ItemIndicator2,
  l as Separator2,
  n as Content2,
  o as Portal2,
  r as Item2,
  s as RadioItem2,
  t as CheckboxItem2,
  u as SubContent2,
} from "../_libs/@radix-ui/react-dropdown-menu+[...].mjs";
import { t as twMerge } from "../_libs/tailwind-merge.mjs";
//#region node_modules/.nitro/vite/services/ssr/assets/routes-D5WfmuCR.js
var import_react = /* @__PURE__ */ __toESM(require_react());
var import_jsx_dev_runtime = require_jsx_dev_runtime();
var import_jspdf_node_min = /* @__PURE__ */ __toESM(require_jspdf_node_min());
var OFFICIAL_SAP_COLLABORATORS = [
  {
    nome: "Alessandro Pires Costa",
    matricula: "10056322",
  },
  {
    nome: "Cássio Messias da Silva",
    matricula: "10056574",
  },
  {
    nome: "Elicarlos Oliveira Irmão",
    matricula: "10056353",
  },
  {
    nome: "Kayque Freitas Fagundes",
    matricula: "10056812",
  },
  {
    nome: "Leandro Donizete Fontana",
    matricula: "10059616",
  },
  {
    nome: "Moacir Borges Junior",
    matricula: "10056600",
  },
];
/**
 * Retorna a data no formato DD-MM-YYYY (com traço), usado no nome do arquivo e na aba.
 */
function formatSapDateDashed(dateOrIso) {
  if (!dateOrIso) {
    const now = /* @__PURE__ */ new Date();
    return `${String(now.getDate()).padStart(2, "0")}-${String(now.getMonth() + 1).padStart(2, "0")}-${now.getFullYear()}`;
  }
  if (typeof dateOrIso === "number") {
    const dObj = new Date(dateOrIso);
    return `${String(dObj.getDate()).padStart(2, "0")}-${String(dObj.getMonth() + 1).padStart(2, "0")}-${dObj.getFullYear()}`;
  }
  if (dateOrIso instanceof Date)
    return `${String(dateOrIso.getDate()).padStart(2, "0")}-${String(dateOrIso.getMonth() + 1).padStart(2, "0")}-${dateOrIso.getFullYear()}`;
  const clean = String(dateOrIso).trim();
  if (clean.includes("-")) {
    const parts = clean.split("T")[0].split("-");
    if (parts.length === 3) {
      if (parts[0].length === 4) {
        const [y, m, d] = parts;
        return `${d.padStart(2, "0")}-${m.padStart(2, "0")}-${y}`;
      }
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
function formatSapDate(dateOrIso) {
  if (!dateOrIso) {
    const now = /* @__PURE__ */ new Date();
    return `${String(now.getDate()).padStart(2, "0")}.${String(now.getMonth() + 1).padStart(2, "0")}.${now.getFullYear()}`;
  }
  if (typeof dateOrIso === "number") {
    const dObj = new Date(dateOrIso);
    return `${String(dObj.getDate()).padStart(2, "0")}.${String(dObj.getMonth() + 1).padStart(2, "0")}.${dObj.getFullYear()}`;
  }
  if (dateOrIso instanceof Date)
    return `${String(dateOrIso.getDate()).padStart(2, "0")}.${String(dateOrIso.getMonth() + 1).padStart(2, "0")}.${dateOrIso.getFullYear()}`;
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
function formatSapTime(tsOrDate) {
  const d = typeof tsOrDate === "number" ? new Date(tsOrDate) : tsOrDate;
  return `${String(d.getHours()).padStart(2, "0")}:${String(d.getMinutes()).padStart(2, "0")}:${String(d.getSeconds()).padStart(2, "0")}`;
}
function mapEventToSapActivity(kind, descricao) {
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
function buildSapIntervalsFromEvents(events, defaultDateIso) {
  const sorted = sortAndNormalizeEvents(events || []);
  if (sorted.length === 0) return createDefaultSapIntervals(defaultDateIso);
  const intervals = [];
  for (let i = 0; i < sorted.length; i++) {
    const ev = sorted[i];
    const next = sorted[i + 1];
    if (ev.kind === "fim_jornada" && !next) break;
    const startTs = ev.timestamp;
    let endTs;
    if (next) endTs = next.timestamp;
    else endTs = startTs + 1800 * 1e3;
    const dStart = new Date(startTs);
    const dEnd = new Date(endTs);
    const atividade = mapEventToSapActivity(ev.kind, ev.descricao);
    const eventoDesvio =
      ev.kind === "parada" || ev.kind === "cancelado"
        ? ev.descricao || ev.observacoes || "Parada operacional"
        : "";
    if (
      dStart.getFullYear() === dEnd.getFullYear() &&
      dStart.getMonth() === dEnd.getMonth() &&
      dStart.getDate() === dEnd.getDate()
    )
      intervals.push({
        id: `intv-${i}`,
        atividade,
        dataInicio: formatSapDate(dStart),
        horaInicio: formatSapTime(dStart),
        dataFim: formatSapDate(dEnd),
        horaFim: formatSapTime(dEnd),
        eventoDesvio,
      });
    else {
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
function createDefaultSapIntervals(defaultDateIso) {
  const refDateStr = formatSapDate(defaultDateIso);
  let prevDateStr = refDateStr;
  if (defaultDateIso && defaultDateIso.includes("-")) {
    const parts = defaultDateIso.split("T")[0].split("-").map(Number);
    if (parts.length === 3) {
      const [y, m, d] = parts;
      prevDateStr = formatSapDate(new Date(y, m - 1, d - 1));
    }
  } else
    prevDateStr = formatSapDate(
      /* @__PURE__ */ new Date(/* @__PURE__ */ new Date().getTime() - 1440 * 60 * 1e3),
    );
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
/**
 * Código VBA exato fornecido pelo usuário para automação SAP IW41.
 */
var SAP_IW41_VBA_MACRO = [
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
var THIS_WORKBOOK_VBA = [
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
async function buildSapIw41Xlsm(options) {
  const { om = "", op = "10", reportDate, collaborators, intervals } = options;
  const dateDashed = formatSapDateDashed(reportDate);
  const dateHeader = formatSapDate(reportDate);
  const cleanOm = (om || "").trim();
  const wb = new Workbook();
  const ws = wb.addSheet(dateDashed);
  ws.setColumnWidth(1, 24);
  ws.setColumnWidth(2, 18);
  ws.setColumnWidth(3, 8);
  ws.setColumnWidth(4, 14);
  ws.setColumnWidth(5, 14);
  ws.setColumnWidth(6, 14);
  ws.setColumnWidth(7, 14);
  ws.setColumnWidth(8, 14);
  ws.setColumnWidth(9, 18);
  ws.setColumnWidth(10, 22);
  ws.setColumnWidth(11, 4);
  ws.setColumnWidth(12, 34);
  ws.setColumnWidth(13, 30);
  ws.setColumnWidth(26, 12);
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
  ws.setValue(1, 12, "AUTOMAÇÃO SAP IW41 (TIC TRENS):");
  ws.setValue(4, 12, "• Atalho de teclado configurado: Ctrl + Shift + R");
  ws.setValue(5, 12, "• Ou pressione Alt + F8 e execute 'Macro_IW44_TIC'");
  ws.setValue(6, 12, "• Certifique-se de que o SAP GUI está aberto e conectado");
  ws.setValue(7, 12, "• A macro consome as linhas automaticamente até a conclusão");
  ws.addFormControl({
    type: "button",
    text: "▶ EXECUTAR NO SAP (IW41) [Ctrl+Shift+R]",
    macro: "Macro_IW44_TIC",
    from: {
      col: 11,
      row: 1,
    },
    width: 290,
    height: 40,
  });
  let currentRow = 2;
  collaborators.forEach((colab, cIdx) => {
    const isLastCollaborator = cIdx === collaborators.length - 1;
    intervals.forEach((intv, iIdx) => {
      const isLastInterval = iIdx === intervals.length - 1;
      const isFinal = isLastCollaborator && isLastInterval;
      const confirmacaoFinal = isFinal ? "x" : "";
      ws.setValue(currentRow, 1, intv.atividade);
      ws.setValue(currentRow, 2, cleanOm || "");
      ws.setValue(currentRow, 3, op || "10");
      ws.setValue(currentRow, 4, colab.matricula);
      ws.setValue(currentRow, 5, intv.dataInicio);
      ws.setValue(currentRow, 6, intv.horaInicio);
      ws.setValue(currentRow, 7, intv.dataFim);
      ws.setValue(currentRow, 8, intv.horaFim);
      ws.setValue(currentRow, 9, intv.eventoDesvio || "");
      ws.setValue(currentRow, 10, confirmacaoFinal);
      ws.setValue(currentRow, 26, isFinal);
      currentRow++;
    });
  });
  const vba = new VbaProject();
  vba.addModule({
    name: "Module1",
    type: "standard",
    code: SAP_IW41_VBA_MACRO,
  });
  vba.addModule({
    name: "ThisWorkbook",
    type: "document",
    code: THIS_WORKBOOK_VBA,
  });
  wb.vbaProject = vba;
  return {
    data: await wb.build(),
    filename: `${dateDashed}.xlsm`,
  };
}
/**
 * Gera e dispara o download do arquivo .xlsm diretamente no navegador do usuário.
 */
async function exportSapIw41ToExcel(options) {
  const { data, filename } = await buildSapIw41Xlsm(options);
  const blob = new Blob([data], { type: "application/vnd.ms-excel.sheet.macroEnabled.12" });
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
var _jsxFileName$5 = "/app/applet/src/components/SapIw41Modal.tsx";
function SapIw41Modal({ isOpen, onClose, header, events = [], onSuccess }) {
  const [om, setOm] = (0, import_react.useState)(header.om || "");
  const [op, setOp] = (0, import_react.useState)("10");
  const [isExporting, setIsExporting] = (0, import_react.useState)(false);
  const initialCollaborators = (0, import_react.useMemo)(() => {
    const rawColabs = (header.colaboradores || "").toLowerCase();
    const hasSpecificSelection = rawColabs.trim().length > 0;
    return OFFICIAL_SAP_COLLABORATORS.map((c) => {
      let isSelected = true;
      if (hasSpecificSelection) {
        const firstName = c.nome.split(" ")[0].toLowerCase();
        const lastName = c.nome.split(" ").slice(-1)[0].toLowerCase();
        isSelected =
          rawColabs.includes(c.nome.toLowerCase()) ||
          rawColabs.includes(firstName) ||
          rawColabs.includes(lastName);
      }
      return {
        ...c,
        selecionado: isSelected,
      };
    });
  }, [header.colaboradores]);
  const [collaborators, setCollaborators] = (0, import_react.useState)(initialCollaborators);
  const [newColabName, setNewColabName] = (0, import_react.useState)("");
  const [newColabMatricula, setNewColabMatricula] = (0, import_react.useState)("");
  const [intervals, setIntervals] = (0, import_react.useState)(() =>
    buildSapIntervalsFromEvents(events, header.data),
  );
  import_react.useEffect(() => {
    if (isOpen) {
      setOm(header.om || "");
      setOp("10");
      setCollaborators(initialCollaborators);
      setIntervals(buildSapIntervalsFromEvents(events, header.data));
      setIsExporting(false);
    }
  }, [isOpen, header.om, header.data, initialCollaborators, events]);
  if (!isOpen) return null;
  const toggleCollaborator = (matricula) => {
    setCollaborators((prev) =>
      prev.map((c) =>
        c.matricula === matricula
          ? {
              ...c,
              selecionado: !c.selecionado,
            }
          : c,
      ),
    );
  };
  const selectAllCollaborators = () => {
    setCollaborators((prev) =>
      prev.map((c) => ({
        ...c,
        selecionado: true,
      })),
    );
  };
  const deselectAllCollaborators = () => {
    setCollaborators((prev) =>
      prev.map((c) => ({
        ...c,
        selecionado: false,
      })),
    );
  };
  const handleAddCustomCollaborator = (e) => {
    e.preventDefault();
    const nome = newColabName.trim();
    const matricula = newColabMatricula.trim().replace(/\D/g, "");
    if (!nome || !matricula) return;
    setCollaborators((prev) => [
      ...prev,
      {
        nome,
        matricula,
        selecionado: true,
      },
    ]);
    setNewColabName("");
    setNewColabMatricula("");
  };
  const selectedCollaborators = collaborators.filter((c) => c.selecionado);
  const totalRows = selectedCollaborators.length * intervals.length;
  const expectedFilename = `${formatSapDateDashed(header.data)}.xlsm`;
  const handleExport = async () => {
    if (selectedCollaborators.length === 0) {
      alert("Selecione ao menos um colaborador para gerar a apropriação.");
      return;
    }
    if (intervals.length === 0) {
      alert("Não há intervalos de atividades para gerar.");
      return;
    }
    try {
      setIsExporting(true);
      const filename = await exportSapIw41ToExcel({
        om,
        op,
        reportDate: header.data,
        collaborators: selectedCollaborators,
        intervals,
      });
      if (onSuccess) onSuccess(filename);
      onClose();
    } catch (err) {
      console.error("Erro ao gerar arquivo Excel com Macro:", err);
      alert("Ocorreu um erro ao gerar a planilha com macro. Tente novamente.");
    } finally {
      setIsExporting(false);
    }
  };
  return /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
    "div",
    {
      className:
        "fixed inset-0 z-50 flex items-center justify-center p-3 sm:p-4 bg-black/60 backdrop-blur-xs animate-in fade-in duration-200",
      children: /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
        "div",
        {
          className:
            "bg-card text-card-foreground border border-border rounded-xl shadow-2xl w-full max-w-3xl max-h-[92vh] flex flex-col overflow-hidden",
          children: [
            /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
              "div",
              {
                className:
                  "flex items-center justify-between px-5 py-4 border-b border-border bg-muted/40",
                children: [
                  /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                    "div",
                    {
                      className: "flex items-center gap-3",
                      children: [
                        /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                          "div",
                          {
                            className:
                              "h-10 w-10 rounded-lg bg-emerald-600/15 text-emerald-600 dark:text-emerald-400 flex items-center justify-center font-bold",
                            children: /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                              FileSpreadsheet,
                              { className: "h-5 w-5" },
                              void 0,
                              false,
                              {
                                fileName: _jsxFileName$5,
                                lineNumber: 164,
                                columnNumber: 15,
                              },
                              this,
                            ),
                          },
                          void 0,
                          false,
                          {
                            fileName: _jsxFileName$5,
                            lineNumber: 163,
                            columnNumber: 13,
                          },
                          this,
                        ),
                        /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                          "div",
                          {
                            children: [
                              /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                "div",
                                {
                                  className: "flex items-center gap-2",
                                  children: [
                                    /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                      "h3",
                                      {
                                        className: "text-base font-bold text-foreground",
                                        children: "Apropriação de Mão de Obra — SAP IW41",
                                      },
                                      void 0,
                                      false,
                                      {
                                        fileName: _jsxFileName$5,
                                        lineNumber: 168,
                                        columnNumber: 17,
                                      },
                                      this,
                                    ),
                                    /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                      "span",
                                      {
                                        className:
                                          "bg-purple-600/10 text-purple-700 dark:text-purple-400 border border-purple-500/20 text-[10px] font-bold px-2 py-0.5 rounded-full uppercase tracking-wider flex items-center gap-1",
                                        children: [
                                          /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                            Play,
                                            { className: "h-2.5 w-2.5 fill-current" },
                                            void 0,
                                            false,
                                            {
                                              fileName: _jsxFileName$5,
                                              lineNumber: 172,
                                              columnNumber: 19,
                                            },
                                            this,
                                          ),
                                          "Excel .xlsm (Macro VBA)",
                                        ],
                                      },
                                      void 0,
                                      true,
                                      {
                                        fileName: _jsxFileName$5,
                                        lineNumber: 171,
                                        columnNumber: 17,
                                      },
                                      this,
                                    ),
                                  ],
                                },
                                void 0,
                                true,
                                {
                                  fileName: _jsxFileName$5,
                                  lineNumber: 167,
                                  columnNumber: 15,
                                },
                                this,
                              ),
                              /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                "p",
                                {
                                  className: "text-xs text-muted-foreground",
                                  children: [
                                    "Arquivo nomeado conforme a data da inspeção (",
                                    expectedFilename,
                                    ") com a macro e botão de execução no SAP",
                                  ],
                                },
                                void 0,
                                true,
                                {
                                  fileName: _jsxFileName$5,
                                  lineNumber: 176,
                                  columnNumber: 15,
                                },
                                this,
                              ),
                            ],
                          },
                          void 0,
                          true,
                          {
                            fileName: _jsxFileName$5,
                            lineNumber: 166,
                            columnNumber: 13,
                          },
                          this,
                        ),
                      ],
                    },
                    void 0,
                    true,
                    {
                      fileName: _jsxFileName$5,
                      lineNumber: 162,
                      columnNumber: 11,
                    },
                    this,
                  ),
                  /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                    "button",
                    {
                      type: "button",
                      onClick: onClose,
                      className:
                        "rounded-lg p-1.5 text-muted-foreground hover:bg-secondary hover:text-foreground transition-colors",
                      children: /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                        X,
                        { className: "h-5 w-5" },
                        void 0,
                        false,
                        {
                          fileName: _jsxFileName$5,
                          lineNumber: 187,
                          columnNumber: 13,
                        },
                        this,
                      ),
                    },
                    void 0,
                    false,
                    {
                      fileName: _jsxFileName$5,
                      lineNumber: 182,
                      columnNumber: 11,
                    },
                    this,
                  ),
                ],
              },
              void 0,
              true,
              {
                fileName: _jsxFileName$5,
                lineNumber: 161,
                columnNumber: 9,
              },
              this,
            ),
            /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
              "div",
              {
                className: "flex-1 overflow-y-auto p-5 space-y-5 text-sm",
                children: [
                  /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                    "div",
                    {
                      className:
                        "grid grid-cols-1 sm:grid-cols-3 gap-3 p-3.5 bg-muted/30 border border-border rounded-lg",
                      children: [
                        /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                          "div",
                          {
                            children: [
                              /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                "label",
                                {
                                  className:
                                    "text-xs font-semibold text-foreground flex items-center gap-1.5 mb-1",
                                  children: [
                                    /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                      Briefcase,
                                      { className: "h-3.5 w-3.5 text-primary" },
                                      void 0,
                                      false,
                                      {
                                        fileName: _jsxFileName$5,
                                        lineNumber: 197,
                                        columnNumber: 17,
                                      },
                                      this,
                                    ),
                                    /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                      "span",
                                      { children: "Ordem de Manutenção (OM)" },
                                      void 0,
                                      false,
                                      {
                                        fileName: _jsxFileName$5,
                                        lineNumber: 198,
                                        columnNumber: 17,
                                      },
                                      this,
                                    ),
                                  ],
                                },
                                void 0,
                                true,
                                {
                                  fileName: _jsxFileName$5,
                                  lineNumber: 196,
                                  columnNumber: 15,
                                },
                                this,
                              ),
                              /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                "input",
                                {
                                  type: "text",
                                  value: om,
                                  onChange: (e) => setOm(e.target.value),
                                  placeholder: "Deixar em branco para preencher no Excel",
                                  className:
                                    "w-full text-xs font-mono bg-background border border-input rounded-md px-2.5 py-1.5 focus:ring-1 focus:ring-primary outline-hidden",
                                },
                                void 0,
                                false,
                                {
                                  fileName: _jsxFileName$5,
                                  lineNumber: 200,
                                  columnNumber: 15,
                                },
                                this,
                              ),
                              /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                "span",
                                {
                                  className: "text-[10px] text-muted-foreground mt-0.5 block",
                                  children: om
                                    ? "Utilizando OM informada no cabeçalho"
                                    : "Coluna Ordem ficará vazia para preenchimento manual na planilha",
                                },
                                void 0,
                                false,
                                {
                                  fileName: _jsxFileName$5,
                                  lineNumber: 207,
                                  columnNumber: 15,
                                },
                                this,
                              ),
                            ],
                          },
                          void 0,
                          true,
                          {
                            fileName: _jsxFileName$5,
                            lineNumber: 195,
                            columnNumber: 13,
                          },
                          this,
                        ),
                        /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                          "div",
                          {
                            children: [
                              /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                "label",
                                {
                                  className:
                                    "text-xs font-semibold text-foreground flex items-center gap-1.5 mb-1",
                                  children: /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                    "span",
                                    { children: "Operação (OP)" },
                                    void 0,
                                    false,
                                    {
                                      fileName: _jsxFileName$5,
                                      lineNumber: 216,
                                      columnNumber: 17,
                                    },
                                    this,
                                  ),
                                },
                                void 0,
                                false,
                                {
                                  fileName: _jsxFileName$5,
                                  lineNumber: 215,
                                  columnNumber: 15,
                                },
                                this,
                              ),
                              /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                "input",
                                {
                                  type: "text",
                                  value: op,
                                  onChange: (e) => setOp(e.target.value),
                                  placeholder: "10",
                                  className:
                                    "w-full text-xs font-mono bg-background border border-input rounded-md px-2.5 py-1.5 focus:ring-1 focus:ring-primary outline-hidden",
                                },
                                void 0,
                                false,
                                {
                                  fileName: _jsxFileName$5,
                                  lineNumber: 218,
                                  columnNumber: 15,
                                },
                                this,
                              ),
                              /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                "span",
                                {
                                  className: "text-[10px] text-muted-foreground mt-0.5 block",
                                  children: "Padrão SAP: operação 10",
                                },
                                void 0,
                                false,
                                {
                                  fileName: _jsxFileName$5,
                                  lineNumber: 225,
                                  columnNumber: 15,
                                },
                                this,
                              ),
                            ],
                          },
                          void 0,
                          true,
                          {
                            fileName: _jsxFileName$5,
                            lineNumber: 214,
                            columnNumber: 13,
                          },
                          this,
                        ),
                        /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                          "div",
                          {
                            children: [
                              /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                "label",
                                {
                                  className:
                                    "text-xs font-semibold text-foreground flex items-center gap-1.5 mb-1",
                                  children: [
                                    /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                      Clock,
                                      { className: "h-3.5 w-3.5 text-primary" },
                                      void 0,
                                      false,
                                      {
                                        fileName: _jsxFileName$5,
                                        lineNumber: 232,
                                        columnNumber: 17,
                                      },
                                      this,
                                    ),
                                    /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                      "span",
                                      { children: "Nome do Arquivo & Aba" },
                                      void 0,
                                      false,
                                      {
                                        fileName: _jsxFileName$5,
                                        lineNumber: 233,
                                        columnNumber: 17,
                                      },
                                      this,
                                    ),
                                  ],
                                },
                                void 0,
                                true,
                                {
                                  fileName: _jsxFileName$5,
                                  lineNumber: 231,
                                  columnNumber: 15,
                                },
                                this,
                              ),
                              /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                "div",
                                {
                                  className:
                                    "w-full text-xs font-mono bg-muted/60 border border-border rounded-md px-2.5 py-1.5 text-foreground flex items-center justify-between",
                                  children: [
                                    /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                      "span",
                                      {
                                        className:
                                          "font-bold text-emerald-600 dark:text-emerald-400 truncate",
                                        children: expectedFilename,
                                      },
                                      void 0,
                                      false,
                                      {
                                        fileName: _jsxFileName$5,
                                        lineNumber: 236,
                                        columnNumber: 17,
                                      },
                                      this,
                                    ),
                                    /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                      "span",
                                      {
                                        className:
                                          "text-[10px] text-muted-foreground font-sans shrink-0 ml-1",
                                        children: "(com traço)",
                                      },
                                      void 0,
                                      false,
                                      {
                                        fileName: _jsxFileName$5,
                                        lineNumber: 239,
                                        columnNumber: 17,
                                      },
                                      this,
                                    ),
                                  ],
                                },
                                void 0,
                                true,
                                {
                                  fileName: _jsxFileName$5,
                                  lineNumber: 235,
                                  columnNumber: 15,
                                },
                                this,
                              ),
                              /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                "span",
                                {
                                  className: "text-[10px] text-muted-foreground mt-0.5 block",
                                  children: "Nome da planilha conforme a data da inspeção",
                                },
                                void 0,
                                false,
                                {
                                  fileName: _jsxFileName$5,
                                  lineNumber: 243,
                                  columnNumber: 15,
                                },
                                this,
                              ),
                            ],
                          },
                          void 0,
                          true,
                          {
                            fileName: _jsxFileName$5,
                            lineNumber: 230,
                            columnNumber: 13,
                          },
                          this,
                        ),
                      ],
                    },
                    void 0,
                    true,
                    {
                      fileName: _jsxFileName$5,
                      lineNumber: 194,
                      columnNumber: 11,
                    },
                    this,
                  ),
                  /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                    "div",
                    {
                      className:
                        "p-3.5 bg-purple-500/10 border border-purple-500/20 rounded-lg flex items-start gap-3",
                      children: [
                        /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                          "div",
                          {
                            className:
                              "h-8 w-8 rounded-lg bg-purple-600/20 text-purple-600 dark:text-purple-400 flex items-center justify-center shrink-0 mt-0.5",
                            children: /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                              Play,
                              { className: "h-4 w-4 fill-current ml-0.5" },
                              void 0,
                              false,
                              {
                                fileName: _jsxFileName$5,
                                lineNumber: 252,
                                columnNumber: 15,
                              },
                              this,
                            ),
                          },
                          void 0,
                          false,
                          {
                            fileName: _jsxFileName$5,
                            lineNumber: 251,
                            columnNumber: 13,
                          },
                          this,
                        ),
                        /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                          "div",
                          {
                            className: "text-xs space-y-1",
                            children: [
                              /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                "div",
                                {
                                  className: "flex items-center gap-2",
                                  children: [
                                    /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                      "span",
                                      {
                                        className: "font-bold text-foreground",
                                        children:
                                          "Macro SAP IW41 Embutida com Atalho e Botão na Planilha",
                                      },
                                      void 0,
                                      false,
                                      {
                                        fileName: _jsxFileName$5,
                                        lineNumber: 256,
                                        columnNumber: 17,
                                      },
                                      this,
                                    ),
                                    /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                      "span",
                                      {
                                        className:
                                          "inline-flex items-center gap-1 bg-purple-600/20 text-purple-700 dark:text-purple-300 font-mono text-[10px] px-1.5 py-0.5 rounded font-bold",
                                        children: "Ctrl + Shift + R",
                                      },
                                      void 0,
                                      false,
                                      {
                                        fileName: _jsxFileName$5,
                                        lineNumber: 259,
                                        columnNumber: 17,
                                      },
                                      this,
                                    ),
                                  ],
                                },
                                void 0,
                                true,
                                {
                                  fileName: _jsxFileName$5,
                                  lineNumber: 255,
                                  columnNumber: 15,
                                },
                                this,
                              ),
                              /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                "p",
                                {
                                  className: "text-muted-foreground leading-relaxed",
                                  children: [
                                    "A planilha gerada inclui a macro ",
                                    /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                      "strong",
                                      { children: "Macro_IW44_TIC" },
                                      void 0,
                                      false,
                                      {
                                        fileName: _jsxFileName$5,
                                        lineNumber: 264,
                                        columnNumber: 50,
                                      },
                                      this,
                                    ),
                                    " integrada, um botão interativo ",
                                    /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                      "strong",
                                      { children: "[▶ EXECUTAR NO SAP]" },
                                      void 0,
                                      false,
                                      {
                                        fileName: _jsxFileName$5,
                                        lineNumber: 265,
                                        columnNumber: 28,
                                      },
                                      this,
                                    ),
                                    " diretamente na planilha e atalho de teclado automático para conectar ao SAP GUI e lançar o HH sem esforço manual.",
                                  ],
                                },
                                void 0,
                                true,
                                {
                                  fileName: _jsxFileName$5,
                                  lineNumber: 263,
                                  columnNumber: 15,
                                },
                                this,
                              ),
                            ],
                          },
                          void 0,
                          true,
                          {
                            fileName: _jsxFileName$5,
                            lineNumber: 254,
                            columnNumber: 13,
                          },
                          this,
                        ),
                      ],
                    },
                    void 0,
                    true,
                    {
                      fileName: _jsxFileName$5,
                      lineNumber: 250,
                      columnNumber: 11,
                    },
                    this,
                  ),
                  /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                    "div",
                    {
                      className: "space-y-2",
                      children: [
                        /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                          "div",
                          {
                            className: "flex items-center justify-between",
                            children: [
                              /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                "div",
                                {
                                  className: "flex items-center gap-2",
                                  children: [
                                    /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                      Users,
                                      { className: "h-4 w-4 text-primary" },
                                      void 0,
                                      false,
                                      {
                                        fileName: _jsxFileName$5,
                                        lineNumber: 275,
                                        columnNumber: 17,
                                      },
                                      this,
                                    ),
                                    /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                      "h4",
                                      {
                                        className:
                                          "text-xs font-bold uppercase tracking-wider text-foreground",
                                        children: [
                                          "Colaboradores e Matrículas (",
                                          selectedCollaborators.length,
                                          " de",
                                          " ",
                                          collaborators.length,
                                          " selecionados)",
                                        ],
                                      },
                                      void 0,
                                      true,
                                      {
                                        fileName: _jsxFileName$5,
                                        lineNumber: 276,
                                        columnNumber: 17,
                                      },
                                      this,
                                    ),
                                  ],
                                },
                                void 0,
                                true,
                                {
                                  fileName: _jsxFileName$5,
                                  lineNumber: 274,
                                  columnNumber: 15,
                                },
                                this,
                              ),
                              /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                "div",
                                {
                                  className: "flex items-center gap-2 text-xs",
                                  children: [
                                    /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                      "button",
                                      {
                                        type: "button",
                                        onClick: selectAllCollaborators,
                                        className: "text-primary hover:underline font-medium",
                                        children: "Selecionar Todos",
                                      },
                                      void 0,
                                      false,
                                      {
                                        fileName: _jsxFileName$5,
                                        lineNumber: 282,
                                        columnNumber: 17,
                                      },
                                      this,
                                    ),
                                    /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                      "span",
                                      {
                                        className: "text-muted-foreground",
                                        children: "•",
                                      },
                                      void 0,
                                      false,
                                      {
                                        fileName: _jsxFileName$5,
                                        lineNumber: 289,
                                        columnNumber: 17,
                                      },
                                      this,
                                    ),
                                    /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                      "button",
                                      {
                                        type: "button",
                                        onClick: deselectAllCollaborators,
                                        className: "text-muted-foreground hover:underline",
                                        children: "Desmarcar Todos",
                                      },
                                      void 0,
                                      false,
                                      {
                                        fileName: _jsxFileName$5,
                                        lineNumber: 290,
                                        columnNumber: 17,
                                      },
                                      this,
                                    ),
                                  ],
                                },
                                void 0,
                                true,
                                {
                                  fileName: _jsxFileName$5,
                                  lineNumber: 281,
                                  columnNumber: 15,
                                },
                                this,
                              ),
                            ],
                          },
                          void 0,
                          true,
                          {
                            fileName: _jsxFileName$5,
                            lineNumber: 273,
                            columnNumber: 13,
                          },
                          this,
                        ),
                        /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                          "div",
                          {
                            className: "grid grid-cols-1 sm:grid-cols-2 gap-2",
                            children: collaborators.map((colab) =>
                              /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                "label",
                                {
                                  className: `flex items-center justify-between p-2.5 rounded-lg border text-xs cursor-pointer transition-colors ${colab.selecionado ? "bg-primary/5 border-primary/40 shadow-2xs" : "bg-background border-border hover:bg-muted/40 opacity-70"}`,
                                  children: [
                                    /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                      "div",
                                      {
                                        className: "flex items-center gap-2.5 min-w-0",
                                        children: [
                                          /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                            "input",
                                            {
                                              type: "checkbox",
                                              checked: colab.selecionado,
                                              onChange: () => toggleCollaborator(colab.matricula),
                                              className:
                                                "rounded border-input text-primary focus:ring-primary h-4 w-4",
                                            },
                                            void 0,
                                            false,
                                            {
                                              fileName: _jsxFileName$5,
                                              lineNumber: 311,
                                              columnNumber: 21,
                                            },
                                            this,
                                          ),
                                          /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                            "div",
                                            {
                                              className: "truncate",
                                              children: [
                                                /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                                  "p",
                                                  {
                                                    className:
                                                      "font-semibold text-foreground truncate",
                                                    children: colab.nome,
                                                  },
                                                  void 0,
                                                  false,
                                                  {
                                                    fileName: _jsxFileName$5,
                                                    lineNumber: 318,
                                                    columnNumber: 23,
                                                  },
                                                  this,
                                                ),
                                                /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                                  "p",
                                                  {
                                                    className:
                                                      "text-[11px] text-muted-foreground font-mono",
                                                    children: [
                                                      "Matrícula: ",
                                                      /* @__PURE__ */ (0,
                                                      import_jsx_dev_runtime.jsxDEV)(
                                                        "strong",
                                                        {
                                                          className: "text-foreground",
                                                          children: colab.matricula,
                                                        },
                                                        void 0,
                                                        false,
                                                        {
                                                          fileName: _jsxFileName$5,
                                                          lineNumber: 320,
                                                          columnNumber: 36,
                                                        },
                                                        this,
                                                      ),
                                                    ],
                                                  },
                                                  void 0,
                                                  true,
                                                  {
                                                    fileName: _jsxFileName$5,
                                                    lineNumber: 319,
                                                    columnNumber: 23,
                                                  },
                                                  this,
                                                ),
                                              ],
                                            },
                                            void 0,
                                            true,
                                            {
                                              fileName: _jsxFileName$5,
                                              lineNumber: 317,
                                              columnNumber: 21,
                                            },
                                            this,
                                          ),
                                        ],
                                      },
                                      void 0,
                                      true,
                                      {
                                        fileName: _jsxFileName$5,
                                        lineNumber: 310,
                                        columnNumber: 19,
                                      },
                                      this,
                                    ),
                                    colab.selecionado &&
                                      /* @__PURE__ */ (void 0)(
                                        Check,
                                        { className: "h-4 w-4 text-primary shrink-0" },
                                        void 0,
                                        false,
                                        {
                                          fileName: _jsxFileName$5,
                                          lineNumber: 324,
                                          columnNumber: 41,
                                        },
                                        this,
                                      ),
                                  ],
                                },
                                colab.matricula,
                                true,
                                {
                                  fileName: _jsxFileName$5,
                                  lineNumber: 302,
                                  columnNumber: 17,
                                },
                                this,
                              ),
                            ),
                          },
                          void 0,
                          false,
                          {
                            fileName: _jsxFileName$5,
                            lineNumber: 300,
                            columnNumber: 13,
                          },
                          this,
                        ),
                        /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                          "form",
                          {
                            onSubmit: handleAddCustomCollaborator,
                            className: "flex items-center gap-2 pt-1 text-xs",
                            children: [
                              /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                "input",
                                {
                                  type: "text",
                                  value: newColabName,
                                  onChange: (e) => setNewColabName(e.target.value),
                                  placeholder: "Adicionar outro colaborador...",
                                  className:
                                    "flex-1 bg-background border border-input rounded px-2.5 py-1 text-xs",
                                },
                                void 0,
                                false,
                                {
                                  fileName: _jsxFileName$5,
                                  lineNumber: 334,
                                  columnNumber: 15,
                                },
                                this,
                              ),
                              /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                "input",
                                {
                                  type: "text",
                                  value: newColabMatricula,
                                  onChange: (e) => setNewColabMatricula(e.target.value),
                                  placeholder: "Matrícula (ex: 10059999)",
                                  className:
                                    "w-40 bg-background border border-input rounded px-2.5 py-1 text-xs font-mono",
                                },
                                void 0,
                                false,
                                {
                                  fileName: _jsxFileName$5,
                                  lineNumber: 341,
                                  columnNumber: 15,
                                },
                                this,
                              ),
                              /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                "button",
                                {
                                  type: "submit",
                                  disabled: !newColabName.trim() || !newColabMatricula.trim(),
                                  className:
                                    "btn-secondary py-1 px-2.5 text-xs inline-flex items-center gap-1 disabled:opacity-40",
                                  children: [
                                    /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                      Plus,
                                      { className: "h-3 w-3" },
                                      void 0,
                                      false,
                                      {
                                        fileName: _jsxFileName$5,
                                        lineNumber: 353,
                                        columnNumber: 17,
                                      },
                                      this,
                                    ),
                                    /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                      "span",
                                      { children: "Adicionar" },
                                      void 0,
                                      false,
                                      {
                                        fileName: _jsxFileName$5,
                                        lineNumber: 354,
                                        columnNumber: 17,
                                      },
                                      this,
                                    ),
                                  ],
                                },
                                void 0,
                                true,
                                {
                                  fileName: _jsxFileName$5,
                                  lineNumber: 348,
                                  columnNumber: 15,
                                },
                                this,
                              ),
                            ],
                          },
                          void 0,
                          true,
                          {
                            fileName: _jsxFileName$5,
                            lineNumber: 330,
                            columnNumber: 13,
                          },
                          this,
                        ),
                      ],
                    },
                    void 0,
                    true,
                    {
                      fileName: _jsxFileName$5,
                      lineNumber: 272,
                      columnNumber: 11,
                    },
                    this,
                  ),
                  /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                    "div",
                    {
                      className: "space-y-2",
                      children: [
                        /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                          "div",
                          {
                            className: "flex items-center justify-between",
                            children: [
                              /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                "div",
                                {
                                  className: "flex items-center gap-2",
                                  children: [
                                    /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                      Clock,
                                      { className: "h-4 w-4 text-primary" },
                                      void 0,
                                      false,
                                      {
                                        fileName: _jsxFileName$5,
                                        lineNumber: 363,
                                        columnNumber: 17,
                                      },
                                      this,
                                    ),
                                    /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                      "h4",
                                      {
                                        className:
                                          "text-xs font-bold uppercase tracking-wider text-foreground",
                                        children: [
                                          "Etapas e Intervalos Operacionais (",
                                          intervals.length,
                                          " etapas)",
                                        ],
                                      },
                                      void 0,
                                      true,
                                      {
                                        fileName: _jsxFileName$5,
                                        lineNumber: 364,
                                        columnNumber: 17,
                                      },
                                      this,
                                    ),
                                  ],
                                },
                                void 0,
                                true,
                                {
                                  fileName: _jsxFileName$5,
                                  lineNumber: 362,
                                  columnNumber: 15,
                                },
                                this,
                              ),
                              /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                "span",
                                {
                                  className: "text-[11px] text-muted-foreground",
                                  children:
                                    events.length > 0
                                      ? "Calculado a partir dos eventos da linha do tempo"
                                      : "Sequência padrão do modelo SAP (6 etapas)",
                                },
                                void 0,
                                false,
                                {
                                  fileName: _jsxFileName$5,
                                  lineNumber: 368,
                                  columnNumber: 15,
                                },
                                this,
                              ),
                            ],
                          },
                          void 0,
                          true,
                          {
                            fileName: _jsxFileName$5,
                            lineNumber: 361,
                            columnNumber: 13,
                          },
                          this,
                        ),
                        /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                          "div",
                          {
                            className: "border border-border rounded-lg overflow-hidden",
                            children: /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                              "div",
                              {
                                className: "max-h-44 overflow-y-auto",
                                children: /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                  "table",
                                  {
                                    className: "w-full text-left text-xs",
                                    children: [
                                      /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                        "thead",
                                        {
                                          className:
                                            "bg-muted/80 text-muted-foreground font-semibold sticky top-0 border-b border-border",
                                          children: /* @__PURE__ */ (0,
                                          import_jsx_dev_runtime.jsxDEV)(
                                            "tr",
                                            {
                                              children: [
                                                /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                                  "th",
                                                  {
                                                    className: "py-2 px-3",
                                                    children: "Atividade",
                                                  },
                                                  void 0,
                                                  false,
                                                  {
                                                    fileName: _jsxFileName$5,
                                                    lineNumber: 380,
                                                    columnNumber: 23,
                                                  },
                                                  this,
                                                ),
                                                /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                                  "th",
                                                  {
                                                    className: "py-2 px-3",
                                                    children: "Início",
                                                  },
                                                  void 0,
                                                  false,
                                                  {
                                                    fileName: _jsxFileName$5,
                                                    lineNumber: 381,
                                                    columnNumber: 23,
                                                  },
                                                  this,
                                                ),
                                                /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                                  "th",
                                                  {
                                                    className: "py-2 px-3",
                                                    children: "Fim",
                                                  },
                                                  void 0,
                                                  false,
                                                  {
                                                    fileName: _jsxFileName$5,
                                                    lineNumber: 382,
                                                    columnNumber: 23,
                                                  },
                                                  this,
                                                ),
                                                /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                                  "th",
                                                  {
                                                    className: "py-2 px-3",
                                                    children: "Desvio",
                                                  },
                                                  void 0,
                                                  false,
                                                  {
                                                    fileName: _jsxFileName$5,
                                                    lineNumber: 383,
                                                    columnNumber: 23,
                                                  },
                                                  this,
                                                ),
                                              ],
                                            },
                                            void 0,
                                            true,
                                            {
                                              fileName: _jsxFileName$5,
                                              lineNumber: 379,
                                              columnNumber: 21,
                                            },
                                            this,
                                          ),
                                        },
                                        void 0,
                                        false,
                                        {
                                          fileName: _jsxFileName$5,
                                          lineNumber: 378,
                                          columnNumber: 19,
                                        },
                                        this,
                                      ),
                                      /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                        "tbody",
                                        {
                                          className: "divide-y divide-border font-mono text-[11px]",
                                          children: intervals.map((intv, idx) =>
                                            /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                              "tr",
                                              {
                                                className: "hover:bg-muted/20",
                                                children: [
                                                  /* @__PURE__ */ (0,
                                                  import_jsx_dev_runtime.jsxDEV)(
                                                    "td",
                                                    {
                                                      className:
                                                        "py-1.5 px-3 font-sans font-medium text-foreground",
                                                      children: intv.atividade,
                                                    },
                                                    void 0,
                                                    false,
                                                    {
                                                      fileName: _jsxFileName$5,
                                                      lineNumber: 389,
                                                      columnNumber: 25,
                                                    },
                                                    this,
                                                  ),
                                                  /* @__PURE__ */ (0,
                                                  import_jsx_dev_runtime.jsxDEV)(
                                                    "td",
                                                    {
                                                      className:
                                                        "py-1.5 px-3 text-muted-foreground",
                                                      children: [
                                                        intv.dataInicio,
                                                        " ",
                                                        intv.horaInicio,
                                                      ],
                                                    },
                                                    void 0,
                                                    true,
                                                    {
                                                      fileName: _jsxFileName$5,
                                                      lineNumber: 392,
                                                      columnNumber: 25,
                                                    },
                                                    this,
                                                  ),
                                                  /* @__PURE__ */ (0,
                                                  import_jsx_dev_runtime.jsxDEV)(
                                                    "td",
                                                    {
                                                      className:
                                                        "py-1.5 px-3 text-muted-foreground",
                                                      children: [intv.dataFim, " ", intv.horaFim],
                                                    },
                                                    void 0,
                                                    true,
                                                    {
                                                      fileName: _jsxFileName$5,
                                                      lineNumber: 395,
                                                      columnNumber: 25,
                                                    },
                                                    this,
                                                  ),
                                                  /* @__PURE__ */ (0,
                                                  import_jsx_dev_runtime.jsxDEV)(
                                                    "td",
                                                    {
                                                      className:
                                                        "py-1.5 px-3 text-amber-600 dark:text-amber-400 font-sans",
                                                      children: intv.eventoDesvio || "-",
                                                    },
                                                    void 0,
                                                    false,
                                                    {
                                                      fileName: _jsxFileName$5,
                                                      lineNumber: 398,
                                                      columnNumber: 25,
                                                    },
                                                    this,
                                                  ),
                                                ],
                                              },
                                              intv.id || idx,
                                              true,
                                              {
                                                fileName: _jsxFileName$5,
                                                lineNumber: 388,
                                                columnNumber: 23,
                                              },
                                              this,
                                            ),
                                          ),
                                        },
                                        void 0,
                                        false,
                                        {
                                          fileName: _jsxFileName$5,
                                          lineNumber: 386,
                                          columnNumber: 19,
                                        },
                                        this,
                                      ),
                                    ],
                                  },
                                  void 0,
                                  true,
                                  {
                                    fileName: _jsxFileName$5,
                                    lineNumber: 377,
                                    columnNumber: 17,
                                  },
                                  this,
                                ),
                              },
                              void 0,
                              false,
                              {
                                fileName: _jsxFileName$5,
                                lineNumber: 376,
                                columnNumber: 15,
                              },
                              this,
                            ),
                          },
                          void 0,
                          false,
                          {
                            fileName: _jsxFileName$5,
                            lineNumber: 375,
                            columnNumber: 13,
                          },
                          this,
                        ),
                      ],
                    },
                    void 0,
                    true,
                    {
                      fileName: _jsxFileName$5,
                      lineNumber: 360,
                      columnNumber: 11,
                    },
                    this,
                  ),
                  /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                    "div",
                    {
                      className:
                        "p-3 bg-emerald-500/10 border border-emerald-500/20 rounded-lg flex items-start gap-3",
                      children: [
                        /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                          FileCheck,
                          {
                            className:
                              "h-5 w-5 text-emerald-600 dark:text-emerald-400 shrink-0 mt-0.5",
                          },
                          void 0,
                          false,
                          {
                            fileName: _jsxFileName$5,
                            lineNumber: 411,
                            columnNumber: 13,
                          },
                          this,
                        ),
                        /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                          "div",
                          {
                            className: "text-xs space-y-1",
                            children: [
                              /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                "p",
                                {
                                  className: "font-semibold text-foreground",
                                  children: [
                                    "Total de ",
                                    totalRows,
                                    " linhas serão geradas no arquivo ",
                                    expectedFilename,
                                  ],
                                },
                                void 0,
                                true,
                                {
                                  fileName: _jsxFileName$5,
                                  lineNumber: 413,
                                  columnNumber: 15,
                                },
                                this,
                              ),
                              /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                "p",
                                {
                                  className: "text-muted-foreground",
                                  children: [
                                    "A planilha segue as colunas e formato da transação ",
                                    /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                      "strong",
                                      { children: "IW41" },
                                      void 0,
                                      false,
                                      {
                                        fileName: _jsxFileName$5,
                                        lineNumber: 417,
                                        columnNumber: 68,
                                      },
                                      this,
                                    ),
                                    ":",
                                    " ",
                                    formatSapDate(header.data),
                                    ", Ordem, OP, Matricula, DATA INICIO, HORA INICIO, DATA FIM, HORA FIM, Evento Desvio e Confirmação Final.",
                                  ],
                                },
                                void 0,
                                true,
                                {
                                  fileName: _jsxFileName$5,
                                  lineNumber: 416,
                                  columnNumber: 15,
                                },
                                this,
                              ),
                              /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                "p",
                                {
                                  className:
                                    "text-[11px] text-emerald-700 dark:text-emerald-300 font-medium",
                                  children:
                                    '✓ Coluna "Confirmação Final ?" conterá "x" na última linha e valor booleano configurado na coluna Z para a automação SAP.',
                                },
                                void 0,
                                false,
                                {
                                  fileName: _jsxFileName$5,
                                  lineNumber: 421,
                                  columnNumber: 15,
                                },
                                this,
                              ),
                            ],
                          },
                          void 0,
                          true,
                          {
                            fileName: _jsxFileName$5,
                            lineNumber: 412,
                            columnNumber: 13,
                          },
                          this,
                        ),
                      ],
                    },
                    void 0,
                    true,
                    {
                      fileName: _jsxFileName$5,
                      lineNumber: 410,
                      columnNumber: 11,
                    },
                    this,
                  ),
                ],
              },
              void 0,
              true,
              {
                fileName: _jsxFileName$5,
                lineNumber: 192,
                columnNumber: 9,
              },
              this,
            ),
            /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
              "div",
              {
                className:
                  "flex items-center justify-between px-5 py-4 border-t border-border bg-muted/20",
                children: [
                  /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                    "button",
                    {
                      type: "button",
                      onClick: onClose,
                      disabled: isExporting,
                      className:
                        "btn-ghost text-xs py-2 px-3 text-muted-foreground hover:text-foreground disabled:opacity-50",
                      children: "Cancelar",
                    },
                    void 0,
                    false,
                    {
                      fileName: _jsxFileName$5,
                      lineNumber: 431,
                      columnNumber: 11,
                    },
                    this,
                  ),
                  /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                    "button",
                    {
                      type: "button",
                      onClick: handleExport,
                      disabled:
                        selectedCollaborators.length === 0 || intervals.length === 0 || isExporting,
                      className:
                        "bg-emerald-600 hover:bg-emerald-700 text-white font-medium py-2 px-4 rounded-md transition-colors flex items-center gap-2 shadow-xs disabled:opacity-50 text-xs sm:text-sm",
                      children: isExporting
                        ? /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                            import_jsx_dev_runtime.Fragment,
                            {
                              children: [
                                /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                  LoaderCircle,
                                  { className: "h-4 w-4 animate-spin" },
                                  void 0,
                                  false,
                                  {
                                    fileName: _jsxFileName$5,
                                    lineNumber: 448,
                                    columnNumber: 17,
                                  },
                                  this,
                                ),
                                /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                  "span",
                                  { children: "Gerando .xlsm com Macro..." },
                                  void 0,
                                  false,
                                  {
                                    fileName: _jsxFileName$5,
                                    lineNumber: 449,
                                    columnNumber: 17,
                                  },
                                  this,
                                ),
                              ],
                            },
                            void 0,
                            true,
                            {
                              fileName: _jsxFileName$5,
                              lineNumber: 447,
                              columnNumber: 15,
                            },
                            this,
                          )
                        : /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                            import_jsx_dev_runtime.Fragment,
                            {
                              children: [
                                /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                  FileSpreadsheet,
                                  { className: "h-4 w-4" },
                                  void 0,
                                  false,
                                  {
                                    fileName: _jsxFileName$5,
                                    lineNumber: 453,
                                    columnNumber: 17,
                                  },
                                  this,
                                ),
                                /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                  "span",
                                  { children: ["Gerar Planilha IW41 (", expectedFilename, ")"] },
                                  void 0,
                                  true,
                                  {
                                    fileName: _jsxFileName$5,
                                    lineNumber: 454,
                                    columnNumber: 17,
                                  },
                                  this,
                                ),
                              ],
                            },
                            void 0,
                            true,
                            {
                              fileName: _jsxFileName$5,
                              lineNumber: 452,
                              columnNumber: 15,
                            },
                            this,
                          ),
                    },
                    void 0,
                    false,
                    {
                      fileName: _jsxFileName$5,
                      lineNumber: 440,
                      columnNumber: 11,
                    },
                    this,
                  ),
                ],
              },
              void 0,
              true,
              {
                fileName: _jsxFileName$5,
                lineNumber: 430,
                columnNumber: 9,
              },
              this,
            ),
          ],
        },
        void 0,
        true,
        {
          fileName: _jsxFileName$5,
          lineNumber: 159,
          columnNumber: 7,
        },
        this,
      ),
    },
    void 0,
    false,
    {
      fileName: _jsxFileName$5,
      lineNumber: 158,
      columnNumber: 5,
    },
    this,
  );
}
var SUPABASE_PROJECT_ID = "ivzutgehaxmvoreypssm";
`${SUPABASE_PROJECT_ID}`;
var SUPABASE_MIGRATION_SQL = `-- ==========================================
-- SCRIPT DE MIGRAÇÃO: RELATÓRIOS E CHECKLISTS
-- Execute este script no SQL Editor do Supabase:
-- https://supabase.com/dashboard/project/${SUPABASE_PROJECT_ID}/sql/new
-- ==========================================

-- 1. Criação da tabela de Relatórios
CREATE TABLE IF NOT EXISTS public.relatorios (
  id TEXT PRIMARY KEY,
  data TEXT NOT NULL,
  nome TEXT NOT NULL DEFAULT '',
  local TEXT DEFAULT '',
  sinais TEXT DEFAULT '',
  linha TEXT DEFAULT '',
  via TEXT DEFAULT '',
  gda TEXT DEFAULT '',
  om TEXT DEFAULT '',
  empresa TEXT DEFAULT '',
  atividade TEXT DEFAULT '',
  colaboradores TEXT DEFAULT '',
  descritivo TEXT DEFAULT '',
  "kmProgramado" TEXT DEFAULT '',
  executado TEXT DEFAULT '',
  "observacaoGeral" TEXT DEFAULT '',
  telefone TEXT DEFAULT '',
  events JSONB DEFAULT '[]'::jsonb,
  photos JSONB DEFAULT '[]'::jsonb,
  totals JSONB DEFAULT '{}'::jsonb,
  "createdAt" BIGINT NOT NULL,
  "updatedAt" BIGINT NOT NULL,
  created_at TIMESTAMPTZ DEFAULT now(),
  updated_at TIMESTAMPTZ DEFAULT now()
);

-- 2. Criação da tabela de Checklists de Equipamentos
CREATE TABLE IF NOT EXISTS public.checklists_equipamento (
  id TEXT PRIMARY KEY,
  data TEXT NOT NULL,
  hora TEXT DEFAULT '',
  equipamento TEXT NOT NULL DEFAULT '',
  prefixo TEXT DEFAULT '',
  "horimetroMotor" TEXT DEFAULT '',
  "horimetroImplementoFerroviario" TEXT DEFAULT '',
  "horimetroGerador" TEXT DEFAULT '',
  "quilometragemCaminhao" TEXT DEFAULT '',
  "horimetroOuKm" TEXT DEFAULT '',
  operador TEXT NOT NULL DEFAULT '',
  turno TEXT DEFAULT '',
  local TEXT DEFAULT '',
  "statusGeral" TEXT DEFAULT 'apto',
  "observacoesGerais" TEXT DEFAULT '',
  items JSONB DEFAULT '[]'::jsonb,
  "fotosCount" INTEGER DEFAULT 0,
  "createdAt" BIGINT NOT NULL,
  "updatedAt" BIGINT NOT NULL,
  created_at TIMESTAMPTZ DEFAULT now(),
  updated_at TIMESTAMPTZ DEFAULT now()
);

-- 3. Habilita RLS (Row Level Security)
ALTER TABLE public.relatorios ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.checklists_equipamento ENABLE ROW LEVEL SECURITY;

-- 4. Cria Políticas de Acesso Público para leitura e escrita
DROP POLICY IF EXISTS "Allow all on relatorios" ON public.relatorios;
CREATE POLICY "Allow all on relatorios" ON public.relatorios
  FOR ALL
  TO public
  USING (true)
  WITH CHECK (true);

DROP POLICY IF EXISTS "Allow all on checklists_equipamento" ON public.checklists_equipamento;
CREATE POLICY "Allow all on checklists_equipamento" ON public.checklists_equipamento
  FOR ALL
  TO public
  USING (true)
  WITH CHECK (true);

-- 5. Concede permissões aos papéis
GRANT ALL ON public.relatorios TO anon, authenticated, service_role;
GRANT ALL ON public.checklists_equipamento TO anon, authenticated, service_role;

-- 6. Habilita Replicação em Tempo Real (Realtime)
DO $$
BEGIN
  IF NOT EXISTS (
    SELECT 1 FROM pg_publication_tables 
    WHERE pubname = 'supabase_realtime' AND tablename = 'relatorios'
  ) THEN
    ALTER PUBLICATION supabase_realtime ADD TABLE public.relatorios;
  END IF;

  IF NOT EXISTS (
    SELECT 1 FROM pg_publication_tables 
    WHERE pubname = 'supabase_realtime' AND tablename = 'checklists_equipamento'
  ) THEN
    ALTER PUBLICATION supabase_realtime ADD TABLE public.checklists_equipamento;
  END IF;
END $$;
`;
var statusListeners = /* @__PURE__ */ new Set();
function notifyStatus(status) {
  statusListeners.forEach((fn) => {
    try {
      fn(status);
    } catch (e) {
      console.warn("Migration listener err:", e);
    }
  });
}
/**
 * Verifica se as tabelas relatorios e checklists_equipamento existem no Supabase
 */
async function checkSupabaseTables() {
  let relatoriosReady = false;
  let checklistsReady = false;
  let reportsCount = 0;
  let checklistsCount = 0;
  let errMsg = null;
  try {
    const { count, error } = await supabase.from("relatorios").select("id", {
      count: "exact",
      head: true,
    });
    if (!error) {
      relatoriosReady = true;
      reportsCount = count || 0;
    } else errMsg = error.message;
  } catch (err) {
    errMsg = err instanceof Error ? err.message : String(err);
  }
  try {
    const { count, error } = await supabase.from("checklists_equipamento").select("id", {
      count: "exact",
      head: true,
    });
    if (!error) {
      checklistsReady = true;
      checklistsCount = count || 0;
    } else if (!errMsg) errMsg = error.message;
  } catch (err) {
    if (!errMsg) errMsg = err instanceof Error ? err.message : String(err);
  }
  const status = {
    tablesReady: relatoriosReady && checklistsReady,
    relatoriosTableReady: relatoriosReady,
    checklistsTableReady: checklistsReady,
    reportsCountInSupabase: reportsCount,
    checklistsCountInSupabase: checklistsCount,
    lastChecked: Date.now(),
    error: errMsg,
  };
  notifyStatus(status);
  return status;
}
/**
 * Converte um SavedCloudReport para o formato de inserção na tabela relatorios do Supabase
 */
function formatReportForSupabase(report) {
  return {
    id: report.id,
    data: report.data || /* @__PURE__ */ new Date().toISOString().split("T")[0],
    nome: report.nome || "",
    local: report.local || "",
    sinais: report.sinais || "",
    linha: report.linha || "",
    via: report.via || "",
    gda: report.gda || "",
    om: report.om || "",
    empresa: report.empresa || "",
    atividade: report.atividade || "",
    colaboradores: report.colaboradores || "",
    descritivo: report.descritivo || "",
    kmProgramado: report.kmProgramado || "",
    executado: report.executado || "",
    observacaoGeral: report.observacaoGeral || "",
    telefone: report.telefone || "",
    events: report.events || [],
    photos: report.photos || [],
    totals: report.totals || {
      jornada: 0,
      atividade: 0,
      preparacao: 0,
      parada: 0,
      producaoTotal: 0,
      defeitosTotal: 0,
    },
    createdAt: Number(report.createdAt || Date.now()),
    updatedAt: Number(report.updatedAt || Date.now()),
  };
}
/**
 * Converte um EquipmentChecklist para o formato de inserção na tabela checklists_equipamento
 */
function formatChecklistForSupabase(chk) {
  return {
    id: chk.id,
    data: chk.data || /* @__PURE__ */ new Date().toISOString().split("T")[0],
    hora: chk.hora || "",
    equipamento: chk.equipamento || "",
    prefixo: chk.prefixo || "",
    horimetroMotor: chk.horimetroMotor || "",
    horimetroImplementoFerroviario: chk.horimetroImplementoFerroviario || "",
    horimetroGerador: chk.horimetroGerador || "",
    quilometragemCaminhao: chk.quilometragemCaminhao || "",
    horimetroOuKm: chk.horimetroOuKm || "",
    operador: chk.operador || "",
    turno: chk.turno || "",
    local: chk.local || "",
    statusGeral: chk.statusGeral || "apto",
    observacoesGerais: chk.observacoesGerais || "",
    items: chk.items || [],
    fotosCount: chk.fotosCount || 0,
    createdAt: Number(chk.createdAt || Date.now()),
    updatedAt: Number(chk.updatedAt || Date.now()),
  };
}
/**
 * Executa a migração de todos os dados locais/servidor para o Supabase
 */
async function migrateAllDataToSupabase(options) {
  const errors = [];
  let migratedReports = 0;
  let migratedChecklists = 0;
  if (!(await checkSupabaseTables()).tablesReady)
    return {
      success: false,
      migratedReports: 0,
      migratedChecklists: 0,
      errors: [
        "As tabelas 'relatorios' e/ou 'checklists_equipamento' ainda não foram criadas no Supabase. Por favor, execute o script SQL no SQL Editor do Supabase antes de migrar.",
      ],
    };
  let allReports = getLocalCache();
  try {
    const res = await fetch("/api/cloud-reports");
    if (res.ok) {
      const serverReports = await res.json();
      if (Array.isArray(serverReports) && serverReports.length > 0)
        allReports = mergeReports(serverReports, allReports);
    }
  } catch (e) {
    console.warn("Could not fetch server reports during migration:", e);
  }
  if (allReports.length > 0) {
    options?.onProgress?.({
      stage: "Migrando relatórios diários...",
      current: 0,
      total: allReports.length,
    });
    for (let i = 0; i < allReports.length; i += 10) {
      const batch = allReports.slice(i, i + 10).map(formatReportForSupabase);
      const { error } = await supabase.from("relatorios").upsert(batch, { onConflict: "id" });
      if (error)
        errors.push(
          `Erro ao migrar lote de relatórios (${i}-${i + batch.length}): ${error.message}`,
        );
      else migratedReports += batch.length;
      options?.onProgress?.({
        stage: "Migrando relatórios diários...",
        current: Math.min(i + 10, allReports.length),
        total: allReports.length,
      });
    }
  }
  const localChecklists = getLocalChecklists();
  if (localChecklists.length > 0) {
    options?.onProgress?.({
      stage: "Migrando checklists de equipamento...",
      current: 0,
      total: localChecklists.length,
    });
    for (let i = 0; i < localChecklists.length; i += 10) {
      const batch = localChecklists.slice(i, i + 10).map(formatChecklistForSupabase);
      const { error } = await supabase
        .from("checklists_equipamento")
        .upsert(batch, { onConflict: "id" });
      if (error)
        errors.push(
          `Erro ao migrar lote de checklists (${i}-${i + batch.length}): ${error.message}`,
        );
      else migratedChecklists += batch.length;
      options?.onProgress?.({
        stage: "Migrando checklists de equipamento...",
        current: Math.min(i + 10, localChecklists.length),
        total: localChecklists.length,
      });
    }
  }
  await checkSupabaseTables();
  return {
    success: errors.length === 0,
    migratedReports,
    migratedChecklists,
    errors,
  };
}
var _jsxFileName$4 = "/app/applet/src/components/SupabaseMigrationModal.tsx";
function SupabaseMigrationModal({ isOpen, onClose, onMigrationComplete }) {
  const [isChecking, setIsChecking] = (0, import_react.useState)(false);
  const [isMigrating, setIsMigrating] = (0, import_react.useState)(false);
  const [copiedSql, setCopiedSql] = (0, import_react.useState)(false);
  const [tableStatus, setTableStatus] = (0, import_react.useState)(null);
  const [progress, setProgress] = (0, import_react.useState)(null);
  const [errorMsg, setErrorMsg] = (0, import_react.useState)(null);
  const checkStatus = async () => {
    setIsChecking(true);
    setErrorMsg(null);
    try {
      const status = await checkSupabaseTables();
      setTableStatus(status);
    } catch (e) {
      console.error("Error checking Supabase tables:", e);
      setErrorMsg("Erro ao verificar tabelas do Supabase. Verifique sua conexão.");
    } finally {
      setIsChecking(false);
    }
  };
  (0, import_react.useEffect)(() => {
    if (isOpen) checkStatus();
  }, [isOpen]);
  if (!isOpen) return null;
  const handleCopySql = () => {
    navigator.clipboard.writeText(SUPABASE_MIGRATION_SQL);
    setCopiedSql(true);
    setTimeout(() => setCopiedSql(false), 2500);
  };
  const handleStartMigration = async () => {
    setIsMigrating(true);
    setErrorMsg(null);
    try {
      const result = await migrateAllDataToSupabase((p) => {
        setProgress(p);
      });
      if (result.success) onMigrationComplete?.();
      else if (result.error) setErrorMsg(result.error);
    } catch (e) {
      console.error("Migration error:", e);
      const msg =
        e instanceof Error ? e.message : "Ocorreu um erro durante a migração para o Supabase.";
      setErrorMsg(msg);
    } finally {
      setIsMigrating(false);
    }
  };
  return /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
    "div",
    {
      id: "supabase-migration-modal-backdrop",
      className:
        "fixed inset-0 z-50 flex items-center justify-center p-4 bg-black/60 backdrop-blur-xs animate-in fade-in duration-200",
      children: /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
        "div",
        {
          id: "supabase-migration-modal-card",
          className:
            "relative w-full max-w-2xl bg-card border border-border text-card-foreground rounded-2xl shadow-2xl flex flex-col max-h-[90vh] overflow-hidden",
          children: [
            /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
              "div",
              {
                className:
                  "flex items-center justify-between px-6 py-4 border-b border-border bg-muted/30",
                children: [
                  /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                    "div",
                    {
                      className: "flex items-center gap-2.5",
                      children: [
                        /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                          "div",
                          {
                            className:
                              "p-2 rounded-xl bg-emerald-500/10 text-emerald-600 dark:text-emerald-400",
                            children: /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                              Database,
                              { className: "w-5 h-5" },
                              void 0,
                              false,
                              {
                                fileName: _jsxFileName$4,
                                lineNumber: 106,
                                columnNumber: 15,
                              },
                              this,
                            ),
                          },
                          void 0,
                          false,
                          {
                            fileName: _jsxFileName$4,
                            lineNumber: 105,
                            columnNumber: 13,
                          },
                          this,
                        ),
                        /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                          "div",
                          {
                            children: [
                              /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                "h2",
                                {
                                  className:
                                    "text-base font-semibold text-foreground flex items-center gap-2",
                                  children: [
                                    "Migração para Banco de Dados Supabase",
                                    /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                      "span",
                                      {
                                        className:
                                          "text-[11px] font-normal px-2 py-0.5 rounded-full bg-emerald-500/10 text-emerald-700 dark:text-emerald-300 border border-emerald-500/20",
                                        children: "Sem limite de cota",
                                      },
                                      void 0,
                                      false,
                                      {
                                        fileName: _jsxFileName$4,
                                        lineNumber: 111,
                                        columnNumber: 17,
                                      },
                                      this,
                                    ),
                                  ],
                                },
                                void 0,
                                true,
                                {
                                  fileName: _jsxFileName$4,
                                  lineNumber: 109,
                                  columnNumber: 15,
                                },
                                this,
                              ),
                              /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                "p",
                                {
                                  className: "text-xs text-muted-foreground",
                                  children:
                                    "Transfira todos os relatórios e checklists do Firebase para o Supabase",
                                },
                                void 0,
                                false,
                                {
                                  fileName: _jsxFileName$4,
                                  lineNumber: 115,
                                  columnNumber: 15,
                                },
                                this,
                              ),
                            ],
                          },
                          void 0,
                          true,
                          {
                            fileName: _jsxFileName$4,
                            lineNumber: 108,
                            columnNumber: 13,
                          },
                          this,
                        ),
                      ],
                    },
                    void 0,
                    true,
                    {
                      fileName: _jsxFileName$4,
                      lineNumber: 104,
                      columnNumber: 11,
                    },
                    this,
                  ),
                  /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                    "button",
                    {
                      id: "close-supabase-modal-btn",
                      onClick: onClose,
                      className:
                        "p-1.5 text-muted-foreground hover:text-foreground rounded-lg hover:bg-muted transition-colors",
                      children: /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                        X,
                        { className: "w-5 h-5" },
                        void 0,
                        false,
                        {
                          fileName: _jsxFileName$4,
                          lineNumber: 125,
                          columnNumber: 13,
                        },
                        this,
                      ),
                    },
                    void 0,
                    false,
                    {
                      fileName: _jsxFileName$4,
                      lineNumber: 120,
                      columnNumber: 11,
                    },
                    this,
                  ),
                ],
              },
              void 0,
              true,
              {
                fileName: _jsxFileName$4,
                lineNumber: 103,
                columnNumber: 9,
              },
              this,
            ),
            /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
              "div",
              {
                className: "px-6 py-5 overflow-y-auto space-y-5 text-xs text-foreground",
                children: [
                  /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                    "div",
                    {
                      className: "p-3.5 rounded-xl border bg-muted/40 flex items-start gap-3",
                      children: [
                        /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                          ShieldCheck,
                          { className: "w-5 h-5 text-emerald-600 shrink-0 mt-0.5" },
                          void 0,
                          false,
                          {
                            fileName: _jsxFileName$4,
                            lineNumber: 133,
                            columnNumber: 13,
                          },
                          this,
                        ),
                        /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                          "div",
                          {
                            className: "space-y-1",
                            children: [
                              /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                "div",
                                {
                                  className:
                                    "font-semibold text-foreground flex items-center gap-2",
                                  children: [
                                    "Projeto Supabase Conectado",
                                    /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                      "code",
                                      {
                                        className:
                                          "text-[10px] px-1.5 py-0.5 rounded bg-muted border font-mono",
                                        children: "ivzutgehaxmvoreypssm",
                                      },
                                      void 0,
                                      false,
                                      {
                                        fileName: _jsxFileName$4,
                                        lineNumber: 137,
                                        columnNumber: 17,
                                      },
                                      this,
                                    ),
                                  ],
                                },
                                void 0,
                                true,
                                {
                                  fileName: _jsxFileName$4,
                                  lineNumber: 135,
                                  columnNumber: 15,
                                },
                                this,
                              ),
                              /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                "p",
                                {
                                  className: "text-muted-foreground leading-relaxed",
                                  children:
                                    "Ao migrar para o Supabase, seu aplicativo ganha sincronização instantânea em tempo real, sem bloqueios de cota diária de leitura do Firebase, funcionando perfeitamente em todos os seus dispositivos.",
                                },
                                void 0,
                                false,
                                {
                                  fileName: _jsxFileName$4,
                                  lineNumber: 141,
                                  columnNumber: 15,
                                },
                                this,
                              ),
                            ],
                          },
                          void 0,
                          true,
                          {
                            fileName: _jsxFileName$4,
                            lineNumber: 134,
                            columnNumber: 13,
                          },
                          this,
                        ),
                      ],
                    },
                    void 0,
                    true,
                    {
                      fileName: _jsxFileName$4,
                      lineNumber: 132,
                      columnNumber: 11,
                    },
                    this,
                  ),
                  /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                    "div",
                    {
                      className: "border border-border rounded-xl p-4 space-y-3 bg-card",
                      children: [
                        /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                          "div",
                          {
                            className: "flex items-center justify-between",
                            children: [
                              /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                "span",
                                {
                                  className:
                                    "font-semibold text-foreground text-xs uppercase tracking-wider",
                                  children: "1. Status das Tabelas no Supabase",
                                },
                                void 0,
                                false,
                                {
                                  fileName: _jsxFileName$4,
                                  lineNumber: 152,
                                  columnNumber: 15,
                                },
                                this,
                              ),
                              /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                "button",
                                {
                                  id: "check-supabase-tables-btn",
                                  onClick: checkStatus,
                                  disabled: isChecking,
                                  className:
                                    "flex items-center gap-1.5 text-xs text-primary hover:underline disabled:opacity-50",
                                  children: [
                                    /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                      RefreshCw,
                                      {
                                        className: `w-3.5 h-3.5 ${isChecking ? "animate-spin" : ""}`,
                                      },
                                      void 0,
                                      false,
                                      {
                                        fileName: _jsxFileName$4,
                                        lineNumber: 161,
                                        columnNumber: 17,
                                      },
                                      this,
                                    ),
                                    isChecking ? "Verificando..." : "Verificar Tabelas",
                                  ],
                                },
                                void 0,
                                true,
                                {
                                  fileName: _jsxFileName$4,
                                  lineNumber: 155,
                                  columnNumber: 15,
                                },
                                this,
                              ),
                            ],
                          },
                          void 0,
                          true,
                          {
                            fileName: _jsxFileName$4,
                            lineNumber: 151,
                            columnNumber: 13,
                          },
                          this,
                        ),
                        /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                          "div",
                          {
                            className: "grid grid-cols-1 sm:grid-cols-2 gap-2.5 pt-1",
                            children: [
                              /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                "div",
                                {
                                  className: `p-3 rounded-lg border flex items-center justify-between ${tableStatus?.relatoriosReady ? "bg-emerald-500/10 border-emerald-500/30 text-emerald-800 dark:text-emerald-300" : "bg-amber-500/10 border-amber-500/30 text-amber-800 dark:text-amber-300"}`,
                                  children: [
                                    /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                      "span",
                                      {
                                        className: "font-medium",
                                        children: "Tabela 'relatorios'",
                                      },
                                      void 0,
                                      false,
                                      {
                                        fileName: _jsxFileName$4,
                                        lineNumber: 174,
                                        columnNumber: 17,
                                      },
                                      this,
                                    ),
                                    tableStatus?.relatoriosReady
                                      ? /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                          "span",
                                          {
                                            className:
                                              "flex items-center gap-1 text-[11px] font-semibold text-emerald-600",
                                            children: [
                                              /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                                CircleCheck,
                                                { className: "w-4 h-4" },
                                                void 0,
                                                false,
                                                {
                                                  fileName: _jsxFileName$4,
                                                  lineNumber: 177,
                                                  columnNumber: 21,
                                                },
                                                this,
                                              ),
                                              " Pronta",
                                            ],
                                          },
                                          void 0,
                                          true,
                                          {
                                            fileName: _jsxFileName$4,
                                            lineNumber: 176,
                                            columnNumber: 19,
                                          },
                                          this,
                                        )
                                      : /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                          "span",
                                          {
                                            className:
                                              "flex items-center gap-1 text-[11px] font-semibold text-amber-600",
                                            children: [
                                              /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                                CircleAlert,
                                                { className: "w-4 h-4" },
                                                void 0,
                                                false,
                                                {
                                                  fileName: _jsxFileName$4,
                                                  lineNumber: 181,
                                                  columnNumber: 21,
                                                },
                                                this,
                                              ),
                                              " Pendente",
                                            ],
                                          },
                                          void 0,
                                          true,
                                          {
                                            fileName: _jsxFileName$4,
                                            lineNumber: 180,
                                            columnNumber: 19,
                                          },
                                          this,
                                        ),
                                  ],
                                },
                                void 0,
                                true,
                                {
                                  fileName: _jsxFileName$4,
                                  lineNumber: 167,
                                  columnNumber: 15,
                                },
                                this,
                              ),
                              /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                "div",
                                {
                                  className: `p-3 rounded-lg border flex items-center justify-between ${tableStatus?.checklistsReady ? "bg-emerald-500/10 border-emerald-500/30 text-emerald-800 dark:text-emerald-300" : "bg-amber-500/10 border-amber-500/30 text-amber-800 dark:text-amber-300"}`,
                                  children: [
                                    /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                      "span",
                                      {
                                        className: "font-medium",
                                        children: "Tabela 'checklists_equipamento'",
                                      },
                                      void 0,
                                      false,
                                      {
                                        fileName: _jsxFileName$4,
                                        lineNumber: 193,
                                        columnNumber: 17,
                                      },
                                      this,
                                    ),
                                    tableStatus?.checklistsReady
                                      ? /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                          "span",
                                          {
                                            className:
                                              "flex items-center gap-1 text-[11px] font-semibold text-emerald-600",
                                            children: [
                                              /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                                CircleCheck,
                                                { className: "w-4 h-4" },
                                                void 0,
                                                false,
                                                {
                                                  fileName: _jsxFileName$4,
                                                  lineNumber: 196,
                                                  columnNumber: 21,
                                                },
                                                this,
                                              ),
                                              " Pronta",
                                            ],
                                          },
                                          void 0,
                                          true,
                                          {
                                            fileName: _jsxFileName$4,
                                            lineNumber: 195,
                                            columnNumber: 19,
                                          },
                                          this,
                                        )
                                      : /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                          "span",
                                          {
                                            className:
                                              "flex items-center gap-1 text-[11px] font-semibold text-amber-600",
                                            children: [
                                              /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                                CircleAlert,
                                                { className: "w-4 h-4" },
                                                void 0,
                                                false,
                                                {
                                                  fileName: _jsxFileName$4,
                                                  lineNumber: 200,
                                                  columnNumber: 21,
                                                },
                                                this,
                                              ),
                                              " Pendente",
                                            ],
                                          },
                                          void 0,
                                          true,
                                          {
                                            fileName: _jsxFileName$4,
                                            lineNumber: 199,
                                            columnNumber: 19,
                                          },
                                          this,
                                        ),
                                  ],
                                },
                                void 0,
                                true,
                                {
                                  fileName: _jsxFileName$4,
                                  lineNumber: 186,
                                  columnNumber: 15,
                                },
                                this,
                              ),
                            ],
                          },
                          void 0,
                          true,
                          {
                            fileName: _jsxFileName$4,
                            lineNumber: 166,
                            columnNumber: 13,
                          },
                          this,
                        ),
                        (!tableStatus?.tablesReady || !tableStatus) &&
                          /* @__PURE__ */ (void 0)(
                            "div",
                            {
                              className: "pt-2 border-t border-border mt-3 space-y-2",
                              children: [
                                /* @__PURE__ */ (void 0)(
                                  "p",
                                  {
                                    className: "text-muted-foreground leading-relaxed",
                                    children: [
                                      "Para criar as tabelas no Supabase com segurança, copie o script SQL abaixo e cole no ",
                                      /* @__PURE__ */ (void 0)(
                                        "strong",
                                        { children: "SQL Editor" },
                                        void 0,
                                        false,
                                        {
                                          fileName: _jsxFileName$4,
                                          lineNumber: 211,
                                          columnNumber: 22,
                                        },
                                        this,
                                      ),
                                      " do seu painel do Supabase:",
                                    ],
                                  },
                                  void 0,
                                  true,
                                  {
                                    fileName: _jsxFileName$4,
                                    lineNumber: 209,
                                    columnNumber: 17,
                                  },
                                  this,
                                ),
                                /* @__PURE__ */ (void 0)(
                                  "div",
                                  {
                                    className: "flex flex-wrap items-center gap-2",
                                    children: [
                                      /* @__PURE__ */ (void 0)(
                                        "button",
                                        {
                                          id: "copy-supabase-sql-btn",
                                          onClick: handleCopySql,
                                          className:
                                            "flex items-center gap-1.5 px-3 py-1.5 rounded-lg bg-primary text-primary-foreground text-xs font-medium hover:bg-primary/90 transition-colors shadow-xs",
                                          children: copiedSql
                                            ? /* @__PURE__ */ (void 0)(
                                                import_jsx_dev_runtime.Fragment,
                                                {
                                                  children: [
                                                    /* @__PURE__ */ (void 0)(
                                                      Check,
                                                      { className: "w-3.5 h-3.5" },
                                                      void 0,
                                                      false,
                                                      {
                                                        fileName: _jsxFileName$4,
                                                        lineNumber: 221,
                                                        columnNumber: 25,
                                                      },
                                                      this,
                                                    ),
                                                    " SQL Copiado!",
                                                  ],
                                                },
                                                void 0,
                                                true,
                                                {
                                                  fileName: _jsxFileName$4,
                                                  lineNumber: 220,
                                                  columnNumber: 23,
                                                },
                                                this,
                                              )
                                            : /* @__PURE__ */ (void 0)(
                                                import_jsx_dev_runtime.Fragment,
                                                {
                                                  children: [
                                                    /* @__PURE__ */ (void 0)(
                                                      Copy,
                                                      { className: "w-3.5 h-3.5" },
                                                      void 0,
                                                      false,
                                                      {
                                                        fileName: _jsxFileName$4,
                                                        lineNumber: 225,
                                                        columnNumber: 25,
                                                      },
                                                      this,
                                                    ),
                                                    " Copiar Código SQL",
                                                  ],
                                                },
                                                void 0,
                                                true,
                                                {
                                                  fileName: _jsxFileName$4,
                                                  lineNumber: 224,
                                                  columnNumber: 23,
                                                },
                                                this,
                                              ),
                                        },
                                        void 0,
                                        false,
                                        {
                                          fileName: _jsxFileName$4,
                                          lineNumber: 214,
                                          columnNumber: 19,
                                        },
                                        this,
                                      ),
                                      /* @__PURE__ */ (void 0)(
                                        "a",
                                        {
                                          id: "open-supabase-dashboard-link",
                                          href: "https://supabase.com/dashboard/project/ivzutgehaxmvoreypssm/sql/new",
                                          target: "_blank",
                                          rel: "noopener noreferrer",
                                          className:
                                            "flex items-center gap-1 px-3 py-1.5 rounded-lg border border-border hover:bg-muted text-xs transition-colors",
                                          children: [
                                            "Abrir SQL Editor no Supabase ",
                                            /* @__PURE__ */ (void 0)(
                                              ExternalLink,
                                              { className: "w-3.5 h-3.5 ml-0.5" },
                                              void 0,
                                              false,
                                              {
                                                fileName: _jsxFileName$4,
                                                lineNumber: 236,
                                                columnNumber: 50,
                                              },
                                              this,
                                            ),
                                          ],
                                        },
                                        void 0,
                                        true,
                                        {
                                          fileName: _jsxFileName$4,
                                          lineNumber: 229,
                                          columnNumber: 19,
                                        },
                                        this,
                                      ),
                                    ],
                                  },
                                  void 0,
                                  true,
                                  {
                                    fileName: _jsxFileName$4,
                                    lineNumber: 213,
                                    columnNumber: 17,
                                  },
                                  this,
                                ),
                              ],
                            },
                            void 0,
                            true,
                            {
                              fileName: _jsxFileName$4,
                              lineNumber: 208,
                              columnNumber: 15,
                            },
                            this,
                          ),
                      ],
                    },
                    void 0,
                    true,
                    {
                      fileName: _jsxFileName$4,
                      lineNumber: 150,
                      columnNumber: 11,
                    },
                    this,
                  ),
                  /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                    "div",
                    {
                      className: "border border-border rounded-xl p-4 space-y-3 bg-card",
                      children: [
                        /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                          "span",
                          {
                            className:
                              "font-semibold text-foreground text-xs uppercase tracking-wider block",
                            children: "2. Transferência de Dados",
                          },
                          void 0,
                          false,
                          {
                            fileName: _jsxFileName$4,
                            lineNumber: 245,
                            columnNumber: 13,
                          },
                          this,
                        ),
                        /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                          "p",
                          {
                            className: "text-muted-foreground",
                            children:
                              "Esta ação coletará todos os relatórios diários de obra e checklists armazenados no seu dispositivo e servidor, formatando-os e gravando-os em lote no Supabase.",
                          },
                          void 0,
                          false,
                          {
                            fileName: _jsxFileName$4,
                            lineNumber: 248,
                            columnNumber: 13,
                          },
                          this,
                        ),
                        progress &&
                          /* @__PURE__ */ (void 0)(
                            "div",
                            {
                              className:
                                "p-3.5 rounded-lg bg-muted/60 border border-border space-y-2",
                              children: [
                                /* @__PURE__ */ (void 0)(
                                  "div",
                                  {
                                    className: "flex items-center justify-between text-xs",
                                    children: [
                                      /* @__PURE__ */ (void 0)(
                                        "span",
                                        {
                                          className: "font-medium text-foreground",
                                          children: progress.message,
                                        },
                                        void 0,
                                        false,
                                        {
                                          fileName: _jsxFileName$4,
                                          lineNumber: 256,
                                          columnNumber: 19,
                                        },
                                        this,
                                      ),
                                      /* @__PURE__ */ (void 0)(
                                        "span",
                                        {
                                          className: "text-muted-foreground",
                                          children:
                                            progress.phase === "completed"
                                              ? "Concluído"
                                              : "Em andamento...",
                                        },
                                        void 0,
                                        false,
                                        {
                                          fileName: _jsxFileName$4,
                                          lineNumber: 257,
                                          columnNumber: 19,
                                        },
                                        this,
                                      ),
                                    ],
                                  },
                                  void 0,
                                  true,
                                  {
                                    fileName: _jsxFileName$4,
                                    lineNumber: 255,
                                    columnNumber: 17,
                                  },
                                  this,
                                ),
                                /* @__PURE__ */ (void 0)(
                                  "div",
                                  {
                                    className: "grid grid-cols-2 gap-2 text-xs pt-1",
                                    children: [
                                      /* @__PURE__ */ (void 0)(
                                        "div",
                                        {
                                          className: "bg-card p-2 rounded border",
                                          children: [
                                            /* @__PURE__ */ (void 0)(
                                              "span",
                                              {
                                                className:
                                                  "text-muted-foreground block text-[11px]",
                                                children: "Relatórios Migrados:",
                                              },
                                              void 0,
                                              false,
                                              {
                                                fileName: _jsxFileName$4,
                                                lineNumber: 263,
                                                columnNumber: 21,
                                              },
                                              this,
                                            ),
                                            /* @__PURE__ */ (void 0)(
                                              "span",
                                              {
                                                className: "text-sm font-semibold text-primary",
                                                children: progress.reportsMigrated,
                                              },
                                              void 0,
                                              false,
                                              {
                                                fileName: _jsxFileName$4,
                                                lineNumber: 266,
                                                columnNumber: 21,
                                              },
                                              this,
                                            ),
                                          ],
                                        },
                                        void 0,
                                        true,
                                        {
                                          fileName: _jsxFileName$4,
                                          lineNumber: 262,
                                          columnNumber: 19,
                                        },
                                        this,
                                      ),
                                      /* @__PURE__ */ (void 0)(
                                        "div",
                                        {
                                          className: "bg-card p-2 rounded border",
                                          children: [
                                            /* @__PURE__ */ (void 0)(
                                              "span",
                                              {
                                                className:
                                                  "text-muted-foreground block text-[11px]",
                                                children: "Checklists Migrados:",
                                              },
                                              void 0,
                                              false,
                                              {
                                                fileName: _jsxFileName$4,
                                                lineNumber: 271,
                                                columnNumber: 21,
                                              },
                                              this,
                                            ),
                                            /* @__PURE__ */ (void 0)(
                                              "span",
                                              {
                                                className: "text-sm font-semibold text-primary",
                                                children: progress.checklistsMigrated,
                                              },
                                              void 0,
                                              false,
                                              {
                                                fileName: _jsxFileName$4,
                                                lineNumber: 274,
                                                columnNumber: 21,
                                              },
                                              this,
                                            ),
                                          ],
                                        },
                                        void 0,
                                        true,
                                        {
                                          fileName: _jsxFileName$4,
                                          lineNumber: 270,
                                          columnNumber: 19,
                                        },
                                        this,
                                      ),
                                    ],
                                  },
                                  void 0,
                                  true,
                                  {
                                    fileName: _jsxFileName$4,
                                    lineNumber: 261,
                                    columnNumber: 17,
                                  },
                                  this,
                                ),
                              ],
                            },
                            void 0,
                            true,
                            {
                              fileName: _jsxFileName$4,
                              lineNumber: 254,
                              columnNumber: 15,
                            },
                            this,
                          ),
                        errorMsg &&
                          /* @__PURE__ */ (void 0)(
                            "div",
                            {
                              className:
                                "p-3 rounded-lg bg-destructive/10 border border-destructive/30 text-destructive text-xs flex items-center gap-2",
                              children: [
                                /* @__PURE__ */ (void 0)(
                                  CircleAlert,
                                  { className: "w-4 h-4 shrink-0" },
                                  void 0,
                                  false,
                                  {
                                    fileName: _jsxFileName$4,
                                    lineNumber: 284,
                                    columnNumber: 17,
                                  },
                                  this,
                                ),
                                /* @__PURE__ */ (void 0)(
                                  "span",
                                  { children: errorMsg },
                                  void 0,
                                  false,
                                  {
                                    fileName: _jsxFileName$4,
                                    lineNumber: 285,
                                    columnNumber: 17,
                                  },
                                  this,
                                ),
                              ],
                            },
                            void 0,
                            true,
                            {
                              fileName: _jsxFileName$4,
                              lineNumber: 283,
                              columnNumber: 15,
                            },
                            this,
                          ),
                        /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                          "div",
                          {
                            className: "pt-2 flex items-center justify-between gap-3",
                            children: [
                              /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                "button",
                                {
                                  id: "start-supabase-migration-btn",
                                  onClick: handleStartMigration,
                                  disabled: isMigrating,
                                  className:
                                    "w-full sm:w-auto flex items-center justify-center gap-2 px-4 py-2.5 rounded-xl bg-emerald-600 hover:bg-emerald-700 text-white font-medium text-xs shadow-xs transition-colors disabled:opacity-50",
                                  children: isMigrating
                                    ? /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                        import_jsx_dev_runtime.Fragment,
                                        {
                                          children: [
                                            /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                              RefreshCw,
                                              { className: "w-4 h-4 animate-spin" },
                                              void 0,
                                              false,
                                              {
                                                fileName: _jsxFileName$4,
                                                lineNumber: 298,
                                                columnNumber: 21,
                                              },
                                              this,
                                            ),
                                            "Migrando Dados para Supabase...",
                                          ],
                                        },
                                        void 0,
                                        true,
                                        {
                                          fileName: _jsxFileName$4,
                                          lineNumber: 297,
                                          columnNumber: 19,
                                        },
                                        this,
                                      )
                                    : /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                        import_jsx_dev_runtime.Fragment,
                                        {
                                          children: [
                                            /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                              Zap,
                                              { className: "w-4 h-4" },
                                              void 0,
                                              false,
                                              {
                                                fileName: _jsxFileName$4,
                                                lineNumber: 303,
                                                columnNumber: 21,
                                              },
                                              this,
                                            ),
                                            "Iniciar Migração dos Dados Agora",
                                            /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                              ArrowRight,
                                              { className: "w-4 h-4" },
                                              void 0,
                                              false,
                                              {
                                                fileName: _jsxFileName$4,
                                                lineNumber: 305,
                                                columnNumber: 21,
                                              },
                                              this,
                                            ),
                                          ],
                                        },
                                        void 0,
                                        true,
                                        {
                                          fileName: _jsxFileName$4,
                                          lineNumber: 302,
                                          columnNumber: 19,
                                        },
                                        this,
                                      ),
                                },
                                void 0,
                                false,
                                {
                                  fileName: _jsxFileName$4,
                                  lineNumber: 290,
                                  columnNumber: 15,
                                },
                                this,
                              ),
                              progress?.phase === "completed" &&
                                /* @__PURE__ */ (void 0)(
                                  "span",
                                  {
                                    className:
                                      "text-xs text-emerald-600 dark:text-emerald-400 font-medium flex items-center gap-1 shrink-0",
                                    children: [
                                      /* @__PURE__ */ (void 0)(
                                        CircleCheck,
                                        { className: "w-4 h-4" },
                                        void 0,
                                        false,
                                        {
                                          fileName: _jsxFileName$4,
                                          lineNumber: 312,
                                          columnNumber: 19,
                                        },
                                        this,
                                      ),
                                      " Migração Finalizada!",
                                    ],
                                  },
                                  void 0,
                                  true,
                                  {
                                    fileName: _jsxFileName$4,
                                    lineNumber: 311,
                                    columnNumber: 17,
                                  },
                                  this,
                                ),
                            ],
                          },
                          void 0,
                          true,
                          {
                            fileName: _jsxFileName$4,
                            lineNumber: 289,
                            columnNumber: 13,
                          },
                          this,
                        ),
                      ],
                    },
                    void 0,
                    true,
                    {
                      fileName: _jsxFileName$4,
                      lineNumber: 244,
                      columnNumber: 11,
                    },
                    this,
                  ),
                ],
              },
              void 0,
              true,
              {
                fileName: _jsxFileName$4,
                lineNumber: 130,
                columnNumber: 9,
              },
              this,
            ),
            /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
              "div",
              {
                className:
                  "px-6 py-3.5 border-t border-border bg-muted/20 flex items-center justify-between text-xs text-muted-foreground",
                children: [
                  /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                    "span",
                    {
                      children: "O aplicativo já continuará salvando automaticamente no Supabase.",
                    },
                    void 0,
                    false,
                    {
                      fileName: _jsxFileName$4,
                      lineNumber: 321,
                      columnNumber: 11,
                    },
                    this,
                  ),
                  /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                    "button",
                    {
                      id: "close-modal-bottom-btn",
                      onClick: onClose,
                      className:
                        "px-4 py-1.5 rounded-lg border border-border hover:bg-muted text-foreground transition-colors",
                      children: "Fechar",
                    },
                    void 0,
                    false,
                    {
                      fileName: _jsxFileName$4,
                      lineNumber: 322,
                      columnNumber: 11,
                    },
                    this,
                  ),
                ],
              },
              void 0,
              true,
              {
                fileName: _jsxFileName$4,
                lineNumber: 320,
                columnNumber: 9,
              },
              this,
            ),
          ],
        },
        void 0,
        true,
        {
          fileName: _jsxFileName$4,
          lineNumber: 98,
          columnNumber: 7,
        },
        this,
      ),
    },
    void 0,
    false,
    {
      fileName: _jsxFileName$4,
      lineNumber: 94,
      columnNumber: 5,
    },
    this,
  );
}
var pad2$1 = (n) => n.toString().padStart(2, "0");
function getImageDimensions(dataUrl) {
  return new Promise((resolve) => {
    const img = new Image();
    img.onload = () => {
      resolve({
        width: img.width || 800,
        height: img.height || 600,
      });
    };
    img.onerror = () => {
      resolve({
        width: 800,
        height: 600,
      });
    };
    img.src = dataUrl;
  });
}
async function buildPdfBlobForReport(header, events, photos) {
  const doc = new import_jspdf_node_min.default({
    unit: "mm",
    format: "a4",
  });
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
  try {
    doc.addImage(TIC_TRENS_LOGO_DATA_URL, "JPEG", 14, y, 48, 15);
    y += 18;
  } catch {
    y += 2;
  }
  doc.setFont("helvetica", "bold");
  doc.setFontSize(13);
  doc.text("Relatório de Atividade", 14, y);
  y += 5;
  if (h.atividade) {
    doc.setFont("helvetica", "bold");
    doc.setFontSize(11);
    const wrapped = doc.splitTextToSize(h.atividade, contentW - 4);
    doc.text(wrapped, 14, y);
    y += wrapped.length * 5 + 1;
  } else y += 1;
  doc.setFontSize(9.5);
  const labelX = 14;
  const valueX = 48;
  const drawKV = (label, value) => {
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
  doc.setFont("helvetica", "bold");
  doc.setFontSize(10);
  doc.text("PARTICIPANTES:", labelX, y);
  y += 3;
  autoTable(doc, {
    startY: y,
    margin: {
      left: marginL,
      right: marginR,
    },
    head: [["Empresa", "Colaboradores"]],
    body: [[h.empresa || "TIC Trens", h.colaboradores || ""]],
    styles: {
      fontSize: 8.5,
      cellPadding: 2,
      lineColor: [0, 0, 0],
      lineWidth: 0.2,
    },
    headStyles: {
      fillColor: [245, 245, 245],
      textColor: 0,
      fontStyle: "bold",
      lineWidth: 0.2,
    },
    columnStyles: { 0: { cellWidth: 45 } },
  });
  y = (doc.lastAutoTable?.finalY ?? y) + 5;
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
  const tableBody = sortedEvents.map((e) => {
    const d = new Date(e.timestamp);
    const dtStr = `${`${pad2$1(d.getDate())}/${pad2$1(d.getMonth() + 1)}/${d.getFullYear()}`}  ${`${pad2$1(d.getHours())}:${pad2$1(d.getMinutes())}:${pad2$1(d.getSeconds())}`}`;
    const label = EVENT_LABEL[e.kind] || e.kind;
    let desc = "";
    if (e.kind === "inicio_atividade") {
      const kms = [e.kmInicial, e.kmFinal].filter(Boolean).join("/");
      desc = `${label}${kms ? ` (km ${kms})` : ""}${e.descricao ? ` — ${e.descricao}` : ""}`;
    } else if (e.kind === "fim_atividade") {
      const kms = [e.kmInicial, e.kmFinal].filter(Boolean).join("/");
      const parts = [];
      if (kms) parts.push(`km ${kms}`);
      if (e.producao) parts.push(`${e.producao} m`);
      if (e.defeitos !== void 0 && e.defeitos !== null && e.defeitos !== "")
        parts.push(`${e.defeitos} def.`);
      desc = `${label}${parts.length ? ` (${parts.join(", ")})` : ""}${e.descricao ? ` — ${e.descricao}` : ""}`;
    } else {
      const parts = [e.descricao, e.observacoes].filter(Boolean);
      if (parts.length > 0) desc = `${label} — ${parts.join(" — ")}`;
      else desc = label;
    }
    return [dtStr, desc];
  });
  autoTable(doc, {
    startY: y,
    margin: {
      left: marginL,
      right: marginR,
    },
    showHead: "never",
    body: tableBody,
    styles: {
      fontSize: 8.5,
      cellPadding: 1.8,
      lineColor: [255, 255, 255],
      lineWidth: 0,
      textColor: [20, 20, 20],
    },
    alternateRowStyles: { fillColor: [248, 249, 250] },
    bodyStyles: { fillColor: [255, 255, 255] },
    columnStyles: {
      0: { cellWidth: 40 },
      1: { cellWidth: contentW - 40 },
    },
  });
  y = (doc.lastAutoTable?.finalY ?? y) + 5;
  doc.setFontSize(9.5);
  const drawSummary = (label, value) => {
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
  doc.setDrawColor(0);
  doc.setLineWidth(0.3);
  const finalBoxHeight = Math.max(y - boxTop + 3, pageH - 2 * boxTop);
  doc.rect(marginL, boxTop, contentW, finalBoxHeight);
  if (photos && photos.length > 0) {
    const targetPhotos = photos.slice(0, 12);
    const photosPerPage = 6;
    const totalPhotoPages = Math.ceil(targetPhotos.length / photosPerPage);
    for (let pageIdx = 0; pageIdx < totalPhotoPages; pageIdx++) {
      doc.addPage("a4", "p");
      const photoPageTop = 10;
      let py = 14;
      try {
        doc.addImage(TIC_TRENS_LOGO_DATA_URL, "JPEG", 14, py, 40, 12.5);
      } catch {}
      doc.setFont("helvetica", "bold");
      doc.setFontSize(12);
      doc.setTextColor(0, 0, 0);
      const pageSuffix = totalPhotoPages > 1 ? ` (${pageIdx + 1}/${totalPhotoPages})` : "";
      doc.text(`ANEXO FOTOGRÁFICO DE INSPEÇÃO${pageSuffix}`, 58, py + 5);
      doc.setFont("helvetica", "normal");
      doc.setFontSize(8.5);
      doc.setTextColor(60, 60, 60);
      const subHeader = `Atividade: ${h.atividade || "Relatório"}  |  Data: ${formatDate(h.data)}  |  Local: ${h.local || "-"}`;
      doc.text(doc.splitTextToSize(subHeader, contentW - 48), 58, py + 9);
      py += 15;
      doc.setDrawColor(180, 185, 190);
      doc.setLineWidth(0.3);
      doc.line(marginL, py, marginL + contentW, py);
      py += 5;
      const pagePhotos = targetPhotos.slice(pageIdx * photosPerPage, (pageIdx + 1) * photosPerPage);
      const colW = 88;
      const rowH = 72;
      for (let i = 0; i < pagePhotos.length; i++) {
        const globalIdx = pageIdx * photosPerPage + i;
        const p = pagePhotos[i];
        const col = i % 2;
        const row = Math.floor(i / 2);
        const cellX = marginL + col * 98;
        const cellY = py + row * 78;
        doc.setDrawColor(200, 205, 210);
        doc.setFillColor(252, 253, 254);
        doc.rect(cellX, cellY, colW, rowH, "FD");
        doc.setFillColor(240, 242, 245);
        doc.rect(cellX, cellY, colW, 6, "F");
        doc.setFont("helvetica", "bold");
        doc.setFontSize(8);
        doc.setTextColor(40, 40, 40);
        doc.text(`FOTO ${pad2$1(globalIdx + 1)}`, cellX + 3, cellY + 4.2);
        const boxX = cellX + 2;
        const boxY = cellY + 7;
        const boxW = colW - 4;
        const boxH = rowH - 15;
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
        doc.setFont("helvetica", "normal");
        doc.setFontSize(7.5);
        doc.setTextColor(30, 30, 30);
        const captionText = p.caption?.trim() || `Registro fotográfico #${globalIdx + 1}`;
        const wrappedCaption = doc.splitTextToSize(captionText, colW - 4);
        doc.text(wrappedCaption, cellX + 2, cellY + rowH - 3);
      }
      const photosBoxH = Math.ceil(pagePhotos.length / 2) * 78 + 22;
      doc.setDrawColor(0);
      doc.setLineWidth(0.3);
      doc.rect(marginL, photoPageTop, contentW, Math.max(photosBoxH, pageH - 2 * photoPageTop));
    }
  }
  return doc.output("blob");
}
async function downloadReportPdf(header, events, photos, filename) {
  const blob = await buildPdfBlobForReport(header, events, photos);
  const name = filename || generatePdfFilename(header, events);
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = name;
  a.click();
  URL.revokeObjectURL(url);
}
var _jsxFileName$3 = "/app/applet/src/components/CloudReportsDatabase.tsx";
function CloudReportsDatabase({ onLoadReportToForm }) {
  const [reports, setReports] = (0, import_react.useState)([]);
  const [searchTerm, setSearchTerm] = (0, import_react.useState)("");
  const [startDate, setStartDate] = (0, import_react.useState)("");
  const [endDate, setEndDate] = (0, import_react.useState)("");
  const [selectedLocal, setSelectedLocal] = (0, import_react.useState)("");
  const [previewReport, setPreviewReport] = (0, import_react.useState)(null);
  const [sapModalReport, setSapModalReport] = (0, import_react.useState)(null);
  const [toastMsg, setToastMsg] = (0, import_react.useState)(null);
  const [isDeduplicating, setIsDeduplicating] = (0, import_react.useState)(false);
  const [isManualSyncing, setIsManualSyncing] = (0, import_react.useState)(false);
  const [isSupabaseModalOpen, setIsSupabaseModalOpen] = (0, import_react.useState)(false);
  (0, import_react.useEffect)(() => {
    return subscribeCloudReports((data) => {
      setReports(data);
    });
  }, []);
  const showToast = (msg) => {
    setToastMsg(msg);
    setTimeout(() => setToastMsg(null), 3500);
  };
  const handleManualSync = async () => {
    setIsManualSyncing(true);
    try {
      const { reports: updated } = await syncCloudReportsAcrossDevices();
      setReports(updated);
      showToast(
        `✓ Sincronizado! ${updated.length} relatório(s) sincronizados entre Computador e Celular.`,
      );
    } catch (e) {
      console.warn("Manual sync error:", e);
      showToast("Aviso ao sincronizar relatórios.");
    } finally {
      setIsManualSyncing(false);
    }
  };
  const handleDeduplicate = async () => {
    setIsDeduplicating(true);
    try {
      const result = await deduplicateCloudReports();
      if (result.duplicatesRemoved > 0)
        showToast(
          `✨ ${result.duplicatesRemoved} relatório(s) repetido(s) eliminado(s)! Mantido somente o último sincronizado de cada dia.`,
        );
      else showToast("✓ Banco otimizado: Não há duplicidades no mesmo dia.");
    } catch (e) {
      console.error(e);
      showToast("Aviso ao limpar duplicados da nuvem.");
    } finally {
      setIsDeduplicating(false);
    }
  };
  const uniqueLocals = (0, import_react.useMemo)(() => {
    const locs = /* @__PURE__ */ new Set();
    reports.forEach((r) => {
      if (r.local && r.local.trim()) locs.add(r.local.trim());
    });
    return Array.from(locs).sort();
  }, [reports]);
  const filteredReports = (0, import_react.useMemo)(() => {
    return reports.filter((r) => {
      if (searchTerm.trim()) {
        const term = searchTerm.toLowerCase().trim();
        if (!(
          (r.nome && r.nome.toLowerCase().includes(term)) ||
          (r.local && r.local.toLowerCase().includes(term)) ||
          (r.atividade && r.atividade.toLowerCase().includes(term)) ||
          (r.empresa && r.empresa.toLowerCase().includes(term)) ||
          (r.om && r.om.toLowerCase().includes(term)) ||
          (r.sinais && r.sinais.toLowerCase().includes(term)) ||
          (r.data && r.data.includes(term)) ||
          (r.colaboradores && r.colaboradores.toLowerCase().includes(term)) ||
          (r.descritivo && r.descritivo.toLowerCase().includes(term)) ||
          (r.id && r.id.toLowerCase().includes(term))
        ))
          return false;
      }
      if (startDate && r.data && r.data < startDate) return false;
      if (endDate && r.data && r.data > endDate) return false;
      if (selectedLocal && r.local !== selectedLocal) return false;
      return true;
    });
  }, [reports, searchTerm, startDate, endDate, selectedLocal]);
  const stats = (0, import_react.useMemo)(() => {
    return {
      totalCount: filteredReports.length,
      totalProd: filteredReports.reduce((acc, r) => acc + (r.totals?.producaoTotal || 0), 0),
      totalDefects: filteredReports.reduce((acc, r) => acc + (r.totals?.defeitosTotal || 0), 0),
      totalJornadaMs: filteredReports.reduce((acc, r) => acc + (r.totals?.jornada || 0), 0),
    };
  }, [filteredReports]);
  const handleDelete = async (r) => {
    if (
      confirm(
        `Tem certeza que deseja excluir permanentemente o relatório de ${formatDate(r.data)} (${r.nome || "sem nome"}) da nuvem?`,
      )
    )
      try {
        setReports((prev) => prev.filter((item) => item.id !== r.id));
        await deleteCloudReport(r.id);
        showToast("Relatório removido do banco de dados na nuvem.");
        if (previewReport?.id === r.id) setPreviewReport(null);
      } catch (err) {
        console.error("Erro ao deletar relatório:", err);
        showToast("Erro ao excluir relatório da nuvem.");
      }
  };
  const handleDownloadSinglePdf = (r) => {
    const { header, events, photos } = cloudReportToHeaderAndEvents(r);
    downloadReportPdf(header, events, photos, generatePdfFilename(header, events));
    showToast(`Baixando PDF de ${r.nome || "Relatório"}...`);
  };
  const handleDownloadSingleExcel = (r) => {
    exportSingleReportToExcel(r);
    showToast(`Baixando Excel de ${r.nome || "Relatório"}...`);
  };
  const handleExportAllExcel = () => {
    if (reports.length === 0) {
      alert("Não há relatórios salvos no banco de dados na nuvem.");
      return;
    }
    exportAllReportsToExcel(
      reports,
      `banco-de-dados-geral-${/* @__PURE__ */ new Date().toISOString().slice(0, 10)}.xlsx`,
    );
    showToast("Exportando banco de dados completo para Excel...");
  };
  const handleExportFilteredExcel = () => {
    if (filteredReports.length === 0) {
      alert("Nenhum relatório encontrado para os filtros selecionados.");
      return;
    }
    exportAllReportsToExcel(
      filteredReports,
      `banco-de-dados-filtrado-${/* @__PURE__ */ new Date().toISOString().slice(0, 10)}.xlsx`,
    );
    showToast("Exportando relatórios filtrados para Excel...");
  };
  return /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
    "div",
    {
      className: "space-y-6",
      children: [
        toastMsg &&
          /* @__PURE__ */ (void 0)(
            "div",
            {
              className:
                "fixed bottom-4 right-4 z-50 flex items-center gap-2 rounded-lg bg-emerald-600 px-4 py-3 text-sm font-medium text-white shadow-lg animate-in fade-in slide-in-from-bottom-2",
              children: [
                /* @__PURE__ */ (void 0)(
                  Check,
                  { className: "h-4 w-4 shrink-0" },
                  void 0,
                  false,
                  {
                    fileName: _jsxFileName$3,
                    lineNumber: 223,
                    columnNumber: 11,
                  },
                  this,
                ),
                /* @__PURE__ */ (void 0)(
                  "span",
                  { children: toastMsg },
                  void 0,
                  false,
                  {
                    fileName: _jsxFileName$3,
                    lineNumber: 224,
                    columnNumber: 11,
                  },
                  this,
                ),
              ],
            },
            void 0,
            true,
            {
              fileName: _jsxFileName$3,
              lineNumber: 222,
              columnNumber: 9,
            },
            this,
          ),
        /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
          "div",
          {
            className: "rounded-xl border border-primary/20 bg-card p-5 shadow-sm",
            children: [
              /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                "div",
                {
                  className: "flex flex-col gap-4 md:flex-row md:items-center md:justify-between",
                  children: [
                    /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                      "div",
                      {
                        className: "flex items-start gap-3",
                        children: [
                          /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                            "div",
                            {
                              className:
                                "flex h-10 w-10 shrink-0 items-center justify-center rounded-lg bg-primary/10 text-primary",
                              children: /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                Cloud,
                                { className: "h-5 w-5" },
                                void 0,
                                false,
                                {
                                  fileName: _jsxFileName$3,
                                  lineNumber: 233,
                                  columnNumber: 15,
                                },
                                this,
                              ),
                            },
                            void 0,
                            false,
                            {
                              fileName: _jsxFileName$3,
                              lineNumber: 232,
                              columnNumber: 13,
                            },
                            this,
                          ),
                          /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                            "div",
                            {
                              children: [
                                /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                  "div",
                                  {
                                    className: "flex flex-wrap items-center gap-2",
                                    children: [
                                      /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                        "h2",
                                        {
                                          className:
                                            "text-lg font-semibold tracking-tight text-foreground",
                                          children: "Banco de Dados de Relatórios na Nuvem",
                                        },
                                        void 0,
                                        false,
                                        {
                                          fileName: _jsxFileName$3,
                                          lineNumber: 237,
                                          columnNumber: 17,
                                        },
                                        this,
                                      ),
                                      /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                        "span",
                                        {
                                          className:
                                            "inline-flex items-center gap-1 rounded-full bg-emerald-500/10 px-2 py-0.5 text-[10px] font-semibold text-emerald-600 dark:text-emerald-400 border border-emerald-500/20",
                                          children: [
                                            /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                              Check,
                                              { className: "h-3 w-3" },
                                              void 0,
                                              false,
                                              {
                                                fileName: _jsxFileName$3,
                                                lineNumber: 241,
                                                columnNumber: 19,
                                              },
                                              this,
                                            ),
                                            "Último sincronizado por dia",
                                          ],
                                        },
                                        void 0,
                                        true,
                                        {
                                          fileName: _jsxFileName$3,
                                          lineNumber: 240,
                                          columnNumber: 17,
                                        },
                                        this,
                                      ),
                                      /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                        "span",
                                        {
                                          className:
                                            "inline-flex items-center gap-1 rounded-full bg-blue-500/10 px-2 py-0.5 text-[10px] font-semibold text-blue-600 dark:text-blue-400 border border-blue-500/20",
                                          children: [
                                            /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                              RefreshCw,
                                              { className: "h-3 w-3" },
                                              void 0,
                                              false,
                                              {
                                                fileName: _jsxFileName$3,
                                                lineNumber: 245,
                                                columnNumber: 19,
                                              },
                                              this,
                                            ),
                                            "Sincronismo PC & Celular Ativo",
                                          ],
                                        },
                                        void 0,
                                        true,
                                        {
                                          fileName: _jsxFileName$3,
                                          lineNumber: 244,
                                          columnNumber: 17,
                                        },
                                        this,
                                      ),
                                    ],
                                  },
                                  void 0,
                                  true,
                                  {
                                    fileName: _jsxFileName$3,
                                    lineNumber: 236,
                                    columnNumber: 15,
                                  },
                                  this,
                                ),
                                /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                  "p",
                                  {
                                    className: "text-xs text-muted-foreground mt-0.5",
                                    children:
                                      "Central de visualização, busca e exportação de todos os relatórios da equipe em formato Excel e PDF.",
                                  },
                                  void 0,
                                  false,
                                  {
                                    fileName: _jsxFileName$3,
                                    lineNumber: 249,
                                    columnNumber: 15,
                                  },
                                  this,
                                ),
                              ],
                            },
                            void 0,
                            true,
                            {
                              fileName: _jsxFileName$3,
                              lineNumber: 235,
                              columnNumber: 13,
                            },
                            this,
                          ),
                        ],
                      },
                      void 0,
                      true,
                      {
                        fileName: _jsxFileName$3,
                        lineNumber: 231,
                        columnNumber: 11,
                      },
                      this,
                    ),
                    /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                      "div",
                      {
                        className: "flex flex-wrap items-center gap-2",
                        children: [
                          /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                            "button",
                            {
                              type: "button",
                              onClick: () => setIsSupabaseModalOpen(true),
                              className:
                                "inline-flex items-center gap-1.5 rounded-lg border border-emerald-500/30 bg-emerald-500/10 hover:bg-emerald-500/20 text-emerald-700 dark:text-emerald-400 py-2 px-3 text-xs font-semibold shadow-sm transition-colors",
                              title:
                                "Gerenciar migração e sincronização com o banco de dados Supabase",
                              children: [
                                /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                  Database,
                                  {
                                    className: "h-3.5 w-3.5 text-emerald-600 dark:text-emerald-400",
                                  },
                                  void 0,
                                  false,
                                  {
                                    fileName: _jsxFileName$3,
                                    lineNumber: 263,
                                    columnNumber: 15,
                                  },
                                  this,
                                ),
                                /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                  "span",
                                  { children: "Migração Supabase" },
                                  void 0,
                                  false,
                                  {
                                    fileName: _jsxFileName$3,
                                    lineNumber: 264,
                                    columnNumber: 15,
                                  },
                                  this,
                                ),
                              ],
                            },
                            void 0,
                            true,
                            {
                              fileName: _jsxFileName$3,
                              lineNumber: 257,
                              columnNumber: 13,
                            },
                            this,
                          ),
                          /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                            "button",
                            {
                              type: "button",
                              onClick: handleManualSync,
                              disabled: isManualSyncing,
                              className:
                                "inline-flex items-center gap-1.5 rounded-lg border border-border bg-background hover:bg-muted text-foreground py-2 px-3 text-xs font-semibold shadow-sm transition-colors disabled:opacity-50",
                              title:
                                "Sincroniza instantaneamente os relatórios entre computador e celular",
                              children: [
                                /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                  RefreshCw,
                                  {
                                    className: `h-3.5 w-3.5 ${isManualSyncing ? "animate-spin text-primary" : ""}`,
                                  },
                                  void 0,
                                  false,
                                  {
                                    fileName: _jsxFileName$3,
                                    lineNumber: 274,
                                    columnNumber: 15,
                                  },
                                  this,
                                ),
                                /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                  "span",
                                  {
                                    children: isManualSyncing
                                      ? "Sincronizando..."
                                      : "Sincronizar PC / Celular",
                                  },
                                  void 0,
                                  false,
                                  {
                                    fileName: _jsxFileName$3,
                                    lineNumber: 277,
                                    columnNumber: 15,
                                  },
                                  this,
                                ),
                              ],
                            },
                            void 0,
                            true,
                            {
                              fileName: _jsxFileName$3,
                              lineNumber: 267,
                              columnNumber: 13,
                            },
                            this,
                          ),
                          /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                            "button",
                            {
                              type: "button",
                              onClick: handleDeduplicate,
                              disabled: isDeduplicating || reports.length === 0,
                              className:
                                "inline-flex items-center gap-1.5 rounded-lg border border-amber-500/30 bg-amber-500/10 hover:bg-amber-500/20 text-amber-600 dark:text-amber-400 py-2 px-3 text-xs font-semibold shadow-sm transition-colors disabled:opacity-50",
                              title:
                                "Garante que não haja relatórios repetidos no mesmo dia, deixando somente o último sincronizado no banco em nuvem",
                              children: [
                                /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                  Sparkles,
                                  {
                                    className: `h-3.5 w-3.5 ${isDeduplicating ? "animate-spin text-amber-500" : ""}`,
                                  },
                                  void 0,
                                  false,
                                  {
                                    fileName: _jsxFileName$3,
                                    lineNumber: 287,
                                    columnNumber: 15,
                                  },
                                  this,
                                ),
                                /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                  "span",
                                  {
                                    children: isDeduplicating
                                      ? "Limpando..."
                                      : "Eliminar Repetidos (Mesmo Dia)",
                                  },
                                  void 0,
                                  false,
                                  {
                                    fileName: _jsxFileName$3,
                                    lineNumber: 290,
                                    columnNumber: 15,
                                  },
                                  this,
                                ),
                              ],
                            },
                            void 0,
                            true,
                            {
                              fileName: _jsxFileName$3,
                              lineNumber: 280,
                              columnNumber: 13,
                            },
                            this,
                          ),
                          /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                            "button",
                            {
                              type: "button",
                              onClick: handleExportAllExcel,
                              disabled: reports.length === 0,
                              className:
                                "btn-primary inline-flex items-center gap-2 py-2 px-3.5 text-xs font-semibold shadow-sm disabled:opacity-50",
                              title:
                                "Baixar todos os relatórios salvos na nuvem em uma planilha Excel completa com abas de resumo e eventos",
                              children: [
                                /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                  FileSpreadsheet,
                                  { className: "h-4 w-4" },
                                  void 0,
                                  false,
                                  {
                                    fileName: _jsxFileName$3,
                                    lineNumber: 300,
                                    columnNumber: 15,
                                  },
                                  this,
                                ),
                                /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                  "span",
                                  { children: "Baixar Banco Geral (Excel .xlsx)" },
                                  void 0,
                                  false,
                                  {
                                    fileName: _jsxFileName$3,
                                    lineNumber: 301,
                                    columnNumber: 15,
                                  },
                                  this,
                                ),
                              ],
                            },
                            void 0,
                            true,
                            {
                              fileName: _jsxFileName$3,
                              lineNumber: 293,
                              columnNumber: 13,
                            },
                            this,
                          ),
                          filteredReports.length < reports.length &&
                            /* @__PURE__ */ (void 0)(
                              "button",
                              {
                                type: "button",
                                onClick: handleExportFilteredExcel,
                                className:
                                  "btn-secondary inline-flex items-center gap-2 py-2 px-3 text-xs",
                                title: "Baixar apenas os resultados filtrados em planilha Excel",
                                children: [
                                  /* @__PURE__ */ (void 0)(
                                    Download,
                                    { className: "h-3.5 w-3.5" },
                                    void 0,
                                    false,
                                    {
                                      fileName: _jsxFileName$3,
                                      lineNumber: 311,
                                      columnNumber: 17,
                                    },
                                    this,
                                  ),
                                  /* @__PURE__ */ (void 0)(
                                    "span",
                                    {
                                      children: ["Baixar Filtrados (", filteredReports.length, ")"],
                                    },
                                    void 0,
                                    true,
                                    {
                                      fileName: _jsxFileName$3,
                                      lineNumber: 312,
                                      columnNumber: 17,
                                    },
                                    this,
                                  ),
                                ],
                              },
                              void 0,
                              true,
                              {
                                fileName: _jsxFileName$3,
                                lineNumber: 305,
                                columnNumber: 15,
                              },
                              this,
                            ),
                        ],
                      },
                      void 0,
                      true,
                      {
                        fileName: _jsxFileName$3,
                        lineNumber: 256,
                        columnNumber: 11,
                      },
                      this,
                    ),
                  ],
                },
                void 0,
                true,
                {
                  fileName: _jsxFileName$3,
                  lineNumber: 230,
                  columnNumber: 9,
                },
                this,
              ),
              /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                "div",
                {
                  className:
                    "mt-4 grid grid-cols-2 gap-3 sm:grid-cols-4 border-t border-border pt-4 text-xs",
                  children: [
                    /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                      "div",
                      {
                        className: "rounded-lg bg-secondary/50 p-2.5",
                        children: [
                          /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                            "span",
                            {
                              className:
                                "text-muted-foreground block text-[10px] uppercase font-semibold",
                              children: "Total de Relatórios",
                            },
                            void 0,
                            false,
                            {
                              fileName: _jsxFileName$3,
                              lineNumber: 321,
                              columnNumber: 13,
                            },
                            this,
                          ),
                          /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                            "span",
                            {
                              className: "text-base font-bold text-foreground font-mono",
                              children: [stats.totalCount, " de ", reports.length],
                            },
                            void 0,
                            true,
                            {
                              fileName: _jsxFileName$3,
                              lineNumber: 324,
                              columnNumber: 13,
                            },
                            this,
                          ),
                        ],
                      },
                      void 0,
                      true,
                      {
                        fileName: _jsxFileName$3,
                        lineNumber: 320,
                        columnNumber: 11,
                      },
                      this,
                    ),
                    /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                      "div",
                      {
                        className: "rounded-lg bg-secondary/50 p-2.5",
                        children: [
                          /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                            "span",
                            {
                              className:
                                "text-muted-foreground block text-[10px] uppercase font-semibold",
                              children: "Produção Acumulada",
                            },
                            void 0,
                            false,
                            {
                              fileName: _jsxFileName$3,
                              lineNumber: 329,
                              columnNumber: 13,
                            },
                            this,
                          ),
                          /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                            "span",
                            {
                              className: "text-base font-bold text-emerald-600 font-mono",
                              children: [stats.totalProd, " m"],
                            },
                            void 0,
                            true,
                            {
                              fileName: _jsxFileName$3,
                              lineNumber: 332,
                              columnNumber: 13,
                            },
                            this,
                          ),
                        ],
                      },
                      void 0,
                      true,
                      {
                        fileName: _jsxFileName$3,
                        lineNumber: 328,
                        columnNumber: 11,
                      },
                      this,
                    ),
                    /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                      "div",
                      {
                        className: "rounded-lg bg-secondary/50 p-2.5",
                        children: [
                          /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                            "span",
                            {
                              className:
                                "text-muted-foreground block text-[10px] uppercase font-semibold",
                              children: "Defeitos Acumulados",
                            },
                            void 0,
                            false,
                            {
                              fileName: _jsxFileName$3,
                              lineNumber: 337,
                              columnNumber: 13,
                            },
                            this,
                          ),
                          /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                            "span",
                            {
                              className: "text-base font-bold text-amber-600 font-mono",
                              children: stats.totalDefects,
                            },
                            void 0,
                            false,
                            {
                              fileName: _jsxFileName$3,
                              lineNumber: 340,
                              columnNumber: 13,
                            },
                            this,
                          ),
                        ],
                      },
                      void 0,
                      true,
                      {
                        fileName: _jsxFileName$3,
                        lineNumber: 336,
                        columnNumber: 11,
                      },
                      this,
                    ),
                    /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                      "div",
                      {
                        className: "rounded-lg bg-secondary/50 p-2.5",
                        children: [
                          /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                            "span",
                            {
                              className:
                                "text-muted-foreground block text-[10px] uppercase font-semibold",
                              children: "Jornada Total",
                            },
                            void 0,
                            false,
                            {
                              fileName: _jsxFileName$3,
                              lineNumber: 345,
                              columnNumber: 13,
                            },
                            this,
                          ),
                          /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                            "span",
                            {
                              className: "text-base font-bold text-foreground font-mono",
                              children: formatDuration(stats.totalJornadaMs),
                            },
                            void 0,
                            false,
                            {
                              fileName: _jsxFileName$3,
                              lineNumber: 348,
                              columnNumber: 13,
                            },
                            this,
                          ),
                        ],
                      },
                      void 0,
                      true,
                      {
                        fileName: _jsxFileName$3,
                        lineNumber: 344,
                        columnNumber: 11,
                      },
                      this,
                    ),
                  ],
                },
                void 0,
                true,
                {
                  fileName: _jsxFileName$3,
                  lineNumber: 319,
                  columnNumber: 9,
                },
                this,
              ),
            ],
          },
          void 0,
          true,
          {
            fileName: _jsxFileName$3,
            lineNumber: 229,
            columnNumber: 7,
          },
          this,
        ),
        /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
          "div",
          {
            className: "rounded-xl border border-border bg-card p-4 shadow-sm space-y-3",
            children: [
              /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                "div",
                {
                  className: "flex flex-col md:flex-row items-center gap-3",
                  children: [
                    /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                      "div",
                      {
                        className: "relative flex-1 w-full",
                        children: [
                          /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                            Search,
                            {
                              className:
                                "absolute left-3 top-1/2 -translate-y-1/2 h-4 w-4 text-muted-foreground",
                            },
                            void 0,
                            false,
                            {
                              fileName: _jsxFileName$3,
                              lineNumber: 360,
                              columnNumber: 13,
                            },
                            this,
                          ),
                          /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                            "input",
                            {
                              type: "text",
                              value: searchTerm,
                              onChange: (e) => setSearchTerm(e.target.value),
                              placeholder:
                                "Pesquisar por responsável, local, atividade, empresa, OM...",
                              className: "input pl-9 pr-8 text-xs w-full",
                            },
                            void 0,
                            false,
                            {
                              fileName: _jsxFileName$3,
                              lineNumber: 361,
                              columnNumber: 13,
                            },
                            this,
                          ),
                          searchTerm &&
                            /* @__PURE__ */ (void 0)(
                              "button",
                              {
                                type: "button",
                                onClick: () => setSearchTerm(""),
                                className:
                                  "absolute right-2.5 top-1/2 -translate-y-1/2 text-muted-foreground hover:text-foreground",
                                children: /* @__PURE__ */ (void 0)(
                                  X,
                                  { className: "h-3.5 w-3.5" },
                                  void 0,
                                  false,
                                  {
                                    fileName: _jsxFileName$3,
                                    lineNumber: 374,
                                    columnNumber: 17,
                                  },
                                  this,
                                ),
                              },
                              void 0,
                              false,
                              {
                                fileName: _jsxFileName$3,
                                lineNumber: 369,
                                columnNumber: 15,
                              },
                              this,
                            ),
                        ],
                      },
                      void 0,
                      true,
                      {
                        fileName: _jsxFileName$3,
                        lineNumber: 359,
                        columnNumber: 11,
                      },
                      this,
                    ),
                    /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                      "div",
                      {
                        className:
                          "grid grid-cols-2 sm:grid-cols-3 gap-2 w-full md:w-auto shrink-0",
                        children: [
                          /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                            "div",
                            {
                              children: /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                "select",
                                {
                                  value: selectedLocal,
                                  onChange: (e) => setSelectedLocal(e.target.value),
                                  className: "input text-xs cursor-pointer w-full",
                                  children: [
                                    /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                      "option",
                                      {
                                        value: "",
                                        children: "-- Todos os Locais --",
                                      },
                                      void 0,
                                      false,
                                      {
                                        fileName: _jsxFileName$3,
                                        lineNumber: 387,
                                        columnNumber: 17,
                                      },
                                      this,
                                    ),
                                    uniqueLocals.map((loc) =>
                                      /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                        "option",
                                        {
                                          value: loc,
                                          children: loc,
                                        },
                                        loc,
                                        false,
                                        {
                                          fileName: _jsxFileName$3,
                                          lineNumber: 389,
                                          columnNumber: 19,
                                        },
                                        this,
                                      ),
                                    ),
                                  ],
                                },
                                void 0,
                                true,
                                {
                                  fileName: _jsxFileName$3,
                                  lineNumber: 382,
                                  columnNumber: 15,
                                },
                                this,
                              ),
                            },
                            void 0,
                            false,
                            {
                              fileName: _jsxFileName$3,
                              lineNumber: 381,
                              columnNumber: 13,
                            },
                            this,
                          ),
                          /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                            "div",
                            {
                              className: "flex items-center gap-1",
                              children: /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                "input",
                                {
                                  type: "date",
                                  value: startDate,
                                  onChange: (e) => setStartDate(e.target.value),
                                  className: "input text-xs w-full",
                                  title: "Data inicial",
                                },
                                void 0,
                                false,
                                {
                                  fileName: _jsxFileName$3,
                                  lineNumber: 397,
                                  columnNumber: 15,
                                },
                                this,
                              ),
                            },
                            void 0,
                            false,
                            {
                              fileName: _jsxFileName$3,
                              lineNumber: 396,
                              columnNumber: 13,
                            },
                            this,
                          ),
                          /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                            "div",
                            {
                              className: "flex items-center gap-1 col-span-2 sm:col-span-1",
                              children: /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                "input",
                                {
                                  type: "date",
                                  value: endDate,
                                  onChange: (e) => setEndDate(e.target.value),
                                  className: "input text-xs w-full",
                                  title: "Data final",
                                },
                                void 0,
                                false,
                                {
                                  fileName: _jsxFileName$3,
                                  lineNumber: 407,
                                  columnNumber: 15,
                                },
                                this,
                              ),
                            },
                            void 0,
                            false,
                            {
                              fileName: _jsxFileName$3,
                              lineNumber: 406,
                              columnNumber: 13,
                            },
                            this,
                          ),
                        ],
                      },
                      void 0,
                      true,
                      {
                        fileName: _jsxFileName$3,
                        lineNumber: 380,
                        columnNumber: 11,
                      },
                      this,
                    ),
                  ],
                },
                void 0,
                true,
                {
                  fileName: _jsxFileName$3,
                  lineNumber: 357,
                  columnNumber: 9,
                },
                this,
              ),
              (searchTerm || startDate || endDate || selectedLocal) &&
                /* @__PURE__ */ (void 0)(
                  "div",
                  {
                    className:
                      "flex items-center justify-between gap-2 pt-1 border-t border-border/60 text-xs text-muted-foreground",
                    children: [
                      /* @__PURE__ */ (void 0)(
                        "div",
                        {
                          className: "flex flex-wrap items-center gap-1.5",
                          children: [
                            /* @__PURE__ */ (void 0)(
                              Funnel,
                              { className: "h-3 w-3 text-primary shrink-0" },
                              void 0,
                              false,
                              {
                                fileName: _jsxFileName$3,
                                lineNumber: 422,
                                columnNumber: 15,
                              },
                              this,
                            ),
                            /* @__PURE__ */ (void 0)(
                              "span",
                              { children: "Filtros ativos:" },
                              void 0,
                              false,
                              {
                                fileName: _jsxFileName$3,
                                lineNumber: 423,
                                columnNumber: 15,
                              },
                              this,
                            ),
                            searchTerm &&
                              /* @__PURE__ */ (void 0)(
                                "span",
                                {
                                  className:
                                    "rounded bg-accent px-2 py-0.5 font-medium text-foreground",
                                  children: ['"', searchTerm, '"'],
                                },
                                void 0,
                                true,
                                {
                                  fileName: _jsxFileName$3,
                                  lineNumber: 425,
                                  columnNumber: 17,
                                },
                                this,
                              ),
                            selectedLocal &&
                              /* @__PURE__ */ (void 0)(
                                "span",
                                {
                                  className:
                                    "rounded bg-accent px-2 py-0.5 font-medium text-foreground",
                                  children: ["Local: ", selectedLocal],
                                },
                                void 0,
                                true,
                                {
                                  fileName: _jsxFileName$3,
                                  lineNumber: 430,
                                  columnNumber: 17,
                                },
                                this,
                              ),
                            startDate &&
                              /* @__PURE__ */ (void 0)(
                                "span",
                                {
                                  className:
                                    "rounded bg-accent px-2 py-0.5 font-medium text-foreground",
                                  children: ["De: ", formatDate(startDate)],
                                },
                                void 0,
                                true,
                                {
                                  fileName: _jsxFileName$3,
                                  lineNumber: 435,
                                  columnNumber: 17,
                                },
                                this,
                              ),
                            endDate &&
                              /* @__PURE__ */ (void 0)(
                                "span",
                                {
                                  className:
                                    "rounded bg-accent px-2 py-0.5 font-medium text-foreground",
                                  children: ["Até: ", formatDate(endDate)],
                                },
                                void 0,
                                true,
                                {
                                  fileName: _jsxFileName$3,
                                  lineNumber: 440,
                                  columnNumber: 17,
                                },
                                this,
                              ),
                          ],
                        },
                        void 0,
                        true,
                        {
                          fileName: _jsxFileName$3,
                          lineNumber: 421,
                          columnNumber: 13,
                        },
                        this,
                      ),
                      /* @__PURE__ */ (void 0)(
                        "button",
                        {
                          type: "button",
                          onClick: () => {
                            setSearchTerm("");
                            setStartDate("");
                            setEndDate("");
                            setSelectedLocal("");
                          },
                          className: "text-xs text-primary hover:underline font-medium shrink-0",
                          children: "Limpar Filtros",
                        },
                        void 0,
                        false,
                        {
                          fileName: _jsxFileName$3,
                          lineNumber: 446,
                          columnNumber: 13,
                        },
                        this,
                      ),
                    ],
                  },
                  void 0,
                  true,
                  {
                    fileName: _jsxFileName$3,
                    lineNumber: 420,
                    columnNumber: 11,
                  },
                  this,
                ),
            ],
          },
          void 0,
          true,
          {
            fileName: _jsxFileName$3,
            lineNumber: 356,
            columnNumber: 7,
          },
          this,
        ),
        /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
          "div",
          {
            className: "rounded-xl border border-border bg-card shadow-sm overflow-hidden",
            children:
              filteredReports.length === 0
                ? /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                    "div",
                    {
                      className: "p-8 text-center text-muted-foreground space-y-2",
                      children: [
                        /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                          Cloud,
                          { className: "h-10 w-10 mx-auto text-muted-foreground/40 stroke-1" },
                          void 0,
                          false,
                          {
                            fileName: _jsxFileName$3,
                            lineNumber: 466,
                            columnNumber: 13,
                          },
                          this,
                        ),
                        /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                          "p",
                          {
                            className: "text-sm font-medium",
                            children: "Nenhum relatório encontrado na nuvem.",
                          },
                          void 0,
                          false,
                          {
                            fileName: _jsxFileName$3,
                            lineNumber: 467,
                            columnNumber: 13,
                          },
                          this,
                        ),
                        /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                          "p",
                          {
                            className: "text-xs text-muted-foreground max-w-md mx-auto",
                            children:
                              reports.length === 0
                                ? "Preencha o formulário e clique em 'Salvar Relatório na Nuvem' para registrar os dados."
                                : "Tente alterar os termos da busca ou limpar os filtros de data e local.",
                          },
                          void 0,
                          false,
                          {
                            fileName: _jsxFileName$3,
                            lineNumber: 468,
                            columnNumber: 13,
                          },
                          this,
                        ),
                      ],
                    },
                    void 0,
                    true,
                    {
                      fileName: _jsxFileName$3,
                      lineNumber: 465,
                      columnNumber: 11,
                    },
                    this,
                  )
                : /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                    "div",
                    {
                      className: "overflow-x-auto",
                      children: /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                        "table",
                        {
                          className: "w-full text-left text-xs",
                          children: [
                            /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                              "thead",
                              {
                                className:
                                  "bg-secondary/60 text-muted-foreground border-b border-border font-semibold uppercase tracking-wider text-[10px]",
                                children: /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                  "tr",
                                  {
                                    children: [
                                      /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                        "th",
                                        {
                                          className: "py-3 px-3",
                                          children: "Data / Envio",
                                        },
                                        void 0,
                                        false,
                                        {
                                          fileName: _jsxFileName$3,
                                          lineNumber: 479,
                                          columnNumber: 19,
                                        },
                                        this,
                                      ),
                                      /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                        "th",
                                        {
                                          className: "py-3 px-3",
                                          children: "Responsável",
                                        },
                                        void 0,
                                        false,
                                        {
                                          fileName: _jsxFileName$3,
                                          lineNumber: 480,
                                          columnNumber: 19,
                                        },
                                        this,
                                      ),
                                      /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                        "th",
                                        {
                                          className: "py-3 px-3",
                                          children: "Atividade / Empresa",
                                        },
                                        void 0,
                                        false,
                                        {
                                          fileName: _jsxFileName$3,
                                          lineNumber: 481,
                                          columnNumber: 19,
                                        },
                                        this,
                                      ),
                                      /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                        "th",
                                        {
                                          className: "py-3 px-3",
                                          children: "Local / OM",
                                        },
                                        void 0,
                                        false,
                                        {
                                          fileName: _jsxFileName$3,
                                          lineNumber: 482,
                                          columnNumber: 19,
                                        },
                                        this,
                                      ),
                                      /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                        "th",
                                        {
                                          className: "py-3 px-3 text-center",
                                          children: "Produção",
                                        },
                                        void 0,
                                        false,
                                        {
                                          fileName: _jsxFileName$3,
                                          lineNumber: 483,
                                          columnNumber: 19,
                                        },
                                        this,
                                      ),
                                      /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                        "th",
                                        {
                                          className: "py-3 px-3 text-center",
                                          children: "Defeitos",
                                        },
                                        void 0,
                                        false,
                                        {
                                          fileName: _jsxFileName$3,
                                          lineNumber: 484,
                                          columnNumber: 19,
                                        },
                                        this,
                                      ),
                                      /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                        "th",
                                        {
                                          className: "py-3 px-3 text-center",
                                          children: "Jornada",
                                        },
                                        void 0,
                                        false,
                                        {
                                          fileName: _jsxFileName$3,
                                          lineNumber: 485,
                                          columnNumber: 19,
                                        },
                                        this,
                                      ),
                                      /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                        "th",
                                        {
                                          className: "py-3 px-3 text-right",
                                          children: "Ações (Excel / PDF)",
                                        },
                                        void 0,
                                        false,
                                        {
                                          fileName: _jsxFileName$3,
                                          lineNumber: 486,
                                          columnNumber: 19,
                                        },
                                        this,
                                      ),
                                    ],
                                  },
                                  void 0,
                                  true,
                                  {
                                    fileName: _jsxFileName$3,
                                    lineNumber: 478,
                                    columnNumber: 17,
                                  },
                                  this,
                                ),
                              },
                              void 0,
                              false,
                              {
                                fileName: _jsxFileName$3,
                                lineNumber: 477,
                                columnNumber: 15,
                              },
                              this,
                            ),
                            /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                              "tbody",
                              {
                                className: "divide-y divide-border/60",
                                children: filteredReports.map((r) =>
                                  /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                    "tr",
                                    {
                                      className: "hover:bg-accent/40 transition-colors",
                                      children: [
                                        /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                          "td",
                                          {
                                            className:
                                              "py-3 px-3 font-medium text-foreground whitespace-nowrap",
                                            children: [
                                              /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                                "div",
                                                {
                                                  className:
                                                    "flex items-center gap-1.5 font-mono text-xs",
                                                  children: [
                                                    /* @__PURE__ */ (0,
                                                    import_jsx_dev_runtime.jsxDEV)(
                                                      Calendar,
                                                      {
                                                        className:
                                                          "h-3.5 w-3.5 text-primary shrink-0",
                                                      },
                                                      void 0,
                                                      false,
                                                      {
                                                        fileName: _jsxFileName$3,
                                                        lineNumber: 494,
                                                        columnNumber: 25,
                                                      },
                                                      this,
                                                    ),
                                                    /* @__PURE__ */ (0,
                                                    import_jsx_dev_runtime.jsxDEV)(
                                                      "span",
                                                      { children: formatDate(r.data) },
                                                      void 0,
                                                      false,
                                                      {
                                                        fileName: _jsxFileName$3,
                                                        lineNumber: 495,
                                                        columnNumber: 25,
                                                      },
                                                      this,
                                                    ),
                                                  ],
                                                },
                                                void 0,
                                                true,
                                                {
                                                  fileName: _jsxFileName$3,
                                                  lineNumber: 493,
                                                  columnNumber: 23,
                                                },
                                                this,
                                              ),
                                              /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                                "div",
                                                {
                                                  className:
                                                    "text-[10px] text-muted-foreground font-mono mt-0.5",
                                                  children: new Date(
                                                    r.createdAt,
                                                  ).toLocaleTimeString("pt-BR", {
                                                    hour: "2-digit",
                                                    minute: "2-digit",
                                                  }),
                                                },
                                                void 0,
                                                false,
                                                {
                                                  fileName: _jsxFileName$3,
                                                  lineNumber: 497,
                                                  columnNumber: 23,
                                                },
                                                this,
                                              ),
                                            ],
                                          },
                                          void 0,
                                          true,
                                          {
                                            fileName: _jsxFileName$3,
                                            lineNumber: 492,
                                            columnNumber: 21,
                                          },
                                          this,
                                        ),
                                        /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                          "td",
                                          {
                                            className: "py-3 px-3",
                                            children: [
                                              /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                                "div",
                                                {
                                                  className:
                                                    "font-semibold text-foreground flex items-center gap-1",
                                                  children: [
                                                    /* @__PURE__ */ (0,
                                                    import_jsx_dev_runtime.jsxDEV)(
                                                      User,
                                                      {
                                                        className:
                                                          "h-3 w-3 text-muted-foreground shrink-0",
                                                      },
                                                      void 0,
                                                      false,
                                                      {
                                                        fileName: _jsxFileName$3,
                                                        lineNumber: 507,
                                                        columnNumber: 25,
                                                      },
                                                      this,
                                                    ),
                                                    /* @__PURE__ */ (0,
                                                    import_jsx_dev_runtime.jsxDEV)(
                                                      "span",
                                                      {
                                                        className: "truncate max-w-[140px]",
                                                        children: r.nome || "Não informado",
                                                      },
                                                      void 0,
                                                      false,
                                                      {
                                                        fileName: _jsxFileName$3,
                                                        lineNumber: 508,
                                                        columnNumber: 25,
                                                      },
                                                      this,
                                                    ),
                                                  ],
                                                },
                                                void 0,
                                                true,
                                                {
                                                  fileName: _jsxFileName$3,
                                                  lineNumber: 506,
                                                  columnNumber: 23,
                                                },
                                                this,
                                              ),
                                              r.colaboradores &&
                                                /* @__PURE__ */ (void 0)(
                                                  "div",
                                                  {
                                                    className:
                                                      "text-[10px] text-muted-foreground truncate max-w-[140px]",
                                                    title: r.colaboradores,
                                                    children: ["Equipe: ", r.colaboradores],
                                                  },
                                                  void 0,
                                                  true,
                                                  {
                                                    fileName: _jsxFileName$3,
                                                    lineNumber: 511,
                                                    columnNumber: 25,
                                                  },
                                                  this,
                                                ),
                                            ],
                                          },
                                          void 0,
                                          true,
                                          {
                                            fileName: _jsxFileName$3,
                                            lineNumber: 505,
                                            columnNumber: 21,
                                          },
                                          this,
                                        ),
                                        /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                          "td",
                                          {
                                            className: "py-3 px-3",
                                            children: [
                                              /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                                "div",
                                                {
                                                  className:
                                                    "font-medium text-foreground truncate max-w-[180px] flex items-center gap-1.5",
                                                  title: r.atividade,
                                                  children: [
                                                    /* @__PURE__ */ (0,
                                                    import_jsx_dev_runtime.jsxDEV)(
                                                      "span",
                                                      {
                                                        className: "truncate",
                                                        children:
                                                          r.atividade || "Sem nome de atividade",
                                                      },
                                                      void 0,
                                                      false,
                                                      {
                                                        fileName: _jsxFileName$3,
                                                        lineNumber: 525,
                                                        columnNumber: 25,
                                                      },
                                                      this,
                                                    ),
                                                    r.photos &&
                                                      r.photos.length > 0 &&
                                                      /* @__PURE__ */ (void 0)(
                                                        "span",
                                                        {
                                                          className:
                                                            "inline-flex items-center gap-0.5 text-[10px] font-medium bg-primary/10 text-primary px-1.5 py-0.2 rounded shrink-0",
                                                          title: `${r.photos.length} foto(s) anexada(s)`,
                                                          children: [
                                                            /* @__PURE__ */ (void 0)(
                                                              Camera,
                                                              { className: "h-3 w-3" },
                                                              void 0,
                                                              false,
                                                              {
                                                                fileName: _jsxFileName$3,
                                                                lineNumber: 531,
                                                                columnNumber: 29,
                                                              },
                                                              this,
                                                            ),
                                                            /* @__PURE__ */ (void 0)(
                                                              "span",
                                                              { children: r.photos.length },
                                                              void 0,
                                                              false,
                                                              {
                                                                fileName: _jsxFileName$3,
                                                                lineNumber: 532,
                                                                columnNumber: 29,
                                                              },
                                                              this,
                                                            ),
                                                          ],
                                                        },
                                                        void 0,
                                                        true,
                                                        {
                                                          fileName: _jsxFileName$3,
                                                          lineNumber: 527,
                                                          columnNumber: 27,
                                                        },
                                                        this,
                                                      ),
                                                  ],
                                                },
                                                void 0,
                                                true,
                                                {
                                                  fileName: _jsxFileName$3,
                                                  lineNumber: 521,
                                                  columnNumber: 23,
                                                },
                                                this,
                                              ),
                                              r.empresa &&
                                                /* @__PURE__ */ (void 0)(
                                                  "div",
                                                  {
                                                    className:
                                                      "text-[10px] text-muted-foreground flex items-center gap-1 mt-0.5",
                                                    children: [
                                                      /* @__PURE__ */ (void 0)(
                                                        Building2,
                                                        { className: "h-3 w-3 shrink-0" },
                                                        void 0,
                                                        false,
                                                        {
                                                          fileName: _jsxFileName$3,
                                                          lineNumber: 538,
                                                          columnNumber: 27,
                                                        },
                                                        this,
                                                      ),
                                                      /* @__PURE__ */ (void 0)(
                                                        "span",
                                                        {
                                                          className: "truncate max-w-[160px]",
                                                          children: r.empresa,
                                                        },
                                                        void 0,
                                                        false,
                                                        {
                                                          fileName: _jsxFileName$3,
                                                          lineNumber: 539,
                                                          columnNumber: 27,
                                                        },
                                                        this,
                                                      ),
                                                    ],
                                                  },
                                                  void 0,
                                                  true,
                                                  {
                                                    fileName: _jsxFileName$3,
                                                    lineNumber: 537,
                                                    columnNumber: 25,
                                                  },
                                                  this,
                                                ),
                                            ],
                                          },
                                          void 0,
                                          true,
                                          {
                                            fileName: _jsxFileName$3,
                                            lineNumber: 520,
                                            columnNumber: 21,
                                          },
                                          this,
                                        ),
                                        /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                          "td",
                                          {
                                            className: "py-3 px-3",
                                            children: [
                                              /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                                "div",
                                                {
                                                  className:
                                                    "font-medium text-foreground flex items-center gap-1",
                                                  children: [
                                                    /* @__PURE__ */ (0,
                                                    import_jsx_dev_runtime.jsxDEV)(
                                                      MapPin,
                                                      {
                                                        className:
                                                          "h-3 w-3 text-amber-500 shrink-0",
                                                      },
                                                      void 0,
                                                      false,
                                                      {
                                                        fileName: _jsxFileName$3,
                                                        lineNumber: 546,
                                                        columnNumber: 25,
                                                      },
                                                      this,
                                                    ),
                                                    /* @__PURE__ */ (0,
                                                    import_jsx_dev_runtime.jsxDEV)(
                                                      "span",
                                                      { children: r.local || "Não inf." },
                                                      void 0,
                                                      false,
                                                      {
                                                        fileName: _jsxFileName$3,
                                                        lineNumber: 547,
                                                        columnNumber: 25,
                                                      },
                                                      this,
                                                    ),
                                                  ],
                                                },
                                                void 0,
                                                true,
                                                {
                                                  fileName: _jsxFileName$3,
                                                  lineNumber: 545,
                                                  columnNumber: 23,
                                                },
                                                this,
                                              ),
                                              /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                                "div",
                                                {
                                                  className:
                                                    "text-[10px] text-muted-foreground font-mono mt-0.5",
                                                  children: [
                                                    r.linha ? `Linha: ${r.linha}` : null,
                                                    r.via ? `Via: ${r.via}` : null,
                                                    r.om ? `OM: ${r.om}` : null,
                                                  ]
                                                    .filter(Boolean)
                                                    .join(" • "),
                                                },
                                                void 0,
                                                false,
                                                {
                                                  fileName: _jsxFileName$3,
                                                  lineNumber: 549,
                                                  columnNumber: 23,
                                                },
                                                this,
                                              ),
                                            ],
                                          },
                                          void 0,
                                          true,
                                          {
                                            fileName: _jsxFileName$3,
                                            lineNumber: 544,
                                            columnNumber: 21,
                                          },
                                          this,
                                        ),
                                        /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                          "td",
                                          {
                                            className:
                                              "py-3 px-3 text-center font-mono font-bold text-emerald-600",
                                            children: [r.totals?.producaoTotal || 0, " m"],
                                          },
                                          void 0,
                                          true,
                                          {
                                            fileName: _jsxFileName$3,
                                            lineNumber: 560,
                                            columnNumber: 21,
                                          },
                                          this,
                                        ),
                                        /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                          "td",
                                          {
                                            className:
                                              "py-3 px-3 text-center font-mono font-bold text-amber-600",
                                            children: r.totals?.defeitosTotal || 0,
                                          },
                                          void 0,
                                          false,
                                          {
                                            fileName: _jsxFileName$3,
                                            lineNumber: 564,
                                            columnNumber: 21,
                                          },
                                          this,
                                        ),
                                        /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                          "td",
                                          {
                                            className:
                                              "py-3 px-3 text-center font-mono text-muted-foreground whitespace-nowrap",
                                            children: formatDuration(r.totals?.jornada || 0),
                                          },
                                          void 0,
                                          false,
                                          {
                                            fileName: _jsxFileName$3,
                                            lineNumber: 568,
                                            columnNumber: 21,
                                          },
                                          this,
                                        ),
                                        /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                          "td",
                                          {
                                            className: "py-3 px-3 text-right whitespace-nowrap",
                                            children: /* @__PURE__ */ (0,
                                            import_jsx_dev_runtime.jsxDEV)(
                                              "div",
                                              {
                                                className: "flex items-center justify-end gap-1.5",
                                                children: [
                                                  /* @__PURE__ */ (0,
                                                  import_jsx_dev_runtime.jsxDEV)(
                                                    "button",
                                                    {
                                                      type: "button",
                                                      onClick: () => handleDownloadSingleExcel(r),
                                                      className:
                                                        "p-1.5 rounded-md border border-emerald-500/30 bg-emerald-50 text-emerald-700 hover:bg-emerald-100 dark:bg-emerald-950/40 dark:text-emerald-300 transition-colors",
                                                      title: "Baixar Relatório em Excel (.xlsx)",
                                                      children: /* @__PURE__ */ (0,
                                                      import_jsx_dev_runtime.jsxDEV)(
                                                        FileSpreadsheet,
                                                        { className: "h-3.5 w-3.5" },
                                                        void 0,
                                                        false,
                                                        {
                                                          fileName: _jsxFileName$3,
                                                          lineNumber: 581,
                                                          columnNumber: 27,
                                                        },
                                                        this,
                                                      ),
                                                    },
                                                    void 0,
                                                    false,
                                                    {
                                                      fileName: _jsxFileName$3,
                                                      lineNumber: 575,
                                                      columnNumber: 25,
                                                    },
                                                    this,
                                                  ),
                                                  /* @__PURE__ */ (0,
                                                  import_jsx_dev_runtime.jsxDEV)(
                                                    "button",
                                                    {
                                                      type: "button",
                                                      onClick: () => setSapModalReport(r),
                                                      className:
                                                        "p-1.5 rounded-md border border-emerald-600/40 bg-emerald-600/15 text-emerald-800 dark:text-emerald-300 hover:bg-emerald-600 hover:text-white transition-colors",
                                                      title:
                                                        "Gerar planilha de apropriação SAP IW41 (.xlsx)",
                                                      children: /* @__PURE__ */ (0,
                                                      import_jsx_dev_runtime.jsxDEV)(
                                                        FileSpreadsheet,
                                                        { className: "h-3.5 w-3.5" },
                                                        void 0,
                                                        false,
                                                        {
                                                          fileName: _jsxFileName$3,
                                                          lineNumber: 591,
                                                          columnNumber: 27,
                                                        },
                                                        this,
                                                      ),
                                                    },
                                                    void 0,
                                                    false,
                                                    {
                                                      fileName: _jsxFileName$3,
                                                      lineNumber: 585,
                                                      columnNumber: 25,
                                                    },
                                                    this,
                                                  ),
                                                  /* @__PURE__ */ (0,
                                                  import_jsx_dev_runtime.jsxDEV)(
                                                    "button",
                                                    {
                                                      type: "button",
                                                      onClick: () => handleDownloadSinglePdf(r),
                                                      className:
                                                        "p-1.5 rounded-md border border-sky-500/30 bg-sky-50 text-sky-700 hover:bg-sky-100 dark:bg-sky-950/40 dark:text-sky-300 transition-colors",
                                                      title: "Baixar Relatório em PDF",
                                                      children: /* @__PURE__ */ (0,
                                                      import_jsx_dev_runtime.jsxDEV)(
                                                        FileText,
                                                        { className: "h-3.5 w-3.5" },
                                                        void 0,
                                                        false,
                                                        {
                                                          fileName: _jsxFileName$3,
                                                          lineNumber: 601,
                                                          columnNumber: 27,
                                                        },
                                                        this,
                                                      ),
                                                    },
                                                    void 0,
                                                    false,
                                                    {
                                                      fileName: _jsxFileName$3,
                                                      lineNumber: 595,
                                                      columnNumber: 25,
                                                    },
                                                    this,
                                                  ),
                                                  /* @__PURE__ */ (0,
                                                  import_jsx_dev_runtime.jsxDEV)(
                                                    "button",
                                                    {
                                                      type: "button",
                                                      onClick: () => setPreviewReport(r),
                                                      className:
                                                        "p-1.5 rounded-md border border-border bg-secondary hover:bg-accent text-foreground transition-colors",
                                                      title: "Visualizar Detalhes",
                                                      children: /* @__PURE__ */ (0,
                                                      import_jsx_dev_runtime.jsxDEV)(
                                                        Eye,
                                                        { className: "h-3.5 w-3.5" },
                                                        void 0,
                                                        false,
                                                        {
                                                          fileName: _jsxFileName$3,
                                                          lineNumber: 611,
                                                          columnNumber: 27,
                                                        },
                                                        this,
                                                      ),
                                                    },
                                                    void 0,
                                                    false,
                                                    {
                                                      fileName: _jsxFileName$3,
                                                      lineNumber: 605,
                                                      columnNumber: 25,
                                                    },
                                                    this,
                                                  ),
                                                  onLoadReportToForm &&
                                                    /* @__PURE__ */ (void 0)(
                                                      "button",
                                                      {
                                                        type: "button",
                                                        onClick: () => {
                                                          onLoadReportToForm(r);
                                                          showToast(
                                                            "Relatório carregado na tela principal!",
                                                          );
                                                        },
                                                        className:
                                                          "px-2 py-1 rounded-md text-[10px] font-semibold bg-primary/10 text-primary hover:bg-primary hover:text-white transition-colors",
                                                        title:
                                                          "Carregar dados deste relatório na tela principal",
                                                        children: "Editar",
                                                      },
                                                      void 0,
                                                      false,
                                                      {
                                                        fileName: _jsxFileName$3,
                                                        lineNumber: 616,
                                                        columnNumber: 27,
                                                      },
                                                      this,
                                                    ),
                                                  /* @__PURE__ */ (0,
                                                  import_jsx_dev_runtime.jsxDEV)(
                                                    "button",
                                                    {
                                                      type: "button",
                                                      onClick: () => handleDelete(r),
                                                      className:
                                                        "p-1.5 rounded-md border border-destructive/30 bg-destructive/10 text-destructive hover:bg-destructive hover:text-white transition-colors",
                                                      title: "Excluir da Nuvem",
                                                      children: /* @__PURE__ */ (0,
                                                      import_jsx_dev_runtime.jsxDEV)(
                                                        Trash2,
                                                        { className: "h-3.5 w-3.5" },
                                                        void 0,
                                                        false,
                                                        {
                                                          fileName: _jsxFileName$3,
                                                          lineNumber: 636,
                                                          columnNumber: 27,
                                                        },
                                                        this,
                                                      ),
                                                    },
                                                    void 0,
                                                    false,
                                                    {
                                                      fileName: _jsxFileName$3,
                                                      lineNumber: 630,
                                                      columnNumber: 25,
                                                    },
                                                    this,
                                                  ),
                                                ],
                                              },
                                              void 0,
                                              true,
                                              {
                                                fileName: _jsxFileName$3,
                                                lineNumber: 573,
                                                columnNumber: 23,
                                              },
                                              this,
                                            ),
                                          },
                                          void 0,
                                          false,
                                          {
                                            fileName: _jsxFileName$3,
                                            lineNumber: 572,
                                            columnNumber: 21,
                                          },
                                          this,
                                        ),
                                      ],
                                    },
                                    r.id,
                                    true,
                                    {
                                      fileName: _jsxFileName$3,
                                      lineNumber: 491,
                                      columnNumber: 19,
                                    },
                                    this,
                                  ),
                                ),
                              },
                              void 0,
                              false,
                              {
                                fileName: _jsxFileName$3,
                                lineNumber: 489,
                                columnNumber: 15,
                              },
                              this,
                            ),
                          ],
                        },
                        void 0,
                        true,
                        {
                          fileName: _jsxFileName$3,
                          lineNumber: 476,
                          columnNumber: 13,
                        },
                        this,
                      ),
                    },
                    void 0,
                    false,
                    {
                      fileName: _jsxFileName$3,
                      lineNumber: 475,
                      columnNumber: 11,
                    },
                    this,
                  ),
          },
          void 0,
          false,
          {
            fileName: _jsxFileName$3,
            lineNumber: 463,
            columnNumber: 7,
          },
          this,
        ),
        previewReport &&
          /* @__PURE__ */ (void 0)(
            "div",
            {
              className:
                "fixed inset-0 z-50 flex items-center justify-center bg-black/60 p-4 backdrop-blur-sm animate-in fade-in",
              children: /* @__PURE__ */ (void 0)(
                "div",
                {
                  className:
                    "relative w-full max-w-3xl max-h-[90vh] overflow-y-auto rounded-xl border border-border bg-card p-6 shadow-2xl space-y-5",
                  children: [
                    /* @__PURE__ */ (void 0)(
                      "div",
                      {
                        className: "flex items-center justify-between border-b border-border pb-3",
                        children: [
                          /* @__PURE__ */ (void 0)(
                            "div",
                            {
                              className: "flex items-center gap-2",
                              children: [
                                /* @__PURE__ */ (void 0)(
                                  Cloud,
                                  { className: "h-5 w-5 text-primary" },
                                  void 0,
                                  false,
                                  {
                                    fileName: _jsxFileName$3,
                                    lineNumber: 654,
                                    columnNumber: 17,
                                  },
                                  this,
                                ),
                                /* @__PURE__ */ (void 0)(
                                  "h3",
                                  {
                                    className: "text-base font-bold text-foreground",
                                    children: "Detalhes do Relatório em Nuvem",
                                  },
                                  void 0,
                                  false,
                                  {
                                    fileName: _jsxFileName$3,
                                    lineNumber: 655,
                                    columnNumber: 17,
                                  },
                                  this,
                                ),
                              ],
                            },
                            void 0,
                            true,
                            {
                              fileName: _jsxFileName$3,
                              lineNumber: 653,
                              columnNumber: 15,
                            },
                            this,
                          ),
                          /* @__PURE__ */ (void 0)(
                            "button",
                            {
                              type: "button",
                              onClick: () => setPreviewReport(null),
                              className:
                                "rounded-lg p-1.5 text-muted-foreground hover:bg-accent hover:text-foreground",
                              children: /* @__PURE__ */ (void 0)(
                                X,
                                { className: "h-5 w-5" },
                                void 0,
                                false,
                                {
                                  fileName: _jsxFileName$3,
                                  lineNumber: 664,
                                  columnNumber: 17,
                                },
                                this,
                              ),
                            },
                            void 0,
                            false,
                            {
                              fileName: _jsxFileName$3,
                              lineNumber: 659,
                              columnNumber: 15,
                            },
                            this,
                          ),
                        ],
                      },
                      void 0,
                      true,
                      {
                        fileName: _jsxFileName$3,
                        lineNumber: 652,
                        columnNumber: 13,
                      },
                      this,
                    ),
                    /* @__PURE__ */ (void 0)(
                      "div",
                      {
                        className:
                          "grid grid-cols-2 sm:grid-cols-3 gap-3 bg-secondary/40 p-4 rounded-lg text-xs",
                        children: [
                          /* @__PURE__ */ (void 0)(
                            "div",
                            {
                              children: [
                                /* @__PURE__ */ (void 0)(
                                  "span",
                                  {
                                    className:
                                      "text-muted-foreground block text-[10px] uppercase font-semibold",
                                    children: "Data do Relatório",
                                  },
                                  void 0,
                                  false,
                                  {
                                    fileName: _jsxFileName$3,
                                    lineNumber: 671,
                                    columnNumber: 17,
                                  },
                                  this,
                                ),
                                /* @__PURE__ */ (void 0)(
                                  "span",
                                  {
                                    className: "font-bold text-foreground font-mono",
                                    children: formatDate(previewReport.data),
                                  },
                                  void 0,
                                  false,
                                  {
                                    fileName: _jsxFileName$3,
                                    lineNumber: 674,
                                    columnNumber: 17,
                                  },
                                  this,
                                ),
                              ],
                            },
                            void 0,
                            true,
                            {
                              fileName: _jsxFileName$3,
                              lineNumber: 670,
                              columnNumber: 15,
                            },
                            this,
                          ),
                          /* @__PURE__ */ (void 0)(
                            "div",
                            {
                              children: [
                                /* @__PURE__ */ (void 0)(
                                  "span",
                                  {
                                    className:
                                      "text-muted-foreground block text-[10px] uppercase font-semibold",
                                    children: "Responsável",
                                  },
                                  void 0,
                                  false,
                                  {
                                    fileName: _jsxFileName$3,
                                    lineNumber: 679,
                                    columnNumber: 17,
                                  },
                                  this,
                                ),
                                /* @__PURE__ */ (void 0)(
                                  "span",
                                  {
                                    className: "font-bold text-foreground",
                                    children: previewReport.nome || "-",
                                  },
                                  void 0,
                                  false,
                                  {
                                    fileName: _jsxFileName$3,
                                    lineNumber: 682,
                                    columnNumber: 17,
                                  },
                                  this,
                                ),
                              ],
                            },
                            void 0,
                            true,
                            {
                              fileName: _jsxFileName$3,
                              lineNumber: 678,
                              columnNumber: 15,
                            },
                            this,
                          ),
                          /* @__PURE__ */ (void 0)(
                            "div",
                            {
                              children: [
                                /* @__PURE__ */ (void 0)(
                                  "span",
                                  {
                                    className:
                                      "text-muted-foreground block text-[10px] uppercase font-semibold",
                                    children: "Local",
                                  },
                                  void 0,
                                  false,
                                  {
                                    fileName: _jsxFileName$3,
                                    lineNumber: 685,
                                    columnNumber: 17,
                                  },
                                  this,
                                ),
                                /* @__PURE__ */ (void 0)(
                                  "span",
                                  {
                                    className: "font-bold text-foreground",
                                    children: previewReport.local || "-",
                                  },
                                  void 0,
                                  false,
                                  {
                                    fileName: _jsxFileName$3,
                                    lineNumber: 688,
                                    columnNumber: 17,
                                  },
                                  this,
                                ),
                              ],
                            },
                            void 0,
                            true,
                            {
                              fileName: _jsxFileName$3,
                              lineNumber: 684,
                              columnNumber: 15,
                            },
                            this,
                          ),
                          /* @__PURE__ */ (void 0)(
                            "div",
                            {
                              children: [
                                /* @__PURE__ */ (void 0)(
                                  "span",
                                  {
                                    className:
                                      "text-muted-foreground block text-[10px] uppercase font-semibold",
                                    children: "Linha / Via",
                                  },
                                  void 0,
                                  false,
                                  {
                                    fileName: _jsxFileName$3,
                                    lineNumber: 691,
                                    columnNumber: 17,
                                  },
                                  this,
                                ),
                                /* @__PURE__ */ (void 0)(
                                  "span",
                                  {
                                    className: "font-medium text-foreground",
                                    children: [
                                      previewReport.linha || "-",
                                      " / ",
                                      previewReport.via || "-",
                                    ],
                                  },
                                  void 0,
                                  true,
                                  {
                                    fileName: _jsxFileName$3,
                                    lineNumber: 694,
                                    columnNumber: 17,
                                  },
                                  this,
                                ),
                              ],
                            },
                            void 0,
                            true,
                            {
                              fileName: _jsxFileName$3,
                              lineNumber: 690,
                              columnNumber: 15,
                            },
                            this,
                          ),
                          /* @__PURE__ */ (void 0)(
                            "div",
                            {
                              children: [
                                /* @__PURE__ */ (void 0)(
                                  "span",
                                  {
                                    className:
                                      "text-muted-foreground block text-[10px] uppercase font-semibold",
                                    children: "GDA / OM",
                                  },
                                  void 0,
                                  false,
                                  {
                                    fileName: _jsxFileName$3,
                                    lineNumber: 699,
                                    columnNumber: 17,
                                  },
                                  this,
                                ),
                                /* @__PURE__ */ (void 0)(
                                  "span",
                                  {
                                    className: "font-medium text-foreground",
                                    children: [
                                      previewReport.gda || "-",
                                      " / ",
                                      previewReport.om || "-",
                                    ],
                                  },
                                  void 0,
                                  true,
                                  {
                                    fileName: _jsxFileName$3,
                                    lineNumber: 702,
                                    columnNumber: 17,
                                  },
                                  this,
                                ),
                              ],
                            },
                            void 0,
                            true,
                            {
                              fileName: _jsxFileName$3,
                              lineNumber: 698,
                              columnNumber: 15,
                            },
                            this,
                          ),
                          /* @__PURE__ */ (void 0)(
                            "div",
                            {
                              children: [
                                /* @__PURE__ */ (void 0)(
                                  "span",
                                  {
                                    className:
                                      "text-muted-foreground block text-[10px] uppercase font-semibold",
                                    children: "Empresa",
                                  },
                                  void 0,
                                  false,
                                  {
                                    fileName: _jsxFileName$3,
                                    lineNumber: 707,
                                    columnNumber: 17,
                                  },
                                  this,
                                ),
                                /* @__PURE__ */ (void 0)(
                                  "span",
                                  {
                                    className: "font-medium text-foreground",
                                    children: previewReport.empresa || "-",
                                  },
                                  void 0,
                                  false,
                                  {
                                    fileName: _jsxFileName$3,
                                    lineNumber: 710,
                                    columnNumber: 17,
                                  },
                                  this,
                                ),
                              ],
                            },
                            void 0,
                            true,
                            {
                              fileName: _jsxFileName$3,
                              lineNumber: 706,
                              columnNumber: 15,
                            },
                            this,
                          ),
                          /* @__PURE__ */ (void 0)(
                            "div",
                            {
                              className: "col-span-2 sm:col-span-3",
                              children: [
                                /* @__PURE__ */ (void 0)(
                                  "span",
                                  {
                                    className:
                                      "text-muted-foreground block text-[10px] uppercase font-semibold",
                                    children: "Nome da Atividade",
                                  },
                                  void 0,
                                  false,
                                  {
                                    fileName: _jsxFileName$3,
                                    lineNumber: 713,
                                    columnNumber: 17,
                                  },
                                  this,
                                ),
                                /* @__PURE__ */ (void 0)(
                                  "span",
                                  {
                                    className: "font-bold text-primary text-sm",
                                    children: previewReport.atividade || "-",
                                  },
                                  void 0,
                                  false,
                                  {
                                    fileName: _jsxFileName$3,
                                    lineNumber: 716,
                                    columnNumber: 17,
                                  },
                                  this,
                                ),
                              ],
                            },
                            void 0,
                            true,
                            {
                              fileName: _jsxFileName$3,
                              lineNumber: 712,
                              columnNumber: 15,
                            },
                            this,
                          ),
                          previewReport.colaboradores &&
                            /* @__PURE__ */ (void 0)(
                              "div",
                              {
                                className: "col-span-2 sm:col-span-3",
                                children: [
                                  /* @__PURE__ */ (void 0)(
                                    "span",
                                    {
                                      className:
                                        "text-muted-foreground block text-[10px] uppercase font-semibold",
                                      children: "Colaboradores",
                                    },
                                    void 0,
                                    false,
                                    {
                                      fileName: _jsxFileName$3,
                                      lineNumber: 722,
                                      columnNumber: 19,
                                    },
                                    this,
                                  ),
                                  /* @__PURE__ */ (void 0)(
                                    "span",
                                    {
                                      className: "text-foreground",
                                      children: previewReport.colaboradores,
                                    },
                                    void 0,
                                    false,
                                    {
                                      fileName: _jsxFileName$3,
                                      lineNumber: 725,
                                      columnNumber: 19,
                                    },
                                    this,
                                  ),
                                ],
                              },
                              void 0,
                              true,
                              {
                                fileName: _jsxFileName$3,
                                lineNumber: 721,
                                columnNumber: 17,
                              },
                              this,
                            ),
                        ],
                      },
                      void 0,
                      true,
                      {
                        fileName: _jsxFileName$3,
                        lineNumber: 669,
                        columnNumber: 13,
                      },
                      this,
                    ),
                    /* @__PURE__ */ (void 0)(
                      "div",
                      {
                        children: [
                          /* @__PURE__ */ (void 0)(
                            "h4",
                            {
                              className:
                                "text-xs font-semibold uppercase tracking-wider text-muted-foreground mb-2",
                              children: [
                                "Linha do Tempo de Eventos (",
                                previewReport.events?.length || 0,
                                ")",
                              ],
                            },
                            void 0,
                            true,
                            {
                              fileName: _jsxFileName$3,
                              lineNumber: 732,
                              columnNumber: 15,
                            },
                            this,
                          ),
                          /* @__PURE__ */ (void 0)(
                            "div",
                            {
                              className:
                                "rounded-lg border border-border overflow-hidden text-xs max-h-60 overflow-y-auto",
                              children: /* @__PURE__ */ (void 0)(
                                "table",
                                {
                                  className: "w-full text-left",
                                  children: [
                                    /* @__PURE__ */ (void 0)(
                                      "thead",
                                      {
                                        className:
                                          "bg-secondary text-muted-foreground text-[10px] uppercase font-semibold",
                                        children: /* @__PURE__ */ (void 0)(
                                          "tr",
                                          {
                                            children: [
                                              /* @__PURE__ */ (void 0)(
                                                "th",
                                                {
                                                  className: "py-2 px-3",
                                                  children: "Início",
                                                },
                                                void 0,
                                                false,
                                                {
                                                  fileName: _jsxFileName$3,
                                                  lineNumber: 739,
                                                  columnNumber: 23,
                                                },
                                                this,
                                              ),
                                              /* @__PURE__ */ (void 0)(
                                                "th",
                                                {
                                                  className: "py-2 px-3",
                                                  children: "Fim",
                                                },
                                                void 0,
                                                false,
                                                {
                                                  fileName: _jsxFileName$3,
                                                  lineNumber: 740,
                                                  columnNumber: 23,
                                                },
                                                this,
                                              ),
                                              /* @__PURE__ */ (void 0)(
                                                "th",
                                                {
                                                  className: "py-2 px-3",
                                                  children: "Duração",
                                                },
                                                void 0,
                                                false,
                                                {
                                                  fileName: _jsxFileName$3,
                                                  lineNumber: 741,
                                                  columnNumber: 23,
                                                },
                                                this,
                                              ),
                                              /* @__PURE__ */ (void 0)(
                                                "th",
                                                {
                                                  className: "py-2 px-3",
                                                  children: "Evento",
                                                },
                                                void 0,
                                                false,
                                                {
                                                  fileName: _jsxFileName$3,
                                                  lineNumber: 742,
                                                  columnNumber: 23,
                                                },
                                                this,
                                              ),
                                              /* @__PURE__ */ (void 0)(
                                                "th",
                                                {
                                                  className: "py-2 px-3",
                                                  children: "Km / Prod.",
                                                },
                                                void 0,
                                                false,
                                                {
                                                  fileName: _jsxFileName$3,
                                                  lineNumber: 743,
                                                  columnNumber: 23,
                                                },
                                                this,
                                              ),
                                              /* @__PURE__ */ (void 0)(
                                                "th",
                                                {
                                                  className: "py-2 px-3",
                                                  children: "Descrição / Obs",
                                                },
                                                void 0,
                                                false,
                                                {
                                                  fileName: _jsxFileName$3,
                                                  lineNumber: 744,
                                                  columnNumber: 23,
                                                },
                                                this,
                                              ),
                                            ],
                                          },
                                          void 0,
                                          true,
                                          {
                                            fileName: _jsxFileName$3,
                                            lineNumber: 738,
                                            columnNumber: 21,
                                          },
                                          this,
                                        ),
                                      },
                                      void 0,
                                      false,
                                      {
                                        fileName: _jsxFileName$3,
                                        lineNumber: 737,
                                        columnNumber: 19,
                                      },
                                      this,
                                    ),
                                    /* @__PURE__ */ (void 0)(
                                      "tbody",
                                      {
                                        className: "divide-y divide-border",
                                        children: (() => {
                                          const sorted = sortAndNormalizeEvents(
                                            previewReport.events || [],
                                          );
                                          return sorted.map((ev, i) => {
                                            const next = sorted[i + 1];
                                            const startStr = formatTime(ev.timestamp);
                                            let endStr = "";
                                            let durStr = "";
                                            if (next) {
                                              endStr = formatTime(next.timestamp);
                                              durStr = formatDuration(
                                                Math.max(0, next.timestamp - ev.timestamp),
                                              );
                                            } else if (ev.kind === "fim_jornada") {
                                              endStr = formatTime(ev.timestamp);
                                              durStr = "00:00:00";
                                            } else {
                                              endStr = "Em andamento";
                                              durStr = "(em andamento)";
                                            }
                                            return /* @__PURE__ */ (void 0)(
                                              "tr",
                                              {
                                                className: "hover:bg-accent/30",
                                                children: [
                                                  /* @__PURE__ */ (void 0)(
                                                    "td",
                                                    {
                                                      className:
                                                        "py-2 px-3 font-mono text-[11px] whitespace-nowrap",
                                                      children: startStr,
                                                    },
                                                    void 0,
                                                    false,
                                                    {
                                                      fileName: _jsxFileName$3,
                                                      lineNumber: 767,
                                                      columnNumber: 29,
                                                    },
                                                    this,
                                                  ),
                                                  /* @__PURE__ */ (void 0)(
                                                    "td",
                                                    {
                                                      className:
                                                        "py-2 px-3 font-mono text-[11px] whitespace-nowrap text-muted-foreground",
                                                      children: endStr,
                                                    },
                                                    void 0,
                                                    false,
                                                    {
                                                      fileName: _jsxFileName$3,
                                                      lineNumber: 770,
                                                      columnNumber: 29,
                                                    },
                                                    this,
                                                  ),
                                                  /* @__PURE__ */ (void 0)(
                                                    "td",
                                                    {
                                                      className:
                                                        "py-2 px-3 font-mono text-[11px] whitespace-nowrap font-semibold",
                                                      children: durStr,
                                                    },
                                                    void 0,
                                                    false,
                                                    {
                                                      fileName: _jsxFileName$3,
                                                      lineNumber: 773,
                                                      columnNumber: 29,
                                                    },
                                                    this,
                                                  ),
                                                  /* @__PURE__ */ (void 0)(
                                                    "td",
                                                    {
                                                      className: "py-2 px-3 font-medium",
                                                      children: EVENT_LABEL[ev.kind] || ev.kind,
                                                    },
                                                    void 0,
                                                    false,
                                                    {
                                                      fileName: _jsxFileName$3,
                                                      lineNumber: 776,
                                                      columnNumber: 29,
                                                    },
                                                    this,
                                                  ),
                                                  /* @__PURE__ */ (void 0)(
                                                    "td",
                                                    {
                                                      className: "py-2 px-3 font-mono",
                                                      children:
                                                        [
                                                          ev.kmInicial
                                                            ? `km ${ev.kmInicial}`
                                                            : null,
                                                          ev.producao ? `${ev.producao}m` : null,
                                                          ev.defeitos ? `${ev.defeitos}def` : null,
                                                        ]
                                                          .filter(Boolean)
                                                          .join(" • ") || "-",
                                                    },
                                                    void 0,
                                                    false,
                                                    {
                                                      fileName: _jsxFileName$3,
                                                      lineNumber: 779,
                                                      columnNumber: 29,
                                                    },
                                                    this,
                                                  ),
                                                  /* @__PURE__ */ (void 0)(
                                                    "td",
                                                    {
                                                      className: "py-2 px-3 text-muted-foreground",
                                                      children:
                                                        [ev.descricao, ev.observacoes]
                                                          .filter(Boolean)
                                                          .join(" — ") || "-",
                                                    },
                                                    void 0,
                                                    false,
                                                    {
                                                      fileName: _jsxFileName$3,
                                                      lineNumber: 788,
                                                      columnNumber: 29,
                                                    },
                                                    this,
                                                  ),
                                                ],
                                              },
                                              ev.id,
                                              true,
                                              {
                                                fileName: _jsxFileName$3,
                                                lineNumber: 766,
                                                columnNumber: 27,
                                              },
                                              this,
                                            );
                                          });
                                        })(),
                                      },
                                      void 0,
                                      false,
                                      {
                                        fileName: _jsxFileName$3,
                                        lineNumber: 747,
                                        columnNumber: 19,
                                      },
                                      this,
                                    ),
                                  ],
                                },
                                void 0,
                                true,
                                {
                                  fileName: _jsxFileName$3,
                                  lineNumber: 736,
                                  columnNumber: 17,
                                },
                                this,
                              ),
                            },
                            void 0,
                            false,
                            {
                              fileName: _jsxFileName$3,
                              lineNumber: 735,
                              columnNumber: 15,
                            },
                            this,
                          ),
                        ],
                      },
                      void 0,
                      true,
                      {
                        fileName: _jsxFileName$3,
                        lineNumber: 731,
                        columnNumber: 13,
                      },
                      this,
                    ),
                    previewReport.photos &&
                      previewReport.photos.length > 0 &&
                      /* @__PURE__ */ (void 0)(
                        "div",
                        {
                          className: "space-y-2",
                          children: [
                            /* @__PURE__ */ (void 0)(
                              "div",
                              {
                                className: "flex items-center gap-2",
                                children: [
                                  /* @__PURE__ */ (void 0)(
                                    Camera,
                                    { className: "h-4 w-4 text-primary" },
                                    void 0,
                                    false,
                                    {
                                      fileName: _jsxFileName$3,
                                      lineNumber: 804,
                                      columnNumber: 19,
                                    },
                                    this,
                                  ),
                                  /* @__PURE__ */ (void 0)(
                                    "h4",
                                    {
                                      className:
                                        "text-xs font-semibold uppercase tracking-wider text-muted-foreground",
                                      children: [
                                        "Anexo Fotográfico (",
                                        previewReport.photos.length,
                                        " ",
                                        previewReport.photos.length === 1 ? "foto" : "fotos",
                                        ")",
                                      ],
                                    },
                                    void 0,
                                    true,
                                    {
                                      fileName: _jsxFileName$3,
                                      lineNumber: 805,
                                      columnNumber: 19,
                                    },
                                    this,
                                  ),
                                ],
                              },
                              void 0,
                              true,
                              {
                                fileName: _jsxFileName$3,
                                lineNumber: 803,
                                columnNumber: 17,
                              },
                              this,
                            ),
                            /* @__PURE__ */ (void 0)(
                              "div",
                              {
                                className: "grid grid-cols-2 sm:grid-cols-3 gap-2.5",
                                children: previewReport.photos.map((p, idx) =>
                                  /* @__PURE__ */ (void 0)(
                                    "div",
                                    {
                                      className:
                                        "border border-border rounded-lg overflow-hidden bg-muted/30 p-1",
                                      children: [
                                        /* @__PURE__ */ (void 0)(
                                          "div",
                                          {
                                            className:
                                              "aspect-4/3 relative overflow-hidden rounded bg-black/5 flex items-center justify-center",
                                            children: [
                                              /* @__PURE__ */ (void 0)(
                                                "img",
                                                {
                                                  src: p.dataUrl,
                                                  alt: `Foto ${idx + 1}`,
                                                  className: "w-full h-full object-contain",
                                                  referrerPolicy: "no-referrer",
                                                },
                                                void 0,
                                                false,
                                                {
                                                  fileName: _jsxFileName$3,
                                                  lineNumber: 817,
                                                  columnNumber: 25,
                                                },
                                                this,
                                              ),
                                              /* @__PURE__ */ (void 0)(
                                                "span",
                                                {
                                                  className:
                                                    "absolute top-1 left-1 bg-background/90 text-foreground font-bold text-[9px] px-1.5 py-0.5 rounded shadow-xs",
                                                  children: [
                                                    "Foto ",
                                                    (idx + 1).toString().padStart(2, "0"),
                                                  ],
                                                },
                                                void 0,
                                                true,
                                                {
                                                  fileName: _jsxFileName$3,
                                                  lineNumber: 823,
                                                  columnNumber: 25,
                                                },
                                                this,
                                              ),
                                            ],
                                          },
                                          void 0,
                                          true,
                                          {
                                            fileName: _jsxFileName$3,
                                            lineNumber: 816,
                                            columnNumber: 23,
                                          },
                                          this,
                                        ),
                                        p.caption &&
                                          /* @__PURE__ */ (void 0)(
                                            "p",
                                            {
                                              className:
                                                "text-[10px] text-muted-foreground p-1 truncate",
                                              title: p.caption,
                                              children: p.caption,
                                            },
                                            void 0,
                                            false,
                                            {
                                              fileName: _jsxFileName$3,
                                              lineNumber: 828,
                                              columnNumber: 25,
                                            },
                                            this,
                                          ),
                                      ],
                                    },
                                    p.id || idx,
                                    true,
                                    {
                                      fileName: _jsxFileName$3,
                                      lineNumber: 812,
                                      columnNumber: 21,
                                    },
                                    this,
                                  ),
                                ),
                              },
                              void 0,
                              false,
                              {
                                fileName: _jsxFileName$3,
                                lineNumber: 810,
                                columnNumber: 17,
                              },
                              this,
                            ),
                          ],
                        },
                        void 0,
                        true,
                        {
                          fileName: _jsxFileName$3,
                          lineNumber: 802,
                          columnNumber: 15,
                        },
                        this,
                      ),
                    /* @__PURE__ */ (void 0)(
                      "div",
                      {
                        className:
                          "flex flex-wrap items-center justify-between gap-3 border-t border-border pt-4",
                        children: [
                          /* @__PURE__ */ (void 0)(
                            "div",
                            {
                              className: "text-xs text-muted-foreground font-mono",
                              children: ["ID: ", previewReport.id],
                            },
                            void 0,
                            true,
                            {
                              fileName: _jsxFileName$3,
                              lineNumber: 843,
                              columnNumber: 15,
                            },
                            this,
                          ),
                          /* @__PURE__ */ (void 0)(
                            "div",
                            {
                              className: "flex flex-wrap items-center gap-2",
                              children: [
                                /* @__PURE__ */ (void 0)(
                                  "button",
                                  {
                                    type: "button",
                                    onClick: () => setSapModalReport(previewReport),
                                    className:
                                      "btn-secondary inline-flex items-center gap-1.5 text-xs py-1.5 px-3 font-semibold text-emerald-700 dark:text-emerald-400 border-emerald-500/40",
                                    title:
                                      "Gerar planilha de apropriação de mão de obra para SAP IW41",
                                    children: [
                                      /* @__PURE__ */ (void 0)(
                                        FileSpreadsheet,
                                        { className: "h-4 w-4 text-emerald-600" },
                                        void 0,
                                        false,
                                        {
                                          fileName: _jsxFileName$3,
                                          lineNumber: 852,
                                          columnNumber: 19,
                                        },
                                        this,
                                      ),
                                      /* @__PURE__ */ (void 0)(
                                        "span",
                                        { children: "Apropriação SAP (IW41)" },
                                        void 0,
                                        false,
                                        {
                                          fileName: _jsxFileName$3,
                                          lineNumber: 853,
                                          columnNumber: 19,
                                        },
                                        this,
                                      ),
                                    ],
                                  },
                                  void 0,
                                  true,
                                  {
                                    fileName: _jsxFileName$3,
                                    lineNumber: 846,
                                    columnNumber: 17,
                                  },
                                  this,
                                ),
                                /* @__PURE__ */ (void 0)(
                                  "button",
                                  {
                                    type: "button",
                                    onClick: () => handleDownloadSingleExcel(previewReport),
                                    className:
                                      "btn-secondary inline-flex items-center gap-1.5 text-xs py-1.5 px-3",
                                    children: [
                                      /* @__PURE__ */ (void 0)(
                                        FileSpreadsheet,
                                        { className: "h-4 w-4 text-emerald-600" },
                                        void 0,
                                        false,
                                        {
                                          fileName: _jsxFileName$3,
                                          lineNumber: 861,
                                          columnNumber: 19,
                                        },
                                        this,
                                      ),
                                      /* @__PURE__ */ (void 0)(
                                        "span",
                                        { children: "Baixar Excel" },
                                        void 0,
                                        false,
                                        {
                                          fileName: _jsxFileName$3,
                                          lineNumber: 862,
                                          columnNumber: 19,
                                        },
                                        this,
                                      ),
                                    ],
                                  },
                                  void 0,
                                  true,
                                  {
                                    fileName: _jsxFileName$3,
                                    lineNumber: 856,
                                    columnNumber: 17,
                                  },
                                  this,
                                ),
                                /* @__PURE__ */ (void 0)(
                                  "button",
                                  {
                                    type: "button",
                                    onClick: () => handleDownloadSinglePdf(previewReport),
                                    className:
                                      "btn-primary inline-flex items-center gap-1.5 text-xs py-1.5 px-3",
                                    children: [
                                      /* @__PURE__ */ (void 0)(
                                        FileText,
                                        { className: "h-4 w-4" },
                                        void 0,
                                        false,
                                        {
                                          fileName: _jsxFileName$3,
                                          lineNumber: 870,
                                          columnNumber: 19,
                                        },
                                        this,
                                      ),
                                      /* @__PURE__ */ (void 0)(
                                        "span",
                                        { children: "Baixar PDF" },
                                        void 0,
                                        false,
                                        {
                                          fileName: _jsxFileName$3,
                                          lineNumber: 871,
                                          columnNumber: 19,
                                        },
                                        this,
                                      ),
                                    ],
                                  },
                                  void 0,
                                  true,
                                  {
                                    fileName: _jsxFileName$3,
                                    lineNumber: 865,
                                    columnNumber: 17,
                                  },
                                  this,
                                ),
                                /* @__PURE__ */ (void 0)(
                                  "button",
                                  {
                                    type: "button",
                                    onClick: () => setPreviewReport(null),
                                    className: "btn-ghost text-xs py-1.5 px-3",
                                    children: "Fechar",
                                  },
                                  void 0,
                                  false,
                                  {
                                    fileName: _jsxFileName$3,
                                    lineNumber: 874,
                                    columnNumber: 17,
                                  },
                                  this,
                                ),
                              ],
                            },
                            void 0,
                            true,
                            {
                              fileName: _jsxFileName$3,
                              lineNumber: 845,
                              columnNumber: 15,
                            },
                            this,
                          ),
                        ],
                      },
                      void 0,
                      true,
                      {
                        fileName: _jsxFileName$3,
                        lineNumber: 842,
                        columnNumber: 13,
                      },
                      this,
                    ),
                  ],
                },
                void 0,
                true,
                {
                  fileName: _jsxFileName$3,
                  lineNumber: 651,
                  columnNumber: 11,
                },
                this,
              ),
            },
            void 0,
            false,
            {
              fileName: _jsxFileName$3,
              lineNumber: 650,
              columnNumber: 9,
            },
            this,
          ),
        sapModalReport &&
          /* @__PURE__ */ (void 0)(
            SapIw41Modal,
            {
              isOpen: !!sapModalReport,
              onClose: () => setSapModalReport(null),
              header: cloudReportToHeaderAndEvents(sapModalReport).header,
              events: cloudReportToHeaderAndEvents(sapModalReport).events,
              onSuccess: (filename) =>
                showToast(`✓ Planilha SAP IW41 gerada com sucesso: ${filename}`),
            },
            void 0,
            false,
            {
              fileName: _jsxFileName$3,
              lineNumber: 888,
              columnNumber: 9,
            },
            this,
          ),
        /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
          SupabaseMigrationModal,
          {
            isOpen: isSupabaseModalOpen,
            onClose: () => setIsSupabaseModalOpen(false),
            onMigrationComplete: () => {
              showToast("✓ Migração para o Supabase concluída com sucesso!");
              handleManualSync();
            },
          },
          void 0,
          false,
          {
            fileName: _jsxFileName$3,
            lineNumber: 898,
            columnNumber: 7,
          },
          this,
        ),
      ],
    },
    void 0,
    true,
    {
      fileName: _jsxFileName$3,
      lineNumber: 219,
      columnNumber: 5,
    },
    this,
  );
}
var _jsxFileName$2 = "/app/applet/src/components/PhotoUploader.tsx";
/**
 * Resizes and compresses an image file to an optimized, lightweight JPEG dataURL.
 * Uses createImageBitmap if supported for speed and low memory, with robust fallbacks.
 */
async function compressImageSafely(file, maxDim = 720, quality = 0.65) {
  if (typeof window !== "undefined" && "createImageBitmap" in window)
    try {
      const bitmap = await createImageBitmap(file);
      let { width: w, height: h } = bitmap;
      if (w > maxDim || h > maxDim)
        if (w > h) {
          h = Math.round((h * maxDim) / w);
          w = maxDim;
        } else {
          w = Math.round((w * maxDim) / h);
          h = maxDim;
        }
      const canvas = document.createElement("canvas");
      canvas.width = Math.max(w, 1);
      canvas.height = Math.max(h, 1);
      const ctx = canvas.getContext("2d", { alpha: false });
      if (ctx) {
        ctx.fillStyle = "#ffffff";
        ctx.fillRect(0, 0, w, h);
        ctx.drawImage(bitmap, 0, 0, w, h);
        bitmap.close();
        const dataUrl = canvas.toDataURL("image/jpeg", quality);
        if (dataUrl && dataUrl.length > 200) return dataUrl;
      }
      bitmap.close();
    } catch (bitmapErr) {
      console.warn("createImageBitmap fallback to ObjectURL:", bitmapErr);
    }
  try {
    const dataUrl = await new Promise((resolve, reject) => {
      const objectUrl = URL.createObjectURL(file);
      const img = new Image();
      img.onload = () => {
        try {
          let w = img.naturalWidth || img.width;
          let h = img.naturalHeight || img.height;
          if (w > maxDim || h > maxDim)
            if (w > h) {
              h = Math.round((h * maxDim) / w);
              w = maxDim;
            } else {
              w = Math.round((w * maxDim) / h);
              h = maxDim;
            }
          const canvas = document.createElement("canvas");
          canvas.width = Math.max(w, 1);
          canvas.height = Math.max(h, 1);
          const ctx = canvas.getContext("2d", { alpha: false });
          if (!ctx) {
            URL.revokeObjectURL(objectUrl);
            reject(/* @__PURE__ */ new Error("Canvas context não disponível"));
            return;
          }
          ctx.fillStyle = "#ffffff";
          ctx.fillRect(0, 0, w, h);
          ctx.drawImage(img, 0, 0, w, h);
          URL.revokeObjectURL(objectUrl);
          resolve(canvas.toDataURL("image/jpeg", quality));
        } catch (canvasErr) {
          URL.revokeObjectURL(objectUrl);
          reject(canvasErr);
        }
      };
      img.onerror = () => {
        URL.revokeObjectURL(objectUrl);
        reject(/* @__PURE__ */ new Error("Falha ao decodificar imagem via ObjectURL"));
      };
      img.src = objectUrl;
    });
    if (dataUrl && dataUrl.length > 200) return dataUrl;
  } catch (objErr) {
    console.warn("ObjectURL fallback to FileReader:", objErr);
  }
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      const rawData = e.target?.result;
      if (!rawData) {
        reject(/* @__PURE__ */ new Error("Não foi possível ler os dados da foto."));
        return;
      }
      const img = new Image();
      img.onload = () => {
        let w = img.width;
        let h = img.height;
        if (w > maxDim || h > maxDim)
          if (w > h) {
            h = Math.round((h * maxDim) / w);
            w = maxDim;
          } else {
            w = Math.round((w * maxDim) / h);
            h = maxDim;
          }
        const canvas = document.createElement("canvas");
        canvas.width = Math.max(w, 1);
        canvas.height = Math.max(h, 1);
        const ctx = canvas.getContext("2d");
        if (!ctx) {
          resolve(rawData);
          return;
        }
        ctx.fillStyle = "#ffffff";
        ctx.fillRect(0, 0, w, h);
        ctx.drawImage(img, 0, 0, w, h);
        resolve(canvas.toDataURL("image/jpeg", quality));
      };
      img.onerror = () => {
        if (rawData.startsWith("data:image")) resolve(rawData);
        else reject(/* @__PURE__ */ new Error("Formato de imagem não suportado pelo navegador."));
      };
      img.src = rawData;
    };
    reader.onerror = () => reject(/* @__PURE__ */ new Error("Erro ao ler arquivo da foto."));
    reader.readAsDataURL(file);
  });
}
function PhotoUploader({ photos, onChange, maxPhotos = 12 }) {
  const galleryInputRef = (0, import_react.useRef)(null);
  const cameraInputRef = (0, import_react.useRef)(null);
  const [isProcessing, setIsProcessing] = (0, import_react.useState)(false);
  const [processingStatus, setProcessingStatus] = (0, import_react.useState)("");
  const [warningMessage, setWarningMessage] = (0, import_react.useState)(null);
  const [successMessage, setSuccessMessage] = (0, import_react.useState)(null);
  const [previewPhoto, setPreviewPhoto] = (0, import_react.useState)(null);
  const [isDragging, setIsDragging] = (0, import_react.useState)(false);
  const processFiles = async (files) => {
    if (!files || files.length === 0) return;
    setIsProcessing(true);
    setWarningMessage(null);
    setSuccessMessage(null);
    try {
      const availableSlots = maxPhotos - photos.length;
      if (availableSlots <= 0) {
        setWarningMessage(`O limite de ${maxPhotos} fotos já foi atingido.`);
        setIsProcessing(false);
        return;
      }
      const filesToProcess = files.slice(0, availableSlots);
      if (files.length > availableSlots)
        setWarningMessage(
          `Você selecionou ${files.length} fotos. Apenas ${availableSlots} ${availableSlots === 1 ? "foto foi adicionada" : "fotos foram adicionadas"} para respeitar o limite de ${maxPhotos} fotos.`,
        );
      const newPhotos = [];
      const failedFiles = [];
      for (let i = 0; i < filesToProcess.length; i++) {
        const file = filesToProcess[i];
        setProcessingStatus(
          `Processando foto ${i + 1} de ${filesToProcess.length} (${file.name || "imagem"})...`,
        );
        try {
          const dataUrl = await compressImageSafely(file, 720, 0.65);
          newPhotos.push({
            id: `photo-${Date.now()}-${Math.random().toString(36).substring(2, 7)}`,
            dataUrl,
            caption: "",
          });
        } catch (err) {
          console.error("Erro ao processar imagem:", file.name, err);
          failedFiles.push(file.name || `Foto #${i + 1}`);
        }
      }
      if (newPhotos.length > 0) {
        onChange([...photos, ...newPhotos]);
        setSuccessMessage(
          `${newPhotos.length} ${newPhotos.length === 1 ? "foto adicionada" : "fotos adicionadas"} com sucesso!`,
        );
        setTimeout(() => setSuccessMessage(null), 4e3);
      }
      if (failedFiles.length > 0)
        setWarningMessage(
          `Não foi possível carregar ${failedFiles.length} foto(s): ${failedFiles.join(", ")}. Tente fotos nos formatos JPG, PNG ou WEBP.`,
        );
    } catch (err) {
      console.error("Erro ao adicionar fotos:", err);
      setWarningMessage("Ocorreu um erro ao carregar as fotos. Tente novamente.");
    } finally {
      setIsProcessing(false);
      setProcessingStatus("");
      if (galleryInputRef.current) galleryInputRef.current.value = "";
      if (cameraInputRef.current) cameraInputRef.current.value = "";
    }
  };
  const handleFileSelect = async (e) => {
    const files = Array.from(e.target.files || []);
    await processFiles(files);
  };
  const handleDragOver = (e) => {
    e.preventDefault();
    e.stopPropagation();
    setIsDragging(true);
  };
  const handleDragLeave = (e) => {
    e.preventDefault();
    e.stopPropagation();
    setIsDragging(false);
  };
  const handleDrop = async (e) => {
    e.preventDefault();
    e.stopPropagation();
    setIsDragging(false);
    const droppedFiles = Array.from(e.dataTransfer.files || []).filter((f) =>
      f.type.startsWith("image/"),
    );
    if (droppedFiles.length > 0) await processFiles(droppedFiles);
  };
  const handleRemovePhoto = (id) => {
    onChange(photos.filter((p) => p.id !== id));
    setWarningMessage(null);
  };
  const handleUpdateCaption = (id, caption) => {
    onChange(
      photos.map((p) =>
        p.id === id
          ? {
              ...p,
              caption,
            }
          : p,
      ),
    );
  };
  return /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
    "div",
    {
      className: `bg-card border rounded-xl p-4 sm:p-5 space-y-4 shadow-sm transition-colors ${isDragging ? "border-primary bg-primary/5" : "border-border"}`,
      onDragOver: handleDragOver,
      onDragLeave: handleDragLeave,
      onDrop: handleDrop,
      children: [
        /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
          "div",
          {
            className:
              "flex flex-col sm:flex-row sm:items-center justify-between gap-3 pb-3 border-b border-border/60",
            children: [
              /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                "div",
                {
                  children: [
                    /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                      "div",
                      {
                        className: "flex items-center gap-2",
                        children: [
                          /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                            Camera,
                            { className: "w-5 h-5 text-primary shrink-0" },
                            void 0,
                            false,
                            {
                              fileName: _jsxFileName$2,
                              lineNumber: 298,
                              columnNumber: 13,
                            },
                            this,
                          ),
                          /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                            "h3",
                            {
                              className: "font-semibold text-foreground text-sm sm:text-base",
                              children: "Fotos para o Relatório PDF (Opcional)",
                            },
                            void 0,
                            false,
                            {
                              fileName: _jsxFileName$2,
                              lineNumber: 299,
                              columnNumber: 13,
                            },
                            this,
                          ),
                          /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                            "span",
                            {
                              className:
                                "text-xs px-2 py-0.5 rounded-full font-semibold bg-primary/10 text-primary border border-primary/20",
                              children: [photos.length, " de ", maxPhotos],
                            },
                            void 0,
                            true,
                            {
                              fileName: _jsxFileName$2,
                              lineNumber: 302,
                              columnNumber: 13,
                            },
                            this,
                          ),
                        ],
                      },
                      void 0,
                      true,
                      {
                        fileName: _jsxFileName$2,
                        lineNumber: 297,
                        columnNumber: 11,
                      },
                      this,
                    ),
                    /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                      "p",
                      {
                        className: "text-xs text-muted-foreground mt-1",
                        children: [
                          "Anexe de 1 a ",
                          maxPhotos,
                          " fotos. Elas são compactadas automaticamente e diagramadas sem distorção no PDF.",
                        ],
                      },
                      void 0,
                      true,
                      {
                        fileName: _jsxFileName$2,
                        lineNumber: 306,
                        columnNumber: 11,
                      },
                      this,
                    ),
                  ],
                },
                void 0,
                true,
                {
                  fileName: _jsxFileName$2,
                  lineNumber: 296,
                  columnNumber: 9,
                },
                this,
              ),
              /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                "div",
                {
                  className: "flex flex-wrap items-center gap-2",
                  children: [
                    /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                      "input",
                      {
                        type: "file",
                        ref: galleryInputRef,
                        onChange: handleFileSelect,
                        accept: "image/*,image/jpeg,image/png,image/webp,image/heic,image/heif",
                        multiple: true,
                        className: "hidden",
                        disabled: photos.length >= maxPhotos || isProcessing,
                      },
                      void 0,
                      false,
                      {
                        fileName: _jsxFileName$2,
                        lineNumber: 315,
                        columnNumber: 11,
                      },
                      this,
                    ),
                    /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                      "input",
                      {
                        type: "file",
                        ref: cameraInputRef,
                        onChange: handleFileSelect,
                        accept: "image/*",
                        capture: "environment",
                        className: "hidden",
                        disabled: photos.length >= maxPhotos || isProcessing,
                      },
                      void 0,
                      false,
                      {
                        fileName: _jsxFileName$2,
                        lineNumber: 324,
                        columnNumber: 11,
                      },
                      this,
                    ),
                    /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                      "button",
                      {
                        type: "button",
                        onClick: () => cameraInputRef.current?.click(),
                        disabled: photos.length >= maxPhotos || isProcessing,
                        className: `btn inline-flex items-center gap-1.5 text-xs sm:text-sm font-semibold px-3 py-2 rounded-lg transition-all ${photos.length >= maxPhotos || isProcessing ? "bg-muted text-muted-foreground cursor-not-allowed" : "bg-secondary text-secondary-foreground hover:bg-secondary/80 border border-border/60 shadow-xs"}`,
                        title: "Tirar foto direto com a câmera do celular",
                        children: [
                          /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                            Camera,
                            { className: "w-4 h-4 text-primary" },
                            void 0,
                            false,
                            {
                              fileName: _jsxFileName$2,
                              lineNumber: 346,
                              columnNumber: 13,
                            },
                            this,
                          ),
                          /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                            "span",
                            { children: "Câmera" },
                            void 0,
                            false,
                            {
                              fileName: _jsxFileName$2,
                              lineNumber: 347,
                              columnNumber: 13,
                            },
                            this,
                          ),
                        ],
                      },
                      void 0,
                      true,
                      {
                        fileName: _jsxFileName$2,
                        lineNumber: 335,
                        columnNumber: 11,
                      },
                      this,
                    ),
                    /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                      "button",
                      {
                        type: "button",
                        onClick: () => galleryInputRef.current?.click(),
                        disabled: photos.length >= maxPhotos || isProcessing,
                        className: `btn inline-flex items-center gap-1.5 text-xs sm:text-sm font-semibold px-3.5 py-2 rounded-lg transition-all ${photos.length >= maxPhotos || isProcessing ? "bg-muted text-muted-foreground cursor-not-allowed" : "bg-primary text-primary-foreground hover:bg-primary/90 shadow-xs"}`,
                        children: isProcessing
                          ? /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                              import_jsx_dev_runtime.Fragment,
                              {
                                children: [
                                  /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                    "div",
                                    {
                                      className:
                                        "w-4 h-4 border-2 border-primary-foreground border-t-transparent rounded-full animate-spin",
                                    },
                                    void 0,
                                    false,
                                    {
                                      fileName: _jsxFileName$2,
                                      lineNumber: 363,
                                      columnNumber: 17,
                                    },
                                    this,
                                  ),
                                  /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                    "span",
                                    { children: "Processando..." },
                                    void 0,
                                    false,
                                    {
                                      fileName: _jsxFileName$2,
                                      lineNumber: 364,
                                      columnNumber: 17,
                                    },
                                    this,
                                  ),
                                ],
                              },
                              void 0,
                              true,
                              {
                                fileName: _jsxFileName$2,
                                lineNumber: 362,
                                columnNumber: 15,
                              },
                              this,
                            )
                          : /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                              import_jsx_dev_runtime.Fragment,
                              {
                                children: [
                                  /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                    ImagePlus,
                                    { className: "w-4 h-4" },
                                    void 0,
                                    false,
                                    {
                                      fileName: _jsxFileName$2,
                                      lineNumber: 368,
                                      columnNumber: 17,
                                    },
                                    this,
                                  ),
                                  /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                    "span",
                                    {
                                      children:
                                        photos.length === 0
                                          ? "Adicionar Fotos"
                                          : photos.length >= maxPhotos
                                            ? "Limite Atingido"
                                            : "Mais Fotos",
                                    },
                                    void 0,
                                    false,
                                    {
                                      fileName: _jsxFileName$2,
                                      lineNumber: 369,
                                      columnNumber: 17,
                                    },
                                    this,
                                  ),
                                ],
                              },
                              void 0,
                              true,
                              {
                                fileName: _jsxFileName$2,
                                lineNumber: 367,
                                columnNumber: 15,
                              },
                              this,
                            ),
                      },
                      void 0,
                      false,
                      {
                        fileName: _jsxFileName$2,
                        lineNumber: 351,
                        columnNumber: 11,
                      },
                      this,
                    ),
                  ],
                },
                void 0,
                true,
                {
                  fileName: _jsxFileName$2,
                  lineNumber: 313,
                  columnNumber: 9,
                },
                this,
              ),
            ],
          },
          void 0,
          true,
          {
            fileName: _jsxFileName$2,
            lineNumber: 295,
            columnNumber: 7,
          },
          this,
        ),
        isProcessing &&
          processingStatus &&
          /* @__PURE__ */ (void 0)(
            "div",
            {
              className:
                "flex items-center gap-2.5 text-xs bg-primary/10 text-primary p-3 rounded-lg border border-primary/20 animate-pulse",
              children: [
                /* @__PURE__ */ (void 0)(
                  "div",
                  {
                    className:
                      "w-4 h-4 border-2 border-primary border-t-transparent rounded-full animate-spin shrink-0",
                  },
                  void 0,
                  false,
                  {
                    fileName: _jsxFileName$2,
                    lineNumber: 385,
                    columnNumber: 11,
                  },
                  this,
                ),
                /* @__PURE__ */ (void 0)(
                  "span",
                  {
                    className: "font-medium",
                    children: processingStatus,
                  },
                  void 0,
                  false,
                  {
                    fileName: _jsxFileName$2,
                    lineNumber: 386,
                    columnNumber: 11,
                  },
                  this,
                ),
              ],
            },
            void 0,
            true,
            {
              fileName: _jsxFileName$2,
              lineNumber: 384,
              columnNumber: 9,
            },
            this,
          ),
        successMessage &&
          /* @__PURE__ */ (void 0)(
            "div",
            {
              className:
                "flex items-center gap-2 text-xs bg-emerald-500/10 text-emerald-700 dark:text-emerald-300 p-3 rounded-lg border border-emerald-500/20 animate-in fade-in",
              children: [
                /* @__PURE__ */ (void 0)(
                  CircleCheck,
                  { className: "w-4 h-4 shrink-0 text-emerald-600 dark:text-emerald-400" },
                  void 0,
                  false,
                  {
                    fileName: _jsxFileName$2,
                    lineNumber: 393,
                    columnNumber: 11,
                  },
                  this,
                ),
                /* @__PURE__ */ (void 0)(
                  "span",
                  {
                    className: "font-medium",
                    children: successMessage,
                  },
                  void 0,
                  false,
                  {
                    fileName: _jsxFileName$2,
                    lineNumber: 394,
                    columnNumber: 11,
                  },
                  this,
                ),
              ],
            },
            void 0,
            true,
            {
              fileName: _jsxFileName$2,
              lineNumber: 392,
              columnNumber: 9,
            },
            this,
          ),
        warningMessage &&
          /* @__PURE__ */ (void 0)(
            "div",
            {
              className:
                "flex items-start gap-2 text-xs bg-amber-500/10 text-amber-700 dark:text-amber-400 p-3 rounded-lg border border-amber-500/20",
              children: [
                /* @__PURE__ */ (void 0)(
                  CircleAlert,
                  { className: "w-4 h-4 shrink-0 mt-0.5" },
                  void 0,
                  false,
                  {
                    fileName: _jsxFileName$2,
                    lineNumber: 401,
                    columnNumber: 11,
                  },
                  this,
                ),
                /* @__PURE__ */ (void 0)(
                  "span",
                  { children: warningMessage },
                  void 0,
                  false,
                  {
                    fileName: _jsxFileName$2,
                    lineNumber: 402,
                    columnNumber: 11,
                  },
                  this,
                ),
              ],
            },
            void 0,
            true,
            {
              fileName: _jsxFileName$2,
              lineNumber: 400,
              columnNumber: 9,
            },
            this,
          ),
        photos.length === 0
          ? /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
              "div",
              {
                onClick: () => galleryInputRef.current?.click(),
                className:
                  "text-center py-8 px-4 border-2 border-dashed border-border/80 rounded-xl bg-muted/20 hover:bg-muted/40 cursor-pointer transition-colors",
                children: [
                  /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                    CloudUpload,
                    { className: "w-9 h-9 text-primary mx-auto mb-2 opacity-80" },
                    void 0,
                    false,
                    {
                      fileName: _jsxFileName$2,
                      lineNumber: 412,
                      columnNumber: 11,
                    },
                    this,
                  ),
                  /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                    "p",
                    {
                      className: "text-xs sm:text-sm text-foreground font-semibold",
                      children: "Clique aqui ou arraste fotos para anexar ao relatório",
                    },
                    void 0,
                    false,
                    {
                      fileName: _jsxFileName$2,
                      lineNumber: 413,
                      columnNumber: 11,
                    },
                    this,
                  ),
                  /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                    "p",
                    {
                      className: "text-[11px] text-muted-foreground mt-1 max-w-md mx-auto",
                      children: [
                        "Você pode tirar fotos pelo celular ou escolher da galeria (formatos JPG, PNG, WEBP, até",
                        " ",
                        maxPhotos,
                        " fotos).",
                      ],
                    },
                    void 0,
                    true,
                    {
                      fileName: _jsxFileName$2,
                      lineNumber: 416,
                      columnNumber: 11,
                    },
                    this,
                  ),
                ],
              },
              void 0,
              true,
              {
                fileName: _jsxFileName$2,
                lineNumber: 408,
                columnNumber: 9,
              },
              this,
            )
          : /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
              "div",
              {
                className: "grid grid-cols-1 sm:grid-cols-2 md:grid-cols-3 gap-3.5",
                children: photos.map((photo, index) =>
                  /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                    "div",
                    {
                      className:
                        "group relative bg-background border border-border rounded-lg overflow-hidden flex flex-col shadow-xs transition-all hover:border-primary/50",
                      children: [
                        /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                          "div",
                          {
                            className:
                              "relative aspect-4/3 bg-muted/40 overflow-hidden flex items-center justify-center p-1",
                            children: [
                              /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                "img",
                                {
                                  src: photo.dataUrl,
                                  alt: `Foto ${index + 1}`,
                                  className: "w-full h-full object-contain rounded cursor-pointer",
                                  onClick: () => setPreviewPhoto(photo),
                                  referrerPolicy: "no-referrer",
                                },
                                void 0,
                                false,
                                {
                                  fileName: _jsxFileName$2,
                                  lineNumber: 429,
                                  columnNumber: 17,
                                },
                                this,
                              ),
                              /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                "span",
                                {
                                  className:
                                    "absolute top-2 left-2 bg-background/90 backdrop-blur-xs text-foreground font-bold text-[10px] px-2 py-0.5 rounded shadow-xs border border-border/60",
                                  children: ["Foto ", (index + 1).toString().padStart(2, "0")],
                                },
                                void 0,
                                true,
                                {
                                  fileName: _jsxFileName$2,
                                  lineNumber: 436,
                                  columnNumber: 17,
                                },
                                this,
                              ),
                              /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                "button",
                                {
                                  type: "button",
                                  onClick: () => setPreviewPhoto(photo),
                                  className:
                                    "absolute bottom-2 left-2 bg-background/80 text-foreground hover:bg-background p-1.5 rounded-md shadow-xs transition-transform opacity-80 hover:opacity-100",
                                  title: "Visualizar em tamanho maior",
                                  children: /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                    Maximize2,
                                    { className: "w-3.5 h-3.5" },
                                    void 0,
                                    false,
                                    {
                                      fileName: _jsxFileName$2,
                                      lineNumber: 447,
                                      columnNumber: 19,
                                    },
                                    this,
                                  ),
                                },
                                void 0,
                                false,
                                {
                                  fileName: _jsxFileName$2,
                                  lineNumber: 441,
                                  columnNumber: 17,
                                },
                                this,
                              ),
                              /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                "button",
                                {
                                  type: "button",
                                  onClick: () => handleRemovePhoto(photo.id),
                                  className:
                                    "absolute top-2 right-2 bg-destructive text-destructive-foreground hover:bg-destructive/90 p-1.5 rounded-full shadow-md transition-transform hover:scale-110",
                                  title: "Remover esta foto",
                                  children: /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                    Trash2,
                                    { className: "w-3.5 h-3.5" },
                                    void 0,
                                    false,
                                    {
                                      fileName: _jsxFileName$2,
                                      lineNumber: 457,
                                      columnNumber: 19,
                                    },
                                    this,
                                  ),
                                },
                                void 0,
                                false,
                                {
                                  fileName: _jsxFileName$2,
                                  lineNumber: 451,
                                  columnNumber: 17,
                                },
                                this,
                              ),
                            ],
                          },
                          void 0,
                          true,
                          {
                            fileName: _jsxFileName$2,
                            lineNumber: 428,
                            columnNumber: 15,
                          },
                          this,
                        ),
                        /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                          "div",
                          {
                            className: "p-2.5 flex-1 flex flex-col justify-between bg-card",
                            children: /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                              "input",
                              {
                                type: "text",
                                value: photo.caption || "",
                                onChange: (e) => handleUpdateCaption(photo.id, e.target.value),
                                placeholder: `Legenda da Foto ${(index + 1).toString().padStart(2, "0")} (opcional)...`,
                                className:
                                  "w-full text-xs px-2.5 py-1.5 bg-muted/30 border border-input rounded focus:outline-none focus:ring-1 focus:ring-primary focus:bg-background transition-colors",
                              },
                              void 0,
                              false,
                              {
                                fileName: _jsxFileName$2,
                                lineNumber: 462,
                                columnNumber: 17,
                              },
                              this,
                            ),
                          },
                          void 0,
                          false,
                          {
                            fileName: _jsxFileName$2,
                            lineNumber: 461,
                            columnNumber: 15,
                          },
                          this,
                        ),
                      ],
                    },
                    photo.id,
                    true,
                    {
                      fileName: _jsxFileName$2,
                      lineNumber: 424,
                      columnNumber: 13,
                    },
                    this,
                  ),
                ),
              },
              void 0,
              false,
              {
                fileName: _jsxFileName$2,
                lineNumber: 422,
                columnNumber: 9,
              },
              this,
            ),
        previewPhoto &&
          /* @__PURE__ */ (void 0)(
            "div",
            {
              className:
                "fixed inset-0 z-50 flex items-center justify-center bg-black/80 p-4 animate-in fade-in",
              onClick: () => setPreviewPhoto(null),
              children: /* @__PURE__ */ (void 0)(
                "div",
                {
                  className:
                    "relative max-w-3xl w-full bg-card border border-border rounded-xl p-4 shadow-2xl space-y-3",
                  onClick: (e) => e.stopPropagation(),
                  children: [
                    /* @__PURE__ */ (void 0)(
                      "div",
                      {
                        className: "flex items-center justify-between pb-2 border-b border-border",
                        children: [
                          /* @__PURE__ */ (void 0)(
                            "h4",
                            {
                              className: "font-semibold text-sm sm:text-base text-foreground",
                              children: "Visualização da Foto",
                            },
                            void 0,
                            false,
                            {
                              fileName: _jsxFileName$2,
                              lineNumber: 486,
                              columnNumber: 15,
                            },
                            this,
                          ),
                          /* @__PURE__ */ (void 0)(
                            "button",
                            {
                              type: "button",
                              onClick: () => setPreviewPhoto(null),
                              className:
                                "p-1.5 rounded-lg text-muted-foreground hover:text-foreground hover:bg-muted",
                              children: /* @__PURE__ */ (void 0)(
                                X,
                                { className: "w-5 h-5" },
                                void 0,
                                false,
                                {
                                  fileName: _jsxFileName$2,
                                  lineNumber: 494,
                                  columnNumber: 17,
                                },
                                this,
                              ),
                            },
                            void 0,
                            false,
                            {
                              fileName: _jsxFileName$2,
                              lineNumber: 489,
                              columnNumber: 15,
                            },
                            this,
                          ),
                        ],
                      },
                      void 0,
                      true,
                      {
                        fileName: _jsxFileName$2,
                        lineNumber: 485,
                        columnNumber: 13,
                      },
                      this,
                    ),
                    /* @__PURE__ */ (void 0)(
                      "div",
                      {
                        className:
                          "max-h-[70vh] flex items-center justify-center overflow-hidden rounded-lg bg-black/5",
                        children: /* @__PURE__ */ (void 0)(
                          "img",
                          {
                            src: previewPhoto.dataUrl,
                            alt: "Visualização",
                            className: "max-h-[65vh] w-auto max-w-full object-contain rounded",
                            referrerPolicy: "no-referrer",
                          },
                          void 0,
                          false,
                          {
                            fileName: _jsxFileName$2,
                            lineNumber: 499,
                            columnNumber: 15,
                          },
                          this,
                        ),
                      },
                      void 0,
                      false,
                      {
                        fileName: _jsxFileName$2,
                        lineNumber: 498,
                        columnNumber: 13,
                      },
                      this,
                    ),
                    previewPhoto.caption &&
                      /* @__PURE__ */ (void 0)(
                        "p",
                        {
                          className: "text-xs text-muted-foreground italic text-center",
                          children: ['"', previewPhoto.caption, '"'],
                        },
                        void 0,
                        true,
                        {
                          fileName: _jsxFileName$2,
                          lineNumber: 508,
                          columnNumber: 15,
                        },
                        this,
                      ),
                  ],
                },
                void 0,
                true,
                {
                  fileName: _jsxFileName$2,
                  lineNumber: 481,
                  columnNumber: 11,
                },
                this,
              ),
            },
            void 0,
            false,
            {
              fileName: _jsxFileName$2,
              lineNumber: 477,
              columnNumber: 9,
            },
            this,
          ),
      ],
    },
    void 0,
    true,
    {
      fileName: _jsxFileName$2,
      lineNumber: 287,
      columnNumber: 5,
    },
    this,
  );
}
function cn(...inputs) {
  return twMerge(clsx(inputs));
}
var _jsxFileName$1 = "/app/applet/src/components/ui/dropdown-menu.tsx";
var DropdownMenu = Root2;
var DropdownMenuTrigger = Trigger;
var DropdownMenuSubTrigger = import_react.forwardRef(
  ({ className, inset, children, ...props }, ref) =>
    /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
      SubTrigger2,
      {
        ref,
        className: cn(
          "flex cursor-default select-none items-center gap-2 rounded-sm px-2 py-1.5 text-sm outline-none focus:bg-accent data-[state=open]:bg-accent [&_svg]:pointer-events-none [&_svg]:size-4 [&_svg]:shrink-0",
          inset && "pl-8",
          className,
        ),
        ...props,
        children: [
          children,
          /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
            ChevronRight,
            { className: "ml-auto" },
            void 0,
            false,
            {
              fileName: _jsxFileName$1,
              lineNumber: 37,
              columnNumber: 5,
            },
            void 0,
          ),
        ],
      },
      void 0,
      true,
      {
        fileName: _jsxFileName$1,
        lineNumber: 27,
        columnNumber: 3,
      },
      void 0,
    ),
);
DropdownMenuSubTrigger.displayName = SubTrigger2.displayName;
var DropdownMenuSubContent = import_react.forwardRef(({ className, ...props }, ref) =>
  /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
    SubContent2,
    {
      ref,
      className: cn(
        "z-50 min-w-[8rem] overflow-hidden rounded-md border bg-popover p-1 text-popover-foreground shadow-lg data-[state=open]:animate-in data-[state=closed]:animate-out data-[state=closed]:fade-out-0 data-[state=open]:fade-in-0 data-[state=closed]:zoom-out-95 data-[state=open]:zoom-in-95 data-[side=bottom]:slide-in-from-top-2 data-[side=left]:slide-in-from-right-2 data-[side=right]:slide-in-from-left-2 data-[side=top]:slide-in-from-bottom-2 origin-(--radix-dropdown-menu-content-transform-origin)",
        className,
      ),
      ...props,
    },
    void 0,
    false,
    {
      fileName: _jsxFileName$1,
      lineNumber: 46,
      columnNumber: 3,
    },
    void 0,
  ),
);
DropdownMenuSubContent.displayName = SubContent2.displayName;
var DropdownMenuContent = import_react.forwardRef(({ className, sideOffset = 4, ...props }, ref) =>
  /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
    Portal2,
    {
      children: /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
        Content2,
        {
          ref,
          sideOffset,
          className: cn(
            "z-50 max-h-[var(--radix-dropdown-menu-content-available-height)] min-w-[8rem] overflow-y-auto overflow-x-hidden rounded-md border bg-popover p-1 text-popover-foreground shadow-md",
            "data-[state=open]:animate-in data-[state=closed]:animate-out data-[state=closed]:fade-out-0 data-[state=open]:fade-in-0 data-[state=closed]:zoom-out-95 data-[state=open]:zoom-in-95 data-[side=bottom]:slide-in-from-top-2 data-[side=left]:slide-in-from-right-2 data-[side=right]:slide-in-from-left-2 data-[side=top]:slide-in-from-bottom-2 origin-(--radix-dropdown-menu-content-transform-origin)",
            className,
          ),
          ...props,
        },
        void 0,
        false,
        {
          fileName: _jsxFileName$1,
          lineNumber: 62,
          columnNumber: 5,
        },
        void 0,
      ),
    },
    void 0,
    false,
    {
      fileName: _jsxFileName$1,
      lineNumber: 61,
      columnNumber: 3,
    },
    void 0,
  ),
);
DropdownMenuContent.displayName = Content2.displayName;
var DropdownMenuItem = import_react.forwardRef(({ className, inset, ...props }, ref) =>
  /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
    Item2,
    {
      ref,
      className: cn(
        "relative flex cursor-default select-none items-center gap-2 rounded-sm px-2 py-1.5 text-sm outline-none transition-colors focus:bg-accent focus:text-accent-foreground data-[disabled]:pointer-events-none data-[disabled]:opacity-50 [&>svg]:size-4 [&>svg]:shrink-0",
        inset && "pl-8",
        className,
      ),
      ...props,
    },
    void 0,
    false,
    {
      fileName: _jsxFileName$1,
      lineNumber: 82,
      columnNumber: 3,
    },
    void 0,
  ),
);
DropdownMenuItem.displayName = Item2.displayName;
var DropdownMenuCheckboxItem = import_react.forwardRef(
  ({ className, children, checked, ...props }, ref) =>
    /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
      CheckboxItem2,
      {
        ref,
        className: cn(
          "relative flex cursor-default select-none items-center rounded-sm py-1.5 pl-8 pr-2 text-sm outline-none transition-colors focus:bg-accent focus:text-accent-foreground data-[disabled]:pointer-events-none data-[disabled]:opacity-50",
          className,
        ),
        checked,
        ...props,
        children: [
          /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
            "span",
            {
              className: "absolute left-2 flex h-3.5 w-3.5 items-center justify-center",
              children: /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                ItemIndicator2,
                {
                  children: /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                    Check,
                    { className: "h-4 w-4" },
                    void 0,
                    false,
                    {
                      fileName: _jsxFileName$1,
                      lineNumber: 109,
                      columnNumber: 9,
                    },
                    void 0,
                  ),
                },
                void 0,
                false,
                {
                  fileName: _jsxFileName$1,
                  lineNumber: 108,
                  columnNumber: 7,
                },
                void 0,
              ),
            },
            void 0,
            false,
            {
              fileName: _jsxFileName$1,
              lineNumber: 107,
              columnNumber: 5,
            },
            void 0,
          ),
          children,
        ],
      },
      void 0,
      true,
      {
        fileName: _jsxFileName$1,
        lineNumber: 98,
        columnNumber: 3,
      },
      void 0,
    ),
);
DropdownMenuCheckboxItem.displayName = CheckboxItem2.displayName;
var DropdownMenuRadioItem = import_react.forwardRef(({ className, children, ...props }, ref) =>
  /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
    RadioItem2,
    {
      ref,
      className: cn(
        "relative flex cursor-default select-none items-center rounded-sm py-1.5 pl-8 pr-2 text-sm outline-none transition-colors focus:bg-accent focus:text-accent-foreground data-[disabled]:pointer-events-none data-[disabled]:opacity-50",
        className,
      ),
      ...props,
      children: [
        /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
          "span",
          {
            className: "absolute left-2 flex h-3.5 w-3.5 items-center justify-center",
            children: /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
              ItemIndicator2,
              {
                children: /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                  Circle,
                  { className: "h-2 w-2 fill-current" },
                  void 0,
                  false,
                  {
                    fileName: _jsxFileName$1,
                    lineNumber: 131,
                    columnNumber: 9,
                  },
                  void 0,
                ),
              },
              void 0,
              false,
              {
                fileName: _jsxFileName$1,
                lineNumber: 130,
                columnNumber: 7,
              },
              void 0,
            ),
          },
          void 0,
          false,
          {
            fileName: _jsxFileName$1,
            lineNumber: 129,
            columnNumber: 5,
          },
          void 0,
        ),
        children,
      ],
    },
    void 0,
    true,
    {
      fileName: _jsxFileName$1,
      lineNumber: 121,
      columnNumber: 3,
    },
    void 0,
  ),
);
DropdownMenuRadioItem.displayName = RadioItem2.displayName;
var DropdownMenuLabel = import_react.forwardRef(({ className, inset, ...props }, ref) =>
  /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
    Label2,
    {
      ref,
      className: cn("px-2 py-1.5 text-sm font-semibold", inset && "pl-8", className),
      ...props,
    },
    void 0,
    false,
    {
      fileName: _jsxFileName$1,
      lineNumber: 145,
      columnNumber: 3,
    },
    void 0,
  ),
);
DropdownMenuLabel.displayName = Label2.displayName;
var DropdownMenuSeparator = import_react.forwardRef(({ className, ...props }, ref) =>
  /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
    Separator2,
    {
      ref,
      className: cn("-mx-1 my-1 h-px bg-muted", className),
      ...props,
    },
    void 0,
    false,
    {
      fileName: _jsxFileName$1,
      lineNumber: 157,
      columnNumber: 3,
    },
    void 0,
  ),
);
DropdownMenuSeparator.displayName = Separator2.displayName;
var DropdownMenuShortcut = ({ className, ...props }) => {
  return /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
    "span",
    {
      className: cn("ml-auto text-xs tracking-widest opacity-60", className),
      ...props,
    },
    void 0,
    false,
    {
      fileName: _jsxFileName$1,
      lineNumber: 167,
      columnNumber: 5,
    },
    void 0,
  );
};
DropdownMenuShortcut.displayName = "DropdownMenuShortcut";
var KEY = "last_defect_rec";
var LEGACY_KEY = "last_defect";
var CHANNEL_NAME = "last_defect_bc_v3";
var currentRecord = {
  value: "",
  timestamp: 0,
};
if (typeof localStorage !== "undefined")
  try {
    const raw = localStorage.getItem(KEY);
    if (raw) {
      const parsed = JSON.parse(raw);
      if (parsed && typeof parsed.value === "string" && typeof parsed.timestamp === "number")
        currentRecord = parsed;
    } else {
      const legacy = localStorage.getItem(LEGACY_KEY);
      if (legacy)
        currentRecord = {
          value: legacy,
          timestamp: 1,
        };
    }
  } catch {}
var localSubscribers = /* @__PURE__ */ new Set();
function getBroadcastChannel() {
  if (typeof window !== "undefined" && "BroadcastChannel" in window)
    try {
      return new BroadcastChannel(CHANNEL_NAME);
    } catch {}
  return null;
}
function updateIfNewer(incoming) {
  if (!incoming || typeof incoming.value !== "string" || typeof incoming.timestamp !== "number")
    return false;
  if (incoming.timestamp > currentRecord.timestamp) {
    currentRecord = {
      value: incoming.value,
      timestamp: incoming.timestamp,
    };
    if (typeof localStorage !== "undefined")
      try {
        localStorage.setItem(KEY, JSON.stringify(currentRecord));
        localStorage.setItem(LEGACY_KEY, currentRecord.value);
      } catch {}
    localSubscribers.forEach((fn) => {
      try {
        fn(currentRecord.value, currentRecord);
      } catch {}
    });
    return true;
  }
  return false;
}
async function fetchLastDefect() {
  try {
    const snap = await getDoc(doc(db, "shared_state", "last_defect"));
    if (snap.exists()) {
      const d = snap.data();
      if (d && typeof d.value === "string" && typeof d.timestamp === "number")
        updateIfNewer({
          value: d.value,
          timestamp: d.timestamp,
        });
    }
  } catch {}
  try {
    const res = await fetch("/api/last-defect", { cache: "no-store" });
    if (res.ok) {
      const data = await res.json();
      if (data && typeof data.value === "string" && typeof data.timestamp === "number")
        updateIfNewer(data);
    }
  } catch {}
  try {
    const { data } = await supabase
      .from("shared_state")
      .select("value, timestamp, updated_at")
      .eq("key", "last_defect")
      .maybeSingle();
    if (data && typeof data.value === "string") {
      const rawTs = data.timestamp;
      let ts = typeof rawTs === "number" ? rawTs : 0;
      if (!ts && data.updated_at) ts = new Date(data.updated_at).getTime();
      updateIfNewer({
        value: data.value,
        timestamp: ts || 1,
      });
    }
  } catch {}
  return currentRecord.value;
}
async function saveLastDefect(value) {
  const newRec = {
    value: value.trim(),
    timestamp: Date.now(),
  };
  updateIfNewer(newRec);
  const bc = getBroadcastChannel();
  if (bc)
    try {
      bc.postMessage({
        type: "LAST_DEFECT_UPDATED",
        record: newRec,
      });
      bc.close();
    } catch {}
  try {
    await setDoc(
      doc(db, "shared_state", "last_defect"),
      {
        ...newRec,
        key: "last_defect",
      },
      { merge: true },
    );
  } catch (err) {
    console.warn("Firestore last_defect save error:", err);
  }
  try {
    const res = await fetch("/api/last-defect", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(newRec),
    });
    if (res.ok) {
      const data = await res.json();
      if (data && typeof data.value === "string" && typeof data.timestamp === "number")
        updateIfNewer(data);
    }
  } catch (err) {
    console.warn("API save warning:", err);
  }
  try {
    await supabase.from("shared_state").upsert(
      {
        key: "last_defect",
        value: newRec.value,
        timestamp: newRec.timestamp,
        updated_at: new Date(newRec.timestamp).toISOString(),
      },
      { onConflict: "key" },
    );
  } catch (err) {
    console.warn("Supabase save warning:", err);
  }
}
function subscribeLastDefect(onChange) {
  onChange(currentRecord.value);
  const subscriber = (val) => {
    onChange(val);
  };
  localSubscribers.add(subscriber);
  const bc = getBroadcastChannel();
  if (bc)
    bc.onmessage = (event) => {
      if (event.data?.type === "LAST_DEFECT_UPDATED" && event.data?.record)
        updateIfNewer(event.data.record);
    };
  const handleStorage = (e) => {
    if (e.key === KEY && typeof e.newValue === "string")
      try {
        const parsed = JSON.parse(e.newValue);
        if (parsed && typeof parsed.value === "string" && typeof parsed.timestamp === "number")
          updateIfNewer(parsed);
      } catch {}
  };
  if (typeof window !== "undefined") window.addEventListener("storage", handleStorage);
  let eventSource = null;
  if (typeof window !== "undefined" && "EventSource" in window)
    try {
      eventSource = new EventSource("/api/last-defect-stream");
      eventSource.onmessage = (event) => {
        try {
          const data = JSON.parse(event.data);
          if (data && typeof data.value === "string" && typeof data.timestamp === "number")
            updateIfNewer(data);
        } catch {}
      };
    } catch {}
  let unsubscribeFirestore = null;
  try {
    unsubscribeFirestore = onSnapshot(doc(db, "shared_state", "last_defect"), (snap) => {
      if (snap.exists()) {
        const d = snap.data();
        if (d && typeof d.value === "string" && typeof d.timestamp === "number")
          updateIfNewer({
            value: d.value,
            timestamp: d.timestamp,
          });
      }
    });
  } catch {}
  let supabaseChannel = null;
  try {
    supabaseChannel = supabase
      .channel("shared_state_last_defect_v3")
      .on(
        "postgres_changes",
        {
          event: "*",
          schema: "public",
          table: "shared_state",
          filter: "key=eq.last_defect",
        },
        (payload) => {
          const newRow = payload.new;
          if (newRow && typeof newRow.value === "string") {
            let ts = typeof newRow.timestamp === "number" ? newRow.timestamp : 0;
            if (!ts && newRow.updated_at) ts = new Date(newRow.updated_at).getTime();
            updateIfNewer({
              value: newRow.value,
              timestamp: ts || Date.now(),
            });
          }
        },
      )
      .subscribe();
  } catch {}
  const intervalId = setInterval(() => {
    fetchLastDefect().catch(() => {});
  }, 2e3);
  return () => {
    localSubscribers.delete(subscriber);
    if (unsubscribeFirestore)
      try {
        unsubscribeFirestore();
      } catch {}
    if (bc)
      try {
        bc.close();
      } catch {}
    if (typeof window !== "undefined") window.removeEventListener("storage", handleStorage);
    if (eventSource)
      try {
        eventSource.close();
      } catch {}
    if (supabaseChannel)
      try {
        supabase.removeChannel(supabaseChannel);
      } catch {}
    clearInterval(intervalId);
  };
}
var _jsxFileName = "/app/applet/src/routes/index.tsx?tsr-split=component";
var STORAGE_KEY = "parte-diaria-v1";
var ACTIVITIES_STORAGE_KEY = "activities-history-v1";
var COMPANIES_STORAGE_KEY = "companies-history-v1";
var TEAM_ROSTER_KEY = "team-roster-v1";
var LOCAL_OPTIONS = [
  "LUZ-BFU",
  "BFU-ABR",
  "ABR-LPA",
  "LPA-PQR",
  "PQR-PRT",
  "PRT-VCL",
  "VCL-JRG",
  "JRG-VAU",
  "VAU-PRU",
  "PRU-CAI",
  "CAI-FDR",
  "FDR-BFI",
  "BFI-FMO",
  "FMO-BTJ",
  "BTJ-CLP",
  "CLP-VPL",
  "VPL-JUN",
  "JUN-LOV",
];
function LocalSelectField({ value, onChange }) {
  const [isEditing, setIsEditing] = (0, import_react.useState)(false);
  const isPredefined = LOCAL_OPTIONS.includes(value);
  if (isEditing || (!isPredefined && value.trim() !== ""))
    return /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
      "div",
      {
        className: "relative flex items-center gap-1.5 w-full",
        children: [
          /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
            "input",
            {
              type: "text",
              value,
              onChange: (e) => onChange(e.target.value),
              placeholder: "Digite ou edite o local...",
              className: "input pr-16",
              autoFocus: true,
            },
            void 0,
            false,
            {
              fileName: _jsxFileName,
              lineNumber: 33,
              columnNumber: 9,
            },
            this,
          ),
          /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
            "button",
            {
              type: "button",
              onClick: () => setIsEditing(false),
              className:
                "absolute right-1.5 px-2 py-1 text-xs font-medium text-muted-foreground hover:text-foreground bg-secondary hover:bg-secondary/80 rounded transition-colors",
              title: "Ver lista de opções",
              children: "Opções",
            },
            void 0,
            false,
            {
              fileName: _jsxFileName,
              lineNumber: 34,
              columnNumber: 9,
            },
            this,
          ),
        ],
      },
      void 0,
      true,
      {
        fileName: _jsxFileName,
        lineNumber: 32,
        columnNumber: 12,
      },
      this,
    );
  return /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
    "div",
    {
      className: "flex items-center gap-1.5 w-full",
      children: [
        /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
          "select",
          {
            value,
            onChange: (e) => {
              if (e.target.value === "__EDIT_MANUALLY__") setIsEditing(true);
              else onChange(e.target.value);
            },
            className: "input cursor-pointer flex-1",
            children: [
              /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                "option",
                {
                  value: "",
                  children: "-- Selecione o Local --",
                },
                void 0,
                false,
                {
                  fileName: _jsxFileName,
                  lineNumber: 47,
                  columnNumber: 9,
                },
                this,
              ),
              LOCAL_OPTIONS.map((opt) =>
                /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                  "option",
                  {
                    value: opt,
                    children: opt,
                  },
                  opt,
                  false,
                  {
                    fileName: _jsxFileName,
                    lineNumber: 48,
                    columnNumber: 35,
                  },
                  this,
                ),
              ),
              /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                "option",
                {
                  value: "__EDIT_MANUALLY__",
                  children: "✏️ Editar texto manualmente...",
                },
                void 0,
                false,
                {
                  fileName: _jsxFileName,
                  lineNumber: 51,
                  columnNumber: 9,
                },
                this,
              ),
            ],
          },
          void 0,
          true,
          {
            fileName: _jsxFileName,
            lineNumber: 40,
            columnNumber: 7,
          },
          this,
        ),
        value &&
          /* @__PURE__ */ (void 0)(
            "button",
            {
              type: "button",
              onClick: () => setIsEditing(true),
              className:
                "p-2 border border-border rounded-md bg-background text-muted-foreground hover:bg-accent hover:text-accent-foreground transition-colors shrink-0",
              title: "Clique para editar o texto do item escolhido",
              "aria-label": "Editar texto do local",
              children: /* @__PURE__ */ (void 0)(
                Pencil,
                { className: "h-4 w-4" },
                void 0,
                false,
                {
                  fileName: _jsxFileName,
                  lineNumber: 54,
                  columnNumber: 11,
                },
                this,
              ),
            },
            void 0,
            false,
            {
              fileName: _jsxFileName,
              lineNumber: 53,
              columnNumber: 17,
            },
            this,
          ),
      ],
    },
    void 0,
    true,
    {
      fileName: _jsxFileName,
      lineNumber: 39,
      columnNumber: 10,
    },
    this,
  );
}
function ActivitySelectField({ value, onChange, activitiesHistory, onSaveHistory }) {
  const [isEditingManually, setIsEditingManually] = (0, import_react.useState)(false);
  const options = (0, import_react.useMemo)(() => {
    const combined = [
      ...[
        "Inspeção de Via Permanente",
        "Substituição de Dormentes",
        "Manutenção de AMV",
        "Alinhamento e Nivelamento de Via",
        "Roçada e Limpeza de Faixa de Domínio",
        "Inspeção e Esmerilhamento de Trilhos",
        "Manutenção de Sinalização e Subestação",
      ],
      ...activitiesHistory,
    ];
    if (value && value.trim() && !combined.includes(value.trim())) combined.push(value.trim());
    const set = new Set(combined.map((a) => a.trim()).filter(Boolean));
    return Array.from(set).sort((a, b) => a.localeCompare(b, "pt-BR"));
  }, [activitiesHistory, value]);
  const isInList = options.includes(value.trim());
  if (isEditingManually || (!isInList && value.trim() !== ""))
    return /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
      "div",
      {
        className: "relative flex items-center gap-1.5 w-full",
        children: [
          /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
            "input",
            {
              type: "text",
              value,
              onChange: (e) => onChange(e.target.value),
              onBlur: (e) => {
                if (e.target.value.trim()) onSaveHistory(e.target.value.trim());
              },
              placeholder: "Digite o nome da atividade...",
              className: "input pr-20",
              autoFocus: isEditingManually,
            },
            void 0,
            false,
            {
              fileName: _jsxFileName,
              lineNumber: 82,
              columnNumber: 9,
            },
            this,
          ),
          /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
            "button",
            {
              type: "button",
              onClick: () => setIsEditingManually(false),
              className:
                "absolute right-1.5 px-2 py-1 text-xs font-medium text-muted-foreground hover:text-foreground bg-secondary hover:bg-secondary/80 rounded transition-colors",
              title: "Ver lista de atividades cadastradas",
              children: "Opções",
            },
            void 0,
            false,
            {
              fileName: _jsxFileName,
              lineNumber: 87,
              columnNumber: 9,
            },
            this,
          ),
        ],
      },
      void 0,
      true,
      {
        fileName: _jsxFileName,
        lineNumber: 81,
        columnNumber: 12,
      },
      this,
    );
  return /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
    "div",
    {
      className: "flex items-center gap-1.5 w-full",
      children: [
        /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
          "select",
          {
            value,
            onChange: (e) => {
              if (e.target.value === "__NEW_ACTIVITY__") {
                setIsEditingManually(true);
                onChange("");
              } else {
                onChange(e.target.value);
                if (e.target.value.trim()) onSaveHistory(e.target.value.trim());
              }
            },
            className: "input cursor-pointer flex-1 font-medium",
            children: [
              /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                "option",
                {
                  value: "",
                  children: ["-- Selecione a atividade (", options.length, ") --"],
                },
                void 0,
                true,
                {
                  fileName: _jsxFileName,
                  lineNumber: 104,
                  columnNumber: 9,
                },
                this,
              ),
              options.map((act) =>
                /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                  "option",
                  {
                    value: act,
                    children: act,
                  },
                  act,
                  false,
                  {
                    fileName: _jsxFileName,
                    lineNumber: 105,
                    columnNumber: 29,
                  },
                  this,
                ),
              ),
              /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                "option",
                {
                  value: "__NEW_ACTIVITY__",
                  children: "✏️ Digitar outra atividade...",
                },
                void 0,
                false,
                {
                  fileName: _jsxFileName,
                  lineNumber: 108,
                  columnNumber: 9,
                },
                this,
              ),
            ],
          },
          void 0,
          true,
          {
            fileName: _jsxFileName,
            lineNumber: 93,
            columnNumber: 7,
          },
          this,
        ),
        /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
          "button",
          {
            type: "button",
            onClick: () => setIsEditingManually(true),
            className:
              "p-2 border border-border rounded-md bg-background text-muted-foreground hover:bg-accent hover:text-accent-foreground transition-colors shrink-0",
            title: "Editar ou digitar o nome da atividade manualmente",
            "aria-label": "Editar nome da atividade",
            children: /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
              Pencil,
              { className: "h-4 w-4" },
              void 0,
              false,
              {
                fileName: _jsxFileName,
                lineNumber: 111,
                columnNumber: 9,
              },
              this,
            ),
          },
          void 0,
          false,
          {
            fileName: _jsxFileName,
            lineNumber: 110,
            columnNumber: 7,
          },
          this,
        ),
      ],
    },
    void 0,
    true,
    {
      fileName: _jsxFileName,
      lineNumber: 92,
      columnNumber: 10,
    },
    this,
  );
}
var DEFAULT_ROSTER = [
  "Alessandro Pires Costa",
  "Cássio Messias da Silva",
  "Elicarlos Oliveira Irmão",
  "Kayque Freitas Fagundes",
  "Leandro Donizete Fontana",
  "Moacir Borges Junior",
];
var uid = () => Math.random().toString(36).slice(2) + Date.now().toString(36);
var pad2 = (n) => n.toString().padStart(2, "0");
var tsToLocalInput = (ts) => {
  const d = new Date(ts);
  return `${d.getFullYear()}-${pad2(d.getMonth() + 1)}-${pad2(d.getDate())}T${pad2(d.getHours())}:${pad2(d.getMinutes())}:${pad2(d.getSeconds())}`;
};
var parseLocalInput = (v) => {
  const t = new Date(v).getTime();
  return Number.isFinite(t) ? t : null;
};
function PartesDiariasPage() {
  const [state, setState] = (0, import_react.useState)({
    header: emptyHeader(),
    events: [],
    roster: DEFAULT_ROSTER,
  });
  const [photos, setPhotos] = (0, import_react.useState)([]);
  const [currentReportId, setCurrentReportId] = (0, import_react.useState)(void 0);
  const [lastCloudSyncTime, setLastCloudSyncTime] = (0, import_react.useState)(null);
  const [isSyncing, setIsSyncing] = (0, import_react.useState)(false);
  const [isOnline, setIsOnline] = (0, import_react.useState)(
    typeof window !== "undefined" ? navigator.onLine : true,
  );
  const [hasPendingSync, setHasPendingSync] = (0, import_react.useState)(false);
  const currentReportIdRef = (0, import_react.useRef)(currentReportId);
  (0, import_react.useEffect)(() => {
    currentReportIdRef.current = currentReportId;
  }, [currentReportId]);
  const [activeTab, setActiveTab] = (0, import_react.useState)("registro");
  const [isSapModalOpen, setIsSapModalOpen] = (0, import_react.useState)(false);
  const [saveNotification, setSaveNotification] = (0, import_react.useState)(null);
  const [pendingKind, setPendingKind] = (0, import_react.useState)(null);
  const [pendingDraft, setPendingDraft] = (0, import_react.useState)({});
  const [now, setNow] = (0, import_react.useState)(Date.now());
  const [newMember, setNewMember] = (0, import_react.useState)("");
  const [sharedLastDefect, setSharedLastDefect] = (0, import_react.useState)("");
  const [activitiesHistory, setActivitiesHistory] = (0, import_react.useState)([]);
  const [companiesHistory, setCompaniesHistory] = (0, import_react.useState)([]);
  const [editingDescEventId, setEditingDescEventId] = (0, import_react.useState)(null);
  const [editingDescValue, setEditingDescValue] = (0, import_react.useState)("");
  const syncToCloud = async (
    overrideEvents,
    overrideHeader,
    overridePhotos,
    customNotification,
  ) => {
    const targetEvents = overrideEvents ?? state.events;
    const targetHeader = overrideHeader ?? state.header;
    const targetPhotos = overridePhotos ?? photos;
    if (targetEvents.length === 0) return;
    try {
      localStorage.setItem(
        STORAGE_KEY,
        JSON.stringify({
          ...state,
          events: targetEvents,
          header: targetHeader,
          photos: targetPhotos,
          currentReportId: currentReportIdRef.current || currentReportId,
          hasPendingSync: true,
        }),
      );
    } catch {}
    if (typeof window !== "undefined" && !navigator.onLine) {
      setIsOnline(false);
      setHasPendingSync(true);
      setSaveNotification(
        customNotification ||
          "📱 Modo Offline: Seus dados estão salvos em segurança no celular. Serão enviados à nuvem quando houver conexão.",
      );
      setTimeout(() => setSaveNotification(null), 4e3);
      return;
    }
    setIsSyncing(true);
    try {
      const saved = await saveReportToCloud(
        targetHeader,
        targetEvents,
        currentReportIdRef.current || currentReportId,
        targetPhotos,
      );
      if (saved && saved.id) {
        setCurrentReportId(saved.id);
        currentReportIdRef.current = saved.id;
        setLastCloudSyncTime(Date.now());
        setHasPendingSync(false);
        setIsOnline(true);
        const nowStr = /* @__PURE__ */ new Date().toLocaleTimeString("pt-BR", {
          hour: "2-digit",
          minute: "2-digit",
          second: "2-digit",
        });
        const notifMsg = customNotification || `☁️ Salvo e sincronizado na nuvem às ${nowStr}`;
        setSaveNotification(notifMsg);
        setTimeout(() => setSaveNotification(null), 4e3);
      }
    } catch (err) {
      console.warn("Auto-save cloud warning:", err);
      setHasPendingSync(true);
      setSaveNotification(
        "📱 Sem sinal com a nuvem. Todos os dados continuam salvos com segurança no seu celular!",
      );
      setTimeout(() => setSaveNotification(null), 4e3);
    } finally {
      setIsSyncing(false);
    }
  };
  const handleHeaderRefreshSync = async () => {
    try {
      localStorage.setItem(
        STORAGE_KEY,
        JSON.stringify({
          ...state,
          photos,
          currentReportId,
          hasPendingSync,
        }),
      );
    } catch {}
    if (state.events.length === 0) {
      setSaveNotification(
        "📱 Dados salvos no celular. Registre ao menos um evento para sincronizar na nuvem.",
      );
      setTimeout(() => setSaveNotification(null), 4e3);
      return;
    }
    if (typeof window !== "undefined" && !navigator.onLine) {
      setIsOnline(false);
      setHasPendingSync(true);
      setSaveNotification(
        "📱 Sem internet no momento: Dados atualizados e salvos no celular com segurança!",
      );
      setTimeout(() => setSaveNotification(null), 4e3);
      return;
    }
    await syncToCloud(
      void 0,
      void 0,
      void 0,
      `☁️ Dados atualizados com sucesso na nuvem às ${/* @__PURE__ */ new Date().toLocaleTimeString("pt-BR")}!`,
    );
  };
  const handleSaveToCloud = async () => {
    if (state.events.length === 0) {
      alert(
        "Nenhum evento registrado ainda. Registre ao menos um evento antes de salvar o relatório na nuvem.",
      );
      return;
    }
    await syncToCloud(
      void 0,
      void 0,
      void 0,
      `Relatório de ${formatDate(state.header.data)} salvo no banco de dados na nuvem!`,
    );
  };
  const handleDownloadExcel = () => {
    if (state.events.length === 0) {
      alert("Nenhum evento registrado no relatório.");
      return;
    }
    exportSingleReportToExcel({
      id: `REL-${state.header.data || "hoje"}`,
      createdAt: Date.now(),
      updatedAt: Date.now(),
      data: state.header.data || "",
      nome: state.header.nome || "",
      local: state.header.local || "",
      sinais: state.header.sinais || "",
      linha: state.header.linha || "",
      via: state.header.via || "",
      gda: state.header.gda || "",
      om: state.header.om || "",
      empresa: state.header.empresa || "",
      atividade: state.header.atividade || "",
      colaboradores: state.header.colaboradores || "",
      descritivo: state.header.descritivo || "",
      kmProgramado: state.header.kmProgramado || "",
      executado: state.header.executado || "",
      observacaoGeral: state.header.observacaoGeral || "",
      telefone: state.header.telefone || "",
      events: state.events,
      photos,
      totals: totals(state.events),
    });
  };
  (0, import_react.useEffect)(() => {
    try {
      let persistentRoster = [];
      let rosterWasInStorage = false;
      const rawRoster = localStorage.getItem(TEAM_ROSTER_KEY);
      if (rawRoster !== null)
        try {
          const parsedRoster = JSON.parse(rawRoster);
          if (Array.isArray(parsedRoster)) {
            persistentRoster = parsedRoster;
            rosterWasInStorage = true;
          }
        } catch {}
      const raw = localStorage.getItem(STORAGE_KEY);
      if (raw) {
        const parsed = JSON.parse(raw);
        const stateRoster = Array.isArray(parsed.roster) ? parsed.roster : [];
        const colabNames =
          typeof parsed.header?.colaboradores === "string"
            ? parsed.header.colaboradores
                .split(";")
                .map((s) => s.trim())
                .filter(Boolean)
            : [];
        const finalRoster = rosterWasInStorage
          ? persistentRoster
          : Array.from(/* @__PURE__ */ new Set([...stateRoster, ...colabNames])).filter(Boolean);
        const loadedHeader = {
          ...emptyHeader(),
          ...(parsed.header ?? {}),
        };
        if (!loadedHeader.linha) loadedHeader.linha = "7";
        setState({
          header: loadedHeader,
          events: parsed.events ?? [],
          roster: finalRoster,
        });
        if (Array.isArray(parsed.photos)) setPhotos(parsed.photos);
        if (parsed.currentReportId) {
          setCurrentReportId(parsed.currentReportId);
          currentReportIdRef.current = parsed.currentReportId;
        }
        if (typeof parsed.hasPendingSync === "boolean") setHasPendingSync(parsed.hasPendingSync);
      } else if (rosterWasInStorage)
        setState((s) => ({
          ...s,
          roster: persistentRoster,
        }));
    } catch {}
  }, []);
  (0, import_react.useEffect)(() => {
    try {
      const raw = localStorage.getItem(ACTIVITIES_STORAGE_KEY);
      if (raw) {
        const parsed = JSON.parse(raw);
        if (Array.isArray(parsed)) setActivitiesHistory(parsed);
      }
    } catch {}
    return subscribeCloudReports((reports) => {
      const cloudActivities = reports.map((r) => r.atividade?.trim()).filter((a) => Boolean(a));
      if (cloudActivities.length > 0)
        setActivitiesHistory((prev) => {
          const combined = Array.from(/* @__PURE__ */ new Set([...prev, ...cloudActivities]));
          try {
            localStorage.setItem(ACTIVITIES_STORAGE_KEY, JSON.stringify(combined));
          } catch {}
          return combined;
        });
    });
  }, []);
  const saveActivityToHistory = (actName) => {
    const trimmed = actName.trim();
    if (!trimmed) return;
    setActivitiesHistory((prev) => {
      if (prev.includes(trimmed)) return prev;
      const updated = [...prev, trimmed];
      try {
        localStorage.setItem(ACTIVITIES_STORAGE_KEY, JSON.stringify(updated));
      } catch {}
      return updated;
    });
  };
  const clearActivitiesHistory = () => {
    if (window.confirm("Tem certeza que deseja apagar todo o histórico de atividades?")) {
      setActivitiesHistory([]);
      try {
        localStorage.removeItem(ACTIVITIES_STORAGE_KEY);
      } catch {}
    }
  };
  (0, import_react.useEffect)(() => {
    try {
      const raw = localStorage.getItem(COMPANIES_STORAGE_KEY);
      if (raw) {
        const parsed = JSON.parse(raw);
        if (Array.isArray(parsed)) setCompaniesHistory(parsed);
      }
    } catch {}
  }, []);
  const saveCompanyToHistory = (compName) => {
    const trimmed = compName.trim();
    if (!trimmed) return;
    setCompaniesHistory((prev) => {
      if (prev.includes(trimmed)) return prev;
      const updated = [...prev, trimmed];
      try {
        localStorage.setItem(COMPANIES_STORAGE_KEY, JSON.stringify(updated));
      } catch {}
      return updated;
    });
  };
  const clearCompaniesHistory = () => {
    if (window.confirm("Tem certeza que deseja apagar todo o histórico de empresas?")) {
      setCompaniesHistory([]);
      try {
        localStorage.removeItem(COMPANIES_STORAGE_KEY);
      } catch {}
    }
  };
  const clearRosterHistory = () => {
    if (window.confirm("Tem certeza que deseja apagar todo o histórico de colaboradores?")) {
      setState((s) => ({
        ...s,
        roster: [],
        header: {
          ...s.header,
          colaboradores: "",
        },
      }));
      try {
        localStorage.setItem(TEAM_ROSTER_KEY, JSON.stringify([]));
      } catch {}
    }
  };
  (0, import_react.useEffect)(() => {
    try {
      localStorage.setItem(
        STORAGE_KEY,
        JSON.stringify({
          ...state,
          photos,
          currentReportId,
          hasPendingSync,
        }),
      );
    } catch {
      try {
        localStorage.setItem(
          STORAGE_KEY,
          JSON.stringify({
            ...state,
            photos: [],
            currentReportId,
            hasPendingSync,
          }),
        );
      } catch {}
    }
    try {
      localStorage.setItem(TEAM_ROSTER_KEY, JSON.stringify(state.roster));
    } catch {}
  }, [state, photos, currentReportId, hasPendingSync]);
  (0, import_react.useEffect)(() => {
    const handleOnline = () => {
      setIsOnline(true);
      if (state.events.length > 0)
        syncToCloud(
          void 0,
          void 0,
          void 0,
          "📶 Conexão reestabelecida! Todos os dados foram sincronizados com a nuvem.",
        );
      else {
        setSaveNotification("📶 Conexão de internet reestabelecida.");
        setTimeout(() => setSaveNotification(null), 3e3);
      }
    };
    const handleOffline = () => {
      setIsOnline(false);
      setSaveNotification(
        "📱 Conexão com a internet perdida. Seus dados continuam sendo salvos em segurança no celular.",
      );
      setTimeout(() => setSaveNotification(null), 4e3);
    };
    window.addEventListener("online", handleOnline);
    window.addEventListener("offline", handleOffline);
    return () => {
      window.removeEventListener("online", handleOnline);
      window.removeEventListener("offline", handleOffline);
    };
  }, [state.events, state.header, photos]);
  (0, import_react.useEffect)(() => {
    if (state.events.length === 0) return;
    const timer = setTimeout(() => {
      syncToCloud();
    }, 1500);
    return () => clearTimeout(timer);
  }, [state.header, photos, state.events.length]);
  (0, import_react.useEffect)(() => {
    const id = setInterval(() => setNow(Date.now()), 1e3);
    return () => clearInterval(id);
  }, []);
  (0, import_react.useEffect)(() => {
    let cancelled = false;
    (async () => {
      try {
        const blob = await (await fetch(logoAsset.url)).blob();
        const reader = new FileReader();
        reader.onloadend = () => {
          if (!cancelled) setLogoDataUrl(reader.result);
        };
        reader.readAsDataURL(blob);
      } catch {}
    })();
    return () => {
      cancelled = true;
    };
  }, []);
  (0, import_react.useEffect)(() => {
    fetchLastDefect()
      .then(setSharedLastDefect)
      .catch(() => {});
    return subscribeLastDefect(setSharedLastDefect);
  }, []);
  const t = (0, import_react.useMemo)(() => totals(state.events), [state.events]);
  const sortedEvents = (0, import_react.useMemo)(
    () => [...state.events].sort((a, b) => a.timestamp - b.timestamp),
    [state.events],
  );
  const jornadaStarted = state.events.some((e) => e.kind === "inicio_jornada");
  const jornadaEnded = state.events.some((e) => e.kind === "fim_jornada");
  const openActivity =
    sortedEvents.filter((e) => e.kind === "inicio_atividade").length >
    sortedEvents.filter((e) => e.kind === "fim_atividade").length;
  const liveJornada = jornadaStarted && !jornadaEnded ? now - sortedEvents[0].timestamp : t.jornada;
  const registerEvent = (kind) => {
    setPendingKind(kind);
    setPendingDraft({});
  };
  const addEvent = (kind, draft) => {
    const ev = {
      id: uid(),
      kind,
      timestamp: Date.now(),
      ...draft,
    };
    const nextEvents = [...state.events, ev];
    setState((s) => ({
      ...s,
      events: nextEvents,
      header:
        draft.local && !s.header.local
          ? {
              ...s.header,
              local: draft.local,
            }
          : s.header,
    }));
    if (kind === "fim_jornada") {
      if (state.header.atividade && state.header.atividade.trim())
        saveActivityToHistory(state.header.atividade);
      if (state.header.empresa && state.header.empresa.trim())
        saveCompanyToHistory(state.header.empresa);
    }
    if (draft.ultimoDefeito && draft.ultimoDefeito.trim()) {
      const v = draft.ultimoDefeito.trim();
      setSharedLastDefect(v);
      saveLastDefect(v).catch(() => {});
    }
    setPendingKind(null);
    setPendingDraft({});
    syncToCloud(
      nextEvents,
      void 0,
      void 0,
      kind === "inicio_jornada"
        ? "☁️ Jornada iniciada e salva na nuvem com sucesso!"
        : kind === "fim_jornada"
          ? "☁️ Jornada encerrada e relatório salvo na nuvem!"
          : "☁️ Evento registrado e atualizado na nuvem!",
    );
  };
  const removeEvent = (id) => {
    const nextEvents = state.events.filter((e) => e.id !== id);
    setState((s) => ({
      ...s,
      events: nextEvents,
    }));
    if (nextEvents.length > 0)
      syncToCloud(nextEvents, void 0, void 0, "☁️ Evento removido e atualizado na nuvem!");
  };
  const updateEvent = (id, patch) => {
    const nextEvents = state.events.map((e) =>
      e.id === id
        ? {
            ...e,
            ...patch,
          }
        : e,
    );
    setState((s) => ({
      ...s,
      events: nextEvents,
    }));
    if (patch.ultimoDefeito && patch.ultimoDefeito.trim()) {
      const v = patch.ultimoDefeito.trim();
      setSharedLastDefect(v);
      saveLastDefect(v).catch(() => {});
    }
    syncToCloud(nextEvents, void 0, void 0, "☁️ Evento atualizado na nuvem!");
  };
  const resetDay = () => {
    if (confirm("Zerar todos os dados deste registro de atividade diária?")) {
      setState((s) => {
        const used = s.header.colaboradores
          .split(";")
          .map((n) => n.trim())
          .filter(Boolean);
        const mergedRoster = Array.from(/* @__PURE__ */ new Set([...s.roster, ...used]));
        return {
          header: emptyHeader(),
          events: [],
          roster: mergedRoster,
        };
      });
      setPhotos([]);
      setCurrentReportId(void 0);
      currentReportIdRef.current = void 0;
    }
  };
  const updateHeader = (patch) => {
    setState((s) => ({
      ...s,
      header: {
        ...s.header,
        ...patch,
      },
    }));
    if (patch.data && patch.data !== state.header.data) {
      const existing = getLocalCache().find((r) => r.data === patch.data);
      if (existing) {
        setCurrentReportId(existing.id);
        currentReportIdRef.current = existing.id;
      } else {
        setCurrentReportId(void 0);
        currentReportIdRef.current = void 0;
      }
    }
  };
  const report = {
    header: state.header,
    events: state.events,
  };
  const sendWhatsApp = () => {
    const msg = encodeURIComponent(buildWhatsAppMessage(report));
    const phone = state.header.telefone.replace(/\D/g, "");
    const url = phone ? `https://wa.me/${phone}?text=${msg}` : `https://wa.me/?text=${msg}`;
    window.open(url, "_blank");
  };
  const pdfFilename = () => generatePdfFilename(state.header, state.events);
  const downloadBlob = (blob, filename) => {
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = filename;
    a.click();
    URL.revokeObjectURL(url);
  };
  const buildPdfBlob = async () => {
    return buildPdfBlobForReport(state.header, state.events, photos);
  };
  const downloadPDF = async () => {
    downloadReportPdf(state.header, state.events, photos, pdfFilename());
  };
  const shareFileWhatsApp = async (blob, filename, mime) => {
    const file = new File([blob], filename, { type: mime });
    const nav = navigator;
    if (nav.canShare && nav.canShare({ files: [file] }))
      try {
        await nav.share({
          files: [file],
          title: "Registro de Atividade Diária",
          text: buildWhatsAppMessage(report),
        });
        return;
      } catch {}
    downloadBlob(blob, filename);
    const msg = encodeURIComponent(
      buildWhatsAppMessage(report) + `\n\n(Arquivo ${filename} baixado — anexe no WhatsApp)`,
    );
    const phone = state.header.telefone.replace(/\D/g, "");
    const url = phone ? `https://wa.me/${phone}?text=${msg}` : `https://wa.me/?text=${msg}`;
    window.open(url, "_blank");
  };
  const sharePdfWhatsApp = async () =>
    shareFileWhatsApp(await buildPdfBlob(), pdfFilename(), "application/pdf");
  const shareExcelWhatsApp = () => {
    const { blob, filename } = buildSingleReportExcelBlob({
      id: currentReportId || `REL-${state.header.data || "hoje"}`,
      createdAt: Date.now(),
      updatedAt: Date.now(),
      data: state.header.data || "",
      nome: state.header.nome || "",
      local: state.header.local || "",
      sinais: state.header.sinais || "",
      linha: state.header.linha || "",
      via: state.header.via || "",
      gda: state.header.gda || "",
      om: state.header.om || "",
      empresa: state.header.empresa || "",
      atividade: state.header.atividade || "",
      colaboradores: state.header.colaboradores || "",
      descritivo: state.header.descritivo || "",
      kmProgramado: state.header.kmProgramado || "",
      executado: state.header.executado || "",
      observacaoGeral: state.header.observacaoGeral || "",
      telefone: state.header.telefone || "",
      events: state.events,
      photos,
      totals: totals(state.events),
    });
    return shareFileWhatsApp(
      blob,
      filename,
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    );
  };
  return /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
    "main",
    {
      className: "min-h-screen bg-background text-foreground pb-12",
      children: [
        saveNotification &&
          /* @__PURE__ */ (void 0)(
            "div",
            {
              className:
                "fixed top-4 right-4 z-50 flex items-center gap-2 rounded-lg bg-emerald-600 px-4 py-3 text-sm font-medium text-white shadow-xl animate-in fade-in slide-in-from-top-2",
              children: [
                /* @__PURE__ */ (void 0)(
                  Check,
                  { className: "h-5 w-5 shrink-0" },
                  void 0,
                  false,
                  {
                    fileName: _jsxFileName,
                    lineNumber: 700,
                    columnNumber: 11,
                  },
                  this,
                ),
                /* @__PURE__ */ (void 0)(
                  "span",
                  { children: saveNotification },
                  void 0,
                  false,
                  {
                    fileName: _jsxFileName,
                    lineNumber: 701,
                    columnNumber: 11,
                  },
                  this,
                ),
              ],
            },
            void 0,
            true,
            {
              fileName: _jsxFileName,
              lineNumber: 699,
              columnNumber: 28,
            },
            this,
          ),
        /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
          "header",
          {
            className: "border-b border-border bg-card",
            children: /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
              "div",
              {
                className: "mx-auto max-w-5xl px-4 py-4 sm:py-5",
                children: /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                  "div",
                  {
                    className: "flex flex-col sm:flex-row sm:items-center justify-between gap-4",
                    children: [
                      /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                        "div",
                        {
                          children: [
                            /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                              "h1",
                              {
                                className: "text-2xl font-semibold tracking-tight sm:text-3xl",
                                children: "Registro de Atividade Diária",
                              },
                              void 0,
                              false,
                              {
                                fileName: _jsxFileName,
                                lineNumber: 708,
                                columnNumber: 15,
                              },
                              this,
                            ),
                            /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                              "p",
                              {
                                className: "mt-0.5 text-xs sm:text-sm text-muted-foreground",
                                children:
                                  "Tempos e movimentos do dia — registre eventos, salve na nuvem, gere relatórios em Excel e PDF.",
                              },
                              void 0,
                              false,
                              {
                                fileName: _jsxFileName,
                                lineNumber: 711,
                                columnNumber: 15,
                              },
                              this,
                            ),
                          ],
                        },
                        void 0,
                        true,
                        {
                          fileName: _jsxFileName,
                          lineNumber: 707,
                          columnNumber: 13,
                        },
                        this,
                      ),
                      /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                        "div",
                        {
                          className: "flex flex-wrap items-center gap-2 self-start sm:self-auto",
                          children: [
                            /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                              "div",
                              {
                                className: "flex bg-secondary p-1 rounded-lg border border-border",
                                children: [
                                  /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                    "button",
                                    {
                                      type: "button",
                                      onClick: () => setActiveTab("registro"),
                                      className: `px-3 py-1.5 text-xs font-semibold rounded-md transition-all flex items-center gap-1.5 ${activeTab === "registro" ? "bg-card text-foreground shadow-sm" : "text-muted-foreground hover:text-foreground"}`,
                                      children: [
                                        /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                          FileText,
                                          { className: "h-3.5 w-3.5" },
                                          void 0,
                                          false,
                                          {
                                            fileName: _jsxFileName,
                                            lineNumber: 720,
                                            columnNumber: 19,
                                          },
                                          this,
                                        ),
                                        /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                          "span",
                                          { children: "Registro de Hoje" },
                                          void 0,
                                          false,
                                          {
                                            fileName: _jsxFileName,
                                            lineNumber: 721,
                                            columnNumber: 19,
                                          },
                                          this,
                                        ),
                                      ],
                                    },
                                    void 0,
                                    true,
                                    {
                                      fileName: _jsxFileName,
                                      lineNumber: 719,
                                      columnNumber: 17,
                                    },
                                    this,
                                  ),
                                  /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                    "button",
                                    {
                                      type: "button",
                                      onClick: () => setActiveTab("banco_nuvem"),
                                      className: `px-3 py-1.5 text-xs font-semibold rounded-md transition-all flex items-center gap-1.5 ${activeTab === "banco_nuvem" ? "bg-primary text-white shadow-sm" : "text-muted-foreground hover:text-foreground"}`,
                                      children: [
                                        /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                          Cloud,
                                          { className: "h-3.5 w-3.5" },
                                          void 0,
                                          false,
                                          {
                                            fileName: _jsxFileName,
                                            lineNumber: 725,
                                            columnNumber: 19,
                                          },
                                          this,
                                        ),
                                        /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                          "span",
                                          { children: "Banco de Dados na Nuvem" },
                                          void 0,
                                          false,
                                          {
                                            fileName: _jsxFileName,
                                            lineNumber: 726,
                                            columnNumber: 19,
                                          },
                                          this,
                                        ),
                                      ],
                                    },
                                    void 0,
                                    true,
                                    {
                                      fileName: _jsxFileName,
                                      lineNumber: 724,
                                      columnNumber: 17,
                                    },
                                    this,
                                  ),
                                ],
                              },
                              void 0,
                              true,
                              {
                                fileName: _jsxFileName,
                                lineNumber: 718,
                                columnNumber: 15,
                              },
                              this,
                            ),
                            /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                              "button",
                              {
                                type: "button",
                                onClick: handleHeaderRefreshSync,
                                disabled: isSyncing,
                                className: `flex items-center gap-1.5 px-3 py-2 text-xs font-semibold rounded-lg border transition-all shadow-sm ${!isOnline ? "bg-amber-500/10 text-amber-700 border-amber-500/40 hover:bg-amber-500/20 dark:text-amber-300" : hasPendingSync ? "bg-amber-500/10 text-amber-700 border-amber-500/40 hover:bg-amber-500/20 dark:text-amber-300" : "bg-emerald-500/10 text-emerald-700 border-emerald-500/40 hover:bg-emerald-500/20 dark:text-emerald-300"}`,
                                title: !isOnline
                                  ? "Sem internet - Dados salvos no celular. Clique para tentar atualizar."
                                  : hasPendingSync
                                    ? "Internet disponível - Há dados pendentes para sincronizar com a nuvem."
                                    : "Sincronizado com a nuvem - Clique para atualizar.",
                                children: [
                                  /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                    RefreshCw,
                                    {
                                      className: `h-3.5 w-3.5 ${isSyncing ? "animate-spin text-primary" : ""}`,
                                    },
                                    void 0,
                                    false,
                                    {
                                      fileName: _jsxFileName,
                                      lineNumber: 731,
                                      columnNumber: 17,
                                    },
                                    this,
                                  ),
                                  /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                    "span",
                                    {
                                      children: isSyncing
                                        ? "Sincronizando..."
                                        : !isOnline
                                          ? "Salvo no celular (Offline)"
                                          : hasPendingSync
                                            ? "Sincronizar Pendente"
                                            : "Atualizar na Nuvem",
                                    },
                                    void 0,
                                    false,
                                    {
                                      fileName: _jsxFileName,
                                      lineNumber: 732,
                                      columnNumber: 17,
                                    },
                                    this,
                                  ),
                                  isOnline
                                    ? /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                        Wifi,
                                        {
                                          className:
                                            "h-3.5 w-3.5 text-emerald-600 dark:text-emerald-400",
                                        },
                                        void 0,
                                        false,
                                        {
                                          fileName: _jsxFileName,
                                          lineNumber: 735,
                                          columnNumber: 29,
                                        },
                                        this,
                                      )
                                    : /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                        WifiOff,
                                        {
                                          className:
                                            "h-3.5 w-3.5 text-amber-600 dark:text-amber-400",
                                        },
                                        void 0,
                                        false,
                                        {
                                          fileName: _jsxFileName,
                                          lineNumber: 735,
                                          columnNumber: 103,
                                        },
                                        this,
                                      ),
                                ],
                              },
                              void 0,
                              true,
                              {
                                fileName: _jsxFileName,
                                lineNumber: 730,
                                columnNumber: 15,
                              },
                              this,
                            ),
                            /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                              Link,
                              {
                                to: "/checklist",
                                className:
                                  "inline-flex items-center gap-1.5 px-3 py-2 text-xs font-semibold rounded-lg bg-primary/10 text-primary hover:bg-primary/20 border border-primary/30 transition-all shadow-xs",
                                title: "Abrir checklist de inspeção de máquinas e equipamentos",
                                children: [
                                  /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                    ClipboardCheck,
                                    { className: "h-3.5 w-3.5" },
                                    void 0,
                                    false,
                                    {
                                      fileName: _jsxFileName,
                                      lineNumber: 739,
                                      columnNumber: 17,
                                    },
                                    this,
                                  ),
                                  /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                    "span",
                                    { children: "Checklist do Equipamento" },
                                    void 0,
                                    false,
                                    {
                                      fileName: _jsxFileName,
                                      lineNumber: 740,
                                      columnNumber: 17,
                                    },
                                    this,
                                  ),
                                ],
                              },
                              void 0,
                              true,
                              {
                                fileName: _jsxFileName,
                                lineNumber: 738,
                                columnNumber: 15,
                              },
                              this,
                            ),
                            /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                              Link,
                              {
                                to: "/dashboard",
                                className: "btn-secondary text-xs py-2 px-3",
                                children: "Dashboard",
                              },
                              void 0,
                              false,
                              {
                                fileName: _jsxFileName,
                                lineNumber: 743,
                                columnNumber: 15,
                              },
                              this,
                            ),
                            /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                              ThemeToggle,
                              { className: "py-2 px-3" },
                              void 0,
                              false,
                              {
                                fileName: _jsxFileName,
                                lineNumber: 747,
                                columnNumber: 15,
                              },
                              this,
                            ),
                            activeTab === "registro" &&
                              /* @__PURE__ */ (void 0)(
                                "div",
                                {
                                  className:
                                    "hidden text-right sm:block ml-2 border-l border-border pl-3",
                                  children: [
                                    /* @__PURE__ */ (void 0)(
                                      "div",
                                      {
                                        className:
                                          "text-[10px] uppercase tracking-wider text-muted-foreground",
                                        children: "Jornada",
                                      },
                                      void 0,
                                      false,
                                      {
                                        fileName: _jsxFileName,
                                        lineNumber: 750,
                                        columnNumber: 19,
                                      },
                                      this,
                                    ),
                                    /* @__PURE__ */ (void 0)(
                                      "div",
                                      {
                                        className: "font-mono text-xl tabular-nums",
                                        children: formatDuration(liveJornada),
                                      },
                                      void 0,
                                      false,
                                      {
                                        fileName: _jsxFileName,
                                        lineNumber: 753,
                                        columnNumber: 19,
                                      },
                                      this,
                                    ),
                                  ],
                                },
                                void 0,
                                true,
                                {
                                  fileName: _jsxFileName,
                                  lineNumber: 749,
                                  columnNumber: 44,
                                },
                                this,
                              ),
                          ],
                        },
                        void 0,
                        true,
                        {
                          fileName: _jsxFileName,
                          lineNumber: 717,
                          columnNumber: 13,
                        },
                        this,
                      ),
                    ],
                  },
                  void 0,
                  true,
                  {
                    fileName: _jsxFileName,
                    lineNumber: 706,
                    columnNumber: 11,
                  },
                  this,
                ),
              },
              void 0,
              false,
              {
                fileName: _jsxFileName,
                lineNumber: 705,
                columnNumber: 9,
              },
              this,
            ),
          },
          void 0,
          false,
          {
            fileName: _jsxFileName,
            lineNumber: 704,
            columnNumber: 7,
          },
          this,
        ),
        /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
          "div",
          {
            className: "mx-auto max-w-5xl px-4 py-6",
            children:
              activeTab === "banco_nuvem"
                ? /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                    CloudReportsDatabase,
                    {
                      onLoadReportToForm: (r) => {
                        const {
                          header,
                          events,
                          photos: cloudPhotos,
                        } = cloudReportToHeaderAndEvents(r);
                        setState((s) => ({
                          ...s,
                          header,
                          events,
                        }));
                        setPhotos(cloudPhotos || []);
                        setCurrentReportId(r.id);
                        currentReportIdRef.current = r.id;
                        setActiveTab("registro");
                        setSaveNotification(
                          `☁️ Relatório de ${formatDate(header.data)} carregado da nuvem!`,
                        );
                        setTimeout(() => setSaveNotification(null), 4e3);
                      },
                    },
                    void 0,
                    false,
                    {
                      fileName: _jsxFileName,
                      lineNumber: 763,
                      columnNumber: 40,
                    },
                    this,
                  )
                : /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                    "div",
                    {
                      className: "grid gap-6 lg:grid-cols-3",
                      children: [
                        /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                          "section",
                          {
                            className: "space-y-6 lg:col-span-1",
                            children: [
                              /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                Card,
                                {
                                  title: "Cabeçalho",
                                  children: /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                    "div",
                                    {
                                      className: "grid grid-cols-2 gap-3",
                                      children: [
                                        /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                          Field,
                                          {
                                            label: "Data",
                                            children: /* @__PURE__ */ (0,
                                            import_jsx_dev_runtime.jsxDEV)(
                                              "input",
                                              {
                                                type: "date",
                                                value: state.header.data,
                                                onChange: (e) =>
                                                  updateHeader({ data: e.target.value }),
                                                className: "input",
                                              },
                                              void 0,
                                              false,
                                              {
                                                fileName: _jsxFileName,
                                                lineNumber: 786,
                                                columnNumber: 21,
                                              },
                                              this,
                                            ),
                                          },
                                          void 0,
                                          false,
                                          {
                                            fileName: _jsxFileName,
                                            lineNumber: 785,
                                            columnNumber: 19,
                                          },
                                          this,
                                        ),
                                        /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                          Field,
                                          {
                                            label: "Responsável",
                                            children: /* @__PURE__ */ (0,
                                            import_jsx_dev_runtime.jsxDEV)(
                                              "input",
                                              {
                                                value: state.header.nome,
                                                onChange: (e) =>
                                                  updateHeader({ nome: e.target.value }),
                                                className: "input",
                                              },
                                              void 0,
                                              false,
                                              {
                                                fileName: _jsxFileName,
                                                lineNumber: 791,
                                                columnNumber: 21,
                                              },
                                              this,
                                            ),
                                          },
                                          void 0,
                                          false,
                                          {
                                            fileName: _jsxFileName,
                                            lineNumber: 790,
                                            columnNumber: 19,
                                          },
                                          this,
                                        ),
                                        /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                          Field,
                                          {
                                            label: "Local",
                                            full: true,
                                            children: /* @__PURE__ */ (0,
                                            import_jsx_dev_runtime.jsxDEV)(
                                              LocalSelectField,
                                              {
                                                value: state.header.local,
                                                onChange: (local) => updateHeader({ local }),
                                              },
                                              void 0,
                                              false,
                                              {
                                                fileName: _jsxFileName,
                                                lineNumber: 796,
                                                columnNumber: 21,
                                              },
                                              this,
                                            ),
                                          },
                                          void 0,
                                          false,
                                          {
                                            fileName: _jsxFileName,
                                            lineNumber: 795,
                                            columnNumber: 19,
                                          },
                                          this,
                                        ),
                                        /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                          Field,
                                          {
                                            label: "Sinais",
                                            full: true,
                                            children: /* @__PURE__ */ (0,
                                            import_jsx_dev_runtime.jsxDEV)(
                                              "input",
                                              {
                                                value: state.header.sinais,
                                                onChange: (e) =>
                                                  updateHeader({
                                                    sinais: maskSinais(e.target.value),
                                                  }),
                                                placeholder: "Ex: 23BFU-24ABR",
                                                className: "input",
                                              },
                                              void 0,
                                              false,
                                              {
                                                fileName: _jsxFileName,
                                                lineNumber: 801,
                                                columnNumber: 21,
                                              },
                                              this,
                                            ),
                                          },
                                          void 0,
                                          false,
                                          {
                                            fileName: _jsxFileName,
                                            lineNumber: 800,
                                            columnNumber: 19,
                                          },
                                          this,
                                        ),
                                        /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                          Field,
                                          {
                                            label: "Linha",
                                            children: /* @__PURE__ */ (0,
                                            import_jsx_dev_runtime.jsxDEV)(
                                              "input",
                                              {
                                                value: state.header.linha,
                                                onChange: (e) =>
                                                  updateHeader({ linha: e.target.value }),
                                                className: "input",
                                              },
                                              void 0,
                                              false,
                                              {
                                                fileName: _jsxFileName,
                                                lineNumber: 806,
                                                columnNumber: 21,
                                              },
                                              this,
                                            ),
                                          },
                                          void 0,
                                          false,
                                          {
                                            fileName: _jsxFileName,
                                            lineNumber: 805,
                                            columnNumber: 19,
                                          },
                                          this,
                                        ),
                                        /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                          Field,
                                          {
                                            label: "Via",
                                            children: /* @__PURE__ */ (0,
                                            import_jsx_dev_runtime.jsxDEV)(
                                              "input",
                                              {
                                                value: state.header.via,
                                                onChange: (e) =>
                                                  updateHeader({ via: e.target.value }),
                                                className: "input",
                                              },
                                              void 0,
                                              false,
                                              {
                                                fileName: _jsxFileName,
                                                lineNumber: 811,
                                                columnNumber: 21,
                                              },
                                              this,
                                            ),
                                          },
                                          void 0,
                                          false,
                                          {
                                            fileName: _jsxFileName,
                                            lineNumber: 810,
                                            columnNumber: 19,
                                          },
                                          this,
                                        ),
                                        /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                          Field,
                                          {
                                            label: "GDA",
                                            children: /* @__PURE__ */ (0,
                                            import_jsx_dev_runtime.jsxDEV)(
                                              "input",
                                              {
                                                value: state.header.gda,
                                                onChange: (e) =>
                                                  updateHeader({ gda: maskGda(e.target.value) }),
                                                placeholder: "Ex: 2026-001235",
                                                className: "input",
                                              },
                                              void 0,
                                              false,
                                              {
                                                fileName: _jsxFileName,
                                                lineNumber: 816,
                                                columnNumber: 21,
                                              },
                                              this,
                                            ),
                                          },
                                          void 0,
                                          false,
                                          {
                                            fileName: _jsxFileName,
                                            lineNumber: 815,
                                            columnNumber: 19,
                                          },
                                          this,
                                        ),
                                        /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                          Field,
                                          {
                                            label: "OM",
                                            children: /* @__PURE__ */ (0,
                                            import_jsx_dev_runtime.jsxDEV)(
                                              "input",
                                              {
                                                value: state.header.om,
                                                onChange: (e) =>
                                                  updateHeader({ om: e.target.value }),
                                                className: "input",
                                              },
                                              void 0,
                                              false,
                                              {
                                                fileName: _jsxFileName,
                                                lineNumber: 821,
                                                columnNumber: 21,
                                              },
                                              this,
                                            ),
                                          },
                                          void 0,
                                          false,
                                          {
                                            fileName: _jsxFileName,
                                            lineNumber: 820,
                                            columnNumber: 19,
                                          },
                                          this,
                                        ),
                                        /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                          Field,
                                          {
                                            label: "Empresa",
                                            full: true,
                                            action: /* @__PURE__ */ (0,
                                            import_jsx_dev_runtime.jsxDEV)(
                                              DropdownMenu,
                                              {
                                                children: [
                                                  /* @__PURE__ */ (0,
                                                  import_jsx_dev_runtime.jsxDEV)(
                                                    DropdownMenuTrigger,
                                                    {
                                                      asChild: true,
                                                      children: /* @__PURE__ */ (0,
                                                      import_jsx_dev_runtime.jsxDEV)(
                                                        "button",
                                                        {
                                                          type: "button",
                                                          className:
                                                            "rounded p-0.5 text-muted-foreground hover:bg-secondary hover:text-foreground transition-colors",
                                                          title: "Opções do histórico",
                                                          "aria-label": "Opções do histórico",
                                                          children: /* @__PURE__ */ (0,
                                                          import_jsx_dev_runtime.jsxDEV)(
                                                            EllipsisVertical,
                                                            { className: "h-4 w-4" },
                                                            void 0,
                                                            false,
                                                            {
                                                              fileName: _jsxFileName,
                                                              lineNumber: 828,
                                                              columnNumber: 29,
                                                            },
                                                            this,
                                                          ),
                                                        },
                                                        void 0,
                                                        false,
                                                        {
                                                          fileName: _jsxFileName,
                                                          lineNumber: 827,
                                                          columnNumber: 27,
                                                        },
                                                        this,
                                                      ),
                                                    },
                                                    void 0,
                                                    false,
                                                    {
                                                      fileName: _jsxFileName,
                                                      lineNumber: 826,
                                                      columnNumber: 25,
                                                    },
                                                    this,
                                                  ),
                                                  /* @__PURE__ */ (0,
                                                  import_jsx_dev_runtime.jsxDEV)(
                                                    DropdownMenuContent,
                                                    {
                                                      align: "end",
                                                      children: /* @__PURE__ */ (0,
                                                      import_jsx_dev_runtime.jsxDEV)(
                                                        DropdownMenuItem,
                                                        {
                                                          className:
                                                            "text-destructive focus:text-destructive flex cursor-pointer items-center gap-2 text-xs",
                                                          onClick: clearCompaniesHistory,
                                                          children: [
                                                            /* @__PURE__ */ (0,
                                                            import_jsx_dev_runtime.jsxDEV)(
                                                              Trash2,
                                                              { className: "h-3.5 w-3.5" },
                                                              void 0,
                                                              false,
                                                              {
                                                                fileName: _jsxFileName,
                                                                lineNumber: 833,
                                                                columnNumber: 29,
                                                              },
                                                              this,
                                                            ),
                                                            "Apagar histórico de empresas",
                                                          ],
                                                        },
                                                        void 0,
                                                        true,
                                                        {
                                                          fileName: _jsxFileName,
                                                          lineNumber: 832,
                                                          columnNumber: 27,
                                                        },
                                                        this,
                                                      ),
                                                    },
                                                    void 0,
                                                    false,
                                                    {
                                                      fileName: _jsxFileName,
                                                      lineNumber: 831,
                                                      columnNumber: 25,
                                                    },
                                                    this,
                                                  ),
                                                ],
                                              },
                                              void 0,
                                              true,
                                              {
                                                fileName: _jsxFileName,
                                                lineNumber: 825,
                                                columnNumber: 55,
                                              },
                                              this,
                                            ),
                                            children: /* @__PURE__ */ (0,
                                            import_jsx_dev_runtime.jsxDEV)(
                                              "div",
                                              {
                                                children: [
                                                  /* @__PURE__ */ (0,
                                                  import_jsx_dev_runtime.jsxDEV)(
                                                    "input",
                                                    {
                                                      list: "companies-history-list",
                                                      value: state.header.empresa,
                                                      onChange: (e) =>
                                                        updateHeader({ empresa: e.target.value }),
                                                      onBlur: (e) => {
                                                        if (e.target.value.trim())
                                                          saveCompanyToHistory(e.target.value);
                                                      },
                                                      className: "input",
                                                    },
                                                    void 0,
                                                    false,
                                                    {
                                                      fileName: _jsxFileName,
                                                      lineNumber: 839,
                                                      columnNumber: 23,
                                                    },
                                                    this,
                                                  ),
                                                  /* @__PURE__ */ (0,
                                                  import_jsx_dev_runtime.jsxDEV)(
                                                    "datalist",
                                                    {
                                                      id: "companies-history-list",
                                                      children: companiesHistory.map((comp) =>
                                                        /* @__PURE__ */ (0,
                                                        import_jsx_dev_runtime.jsxDEV)(
                                                          "option",
                                                          { value: comp },
                                                          comp,
                                                          false,
                                                          {
                                                            fileName: _jsxFileName,
                                                            lineNumber: 847,
                                                            columnNumber: 55,
                                                          },
                                                          this,
                                                        ),
                                                      ),
                                                    },
                                                    void 0,
                                                    false,
                                                    {
                                                      fileName: _jsxFileName,
                                                      lineNumber: 846,
                                                      columnNumber: 23,
                                                    },
                                                    this,
                                                  ),
                                                ],
                                              },
                                              void 0,
                                              true,
                                              {
                                                fileName: _jsxFileName,
                                                lineNumber: 838,
                                                columnNumber: 21,
                                              },
                                              this,
                                            ),
                                          },
                                          void 0,
                                          false,
                                          {
                                            fileName: _jsxFileName,
                                            lineNumber: 825,
                                            columnNumber: 19,
                                          },
                                          this,
                                        ),
                                        /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                          Field,
                                          {
                                            label: "Nome da Atividade",
                                            full: true,
                                            action: /* @__PURE__ */ (0,
                                            import_jsx_dev_runtime.jsxDEV)(
                                              DropdownMenu,
                                              {
                                                children: [
                                                  /* @__PURE__ */ (0,
                                                  import_jsx_dev_runtime.jsxDEV)(
                                                    DropdownMenuTrigger,
                                                    {
                                                      asChild: true,
                                                      children: /* @__PURE__ */ (0,
                                                      import_jsx_dev_runtime.jsxDEV)(
                                                        "button",
                                                        {
                                                          type: "button",
                                                          className:
                                                            "rounded p-0.5 text-muted-foreground hover:bg-secondary hover:text-foreground transition-colors",
                                                          title: "Opções do histórico",
                                                          "aria-label": "Opções do histórico",
                                                          children: /* @__PURE__ */ (0,
                                                          import_jsx_dev_runtime.jsxDEV)(
                                                            EllipsisVertical,
                                                            { className: "h-4 w-4" },
                                                            void 0,
                                                            false,
                                                            {
                                                              fileName: _jsxFileName,
                                                              lineNumber: 854,
                                                              columnNumber: 29,
                                                            },
                                                            this,
                                                          ),
                                                        },
                                                        void 0,
                                                        false,
                                                        {
                                                          fileName: _jsxFileName,
                                                          lineNumber: 853,
                                                          columnNumber: 27,
                                                        },
                                                        this,
                                                      ),
                                                    },
                                                    void 0,
                                                    false,
                                                    {
                                                      fileName: _jsxFileName,
                                                      lineNumber: 852,
                                                      columnNumber: 25,
                                                    },
                                                    this,
                                                  ),
                                                  /* @__PURE__ */ (0,
                                                  import_jsx_dev_runtime.jsxDEV)(
                                                    DropdownMenuContent,
                                                    {
                                                      align: "end",
                                                      children: /* @__PURE__ */ (0,
                                                      import_jsx_dev_runtime.jsxDEV)(
                                                        DropdownMenuItem,
                                                        {
                                                          className:
                                                            "text-destructive focus:text-destructive flex cursor-pointer items-center gap-2 text-xs",
                                                          onClick: clearActivitiesHistory,
                                                          children: [
                                                            /* @__PURE__ */ (0,
                                                            import_jsx_dev_runtime.jsxDEV)(
                                                              Trash2,
                                                              { className: "h-3.5 w-3.5" },
                                                              void 0,
                                                              false,
                                                              {
                                                                fileName: _jsxFileName,
                                                                lineNumber: 859,
                                                                columnNumber: 29,
                                                              },
                                                              this,
                                                            ),
                                                            "Apagar histórico de atividades",
                                                          ],
                                                        },
                                                        void 0,
                                                        true,
                                                        {
                                                          fileName: _jsxFileName,
                                                          lineNumber: 858,
                                                          columnNumber: 27,
                                                        },
                                                        this,
                                                      ),
                                                    },
                                                    void 0,
                                                    false,
                                                    {
                                                      fileName: _jsxFileName,
                                                      lineNumber: 857,
                                                      columnNumber: 25,
                                                    },
                                                    this,
                                                  ),
                                                ],
                                              },
                                              void 0,
                                              true,
                                              {
                                                fileName: _jsxFileName,
                                                lineNumber: 851,
                                                columnNumber: 65,
                                              },
                                              this,
                                            ),
                                            children: /* @__PURE__ */ (0,
                                            import_jsx_dev_runtime.jsxDEV)(
                                              ActivitySelectField,
                                              {
                                                value: state.header.atividade,
                                                onChange: (atividade) =>
                                                  updateHeader({ atividade }),
                                                activitiesHistory,
                                                onSaveHistory: saveActivityToHistory,
                                              },
                                              void 0,
                                              false,
                                              {
                                                fileName: _jsxFileName,
                                                lineNumber: 864,
                                                columnNumber: 21,
                                              },
                                              this,
                                            ),
                                          },
                                          void 0,
                                          false,
                                          {
                                            fileName: _jsxFileName,
                                            lineNumber: 851,
                                            columnNumber: 19,
                                          },
                                          this,
                                        ),
                                        /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                          Field,
                                          {
                                            label: "Equipe / Colaboradores",
                                            full: true,
                                            action:
                                              state.roster.length > 0
                                                ? /* @__PURE__ */ (0,
                                                  import_jsx_dev_runtime.jsxDEV)(
                                                    DropdownMenu,
                                                    {
                                                      children: [
                                                        /* @__PURE__ */ (0,
                                                        import_jsx_dev_runtime.jsxDEV)(
                                                          DropdownMenuTrigger,
                                                          {
                                                            asChild: true,
                                                            children: /* @__PURE__ */ (0,
                                                            import_jsx_dev_runtime.jsxDEV)(
                                                              "button",
                                                              {
                                                                type: "button",
                                                                className:
                                                                  "rounded p-0.5 text-muted-foreground hover:bg-secondary hover:text-foreground transition-colors",
                                                                title: "Opções da equipe",
                                                                "aria-label": "Opções da equipe",
                                                                children: /* @__PURE__ */ (0,
                                                                import_jsx_dev_runtime.jsxDEV)(
                                                                  EllipsisVertical,
                                                                  { className: "h-4 w-4" },
                                                                  void 0,
                                                                  false,
                                                                  {
                                                                    fileName: _jsxFileName,
                                                                    lineNumber: 871,
                                                                    columnNumber: 31,
                                                                  },
                                                                  this,
                                                                ),
                                                              },
                                                              void 0,
                                                              false,
                                                              {
                                                                fileName: _jsxFileName,
                                                                lineNumber: 870,
                                                                columnNumber: 29,
                                                              },
                                                              this,
                                                            ),
                                                          },
                                                          void 0,
                                                          false,
                                                          {
                                                            fileName: _jsxFileName,
                                                            lineNumber: 869,
                                                            columnNumber: 27,
                                                          },
                                                          this,
                                                        ),
                                                        /* @__PURE__ */ (0,
                                                        import_jsx_dev_runtime.jsxDEV)(
                                                          DropdownMenuContent,
                                                          {
                                                            align: "end",
                                                            children: /* @__PURE__ */ (0,
                                                            import_jsx_dev_runtime.jsxDEV)(
                                                              DropdownMenuItem,
                                                              {
                                                                className:
                                                                  "text-destructive focus:text-destructive flex cursor-pointer items-center gap-2 text-xs",
                                                                onClick: clearRosterHistory,
                                                                children: [
                                                                  /* @__PURE__ */ (0,
                                                                  import_jsx_dev_runtime.jsxDEV)(
                                                                    Trash2,
                                                                    { className: "h-3.5 w-3.5" },
                                                                    void 0,
                                                                    false,
                                                                    {
                                                                      fileName: _jsxFileName,
                                                                      lineNumber: 876,
                                                                      columnNumber: 31,
                                                                    },
                                                                    this,
                                                                  ),
                                                                  "Apagar histórico de colaboradores",
                                                                ],
                                                              },
                                                              void 0,
                                                              true,
                                                              {
                                                                fileName: _jsxFileName,
                                                                lineNumber: 875,
                                                                columnNumber: 29,
                                                              },
                                                              this,
                                                            ),
                                                          },
                                                          void 0,
                                                          false,
                                                          {
                                                            fileName: _jsxFileName,
                                                            lineNumber: 874,
                                                            columnNumber: 27,
                                                          },
                                                          this,
                                                        ),
                                                      ],
                                                    },
                                                    void 0,
                                                    true,
                                                    {
                                                      fileName: _jsxFileName,
                                                      lineNumber: 868,
                                                      columnNumber: 96,
                                                    },
                                                    this,
                                                  )
                                                : null,
                                            children: /* @__PURE__ */ (0,
                                            import_jsx_dev_runtime.jsxDEV)(
                                              TeamPicker,
                                              {
                                                roster: state.roster,
                                                selected: state.header.colaboradores,
                                                onSelectedChange: (v) =>
                                                  updateHeader({ colaboradores: v }),
                                                onRosterChange: (roster) =>
                                                  setState((s) => ({
                                                    ...s,
                                                    roster,
                                                  })),
                                                newMember,
                                                setNewMember,
                                              },
                                              void 0,
                                              false,
                                              {
                                                fileName: _jsxFileName,
                                                lineNumber: 881,
                                                columnNumber: 21,
                                              },
                                              this,
                                            ),
                                          },
                                          void 0,
                                          false,
                                          {
                                            fileName: _jsxFileName,
                                            lineNumber: 868,
                                            columnNumber: 19,
                                          },
                                          this,
                                        ),
                                        /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                          Field,
                                          {
                                            label: "Descritivo da Atividade",
                                            full: true,
                                            children: /* @__PURE__ */ (0,
                                            import_jsx_dev_runtime.jsxDEV)(
                                              "textarea",
                                              {
                                                value: state.header.descritivo,
                                                onChange: (e) =>
                                                  updateHeader({ descritivo: e.target.value }),
                                                className: "input min-h-16",
                                              },
                                              void 0,
                                              false,
                                              {
                                                fileName: _jsxFileName,
                                                lineNumber: 889,
                                                columnNumber: 21,
                                              },
                                              this,
                                            ),
                                          },
                                          void 0,
                                          false,
                                          {
                                            fileName: _jsxFileName,
                                            lineNumber: 888,
                                            columnNumber: 19,
                                          },
                                          this,
                                        ),
                                        /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                          Field,
                                          {
                                            label: "Programado",
                                            full: true,
                                            children: /* @__PURE__ */ (0,
                                            import_jsx_dev_runtime.jsxDEV)(
                                              "input",
                                              {
                                                value: state.header.kmProgramado,
                                                onChange: (e) =>
                                                  updateHeader({
                                                    kmProgramado: maskProgramadoExecutado(
                                                      e.target.value,
                                                    ),
                                                  }),
                                                placeholder: "Ex: 25/21-26/21",
                                                className: "input",
                                              },
                                              void 0,
                                              false,
                                              {
                                                fileName: _jsxFileName,
                                                lineNumber: 894,
                                                columnNumber: 21,
                                              },
                                              this,
                                            ),
                                          },
                                          void 0,
                                          false,
                                          {
                                            fileName: _jsxFileName,
                                            lineNumber: 893,
                                            columnNumber: 19,
                                          },
                                          this,
                                        ),
                                        /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                          Field,
                                          {
                                            label: "Executado",
                                            full: true,
                                            action: state.header.kmProgramado
                                              ? /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                                  "button",
                                                  {
                                                    type: "button",
                                                    onClick: () =>
                                                      updateHeader({
                                                        executado: state.header.kmProgramado,
                                                      }),
                                                    className:
                                                      "text-xs font-medium text-primary hover:underline flex items-center gap-1 transition-colors",
                                                    title: "Copiar mesmo valor do Programado",
                                                    children: [
                                                      /* @__PURE__ */ (0,
                                                      import_jsx_dev_runtime.jsxDEV)(
                                                        Copy,
                                                        { className: "h-3 w-3" },
                                                        void 0,
                                                        false,
                                                        {
                                                          fileName: _jsxFileName,
                                                          lineNumber: 901,
                                                          columnNumber: 27,
                                                        },
                                                        this,
                                                      ),
                                                      /* @__PURE__ */ (0,
                                                      import_jsx_dev_runtime.jsxDEV)(
                                                        "span",
                                                        { children: "Mesmo do programado" },
                                                        void 0,
                                                        false,
                                                        {
                                                          fileName: _jsxFileName,
                                                          lineNumber: 902,
                                                          columnNumber: 27,
                                                        },
                                                        this,
                                                      ),
                                                    ],
                                                  },
                                                  void 0,
                                                  true,
                                                  {
                                                    fileName: _jsxFileName,
                                                    lineNumber: 898,
                                                    columnNumber: 85,
                                                  },
                                                  this,
                                                )
                                              : null,
                                            children: /* @__PURE__ */ (0,
                                            import_jsx_dev_runtime.jsxDEV)(
                                              "div",
                                              {
                                                className: "space-y-1.5",
                                                children: [
                                                  /* @__PURE__ */ (0,
                                                  import_jsx_dev_runtime.jsxDEV)(
                                                    "input",
                                                    {
                                                      value: state.header.executado,
                                                      onChange: (e) =>
                                                        updateHeader({
                                                          executado: maskProgramadoExecutado(
                                                            e.target.value,
                                                          ),
                                                        }),
                                                      placeholder: "Ex: 25/21-26/21",
                                                      className: "input",
                                                    },
                                                    void 0,
                                                    false,
                                                    {
                                                      fileName: _jsxFileName,
                                                      lineNumber: 905,
                                                      columnNumber: 23,
                                                    },
                                                    this,
                                                  ),
                                                  state.header.kmProgramado &&
                                                    state.header.executado !==
                                                      state.header.kmProgramado &&
                                                    /* @__PURE__ */ (void 0)(
                                                      "button",
                                                      {
                                                        type: "button",
                                                        onClick: () =>
                                                          updateHeader({
                                                            executado: state.header.kmProgramado,
                                                          }),
                                                        className:
                                                          "inline-flex items-center gap-1 text-xs text-muted-foreground hover:text-primary transition-colors",
                                                        children: [
                                                          /* @__PURE__ */ (void 0)(
                                                            Copy,
                                                            { className: "h-3 w-3" },
                                                            void 0,
                                                            false,
                                                            {
                                                              fileName: _jsxFileName,
                                                              lineNumber: 911,
                                                              columnNumber: 29,
                                                            },
                                                            this,
                                                          ),
                                                          /* @__PURE__ */ (void 0)(
                                                            "span",
                                                            {
                                                              children: [
                                                                "Usar valor do programado (",
                                                                state.header.kmProgramado,
                                                                ")",
                                                              ],
                                                            },
                                                            void 0,
                                                            true,
                                                            {
                                                              fileName: _jsxFileName,
                                                              lineNumber: 912,
                                                              columnNumber: 29,
                                                            },
                                                            this,
                                                          ),
                                                        ],
                                                      },
                                                      void 0,
                                                      true,
                                                      {
                                                        fileName: _jsxFileName,
                                                        lineNumber: 908,
                                                        columnNumber: 109,
                                                      },
                                                      this,
                                                    ),
                                                ],
                                              },
                                              void 0,
                                              true,
                                              {
                                                fileName: _jsxFileName,
                                                lineNumber: 904,
                                                columnNumber: 21,
                                              },
                                              this,
                                            ),
                                          },
                                          void 0,
                                          false,
                                          {
                                            fileName: _jsxFileName,
                                            lineNumber: 898,
                                            columnNumber: 19,
                                          },
                                          this,
                                        ),
                                        /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                          Field,
                                          {
                                            label: "Observação",
                                            full: true,
                                            children: /* @__PURE__ */ (0,
                                            import_jsx_dev_runtime.jsxDEV)(
                                              "textarea",
                                              {
                                                value: state.header.observacaoGeral,
                                                onChange: (e) =>
                                                  updateHeader({ observacaoGeral: e.target.value }),
                                                className: "input min-h-16",
                                              },
                                              void 0,
                                              false,
                                              {
                                                fileName: _jsxFileName,
                                                lineNumber: 917,
                                                columnNumber: 21,
                                              },
                                              this,
                                            ),
                                          },
                                          void 0,
                                          false,
                                          {
                                            fileName: _jsxFileName,
                                            lineNumber: 916,
                                            columnNumber: 19,
                                          },
                                          this,
                                        ),
                                      ],
                                    },
                                    void 0,
                                    true,
                                    {
                                      fileName: _jsxFileName,
                                      lineNumber: 784,
                                      columnNumber: 17,
                                    },
                                    this,
                                  ),
                                },
                                void 0,
                                false,
                                {
                                  fileName: _jsxFileName,
                                  lineNumber: 783,
                                  columnNumber: 15,
                                },
                                this,
                              ),
                              /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                Card,
                                {
                                  title: "Registrar evento",
                                  children: [
                                    /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                      "div",
                                      {
                                        className: `mb-3 flex items-center justify-between rounded-lg border px-3 py-2 text-xs font-medium ${!isOnline ? "border-amber-500/30 bg-amber-500/10 text-amber-800 dark:text-amber-200" : "border-emerald-500/30 bg-emerald-500/10 text-emerald-800 dark:text-emerald-200"}`,
                                        children: [
                                          /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                            "div",
                                            {
                                              className: "flex items-center gap-2",
                                              children: [
                                                isOnline
                                                  ? /* @__PURE__ */ (0,
                                                    import_jsx_dev_runtime.jsxDEV)(
                                                      Cloud,
                                                      {
                                                        className: `h-4 w-4 shrink-0 text-emerald-600 dark:text-emerald-400 ${isSyncing ? "animate-bounce" : ""}`,
                                                      },
                                                      void 0,
                                                      false,
                                                      {
                                                        fileName: _jsxFileName,
                                                        lineNumber: 927,
                                                        columnNumber: 33,
                                                      },
                                                      this,
                                                    )
                                                  : /* @__PURE__ */ (0,
                                                    import_jsx_dev_runtime.jsxDEV)(
                                                      WifiOff,
                                                      {
                                                        className:
                                                          "h-4 w-4 shrink-0 text-amber-600 dark:text-amber-400",
                                                      },
                                                      void 0,
                                                      false,
                                                      {
                                                        fileName: _jsxFileName,
                                                        lineNumber: 927,
                                                        columnNumber: 152,
                                                      },
                                                      this,
                                                    ),
                                                /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                                  "div",
                                                  {
                                                    children: [
                                                      /* @__PURE__ */ (0,
                                                      import_jsx_dev_runtime.jsxDEV)(
                                                        "span",
                                                        {
                                                          className: "font-semibold block",
                                                          children: isOnline
                                                            ? "Salvo na nuvem e no celular"
                                                            : "Salvo no celular (Modo Offline)",
                                                        },
                                                        void 0,
                                                        false,
                                                        {
                                                          fileName: _jsxFileName,
                                                          lineNumber: 929,
                                                          columnNumber: 23,
                                                        },
                                                        this,
                                                      ),
                                                      /* @__PURE__ */ (0,
                                                      import_jsx_dev_runtime.jsxDEV)(
                                                        "span",
                                                        {
                                                          className:
                                                            "text-[10px] text-muted-foreground block",
                                                          children: isSyncing
                                                            ? "Sincronizando agora..."
                                                            : !isOnline
                                                              ? "Sem internet — enviará para a nuvem quando reconectar"
                                                              : lastCloudSyncTime
                                                                ? `Última sincronização às ${new Date(lastCloudSyncTime).toLocaleTimeString("pt-BR")}`
                                                                : "Ativo — dados seguros no celular",
                                                        },
                                                        void 0,
                                                        false,
                                                        {
                                                          fileName: _jsxFileName,
                                                          lineNumber: 932,
                                                          columnNumber: 23,
                                                        },
                                                        this,
                                                      ),
                                                    ],
                                                  },
                                                  void 0,
                                                  true,
                                                  {
                                                    fileName: _jsxFileName,
                                                    lineNumber: 928,
                                                    columnNumber: 21,
                                                  },
                                                  this,
                                                ),
                                              ],
                                            },
                                            void 0,
                                            true,
                                            {
                                              fileName: _jsxFileName,
                                              lineNumber: 926,
                                              columnNumber: 19,
                                            },
                                            this,
                                          ),
                                          /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                            "span",
                                            {
                                              className: `inline-flex items-center rounded-full px-2 py-0.5 text-[10px] font-semibold ${!isOnline ? "bg-amber-500/20 text-amber-700 dark:text-amber-300" : "bg-emerald-500/20 text-emerald-700 dark:text-emerald-300"}`,
                                              children: !isOnline ? "Offline 📱" : "Nuvem Ativa ⚡",
                                            },
                                            void 0,
                                            false,
                                            {
                                              fileName: _jsxFileName,
                                              lineNumber: 937,
                                              columnNumber: 19,
                                            },
                                            this,
                                          ),
                                        ],
                                      },
                                      void 0,
                                      true,
                                      {
                                        fileName: _jsxFileName,
                                        lineNumber: 925,
                                        columnNumber: 17,
                                      },
                                      this,
                                    ),
                                    /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                      "div",
                                      {
                                        className: "grid grid-cols-2 gap-2",
                                        children: [
                                          /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                            EventBtn,
                                            {
                                              onClick: () => registerEvent("inicio_jornada"),
                                              disabled: jornadaStarted,
                                              variant: "primary",
                                              children: "Início da jornada",
                                            },
                                            void 0,
                                            false,
                                            {
                                              fileName: _jsxFileName,
                                              lineNumber: 943,
                                              columnNumber: 19,
                                            },
                                            this,
                                          ),
                                          /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                            EventBtn,
                                            {
                                              onClick: () => registerEvent("fim_jornada"),
                                              disabled:
                                                !jornadaStarted || jornadaEnded || openActivity,
                                              variant: "danger",
                                              children: "Fim da jornada",
                                            },
                                            void 0,
                                            false,
                                            {
                                              fileName: _jsxFileName,
                                              lineNumber: 946,
                                              columnNumber: 19,
                                            },
                                            this,
                                          ),
                                          /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                            EventBtn,
                                            {
                                              onClick: () => registerEvent("preparacao"),
                                              disabled: !jornadaStarted || jornadaEnded,
                                              children: "Preparação/Manutenção",
                                            },
                                            void 0,
                                            false,
                                            {
                                              fileName: _jsxFileName,
                                              lineNumber: 949,
                                              columnNumber: 19,
                                            },
                                            this,
                                          ),
                                          /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                            EventBtn,
                                            {
                                              onClick: () => registerEvent("parada"),
                                              disabled: !jornadaStarted || jornadaEnded,
                                              children: "Parada",
                                            },
                                            void 0,
                                            false,
                                            {
                                              fileName: _jsxFileName,
                                              lineNumber: 952,
                                              columnNumber: 19,
                                            },
                                            this,
                                          ),
                                          /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                            EventBtn,
                                            {
                                              onClick: () => registerEvent("inicio_atividade"),
                                              disabled:
                                                !jornadaStarted || jornadaEnded || openActivity,
                                              variant: "success",
                                              children: "Início da atividade",
                                            },
                                            void 0,
                                            false,
                                            {
                                              fileName: _jsxFileName,
                                              lineNumber: 955,
                                              columnNumber: 19,
                                            },
                                            this,
                                          ),
                                          /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                            EventBtn,
                                            {
                                              onClick: () => registerEvent("fim_atividade"),
                                              disabled: !openActivity,
                                              variant: "warning",
                                              children: "Fim da atividade",
                                            },
                                            void 0,
                                            false,
                                            {
                                              fileName: _jsxFileName,
                                              lineNumber: 958,
                                              columnNumber: 19,
                                            },
                                            this,
                                          ),
                                          /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                            EventBtn,
                                            {
                                              onClick: () =>
                                                registerEvent("deslocamento_ferroviario"),
                                              disabled: !jornadaStarted || jornadaEnded,
                                              children: "Deslocamento Ferroviário",
                                            },
                                            void 0,
                                            false,
                                            {
                                              fileName: _jsxFileName,
                                              lineNumber: 961,
                                              columnNumber: 19,
                                            },
                                            this,
                                          ),
                                          /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                            EventBtn,
                                            {
                                              onClick: () =>
                                                registerEvent("deslocamento_rodoviario"),
                                              disabled: !jornadaStarted || jornadaEnded,
                                              children: "Deslocamento Rodoviário",
                                            },
                                            void 0,
                                            false,
                                            {
                                              fileName: _jsxFileName,
                                              lineNumber: 964,
                                              columnNumber: 19,
                                            },
                                            this,
                                          ),
                                          /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                            EventBtn,
                                            {
                                              onClick: () => registerEvent("pronto_cco"),
                                              disabled: !jornadaStarted || jornadaEnded,
                                              children: "Pronto pra sair ao CCO",
                                            },
                                            void 0,
                                            false,
                                            {
                                              fileName: _jsxFileName,
                                              lineNumber: 967,
                                              columnNumber: 19,
                                            },
                                            this,
                                          ),
                                          /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                            EventBtn,
                                            {
                                              onClick: () => registerEvent("pronto_patio"),
                                              disabled: !jornadaStarted || jornadaEnded,
                                              children: "Pronto pra sair no pátio",
                                            },
                                            void 0,
                                            false,
                                            {
                                              fileName: _jsxFileName,
                                              lineNumber: 970,
                                              columnNumber: 19,
                                            },
                                            this,
                                          ),
                                          /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                            EventBtn,
                                            {
                                              onClick: () => registerEvent("pedido_cco"),
                                              disabled: !jornadaStarted || jornadaEnded,
                                              children: "Pedido ao CCO",
                                            },
                                            void 0,
                                            false,
                                            {
                                              fileName: _jsxFileName,
                                              lineNumber: 973,
                                              columnNumber: 19,
                                            },
                                            this,
                                          ),
                                          /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                            EventBtn,
                                            {
                                              onClick: () => registerEvent("liberacao_acesso"),
                                              disabled: !jornadaStarted || jornadaEnded,
                                              children: "Liberação de acesso a via",
                                            },
                                            void 0,
                                            false,
                                            {
                                              fileName: _jsxFileName,
                                              lineNumber: 976,
                                              columnNumber: 19,
                                            },
                                            this,
                                          ),
                                          /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                            EventBtn,
                                            {
                                              onClick: () => registerEvent("chegada_local"),
                                              disabled: !jornadaStarted || jornadaEnded,
                                              children: "Chegada ao Local",
                                            },
                                            void 0,
                                            false,
                                            {
                                              fileName: _jsxFileName,
                                              lineNumber: 979,
                                              columnNumber: 19,
                                            },
                                            this,
                                          ),
                                          /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                            EventBtn,
                                            {
                                              onClick: () => registerEvent("retorno_base"),
                                              disabled: !jornadaStarted || jornadaEnded,
                                              children: "Retorno a base",
                                            },
                                            void 0,
                                            false,
                                            {
                                              fileName: _jsxFileName,
                                              lineNumber: 982,
                                              columnNumber: 19,
                                            },
                                            this,
                                          ),
                                          /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                            EventBtn,
                                            {
                                              onClick: () =>
                                                registerEvent("dimensionamento_defeito"),
                                              disabled: !jornadaStarted || jornadaEnded,
                                              children: "Dimensionamento Defeito",
                                            },
                                            void 0,
                                            false,
                                            {
                                              fileName: _jsxFileName,
                                              lineNumber: 985,
                                              columnNumber: 19,
                                            },
                                            this,
                                          ),
                                          /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                            EventBtn,
                                            {
                                              onClick: () => registerEvent("refeicao"),
                                              disabled: !jornadaStarted || jornadaEnded,
                                              children: "Janta/Almoço",
                                            },
                                            void 0,
                                            false,
                                            {
                                              fileName: _jsxFileName,
                                              lineNumber: 988,
                                              columnNumber: 19,
                                            },
                                            this,
                                          ),
                                          /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                            EventBtn,
                                            {
                                              onClick: () => registerEvent("cancelado"),
                                              disabled: !jornadaStarted || jornadaEnded,
                                              variant: "danger",
                                              children: "Cancelado",
                                            },
                                            void 0,
                                            false,
                                            {
                                              fileName: _jsxFileName,
                                              lineNumber: 991,
                                              columnNumber: 19,
                                            },
                                            this,
                                          ),
                                          /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                            EventBtn,
                                            {
                                              onClick: () => registerEvent("outros"),
                                              disabled: !jornadaStarted || jornadaEnded,
                                              children: "Outros...",
                                            },
                                            void 0,
                                            false,
                                            {
                                              fileName: _jsxFileName,
                                              lineNumber: 994,
                                              columnNumber: 19,
                                            },
                                            this,
                                          ),
                                        ],
                                      },
                                      void 0,
                                      true,
                                      {
                                        fileName: _jsxFileName,
                                        lineNumber: 942,
                                        columnNumber: 17,
                                      },
                                      this,
                                    ),
                                    jornadaEnded &&
                                      /* @__PURE__ */ (void 0)(
                                        "p",
                                        {
                                          className:
                                            "mt-3 rounded-md bg-accent/60 px-3 py-2 text-xs text-accent-foreground",
                                          children: "Jornada encerrada. Gere o relatório abaixo.",
                                        },
                                        void 0,
                                        false,
                                        {
                                          fileName: _jsxFileName,
                                          lineNumber: 998,
                                          columnNumber: 34,
                                        },
                                        this,
                                      ),
                                  ],
                                },
                                void 0,
                                true,
                                {
                                  fileName: _jsxFileName,
                                  lineNumber: 924,
                                  columnNumber: 15,
                                },
                                this,
                              ),
                              /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                PhotoUploader,
                                {
                                  photos,
                                  onChange: setPhotos,
                                },
                                void 0,
                                false,
                                {
                                  fileName: _jsxFileName,
                                  lineNumber: 1003,
                                  columnNumber: 15,
                                },
                                this,
                              ),
                              /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                Card,
                                {
                                  title: "Relatório",
                                  children: /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                    "div",
                                    {
                                      className: "flex flex-col gap-2",
                                      children: [
                                        /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                          "button",
                                          {
                                            type: "button",
                                            onClick: handleSaveToCloud,
                                            disabled: state.events.length === 0,
                                            className:
                                              "w-full bg-emerald-600 hover:bg-emerald-700 text-white font-medium py-2.5 px-3 rounded-md transition-colors flex items-center justify-center gap-2 shadow-sm disabled:opacity-50 text-xs sm:text-sm",
                                            title:
                                              "Salvar este relatório no banco de dados na nuvem para acesso de toda a equipe",
                                            children: [
                                              /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                                Cloud,
                                                { className: "h-4 w-4" },
                                                void 0,
                                                false,
                                                {
                                                  fileName: _jsxFileName,
                                                  lineNumber: 1008,
                                                  columnNumber: 21,
                                                },
                                                this,
                                              ),
                                              /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                                "span",
                                                { children: "Salvar Relatório na Nuvem" },
                                                void 0,
                                                false,
                                                {
                                                  fileName: _jsxFileName,
                                                  lineNumber: 1009,
                                                  columnNumber: 21,
                                                },
                                                this,
                                              ),
                                            ],
                                          },
                                          void 0,
                                          true,
                                          {
                                            fileName: _jsxFileName,
                                            lineNumber: 1007,
                                            columnNumber: 19,
                                          },
                                          this,
                                        ),
                                        /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                          "button",
                                          {
                                            onClick: sendWhatsApp,
                                            className: "btn-primary text-xs sm:text-sm",
                                            disabled: state.events.length === 0,
                                            children: "Enviar resumo por WhatsApp",
                                          },
                                          void 0,
                                          false,
                                          {
                                            fileName: _jsxFileName,
                                            lineNumber: 1012,
                                            columnNumber: 19,
                                          },
                                          this,
                                        ),
                                        /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                          "div",
                                          {
                                            className:
                                              "text-[11px] font-semibold uppercase tracking-wider text-muted-foreground mt-2",
                                            children: "Apropriação SAP & Excel",
                                          },
                                          void 0,
                                          false,
                                          {
                                            fileName: _jsxFileName,
                                            lineNumber: 1016,
                                            columnNumber: 19,
                                          },
                                          this,
                                        ),
                                        /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                          "button",
                                          {
                                            type: "button",
                                            onClick: () => setIsSapModalOpen(true),
                                            className:
                                              "w-full bg-emerald-600 hover:bg-emerald-700 text-white font-semibold py-2.5 px-3 rounded-md transition-colors flex items-center justify-between gap-2 shadow-xs text-xs sm:text-sm border border-emerald-500",
                                            title:
                                              "Gerar arquivo em Excel (.xlsm) com Macro integrada para apropriação no SAP via transação IW41",
                                            children: [
                                              /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                                "div",
                                                {
                                                  className: "flex items-center gap-2",
                                                  children: [
                                                    /* @__PURE__ */ (0,
                                                    import_jsx_dev_runtime.jsxDEV)(
                                                      FileSpreadsheet,
                                                      { className: "h-4 w-4 shrink-0" },
                                                      void 0,
                                                      false,
                                                      {
                                                        fileName: _jsxFileName,
                                                        lineNumber: 1023,
                                                        columnNumber: 23,
                                                      },
                                                      this,
                                                    ),
                                                    /* @__PURE__ */ (0,
                                                    import_jsx_dev_runtime.jsxDEV)(
                                                      "span",
                                                      { children: "Apropriação SAP (IW41)" },
                                                      void 0,
                                                      false,
                                                      {
                                                        fileName: _jsxFileName,
                                                        lineNumber: 1024,
                                                        columnNumber: 23,
                                                      },
                                                      this,
                                                    ),
                                                  ],
                                                },
                                                void 0,
                                                true,
                                                {
                                                  fileName: _jsxFileName,
                                                  lineNumber: 1022,
                                                  columnNumber: 21,
                                                },
                                                this,
                                              ),
                                              /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                                "span",
                                                {
                                                  className:
                                                    "bg-black/25 text-[10px] font-mono px-1.5 py-0.5 rounded text-white/90 shrink-0",
                                                  children: ".xlsm + Macro",
                                                },
                                                void 0,
                                                false,
                                                {
                                                  fileName: _jsxFileName,
                                                  lineNumber: 1026,
                                                  columnNumber: 21,
                                                },
                                                this,
                                              ),
                                            ],
                                          },
                                          void 0,
                                          true,
                                          {
                                            fileName: _jsxFileName,
                                            lineNumber: 1021,
                                            columnNumber: 19,
                                          },
                                          this,
                                        ),
                                        /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                          "div",
                                          {
                                            className: "grid grid-cols-2 gap-2",
                                            children: [
                                              /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                                "button",
                                                {
                                                  type: "button",
                                                  onClick: handleDownloadExcel,
                                                  className:
                                                    "btn-secondary text-xs text-emerald-700 dark:text-emerald-400 font-semibold border-emerald-500/30 flex items-center justify-center gap-1",
                                                  disabled: state.events.length === 0,
                                                  title:
                                                    "Baixar relatório diário em formato de planilha Excel (.xlsx)",
                                                  children: [
                                                    /* @__PURE__ */ (0,
                                                    import_jsx_dev_runtime.jsxDEV)(
                                                      FileSpreadsheet,
                                                      { className: "h-3.5 w-3.5" },
                                                      void 0,
                                                      false,
                                                      {
                                                        fileName: _jsxFileName,
                                                        lineNumber: 1033,
                                                        columnNumber: 23,
                                                      },
                                                      this,
                                                    ),
                                                    /* @__PURE__ */ (0,
                                                    import_jsx_dev_runtime.jsxDEV)(
                                                      "span",
                                                      { children: "Baixar Excel" },
                                                      void 0,
                                                      false,
                                                      {
                                                        fileName: _jsxFileName,
                                                        lineNumber: 1034,
                                                        columnNumber: 23,
                                                      },
                                                      this,
                                                    ),
                                                  ],
                                                },
                                                void 0,
                                                true,
                                                {
                                                  fileName: _jsxFileName,
                                                  lineNumber: 1032,
                                                  columnNumber: 21,
                                                },
                                                this,
                                              ),
                                              /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                                "button",
                                                {
                                                  type: "button",
                                                  onClick: shareExcelWhatsApp,
                                                  className:
                                                    "btn-secondary text-xs text-emerald-700 dark:text-emerald-400 font-semibold border-emerald-500/30 flex items-center justify-center gap-1.5",
                                                  disabled: state.events.length === 0,
                                                  title:
                                                    "Compartilhar ou anexar relatório em planilha Excel via WhatsApp",
                                                  children: [
                                                    /* @__PURE__ */ (0,
                                                    import_jsx_dev_runtime.jsxDEV)(
                                                      FileSpreadsheet,
                                                      { className: "h-3.5 w-3.5" },
                                                      void 0,
                                                      false,
                                                      {
                                                        fileName: _jsxFileName,
                                                        lineNumber: 1037,
                                                        columnNumber: 23,
                                                      },
                                                      this,
                                                    ),
                                                    /* @__PURE__ */ (0,
                                                    import_jsx_dev_runtime.jsxDEV)(
                                                      "span",
                                                      { children: "Excel WhatsApp" },
                                                      void 0,
                                                      false,
                                                      {
                                                        fileName: _jsxFileName,
                                                        lineNumber: 1038,
                                                        columnNumber: 23,
                                                      },
                                                      this,
                                                    ),
                                                  ],
                                                },
                                                void 0,
                                                true,
                                                {
                                                  fileName: _jsxFileName,
                                                  lineNumber: 1036,
                                                  columnNumber: 21,
                                                },
                                                this,
                                              ),
                                            ],
                                          },
                                          void 0,
                                          true,
                                          {
                                            fileName: _jsxFileName,
                                            lineNumber: 1031,
                                            columnNumber: 19,
                                          },
                                          this,
                                        ),
                                        /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                          "div",
                                          {
                                            className:
                                              "text-[11px] font-semibold uppercase tracking-wider text-muted-foreground mt-2",
                                            children: "PDF",
                                          },
                                          void 0,
                                          false,
                                          {
                                            fileName: _jsxFileName,
                                            lineNumber: 1042,
                                            columnNumber: 19,
                                          },
                                          this,
                                        ),
                                        /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                          "div",
                                          {
                                            className: "grid grid-cols-2 gap-2",
                                            children: [
                                              /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                                "button",
                                                {
                                                  onClick: downloadPDF,
                                                  className: "btn-secondary text-xs",
                                                  disabled: state.events.length === 0,
                                                  children: "Baixar PDF",
                                                },
                                                void 0,
                                                false,
                                                {
                                                  fileName: _jsxFileName,
                                                  lineNumber: 1046,
                                                  columnNumber: 21,
                                                },
                                                this,
                                              ),
                                              /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                                "button",
                                                {
                                                  onClick: sharePdfWhatsApp,
                                                  className: "btn-secondary text-xs",
                                                  disabled: state.events.length === 0,
                                                  children: "PDF via WhatsApp",
                                                },
                                                void 0,
                                                false,
                                                {
                                                  fileName: _jsxFileName,
                                                  lineNumber: 1049,
                                                  columnNumber: 21,
                                                },
                                                this,
                                              ),
                                            ],
                                          },
                                          void 0,
                                          true,
                                          {
                                            fileName: _jsxFileName,
                                            lineNumber: 1045,
                                            columnNumber: 19,
                                          },
                                          this,
                                        ),
                                        /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                          "button",
                                          {
                                            onClick: resetDay,
                                            className: "btn-ghost mt-2 text-destructive text-xs",
                                            children: "Zerar dia",
                                          },
                                          void 0,
                                          false,
                                          {
                                            fileName: _jsxFileName,
                                            lineNumber: 1054,
                                            columnNumber: 19,
                                          },
                                          this,
                                        ),
                                      ],
                                    },
                                    void 0,
                                    true,
                                    {
                                      fileName: _jsxFileName,
                                      lineNumber: 1006,
                                      columnNumber: 17,
                                    },
                                    this,
                                  ),
                                },
                                void 0,
                                false,
                                {
                                  fileName: _jsxFileName,
                                  lineNumber: 1005,
                                  columnNumber: 15,
                                },
                                this,
                              ),
                            ],
                          },
                          void 0,
                          true,
                          {
                            fileName: _jsxFileName,
                            lineNumber: 782,
                            columnNumber: 13,
                          },
                          this,
                        ),
                        /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                          "section",
                          {
                            className: "space-y-6 lg:col-span-2",
                            children: [
                              /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                LastDefectCard,
                                {
                                  value: sharedLastDefect,
                                  onSave: (val) => {
                                    const v = val.trim();
                                    setSharedLastDefect(v);
                                    saveLastDefect(v).catch(() => {});
                                  },
                                },
                                void 0,
                                false,
                                {
                                  fileName: _jsxFileName,
                                  lineNumber: 1063,
                                  columnNumber: 15,
                                },
                                this,
                              ),
                              /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                Card,
                                {
                                  title: "Resumo",
                                  children: /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                    "div",
                                    {
                                      className: "grid grid-cols-2 gap-3 sm:grid-cols-3",
                                      children: [
                                        /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                          Stat,
                                          {
                                            label: "Jornada",
                                            value: formatDuration(liveJornada),
                                            tone: "primary",
                                          },
                                          void 0,
                                          false,
                                          {
                                            fileName: _jsxFileName,
                                            lineNumber: 1071,
                                            columnNumber: 19,
                                          },
                                          this,
                                        ),
                                        /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                          Stat,
                                          {
                                            label: "Atividade",
                                            value: formatDuration(t.atividade),
                                            tone: "success",
                                          },
                                          void 0,
                                          false,
                                          {
                                            fileName: _jsxFileName,
                                            lineNumber: 1072,
                                            columnNumber: 19,
                                          },
                                          this,
                                        ),
                                        /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                          Stat,
                                          {
                                            label: "Preparação",
                                            value: formatDuration(t.preparacao),
                                          },
                                          void 0,
                                          false,
                                          {
                                            fileName: _jsxFileName,
                                            lineNumber: 1073,
                                            columnNumber: 19,
                                          },
                                          this,
                                        ),
                                        /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                          Stat,
                                          {
                                            label: "Paradas",
                                            value: formatDuration(t.parada),
                                            tone: "warning",
                                          },
                                          void 0,
                                          false,
                                          {
                                            fileName: _jsxFileName,
                                            lineNumber: 1074,
                                            columnNumber: 19,
                                          },
                                          this,
                                        ),
                                        /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                          Stat,
                                          {
                                            label: "Produção (m)",
                                            value: String(t.producaoTotal || 0),
                                          },
                                          void 0,
                                          false,
                                          {
                                            fileName: _jsxFileName,
                                            lineNumber: 1075,
                                            columnNumber: 19,
                                          },
                                          this,
                                        ),
                                        /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                          Stat,
                                          {
                                            label: "Defeitos",
                                            value: String(t.defeitosTotal || 0),
                                          },
                                          void 0,
                                          false,
                                          {
                                            fileName: _jsxFileName,
                                            lineNumber: 1076,
                                            columnNumber: 19,
                                          },
                                          this,
                                        ),
                                      ],
                                    },
                                    void 0,
                                    true,
                                    {
                                      fileName: _jsxFileName,
                                      lineNumber: 1070,
                                      columnNumber: 17,
                                    },
                                    this,
                                  ),
                                },
                                void 0,
                                false,
                                {
                                  fileName: _jsxFileName,
                                  lineNumber: 1069,
                                  columnNumber: 15,
                                },
                                this,
                              ),
                              /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                Card,
                                {
                                  title: "Linha do tempo",
                                  children:
                                    sortedEvents.length === 0
                                      ? /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                          "p",
                                          {
                                            className: "text-sm text-muted-foreground",
                                            children:
                                              'Nenhum evento registrado ainda. Comece pelo "Início da jornada".',
                                          },
                                          void 0,
                                          false,
                                          {
                                            fileName: _jsxFileName,
                                            lineNumber: 1081,
                                            columnNumber: 46,
                                          },
                                          this,
                                        )
                                      : /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                          "ol",
                                          {
                                            className: "space-y-3",
                                            children: sortedEvents.map((e, i) => {
                                              const startTs = e.timestamp;
                                              const next = sortedEvents[i + 1];
                                              const hasSubsequent = Boolean(next);
                                              let endStr = "";
                                              let durStr = "";
                                              let isOngoing = false;
                                              if (hasSubsequent) {
                                                const endTs = next.timestamp;
                                                const durMs = Math.max(0, endTs - startTs);
                                                endStr = formatTime(endTs);
                                                durStr = formatDuration(durMs);
                                              } else if (e.kind === "fim_jornada") {
                                                endStr = formatTime(startTs);
                                                durStr = "00:00:00";
                                              } else {
                                                isOngoing = true;
                                                const durMs = Math.max(0, now - startTs);
                                                endStr = "Em andamento";
                                                durStr = `${formatDuration(durMs)} (em andamento)`;
                                              }
                                              return /* @__PURE__ */ (0,
                                              import_jsx_dev_runtime.jsxDEV)(
                                                "li",
                                                {
                                                  className:
                                                    "flex flex-col gap-2 rounded-lg border border-border bg-card p-3 shadow-xs",
                                                  children: [
                                                    /* @__PURE__ */ (0,
                                                    import_jsx_dev_runtime.jsxDEV)(
                                                      "div",
                                                      {
                                                        className:
                                                          "flex flex-wrap items-center justify-between gap-2 border-b border-border/50 pb-2",
                                                        children: [
                                                          /* @__PURE__ */ (0,
                                                          import_jsx_dev_runtime.jsxDEV)(
                                                            "div",
                                                            {
                                                              className:
                                                                "flex flex-wrap items-center gap-2",
                                                              children: [
                                                                /* @__PURE__ */ (0,
                                                                import_jsx_dev_runtime.jsxDEV)(
                                                                  "span",
                                                                  {
                                                                    className: `badge badge-${e.kind}`,
                                                                    children: EVENT_LABEL[e.kind],
                                                                  },
                                                                  void 0,
                                                                  false,
                                                                  {
                                                                    fileName: _jsxFileName,
                                                                    lineNumber: 1108,
                                                                    columnNumber: 31,
                                                                  },
                                                                  this,
                                                                ),
                                                                isOngoing
                                                                  ? /* @__PURE__ */ (0,
                                                                    import_jsx_dev_runtime.jsxDEV)(
                                                                      "span",
                                                                      {
                                                                        className:
                                                                          "inline-flex items-center gap-1.5 rounded-full bg-amber-500/10 px-2.5 py-0.5 text-xs font-semibold text-amber-600 dark:text-amber-400 border border-amber-500/20 animate-pulse",
                                                                        children: [
                                                                          /* @__PURE__ */ (0,
                                                                          import_jsx_dev_runtime.jsxDEV)(
                                                                            "span",
                                                                            {
                                                                              className:
                                                                                "h-2 w-2 rounded-full bg-amber-500",
                                                                            },
                                                                            void 0,
                                                                            false,
                                                                            {
                                                                              fileName:
                                                                                _jsxFileName,
                                                                              lineNumber: 1110,
                                                                              columnNumber: 35,
                                                                            },
                                                                            this,
                                                                          ),
                                                                          /* @__PURE__ */ (0,
                                                                          import_jsx_dev_runtime.jsxDEV)(
                                                                            "span",
                                                                            {
                                                                              children: [
                                                                                "Duração: ",
                                                                                durStr,
                                                                              ],
                                                                            },
                                                                            void 0,
                                                                            true,
                                                                            {
                                                                              fileName:
                                                                                _jsxFileName,
                                                                              lineNumber: 1111,
                                                                              columnNumber: 35,
                                                                            },
                                                                            this,
                                                                          ),
                                                                        ],
                                                                      },
                                                                      void 0,
                                                                      true,
                                                                      {
                                                                        fileName: _jsxFileName,
                                                                        lineNumber: 1109,
                                                                        columnNumber: 44,
                                                                      },
                                                                      this,
                                                                    )
                                                                  : /* @__PURE__ */ (0,
                                                                    import_jsx_dev_runtime.jsxDEV)(
                                                                      "span",
                                                                      {
                                                                        className:
                                                                          "text-xs font-mono font-medium text-muted-foreground bg-secondary/80 px-2.5 py-0.5 rounded-md border border-border/40",
                                                                        children: [
                                                                          "Duração:",
                                                                          " ",
                                                                          /* @__PURE__ */ (0,
                                                                          import_jsx_dev_runtime.jsxDEV)(
                                                                            "strong",
                                                                            {
                                                                              className:
                                                                                "text-foreground font-bold",
                                                                              children: durStr,
                                                                            },
                                                                            void 0,
                                                                            false,
                                                                            {
                                                                              fileName:
                                                                                _jsxFileName,
                                                                              lineNumber: 1114,
                                                                              columnNumber: 35,
                                                                            },
                                                                            this,
                                                                          ),
                                                                        ],
                                                                      },
                                                                      void 0,
                                                                      true,
                                                                      {
                                                                        fileName: _jsxFileName,
                                                                        lineNumber: 1112,
                                                                        columnNumber: 43,
                                                                      },
                                                                      this,
                                                                    ),
                                                              ],
                                                            },
                                                            void 0,
                                                            true,
                                                            {
                                                              fileName: _jsxFileName,
                                                              lineNumber: 1107,
                                                              columnNumber: 29,
                                                            },
                                                            this,
                                                          ),
                                                          /* @__PURE__ */ (0,
                                                          import_jsx_dev_runtime.jsxDEV)(
                                                            "div",
                                                            {
                                                              className:
                                                                "flex items-center gap-2 text-xs",
                                                              children: [
                                                                /* @__PURE__ */ (0,
                                                                import_jsx_dev_runtime.jsxDEV)(
                                                                  "button",
                                                                  {
                                                                    type: "button",
                                                                    onClick: () => {
                                                                      setEditingDescEventId(e.id);
                                                                      setEditingDescValue(
                                                                        e.descricao || "",
                                                                      );
                                                                    },
                                                                    className:
                                                                      "inline-flex items-center gap-1 text-muted-foreground hover:text-primary transition-colors",
                                                                    title:
                                                                      "Editar descrição do evento",
                                                                    children: [
                                                                      /* @__PURE__ */ (0,
                                                                      import_jsx_dev_runtime.jsxDEV)(
                                                                        Pencil,
                                                                        { className: "h-3 w-3" },
                                                                        void 0,
                                                                        false,
                                                                        {
                                                                          fileName: _jsxFileName,
                                                                          lineNumber: 1123,
                                                                          columnNumber: 33,
                                                                        },
                                                                        this,
                                                                      ),
                                                                      /* @__PURE__ */ (0,
                                                                      import_jsx_dev_runtime.jsxDEV)(
                                                                        "span",
                                                                        {
                                                                          children: e.descricao
                                                                            ? "Editar desc."
                                                                            : "+ Descrição",
                                                                        },
                                                                        void 0,
                                                                        false,
                                                                        {
                                                                          fileName: _jsxFileName,
                                                                          lineNumber: 1124,
                                                                          columnNumber: 33,
                                                                        },
                                                                        this,
                                                                      ),
                                                                    ],
                                                                  },
                                                                  void 0,
                                                                  true,
                                                                  {
                                                                    fileName: _jsxFileName,
                                                                    lineNumber: 1119,
                                                                    columnNumber: 31,
                                                                  },
                                                                  this,
                                                                ),
                                                                /* @__PURE__ */ (0,
                                                                import_jsx_dev_runtime.jsxDEV)(
                                                                  "span",
                                                                  {
                                                                    className:
                                                                      "text-muted-foreground",
                                                                    children: "•",
                                                                  },
                                                                  void 0,
                                                                  false,
                                                                  {
                                                                    fileName: _jsxFileName,
                                                                    lineNumber: 1126,
                                                                    columnNumber: 31,
                                                                  },
                                                                  this,
                                                                ),
                                                                /* @__PURE__ */ (0,
                                                                import_jsx_dev_runtime.jsxDEV)(
                                                                  "button",
                                                                  {
                                                                    type: "button",
                                                                    onClick: () =>
                                                                      removeEvent(e.id),
                                                                    className:
                                                                      "text-muted-foreground hover:text-destructive transition-colors",
                                                                    children: "remover",
                                                                  },
                                                                  void 0,
                                                                  false,
                                                                  {
                                                                    fileName: _jsxFileName,
                                                                    lineNumber: 1127,
                                                                    columnNumber: 31,
                                                                  },
                                                                  this,
                                                                ),
                                                              ],
                                                            },
                                                            void 0,
                                                            true,
                                                            {
                                                              fileName: _jsxFileName,
                                                              lineNumber: 1118,
                                                              columnNumber: 29,
                                                            },
                                                            this,
                                                          ),
                                                        ],
                                                      },
                                                      void 0,
                                                      true,
                                                      {
                                                        fileName: _jsxFileName,
                                                        lineNumber: 1106,
                                                        columnNumber: 27,
                                                      },
                                                      this,
                                                    ),
                                                    /* @__PURE__ */ (0,
                                                    import_jsx_dev_runtime.jsxDEV)(
                                                      "div",
                                                      {
                                                        className:
                                                          "grid grid-cols-1 md:grid-cols-2 gap-3 text-xs pt-1",
                                                        children: [
                                                          /* @__PURE__ */ (0,
                                                          import_jsx_dev_runtime.jsxDEV)(
                                                            "div",
                                                            {
                                                              className:
                                                                "flex flex-wrap items-center gap-x-3 gap-y-1.5 bg-secondary/30 p-2 rounded-md border border-border/40",
                                                              children: [
                                                                /* @__PURE__ */ (0,
                                                                import_jsx_dev_runtime.jsxDEV)(
                                                                  "div",
                                                                  {
                                                                    className:
                                                                      "flex items-center gap-1.5",
                                                                    children: [
                                                                      /* @__PURE__ */ (0,
                                                                      import_jsx_dev_runtime.jsxDEV)(
                                                                        "span",
                                                                        {
                                                                          className:
                                                                            "text-[10px] font-semibold text-muted-foreground uppercase tracking-wider",
                                                                          children: "Início:",
                                                                        },
                                                                        void 0,
                                                                        false,
                                                                        {
                                                                          fileName: _jsxFileName,
                                                                          lineNumber: 1137,
                                                                          columnNumber: 33,
                                                                        },
                                                                        this,
                                                                      ),
                                                                      /* @__PURE__ */ (0,
                                                                      import_jsx_dev_runtime.jsxDEV)(
                                                                        "input",
                                                                        {
                                                                          type: "datetime-local",
                                                                          step: 1,
                                                                          value: tsToLocalInput(
                                                                            e.timestamp,
                                                                          ),
                                                                          onChange: (ev) => {
                                                                            const ts =
                                                                              parseLocalInput(
                                                                                ev.target.value,
                                                                              );
                                                                            if (ts)
                                                                              updateEvent(e.id, {
                                                                                timestamp: ts,
                                                                              });
                                                                          },
                                                                          className:
                                                                            "input px-1.5 py-0.5 text-xs font-mono w-auto",
                                                                        },
                                                                        void 0,
                                                                        false,
                                                                        {
                                                                          fileName: _jsxFileName,
                                                                          lineNumber: 1140,
                                                                          columnNumber: 33,
                                                                        },
                                                                        this,
                                                                      ),
                                                                    ],
                                                                  },
                                                                  void 0,
                                                                  true,
                                                                  {
                                                                    fileName: _jsxFileName,
                                                                    lineNumber: 1136,
                                                                    columnNumber: 31,
                                                                  },
                                                                  this,
                                                                ),
                                                                /* @__PURE__ */ (0,
                                                                import_jsx_dev_runtime.jsxDEV)(
                                                                  "div",
                                                                  {
                                                                    className:
                                                                      "flex items-center gap-1.5",
                                                                    children: [
                                                                      /* @__PURE__ */ (0,
                                                                      import_jsx_dev_runtime.jsxDEV)(
                                                                        "span",
                                                                        {
                                                                          className:
                                                                            "text-[10px] font-semibold text-muted-foreground uppercase tracking-wider",
                                                                          children: "Fim:",
                                                                        },
                                                                        void 0,
                                                                        false,
                                                                        {
                                                                          fileName: _jsxFileName,
                                                                          lineNumber: 1148,
                                                                          columnNumber: 33,
                                                                        },
                                                                        this,
                                                                      ),
                                                                      /* @__PURE__ */ (0,
                                                                      import_jsx_dev_runtime.jsxDEV)(
                                                                        "span",
                                                                        {
                                                                          className:
                                                                            "font-mono font-semibold text-foreground bg-background px-2 py-0.5 rounded border border-border/60",
                                                                          children: endStr,
                                                                        },
                                                                        void 0,
                                                                        false,
                                                                        {
                                                                          fileName: _jsxFileName,
                                                                          lineNumber: 1151,
                                                                          columnNumber: 33,
                                                                        },
                                                                        this,
                                                                      ),
                                                                    ],
                                                                  },
                                                                  void 0,
                                                                  true,
                                                                  {
                                                                    fileName: _jsxFileName,
                                                                    lineNumber: 1147,
                                                                    columnNumber: 31,
                                                                  },
                                                                  this,
                                                                ),
                                                              ],
                                                            },
                                                            void 0,
                                                            true,
                                                            {
                                                              fileName: _jsxFileName,
                                                              lineNumber: 1135,
                                                              columnNumber: 29,
                                                            },
                                                            this,
                                                          ),
                                                          /* @__PURE__ */ (0,
                                                          import_jsx_dev_runtime.jsxDEV)(
                                                            "div",
                                                            {
                                                              className:
                                                                "flex flex-wrap items-center gap-x-3 gap-y-1 text-xs text-muted-foreground self-center",
                                                              children: [
                                                                e.local &&
                                                                  /* @__PURE__ */ (void 0)(
                                                                    "span",
                                                                    {
                                                                      children: [
                                                                        "local: ",
                                                                        /* @__PURE__ */ (void 0)(
                                                                          "b",
                                                                          {
                                                                            className:
                                                                              "text-foreground",
                                                                            children: e.local,
                                                                          },
                                                                          void 0,
                                                                          false,
                                                                          {
                                                                            fileName: _jsxFileName,
                                                                            lineNumber: 1160,
                                                                            columnNumber: 42,
                                                                          },
                                                                          this,
                                                                        ),
                                                                      ],
                                                                    },
                                                                    void 0,
                                                                    true,
                                                                    {
                                                                      fileName: _jsxFileName,
                                                                      lineNumber: 1159,
                                                                      columnNumber: 43,
                                                                    },
                                                                    this,
                                                                  ),
                                                                e.kmInicial &&
                                                                  /* @__PURE__ */ (void 0)(
                                                                    "span",
                                                                    {
                                                                      children: [
                                                                        "km inicial: ",
                                                                        /* @__PURE__ */ (void 0)(
                                                                          "b",
                                                                          {
                                                                            className:
                                                                              "text-foreground",
                                                                            children: e.kmInicial,
                                                                          },
                                                                          void 0,
                                                                          false,
                                                                          {
                                                                            fileName: _jsxFileName,
                                                                            lineNumber: 1163,
                                                                            columnNumber: 47,
                                                                          },
                                                                          this,
                                                                        ),
                                                                      ],
                                                                    },
                                                                    void 0,
                                                                    true,
                                                                    {
                                                                      fileName: _jsxFileName,
                                                                      lineNumber: 1162,
                                                                      columnNumber: 47,
                                                                    },
                                                                    this,
                                                                  ),
                                                                e.kmFinal &&
                                                                  /* @__PURE__ */ (void 0)(
                                                                    "span",
                                                                    {
                                                                      children: [
                                                                        "km final: ",
                                                                        /* @__PURE__ */ (void 0)(
                                                                          "b",
                                                                          {
                                                                            className:
                                                                              "text-foreground",
                                                                            children: e.kmFinal,
                                                                          },
                                                                          void 0,
                                                                          false,
                                                                          {
                                                                            fileName: _jsxFileName,
                                                                            lineNumber: 1166,
                                                                            columnNumber: 45,
                                                                          },
                                                                          this,
                                                                        ),
                                                                      ],
                                                                    },
                                                                    void 0,
                                                                    true,
                                                                    {
                                                                      fileName: _jsxFileName,
                                                                      lineNumber: 1165,
                                                                      columnNumber: 45,
                                                                    },
                                                                    this,
                                                                  ),
                                                                e.producao &&
                                                                  /* @__PURE__ */ (void 0)(
                                                                    "span",
                                                                    {
                                                                      children: [
                                                                        "produção: ",
                                                                        /* @__PURE__ */ (void 0)(
                                                                          "b",
                                                                          {
                                                                            className:
                                                                              "text-foreground",
                                                                            children: [
                                                                              e.producao,
                                                                              " m",
                                                                            ],
                                                                          },
                                                                          void 0,
                                                                          true,
                                                                          {
                                                                            fileName: _jsxFileName,
                                                                            lineNumber: 1169,
                                                                            columnNumber: 45,
                                                                          },
                                                                          this,
                                                                        ),
                                                                      ],
                                                                    },
                                                                    void 0,
                                                                    true,
                                                                    {
                                                                      fileName: _jsxFileName,
                                                                      lineNumber: 1168,
                                                                      columnNumber: 46,
                                                                    },
                                                                    this,
                                                                  ),
                                                                e.defeitos &&
                                                                  /* @__PURE__ */ (void 0)(
                                                                    "span",
                                                                    {
                                                                      children: [
                                                                        "defeitos: ",
                                                                        /* @__PURE__ */ (void 0)(
                                                                          "b",
                                                                          {
                                                                            className:
                                                                              "text-foreground",
                                                                            children: e.defeitos,
                                                                          },
                                                                          void 0,
                                                                          false,
                                                                          {
                                                                            fileName: _jsxFileName,
                                                                            lineNumber: 1172,
                                                                            columnNumber: 45,
                                                                          },
                                                                          this,
                                                                        ),
                                                                      ],
                                                                    },
                                                                    void 0,
                                                                    true,
                                                                    {
                                                                      fileName: _jsxFileName,
                                                                      lineNumber: 1171,
                                                                      columnNumber: 46,
                                                                    },
                                                                    this,
                                                                  ),
                                                                e.ultimoDefeito &&
                                                                  /* @__PURE__ */ (void 0)(
                                                                    "span",
                                                                    {
                                                                      children: [
                                                                        "último def.: ",
                                                                        /* @__PURE__ */ (void 0)(
                                                                          "b",
                                                                          {
                                                                            className:
                                                                              "text-foreground",
                                                                            children:
                                                                              e.ultimoDefeito,
                                                                          },
                                                                          void 0,
                                                                          false,
                                                                          {
                                                                            fileName: _jsxFileName,
                                                                            lineNumber: 1175,
                                                                            columnNumber: 48,
                                                                          },
                                                                          this,
                                                                        ),
                                                                      ],
                                                                    },
                                                                    void 0,
                                                                    true,
                                                                    {
                                                                      fileName: _jsxFileName,
                                                                      lineNumber: 1174,
                                                                      columnNumber: 51,
                                                                    },
                                                                    this,
                                                                  ),
                                                                e.descricao &&
                                                                  /* @__PURE__ */ (void 0)(
                                                                    "span",
                                                                    {
                                                                      className: "basis-full",
                                                                      children: [
                                                                        "descrição: ",
                                                                        /* @__PURE__ */ (void 0)(
                                                                          "b",
                                                                          {
                                                                            className:
                                                                              "text-foreground",
                                                                            children: e.descricao,
                                                                          },
                                                                          void 0,
                                                                          false,
                                                                          {
                                                                            fileName: _jsxFileName,
                                                                            lineNumber: 1178,
                                                                            columnNumber: 46,
                                                                          },
                                                                          this,
                                                                        ),
                                                                      ],
                                                                    },
                                                                    void 0,
                                                                    true,
                                                                    {
                                                                      fileName: _jsxFileName,
                                                                      lineNumber: 1177,
                                                                      columnNumber: 47,
                                                                    },
                                                                    this,
                                                                  ),
                                                                e.observacoes &&
                                                                  /* @__PURE__ */ (void 0)(
                                                                    "span",
                                                                    {
                                                                      className: "basis-full",
                                                                      children: [
                                                                        "obs: ",
                                                                        e.observacoes,
                                                                      ],
                                                                    },
                                                                    void 0,
                                                                    true,
                                                                    {
                                                                      fileName: _jsxFileName,
                                                                      lineNumber: 1180,
                                                                      columnNumber: 49,
                                                                    },
                                                                    this,
                                                                  ),
                                                              ],
                                                            },
                                                            void 0,
                                                            true,
                                                            {
                                                              fileName: _jsxFileName,
                                                              lineNumber: 1158,
                                                              columnNumber: 29,
                                                            },
                                                            this,
                                                          ),
                                                        ],
                                                      },
                                                      void 0,
                                                      true,
                                                      {
                                                        fileName: _jsxFileName,
                                                        lineNumber: 1133,
                                                        columnNumber: 27,
                                                      },
                                                      this,
                                                    ),
                                                    editingDescEventId === e.id &&
                                                      /* @__PURE__ */ (void 0)(
                                                        "div",
                                                        {
                                                          className:
                                                            "mt-1 flex flex-col gap-2 rounded-md border border-primary/30 bg-primary/5 p-2.5",
                                                          children: [
                                                            /* @__PURE__ */ (void 0)(
                                                              "div",
                                                              {
                                                                className:
                                                                  "flex items-center justify-between",
                                                                children: [
                                                                  /* @__PURE__ */ (void 0)(
                                                                    "span",
                                                                    {
                                                                      className:
                                                                        "text-xs font-semibold text-foreground",
                                                                      children:
                                                                        "Editar Descrição do Evento",
                                                                    },
                                                                    void 0,
                                                                    false,
                                                                    {
                                                                      fileName: _jsxFileName,
                                                                      lineNumber: 1186,
                                                                      columnNumber: 33,
                                                                    },
                                                                    this,
                                                                  ),
                                                                  /* @__PURE__ */ (void 0)(
                                                                    "span",
                                                                    {
                                                                      className:
                                                                        "text-[10px] text-muted-foreground",
                                                                      children: EVENT_LABEL[e.kind],
                                                                    },
                                                                    void 0,
                                                                    false,
                                                                    {
                                                                      fileName: _jsxFileName,
                                                                      lineNumber: 1189,
                                                                      columnNumber: 33,
                                                                    },
                                                                    this,
                                                                  ),
                                                                ],
                                                              },
                                                              void 0,
                                                              true,
                                                              {
                                                                fileName: _jsxFileName,
                                                                lineNumber: 1185,
                                                                columnNumber: 31,
                                                              },
                                                              this,
                                                            ),
                                                            /* @__PURE__ */ (void 0)(
                                                              "textarea",
                                                              {
                                                                value: editingDescValue,
                                                                onChange: (ev) =>
                                                                  setEditingDescValue(
                                                                    ev.target.value,
                                                                  ),
                                                                className:
                                                                  "input min-h-16 text-xs bg-background",
                                                                placeholder:
                                                                  "Digite a descrição do evento...",
                                                                autoFocus: true,
                                                              },
                                                              void 0,
                                                              false,
                                                              {
                                                                fileName: _jsxFileName,
                                                                lineNumber: 1193,
                                                                columnNumber: 31,
                                                              },
                                                              this,
                                                            ),
                                                            /* @__PURE__ */ (void 0)(
                                                              "div",
                                                              {
                                                                className: "flex justify-end gap-2",
                                                                children: [
                                                                  /* @__PURE__ */ (void 0)(
                                                                    "button",
                                                                    {
                                                                      type: "button",
                                                                      onClick: () =>
                                                                        setEditingDescEventId(null),
                                                                      className:
                                                                        "btn-ghost py-1 px-2.5 text-xs",
                                                                      children: "Cancelar",
                                                                    },
                                                                    void 0,
                                                                    false,
                                                                    {
                                                                      fileName: _jsxFileName,
                                                                      lineNumber: 1195,
                                                                      columnNumber: 33,
                                                                    },
                                                                    this,
                                                                  ),
                                                                  /* @__PURE__ */ (void 0)(
                                                                    "button",
                                                                    {
                                                                      type: "button",
                                                                      onClick: () => {
                                                                        updateEvent(e.id, {
                                                                          descricao:
                                                                            editingDescValue.trim(),
                                                                        });
                                                                        setEditingDescEventId(null);
                                                                      },
                                                                      className:
                                                                        "btn-primary py-1 px-3 text-xs",
                                                                      children: "Salvar Descrição",
                                                                    },
                                                                    void 0,
                                                                    false,
                                                                    {
                                                                      fileName: _jsxFileName,
                                                                      lineNumber: 1198,
                                                                      columnNumber: 33,
                                                                    },
                                                                    this,
                                                                  ),
                                                                ],
                                                              },
                                                              void 0,
                                                              true,
                                                              {
                                                                fileName: _jsxFileName,
                                                                lineNumber: 1194,
                                                                columnNumber: 31,
                                                              },
                                                              this,
                                                            ),
                                                          ],
                                                        },
                                                        void 0,
                                                        true,
                                                        {
                                                          fileName: _jsxFileName,
                                                          lineNumber: 1184,
                                                          columnNumber: 59,
                                                        },
                                                        this,
                                                      ),
                                                  ],
                                                },
                                                e.id,
                                                true,
                                                {
                                                  fileName: _jsxFileName,
                                                  lineNumber: 1105,
                                                  columnNumber: 24,
                                                },
                                                this,
                                              );
                                            }),
                                          },
                                          void 0,
                                          false,
                                          {
                                            fileName: _jsxFileName,
                                            lineNumber: 1083,
                                            columnNumber: 26,
                                          },
                                          this,
                                        ),
                                },
                                void 0,
                                false,
                                {
                                  fileName: _jsxFileName,
                                  lineNumber: 1080,
                                  columnNumber: 15,
                                },
                                this,
                              ),
                            ],
                          },
                          void 0,
                          true,
                          {
                            fileName: _jsxFileName,
                            lineNumber: 1062,
                            columnNumber: 13,
                          },
                          this,
                        ),
                      ],
                    },
                    void 0,
                    true,
                    {
                      fileName: _jsxFileName,
                      lineNumber: 780,
                      columnNumber: 15,
                    },
                    this,
                  ),
          },
          void 0,
          false,
          {
            fileName: _jsxFileName,
            lineNumber: 762,
            columnNumber: 7,
          },
          this,
        ),
        pendingKind &&
          /* @__PURE__ */ (void 0)(
            ActivityModal,
            {
              kind: pendingKind,
              draft: pendingDraft,
              onChange: setPendingDraft,
              onCancel: () => {
                setPendingKind(null);
                setPendingDraft({});
              },
              onConfirm: () => addEvent(pendingKind, pendingDraft),
            },
            void 0,
            false,
            {
              fileName: _jsxFileName,
              lineNumber: 1216,
              columnNumber: 23,
            },
            this,
          ),
        isSapModalOpen &&
          /* @__PURE__ */ (void 0)(
            SapIw41Modal,
            {
              isOpen: isSapModalOpen,
              onClose: () => setIsSapModalOpen(false),
              header: state.header,
              events: state.events,
              onSuccess: (filename) =>
                setSaveNotification(`✓ Planilha SAP IW41 gerada com sucesso: ${filename}`),
            },
            void 0,
            false,
            {
              fileName: _jsxFileName,
              lineNumber: 1221,
              columnNumber: 26,
            },
            this,
          ),
        saveNotification &&
          /* @__PURE__ */ (void 0)(
            "div",
            {
              className:
                "fixed bottom-5 right-5 z-50 flex items-center gap-3 rounded-xl bg-emerald-600 px-4 py-3 text-xs sm:text-sm font-semibold text-white shadow-2xl transition-all animate-in fade-in slide-in-from-bottom-4 border border-emerald-400/40",
              children: [
                /* @__PURE__ */ (void 0)(
                  Cloud,
                  { className: "h-4 w-4 shrink-0 animate-pulse text-white" },
                  void 0,
                  false,
                  {
                    fileName: _jsxFileName,
                    lineNumber: 1224,
                    columnNumber: 11,
                  },
                  this,
                ),
                /* @__PURE__ */ (void 0)(
                  "span",
                  { children: saveNotification },
                  void 0,
                  false,
                  {
                    fileName: _jsxFileName,
                    lineNumber: 1225,
                    columnNumber: 11,
                  },
                  this,
                ),
              ],
            },
            void 0,
            true,
            {
              fileName: _jsxFileName,
              lineNumber: 1223,
              columnNumber: 28,
            },
            this,
          ),
      ],
    },
    void 0,
    true,
    {
      fileName: _jsxFileName,
      lineNumber: 698,
      columnNumber: 10,
    },
    this,
  );
}
function Card({ title, children }) {
  return /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
    "div",
    {
      className: "rounded-xl border border-border bg-card p-4 shadow-sm",
      children: [
        /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
          "h2",
          {
            className: "mb-3 text-sm font-semibold uppercase tracking-wider text-muted-foreground",
            children: title,
          },
          void 0,
          false,
          {
            fileName: _jsxFileName,
            lineNumber: 1237,
            columnNumber: 7,
          },
          this,
        ),
        children,
      ],
    },
    void 0,
    true,
    {
      fileName: _jsxFileName,
      lineNumber: 1236,
      columnNumber: 10,
    },
    this,
  );
}
function Field({ label, full, action, children }) {
  return /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
    "div",
    {
      className: `flex flex-col gap-1 text-xs font-medium text-muted-foreground ${full ? "col-span-2" : ""}`,
      children: [
        /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
          "div",
          {
            className: "flex items-center justify-between min-h-[20px]",
            children: [
              /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                "span",
                { children: label },
                void 0,
                false,
                {
                  fileName: _jsxFileName,
                  lineNumber: 1256,
                  columnNumber: 9,
                },
                this,
              ),
              action,
            ],
          },
          void 0,
          true,
          {
            fileName: _jsxFileName,
            lineNumber: 1255,
            columnNumber: 7,
          },
          this,
        ),
        children,
      ],
    },
    void 0,
    true,
    {
      fileName: _jsxFileName,
      lineNumber: 1254,
      columnNumber: 10,
    },
    this,
  );
}
function Stat({ label, value, tone }) {
  const toneClass =
    tone === "primary"
      ? "text-primary"
      : tone === "success"
        ? "text-[color:var(--success)]"
        : tone === "warning"
          ? "text-[color:var(--warning)]"
          : "text-foreground";
  return /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
    "div",
    {
      className: "rounded-lg border border-border bg-background px-3 py-3",
      children: [
        /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
          "div",
          {
            className: "text-[10px] font-semibold uppercase tracking-widest text-muted-foreground",
            children: label,
          },
          void 0,
          false,
          {
            fileName: _jsxFileName,
            lineNumber: 1273,
            columnNumber: 7,
          },
          this,
        ),
        /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
          "div",
          {
            className: `mt-1 font-mono text-xl tabular-nums ${toneClass}`,
            children: value,
          },
          void 0,
          false,
          {
            fileName: _jsxFileName,
            lineNumber: 1276,
            columnNumber: 7,
          },
          this,
        ),
      ],
    },
    void 0,
    true,
    {
      fileName: _jsxFileName,
      lineNumber: 1272,
      columnNumber: 10,
    },
    this,
  );
}
function EventBtn({ children, onClick, disabled, variant }) {
  return /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
    "button",
    {
      onClick,
      disabled,
      className: `rounded-lg px-3 py-3 text-sm font-medium transition-colors disabled:cursor-not-allowed disabled:opacity-40 ${variant === "primary" ? "bg-primary text-primary-foreground hover:bg-primary/90" : variant === "success" ? "bg-[color:var(--success)] text-[color:var(--success-foreground)] hover:opacity-90" : variant === "warning" ? "bg-[color:var(--warning)] text-[color:var(--warning-foreground)] hover:opacity-90" : variant === "danger" ? "bg-destructive text-destructive-foreground hover:bg-destructive/90" : "bg-secondary text-secondary-foreground hover:bg-secondary/80"}`,
      children,
    },
    void 0,
    false,
    {
      fileName: _jsxFileName,
      lineNumber: 1292,
      columnNumber: 10,
    },
    this,
  );
}
function ActivityModal({ kind, draft, onChange, onCancel, onConfirm }) {
  const [errorMsg, setErrorMsg] = (0, import_react.useState)("");
  const isStart = kind === "inicio_atividade";
  const isEnd = kind === "fim_atividade";
  const isDefect = kind === "dimensionamento_defeito";
  const isCancelado = kind === "cancelado";
  const isProntoPatio = kind === "pronto_patio";
  const descLabel =
    kind === "cancelado"
      ? "Motivo do cancelamento *"
      : kind === "parada"
        ? "Motivo da parada"
        : kind === "outros"
          ? "Descreva o evento"
          : "Descrição (opcional)";
  const handleConfirm = () => {
    if (isCancelado && (!draft.descricao || !draft.descricao.trim())) {
      setErrorMsg("É obrigatório informar o motivo do cancelamento.");
      return;
    }
    if (isProntoPatio && (!draft.local || !draft.local.trim())) {
      setErrorMsg("É obrigatório informar o local ao acionar Pronto pra sair no pátio.");
      return;
    }
    setErrorMsg("");
    onConfirm();
  };
  return /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
    "div",
    {
      className: "fixed inset-0 z-50 flex items-center justify-center bg-black/40 px-4",
      children: /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
        "div",
        {
          className: "w-full max-w-md rounded-xl border border-border bg-card p-5 shadow-xl",
          children: [
            /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
              "h3",
              {
                className: "text-lg font-semibold",
                children: EVENT_LABEL[kind],
              },
              void 0,
              false,
              {
                fileName: _jsxFileName,
                lineNumber: 1330,
                columnNumber: 9,
              },
              this,
            ),
            /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
              "p",
              {
                className: "mt-1 text-xs text-muted-foreground",
                children: ["Hora será registrada como agora: ", formatTime(Date.now())],
              },
              void 0,
              true,
              {
                fileName: _jsxFileName,
                lineNumber: 1331,
                columnNumber: 9,
              },
              this,
            ),
            /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
              "div",
              {
                className: "mt-4 grid grid-cols-2 gap-3",
                children: [
                  isStart &&
                    /* @__PURE__ */ (void 0)(
                      Field,
                      {
                        label: "Km/poste inicial",
                        full: true,
                        children: /* @__PURE__ */ (void 0)(
                          "input",
                          {
                            autoFocus: true,
                            value: draft.kmInicial ?? "",
                            onChange: (e) =>
                              onChange({
                                ...draft,
                                kmInicial: e.target.value,
                              }),
                            className: "input",
                          },
                          void 0,
                          false,
                          {
                            fileName: _jsxFileName,
                            lineNumber: 1336,
                            columnNumber: 15,
                          },
                          this,
                        ),
                      },
                      void 0,
                      false,
                      {
                        fileName: _jsxFileName,
                        lineNumber: 1335,
                        columnNumber: 23,
                      },
                      this,
                    ),
                  isEnd &&
                    /* @__PURE__ */ (void 0)(
                      import_jsx_dev_runtime.Fragment,
                      {
                        children: [
                          /* @__PURE__ */ (void 0)(
                            Field,
                            {
                              label: "Km/poste final",
                              full: true,
                              children: /* @__PURE__ */ (void 0)(
                                "input",
                                {
                                  autoFocus: true,
                                  value: draft.kmFinal ?? "",
                                  onChange: (e) =>
                                    onChange({
                                      ...draft,
                                      kmFinal: e.target.value,
                                    }),
                                  className: "input",
                                },
                                void 0,
                                false,
                                {
                                  fileName: _jsxFileName,
                                  lineNumber: 1343,
                                  columnNumber: 17,
                                },
                                this,
                              ),
                            },
                            void 0,
                            false,
                            {
                              fileName: _jsxFileName,
                              lineNumber: 1342,
                              columnNumber: 15,
                            },
                            this,
                          ),
                          /* @__PURE__ */ (void 0)(
                            Field,
                            {
                              label: "Produção (m)",
                              children: /* @__PURE__ */ (void 0)(
                                "input",
                                {
                                  inputMode: "numeric",
                                  value: draft.producao ?? "",
                                  onChange: (e) =>
                                    onChange({
                                      ...draft,
                                      producao: e.target.value,
                                    }),
                                  className: "input",
                                },
                                void 0,
                                false,
                                {
                                  fileName: _jsxFileName,
                                  lineNumber: 1349,
                                  columnNumber: 17,
                                },
                                this,
                              ),
                            },
                            void 0,
                            false,
                            {
                              fileName: _jsxFileName,
                              lineNumber: 1348,
                              columnNumber: 15,
                            },
                            this,
                          ),
                          /* @__PURE__ */ (void 0)(
                            Field,
                            {
                              label: "Nº defeitos",
                              children: /* @__PURE__ */ (void 0)(
                                "input",
                                {
                                  inputMode: "numeric",
                                  value: draft.defeitos ?? "",
                                  onChange: (e) =>
                                    onChange({
                                      ...draft,
                                      defeitos: e.target.value,
                                    }),
                                  className: "input",
                                },
                                void 0,
                                false,
                                {
                                  fileName: _jsxFileName,
                                  lineNumber: 1355,
                                  columnNumber: 17,
                                },
                                this,
                              ),
                            },
                            void 0,
                            false,
                            {
                              fileName: _jsxFileName,
                              lineNumber: 1354,
                              columnNumber: 15,
                            },
                            this,
                          ),
                          /* @__PURE__ */ (void 0)(
                            Field,
                            {
                              label: "Nº último defeito",
                              full: true,
                              children: /* @__PURE__ */ (void 0)(
                                "input",
                                {
                                  value: draft.ultimoDefeito ?? "",
                                  onChange: (e) =>
                                    onChange({
                                      ...draft,
                                      ultimoDefeito: e.target.value,
                                    }),
                                  className: "input",
                                },
                                void 0,
                                false,
                                {
                                  fileName: _jsxFileName,
                                  lineNumber: 1361,
                                  columnNumber: 17,
                                },
                                this,
                              ),
                            },
                            void 0,
                            false,
                            {
                              fileName: _jsxFileName,
                              lineNumber: 1360,
                              columnNumber: 15,
                            },
                            this,
                          ),
                        ],
                      },
                      void 0,
                      true,
                      {
                        fileName: _jsxFileName,
                        lineNumber: 1341,
                        columnNumber: 21,
                      },
                      this,
                    ),
                  isDefect &&
                    /* @__PURE__ */ (void 0)(
                      Field,
                      {
                        label: "Nº do novo defeito",
                        full: true,
                        children: /* @__PURE__ */ (void 0)(
                          "input",
                          {
                            autoFocus: true,
                            value: draft.ultimoDefeito ?? "",
                            onChange: (e) =>
                              onChange({
                                ...draft,
                                ultimoDefeito: e.target.value,
                              }),
                            className: "input",
                          },
                          void 0,
                          false,
                          {
                            fileName: _jsxFileName,
                            lineNumber: 1368,
                            columnNumber: 15,
                          },
                          this,
                        ),
                      },
                      void 0,
                      false,
                      {
                        fileName: _jsxFileName,
                        lineNumber: 1367,
                        columnNumber: 24,
                      },
                      this,
                    ),
                  isProntoPatio &&
                    /* @__PURE__ */ (void 0)(
                      Field,
                      {
                        label: "Local *",
                        full: true,
                        children: /* @__PURE__ */ (void 0)(
                          "input",
                          {
                            autoFocus: true,
                            value: draft.local ?? "",
                            onChange: (e) => {
                              if (errorMsg) setErrorMsg("");
                              onChange({
                                ...draft,
                                local: e.target.value,
                              });
                            },
                            className: `input ${errorMsg && !draft.local?.trim() ? "border-destructive focus-visible:ring-destructive" : ""}`,
                            placeholder: "Informe o local do pátio...",
                          },
                          void 0,
                          false,
                          {
                            fileName: _jsxFileName,
                            lineNumber: 1374,
                            columnNumber: 15,
                          },
                          this,
                        ),
                      },
                      void 0,
                      false,
                      {
                        fileName: _jsxFileName,
                        lineNumber: 1373,
                        columnNumber: 29,
                      },
                      this,
                    ),
                  /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                    Field,
                    {
                      label: descLabel,
                      full: true,
                      children: [
                        /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                          "textarea",
                          {
                            value: draft.descricao ?? "",
                            onChange: (e) => {
                              if (errorMsg) setErrorMsg("");
                              onChange({
                                ...draft,
                                descricao: e.target.value,
                              });
                            },
                            className: `input min-h-20 ${errorMsg ? "border-destructive focus-visible:ring-destructive" : ""}`,
                            placeholder: isCancelado
                              ? "Digite o motivo do cancelamento obrigatoriamente..."
                              : void 0,
                            autoFocus: isCancelado,
                          },
                          void 0,
                          false,
                          {
                            fileName: _jsxFileName,
                            lineNumber: 1383,
                            columnNumber: 13,
                          },
                          this,
                        ),
                        errorMsg &&
                          /* @__PURE__ */ (void 0)(
                            "span",
                            {
                              className: "text-xs font-medium text-destructive mt-1 block",
                              children: errorMsg,
                            },
                            void 0,
                            false,
                            {
                              fileName: _jsxFileName,
                              lineNumber: 1390,
                              columnNumber: 26,
                            },
                            this,
                          ),
                      ],
                    },
                    void 0,
                    true,
                    {
                      fileName: _jsxFileName,
                      lineNumber: 1382,
                      columnNumber: 11,
                    },
                    this,
                  ),
                  /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                    Field,
                    {
                      label: "Observações",
                      full: true,
                      children: /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                        "textarea",
                        {
                          value: draft.observacoes ?? "",
                          onChange: (e) =>
                            onChange({
                              ...draft,
                              observacoes: e.target.value,
                            }),
                          className: "input min-h-20",
                        },
                        void 0,
                        false,
                        {
                          fileName: _jsxFileName,
                          lineNumber: 1393,
                          columnNumber: 13,
                        },
                        this,
                      ),
                    },
                    void 0,
                    false,
                    {
                      fileName: _jsxFileName,
                      lineNumber: 1392,
                      columnNumber: 11,
                    },
                    this,
                  ),
                ],
              },
              void 0,
              true,
              {
                fileName: _jsxFileName,
                lineNumber: 1334,
                columnNumber: 9,
              },
              this,
            ),
            /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
              "div",
              {
                className: "mt-5 flex justify-end gap-2",
                children: [
                  /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                    "button",
                    {
                      onClick: onCancel,
                      className: "btn-ghost",
                      children: "Cancelar",
                    },
                    void 0,
                    false,
                    {
                      fileName: _jsxFileName,
                      lineNumber: 1400,
                      columnNumber: 11,
                    },
                    this,
                  ),
                  /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                    "button",
                    {
                      onClick: handleConfirm,
                      className: "btn-primary",
                      children: "Registrar",
                    },
                    void 0,
                    false,
                    {
                      fileName: _jsxFileName,
                      lineNumber: 1403,
                      columnNumber: 11,
                    },
                    this,
                  ),
                ],
              },
              void 0,
              true,
              {
                fileName: _jsxFileName,
                lineNumber: 1399,
                columnNumber: 9,
              },
              this,
            ),
          ],
        },
        void 0,
        true,
        {
          fileName: _jsxFileName,
          lineNumber: 1329,
          columnNumber: 7,
        },
        this,
      ),
    },
    void 0,
    false,
    {
      fileName: _jsxFileName,
      lineNumber: 1328,
      columnNumber: 10,
    },
    this,
  );
}
function TeamPicker({
  roster,
  selected,
  onSelectedChange,
  onRosterChange,
  newMember,
  setNewMember,
}) {
  const selectedList = selected
    .split(";")
    .map((n) => n.trim())
    .filter(Boolean);
  const isSelected = (n) => selectedList.includes(n);
  const toggle = (n) => {
    onSelectedChange(
      (isSelected(n) ? selectedList.filter((x) => x !== n) : [...selectedList, n]).join("; "),
    );
  };
  const addMember = () => {
    const name = newMember.trim();
    if (!name) return;
    const updatedRoster = roster.includes(name) ? roster : [...roster, name];
    if (!roster.includes(name)) onRosterChange(updatedRoster);
    if (!selectedList.includes(name)) onSelectedChange([...selectedList, name].join("; "));
    setNewMember("");
  };
  const removeFromRoster = (n, e) => {
    e.stopPropagation();
    e.preventDefault();
    onRosterChange(roster.filter((x) => x !== n));
    if (isSelected(n)) onSelectedChange(selectedList.filter((x) => x !== n).join("; "));
  };
  return /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
    "div",
    {
      className: "space-y-3",
      children: [
        /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
          "div",
          {
            className: "flex gap-2",
            children: [
              /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                "input",
                {
                  list: "roster-datalist",
                  value: newMember,
                  onChange: (e) => setNewMember(e.target.value),
                  onKeyDown: (e) => {
                    if (e.key === "Enter") {
                      e.preventDefault();
                      addMember();
                    }
                  },
                  className: "input flex-1",
                },
                void 0,
                false,
                {
                  fileName: _jsxFileName,
                  lineNumber: 1453,
                  columnNumber: 9,
                },
                this,
              ),
              /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                "datalist",
                {
                  id: "roster-datalist",
                  children: roster.map((name) =>
                    /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                      "option",
                      { value: name },
                      name,
                      false,
                      {
                        fileName: _jsxFileName,
                        lineNumber: 1460,
                        columnNumber: 31,
                      },
                      this,
                    ),
                  ),
                },
                void 0,
                false,
                {
                  fileName: _jsxFileName,
                  lineNumber: 1459,
                  columnNumber: 9,
                },
                this,
              ),
              /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                "button",
                {
                  type: "button",
                  onClick: addMember,
                  className: "btn-secondary px-3",
                  children: "Adicionar",
                },
                void 0,
                false,
                {
                  fileName: _jsxFileName,
                  lineNumber: 1462,
                  columnNumber: 9,
                },
                this,
              ),
            ],
          },
          void 0,
          true,
          {
            fileName: _jsxFileName,
            lineNumber: 1452,
            columnNumber: 7,
          },
          this,
        ),
        roster.length > 0
          ? /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
              "div",
              {
                className: "space-y-1.5",
                children: [
                  /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                    "div",
                    {
                      className: "text-[11px] font-medium text-muted-foreground",
                      children: "Clique no nome para escolher / desmarcar:",
                    },
                    void 0,
                    false,
                    {
                      fileName: _jsxFileName,
                      lineNumber: 1468,
                      columnNumber: 11,
                    },
                    this,
                  ),
                  /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                    "div",
                    {
                      className: "flex flex-wrap gap-1.5",
                      children: roster.map((n) => {
                        const on = isSelected(n);
                        return /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                          "div",
                          {
                            className: `inline-flex items-center rounded-full border text-xs transition-colors overflow-hidden ${on ? "bg-primary text-primary-foreground border-primary" : "bg-background text-foreground border-border hover:bg-accent"}`,
                            children: [
                              /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                "button",
                                {
                                  type: "button",
                                  onClick: () => toggle(n),
                                  className:
                                    "flex items-center gap-1.5 px-3 py-1 text-xs font-medium focus:outline-none cursor-pointer",
                                  children: [
                                    on &&
                                      /* @__PURE__ */ (void 0)(
                                        "span",
                                        {
                                          className: "font-bold",
                                          children: "✓",
                                        },
                                        void 0,
                                        false,
                                        {
                                          fileName: _jsxFileName,
                                          lineNumber: 1476,
                                          columnNumber: 28,
                                        },
                                        this,
                                      ),
                                    /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                      "span",
                                      { children: n },
                                      void 0,
                                      false,
                                      {
                                        fileName: _jsxFileName,
                                        lineNumber: 1477,
                                        columnNumber: 21,
                                      },
                                      this,
                                    ),
                                  ],
                                },
                                void 0,
                                true,
                                {
                                  fileName: _jsxFileName,
                                  lineNumber: 1475,
                                  columnNumber: 19,
                                },
                                this,
                              ),
                              /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                                "button",
                                {
                                  type: "button",
                                  onClick: (e) => removeFromRoster(n, e),
                                  className: `px-2.5 py-1 text-xs font-bold transition-all border-l cursor-pointer ${on ? "text-primary-foreground/80 hover:text-white hover:bg-destructive border-primary-foreground/20" : "text-muted-foreground hover:text-destructive hover:bg-destructive/10 border-border"}`,
                                  title: `Remover ${n} do histórico`,
                                  "aria-label": `Remover ${n} do histórico`,
                                  children: "×",
                                },
                                void 0,
                                false,
                                {
                                  fileName: _jsxFileName,
                                  lineNumber: 1479,
                                  columnNumber: 19,
                                },
                                this,
                              ),
                            ],
                          },
                          n,
                          true,
                          {
                            fileName: _jsxFileName,
                            lineNumber: 1474,
                            columnNumber: 18,
                          },
                          this,
                        );
                      }),
                    },
                    void 0,
                    false,
                    {
                      fileName: _jsxFileName,
                      lineNumber: 1471,
                      columnNumber: 11,
                    },
                    this,
                  ),
                ],
              },
              void 0,
              true,
              {
                fileName: _jsxFileName,
                lineNumber: 1467,
                columnNumber: 28,
              },
              this,
            )
          : /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
              "span",
              {
                className: "text-xs text-muted-foreground",
                children: "Nenhum nome cadastrado. Digite acima para adicionar.",
              },
              void 0,
              false,
              {
                fileName: _jsxFileName,
                lineNumber: 1485,
                columnNumber: 18,
              },
              this,
            ),
        selectedList.length > 0 &&
          /* @__PURE__ */ (void 0)(
            "div",
            {
              className: "rounded-md bg-secondary/50 p-2 text-[11px] text-foreground",
              children: [
                /* @__PURE__ */ (void 0)(
                  "span",
                  {
                    className: "font-semibold",
                    children: ["Selecionados para a atividade (", selectedList.length, "):"],
                  },
                  void 0,
                  true,
                  {
                    fileName: _jsxFileName,
                    lineNumber: 1490,
                    columnNumber: 11,
                  },
                  this,
                ),
                " ",
                selectedList.join("; "),
              ],
            },
            void 0,
            true,
            {
              fileName: _jsxFileName,
              lineNumber: 1489,
              columnNumber: 35,
            },
            this,
          ),
      ],
    },
    void 0,
    true,
    {
      fileName: _jsxFileName,
      lineNumber: 1451,
      columnNumber: 10,
    },
    this,
  );
}
function LastDefectCard({ value: last, onSave }) {
  const [editing, setEditing] = (0, import_react.useState)(false);
  const [value, setValue] = (0, import_react.useState)(last);
  (0, import_react.useEffect)(() => {
    if (!editing) setValue(last);
  }, [last, editing]);
  return /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
    "div",
    {
      className: "rounded-xl border-2 border-primary/40 bg-primary/5 p-4 shadow-sm",
      children: /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
        "div",
        {
          className: "flex items-center justify-between gap-3",
          children: [
            /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
              "div",
              {
                className: "min-w-0",
                children: [
                  /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                    "div",
                    {
                      className: "text-[11px] font-semibold uppercase tracking-wider text-primary",
                      children: "Último defeito encontrado",
                    },
                    void 0,
                    false,
                    {
                      fileName: _jsxFileName,
                      lineNumber: 1512,
                      columnNumber: 11,
                    },
                    this,
                  ),
                  editing
                    ? /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                        "input",
                        {
                          autoFocus: true,
                          value,
                          onChange: (e) => setValue(e.target.value),
                          onKeyDown: (e) => {
                            if (e.key === "Enter") {
                              onSave(value.trim());
                              setEditing(false);
                            }
                            if (e.key === "Escape") setEditing(false);
                          },
                          className: "input mt-1 text-2xl font-bold",
                        },
                        void 0,
                        false,
                        {
                          fileName: _jsxFileName,
                          lineNumber: 1515,
                          columnNumber: 22,
                        },
                        this,
                      )
                    : /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                        "div",
                        {
                          className: "mt-1 truncate text-3xl font-bold text-foreground",
                          children:
                            last ||
                            /* @__PURE__ */ (void 0)(
                              "span",
                              {
                                className: "text-muted-foreground text-xl font-normal",
                                children: "— nenhum registrado —",
                              },
                              void 0,
                              false,
                              {
                                fileName: _jsxFileName,
                                lineNumber: 1524,
                                columnNumber: 24,
                              },
                              this,
                            ),
                        },
                        void 0,
                        false,
                        {
                          fileName: _jsxFileName,
                          lineNumber: 1523,
                          columnNumber: 59,
                        },
                        this,
                      ),
                ],
              },
              void 0,
              true,
              {
                fileName: _jsxFileName,
                lineNumber: 1511,
                columnNumber: 9,
              },
              this,
            ),
            editing
              ? /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                  "div",
                  {
                    className: "flex gap-2 shrink-0",
                    children: [
                      /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                        "button",
                        {
                          className: "btn-primary",
                          onClick: () => {
                            onSave(value.trim());
                            setEditing(false);
                          },
                          children: "Salvar",
                        },
                        void 0,
                        false,
                        {
                          fileName: _jsxFileName,
                          lineNumber: 1530,
                          columnNumber: 13,
                        },
                        this,
                      ),
                      /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                        "button",
                        {
                          className: "btn-ghost",
                          onClick: () => setEditing(false),
                          children: "Cancelar",
                        },
                        void 0,
                        false,
                        {
                          fileName: _jsxFileName,
                          lineNumber: 1536,
                          columnNumber: 13,
                        },
                        this,
                      ),
                    ],
                  },
                  void 0,
                  true,
                  {
                    fileName: _jsxFileName,
                    lineNumber: 1529,
                    columnNumber: 20,
                  },
                  this,
                )
              : /* @__PURE__ */ (0, import_jsx_dev_runtime.jsxDEV)(
                  "button",
                  {
                    className: "btn-secondary shrink-0",
                    onClick: () => setEditing(true),
                    children: "Editar",
                  },
                  void 0,
                  false,
                  {
                    fileName: _jsxFileName,
                    lineNumber: 1539,
                    columnNumber: 20,
                  },
                  this,
                ),
          ],
        },
        void 0,
        true,
        {
          fileName: _jsxFileName,
          lineNumber: 1510,
          columnNumber: 7,
        },
        this,
      ),
    },
    void 0,
    false,
    {
      fileName: _jsxFileName,
      lineNumber: 1509,
      columnNumber: 10,
    },
    this,
  );
}
//#endregion
export { PartesDiariasPage as component };
