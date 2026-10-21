import { r as __exportAll } from "../_runtime.mjs";
//#region node_modules/@node-projects/excelforge/dist/core/types.js
var CellError = class CellError {
  constructor(error) {
    this.error = error;
  }
  toString() {
    return this.error;
  }
  static {
    this.NULL = new CellError("#NULL!");
  }
  static {
    this.DIV0 = new CellError("#DIV/0!");
  }
  static {
    this.VALUE = new CellError("#VALUE!");
  }
  static {
    this.REF = new CellError("#REF!");
  }
  static {
    this.NAME = new CellError("#NAME?");
  }
  static {
    this.NUM = new CellError("#NUM!");
  }
  static {
    this.NA = new CellError("#N/A");
  }
  static {
    this.GETTING = new CellError("#GETTING_DATA");
  }
};
//#endregion
//#region node_modules/@node-projects/excelforge/dist/utils/helpers.js
var _colCache = [];
function colIndexToLetter(n) {
  if (_colCache[n]) return _colCache[n];
  let s = "",
    v = n;
  while (v > 0) {
    const r = (v - 1) % 26;
    s = String.fromCharCode(65 + r) + s;
    v = Math.floor((v - 1) / 26);
  }
  _colCache[n] = s;
  return s;
}
function colLetterToIndex(col) {
  let n = 0;
  for (let i = 0; i < col.length; i++) n = n * 26 + (col.charCodeAt(i) - 64);
  return n;
}
function cellRefToIndices(ref) {
  const m = ref.match(/^(\$?)([A-Z]+)(\$?)(\d+)$/);
  if (!m) throw new Error(`Invalid cell ref: ${ref}`);
  return {
    col: colLetterToIndex(m[2]),
    row: parseInt(m[4], 10),
  };
}
function indicesToCellRef(row, col, abs = false) {
  const a = abs ? "$" : "";
  return `${a}${colIndexToLetter(col)}${a}${row}`;
}
function parseRange(range) {
  const [start, end] = range.split(":");
  const s = cellRefToIndices(start.replace(/\$/g, ""));
  const e = end ? cellRefToIndices(end.replace(/\$/g, "")) : s;
  return {
    startRow: s.row,
    startCol: s.col,
    endRow: e.row,
    endCol: e.col,
  };
}
var _xmlEsc = {
  "&": "&amp;",
  "<": "&lt;",
  ">": "&gt;",
  '"': "&quot;",
  "'": "&apos;",
};
var _xmlRe = /[&<>"']/g;
function escapeXml(s) {
  return _xmlRe.test(s) ? ((_xmlRe.lastIndex = 0), s.replace(_xmlRe, (ch) => _xmlEsc[ch])) : s;
}
var enc$1 = new TextEncoder();
function strToBytes(s) {
  return enc$1.encode(s);
}
function dateToSerial(d, date1904 = false) {
  const epoch = date1904 ? Date.UTC(1904, 0, 1) : Date.UTC(1899, 11, 30);
  const serial = (d.getTime() - epoch) / 864e5;
  return date1904 ? serial : serial >= 60 ? serial : serial;
}
function pxToEmu(px) {
  return Math.round((px * 914400) / 96);
}
function base64ToBytes(b64) {
  const bin = atob(b64);
  const out = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) out[i] = bin.charCodeAt(i);
  return out;
}
//#endregion
//#region node_modules/@node-projects/excelforge/dist/core/Worksheet.js
var _xmlContentRe = /[&<>]/g;
var _xmlContentEsc = {
  "&": "&amp;",
  "<": "&lt;",
  ">": "&gt;",
};
function escapeXmlContent(s) {
  return _xmlContentRe.test(s)
    ? ((_xmlContentRe.lastIndex = 0), s.replace(_xmlContentRe, (ch) => _xmlContentEsc[ch]))
    : s;
}
function colorEl(c) {
  if (c.startsWith("theme:")) {
    const rest = c.slice(6);
    const tintIdx = rest.indexOf(":tint:");
    if (tintIdx >= 0)
      return `<color theme="${rest.slice(0, tintIdx)}" tint="${rest.slice(tintIdx + 6)}"/>`;
    return `<color theme="${rest}"/>`;
  }
  return `<color rgb="${c.startsWith("#") ? "FF" + c.slice(1) : c}"/>`;
}
var SUBTOTAL_FN = {
  average: 101,
  count: 102,
  countNums: 103,
  max: 104,
  min: 105,
  stdDev: 107,
  sum: 109,
  var: 110,
  vars: 111,
};
var Worksheet = class {
  constructor(name, options = {}) {
    this.cells = /* @__PURE__ */ new Map();
    this.merges = [];
    this.images = [];
    this.cellImages = [];
    this.charts = [];
    this.conditionalFormats = [];
    this.tables = [];
    this.pivotTables = [];
    this.sparklines = [];
    this.formControls = [];
    this.shapes = [];
    this.wordArt = [];
    this.queryTables = [];
    this.tableSlicers = [];
    this.mathEquations = [];
    this.oleObjects = [];
    this.colDefs = /* @__PURE__ */ new Map();
    this.rowDefs = /* @__PURE__ */ new Map();
    this.dataValidations = /* @__PURE__ */ new Map();
    this.rowBreaks = [];
    this.colBreaks = [];
    this.preservedXml = [];
    this._nextSharedIdx = 0;
    this._isChartSheet = false;
    this._isDialogSheet = false;
    this.sheetIndex = 0;
    this.rId = "";
    this.drawingRId = "";
    this.legacyDrawingRId = "";
    this.tableRIds = [];
    this.ctrlPropRIds = [];
    this.slicerRId = "";
    this._slicerDrawingInfo = [];
    this._cellImageVm = /* @__PURE__ */ new Map();
    this.ignoreErrors = [];
    this.oleRIds = [];
    this.oleIconRIds = [];
    this.name = name;
    this.options = {
      ...options,
      name,
    };
  }
  markDirty() {
    this._workbook?.markDirty(this.name);
  }
  setActive() {
    this._workbook?.setActiveSheet(this.name);
  }
  getCell(row, col) {
    let rowMap = this.cells.get(row);
    if (!rowMap) {
      rowMap = /* @__PURE__ */ new Map();
      this.cells.set(row, rowMap);
    }
    let cell = rowMap.get(col);
    if (!cell) {
      cell = {};
      rowMap.set(col, cell);
    }
    return cell;
  }
  getCellByRef(ref) {
    const { row, col } = cellRefToIndices(ref);
    return this.getCell(row, col);
  }
  setCell(row, col, cell) {
    let rowMap = this.cells.get(row);
    if (!rowMap) {
      rowMap = /* @__PURE__ */ new Map();
      this.cells.set(row, rowMap);
    }
    rowMap.set(col, cell);
    return this;
  }
  setValue(row, col, value) {
    const cell = this.getCell(row, col);
    delete cell.formula;
    delete cell.arrayFormula;
    delete cell.richText;
    delete cell._sharedIdx;
    delete cell._sharedRef;
    delete cell._dynamic;
    cell.value = value;
    return this;
  }
  setFormula(row, col, formula) {
    const cell = this.getCell(row, col);
    delete cell.arrayFormula;
    delete cell.richText;
    delete cell._sharedIdx;
    delete cell._sharedRef;
    delete cell._dynamic;
    cell.formula = formula;
    return this;
  }
  setDynamicArrayFormula(row, col, formula) {
    const cell = this.getCell(row, col);
    delete cell.formula;
    delete cell.richText;
    delete cell._sharedIdx;
    delete cell._sharedRef;
    cell.arrayFormula = formula;
    cell._dynamic = true;
    return this;
  }
  setSharedFormula(masterRow, masterCol, formula, rangeRef) {
    const si = this._nextSharedIdx++;
    const master = this.getCell(masterRow, masterCol);
    master.formula = formula;
    master._sharedRef = rangeRef;
    master._sharedIdx = si;
    const { startRow, startCol, endRow, endCol } = parseRange(rangeRef);
    for (let r = startRow; r <= endRow; r++)
      for (let c = startCol; c <= endCol; c++) {
        if (r === masterRow && c === masterCol) continue;
        const dep = this.getCell(r, c);
        dep._sharedIdx = si;
      }
    return this;
  }
  setStyle(row, col, style) {
    this.getCell(row, col).style = style;
    return this;
  }
  writeArray(startRow, startCol, data) {
    for (let r = 0; r < data.length; r++)
      for (let c = 0; c < data[r].length; c++)
        this.setValue(startRow + r, startCol + c, data[r][c]);
    return this;
  }
  writeColumn(startRow, col, values) {
    values.forEach((v, i) => this.setValue(startRow + i, col, v));
    return this;
  }
  writeRow(row, startCol, values) {
    values.forEach((v, i) => this.setValue(row, startCol + i, v));
    return this;
  }
  setColumn(col, def) {
    this.colDefs.set(col, def);
    return this;
  }
  setColumnWidth(col, width) {
    const d = this.colDefs.get(col) ?? {};
    this.colDefs.set(col, {
      ...d,
      width,
      customWidth: true,
    });
    return this;
  }
  setRow(row, def) {
    this.rowDefs.set(row, def);
    return this;
  }
  setRowHeight(row, height) {
    const d = this.rowDefs.get(row) ?? {};
    this.rowDefs.set(row, {
      ...d,
      height,
    });
    return this;
  }
  merge(startRow, startCol, endRow, endCol) {
    this.merges.push({
      startRow,
      startCol,
      endRow,
      endCol,
    });
    return this;
  }
  mergeByRef(ref) {
    const [s, e] = ref.split(":");
    const sc = cellRefToIndices(s);
    const ec = cellRefToIndices(e);
    return this.merge(sc.row, sc.col, ec.row, ec.col);
  }
  getMerges() {
    return this.merges;
  }
  addImage(img) {
    this.images.push(img);
    return this;
  }
  getImages() {
    return this.images;
  }
  addCellImage(img) {
    this.cellImages.push(img);
    return this;
  }
  getCellImages() {
    return this.cellImages;
  }
  getComments() {
    const out = [];
    for (const [r, rowMap] of this.cells)
      for (const [c, cell] of rowMap)
        if (cell.comment)
          out.push({
            row: r,
            col: c,
            comment: cell.comment,
          });
    return out;
  }
  addChart(chart) {
    this.charts.push(chart);
    return this;
  }
  getCharts() {
    return this.charts;
  }
  addConditionalFormat(cf) {
    this.conditionalFormats.push(cf);
    return this;
  }
  getConditionalFormats() {
    return this.conditionalFormats;
  }
  getDataValidations() {
    return this.dataValidations;
  }
  addTable(table) {
    this.tables.push(table);
    if (table.totalsRow && table.columns?.length) {
      const { startRow, startCol, endRow } = parseRange(table.ref);
      const dataStart = startRow + 1;
      const dataEnd = endRow - 1;
      table.columns.forEach((col, i) => {
        const colIdx = startCol + i;
        if (col.totalsRowLabel) this.setValue(endRow, colIdx, col.totalsRowLabel);
        else if (col.totalsRowFunction && col.totalsRowFunction !== "none") {
          const fn = SUBTOTAL_FN[col.totalsRowFunction];
          if (fn !== void 0) {
            const letter = colIndexToLetter(colIdx);
            this.setFormula(
              endRow,
              colIdx,
              `SUBTOTAL(${fn},${letter}${dataStart}:${letter}${dataEnd})`,
            );
          }
        }
      });
    }
    return this;
  }
  getTables() {
    return this.tables;
  }
  addPivotTable(pt) {
    this.pivotTables.push(pt);
    return this;
  }
  getPivotTables() {
    return this.pivotTables;
  }
  readRange(ref) {
    const { startRow, startCol, endRow, endCol } = parseRange(ref);
    const result = [];
    for (let r = startRow; r <= endRow; r++) {
      const row = [];
      const rowMap = this.cells.get(r);
      for (let c = startCol; c <= endCol; c++) {
        const cell = rowMap?.get(c);
        row.push(cell?.value ?? null);
      }
      result.push(row);
    }
    return result;
  }
  readAllCells() {
    const out = [];
    for (const [r, rowMap] of this.cells)
      for (const [c, cell] of rowMap)
        out.push({
          row: r,
          col: c,
          cell,
        });
    return out;
  }
  getUsedRange() {
    let minR = Infinity,
      maxR = 0,
      minC = Infinity,
      maxC = 0;
    for (const [r, rowMap] of this.cells)
      for (const [c] of rowMap) {
        if (r < minR) minR = r;
        if (r > maxR) maxR = r;
        if (c < minC) minC = c;
        if (c > maxC) maxC = c;
      }
    return maxR === 0
      ? null
      : {
          startRow: minR,
          startCol: minC,
          endRow: maxR,
          endCol: maxC,
        };
  }
  getColumn(col) {
    return this.colDefs.get(col);
  }
  getRow(row) {
    return this.rowDefs.get(row);
  }
  insertRows(atRow, count) {
    const rows = [...this.cells.keys()].filter((r) => r >= atRow).sort((a, b) => b - a);
    for (const r of rows) {
      const rowMap = this.cells.get(r);
      this.cells.delete(r);
      this.cells.set(r + count, rowMap);
    }
    const rdKeys = [...this.rowDefs.keys()].filter((r) => r >= atRow).sort((a, b) => b - a);
    for (const r of rdKeys) {
      const d = this.rowDefs.get(r);
      this.rowDefs.delete(r);
      this.rowDefs.set(r + count, d);
    }
    for (const m of this.merges) {
      if (m.startRow >= atRow) m.startRow += count;
      if (m.endRow >= atRow) m.endRow += count;
    }
    return this;
  }
  deleteRows(atRow, count) {
    for (let r = atRow; r < atRow + count; r++) {
      this.cells.delete(r);
      this.rowDefs.delete(r);
    }
    const rows = [...this.cells.keys()].filter((r) => r >= atRow + count).sort((a, b) => a - b);
    for (const r of rows) {
      const rowMap = this.cells.get(r);
      this.cells.delete(r);
      this.cells.set(r - count, rowMap);
    }
    const rdKeys = [...this.rowDefs.keys()].filter((r) => r >= atRow + count).sort((a, b) => a - b);
    for (const r of rdKeys) {
      const d = this.rowDefs.get(r);
      this.rowDefs.delete(r);
      this.rowDefs.set(r - count, d);
    }
    this.merges = this.merges.filter((m) => !(m.startRow >= atRow && m.endRow < atRow + count));
    for (const m of this.merges) {
      if (m.startRow >= atRow + count) m.startRow -= count;
      if (m.endRow >= atRow + count) m.endRow -= count;
    }
    return this;
  }
  insertColumns(atCol, count) {
    for (const [, rowMap] of this.cells) {
      const cols = [...rowMap.keys()].filter((c) => c >= atCol).sort((a, b) => b - a);
      for (const c of cols) {
        const cell = rowMap.get(c);
        rowMap.delete(c);
        rowMap.set(c + count, cell);
      }
    }
    const cdKeys = [...this.colDefs.keys()].filter((c) => c >= atCol).sort((a, b) => b - a);
    for (const c of cdKeys) {
      const d = this.colDefs.get(c);
      this.colDefs.delete(c);
      this.colDefs.set(c + count, d);
    }
    for (const m of this.merges) {
      if (m.startCol >= atCol) m.startCol += count;
      if (m.endCol >= atCol) m.endCol += count;
    }
    return this;
  }
  copyRange(srcRef, targetRow, targetCol) {
    const { startRow, startCol, endRow, endCol } = parseRange(srcRef);
    for (let r = startRow; r <= endRow; r++)
      for (let c = startCol; c <= endCol; c++) {
        const src = this.getCell(r, c);
        const dr = targetRow + (r - startRow);
        const dc = targetCol + (c - startCol);
        if (src.value != null) this.setValue(dr, dc, src.value);
        if (src.formula) this.setFormula(dr, dc, src.formula);
        if (src.style) this.setStyle(dr, dc, { ...src.style });
      }
    return this;
  }
  moveRange(srcRef, targetRow, targetCol) {
    const { startRow, startCol, endRow, endCol } = parseRange(srcRef);
    this.copyRange(srcRef, targetRow, targetCol);
    targetRow + (endRow - startRow);
    targetCol + (endCol - startCol);
    for (let r = startRow; r <= endRow; r++)
      for (let c = startCol; c <= endCol; c++) {
        const dr = targetRow + (r - startRow);
        const dc = targetCol + (c - startCol);
        if (dr === r && dc === c) continue;
        const rowMap = this.cells.get(r);
        if (rowMap) rowMap.delete(c);
      }
    return this;
  }
  sortRange(ref, sortCol, order = "asc") {
    const { startRow, startCol, endRow, endCol } = parseRange(ref);
    const rows = [];
    for (let r = startRow; r <= endRow; r++) {
      const rowMap = this.cells.get(r);
      const subset = /* @__PURE__ */ new Map();
      for (let c = startCol; c <= endCol; c++) {
        const cell = rowMap?.get(c);
        if (cell) subset.set(c, { ...cell });
      }
      rows.push({
        rowIdx: r,
        cells: subset,
      });
    }
    rows.sort((a, b) => {
      const va = a.cells.get(sortCol)?.value;
      const vb = b.cells.get(sortCol)?.value;
      const na = typeof va === "number" ? va : typeof va === "string" ? va : "";
      const nb = typeof vb === "number" ? vb : typeof vb === "string" ? vb : "";
      let cmp = 0;
      if (typeof na === "number" && typeof nb === "number") cmp = na - nb;
      else cmp = String(na).localeCompare(String(nb));
      return order === "desc" ? -cmp : cmp;
    });
    for (let i = 0; i < rows.length; i++) {
      const r = startRow + i;
      for (let c = startCol; c <= endCol; c++) {
        const cell = rows[i].cells.get(c);
        const rowMap = this.cells.get(r) ?? /* @__PURE__ */ new Map();
        if (!this.cells.has(r)) this.cells.set(r, rowMap);
        if (cell) rowMap.set(c, cell);
        else rowMap.delete(c);
      }
    }
    return this;
  }
  fillNumber(startRow, col, count, startValue = 0, step = 1) {
    for (let i = 0; i < count; i++) this.setValue(startRow + i, col, startValue + i * step);
    return this;
  }
  fillDate(startRow, col, count, startDate, unit = "day", step = 1) {
    for (let i = 0; i < count; i++) {
      const d = new Date(startDate);
      switch (unit) {
        case "day":
          d.setDate(d.getDate() + i * step);
          break;
        case "week":
          d.setDate(d.getDate() + i * step * 7);
          break;
        case "month":
          d.setMonth(d.getMonth() + i * step);
          break;
        case "year":
          d.setFullYear(d.getFullYear() + i * step);
          break;
      }
      this.setValue(startRow + i, col, d);
    }
    return this;
  }
  fillList(startRow, col, list, count) {
    for (let i = 0; i < count; i++) this.setValue(startRow + i, col, list[i % list.length]);
    return this;
  }
  autoFitColumns(minWidth = 8, maxWidth = 60) {
    const range = this.getUsedRange();
    if (!range) return this;
    for (let c = range.startCol; c <= range.endCol; c++) {
      let maxLen = 0;
      for (let r = range.startRow; r <= range.endRow; r++) {
        const cell = this.cells.get(r)?.get(c);
        if (cell?.value != null) {
          const s = String(cell.value);
          if (s.length > maxLen) maxLen = s.length;
        }
        if (cell?.richText) {
          const s = cell.richText.map((r) => r.text).join("");
          if (s.length > maxLen) maxLen = s.length;
        }
      }
      if (maxLen > 0) {
        const w = Math.max(minWidth, Math.min(maxWidth, maxLen * 1.2 + 2));
        this.setColumn(c, {
          ...(this.colDefs.get(c) ?? {}),
          width: w,
          customWidth: true,
        });
      }
    }
    return this;
  }
  duplicateRow(sourceRow, targetRow) {
    this.insertRows(targetRow, 1);
    const srcMap = this.cells.get(sourceRow >= targetRow ? sourceRow + 1 : sourceRow);
    if (srcMap) {
      const newMap = /* @__PURE__ */ new Map();
      for (const [c, cell] of srcMap)
        newMap.set(c, {
          ...cell,
          style: cell.style ? { ...cell.style } : void 0,
        });
      this.cells.set(targetRow, newMap);
    }
    const srcDef = this.rowDefs.get(sourceRow >= targetRow ? sourceRow + 1 : sourceRow);
    if (srcDef) this.rowDefs.set(targetRow, { ...srcDef });
    return this;
  }
  spliceRows(startRow, deleteCount, newRows) {
    if (deleteCount > 0) this.deleteRows(startRow, deleteCount);
    if (newRows && newRows.length > 0) {
      this.insertRows(startRow, newRows.length);
      for (let i = 0; i < newRows.length; i++) this.writeRow(startRow + i, 1, newRows[i]);
    }
    return this;
  }
  setAutoFilter(ref, opts) {
    this.autoFilter = { ref };
    if (opts?.columns) this._filterColumns = opts.columns;
    return this;
  }
  addSparkline(s) {
    this.sparklines.push(s);
    return this;
  }
  getSparklines() {
    return this.sparklines;
  }
  addDataValidation(sqref, dv) {
    this.dataValidations.set(sqref, dv);
    return this;
  }
  addRowBreak(row, manual = true) {
    this.rowBreaks.push({
      id: row,
      manual,
    });
    return this;
  }
  addColBreak(col, manual = true) {
    this.colBreaks.push({
      id: col,
      manual,
    });
    return this;
  }
  getRowBreaks() {
    return this.rowBreaks;
  }
  getColBreaks() {
    return this.colBreaks;
  }
  addFormControl(ctrl) {
    if (!ctrl.to && (ctrl.width || ctrl.height)) {
      const COL_PX = 64,
        ROW_PX = 20;
      const w = ctrl.width ?? 100,
        h = ctrl.height ?? 30;
      const endColFrac = ctrl.from.col + w / COL_PX;
      const endRowFrac = ctrl.from.row + h / ROW_PX;
      ctrl = {
        ...ctrl,
        to: {
          col: Math.floor(endColFrac),
          row: Math.floor(endRowFrac),
          colOff: Math.round((endColFrac % 1) * COL_PX),
          rowOff: Math.round((endRowFrac % 1) * ROW_PX),
        },
      };
    }
    this.formControls.push(ctrl);
    return this;
  }
  getFormControls() {
    return this.formControls;
  }
  addShape(shape) {
    this.shapes.push(shape);
    return this;
  }
  getShapes() {
    return this.shapes;
  }
  addOleObject(obj) {
    this.oleObjects.push(obj);
    return this;
  }
  getOleObjects() {
    return this.oleObjects;
  }
  addWordArt(wa) {
    this.wordArt.push(wa);
    return this;
  }
  getWordArt() {
    return this.wordArt;
  }
  addQueryTable(qt) {
    this.queryTables.push(qt);
    return this;
  }
  getQueryTables() {
    return this.queryTables;
  }
  addTableSlicer(slicer) {
    this.tableSlicers.push(slicer);
    return this;
  }
  getTableSlicers() {
    return this.tableSlicers;
  }
  addMathEquation(eq) {
    this.mathEquations.push(eq);
    return this;
  }
  getMathEquations() {
    return this.mathEquations;
  }
  addIgnoredError(sqref, opts) {
    this.ignoreErrors.push({
      sqref,
      ...opts,
    });
    return this;
  }
  getIgnoredErrors() {
    return this.ignoreErrors;
  }
  addPreservedXml(xml) {
    this.preservedXml.push(xml);
    return this;
  }
  freeze(row, col) {
    this.freezePane = {
      row,
      col,
    };
    return this;
  }
  toXml(styles, shared) {
    const fitToPage = this.pageSetup?.fitToPage;
    const tabColor = this.options?.tabColor;
    const sheetPrXml =
      fitToPage || tabColor
        ? `<sheetPr>${tabColor ? `<tabColor rgb="${tabColor}"/>` : ""}${fitToPage ? '<pageSetUpPr fitToPage="1"/>' : ""}</sheetPr>`
        : "";
    const sheetViewXml = this._sheetViewXml();
    const colsXml = this._colsXml(styles);
    const sheetDataXml = this._sheetDataXml(styles, shared);
    const mergesXml = this._mergesXml();
    const cfXml = this._conditionalFormatXml(styles);
    const dvXml = this._dataValidationsXml();
    const autoFilterXml =
      this.autoFilter && !this.tables.some((t) => t.ref === this.autoFilter.ref)
        ? this._autoFilterXml()
        : "";
    const tablePartsXml = this.tables.length
      ? `<tableParts count="${this.tables.length}">${this.tableRIds.map((rId) => `<tablePart r:id="${rId}"/>`).join("")}</tableParts>`
      : "";
    const drawingXml = this.drawingRId ? `<drawing r:id="${this.drawingRId}"/>` : "";
    const legacyDrawingXml = this.legacyDrawingRId
      ? `<legacyDrawing r:id="${this.legacyDrawingRId}"/>`
      : "";
    const controlsXml = this._formControlsXml();
    const oleObjectsXml = this._oleObjectsXml();
    const sparklineXml = this._sparklineXml();
    const customIconExtXml = this._customIconExtXml();
    const slicerExtXml = this.slicerRId
      ? `<extLst><ext uri="{A8765BA9-456A-4dab-B4F3-ACF838C121DE}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><x14:slicerList><x14:slicer r:id="${this.slicerRId}"/></x14:slicerList></ext></extLst>`
      : "";
    const ignoredErrorsXml = this._ignoredErrorsXml();
    const protectionXml = this._protectionXml();
    const pageSetupXml = this._pageSetupXml();
    const pageMarginsXml = this._pageMarginsXml();
    const headerFooterXml = this._headerFooterXml();
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"
  xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision"
  xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
  xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing">
${sheetPrXml}
${sheetViewXml}
${colsXml}
${sheetDataXml}
${protectionXml}
${autoFilterXml}
${mergesXml}
${cfXml}
${dvXml}
${this._printOptionsXml()}
${pageMarginsXml}
${pageSetupXml}
${headerFooterXml}
${this._pageBreaksXml("rowBreaks", this.rowBreaks, 16383)}
${this._pageBreaksXml("colBreaks", this.colBreaks, 1048575)}
${ignoredErrorsXml}
${drawingXml}
${legacyDrawingXml}
${oleObjectsXml}
${controlsXml}
${sparklineXml}
${customIconExtXml}
${tablePartsXml}
${slicerExtXml}
${this.preservedXml.join("\n")}
</worksheet>`;
  }
  toChartSheetXml() {
    const pageMarginsXml = this._pageMarginsXml();
    const pageSetupXml = this._pageSetupXml();
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<chartsheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheetPr/>
<sheetViews><sheetView zoomScale="99" workbookViewId="0" zoomToFit="1"/></sheetViews>
${pageMarginsXml || '<pageMargins left="0.7" right="0.7" top="0.78740157499999996" bottom="0.78740157499999996" header="0.3" footer="0.3"/>'}
${pageSetupXml}
<drawing r:id="${this.drawingRId}"/>
</chartsheet>`;
  }
  toDialogSheetXml(_styles, _shared) {
    const pageMarginsXml = this._pageMarginsXml();
    const legacyDrawingXml = this.legacyDrawingRId
      ? `<legacyDrawing r:id="${this.legacyDrawingRId}"/>`
      : "";
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<dialogsheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheetViews><sheetView showRowColHeaders="0" showZeros="0" showOutlineSymbols="0" workbookViewId="0"/></sheetViews>
<sheetFormatPr baseColWidth="10" defaultColWidth="1" defaultRowHeight="5.65" customHeight="1"/>
<sheetProtection sheet="1"/>
${pageMarginsXml || '<pageMargins left="0.7" right="0.7" top="0.78740157499999996" bottom="0.78740157499999996" header="0.3" footer="0.3"/>'}
${legacyDrawingXml}
</dialogsheet>`;
  }
  _sheetViewXml() {
    const v = this.view ?? {};
    const attrs = [
      'workbookViewId="0"',
      v.showGridLines === false ? 'showGridLines="0"' : "",
      v.showRowColHeaders === false ? 'showRowColHeaders="0"' : "",
      v.zoomScale !== void 0 ? `zoomScale="${v.zoomScale}"` : "",
      v.rightToLeft ? 'rightToLeft="1"' : "",
      v.tabSelected ? 'tabSelected="1"' : "",
      v.view ? `view="${v.view}"` : "",
    ].filter(Boolean);
    let paneXml = "";
    if (this.freezePane) {
      const { row = 0, col = 0 } = this.freezePane;
      const topLeft =
        row && col
          ? indicesToCellRef(row + 1, col + 1)
          : row
            ? indicesToCellRef(row + 1, 1)
            : indicesToCellRef(1, col + 1);
      const activePane = row && col ? "bottomRight" : row ? "bottomLeft" : "topRight";
      paneXml = `<pane xSplit="${col}" ySplit="${row}" topLeftCell="${topLeft}" activePane="${activePane}" state="frozen"/>
<selection pane="${activePane}" activeCell="${topLeft}" sqref="${topLeft}"/>`;
    }
    return `<sheetViews><sheetView ${attrs.join(" ")}>${paneXml}</sheetView></sheetViews>`;
  }
  _colsXml(styles) {
    if (!this.colDefs.size) return "";
    return `<cols>${[...this.colDefs.entries()]
      .sort((a, b) => a[0] - b[0])
      .map(([idx, def]) => {
        const styleIdx = def.style ? styles.register(def.style) : 0;
        return (
          `<col min="${idx}" max="${idx}"` +
          (def.width ? ` width="${def.width}" customWidth="1"` : "") +
          (def.hidden ? ` hidden="1"` : "") +
          (def.bestFit ? ` bestFit="1"` : "") +
          (def.outlineLevel ? ` outlineLevel="${def.outlineLevel}"` : "") +
          (def.collapsed ? ` collapsed="1"` : "") +
          (styleIdx ? ` style="${styleIdx}"` : "") +
          "/>"
        );
      })
      .join("")}</cols>`;
  }
  _sheetDataXml(styles, shared) {
    const sortedRows = [...this.cells.entries()].sort((a, b) => a[0] - b[0]);
    const out = ["<sheetData>"];
    for (let ri = 0; ri < sortedRows.length; ri++) {
      const [rowIdx, colMap] = sortedRows[ri];
      const rowDef = this.rowDefs.get(rowIdx);
      const rowStyleIdx = rowDef?.style ? styles.register(rowDef.style) : 0;
      let attrs = `r="${rowIdx}"`;
      if (rowDef?.height) attrs += ` ht="${rowDef.height}" customHeight="1"`;
      if (rowDef?.hidden) attrs += ' hidden="1"';
      if (rowDef?.outlineLevel) attrs += ` outlineLevel="${rowDef.outlineLevel}"`;
      if (rowDef?.collapsed) attrs += ' collapsed="1"';
      if (rowStyleIdx) attrs += ` s="${rowStyleIdx}" customFormat="1"`;
      if (rowDef?.thickTop) attrs += ' thickTop="1"';
      if (rowDef?.thickBot) attrs += ' thickBot="1"';
      out.push(`<row ${attrs}>`);
      const sortedCells = [...colMap.entries()].sort((a, b) => a[0] - b[0]);
      for (let ci = 0; ci < sortedCells.length; ci++)
        out.push(this._cellXml(rowIdx, sortedCells[ci][0], sortedCells[ci][1], styles, shared));
      out.push("</row>");
    }
    out.push("</sheetData>");
    return out.join("");
  }
  _cellXml(row, col, cell, styles, shared) {
    const ref = `${colIndexToLetter(col)}${row}`;
    const styleIdx = cell.style ? styles.register(cell.style) : 0;
    const sAttr = styleIdx ? ` s="${styleIdx}"` : "";
    const vmIdx = this._cellImageVm.get(ref);
    const vmAttr = vmIdx !== void 0 ? ` vm="${vmIdx}"` : "";
    if (cell.arrayFormula)
      return `<c r="${ref}"${sAttr}${vmAttr}>${`<f t="array" ref="${ref}">${escapeXmlContent(cell.arrayFormula)}</f>`}<v>0</v></c>`;
    if (cell._sharedIdx !== void 0) {
      const si = cell._sharedIdx;
      const sharedRef = cell._sharedRef;
      if (cell.formula && sharedRef)
        return `<c r="${ref}"${sAttr}${vmAttr}>${`<f t="shared" ref="${sharedRef}" si="${si}">${escapeXmlContent(cell.formula)}</f>`}</c>`;
      return `<c r="${ref}"${sAttr}${vmAttr}><f t="shared" si="${si}"/></c>`;
    }
    if (cell.formula)
      return `<c r="${ref}"${sAttr}${vmAttr}>${`<f>${escapeXmlContent(cell.formula)}</f>`}</c>`;
    if (cell.richText)
      return `<c r="${ref}" t="s"${sAttr}${vmAttr}><v>${shared.internRichText(cell.richText)}</v></c>`;
    const v = cell.value;
    if (v === null || v === void 0) {
      if (vmAttr) return `<c r="${ref}"${sAttr} t="e"${vmAttr}><v>#VALUE!</v></c>`;
      return styleIdx ? `<c r="${ref}"${sAttr}/>` : "";
    }
    if (v instanceof CellError)
      return `<c r="${ref}" t="e"${sAttr}${vmAttr}><v>${escapeXml(v.error)}</v></c>`;
    if (typeof v === "boolean")
      return `<c r="${ref}" t="b"${sAttr}${vmAttr}><v>${v ? 1 : 0}</v></c>`;
    if (v instanceof Date) return `<c r="${ref}"${sAttr}${vmAttr}><v>${dateToSerial(v)}</v></c>`;
    if (typeof v === "number") return `<c r="${ref}"${sAttr}${vmAttr}><v>${v}</v></c>`;
    return `<c r="${ref}" t="s"${sAttr}${vmAttr}><v>${shared.intern(v)}</v></c>`;
  }
  _mergesXml() {
    if (!this.merges.length) return "";
    const items = this.merges.map((m) => {
      return `<mergeCell ref="${`${colIndexToLetter(m.startCol)}${m.startRow}`}:${`${colIndexToLetter(m.endCol)}${m.endRow}`}"/>`;
    });
    return `<mergeCells count="${this.merges.length}">${items.join("")}</mergeCells>`;
  }
  _conditionalFormatXml(styles) {
    return this.conditionalFormats
      .map((cf) => {
        const dxfId = cf.style ? styles.registerDxf(cf.style) : void 0;
        let inner = "";
        if (cf.colorScale?.type === "colorScale") {
          const cs = cf.colorScale;
          inner = `<colorScale>${cs.cfvo.map((v) => `<cfvo type="${v.type}"${v.val ? ` val="${v.val}"` : ""}/>`).join("")}${cs.color.map((c) => colorEl(c)).join("")}</colorScale>`;
        } else if (cf.dataBar?.type === "dataBar") {
          const db = cf.dataBar;
          const minCfvo = `<cfvo type="${db.minType ?? "min"}"${db.minVal != null ? ` val="${db.minVal}"` : ""}/>`;
          const maxCfvo = `<cfvo type="${db.maxType ?? "max"}"${db.maxVal != null ? ` val="${db.maxVal}"` : ""}/>`;
          const color = db.color ?? db.minColor ?? "FF638EC6";
          inner = `<dataBar${db.showValue === false ? ' showValue="0"' : ""}>${minCfvo}${maxCfvo}${colorEl(color)}</dataBar>`;
        } else if (cf.iconSet?.type === "iconSet") {
          const is = cf.iconSet;
          const cfvos = is.cfvo
            .map((v) => `<cfvo type="${v.type}"${v.val ? ` val="${v.val}"` : ""}/>`)
            .join("");
          inner = `<iconSet iconSet="${is.iconSet}"${is.showValue === false ? ' showValue="0"' : ""}${is.reverse ? ' reverse="1"' : ""}>${cfvos}</iconSet>`;
        }
        const ruleAttrs = [
          `type="${cf.type}"`,
          cf.operator ? `operator="${cf.operator}"` : "",
          dxfId !== void 0 ? `dxfId="${dxfId}"` : "",
          `priority="${cf.priority ?? 1}"`,
          cf.aboveAverage === false ? `aboveAverage="0"` : "",
          cf.percent ? `percent="1"` : "",
          cf.rank ? `rank="${cf.rank}"` : "",
          cf.timePeriod ? `timePeriod="${cf.timePeriod}"` : "",
          cf.text ? `text="${escapeXml(cf.text)}"` : "",
        ]
          .filter(Boolean)
          .join(" ");
        const fml1 = cf.formula ? `<formula>${escapeXml(cf.formula)}</formula>` : "";
        const fml2 = cf.formula2 ? `<formula>${escapeXml(cf.formula2)}</formula>` : "";
        return `<conditionalFormatting sqref="${cf.sqref}"><cfRule ${ruleAttrs}>${fml1}${fml2}${inner}</cfRule></conditionalFormatting>`;
      })
      .join("");
  }
  _dataValidationsXml() {
    if (!this.dataValidations.size) return "";
    const items = [...this.dataValidations.entries()].map(([sqref, dv]) => {
      const formula1 =
        dv.type === "list" && dv.list
          ? `<formula1>"${dv.list.join(",")}"</formula1>`
          : dv.formula1
            ? `<formula1>${escapeXml(dv.formula1)}</formula1>`
            : "";
      const formula2 = dv.formula2 ? `<formula2>${escapeXml(dv.formula2)}</formula2>` : "";
      return `<dataValidation ${[
        `type="${dv.type}"`,
        dv.operator ? `operator="${dv.operator}"` : "",
        `sqref="${sqref}"`,
        dv.showDropDown !== false && dv.type === "list" ? "" : 'showDropDown="1"',
        dv.allowBlank !== false ? 'allowBlank="1"' : "",
        dv.showErrorAlert ? 'showErrorMessage="1"' : "",
        dv.errorTitle ? `errorTitle="${escapeXml(dv.errorTitle)}"` : "",
        dv.error ? `error="${escapeXml(dv.error)}"` : "",
        dv.showInputMessage ? 'showInputMessage="1"' : "",
        dv.promptTitle ? `promptTitle="${escapeXml(dv.promptTitle)}"` : "",
        dv.prompt ? `prompt="${escapeXml(dv.prompt)}"` : "",
      ]
        .filter(Boolean)
        .join(" ")}>${formula1}${formula2}</dataValidation>`;
    });
    return `<dataValidations count="${this.dataValidations.size}">${items.join("")}</dataValidations>`;
  }
  _protectionXml() {
    const p = this.protection;
    if (!p) return "";
    return `<sheetProtection ${[
      p.sheet ? 'sheet="1"' : "",
      p.password ? `password="${hashPassword(p.password)}"` : "",
      p.selectLockedCells === false ? 'selectLockedCells="0"' : "",
      p.selectUnlockedCells === false ? 'selectUnlockedCells="0"' : "",
      p.formatCells ? 'formatCells="0"' : "",
      p.formatColumns ? 'formatColumns="0"' : "",
      p.formatRows ? 'formatRows="0"' : "",
      p.insertColumns ? 'insertColumns="0"' : "",
      p.insertRows ? 'insertRows="0"' : "",
      p.insertHyperlinks ? 'insertHyperlinks="0"' : "",
      p.deleteColumns ? 'deleteColumns="0"' : "",
      p.deleteRows ? 'deleteRows="0"' : "",
      p.sort ? 'sort="0"' : "",
      p.autoFilter ? 'autoFilter="0"' : "",
      p.pivotTables ? 'pivotTables="0"' : "",
    ]
      .filter(Boolean)
      .join(" ")}/>`;
  }
  _pageSetupXml() {
    const p = this.pageSetup;
    if (!p) return "";
    return `<pageSetup ${[
      p.paperSize ? `paperSize="${p.paperSize}"` : "",
      p.orientation ? `orientation="${p.orientation}"` : "",
      p.fitToWidth !== void 0 ? `fitToWidth="${p.fitToWidth}"` : "",
      p.fitToHeight !== void 0 ? `fitToHeight="${p.fitToHeight}"` : "",
      p.scale ? `scale="${p.scale}"` : "",
      p.horizontalDpi ? `horizontalDpi="${p.horizontalDpi}"` : "",
      p.verticalDpi ? `verticalDpi="${p.verticalDpi}"` : "",
      p.firstPageNumber ? `firstPageNumber="${p.firstPageNumber}" useFirstPageNumber="1"` : "",
    ]
      .filter(Boolean)
      .join(" ")}/>`;
  }
  _pageMarginsXml() {
    const d = {
      left: 0.7,
      right: 0.7,
      top: 0.75,
      bottom: 0.75,
      header: 0.3,
      footer: 0.3,
      ...this.pageMargins,
    };
    return `<pageMargins left="${d.left}" right="${d.right}" top="${d.top}" bottom="${d.bottom}" header="${d.header}" footer="${d.footer}"/>`;
  }
  _headerFooterXml() {
    const hf = this.headerFooter;
    if (!hf) return "";
    const attrs = [
      hf.differentOddEven ? 'differentOddEven="1"' : "",
      hf.differentFirst ? 'differentFirst="1"' : "",
    ]
      .filter(Boolean)
      .join(" ");
    const oh = hf.oddHeader ? `<oddHeader>${escapeXml(hf.oddHeader)}</oddHeader>` : "";
    const of_ = hf.oddFooter ? `<oddFooter>${escapeXml(hf.oddFooter)}</oddFooter>` : "";
    const eh = hf.evenHeader ? `<evenHeader>${escapeXml(hf.evenHeader)}</evenHeader>` : "";
    const ef = hf.evenFooter ? `<evenFooter>${escapeXml(hf.evenFooter)}</evenFooter>` : "";
    const fh = hf.firstHeader ? `<firstHeader>${escapeXml(hf.firstHeader)}</firstHeader>` : "";
    const ff = hf.firstFooter ? `<firstFooter>${escapeXml(hf.firstFooter)}</firstFooter>` : "";
    return `<headerFooter${attrs ? " " + attrs : ""}>${oh}${of_}${eh}${ef}${fh}${ff}</headerFooter>`;
  }
  _printOptionsXml() {
    const p = this.printOptions;
    if (!p) return "";
    const attrs = [
      p.gridLines ? 'gridLines="1"' : "",
      p.gridLinesSet ? 'gridLinesSet="1"' : "",
      p.headings ? 'headings="1"' : "",
      p.centerHorizontal ? 'horizontalCentered="1"' : "",
      p.centerVertical ? 'verticalCentered="1"' : "",
    ]
      .filter(Boolean)
      .join(" ");
    return attrs ? `<printOptions ${attrs}/>` : "";
  }
  _pageBreaksXml(tag, breaks, maxVal) {
    if (!breaks.length) return "";
    const manualCount = breaks.filter((b) => b.manual !== false).length;
    const brks = breaks
      .map((b) => `<brk id="${b.id}" max="${maxVal}"${b.manual !== false ? ' man="1"' : ""}/>`)
      .join("");
    return `<${tag} count="${breaks.length}" manualBreakCount="${manualCount}">${brks}</${tag}>`;
  }
  _formControlsXml() {
    if (!this.formControls.length || !this.ctrlPropRIds.length) return "";
    const baseShapeId = 1025 + this.sheetIndex * 1e3;
    let commentCount = 0;
    for (const rowMap of this.cells.values())
      for (const c of rowMap.values()) if (c.comment) commentCount++;
    return `<mc:AlternateContent><mc:Choice Requires="x14"><controls>${this.formControls
      .map((ctrl, i) => {
        const shapeId = ctrl._shapeId ?? baseShapeId + commentCount + i;
        const ctrlPropRId = this.ctrlPropRIds[i];
        if (!ctrlPropRId) return "";
        return `<mc:AlternateContent><mc:Choice Requires="x14"><control shapeId="${shapeId}" r:id="${ctrlPropRId}" name="${escapeXml(ctrl.text ?? `${ctrl.type} ${i + 1}`)}"><controlPr defaultSize="0" print="0" autoFill="0" autoPict="0"${ctrl.macro ? ` macro="${escapeXml(ctrl.macro)}"` : ""}><anchor moveWithCells="1"><from><xdr:col>${ctrl.from.col}</xdr:col><xdr:colOff>${ctrl.from.colOff ?? 0}</xdr:colOff><xdr:row>${ctrl.from.row}</xdr:row><xdr:rowOff>${ctrl.from.rowOff ?? 0}</xdr:rowOff></from><to><xdr:col>${ctrl.to.col}</xdr:col><xdr:colOff>${ctrl.to.colOff ?? 0}</xdr:colOff><xdr:row>${ctrl.to.row}</xdr:row><xdr:rowOff>${ctrl.to.rowOff ?? 0}</xdr:rowOff></to></anchor></controlPr></control></mc:Choice></mc:AlternateContent>`;
      })
      .join("")}</controls></mc:Choice></mc:AlternateContent>`;
  }
  _oleObjectsXml() {
    if (!this.oleObjects.length || !this.oleRIds.length) return "";
    const baseShapeId = 2025 + this.sheetIndex * 1e3;
    return `<oleObjects>${this.oleObjects
      .map((ole, i) => {
        const shapeId = baseShapeId + i;
        const rId = this.oleRIds[i];
        const progId = ole.progId ?? "Package";
        const link = ole.linkToFile && ole.linkPath ? ` link="${escapeXml(ole.linkPath)}"` : "";
        return `<oleObject progId="${escapeXml(progId)}" shapeId="${shapeId}" r:id="${rId}"${link}/>`;
      })
      .join("")}</oleObjects>`;
  }
  _sparklineXml() {
    if (!this.sparklines.length) return "";
    const colorEl = (name, c) => {
      if (!c) return "";
      return `<x14:${name} rgb="${c.startsWith("#") ? "FF" + c.slice(1) : c}"/>`;
    };
    return `<extLst><ext uri="{05C60535-1F16-4fd2-B633-F4F36F0B64E0}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">${`<x14:sparklineGroups xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">${this.sparklines
      .map((s) => {
        return `<x14:sparklineGroup ${[
          `type="${s.type === "bar" ? "column" : s.type}"`,
          s.lineWidth !== void 0 ? `lineWeight="${s.lineWidth}"` : "",
          s.showMarkers ? 'markers="1"' : "",
          s.showFirst ? 'first="1"' : "",
          s.showLast ? 'last="1"' : "",
          s.showHigh ? 'high="1"' : "",
          s.showLow ? 'low="1"' : "",
          s.showNegative ? 'negative="1"' : "",
          s.minAxisType ? `minAxisType="${s.minAxisType}"` : "",
          s.maxAxisType ? `maxAxisType="${s.maxAxisType}"` : "",
        ]
          .filter(Boolean)
          .join(" ")}>${[
          colorEl("colorSeries", s.color),
          colorEl("colorHigh", s.highColor),
          colorEl("colorLow", s.lowColor),
          colorEl("colorFirst", s.firstColor),
          colorEl("colorLast", s.lastColor),
          colorEl("colorNegative", s.negativeColor),
          colorEl("colorMarkers", s.markersColor),
        ].join(
          "",
        )}${`<x14:sparklines><x14:sparkline><xm:f>${escapeXml(s.dataRange.includes("!") ? s.dataRange : `${this.name}!${s.dataRange}`)}</xm:f><xm:sqref>${s.location}</xm:sqref></x14:sparkline></x14:sparklines>`}</x14:sparklineGroup>`;
      })
      .join("")}</x14:sparklineGroups>`}</ext></extLst>`;
  }
  _customIconExtXml() {
    const customCFs = this.conditionalFormats.filter(
      (cf) => cf.iconSet?.type === "iconSet" && "custom" in cf.iconSet && cf.iconSet.custom?.length,
    );
    if (!customCFs.length) return "";
    return `<extLst><ext uri="{78C0D931-6437-407d-A8EE-F0AAD7539E65}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">${customCFs
      .map((cf, i) => {
        const is = cf.iconSet;
        const cfvos = is.cfvo
          .map((v) => `<x14:cfvo type="${v.type}"${v.val ? ` val="${v.val}"` : ""}/>`)
          .join("");
        const icons = is.custom
          .map((ci) => `<x14:cfIcon iconSet="${ci.iconSet}" iconId="${ci.iconId}"/>`)
          .join("");
        return `<x14:conditionalFormatting xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main"><x14:cfRule type="iconSet" id="{${this._uuid()}}"><x14:iconSet iconSet="${is.iconSet}" custom="1"${is.showValue === false ? ' showValue="0"' : ""}${is.reverse ? ' reverse="1"' : ""}>${cfvos}${icons}</x14:iconSet></x14:cfRule><xm:sqref>${cf.sqref}</xm:sqref></x14:conditionalFormatting>`;
      })
      .join("")}</ext></extLst>`;
  }
  _uuid() {
    const h = "0123456789ABCDEF";
    let u = "";
    for (let i = 0; i < 36; i++)
      if (i === 8 || i === 13 || i === 18 || i === 23) u += "-";
      else if (i === 14) u += "4";
      else if (i === 19) u += h[((Math.random() * 4) | 0) + 8];
      else u += h[(Math.random() * 16) | 0];
    return u;
  }
  _autoFilterXml() {
    if (!this.autoFilter) return "";
    const cols = this._filterColumns;
    if (!cols || cols.length === 0) return `<autoFilter ref="${this.autoFilter.ref}"/>`;
    const colXml = cols
      .map((fc) => {
        const colId = fc.col - 1;
        let inner = "";
        if (fc.type === "custom")
          inner = `<customFilters><customFilter operator="${fc.operator ?? "greaterThan"}" val="${fc.val ?? ""}"/></customFilters>`;
        else if (fc.type === "top10")
          inner = `<top10 ${[
            fc.top === false ? 'top="0"' : "",
            fc.percent ? 'percent="1"' : "",
            `val="${fc.val ?? 10}"`,
          ]
            .filter(Boolean)
            .join(" ")}/>`;
        else if (fc.type === "value" && fc.items)
          inner = `<filters>${fc.items.map((v) => `<filter val="${escapeXml(String(v))}"/>`).join("")}</filters>`;
        else if (fc.type === "dynamic")
          inner = `<dynamicFilter type="${fc.dynamicType ?? "aboveAverage"}"/>`;
        return `<filterColumn colId="${colId}">${inner}</filterColumn>`;
      })
      .join("");
    return `<autoFilter ref="${this.autoFilter.ref}">${colXml}</autoFilter>`;
  }
  _ignoredErrorsXml() {
    if (!this.ignoreErrors.length) return "";
    return `<ignoredErrors>${this.ignoreErrors
      .map((ie) => {
        const attrs = [`sqref="${ie.sqref}"`];
        if (ie.numberStoredAsText) attrs.push('numberStoredAsText="1"');
        if (ie.formula) attrs.push('formula="1"');
        if (ie.formulaRange) attrs.push('formulaRange="1"');
        if (ie.unlockedFormula) attrs.push('unlockedFormula="1"');
        if (ie.evalError) attrs.push('evalError="1"');
        if (ie.twoDigitTextYear) attrs.push('twoDigitTextYear="1"');
        if (ie.emptyRef) attrs.push('emptyRef="1"');
        if (ie.listDataValidation) attrs.push('listDataValidation="1"');
        if (ie.calculatedColumn) attrs.push('calculatedColumn="1"');
        return `<ignoredError ${attrs.join(" ")}/>`;
      })
      .join("")}</ignoredErrors>`;
  }
  toDrawingXml(imageRIds, chartRIds) {
    const parts = [];
    const EMU = pxToEmu;
    this.images.forEach((img, i) => {
      const rId = imageRIds[i];
      const w = EMU(img.width ?? 100);
      const h = EMU(img.height ?? 100);
      let anchor;
      let closeTag;
      if (img.position) {
        anchor = `<xdr:absoluteAnchor><xdr:pos x="${EMU(img.position.x)}" y="${EMU(img.position.y)}"/><xdr:ext cx="${w}" cy="${h}"/>`;
        closeTag = `</xdr:absoluteAnchor>`;
      } else {
        const from = img.from;
        const to = img.to;
        const fromXml = `<xdr:from><xdr:col>${from.col}</xdr:col><xdr:colOff>${from.colOff ?? 0}</xdr:colOff><xdr:row>${from.row}</xdr:row><xdr:rowOff>${from.rowOff ?? 0}</xdr:rowOff></xdr:from>`;
        if (to) {
          anchor = `<xdr:twoCellAnchor editAs="oneCell">${fromXml}${`<xdr:to><xdr:col>${to.col}</xdr:col><xdr:colOff>${to.colOff ?? 0}</xdr:colOff><xdr:row>${to.row}</xdr:row><xdr:rowOff>${to.rowOff ?? 0}</xdr:rowOff></xdr:to>`}`;
          closeTag = `</xdr:twoCellAnchor>`;
        } else {
          anchor = `<xdr:oneCellAnchor>${fromXml}<xdr:ext cx="${w}" cy="${h}"/>`;
          closeTag = `</xdr:oneCellAnchor>`;
        }
      }
      const picXml = `<xdr:pic>
  <xdr:nvPicPr>
    <xdr:cNvPr id="${i + 2}" name="Image ${i + 1}"${img.altText ? ` descr="${escapeXml(img.altText)}"` : ""}/>
    <xdr:cNvPicPr><a:picLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" noChangeAspect="1"/></xdr:cNvPicPr>
  </xdr:nvPicPr>
  <xdr:blipFill>
    <a:blip xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" r:embed="${rId}"/>
    <a:stretch xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:fillRect/></a:stretch>
  </xdr:blipFill>
  <xdr:spPr>
    <a:xfrm xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
      <a:off x="0" y="0"/><a:ext cx="${w}" cy="${h}"/>
    </a:xfrm>
    <a:prstGeom xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" prst="rect"><a:avLst/></a:prstGeom>
  </xdr:spPr>
</xdr:pic>`;
      parts.push(`${anchor}${picXml}<xdr:clientData/>${closeTag}`);
    });
    this.charts.forEach((chart, i) => {
      const rId = chartRIds[i];
      const from = chart.from;
      const to = chart.to;
      const graphicXml = `<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
    <c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="${rId}"/>
  </a:graphicData>
</a:graphic>`;
      if (this._isChartSheet)
        parts.push(
          `<xdr:absoluteAnchor><xdr:pos x="0" y="0"/><xdr:ext cx="9294091" cy="6003636"/><xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="${this.images.length + i + 2}" name="Chart ${i + 1}"/><xdr:cNvGraphicFramePr><a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" noGrp="1"/></xdr:cNvGraphicFramePr></xdr:nvGraphicFramePr><xdr:xfrm><a:off xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" x="0" y="0"/><a:ext xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" cx="0" cy="0"/></xdr:xfrm>${graphicXml}</xdr:graphicFrame><xdr:clientData/></xdr:absoluteAnchor>`,
        );
      else {
        const fromXml = `<xdr:from><xdr:col>${from.col}</xdr:col><xdr:colOff>${from.colOff ?? 0}</xdr:colOff><xdr:row>${from.row}</xdr:row><xdr:rowOff>${from.rowOff ?? 0}</xdr:rowOff></xdr:from>`;
        const toXml = `<xdr:to><xdr:col>${to.col}</xdr:col><xdr:colOff>${to.colOff ?? 0}</xdr:colOff><xdr:row>${to.row}</xdr:row><xdr:rowOff>${to.rowOff ?? 0}</xdr:rowOff></xdr:to>`;
        parts.push(
          `<xdr:twoCellAnchor editAs="oneCell">${fromXml}${toXml}<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="${this.images.length + i + 2}" name="Chart ${i + 1}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" x="0" y="0"/><a:ext xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" cx="0" cy="0"/></xdr:xfrm>${graphicXml}</xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor>`,
        );
      }
    });
    let shapeCounter = this.images.length + this.charts.length + 2;
    const toHex6 = (c) => {
      let h = c.replace(/^#/, "");
      if (h.length === 8) h = h.substring(2);
      return h;
    };
    this.shapes.forEach((shape) => {
      const id = shapeCounter++;
      const from = shape.from;
      const to = shape.to;
      const fromXml = `<xdr:from><xdr:col>${from.col}</xdr:col><xdr:colOff>${from.colOff ?? 0}</xdr:colOff><xdr:row>${from.row}</xdr:row><xdr:rowOff>${from.rowOff ?? 0}</xdr:rowOff></xdr:from>`;
      const toXml = `<xdr:to><xdr:col>${to.col}</xdr:col><xdr:colOff>${to.colOff ?? 0}</xdr:colOff><xdr:row>${to.row}</xdr:row><xdr:rowOff>${to.rowOff ?? 0}</xdr:rowOff></xdr:to>`;
      const fillXml = shape.fillColor
        ? `<a:solidFill><a:srgbClr val="${toHex6(shape.fillColor)}"/></a:solidFill>`
        : "<a:noFill/>";
      const lineXml = shape.lineColor
        ? `<a:ln${shape.lineWidth ? ` w="${shape.lineWidth * 12700}"` : ""}><a:solidFill><a:srgbClr val="${toHex6(shape.lineColor)}"/></a:solidFill></a:ln>`
        : "";
      const rotAttr = shape.rotation ? ` rot="${shape.rotation * 6e4}"` : "";
      const textXml = shape.text
        ? `<xdr:txBody><a:bodyPr vertOverflow="clip" wrap="square" rtlCol="0" anchor="ctr"/><a:lstStyle/><a:p><a:pPr algn="ctr"/><a:r><a:rPr lang="en-US"${shape.font?.bold ? ' b="1"' : ""}${shape.font?.size ? ` sz="${shape.font.size * 100}"` : ""}/><a:t>${escapeXml(shape.text)}</a:t></a:r></a:p></xdr:txBody>`
        : "";
      parts.push(
        `<xdr:twoCellAnchor editAs="oneCell">${fromXml}${toXml}<xdr:sp><xdr:nvSpPr><xdr:cNvPr id="${id}" name="Shape ${id}"/><xdr:cNvSpPr/></xdr:nvSpPr><xdr:spPr><a:xfrm${rotAttr}><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></a:xfrm><a:prstGeom prst="${shape.type}"><a:avLst/></a:prstGeom>${fillXml}${lineXml}</xdr:spPr>${textXml}</xdr:sp><xdr:clientData/></xdr:twoCellAnchor>`,
      );
    });
    this.wordArt.forEach((wa) => {
      const id = shapeCounter++;
      const from = wa.from;
      const to = wa.to;
      const fromXml = `<xdr:from><xdr:col>${from.col}</xdr:col><xdr:colOff>${from.colOff ?? 0}</xdr:colOff><xdr:row>${from.row}</xdr:row><xdr:rowOff>${from.rowOff ?? 0}</xdr:rowOff></xdr:from>`;
      const toXml = `<xdr:to><xdr:col>${to.col}</xdr:col><xdr:colOff>${to.colOff ?? 0}</xdr:colOff><xdr:row>${to.row}</xdr:row><xdr:rowOff>${to.rowOff ?? 0}</xdr:rowOff></xdr:to>`;
      const preset = wa.preset ?? "textPlain";
      const fontSize = wa.font?.size ? wa.font.size * 100 : 3600;
      const fillHex = wa.fillColor ? toHex6(wa.fillColor) : "";
      const outlineHex = wa.outlineColor ? toHex6(wa.outlineColor) : "";
      const textFillXml = fillHex
        ? `<a:solidFill><a:srgbClr val="${fillHex}"/></a:solidFill>`
        : '<a:solidFill><a:schemeClr val="tx1"/></a:solidFill>';
      const textOutlineXml = outlineHex
        ? `<a:ln w="12700"><a:solidFill><a:srgbClr val="${outlineHex}"/></a:solidFill><a:prstDash val="solid"/></a:ln>`
        : "";
      const effectXml = `<a:effectLst><a:outerShdw dist="38100" dir="2700000" algn="bl" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="40000"/></a:srgbClr></a:outerShdw></a:effectLst>`;
      const fontAttrs = [
        'lang="en-US"',
        `sz="${fontSize}"`,
        wa.font?.bold ? 'b="1"' : 'b="1"',
        wa.font?.italic ? 'i="1"' : "",
        'cap="none" spc="0"',
      ]
        .filter(Boolean)
        .join(" ");
      const fontFace = wa.font?.name ?? "Calibri";
      const prstWarpXml =
        preset !== "textPlain" ? `<a:prstTxWarp prst="${preset}"><a:avLst/></a:prstTxWarp>` : "";
      parts.push(
        `<xdr:twoCellAnchor editAs="oneCell">${fromXml}${toXml}<xdr:sp macro="" textlink=""><xdr:nvSpPr><xdr:cNvPr id="${id}" name="WordArt ${id}"/><xdr:cNvSpPr/></xdr:nvSpPr><xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></xdr:spPr><xdr:txBody><a:bodyPr wrap="square" lIns="91440" tIns="45720" rIns="91440" bIns="45720">${prstWarpXml}<a:spAutoFit/></a:bodyPr><a:lstStyle/><a:p><a:pPr algn="ctr"/><a:r><a:rPr ${fontAttrs}>${textOutlineXml}${textFillXml}${effectXml}<a:latin typeface="${escapeXml(fontFace)}"/></a:rPr><a:t>${escapeXml(wa.text)}</a:t></a:r></a:p></xdr:txBody></xdr:sp><xdr:clientData/></xdr:twoCellAnchor>`,
      );
    });
    this.mathEquations.forEach((eq) => {
      const id = shapeCounter++;
      const from = eq.from;
      const fromXml = `<xdr:from><xdr:col>${from.col}</xdr:col><xdr:colOff>${from.colOff ?? 0}</xdr:colOff><xdr:row>${from.row}</xdr:row><xdr:rowOff>${from.rowOff ?? 0}</xdr:rowOff></xdr:from>`;
      const w = eq.width ?? 18e5;
      const h = eq.height ?? 5e5;
      const extXml = `<xdr:ext cx="${w}" cy="${h}"/>`;
      const fontSize = eq.fontSize ? eq.fontSize * 100 : 1100;
      const fontName = eq.fontName ?? "Cambria Math";
      const rPr = `<a:rPr lang="en-US" sz="${fontSize}" i="1"><a:latin typeface="${escapeXml(fontName)}" panose="02040503050406030204" pitchFamily="18" charset="0"/></a:rPr>`;
      const ctrlPr = `<m:ctrlPr>${rPr}</m:ctrlPr>`;
      const buildOmml = (el) => {
        switch (el.type) {
          case "text":
            return `<m:r>${rPr}<m:t>${escapeXml(el.text ?? "")}</m:t></m:r>`;
          case "frac":
            return `<m:f><m:fPr>${el.hideDegree ? '<m:type m:val="noBar"/>' : ""}${ctrlPr}</m:fPr><m:num>${(el.base ?? []).map(buildOmml).join("")}</m:num><m:den>${(el.argument ?? []).map(buildOmml).join("")}</m:den></m:f>`;
          case "sup":
            return `<m:sSup><m:sSupPr>${ctrlPr}</m:sSupPr><m:e>${(el.base ?? []).map(buildOmml).join("")}</m:e><m:sup>${(el.argument ?? []).map(buildOmml).join("")}</m:sup></m:sSup>`;
          case "sub":
            return `<m:sSub><m:sSubPr>${ctrlPr}</m:sSubPr><m:e>${(el.base ?? []).map(buildOmml).join("")}</m:e><m:sub>${(el.argument ?? []).map(buildOmml).join("")}</m:sub></m:sSub>`;
          case "subSup":
            return `<m:sSubSup><m:sSubSupPr>${ctrlPr}</m:sSubSupPr><m:e>${(el.base ?? []).map(buildOmml).join("")}</m:e><m:sub>${(el.subscript ?? []).map(buildOmml).join("")}</m:sub><m:sup>${(el.superscript ?? []).map(buildOmml).join("")}</m:sup></m:sSubSup>`;
          case "nary":
            return `<m:nary><m:naryPr><m:chr m:val="${escapeXml(el.operator ?? "∑")}"/>${ctrlPr}</m:naryPr><m:sub>${(el.lower ?? []).map(buildOmml).join("")}</m:sub><m:sup>${(el.upper ?? []).map(buildOmml).join("")}</m:sup><m:e>${(el.body ?? []).map(buildOmml).join("")}</m:e></m:nary>`;
          case "rad":
            return `<m:rad><m:radPr>${el.hideDegree ? '<m:degHide m:val="1"/>' : ""}${ctrlPr}</m:radPr><m:deg>${(el.degree ?? []).map(buildOmml).join("")}</m:deg><m:e>${(el.body ?? []).map(buildOmml).join("")}</m:e></m:rad>`;
          case "delim": {
            const open = el.open ?? "(";
            const close = el.close ?? ")";
            return `<m:d><m:dPr>${open !== "(" ? `<m:begChr m:val="${escapeXml(open)}"/>` : ""}${close !== ")" ? `<m:endChr m:val="${escapeXml(close)}"/>` : ""}${ctrlPr}</m:dPr><m:e>${(el.body ?? []).map(buildOmml).join("")}</m:e></m:d>`;
          }
          case "func":
            return `<m:func><m:funcPr>${ctrlPr}</m:funcPr><m:fName>${(el.base ?? []).map(buildOmml).join("")}</m:fName><m:e>${(el.argument ?? []).map(buildOmml).join("")}</m:e></m:func>`;
          case "accent":
            return `<m:acc><m:accPr>${el.operator ? `<m:chr m:val="${escapeXml(el.operator)}"/>` : ""}${ctrlPr}</m:accPr><m:e>${(el.body ?? []).map(buildOmml).join("")}</m:e></m:acc>`;
          case "bar":
            return `<m:bar><m:barPr>${ctrlPr}</m:barPr><m:e>${(el.body ?? []).map(buildOmml).join("")}</m:e></m:bar>`;
          case "matrix":
            return `<m:m><m:mPr>${ctrlPr}</m:mPr>${(el.rows ?? []).map((row) => `<m:mr>${row.map((c) => `<m:e>${buildOmml(c)}</m:e>`).join("")}</m:mr>`).join("")}</m:m>`;
          case "eqArr":
            return `<m:eqArr><m:eqArrPr>${ctrlPr}</m:eqArrPr>${(el.rows ?? []).map((row) => `<m:e>${row.map(buildOmml).join("")}</m:e>`).join("")}</m:eqArr>`;
          case "limLow":
            return `<m:limLow><m:limLowPr>${ctrlPr}</m:limLowPr><m:e>${(el.base ?? []).map(buildOmml).join("")}</m:e><m:lim>${(el.argument ?? []).map(buildOmml).join("")}</m:lim></m:limLow>`;
          case "limUpp":
            return `<m:limUpp><m:limUppPr>${ctrlPr}</m:limUppPr><m:e>${(el.base ?? []).map(buildOmml).join("")}</m:e><m:lim>${(el.argument ?? []).map(buildOmml).join("")}</m:lim></m:limUpp>`;
          case "groupChar":
            return `<m:groupChr><m:groupChrPr>${el.operator ? `<m:chr m:val="${escapeXml(el.operator)}"/>` : ""}${ctrlPr}</m:groupChrPr><m:e>${(el.body ?? []).map(buildOmml).join("")}</m:e></m:groupChr>`;
          default:
            return `<m:r>${rPr}<m:t>${escapeXml(el.text ?? "")}</m:t></m:r>`;
        }
      };
      const ommlXml = `<m:oMathPara xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><m:oMathParaPr><m:jc m:val="centerGroup"/></m:oMathParaPr><m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">${eq.elements.map(buildOmml).join("")}</m:oMath></m:oMathPara>`;
      const flattenText = (el) => {
        if (el.text) return el.text;
        const parts = [
          el.base,
          el.argument,
          el.body,
          el.lower,
          el.upper,
          el.superscript,
          el.subscript,
          el.degree,
        ]
          .filter(Boolean)
          .map((arr) => arr.map(flattenText).join(""))
          .join("");
        if (el.type === "frac")
          return `(${(el.base ?? []).map(flattenText).join("")})/(${(el.argument ?? []).map(flattenText).join("")})`;
        if (el.type === "sup")
          return `${(el.base ?? []).map(flattenText).join("")}^${(el.argument ?? []).map(flattenText).join("")}`;
        if (el.type === "nary")
          return `${el.operator ?? "∑"}(${(el.body ?? []).map(flattenText).join("")})`;
        if (el.type === "delim")
          return `${el.open ?? "("}${(el.body ?? []).map(flattenText).join("")}${el.close ?? ")"}`;
        if (el.type === "rad") return `√(${(el.body ?? []).map(flattenText).join("")})`;
        return parts;
      };
      const fallbackText = eq.elements.map(flattenText).join("");
      parts.push(
        `<xdr:oneCellAnchor>${fromXml}${extXml}<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main" Requires="a14"><xdr:sp macro="" textlink=""><xdr:nvSpPr><xdr:cNvPr id="${id}" name="MathEq ${id}"/><xdr:cNvSpPr txBox="1"/></xdr:nvSpPr><xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${w}" cy="${h}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></xdr:spPr><xdr:txBody><a:bodyPr vertOverflow="clip" horzOverflow="clip" wrap="none" lIns="0" tIns="0" rIns="0" bIns="0" rtlCol="0" anchor="t"><a:spAutoFit/></a:bodyPr><a:lstStyle/><a:p><a14:m>${ommlXml}</a14:m><a:endParaRPr lang="en-US" sz="${fontSize}"/></a:p></xdr:txBody></xdr:sp></mc:Choice><mc:Fallback><xdr:sp macro="" textlink=""><xdr:nvSpPr><xdr:cNvPr id="${id}" name="MathEq ${id}"/><xdr:cNvSpPr txBox="1"/></xdr:nvSpPr><xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${w}" cy="${h}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></xdr:spPr><xdr:txBody><a:bodyPr vertOverflow="clip" horzOverflow="clip" wrap="none" lIns="0" tIns="0" rIns="0" bIns="0" rtlCol="0" anchor="t"><a:spAutoFit/></a:bodyPr><a:lstStyle/><a:p><a:r><a:rPr lang="en-US" sz="${fontSize}" i="0"><a:latin typeface="${escapeXml(fontName)}" panose="02040503050406030204" pitchFamily="18" charset="0"/></a:rPr><a:t>${escapeXml(fallbackText)}</a:t></a:r><a:endParaRPr lang="en-US" sz="${fontSize}"/></a:p></xdr:txBody></xdr:sp></mc:Fallback></mc:AlternateContent><xdr:clientData/></xdr:oneCellAnchor>`,
      );
    });
    this.oleObjects.forEach((ole, i) => {
      const id = shapeCounter++;
      const from = ole.from;
      const to = ole.to;
      const fromXml = `<xdr:from><xdr:col>${from.col}</xdr:col><xdr:colOff>${from.colOff ?? 0}</xdr:colOff><xdr:row>${from.row}</xdr:row><xdr:rowOff>${from.rowOff ?? 0}</xdr:rowOff></xdr:from>`;
      const toXml = `<xdr:to><xdr:col>${to.col}</xdr:col><xdr:colOff>${to.colOff ?? 0}</xdr:colOff><xdr:row>${to.row}</xdr:row><xdr:rowOff>${to.rowOff ?? 0}</xdr:rowOff></xdr:to>`;
      this.oleRIds[i];
      const iconRId = this.oleIconRIds[i] ?? "";
      ole.progId;
      ole.displayAsIcon;
      ole.linkToFile && ole.linkPath && `${escapeXml(ole.linkPath)}`;
      iconRId && `${iconRId}`;
      parts.push(
        `<xdr:twoCellAnchor editAs="oneCell">${fromXml}${toXml}<xdr:sp macro=""><xdr:nvSpPr><xdr:cNvPr id="${id}" name="${escapeXml(ole.name)}" hidden="1"/><xdr:cNvSpPr/></xdr:nvSpPr><xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:sp><xdr:clientData/></xdr:twoCellAnchor>`,
      );
    });
    for (const slicerInfo of this._slicerDrawingInfo) {
      const id = shapeCounter++;
      const { row: sRow, col: sCol } = slicerInfo.cell
        ? cellRefToIndices(slicerInfo.cell)
        : {
            row: 1,
            col: 6,
          };
      const fromCol = sCol - 1,
        fromRow = sRow - 1;
      const toCol = fromCol + 2,
        toRow = fromRow + 12;
      parts.push(
        `<xdr:twoCellAnchor editAs="oneCell"><xdr:from><xdr:col>${fromCol}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${fromRow}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>${toCol}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${toRow}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice xmlns:sle15="http://schemas.microsoft.com/office/drawing/2012/slicer" Requires="sle15"><xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="${id}" name="${escapeXml(slicerInfo.name)}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/drawing/2010/slicer"><sle:slicer xmlns:sle="http://schemas.microsoft.com/office/drawing/2010/slicer" name="${escapeXml(slicerInfo.name)}"/></a:graphicData></a:graphic></xdr:graphicFrame></mc:Choice><mc:Fallback/></mc:AlternateContent><xdr:clientData/></xdr:twoCellAnchor>`,
      );
    }
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
${parts.join("\n")}
</xdr:wsDr>`;
  }
};
function hashPassword(pwd) {
  let hash = 0;
  for (let i = pwd.length - 1; i >= 0; i--) {
    hash = ((hash >> 14) & 1) | ((hash << 1) & 32767);
    hash ^= pwd.charCodeAt(i);
  }
  hash = ((hash >> 14) & 1) | ((hash << 1) & 32767);
  hash ^= pwd.length;
  hash ^= 52811;
  return hash.toString(16).toUpperCase().padStart(4, "0");
}
//#endregion
//#region node_modules/@node-projects/excelforge/dist/styles/StyleRegistry.js
function colorAttrXml(tag, color) {
  if (!color) return "";
  if (color.startsWith("theme:")) {
    const rest = color.slice(6);
    const tintIdx = rest.indexOf(":tint:");
    if (tintIdx >= 0)
      return `<${tag} theme="${rest.slice(0, tintIdx)}" tint="${rest.slice(tintIdx + 6)}"/>`;
    return `<${tag} theme="${rest}"/>`;
  }
  return `<${tag} rgb="${color.startsWith("#") ? "FF" + color.slice(1) : color}"/>`;
}
function fontXml(f) {
  const parts = [];
  if (f.bold) parts.push("<b/>");
  if (f.italic) parts.push("<i/>");
  if (f.strike) parts.push("<strike/>");
  if (f.underline && f.underline !== "none") parts.push(`<u val="${f.underline}"/>`);
  if (f.vertAlign) parts.push(`<vertAlign val="${f.vertAlign}"/>`);
  if (f.size) parts.push(`<sz val="${f.size}"/>`);
  if (f.color) parts.push(colorAttrXml("color", f.color));
  if (f.name) parts.push(`<name val="${escapeXml(f.name)}"/>`);
  if (f.family) parts.push(`<family val="${f.family}"/>`);
  if (f.scheme) parts.push(`<scheme val="${f.scheme}"/>`);
  if (f.charset) parts.push(`<charset val="${f.charset}"/>`);
  return parts.join("");
}
function fillXml(fill) {
  if (fill.type === "pattern") {
    const f = fill;
    const fg = colorAttrXml("fgColor", f.fgColor);
    const bg = colorAttrXml("bgColor", f.bgColor);
    return `<patternFill${f.pattern ? ` patternType="${f.pattern}"` : ""}>${fg}${bg}</patternFill>`;
  }
  const f = fill;
  const stops = f.stops
    .map((s) => `<stop position="${s.position}">${colorAttrXml("color", s.color)}</stop>`)
    .join("");
  return `<gradientFill ${[
    f.gradientType ? `type="${f.gradientType}"` : "",
    f.degree !== void 0 ? `degree="${f.degree}"` : "",
    f.left !== void 0 ? `left="${f.left}"` : "",
    f.right !== void 0 ? `right="${f.right}"` : "",
    f.top !== void 0 ? `top="${f.top}"` : "",
    f.bottom !== void 0 ? `bottom="${f.bottom}"` : "",
  ]
    .filter(Boolean)
    .join(" ")}>${stops}</gradientFill>`;
}
function borderSideXml(tag, s) {
  if (!s) return `<${tag}/>`;
  const color = colorAttrXml("color", s.color);
  return s.style ? `<${tag} style="${s.style}">${color}</${tag}>` : `<${tag}/>`;
}
function borderXml(b) {
  const diag = [b.diagonalUp ? 'diagonalUp="1"' : "", b.diagonalDown ? 'diagonalDown="1"' : ""]
    .filter(Boolean)
    .join(" ");
  return (
    `<border${diag ? " " + diag : ""}>` +
    borderSideXml("left", b.left) +
    borderSideXml("right", b.right) +
    borderSideXml("top", b.top) +
    borderSideXml("bottom", b.bottom) +
    borderSideXml("diagonal", b.diagonal) +
    "</border>"
  );
}
function alignmentXml(a) {
  return `<alignment ${[
    a.horizontal ? `horizontal="${a.horizontal}"` : "",
    a.vertical ? `vertical="${a.vertical}"` : "",
    a.wrapText ? `wrapText="1"` : "",
    a.shrinkToFit ? `shrinkToFit="1"` : "",
    a.textRotation !== void 0 ? `textRotation="${a.textRotation}"` : "",
    a.indent ? `indent="${a.indent}"` : "",
    a.readingOrder !== void 0 ? `readingOrder="${a.readingOrder}"` : "",
  ]
    .filter(Boolean)
    .join(" ")}/>`;
}
var StyleRegistry = class {
  constructor() {
    this.fonts = [];
    this.fills = [];
    this.borders = [];
    this.fontIdx = /* @__PURE__ */ new Map();
    this.fillIdx = /* @__PURE__ */ new Map();
    this.borderIdx = /* @__PURE__ */ new Map();
    this.numFmts = /* @__PURE__ */ new Map();
    this.xfs = [];
    this.styleKey = /* @__PURE__ */ new Map();
    this.dxfs = [];
    this.nextNumFmtId = 164;
    this.cellStyleXfs = [];
    this.cellStyleNames = [];
    this.tableStyleDefs = [];
    const defFont = fontXml({
      name: "Calibri",
      size: 11,
      scheme: "minor",
    });
    this.fonts.push(defFont);
    this.fontIdx.set(defFont, 0);
    const fill0 = fillXml({
      type: "pattern",
      pattern: "none",
    });
    const fill1 = fillXml({
      type: "pattern",
      pattern: "gray125",
    });
    this.fills.push(fill0, fill1);
    this.fillIdx.set(fill0, 0);
    this.fillIdx.set(fill1, 1);
    const defBorder = borderXml({});
    this.borders.push(defBorder);
    this.borderIdx.set(defBorder, 0);
    this.cellStyleXfs.push(`<xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>`);
    this.cellStyleNames.push({
      name: "Normal",
      xfId: 0,
      builtinId: 0,
    });
    this.xfs.push(`<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>`);
  }
  registerNamedStyle(name, style, builtinId) {
    const fontId = this.internFont(style.font);
    const fillId = this.internFill(style.fill);
    const borderId = this.internBorder(style.border);
    const xml = `<xf numFmtId="${this.internNumFmt(style.numberFormat, style.numFmtId)}" fontId="${fontId}" fillId="${fillId}" borderId="${borderId}"${style.font ? ' applyFont="1"' : ""}${style.fill ? ' applyFill="1"' : ""}${style.border ? ' applyBorder="1"' : ""}${style.alignment ? ' applyAlignment="1"' : ""}${style.numberFormat || style.numFmtId !== void 0 ? ' applyNumberFormat="1"' : ""}>${style.alignment ? alignmentXml(style.alignment) : ""}</xf>`;
    this.cellStyleXfs.push(xml);
    const xfId = this.cellStyleXfs.length - 1;
    this.cellStyleNames.push({
      name,
      xfId,
      builtinId,
    });
    return xfId;
  }
  internFont(f) {
    if (!f) return 0;
    const xml = fontXml(f);
    const existing = this.fontIdx.get(xml);
    if (existing !== void 0) return existing;
    const idx = this.fonts.length;
    this.fonts.push(xml);
    this.fontIdx.set(xml, idx);
    return idx;
  }
  internFill(f) {
    if (!f) return 0;
    const xml = fillXml(f);
    const existing = this.fillIdx.get(xml);
    if (existing !== void 0) return existing;
    const idx = this.fills.length;
    this.fills.push(xml);
    this.fillIdx.set(xml, idx);
    return idx;
  }
  internBorder(b) {
    if (!b) return 0;
    const xml = borderXml(b);
    const existing = this.borderIdx.get(xml);
    if (existing !== void 0) return existing;
    const idx = this.borders.length;
    this.borders.push(xml);
    this.borderIdx.set(xml, idx);
    return idx;
  }
  internNumFmt(fmt, builtinId) {
    if (builtinId !== void 0) return builtinId;
    if (!fmt) return 0;
    if (this.numFmts.has(fmt.formatCode)) return this.numFmts.get(fmt.formatCode);
    const id = this.nextNumFmtId++;
    this.numFmts.set(fmt.formatCode, id);
    return id;
  }
  register(style) {
    if (!style) return 0;
    const key = JSON.stringify(style);
    if (this.styleKey.has(key)) return this.styleKey.get(key);
    const fontId = this.internFont(style.font);
    const fillId = this.internFill(style.fill);
    const borderId = this.internBorder(style.border);
    const numFmtId = this.internNumFmt(style.numberFormat, style.numFmtId);
    const applyFont = style.font ? ' applyFont="1"' : "";
    const applyFill = style.fill ? ' applyFill="1"' : "";
    const applyBorder = style.border ? ' applyBorder="1"' : "";
    const applyAlignment = style.alignment ? ' applyAlignment="1"' : "";
    const applyNumFmt =
      style.numberFormat || style.numFmtId !== void 0 ? ' applyNumberFormat="1"' : "";
    const applyProtection =
      style.locked !== void 0 || style.hidden !== void 0 ? ' applyProtection="1"' : "";
    const align = style.alignment ? alignmentXml(style.alignment) : "";
    const prot =
      style.locked !== void 0 || style.hidden !== void 0
        ? `<protection${style.locked !== void 0 ? ` locked="${style.locked ? "1" : "0"}"` : ""}${style.hidden !== void 0 ? ` hidden="${style.hidden ? "1" : "0"}"` : ""}/>`
        : "";
    const xml = `<xf numFmtId="${numFmtId}" fontId="${fontId}" fillId="${fillId}" borderId="${borderId}" xfId="${style.namedStyleId ?? 0}"${applyFont}${applyFill}${applyBorder}${applyAlignment}${applyNumFmt}${applyProtection}>${align}${prot}</xf>`;
    this.xfs.push(xml);
    const idx = this.xfs.length - 1;
    this.styleKey.set(key, idx);
    return idx;
  }
  registerDxf(style) {
    const parts = [];
    if (style.font) parts.push(`<font>${fontXml(style.font)}</font>`);
    if (style.numberFormat) {
      const id = this.internNumFmt(style.numberFormat);
      parts.push(
        `<numFmt numFmtId="${id}" formatCode="${escapeXml(style.numberFormat.formatCode)}"/>`,
      );
    }
    if (style.fill) parts.push(`<fill>${fillXml(style.fill)}</fill>`);
    if (style.alignment)
      parts.push(
        `<alignment${style.alignment.horizontal ? ` horizontal="${style.alignment.horizontal}"` : ""}${style.alignment.vertical ? ` vertical="${style.alignment.vertical}"` : ""}${style.alignment.wrapText ? ' wrapText="1"' : ""}${style.alignment.textRotation ? ` textRotation="${style.alignment.textRotation}"` : ""}/>`,
      );
    if (style.border) parts.push(borderXml(style.border));
    const xml = parts.join("");
    this.dxfs.push(xml);
    return this.dxfs.length - 1;
  }
  prependRawDxfs(rawInners) {
    this.dxfs.unshift(...rawInners);
  }
  registerTableStyle(name, def) {
    const elements = [];
    if (def.headerRow)
      elements.push({
        type: "headerRow",
        dxfId: this.registerDxf(def.headerRow),
      });
    if (def.totalRow)
      elements.push({
        type: "totalRow",
        dxfId: this.registerDxf(def.totalRow),
      });
    if (def.dataRow1)
      elements.push({
        type: "firstRowStripe",
        dxfId: this.registerDxf(def.dataRow1),
      });
    if (def.dataRow2)
      elements.push({
        type: "secondRowStripe",
        dxfId: this.registerDxf(def.dataRow2),
      });
    this.tableStyleDefs.push({
      name,
      elements,
    });
  }
  toXml() {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
${this.numFmts.size ? `<numFmts count="${this.numFmts.size}">${[...this.numFmts.entries()].map(([fmt, id]) => `<numFmt numFmtId="${id}" formatCode="${escapeXml(fmt)}"/>`).join("")}</numFmts>` : ""}
<fonts count="${this.fonts.length}">${this.fonts.map((f) => `<font>${f}</font>`).join("")}</fonts>
<fills count="${this.fills.length}">${this.fills.map((f) => `<fill>${f}</fill>`).join("")}</fills>
<borders count="${this.borders.length}">${this.borders.join("")}</borders>
<cellStyleXfs count="${this.cellStyleXfs.length}">${this.cellStyleXfs.join("")}</cellStyleXfs>
<cellXfs count="${this.xfs.length}">${this.xfs.join("")}</cellXfs>
<cellStyles count="${this.cellStyleNames.length}">${this.cellStyleNames.map((cs) => `<cellStyle name="${escapeXml(cs.name)}" xfId="${cs.xfId}"${cs.builtinId !== void 0 ? ` builtinId="${cs.builtinId}"` : ""}/>`).join("")}</cellStyles>
${this.dxfs.length ? `<dxfs count="${this.dxfs.length}">${this.dxfs.map((d) => `<dxf>${d}</dxf>`).join("")}</dxfs>` : ""}
${this.tableStyleDefs.length ? `<tableStyles count="${this.tableStyleDefs.length}">${this.tableStyleDefs.map((ts) => `<tableStyle name="${escapeXml(ts.name)}" count="${ts.elements.length}">${ts.elements.map((el) => `<tableStyleElement type="${el.type}" dxfId="${el.dxfId}"/>`).join("")}</tableStyle>`).join("")}</tableStyles>` : ""}
</styleSheet>`;
  }
};
//#endregion
//#region node_modules/@node-projects/excelforge/dist/core/SharedStrings.js
var SharedStrings = class {
  constructor() {
    this.table = /* @__PURE__ */ new Map();
    this.strings = [];
    this._count = 0;
  }
  get count() {
    return this._count;
  }
  get uniqueCount() {
    return this.strings.length;
  }
  intern(s) {
    this._count++;
    const existing = this.table.get(s);
    if (existing !== void 0) return existing;
    const idx = this.strings.length;
    this.strings.push(s);
    this.table.set(s, idx);
    return idx;
  }
  internRichText(runs) {
    const key = JSON.stringify(runs);
    this._count++;
    const existing = this.table.get(key);
    if (existing !== void 0) return existing;
    const xml = runs
      .map((r) => {
        return `<r>${r.font ? richFontXml(r.font) : ""}<t xml:space="preserve">${escapeXml(r.text)}</t></r>`;
      })
      .join("");
    const idx = this.strings.length;
    this.strings.push("\0RICH\0" + xml);
    this.table.set(key, idx);
    return idx;
  }
  toXml() {
    const items = this.strings
      .map((s) => {
        if (s.startsWith("\0RICH\0")) return `<si>${s.slice(6)}</si>`;
        return `<si><t${s !== s.trim() || s.includes("\n") ? ' xml:space="preserve"' : ""}>${escapeXml(s)}</t></si>`;
      })
      .join("");
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="${this._count}" uniqueCount="${this.strings.length}">
${items}
</sst>`;
  }
};
function richFontXml(f) {
  const parts = [];
  if (f.bold) parts.push("<b/>");
  if (f.italic) parts.push("<i/>");
  if (f.strike) parts.push("<strike/>");
  if (f.underline && f.underline !== "none") parts.push(`<u val="${f.underline}"/>`);
  if (f.size) parts.push(`<sz val="${f.size}"/>`);
  if (f.color)
    if (f.color.startsWith("theme:")) parts.push(`<color theme="${f.color.slice(6)}"/>`);
    else
      parts.push(`<color rgb="${f.color.startsWith("#") ? "FF" + f.color.slice(1) : f.color}"/>`);
  if (f.name) parts.push(`<name val="${escapeXml(f.name)}"/>`);
  return parts.length ? `<rPr>${parts.join("")}</rPr>` : "";
}
//#endregion
//#region node_modules/@node-projects/excelforge/dist/features/TableBuilder.js
function buildTableXml(table, tableId) {
  const { startRow, startCol, endRow, endCol } = parseRange(table.ref);
  const displayName = table.displayName ?? table.name;
  const colsXml = table.columns
    .map((col, i) => {
      startCol + i;
      const id = i + 1;
      const totFn = col.totalsRowFunction ?? "none";
      const totAttrs = table.totalsRow
        ? ` totalsRowFunction="${totFn}"${totFn !== "none" && col.totalsRowFormula ? ` totalsRowFormula="${escapeXml(col.totalsRowFormula)}"` : ""}`
        : "";
      const totLabel =
        table.totalsRow && col.totalsRowLabel
          ? ` totalsRowLabel="${escapeXml(col.totalsRowLabel)}"`
          : "";
      return `<tableColumn id="${id}" name="${escapeXml(col.name)}"${totAttrs}${totLabel}/>`;
    })
    .join("");
  const styleAttrs = [
    `name="${escapeXml(table.style ?? "TableStyleMedium2")}"`,
    table.showFirstColumn ? 'showFirstColumn="1"' : "",
    table.showLastColumn ? 'showLastColumn="1"' : "",
    table.showRowStripes !== false ? 'showRowStripes="1"' : "",
    table.showColumnStripes ? 'showColumnStripes="1"' : "",
  ]
    .filter(Boolean)
    .join(" ");
  const totalsRow = table.totalsRow ? 1 : 0;
  const autoFilterRef = `${colIndexToLetter(startCol)}${startRow}:${colIndexToLetter(endCol)}${endRow - totalsRow}`;
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  id="${tableId}" name="${escapeXml(table.name)}" displayName="${escapeXml(displayName)}"
  ref="${table.ref}" totalsRowCount="${totalsRow}">
  <autoFilter ref="${autoFilterRef}"/>
  <tableColumns count="${table.columns.length}">${colsXml}</tableColumns>
  <tableStyleInfo ${styleAttrs}/>
</table>`;
}
//#endregion
//#region node_modules/@node-projects/excelforge/dist/utils/zip.js
var _crcTable = null;
function crc32Table() {
  if (_crcTable) return _crcTable;
  _crcTable = /* @__PURE__ */ new Uint32Array(256);
  for (let i = 0; i < 256; i++) {
    let c = i;
    for (let j = 0; j < 8; j++) c = c & 1 ? 3988292384 ^ (c >>> 1) : c >>> 1;
    _crcTable[i] = c;
  }
  return _crcTable;
}
function crc32(data) {
  const t = crc32Table();
  let crc = 4294967295;
  for (let i = 0; i < data.length; i++) crc = (crc >>> 8) ^ t[(crc ^ data[i]) & 255];
  return (crc ^ 4294967295) >>> 0;
}
var BitWriter = class {
  constructor(initialSize = 65536) {
    this.pos = 0;
    this.bits = 0;
    this.bitLen = 0;
    this.buf = new Uint8Array(initialSize);
  }
  grow() {
    const next = new Uint8Array(this.buf.length * 2);
    next.set(this.buf);
    this.buf = next;
  }
  writeBits(val, n) {
    this.bits |= (val & ((1 << n) - 1)) << this.bitLen;
    this.bitLen += n;
    while (this.bitLen >= 8) {
      if (this.pos >= this.buf.length) this.grow();
      this.buf[this.pos++] = this.bits & 255;
      this.bits >>>= 8;
      this.bitLen -= 8;
    }
  }
  writeByte(b) {
    if (this.pos >= this.buf.length) this.grow();
    this.buf[this.pos++] = b;
  }
  flush() {
    if (this.bitLen > 0) {
      this.writeByte(this.bits & 255);
      this.bits = 0;
      this.bitLen = 0;
    }
  }
  toBytes() {
    return this.buf.subarray(0, this.pos);
  }
};
function buildCodeLengths(freq, maxBits) {
  const n = freq.length;
  const lengths = new Uint8Array(n);
  const nonZero = [];
  for (let i = 0; i < n; i++) if (freq[i] > 0) nonZero.push(i);
  if (nonZero.length === 0) return lengths;
  if (nonZero.length === 1) {
    lengths[nonZero[0]] = 1;
    return lengths;
  }
  const heap = nonZero.map((s) => ({
    freq: freq[s],
    sym: s,
  }));
  const siftDown = (arr, i, end) => {
    while (true) {
      let min = i;
      const l = 2 * i + 1,
        r = 2 * i + 2;
      if (l < end && arr[l].freq < arr[min].freq) min = l;
      if (r < end && arr[r].freq < arr[min].freq) min = r;
      if (min === i) break;
      const tmp = arr[i];
      arr[i] = arr[min];
      arr[min] = tmp;
      i = min;
    }
  };
  for (let i = (heap.length >> 1) - 1; i >= 0; i--) siftDown(heap, i, heap.length);
  let hSize = heap.length;
  while (hSize > 1) {
    const a = heap[0];
    heap[0] = heap[--hSize];
    siftDown(heap, 0, hSize);
    const b = heap[0];
    heap[0] = {
      freq: a.freq + b.freq,
      sym: -1,
      left: a,
      right: b,
    };
    siftDown(heap, 0, hSize);
  }
  const setDepth = (node, depth) => {
    if (!node) return;
    if (node.sym >= 0) {
      lengths[node.sym] = Math.min(depth, maxBits);
      return;
    }
    setDepth(node.left, depth + 1);
    setDepth(node.right, depth + 1);
  };
  setDepth(heap[0], 0);
  {
    const computeKraft = () => {
      let k = 0;
      for (let i = 0; i < n; i++) if (lengths[i] > 0) k += 1 << (maxBits - lengths[i]);
      return k;
    };
    const target = 1 << maxBits;
    let kraft = computeKraft();
    if (kraft > target) {
      let toFix = kraft - target;
      for (let bl = maxBits - 1; bl >= 1 && toFix > 0; bl--) {
        const freed = 1 << (maxBits - bl - 1);
        for (let i = 0; i < n && toFix > 0; i++)
          if (lengths[i] === bl) {
            lengths[i] = bl + 1;
            toFix -= freed;
          }
      }
    }
    kraft = computeKraft();
    while (kraft < target) {
      let fixed = false;
      for (let i = 0; i < n; i++)
        if (lengths[i] === maxBits) {
          lengths[i]--;
          kraft++;
          fixed = true;
          break;
        }
      if (!fixed) break;
    }
  }
  return lengths;
}
function buildCodes(lengths) {
  const maxLen = Math.max(...lengths);
  const blCount = new Uint32Array(maxLen + 1);
  for (const l of lengths) if (l > 0) blCount[l]++;
  const nextCode = new Uint32Array(maxLen + 1);
  let code = 0;
  for (let bits = 1; bits <= maxLen; bits++) {
    code = (code + blCount[bits - 1]) << 1;
    nextCode[bits] = code;
  }
  const codes = new Uint32Array(lengths.length);
  for (let i = 0; i < lengths.length; i++) if (lengths[i] > 0) codes[i] = nextCode[lengths[i]]++;
  return codes;
}
function reverseBits(val, len) {
  let r = 0;
  for (let i = 0; i < len; i++) {
    r = (r << 1) | (val & 1);
    val >>= 1;
  }
  return r;
}
function fixedLitLengths() {
  const l = /* @__PURE__ */ new Uint8Array(288);
  for (let i = 0; i <= 143; i++) l[i] = 8;
  for (let i = 144; i <= 255; i++) l[i] = 9;
  for (let i = 256; i <= 279; i++) l[i] = 7;
  for (let i = 280; i <= 287; i++) l[i] = 8;
  return l;
}
function fixedDistLengths() {
  const l = /* @__PURE__ */ new Uint8Array(32);
  l.fill(5);
  return l;
}
var LENGTH_EXTRA = [
  0, 0, 0, 0, 0, 0, 0, 0, 1, 1, 1, 1, 2, 2, 2, 2, 3, 3, 3, 3, 4, 4, 4, 4, 5, 5, 5, 5, 0, 0, 0,
];
var LENGTH_BASE = [
  3, 4, 5, 6, 7, 8, 9, 10, 11, 13, 15, 17, 19, 23, 27, 31, 35, 43, 51, 59, 67, 83, 99, 115, 131,
  163, 195, 227, 258, 0, 0,
];
var DIST_EXTRA = [
  0, 0, 0, 0, 1, 1, 2, 2, 3, 3, 4, 4, 5, 5, 6, 6, 7, 7, 8, 8, 9, 9, 10, 10, 11, 11, 12, 12, 13, 13,
];
var DIST_BASE = [
  1, 2, 3, 4, 5, 7, 9, 13, 17, 25, 33, 49, 65, 97, 129, 193, 257, 385, 513, 769, 1025, 1537, 2049,
  3073, 4097, 6145, 8193, 12289, 16385, 24577,
];
var _lenLUT = new Array(259);
for (let i = 0; i < 29; i++) {
  const base = LENGTH_BASE[i],
    extra = LENGTH_EXTRA[i];
  for (let v = base; v < base + (1 << extra); v++)
    if (v <= 258) _lenLUT[v] = [257 + i, extra, v - base];
}
_lenLUT[258] = [285, 0, 0];
function lenCode(len) {
  return _lenLUT[len];
}
var _distLUT = new Array(32769);
for (let i = 0; i < 30; i++) {
  const base = DIST_BASE[i],
    extra = DIST_EXTRA[i];
  for (let v = base; v < base + (1 << extra); v++)
    if (v <= 32768) _distLUT[v] = [i, extra, v - base];
}
function distCode(dist) {
  return _distLUT[dist];
}
var WSIZE = 32768;
var MAX_MATCH = 258;
var MIN_MATCH = 3;
var CHAIN_LEN_FAST = 8;
var CHAIN_LEN_DEFAULT = 32;
var CHAIN_LEN_BEST = 128;
function lz77(data, effort) {
  const chainLen = effort <= 1 ? CHAIN_LEN_FAST : effort <= 6 ? CHAIN_LEN_DEFAULT : CHAIN_LEN_BEST;
  const n = data.length;
  let cap = Math.min(n + 1, 65536);
  let litLen = new Uint16Array(cap);
  let dist = new Uint16Array(cap);
  let count = 0;
  const ensure = () => {
    if (count >= cap) {
      cap = cap * 2;
      const nl = new Uint16Array(cap);
      nl.set(litLen);
      litLen = nl;
      const nd = new Uint16Array(cap);
      nd.set(dist);
      dist = nd;
    }
  };
  const HSIZE = 65536;
  const head = new Int32Array(HSIZE).fill(-1);
  const prev = new Int32Array(Math.min(n, WSIZE)).fill(-1);
  const hash3 = (pos) =>
    (((data[pos] * 4129) ^ (data[pos + 1] * 155) ^ data[pos + 2]) & (HSIZE - 1)) >>> 0;
  let i = 0;
  const emitLit = (b) => {
    ensure();
    litLen[count] = b + 1;
    dist[count] = 0;
    count++;
  };
  const emitMatch = (len, d) => {
    ensure();
    litLen[count] = len;
    dist[count] = d;
    count++;
  };
  while (i < n) {
    if (i + 2 >= n) {
      emitLit(data[i++]);
      continue;
    }
    const h = hash3(i);
    let matchLen = MIN_MATCH - 1;
    let matchDist = 0;
    let chain = head[h];
    let steps = 0;
    while (chain >= 0 && steps < chainLen) {
      const d = i - chain;
      if (d > WSIZE) break;
      let mLen = 0;
      const limit = Math.min(MAX_MATCH, n - i);
      while (mLen < limit && data[chain + mLen] === data[i + mLen]) mLen++;
      if (mLen > matchLen) {
        matchLen = mLen;
        matchDist = d;
      }
      if (matchLen === MAX_MATCH) break;
      chain = prev[chain & (WSIZE - 1)];
      steps++;
    }
    prev[i & (WSIZE - 1)] = head[h];
    head[h] = i;
    if (matchLen >= MIN_MATCH) {
      if (effort >= 6 && i + 3 < n) {
        const h2 = hash3(i + 1);
        let chain2 = head[h2];
        let lazyLen = 0,
          lazyDist = 0;
        let steps2 = 0;
        while (chain2 >= 0 && steps2 < chainLen) {
          const d2 = i + 1 - chain2;
          if (d2 > WSIZE) break;
          let ml = 0;
          const lim = Math.min(MAX_MATCH, n - i - 1);
          while (ml < lim && data[chain2 + ml] === data[i + 1 + ml]) ml++;
          if (ml > lazyLen) {
            lazyLen = ml;
            lazyDist = d2;
          }
          chain2 = prev[chain2 & (WSIZE - 1)];
          steps2++;
        }
        if (lazyLen > matchLen + 1) {
          emitLit(data[i]);
          i++;
          prev[i & (WSIZE - 1)] = head[h2];
          head[h2] = i;
          emitMatch(lazyLen, lazyDist);
          for (let k = 1; k < lazyLen; k++) {
            i++;
            if (i + 2 < n) {
              const hk = hash3(i);
              prev[i & (WSIZE - 1)] = head[hk];
              head[hk] = i;
            }
          }
          i++;
          continue;
        }
      }
      emitMatch(matchLen, matchDist);
      for (let k = 0; k < matchLen; k++)
        if (i + k + 2 < n) {
          const hk = hash3(i + k);
          prev[(i + k) & (WSIZE - 1)] = head[hk];
          head[hk] = i + k;
        }
      i += matchLen;
    } else emitLit(data[i++]);
  }
  return {
    litLen,
    dist,
    count,
  };
}
function encodeCodeLengths(lengths, bw, clCodes, clLens) {
  let i = 0;
  while (i < lengths.length) {
    const l = lengths[i];
    if (l === 0) {
      let run = 0;
      while (i + run < lengths.length && lengths[i + run] === 0 && run < 138) run++;
      if (run < 3) {
        bw.writeBits(reverseBits(clCodes[0], clLens[0]), clLens[0]);
        i++;
      } else if (run <= 10) {
        bw.writeBits(reverseBits(clCodes[17], clLens[17]), clLens[17]);
        bw.writeBits(run - 3, 3);
        i += run;
      } else {
        bw.writeBits(reverseBits(clCodes[18], clLens[18]), clLens[18]);
        bw.writeBits(run - 11, 7);
        i += run;
      }
    } else {
      bw.writeBits(reverseBits(clCodes[l], clLens[l]), clLens[l]);
      i++;
      let run = 0;
      while (i + run < lengths.length && lengths[i + run] === l && run < 6) run++;
      if (run >= 3) {
        bw.writeBits(reverseBits(clCodes[16], clLens[16]), clLens[16]);
        bw.writeBits(run - 3, 2);
        i += run;
      }
    }
  }
}
function deflateBlock(tokens, bw, isLast, useDynamic) {
  const litFreq = /* @__PURE__ */ new Uint32Array(286);
  const distFreq = /* @__PURE__ */ new Uint32Array(30);
  litFreq[256] = 1;
  const { litLen: tLitLen, dist: tDist, count: tCount } = tokens;
  for (let t = 0; t < tCount; t++)
    if (tDist[t] === 0) litFreq[tLitLen[t] - 1]++;
    else {
      const [lSym] = lenCode(tLitLen[t]);
      litFreq[lSym]++;
      const [dSym] = distCode(tDist[t]);
      distFreq[dSym]++;
    }
  let litLens, distLens;
  let litCodes, distCodes;
  if (useDynamic) {
    litLens = buildCodeLengths(litFreq, 15);
    distLens = buildCodeLengths(distFreq, 15);
    if (distLens.every((l) => l === 0)) distLens[0] = 1;
    litCodes = buildCodes(litLens);
    distCodes = buildCodes(distLens);
    const hlit = findMaxUsed(litLens, 257, 286) + 1;
    const hdist = findMaxUsed(distLens, 1, 30) + 1;
    const allLens = [...litLens.subarray(0, hlit), ...distLens.subarray(0, hdist)];
    const clFreq = /* @__PURE__ */ new Uint32Array(19);
    simulateCL(allLens, clFreq);
    const CL_ORDER = [16, 17, 18, 0, 8, 7, 9, 6, 10, 5, 11, 4, 12, 3, 13, 2, 14, 1, 15];
    const clLens_ = buildCodeLengths(clFreq, 7);
    const clCodes = buildCodes(clLens_);
    const hclen =
      findMaxUsedArr(
        CL_ORDER.map((i) => clLens_[i]),
        4,
        19,
      ) + 1;
    bw.writeBits(isLast ? 1 : 0, 1);
    bw.writeBits(2, 2);
    bw.writeBits(hlit - 257, 5);
    bw.writeBits(hdist - 1, 5);
    bw.writeBits(hclen - 4, 4);
    for (let i = 0; i < hclen; i++) bw.writeBits(clLens_[CL_ORDER[i]], 3);
    encodeCodeLengths(allLens, bw, clCodes, clLens_);
  } else {
    litLens = fixedLitLengths();
    distLens = fixedDistLengths();
    litCodes = buildCodes(litLens);
    distCodes = buildCodes(distLens);
    bw.writeBits(isLast ? 1 : 0, 1);
    bw.writeBits(1, 2);
  }
  for (let t = 0; t < tCount; t++)
    if (tDist[t] === 0) {
      const sym = tLitLen[t] - 1;
      const l = litLens[sym];
      bw.writeBits(reverseBits(litCodes[sym], l), l);
    } else {
      const [lSym, lExtra, lVal] = lenCode(tLitLen[t]);
      const ll = litLens[lSym];
      bw.writeBits(reverseBits(litCodes[lSym], ll), ll);
      if (lExtra > 0) bw.writeBits(lVal, lExtra);
      const [dSym, dExtra, dVal] = distCode(tDist[t]);
      const dl = distLens[dSym];
      bw.writeBits(reverseBits(distCodes[dSym], dl), dl);
      if (dExtra > 0) bw.writeBits(dVal, dExtra);
    }
  const eobLen = litLens[256];
  bw.writeBits(reverseBits(litCodes[256], eobLen), eobLen);
}
function findMaxUsed(arr, min, max) {
  let r = min - 1;
  for (let i = min; i < Math.min(arr.length, max); i++) if (arr[i] > 0) r = i;
  return Math.max(r, min - 1);
}
function findMaxUsedArr(arr, min, max) {
  let r = min - 1;
  for (let i = min; i < Math.min(arr.length, max); i++) if (arr[i] > 0) r = i;
  return Math.max(r, min - 1);
}
function simulateCL(lengths, freq) {
  let i = 0;
  while (i < lengths.length) {
    const l = lengths[i];
    if (l === 0) {
      let run = 0;
      while (i + run < lengths.length && lengths[i + run] === 0 && run < 138) run++;
      if (run < 3) {
        freq[0]++;
        i++;
      } else if (run <= 10) {
        freq[17]++;
        i += run;
      } else {
        freq[18]++;
        i += run;
      }
    } else {
      freq[l]++;
      i++;
      let run = 0;
      while (i + run < lengths.length && lengths[i + run] === l && run < 6) run++;
      if (run >= 3) {
        freq[16]++;
        i += run;
      }
    }
  }
}
var BLOCK_SIZE = 65536;
function deflateRaw(data, level = 6) {
  if (level === 0) {
    const bw = new BitWriter();
    let offset = 0;
    while (offset < data.length || data.length === 0) {
      const size = Math.min(BLOCK_SIZE, data.length - offset);
      const isLast = offset + size >= data.length;
      bw.flush();
      bw.writeBits(isLast ? 1 : 0, 1);
      bw.writeBits(0, 2);
      bw.flush();
      const len = size;
      bw.writeByte(len & 255);
      bw.writeByte((len >> 8) & 255);
      bw.writeByte(~len & 255);
      bw.writeByte((~len >> 8) & 255);
      while (bw.pos + size > bw.buf.length)
        bw.buf = (() => {
          const n = new Uint8Array(bw.buf.length * 2);
          n.set(bw.buf);
          return n;
        })();
      bw.buf.set(data.subarray(offset, offset + size), bw.pos);
      bw.pos += size;
      offset += size;
      if (data.length === 0) break;
    }
    return bw.toBytes();
  }
  const effort = Math.max(1, Math.min(9, level));
  const useDynamic = effort >= 2;
  const bw = new BitWriter();
  let offset = 0;
  while (offset < data.length || data.length === 0) {
    const chunk = data.subarray(offset, offset + BLOCK_SIZE);
    const isLast = offset + BLOCK_SIZE >= data.length;
    deflateBlock(lz77(chunk, effort), bw, isLast, useDynamic);
    offset += chunk.length;
    if (data.length === 0) {
      deflateBlock(
        {
          litLen: /* @__PURE__ */ new Uint16Array(0),
          dist: /* @__PURE__ */ new Uint16Array(0),
          count: 0,
        },
        bw,
        true,
        useDynamic,
      );
      break;
    }
  }
  bw.flush();
  return bw.toBytes();
}
function writeUint16LE(buf, o, v) {
  buf[o] = v & 255;
  buf[o + 1] = (v >> 8) & 255;
}
function writeUint32LE(buf, o, v) {
  buf[o] = v & 255;
  buf[o + 1] = (v >> 8) & 255;
  buf[o + 2] = (v >> 16) & 255;
  buf[o + 3] = (v >> 24) & 255;
}
var textEnc = new TextEncoder();
var DEFAULT_NO_COMPRESS = /* @__PURE__ */ new Set([
  "png",
  "jpg",
  "jpeg",
  "gif",
  "tiff",
  "emf",
  "wmf",
  "bmp",
  "webp",
]);
function buildZip(entries, opts = {}) {
  const globalLevel = opts.level ?? 6;
  const noCompress = opts.noCompress
    ? new Set(opts.noCompress.map((e) => e.toLowerCase()))
    : DEFAULT_NO_COMPRESS;
  const localParts = [];
  const centralDir = [];
  let offset = 0;
  for (const entry of entries) {
    const nameBytes = textEnc.encode(entry.name);
    const ext = entry.name.split(".").pop()?.toLowerCase() ?? "";
    const entryLevel = entry.level ?? (noCompress.has(ext) ? 0 : globalLevel);
    const useDeflate = entryLevel > 0 && entry.data.length > 0;
    const rawCrc = crc32(entry.data);
    let compData;
    let method;
    if (useDeflate) {
      compData = deflateRaw(entry.data, entryLevel);
      if (compData.length >= entry.data.length) {
        compData = entry.data;
        method = 0;
      } else method = 8;
    } else {
      compData = entry.data;
      method = 0;
    }
    const lh = new Uint8Array(30 + nameBytes.length);
    writeUint32LE(lh, 0, 67324752);
    writeUint16LE(lh, 4, 20);
    writeUint16LE(lh, 6, 0);
    writeUint16LE(lh, 8, method);
    writeUint16LE(lh, 10, 0);
    writeUint16LE(lh, 12, 0);
    writeUint32LE(lh, 14, rawCrc);
    writeUint32LE(lh, 18, compData.length);
    writeUint32LE(lh, 22, entry.data.length);
    writeUint16LE(lh, 26, nameBytes.length);
    writeUint16LE(lh, 28, 0);
    lh.set(nameBytes, 30);
    const cd = new Uint8Array(46 + nameBytes.length);
    writeUint32LE(cd, 0, 33639248);
    writeUint16LE(cd, 4, 20);
    writeUint16LE(cd, 6, 20);
    writeUint16LE(cd, 8, 0);
    writeUint16LE(cd, 10, method);
    writeUint16LE(cd, 12, 0);
    writeUint16LE(cd, 14, 0);
    writeUint32LE(cd, 16, rawCrc);
    writeUint32LE(cd, 20, compData.length);
    writeUint32LE(cd, 24, entry.data.length);
    writeUint16LE(cd, 28, nameBytes.length);
    writeUint16LE(cd, 30, 0);
    writeUint16LE(cd, 32, 0);
    writeUint16LE(cd, 34, 0);
    writeUint16LE(cd, 36, 0);
    writeUint32LE(cd, 38, 0);
    writeUint32LE(cd, 42, offset);
    cd.set(nameBytes, 46);
    localParts.push(lh, compData);
    centralDir.push(cd);
    offset += lh.length + compData.length;
  }
  const cdStart = offset;
  const cdSize = centralDir.reduce((s, c) => s + c.length, 0);
  const eocd = /* @__PURE__ */ new Uint8Array(22);
  writeUint32LE(eocd, 0, 101010256);
  writeUint16LE(eocd, 4, 0);
  writeUint16LE(eocd, 6, 0);
  writeUint16LE(eocd, 8, centralDir.length);
  writeUint16LE(eocd, 10, centralDir.length);
  writeUint32LE(eocd, 12, cdSize);
  writeUint32LE(eocd, 16, cdStart);
  writeUint16LE(eocd, 20, 0);
  const all = [...localParts, ...centralDir, eocd];
  const total = all.reduce((s, a) => s + a.length, 0);
  const out = new Uint8Array(total);
  let pos = 0;
  for (const a of all) {
    out.set(a, pos);
    pos += a.length;
  }
  return out;
}
//#endregion
//#region node_modules/@node-projects/excelforge/dist/utils/zipReader.js
function readUint16LE(buf, o) {
  return buf[o] | (buf[o + 1] << 8);
}
function readUint32LE(buf, o) {
  return (buf[o] | (buf[o + 1] << 8) | (buf[o + 2] << 16) | (buf[o + 3] << 24)) >>> 0;
}
async function inflateAsync(data) {
  const ds = new DecompressionStream("deflate-raw");
  const writer = ds.writable.getWriter();
  const reader = ds.readable.getReader();
  writer.write(data);
  writer.close();
  const chunks = [];
  let done = false;
  while (!done) {
    const { value, done: d } = await reader.read();
    if (value) chunks.push(value);
    done = d;
  }
  const total = chunks.reduce((s, c) => s + c.length, 0);
  const out = new Uint8Array(total);
  let pos = 0;
  for (const c of chunks) {
    out.set(c, pos);
    pos += c.length;
  }
  return out;
}
async function readZip(data) {
  const dec = new TextDecoder("utf-8");
  const entries = /* @__PURE__ */ new Map();
  let eocdPos = -1;
  for (let i = data.length - 22; i >= 0; i--)
    if (data[i] === 80 && data[i + 1] === 75 && data[i + 2] === 5 && data[i + 3] === 6) {
      eocdPos = i;
      break;
    }
  if (eocdPos < 0) throw new Error("Not a valid ZIP file (EOCD not found)");
  const cdOffset = readUint32LE(data, eocdPos + 16);
  const cdCount = readUint16LE(data, eocdPos + 8);
  let pos = cdOffset;
  for (let i = 0; i < cdCount; i++) {
    if (readUint32LE(data, pos) !== 33639248) throw new Error("Invalid central directory entry");
    const method = readUint16LE(data, pos + 10);
    const compSize = readUint32LE(data, pos + 20);
    readUint32LE(data, pos + 24);
    const nameLen = readUint16LE(data, pos + 28);
    const extraLen = readUint16LE(data, pos + 30);
    const commentLen = readUint16LE(data, pos + 32);
    const localOffset = readUint32LE(data, pos + 42);
    const nameBytes = data.subarray(pos + 46, pos + 46 + nameLen);
    const name = dec.decode(nameBytes);
    pos += 46 + nameLen + extraLen + commentLen;
    const lhBase = localOffset;
    if (readUint32LE(data, lhBase) !== 67324752) throw new Error("Invalid local file header");
    const lhNameLen = readUint16LE(data, lhBase + 26);
    const lhExtraLen = readUint16LE(data, lhBase + 28);
    const dataStart = lhBase + 30 + lhNameLen + lhExtraLen;
    const compData = data.subarray(dataStart, dataStart + compSize);
    let fileData;
    if (method === 0) fileData = compData.slice();
    else if (method === 8) fileData = await inflateAsync(compData);
    else fileData = compData.slice();
    if (!name.endsWith("/"))
      entries.set(name, {
        name,
        data: fileData,
      });
  }
  return entries;
}
function entryText(entry) {
  return new TextDecoder("utf-8").decode(entry.data);
}
//#endregion
//#region node_modules/@node-projects/excelforge/dist/utils/xmlParser.js
function nodeToXml(node) {
  const attrs = Object.entries(node.attrs)
    .map(([k, v]) => `${k}="${escAttr(v)}"`)
    .join(" ");
  const open = attrs ? `<${node.tag} ${attrs}` : `<${node.tag}`;
  const inner = (node.text ?? "") + node.children.map(nodeToXml).join("");
  if (!inner) return `${open}/>`;
  return `${open}>${inner}</${node.tag}>`;
}
function escAttr(s) {
  return s
    .replace(/&/g, "&amp;")
    .replace(/"/g, "&quot;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}
function parseXml(xml) {
  const root = [];
  const stack = [];
  let i = 0;
  if (xml.startsWith("<?")) i = xml.indexOf("?>") + 2;
  while (i < xml.length) {
    if (xml[i] !== "<") {
      const end = xml.indexOf("<", i);
      const text = end < 0 ? xml.slice(i) : xml.slice(i, end);
      if (stack.length && text.trim()) {
        const top = stack[stack.length - 1];
        top.text = (top.text ?? "") + decodeEntities(text);
      }
      i = end < 0 ? xml.length : end;
      continue;
    }
    if (xml.startsWith("<!--", i)) {
      const end = xml.indexOf("-->", i);
      if (stack.length) {
        const top = stack[stack.length - 1];
        top.text = (top.text ?? "") + xml.slice(i, end + 3);
      }
      i = end + 3;
      continue;
    }
    if (xml.startsWith("<![CDATA[", i)) {
      const end = xml.indexOf("]]>", i);
      const cdata = xml.slice(i + 9, end);
      if (stack.length) {
        const top = stack[stack.length - 1];
        top.text = (top.text ?? "") + cdata;
      }
      i = end + 3;
      continue;
    }
    if (xml[i + 1] === "/") {
      const end = xml.indexOf(">", i);
      const node = stack.pop();
      if (stack.length) stack[stack.length - 1].children.push(node);
      else root.push(node);
      i = end + 1;
      continue;
    }
    if (xml[i + 1] === "?") {
      i = xml.indexOf("?>", i) + 2;
      continue;
    }
    const end = findTagEnd(xml, i);
    const raw = xml.slice(i + 1, end);
    const selfClose = raw.endsWith("/");
    const { tag, attrs } = parseTag(selfClose ? raw.slice(0, -1).trim() : raw.trim());
    const node = {
      tag,
      attrs,
      children: [],
    };
    if (selfClose)
      if (stack.length) stack[stack.length - 1].children.push(node);
      else root.push(node);
    else stack.push(node);
    i = end + 1;
  }
  while (stack.length > 1) {
    const node = stack.pop();
    stack[stack.length - 1].children.push(node);
  }
  if (stack.length === 1) root.push(stack[0]);
  if (root.length === 0) throw new Error("Empty XML document");
  return root[0];
}
function findTagEnd(xml, start) {
  let inQuote = null;
  for (let i = start + 1; i < xml.length; i++) {
    const c = xml[i];
    if (inQuote) {
      if (c === inQuote) inQuote = null;
    } else if (c === '"' || c === "'") inQuote = c;
    else if (c === ">") return i;
  }
  return xml.length - 1;
}
function parseTag(s) {
  const attrs = {};
  const spaceIdx = s.search(/\s/);
  const tag = spaceIdx < 0 ? s : s.slice(0, spaceIdx);
  if (spaceIdx < 0)
    return {
      tag,
      attrs,
    };
  const rest = s.slice(spaceIdx);
  const re = /(\S+?)\s*=\s*(?:"([^"]*)"|'([^']*)')/g;
  let m;
  while ((m = re.exec(rest)) !== null) attrs[m[1]] = decodeEntities(m[2] ?? m[3] ?? "");
  return {
    tag,
    attrs,
  };
}
function decodeEntities(s) {
  return s
    .replace(/&amp;/g, "&")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&#(\d+);/g, (_, n) => String.fromCharCode(parseInt(n, 10)))
    .replace(/&#x([0-9a-fA-F]+);/g, (_, h) => String.fromCharCode(parseInt(h, 16)));
}
function child(node, localTag) {
  return node.children.find((c) => localName(c.tag) === localTag);
}
function children(node, localTag) {
  return node.children.filter((c) => localName(c.tag) === localTag);
}
function attr(node, name) {
  if (node.attrs[name] !== void 0) return node.attrs[name];
  const local = localName(name);
  for (const [k, v] of Object.entries(node.attrs)) if (localName(k) === local) return v;
}
function localName(tag) {
  const colon = tag.indexOf(":");
  return colon < 0 ? tag : tag.slice(colon + 1);
}
//#endregion
//#region node_modules/@node-projects/excelforge/dist/core/properties.js
function isoDate(d) {
  return d.toISOString();
}
function buildCoreXml(p) {
  const t = (tag, val) => (val !== void 0 ? `<${tag}>${escapeXml(val)}</${tag}>` : "");
  const dt = (tag, val) =>
    val ? `<${tag} xsi:type="dcterms:W3CDTF">${isoDate(val)}</${tag}>` : "";
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties
  xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
  xmlns:dc="http://purl.org/dc/elements/1.1/"
  xmlns:dcterms="http://purl.org/dc/terms/"
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
${t("dc:title", p.title)}
${t("dc:subject", p.subject)}
${t("dc:creator", p.creator)}
${t("cp:keywords", p.keywords)}
${t("dc:description", p.description)}
${t("cp:lastModifiedBy", p.lastModifiedBy)}
${t("cp:revision", p.revision ?? "1")}
${t("dc:language", p.language)}
${t("dc:identifier", p.identifier)}
${t("cp:version", p.version)}
${t("cp:category", p.category)}
${t("cp:contentStatus", p.contentStatus)}
${dt("dcterms:created", p.created ?? /* @__PURE__ */ new Date())}
${dt("dcterms:modified", p.modified ?? /* @__PURE__ */ new Date())}
</cp:coreProperties>`;
}
function buildAppXml(p, extraRaw) {
  const t = (tag, val) => (val !== void 0 ? `<${tag}>${escapeXml(String(val))}</${tag}>` : "");
  let headingPairsXml = "";
  let titlesXml = "";
  if (p.headingPairs?.length) {
    const pairs = p.headingPairs
      .map(
        (hp) =>
          `<vt:variant><vt:lpstr>${escapeXml(hp.name)}</vt:lpstr></vt:variant><vt:variant><vt:i4>${hp.count}</vt:i4></vt:variant>`,
      )
      .join("");
    headingPairsXml = `<HeadingPairs><vt:vector size="${p.headingPairs.length * 2}" baseType="variant">${pairs}</vt:vector></HeadingPairs>`;
  }
  if (p.titlesOfParts?.length) {
    const titles = p.titlesOfParts.map((s) => `<vt:lpstr>${escapeXml(s)}</vt:lpstr>`).join("");
    titlesXml = `<TitlesOfParts><vt:vector size="${p.titlesOfParts.length}" baseType="lpstr">${titles}</vt:vector></TitlesOfParts>`;
  }
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
  xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
${t("Application", p.application ?? "ExcelForge")}
${t("AppVersion", p.appVersion)}
${t("Company", p.company)}
${t("Manager", p.manager)}
${t("DocSecurity", p.docSecurity ?? 0)}
${t("ScaleCrop", p.scaleCrop ?? false)}
${t("LinksUpToDate", p.linksUpToDate ?? false)}
${t("SharedDoc", p.sharedDoc ?? false)}
${t("HyperlinksChanged", p.hyperlinksChanged ?? false)}
${t("Characters", p.characters)}
${t("CharactersWithSpaces", p.charactersByWord)}
${t("Words", p.words)}
${t("Lines", p.lines)}
${t("Paragraphs", p.paragraphs)}
${t("Pages", p.pages)}
${t("Slides", p.slides)}
${t("Notes", p.notes)}
${t("HiddenSlides", p.hiddenSlides)}
${t("MMClips", p.mmClips)}
${t("Template", p.template)}
${t("PresentationFormat", p.presentationFormat)}
${t("TotalTime", p.totalTime)}
${t("HyperlinkBase", p.hyperlinkBase)}
${headingPairsXml}
${titlesXml}
${extraRaw ?? ""}
</Properties>`;
}
function buildCustomXml(props) {
  let pid = 2;
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"
  xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
${props
  .map((p) => {
    let valXml;
    const v = p.value;
    switch (v.type) {
      case "string":
        valXml = `<vt:lpwstr>${escapeXml(v.value)}</vt:lpwstr>`;
        break;
      case "int":
        valXml = `<vt:i4>${v.value}</vt:i4>`;
        break;
      case "decimal":
        valXml = `<vt:decimal>${v.value}</vt:decimal>`;
        break;
      case "bool":
        valXml = `<vt:bool>${v.value}</vt:bool>`;
        break;
      case "date":
        valXml = `<vt:filetime>${v.value.toISOString()}</vt:filetime>`;
        break;
      case "r8":
        valXml = `<vt:r8>${v.value}</vt:r8>`;
        break;
      case "i8":
        valXml = `<vt:i8>${v.value}</vt:i8>`;
        break;
      case "error":
        valXml = `<vt:error>${escapeXml(v.value)}</vt:error>`;
        break;
      default:
        valXml = `<vt:lpwstr></vt:lpwstr>`;
        break;
    }
    return `<property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="${pid++}" name="${escapeXml(p.name)}">${valXml}</property>`;
  })
  .join("\n")}
</Properties>`;
}
function parseCoreXml(xml) {
  const root = parseXml(xml);
  const g = (tag) => {
    const n = root.children.find((c) => localName(c.tag) === tag);
    return n?.text ?? n?.children[0]?.text;
  };
  const gDate = (tag) => {
    const s = g(tag);
    return s ? new Date(s) : void 0;
  };
  return {
    title: g("title"),
    subject: g("subject"),
    creator: g("creator"),
    keywords: g("keywords"),
    description: g("description"),
    lastModifiedBy: g("lastModifiedBy"),
    revision: g("revision"),
    created: gDate("created"),
    modified: gDate("modified"),
    category: g("category"),
    contentStatus: g("contentStatus"),
    language: g("language"),
    identifier: g("identifier"),
    version: g("version"),
  };
}
function parseAppXml(xml) {
  const root = parseXml(xml);
  const g = (tag) => root.children.find((c) => localName(c.tag) === tag)?.text;
  const gNum = (tag) => {
    const v = g(tag);
    return v !== void 0 ? Number(v) : void 0;
  };
  const gBool = (tag) => {
    const v = g(tag);
    return v !== void 0 ? v === "true" || v === "1" : void 0;
  };
  let headingPairs;
  const hpNode = root.children.find((c) => localName(c.tag) === "HeadingPairs");
  if (hpNode) {
    const vec = hpNode.children[0];
    if (vec) {
      const variants = vec.children.filter((c) => localName(c.tag) === "variant");
      headingPairs = [];
      for (let i = 0; i < variants.length - 1; i += 2) {
        const nameNode = variants[i].children[0];
        const countNode = variants[i + 1].children[0];
        headingPairs.push({
          name: nameNode?.text ?? "",
          count: parseInt(countNode?.text ?? "0", 10),
        });
      }
    }
  }
  let titlesOfParts;
  const topNode = root.children.find((c) => localName(c.tag) === "TitlesOfParts");
  if (topNode) {
    const vec = topNode.children[0];
    if (vec) titlesOfParts = vec.children.map((c) => c.text ?? "");
  }
  const knownTags = /* @__PURE__ */ new Set([
    "Application",
    "AppVersion",
    "Company",
    "Manager",
    "DocSecurity",
    "ScaleCrop",
    "LinksUpToDate",
    "SharedDoc",
    "HyperlinksChanged",
    "Characters",
    "CharactersWithSpaces",
    "Words",
    "Lines",
    "Paragraphs",
    "Pages",
    "Slides",
    "Notes",
    "HiddenSlides",
    "MMClips",
    "Template",
    "PresentationFormat",
    "TotalTime",
    "HyperlinkBase",
    "HeadingPairs",
    "TitlesOfParts",
  ]);
  const unknownRaw = root.children
    .filter((c) => !knownTags.has(localName(c.tag)))
    .map((n) => {
      const attrs = Object.entries(n.attrs)
        .map(([k, v]) => ` ${k}="${v}"`)
        .join("");
      const inner = n.text ?? "";
      if (!inner && !n.children.length) return `<${n.tag}${attrs}/>`;
      return `<${n.tag}${attrs}>${inner}</${n.tag}>`;
    })
    .join("\n");
  return {
    props: {
      application: g("Application"),
      appVersion: g("AppVersion"),
      company: g("Company"),
      manager: g("Manager"),
      docSecurity: gNum("DocSecurity"),
      scaleCrop: gBool("ScaleCrop"),
      linksUpToDate: gBool("LinksUpToDate"),
      sharedDoc: gBool("SharedDoc"),
      hyperlinksChanged: gBool("HyperlinksChanged"),
      characters: gNum("Characters"),
      charactersByWord: gNum("CharactersWithSpaces"),
      words: gNum("Words"),
      lines: gNum("Lines"),
      paragraphs: gNum("Paragraphs"),
      pages: gNum("Pages"),
      slides: gNum("Slides"),
      notes: gNum("Notes"),
      hiddenSlides: gNum("HiddenSlides"),
      mmClips: gNum("MMClips"),
      template: g("Template"),
      presentationFormat: g("PresentationFormat"),
      totalTime: gNum("TotalTime"),
      hyperlinkBase: g("HyperlinkBase"),
      headingPairs,
      titlesOfParts,
    },
    unknownRaw,
  };
}
function parseCustomXml(xml) {
  return parseXml(xml)
    .children.filter((c) => localName(c.tag) === "property")
    .map((p) => {
      const name = p.attrs["name"] ?? "";
      const valNode = p.children[0];
      if (!valNode) return null;
      const vTag = localName(valNode.tag);
      const text = valNode.text ?? "";
      let value;
      switch (vTag) {
        case "lpwstr":
        case "lpstr":
        case "bstr":
          value = {
            type: "string",
            value: text,
          };
          break;
        case "i4":
        case "int":
          value = {
            type: "int",
            value: parseInt(text, 10),
          };
          break;
        case "decimal":
          value = {
            type: "decimal",
            value: parseFloat(text),
          };
          break;
        case "bool":
          value = {
            type: "bool",
            value: text === "true" || text === "1",
          };
          break;
        case "filetime":
          value = {
            type: "date",
            value: new Date(text),
          };
          break;
        case "r8":
          value = {
            type: "r8",
            value: parseFloat(text),
          };
          break;
        case "i8":
          value = {
            type: "i8",
            value: BigInt(text),
          };
          break;
        case "error":
          value = {
            type: "error",
            value: text,
          };
          break;
        default:
          value = {
            type: "string",
            value: text,
          };
          break;
      }
      return {
        name,
        value,
      };
    })
    .filter(Boolean);
}
//#endregion
//#region node_modules/@node-projects/excelforge/dist/features/FormControlBuilderCommon.js
var VML_OBJ_TYPE = {
  button: "Button",
  checkBox: "Checkbox",
  comboBox: "Drop",
  listBox: "List",
  optionButton: "Radio",
  groupBox: "GBox",
  label: "Label",
  scrollBar: "Scroll",
  spinner: "Spin",
  dialog: "Dialog",
};
var CTRL_OBJ_TYPE = {
  ...VML_OBJ_TYPE,
  checkBox: "CheckBox",
  dialog: "Dialog",
};
var OBJ_TYPE_TO_CTRL = {
  ...Object.fromEntries(Object.entries(VML_OBJ_TYPE).map(([k, v]) => [v, k])),
  CheckBox: "checkBox",
};
var CHECKED_MAP = {
  unchecked: "0",
  checked: "1",
  mixed: "2",
};
var CHECKED_REV = {
  0: "unchecked",
  1: "checked",
  2: "mixed",
};
//#endregion
//#region node_modules/@node-projects/excelforge/dist/core/WorkbookReader.js
function parseRels(xml) {
  const map = /* @__PURE__ */ new Map();
  try {
    const root = parseXml(xml);
    for (const c of root.children)
      if (localName(c.tag) === "Relationship") {
        const entry = {
          type: c.attrs["Type"] ?? "",
          target: c.attrs["Target"] ?? "",
        };
        if (c.attrs["TargetMode"]) entry.targetMode = c.attrs["TargetMode"];
        map.set(c.attrs["Id"] ?? "", entry);
      }
  } catch {}
  return map;
}
function parseContentTypes(xml) {
  const map = /* @__PURE__ */ new Map();
  try {
    const root = parseXml(xml);
    for (const c of root.children)
      if (localName(c.tag) === "Override") {
        const part = c.attrs["PartName"] ?? "";
        const ct = c.attrs["ContentType"] ?? "";
        map.set(part.startsWith("/") ? part.slice(1) : part, ct);
      }
  } catch {}
  return map;
}
function parseStyles(xml) {
  const root = parseXml(xml);
  const numFmts = /* @__PURE__ */ new Map();
  const xfs = [];
  const builtinFmts = {
    0: "General",
    1: "0",
    2: "0.00",
    3: "#,##0",
    4: "#,##0.00",
    9: "0%",
    10: "0.00%",
    11: "0.00E+00",
    12: "# ?/?",
    13: "# ??/??",
    14: "mm-dd-yy",
    15: "d-mmm-yy",
    16: "d-mmm",
    17: "mmm-yy",
    18: "h:mm AM/PM",
    19: "h:mm:ss AM/PM",
    20: "h:mm",
    21: "h:mm:ss",
    22: "m/d/yy h:mm",
    37: "#,##0 ;(#,##0)",
    38: "#,##0 ;[Red](#,##0)",
    39: "#,##0.00;(#,##0.00)",
    40: "#,##0.00;[Red](#,##0.00)",
    45: "mm:ss",
    46: "[h]:mm:ss",
    47: "mmss.0",
    48: "##0.0E+0",
    49: "@",
  };
  const numFmtNode = find(root, "numFmts");
  if (numFmtNode)
    for (const n of children(numFmtNode, "numFmt")) {
      const id = parseInt(n.attrs["numFmtId"] ?? "0", 10);
      const code = n.attrs["formatCode"] ?? "";
      numFmts.set(id, code);
    }
  const fontsNode = find(root, "fonts");
  const fonts = [];
  if (fontsNode) for (const fn of children(fontsNode, "font")) fonts.push(parseFont(fn));
  const fillsNode = find(root, "fills");
  const fills = [];
  if (fillsNode) for (const fn of children(fillsNode, "fill")) fills.push(parseFill(fn));
  const bordersNode = find(root, "borders");
  const borders = [];
  if (bordersNode) for (const bn of children(bordersNode, "border")) borders.push(parseBorder(bn));
  const cellXfsNode = find(root, "cellXfs");
  if (cellXfsNode)
    for (const xf of children(cellXfsNode, "xf")) {
      const fontId = parseInt(xf.attrs["fontId"] ?? "0", 10);
      const fillId = parseInt(xf.attrs["fillId"] ?? "0", 10);
      const borderId = parseInt(xf.attrs["borderId"] ?? "0", 10);
      const numFmtId = parseInt(xf.attrs["numFmtId"] ?? "0", 10);
      const applyFont = xf.attrs["applyFont"] === "1";
      const applyFill = xf.attrs["applyFill"] === "1";
      const applyBorder = xf.attrs["applyBorder"] === "1";
      const applyAlignment = xf.attrs["applyAlignment"] === "1";
      const applyNumFmt = xf.attrs["applyNumberFormat"] === "1";
      const applyProt = xf.attrs["applyProtection"] === "1";
      const style = {};
      if (applyFont && fonts[fontId]) style.font = fonts[fontId];
      if (applyFill && fills[fillId]) style.fill = fills[fillId];
      if (applyBorder && borders[borderId]) style.border = borders[borderId];
      if (applyNumFmt && numFmtId !== 0) {
        if (numFmts.has(numFmtId)) style.numberFormat = { formatCode: numFmts.get(numFmtId) };
        else if (builtinFmts[numFmtId]) style.numFmtId = numFmtId;
      }
      const alignNode = child(xf, "alignment");
      if (applyAlignment && alignNode) style.alignment = parseAlignment(alignNode);
      const protNode = child(xf, "protection");
      if (applyProt && protNode) {
        if (protNode.attrs["locked"] !== void 0) style.locked = protNode.attrs["locked"] !== "0";
        if (protNode.attrs["hidden"] !== void 0) style.hidden = protNode.attrs["hidden"] !== "0";
      }
      xfs.push(style);
    }
  const dxfs = [];
  const dxfsNode = find(root, "dxfs");
  if (dxfsNode)
    for (const dxf of children(dxfsNode, "dxf")) {
      const style = {};
      const fontNode = child(dxf, "font");
      if (fontNode) style.font = parseFont(fontNode);
      const fillNode = child(dxf, "fill");
      if (fillNode) style.fill = parseFill(fillNode);
      const borderNode = child(dxf, "border");
      if (borderNode) style.border = parseBorder(borderNode);
      const numFmtDxf = child(dxf, "numFmt");
      if (numFmtDxf) {
        const code = numFmtDxf.attrs["formatCode"] ?? "";
        if (code) style.numberFormat = { formatCode: code };
      }
      const alignDxf = child(dxf, "alignment");
      if (alignDxf) style.alignment = parseAlignment(alignDxf);
      dxfs.push(style);
    }
  return {
    xfs,
    numFmts,
    dxfs,
  };
}
function find(node, localTag) {
  if (localName(node.tag) === localTag) return node;
  for (const c of node.children) {
    const r = find(c, localTag);
    if (r) return r;
  }
}
function parseOoxmlColor(node) {
  if (!node) return void 0;
  if (node.attrs["rgb"]) return node.attrs["rgb"];
  if (node.attrs["theme"])
    return `theme:${node.attrs["theme"]}${node.attrs["tint"] ? ":tint:" + node.attrs["tint"] : ""}`;
}
function parseFont(node) {
  const f = {};
  for (const c of node.children)
    switch (localName(c.tag)) {
      case "b":
        f.bold = c.attrs["val"] !== "0";
        break;
      case "i":
        f.italic = c.attrs["val"] !== "0";
        break;
      case "strike":
        f.strike = c.attrs["val"] !== "0";
        break;
      case "u":
        f.underline = c.attrs["val"] ?? "single";
        break;
      case "sz":
        f.size = parseFloat(c.attrs["val"] ?? "11");
        break;
      case "name":
        f.name = c.attrs["val"];
        break;
      case "family":
        f.family = parseInt(c.attrs["val"] ?? "0", 10);
        break;
      case "scheme":
        f.scheme = c.attrs["val"];
        break;
      case "charset":
        f.charset = parseInt(c.attrs["val"] ?? "0", 10);
        break;
      case "vertAlign":
        f.vertAlign = c.attrs["val"];
        break;
      case "color": {
        const rgb = c.attrs["rgb"];
        const theme = c.attrs["theme"];
        if (rgb) f.color = rgb;
        else if (theme)
          f.color = `theme:${theme}${c.attrs["tint"] ? ":tint:" + c.attrs["tint"] : ""}`;
        break;
      }
    }
  return f;
}
function parseFill(node) {
  const pattern = child(node, "patternFill");
  if (pattern) {
    const fg = child(pattern, "fgColor");
    const bg = child(pattern, "bgColor");
    return {
      type: "pattern",
      pattern: pattern.attrs["patternType"],
      fgColor:
        fg?.attrs["rgb"] ??
        (fg?.attrs["theme"]
          ? `theme:${fg.attrs["theme"]}${fg.attrs["tint"] ? ":tint:" + fg.attrs["tint"] : ""}`
          : void 0),
      bgColor:
        bg?.attrs["rgb"] ??
        (bg?.attrs["theme"]
          ? `theme:${bg.attrs["theme"]}${bg.attrs["tint"] ? ":tint:" + bg.attrs["tint"] : ""}`
          : void 0),
    };
  }
  const gradient = child(node, "gradientFill");
  if (gradient) {
    const stops = children(gradient, "stop")
      .concat(children(gradient, "gradientStop"))
      .map((s) => {
        const colorNode = child(s, "color");
        return {
          position: parseFloat(s.attrs["position"] ?? "0"),
          color:
            colorNode?.attrs["rgb"] ??
            (colorNode?.attrs["theme"]
              ? `theme:${colorNode.attrs["theme"]}${colorNode.attrs["tint"] ? ":tint:" + colorNode.attrs["tint"] : ""}`
              : "FF000000"),
        };
      });
    return {
      type: "gradient",
      gradientType: gradient.attrs["type"],
      degree: gradient.attrs["degree"] ? parseFloat(gradient.attrs["degree"]) : void 0,
      stops,
    };
  }
  return {
    type: "pattern",
    pattern: "none",
  };
}
function parseBorder(node) {
  const parseSide = (tag) => {
    const n = child(node, tag);
    if (!n) return void 0;
    const style = n.attrs["style"];
    const color = child(n, "color");
    if (!style && !color) return void 0;
    return {
      style,
      color:
        color?.attrs["rgb"] ??
        (color?.attrs["theme"]
          ? `theme:${color.attrs["theme"]}${color.attrs["tint"] ? ":tint:" + color.attrs["tint"] : ""}`
          : void 0),
    };
  };
  return {
    left: parseSide("left"),
    right: parseSide("right"),
    top: parseSide("top"),
    bottom: parseSide("bottom"),
    diagonal: parseSide("diagonal"),
    diagonalUp: node.attrs["diagonalUp"] === "1",
    diagonalDown: node.attrs["diagonalDown"] === "1",
  };
}
function parseAlignment(node) {
  const a = {};
  if (node.attrs["horizontal"]) a.horizontal = node.attrs["horizontal"];
  if (node.attrs["vertical"]) a.vertical = node.attrs["vertical"];
  if (node.attrs["wrapText"]) a.wrapText = node.attrs["wrapText"] !== "0";
  if (node.attrs["shrinkToFit"]) a.shrinkToFit = node.attrs["shrinkToFit"] !== "0";
  if (node.attrs["textRotation"]) a.textRotation = parseInt(node.attrs["textRotation"], 10);
  if (node.attrs["indent"]) a.indent = parseInt(node.attrs["indent"], 10);
  if (node.attrs["readingOrder"]) a.readingOrder = parseInt(node.attrs["readingOrder"], 10);
  return a;
}
function parseSharedStrings(xml) {
  return children(parseXml(xml), "si").map((si) => {
    const t = child(si, "t");
    const runs = children(si, "r");
    if (t && !runs.length) return { text: t.text ?? "" };
    const richRuns = runs.map((r) => {
      const runText = child(r, "t")?.text ?? "";
      const rPr = child(r, "rPr");
      const run = { text: runText };
      if (rPr) run.font = parseFont(rPr);
      return run;
    });
    const plainText = richRuns.map((r) => r.text).join("");
    if (richRuns.some((r) => r.font && Object.keys(r.font).length > 0))
      return {
        text: plainText,
        richText: richRuns,
      };
    return { text: plainText };
  });
}
function parseWorksheet(xml, name, styles, sharedStrings) {
  const ws = new Worksheet(name);
  const root = parseXml(xml);
  const unknownParts = [];
  const tableRIds = [];
  let legacyDrawingRId = "";
  let drawingRId = "";
  const ctrlPropRIds = [];
  const vmCells = /* @__PURE__ */ new Map();
  const knownTags = /* @__PURE__ */ new Set([
    "sheetPr",
    "dimension",
    "sheetViews",
    "sheetFormatPr",
    "cols",
    "sheetData",
    "mergeCells",
    "conditionalFormatting",
    "dataValidations",
    "sheetProtection",
    "printOptions",
    "pageMargins",
    "pageSetup",
    "headerFooter",
    "drawing",
    "tableParts",
    "autoFilter",
    "rowBreaks",
    "colBreaks",
    "picture",
    "oleObjects",
    "ctrlProps",
    "legacyDrawing",
    "AlternateContent",
    "extLst",
  ]);
  for (const node of root.children) {
    const tag = localName(node.tag);
    switch (tag) {
      case "sheetViews":
        parseSheetViews(node, ws);
        break;
      case "cols":
        parseCols(node, ws, styles);
        break;
      case "sheetData":
        parseSheetData(node, ws, styles, sharedStrings, vmCells);
        break;
      case "mergeCells":
        parseMerges(node, ws);
        break;
      case "autoFilter":
        ws.autoFilter = { ref: node.attrs["ref"] ?? "" };
        break;
      case "tableParts":
        for (const tp of children(node, "tablePart")) {
          const rid = tp.attrs["r:id"] ?? "";
          if (rid) tableRIds.push(rid);
        }
        break;
      case "sheetProtection":
        parseProtection(node, ws);
        break;
      case "pageMargins":
        parsePageMargins(node, ws);
        break;
      case "pageSetup":
        parsePageSetup(node, ws);
        break;
      case "headerFooter":
        parseHeaderFooter(node, ws);
        break;
      case "printOptions":
        parsePrintOptions(node, ws);
        break;
      case "conditionalFormatting":
        parseConditionalFormatting(node, ws, styles);
        break;
      case "dataValidations":
        parseDataValidations(node, ws);
        break;
      case "rowBreaks":
        for (const brk of children(node, "brk")) {
          const id = parseInt(brk.attrs["id"] ?? "0", 10);
          if (id > 0) ws.addRowBreak(id, brk.attrs["man"] === "1");
        }
        break;
      case "colBreaks":
        for (const brk of children(node, "brk")) {
          const id = parseInt(brk.attrs["id"] ?? "0", 10);
          if (id > 0) ws.addColBreak(id, brk.attrs["man"] === "1");
        }
        break;
      case "drawing":
        drawingRId = node.attrs["r:id"] ?? "";
        break;
      case "legacyDrawing":
        legacyDrawingRId = node.attrs["r:id"] ?? "";
        break;
      case "AlternateContent": {
        const choiceNode = node.children.find(
          (c) => typeof c !== "string" && localName(c.tag) === "Choice",
        );
        const controlsNode = choiceNode
          ? choiceNode.children.find(
              (c) => typeof c !== "string" && localName(c.tag) === "controls",
            )
          : void 0;
        if (controlsNode)
          for (const acNode of controlsNode.children) {
            if (typeof acNode === "string") continue;
            let ctrlNode;
            if (localName(acNode.tag) === "control") ctrlNode = acNode;
            else if (localName(acNode.tag) === "AlternateContent")
              ctrlNode = acNode.children
                .find((c) => typeof c !== "string" && localName(c.tag) === "Choice")
                ?.children.find((c) => typeof c !== "string" && localName(c.tag) === "control");
            if (!ctrlNode) continue;
            const controlPr = ctrlNode.children.find(
              (c) => typeof c !== "string" && localName(c.tag) === "controlPr",
            );
            const ctrlRId = ctrlNode.attrs["r:id"] ?? controlPr?.attrs["r:id"] ?? "";
            if (ctrlRId) ctrlPropRIds.push(ctrlRId);
          }
        break;
      }
      case "extLst":
        parseSparklineExtensions(node, ws);
        break;
      default:
        if (!knownTags.has(tag)) unknownParts.push(nodeToXml(node));
        break;
    }
  }
  return {
    ws,
    originalXml: xml,
    unknownParts,
    tableRIds,
    legacyDrawingRId,
    ctrlPropRIds,
    drawingRId,
    vmCells,
  };
}
function parseSheetViews(node, ws) {
  const sv = children(node, "sheetView")[0];
  if (!sv) return;
  ws.view = {
    showGridLines: sv.attrs["showGridLines"] !== "0",
    showRowColHeaders: sv.attrs["showRowColHeaders"] !== "0",
    zoomScale: sv.attrs["zoomScale"] ? parseInt(sv.attrs["zoomScale"], 10) : void 0,
    rightToLeft: sv.attrs["rightToLeft"] === "1",
    tabSelected: sv.attrs["tabSelected"] === "1",
    view: sv.attrs["view"],
  };
  const pane = child(sv, "pane");
  if (pane && pane.attrs["state"] === "frozen")
    ws.freezePane = {
      col: pane.attrs["xSplit"] ? parseInt(pane.attrs["xSplit"], 10) : void 0,
      row: pane.attrs["ySplit"] ? parseInt(pane.attrs["ySplit"], 10) : void 0,
    };
}
function parseCols(node, ws, styles) {
  for (const col of children(node, "col")) {
    const min = parseInt(col.attrs["min"] ?? "1", 10);
    const max = parseInt(col.attrs["max"] ?? "1", 10);
    const def = {
      width: col.attrs["width"] ? parseFloat(col.attrs["width"]) : void 0,
      hidden: col.attrs["hidden"] === "1",
      customWidth: col.attrs["customWidth"] === "1",
      outlineLevel: col.attrs["outlineLevel"] ? parseInt(col.attrs["outlineLevel"], 10) : void 0,
      style: col.attrs["style"] ? styles.xfs[parseInt(col.attrs["style"], 10)] : void 0,
    };
    for (let c = min; c <= max; c++) ws.setColumn(c, def);
  }
}
function parseSheetData(node, ws, styles, sharedStrings, vmCells) {
  for (const rowNode of children(node, "row")) {
    const rowIdx = parseInt(rowNode.attrs["r"] ?? "0", 10);
    if (!rowIdx) continue;
    const rowDef = {};
    if (rowNode.attrs["ht"]) rowDef.height = parseFloat(rowNode.attrs["ht"]);
    if (rowNode.attrs["hidden"]) rowDef.hidden = rowNode.attrs["hidden"] === "1";
    if (rowNode.attrs["outlineLevel"])
      rowDef.outlineLevel = parseInt(rowNode.attrs["outlineLevel"], 10);
    if (rowNode.attrs["collapsed"]) rowDef.collapsed = rowNode.attrs["collapsed"] === "1";
    if (rowNode.attrs["s"]) rowDef.style = styles.xfs[parseInt(rowNode.attrs["s"], 10)];
    if (Object.keys(rowDef).length) ws.setRow(rowIdx, rowDef);
    for (const cNode of children(rowNode, "c")) {
      const ref = cNode.attrs["r"] ?? "";
      if (!ref) continue;
      const { row, col } = cellRefToIndices(ref);
      const styleIdx = cNode.attrs["s"] ? parseInt(cNode.attrs["s"], 10) : 0;
      const cellStyle = styleIdx > 0 ? styles.xfs[styleIdx] : void 0;
      const t = cNode.attrs["t"] ?? "";
      const fNode = child(cNode, "f");
      const vNode = child(cNode, "v");
      const cell = {};
      if (cellStyle) cell.style = cellStyle;
      const vmAttr = cNode.attrs["vm"];
      if (vmAttr && vmCells) vmCells.set(ref, parseInt(vmAttr, 10));
      if (fNode)
        if (fNode.attrs["t"] === "array") cell.arrayFormula = fNode.text ?? "";
        else cell.formula = fNode.text ?? "";
      else if (vNode) {
        const raw = vNode.text ?? "";
        switch (t) {
          case "s": {
            const ss = sharedStrings[parseInt(raw, 10)];
            if (ss) {
              cell.value = ss.text;
              if (ss.richText) cell.richText = ss.richText;
            } else cell.value = "";
            break;
          }
          case "b":
            cell.value = raw === "1" || raw === "true";
            break;
          case "str":
          case "inlineStr": {
            const is = child(cNode, "is");
            cell.value = is ? (child(is, "t")?.text ?? raw) : raw;
            break;
          }
          case "e":
            cell.value = new CellError(raw);
            break;
          default: {
            const n = parseFloat(raw);
            cell.value = isNaN(n) ? raw : n;
            break;
          }
        }
      }
      if (Object.keys(cell).length || cell.value !== void 0) ws.setCell(row, col, cell);
    }
  }
}
function parseMerges(node, ws) {
  for (const m of children(node, "mergeCell")) {
    const ref = m.attrs["ref"] ?? "";
    if (ref.includes(":")) ws.mergeByRef(ref);
  }
}
function parseProtection(node, ws) {
  ws.protection = {
    sheet: node.attrs["sheet"] !== "0",
    password: void 0,
    selectLockedCells: node.attrs["selectLockedCells"] !== "0",
    selectUnlockedCells: node.attrs["selectUnlockedCells"] !== "0",
    formatCells: node.attrs["formatCells"] === "0",
    formatColumns: node.attrs["formatColumns"] === "0",
    formatRows: node.attrs["formatRows"] === "0",
    insertColumns: node.attrs["insertColumns"] === "0",
    insertRows: node.attrs["insertRows"] === "0",
    insertHyperlinks: node.attrs["insertHyperlinks"] === "0",
    deleteColumns: node.attrs["deleteColumns"] === "0",
    deleteRows: node.attrs["deleteRows"] === "0",
    sort: node.attrs["sort"] === "0",
    autoFilter: node.attrs["autoFilter"] === "0",
    pivotTables: node.attrs["pivotTables"] === "0",
  };
}
function parsePageMargins(node, ws) {
  ws.pageMargins = {
    left: parseFloat(node.attrs["left"] ?? "0.7"),
    right: parseFloat(node.attrs["right"] ?? "0.7"),
    top: parseFloat(node.attrs["top"] ?? "0.75"),
    bottom: parseFloat(node.attrs["bottom"] ?? "0.75"),
    header: parseFloat(node.attrs["header"] ?? "0.3"),
    footer: parseFloat(node.attrs["footer"] ?? "0.3"),
  };
}
function parsePageSetup(node, ws) {
  ws.pageSetup = {
    paperSize: node.attrs["paperSize"] ? parseInt(node.attrs["paperSize"], 10) : void 0,
    orientation: node.attrs["orientation"],
    fitToPage: node.attrs["fitToPage"] === "1",
    fitToWidth: node.attrs["fitToWidth"] ? parseInt(node.attrs["fitToWidth"], 10) : void 0,
    fitToHeight: node.attrs["fitToHeight"] ? parseInt(node.attrs["fitToHeight"], 10) : void 0,
    scale: node.attrs["scale"] ? parseInt(node.attrs["scale"], 10) : void 0,
    horizontalDpi: node.attrs["horizontalDpi"] ? parseInt(node.attrs["horizontalDpi"], 10) : void 0,
    verticalDpi: node.attrs["verticalDpi"] ? parseInt(node.attrs["verticalDpi"], 10) : void 0,
  };
}
function parseHeaderFooter(node, ws) {
  ws.headerFooter = {
    oddHeader: child(node, "oddHeader")?.text,
    oddFooter: child(node, "oddFooter")?.text,
    evenHeader: child(node, "evenHeader")?.text,
    evenFooter: child(node, "evenFooter")?.text,
    firstHeader: child(node, "firstHeader")?.text,
    firstFooter: child(node, "firstFooter")?.text,
    differentOddEven: node.attrs["differentOddEven"] === "1",
    differentFirst: node.attrs["differentFirst"] === "1",
  };
}
function parsePrintOptions(node, ws) {
  ws.printOptions = {
    gridLines: node.attrs["gridLines"] === "1",
    gridLinesSet: node.attrs["gridLinesSet"] === "1",
    headings: node.attrs["headings"] === "1",
    centerHorizontal: node.attrs["horizontalCentered"] === "1",
    centerVertical: node.attrs["verticalCentered"] === "1",
  };
}
function parseConditionalFormatting(node, ws, styles) {
  const sqref = node.attrs["sqref"] ?? "";
  for (const rule of children(node, "cfRule")) {
    const cf = {
      sqref,
      type: rule.attrs["type"] ?? "cellIs",
    };
    if (rule.attrs["operator"]) cf.operator = rule.attrs["operator"];
    if (rule.attrs["priority"]) cf.priority = parseInt(rule.attrs["priority"], 10);
    if (rule.attrs["text"]) cf.text = rule.attrs["text"];
    if (rule.attrs["aboveAverage"] === "0") cf.aboveAverage = false;
    if (rule.attrs["percent"] === "1") cf.percent = true;
    if (rule.attrs["rank"]) cf.rank = parseInt(rule.attrs["rank"], 10);
    if (rule.attrs["timePeriod"]) cf.timePeriod = rule.attrs["timePeriod"];
    if (rule.attrs["dxfId"] !== void 0) {
      const dxfId = parseInt(rule.attrs["dxfId"], 10);
      if (styles.dxfs[dxfId]) cf.style = styles.dxfs[dxfId];
    }
    const formulas = children(rule, "formula");
    if (formulas[0]?.text) cf.formula = formulas[0].text;
    if (formulas[1]?.text) cf.formula2 = formulas[1].text;
    const csNode = child(rule, "colorScale");
    if (csNode)
      cf.colorScale = {
        type: "colorScale",
        cfvo: children(csNode, "cfvo").map((c) => ({
          type: c.attrs["type"] ?? "min",
          val: c.attrs["val"],
        })),
        color: children(csNode, "color").map((c) => parseOoxmlColor(c) ?? ""),
      };
    const dbNode = child(rule, "dataBar");
    if (dbNode) {
      const cfvos = children(dbNode, "cfvo");
      const colorNode = child(dbNode, "color");
      cf.dataBar = {
        type: "dataBar",
        showValue: dbNode.attrs["showValue"] !== "0" ? void 0 : false,
        minType: cfvos[0]?.attrs["type"],
        minVal: cfvos[0]?.attrs["val"],
        maxType: cfvos[1]?.attrs["type"],
        maxVal: cfvos[1]?.attrs["val"],
        color: parseOoxmlColor(colorNode),
      };
    }
    const isNode = child(rule, "iconSet");
    if (isNode) {
      const cfvos = children(isNode, "cfvo").map((c) => ({
        type: c.attrs["type"] ?? "percent",
        val: c.attrs["val"],
      }));
      cf.iconSet = {
        type: "iconSet",
        iconSet: isNode.attrs["iconSet"] ?? "3TrafficLights1",
        cfvo: cfvos,
        showValue: isNode.attrs["showValue"] === "0" ? false : void 0,
        reverse: isNode.attrs["reverse"] === "1" ? true : void 0,
      };
    }
    ws.addConditionalFormat(cf);
  }
}
function parseDataValidations(node, ws) {
  for (const dv of children(node, "dataValidation")) {
    const sqref = dv.attrs["sqref"] ?? "";
    if (!sqref) continue;
    const type = dv.attrs["type"] ?? "whole";
    const val = { type };
    if (dv.attrs["operator"]) val.operator = dv.attrs["operator"];
    if (dv.attrs["allowBlank"] === "1") val.allowBlank = true;
    if (dv.attrs["showErrorMessage"] === "1") val.showErrorAlert = true;
    if (dv.attrs["errorTitle"]) val.errorTitle = dv.attrs["errorTitle"];
    if (dv.attrs["error"]) val.error = dv.attrs["error"];
    if (dv.attrs["showInputMessage"] === "1") val.showInputMessage = true;
    if (dv.attrs["promptTitle"]) val.promptTitle = dv.attrs["promptTitle"];
    if (dv.attrs["prompt"]) val.prompt = dv.attrs["prompt"];
    if (dv.attrs["showDropDown"] === "1") val.showDropDown = false;
    const f1 = child(dv, "formula1");
    const f2 = child(dv, "formula2");
    if (f1?.text)
      if (type === "list" && f1.text.startsWith('"') && f1.text.endsWith('"'))
        val.list = f1.text.slice(1, -1).split(",");
      else val.formula1 = f1.text;
    if (f2?.text) val.formula2 = f2.text;
    ws.addDataValidation(sqref, val);
  }
}
function parseTableXml(xml, styles) {
  try {
    const root = parseXml(xml);
    if (localName(root.tag) !== "table") return null;
    const name = root.attrs["name"] ?? "";
    const displayName = root.attrs["displayName"] ?? name;
    const ref = root.attrs["ref"] ?? "";
    const totalsCount = parseInt(root.attrs["totalsRowCount"] ?? "0", 10);
    const columns = [];
    const colsNode = find(root, "tableColumns");
    if (colsNode)
      for (const col of children(colsNode, "tableColumn")) {
        const tc = { name: col.attrs["name"] ?? "" };
        if (col.attrs["totalsRowFunction"]) tc.totalsRowFunction = col.attrs["totalsRowFunction"];
        if (col.attrs["totalsRowFormula"]) tc.totalsRowFormula = col.attrs["totalsRowFormula"];
        if (col.attrs["totalsRowLabel"]) tc.totalsRowLabel = col.attrs["totalsRowLabel"];
        if (col.attrs["filterButton"] === "0") tc.filterButton = false;
        const dataDxfId = col.attrs["dataDxfId"];
        if (dataDxfId !== void 0) {
          const parsedStyle = styles.dxfs[parseInt(dataDxfId, 10)];
          if (parsedStyle) tc.style = parsedStyle;
        }
        columns.push(tc);
      }
    const table = {
      name,
      ref,
      columns,
    };
    if (displayName && displayName !== name) table.displayName = displayName;
    if (totalsCount > 0) table.totalsRow = true;
    const styleNode = find(root, "tableStyleInfo");
    if (styleNode) {
      if (styleNode.attrs["name"]) table.style = styleNode.attrs["name"];
      if (styleNode.attrs["showFirstColumn"] === "1") table.showFirstColumn = true;
      if (styleNode.attrs["showLastColumn"] === "1") table.showLastColumn = true;
      if (styleNode.attrs["showRowStripes"] === "1") table.showRowStripes = true;
      if (styleNode.attrs["showColumnStripes"] === "1") table.showColumnStripes = true;
    }
    return table;
  } catch {
    return null;
  }
}
var EMU_PX = 9525;
function parseAnchorPos(node) {
  return {
    col: parseInt(child(node, "col")?.text ?? "0", 10),
    row: parseInt(child(node, "row")?.text ?? "0", 10),
    colOff: parseInt(child(node, "colOff")?.text ?? "0", 10),
    rowOff: parseInt(child(node, "rowOff")?.text ?? "0", 10),
  };
}
function extToFormat(ext) {
  return (
    {
      png: "png",
      jpg: "jpeg",
      jpeg: "jpeg",
      gif: "gif",
      emf: "emf",
      wmf: "wmf",
      tiff: "tiff",
      tif: "tiff",
      svg: "svg",
      ico: "ico",
      webp: "webp",
      bmp: "bmp",
    }[ext.toLowerCase()] ?? "png"
  );
}
function parseDrawingObjects(drawXml, drawRels, drawDir, getEntry, ws) {
  const root = parseXml(drawXml);
  for (const anchor of root.children) {
    if (typeof anchor === "string") continue;
    const tag = localName(anchor.tag);
    if (tag !== "twoCellAnchor" && tag !== "oneCellAnchor" && tag !== "absoluteAnchor") continue;
    const anchorFrom = parseDrawingAnchorFrom(anchor, tag);
    const pic = findDescendant(anchor, "pic");
    if (pic) {
      parseDrawingImage(pic, anchor, tag, anchorFrom, drawRels, drawDir, getEntry, ws);
      continue;
    }
    const graphicFrame = findDescendant(anchor, "graphicFrame");
    if (graphicFrame) {
      parseDrawingChart(graphicFrame, anchor, tag, drawRels, drawDir, getEntry, ws);
      continue;
    }
    const oMath = findDescendant(anchor, "oMath");
    if (oMath) {
      parseDrawingMathEquation(oMath, anchor, tag, anchorFrom, ws);
      continue;
    }
    const sp = findDescendant(anchor, "sp");
    if (sp) {
      parseDrawingShape(sp, anchor, tag, ws);
      continue;
    }
  }
}
function parseDrawingChart(graphicFrame, anchor, tag, drawRels, drawDir, getEntry, ws) {
  const chartRef = findDescendant(graphicFrame, "chart");
  if (!chartRef || !drawRels) return;
  const rId = chartRef.attrs["r:id"] ?? attr(chartRef, "r:id") ?? "";
  if (!rId) return;
  const rel = drawRels.get(rId);
  if (!rel) return;
  const chartEntry = getEntry(
    rel.target.startsWith("/") ? rel.target.slice(1) : resolvePath(drawDir, rel.target),
  );
  if (!chartEntry) return;
  const chartRoot = parseXml(new TextDecoder().decode(chartEntry.data));
  let from = {
    col: 0,
    row: 0,
  };
  let to = {
    col: 8,
    row: 15,
  };
  if (tag === "twoCellAnchor") {
    const fromNode = child(anchor, "from");
    const toNode = child(anchor, "to");
    if (fromNode)
      from = {
        col: parseInt(child(fromNode, "col")?.text ?? "0", 10),
        row: parseInt(child(fromNode, "row")?.text ?? "0", 10),
      };
    if (toNode)
      to = {
        col: parseInt(child(toNode, "col")?.text ?? "0", 10),
        row: parseInt(child(toNode, "row")?.text ?? "0", 10),
      };
  }
  const plotArea = findDescendant(chartRoot, "plotArea");
  if (!plotArea) return;
  const chartTypeMap = {
    barChart: "column",
    bar3DChart: "column",
    lineChart: "line",
    line3DChart: "line",
    areaChart: "area",
    area3DChart: "area",
    pieChart: "pie",
    pie3DChart: "pie",
    doughnutChart: "doughnut",
    scatterChart: "scatter",
    bubbleChart: "bubble",
    radarChart: "radar",
    stockChart: "stock",
  };
  let chartType = "column";
  let chartGroupNode;
  for (const plotChild of plotArea.children) {
    if (typeof plotChild === "string") continue;
    const localTag = localName(plotChild.tag);
    if (chartTypeMap[localTag]) {
      chartType = chartTypeMap[localTag];
      chartGroupNode = plotChild;
      if (localTag === "barChart" || localTag === "bar3DChart") {
        if ((child(plotChild, "barDir")?.attrs["val"] ?? "") === "bar") chartType = "bar";
        const grpVal = child(plotChild, "grouping")?.attrs["val"] ?? "";
        if (grpVal === "stacked") chartType = chartType === "bar" ? "barStacked" : "columnStacked";
        if (grpVal === "percentStacked")
          chartType = chartType === "bar" ? "barStacked100" : "columnStacked100";
      }
      if (localTag === "lineChart") {
        const marker = findDescendant(plotChild, "marker");
        if (marker) {
          const markerVal = child(marker, "symbol")?.attrs["val"] ?? "";
          if (markerVal && markerVal !== "none") chartType = "lineMarker";
        }
      }
      if (localTag === "scatterChart") {
        if (child(plotChild, "scatterStyle")?.attrs["val"]?.includes("Smooth"))
          chartType = "scatterSmooth";
      }
      if (localTag === "radarChart") {
        if (child(plotChild, "radarStyle")?.attrs["val"] === "filled") chartType = "radarFilled";
      }
      break;
    }
  }
  if (!chartGroupNode) return;
  const series = [];
  for (const ser of children(chartGroupNode, "ser")) {
    const s = { values: "" };
    const tx = child(ser, "tx");
    if (tx) {
      const strRef = findDescendant(tx, "strRef");
      if (strRef) {
        const f = child(strRef, "f");
        if (f?.text) s.name = f.text;
        const strCache = child(strRef, "strCache");
        if (strCache) {
          const pt = findDescendant(strCache, "pt");
          const v = pt ? findDescendant(pt, "v") : void 0;
          if (v?.text) s.name = v.text;
        }
      }
      const v = findDescendant(tx, "v");
      if (v?.text && !s.name) s.name = v.text;
    }
    const val = child(ser, "val") || child(ser, "yVal");
    if (val) {
      const numRef = findDescendant(val, "numRef");
      if (numRef) {
        const f = child(numRef, "f");
        if (f?.text) s.values = f.text;
      }
    }
    const cat = child(ser, "cat") || child(ser, "xVal");
    if (cat) {
      const strRef = findDescendant(cat, "strRef");
      const numRef = findDescendant(cat, "numRef");
      const ref = strRef || numRef;
      if (ref) {
        const f = child(ref, "f");
        if (f?.text) s.categories = f.text;
      }
    }
    const spPr = child(ser, "spPr");
    if (spPr) {
      const solidFill = child(spPr, "solidFill");
      if (solidFill) {
        const srgb = child(solidFill, "srgbClr");
        if (srgb?.attrs["val"]) s.color = "#" + srgb.attrs["val"];
      }
    }
    if (s.values) series.push(s);
  }
  if (!series.length) return;
  let title;
  const titleNode = child(chartRoot, "title") || findDescendant(chartRoot, "title");
  if (titleNode) {
    const aT = findDescendant(titleNode, "t");
    if (aT?.text) title = aT.text;
  }
  let xAxis;
  let yAxis;
  const catAx = findDescendant(plotArea, "catAx") || findDescendant(plotArea, "dateAx");
  const valAx = findDescendant(plotArea, "valAx");
  if (catAx) {
    xAxis = {};
    const axTitle = findDescendant(catAx, "title");
    if (axTitle) {
      const t = findDescendant(axTitle, "t");
      if (t?.text) xAxis.title = t.text;
    }
  }
  if (valAx) {
    yAxis = {};
    const axTitle = findDescendant(valAx, "title");
    if (axTitle) {
      const t = findDescendant(axTitle, "t");
      if (t?.text) yAxis.title = t.text;
    }
    if (child(valAx, "majorGridlines")) yAxis.gridLines = true;
  }
  let legend = false;
  const legendNode = findDescendant(chartRoot, "legend");
  if (legendNode) legend = child(legendNode, "legendPos")?.attrs["val"] || true;
  const varyColorsVal = child(chartGroupNode, "varyColors")?.attrs["val"] !== "0";
  const chart = {
    type: chartType,
    series,
    from,
    to,
  };
  if (title) chart.title = title;
  if (xAxis) chart.xAxis = xAxis;
  if (yAxis) chart.yAxis = yAxis;
  if (legend !== false) chart.legend = legend;
  if (varyColorsVal) chart.varyColors = true;
  ws.addChart(chart);
}
function parseDrawingShape(sp, anchor, tag, ws) {
  let from = {
    col: 0,
    row: 0,
  };
  let to = {
    col: 2,
    row: 2,
  };
  if (tag === "twoCellAnchor") {
    const fromNode = child(anchor, "from");
    const toNode = child(anchor, "to");
    if (fromNode)
      from = {
        col: parseInt(child(fromNode, "col")?.text ?? "0", 10),
        row: parseInt(child(fromNode, "row")?.text ?? "0", 10),
      };
    if (toNode)
      to = {
        col: parseInt(child(toNode, "col")?.text ?? "0", 10),
        row: parseInt(child(toNode, "row")?.text ?? "0", 10),
      };
  }
  const spPr = findDescendant(sp, "spPr");
  let shapeType = "rect";
  let fillColor;
  let lineColor;
  let lineWidth;
  let rotation;
  if (spPr) {
    const prstGeom = findDescendant(spPr, "prstGeom");
    if (prstGeom) {
      const prst = prstGeom.attrs["prst"] ?? "";
      if (
        /* @__PURE__ */ new Set([
          "rect",
          "roundRect",
          "ellipse",
          "triangle",
          "diamond",
          "pentagon",
          "hexagon",
          "octagon",
          "star5",
          "star6",
          "rightArrow",
          "leftArrow",
          "upArrow",
          "downArrow",
          "line",
          "curvedConnector3",
          "callout1",
          "callout2",
          "cloud",
          "heart",
          "lightningBolt",
          "sun",
          "moon",
          "smileyFace",
          "flowChartProcess",
          "flowChartDecision",
          "flowChartTerminator",
          "flowChartDocument",
        ]).has(prst)
      )
        shapeType = prst;
    }
    const solidFill = child(spPr, "solidFill");
    if (solidFill) {
      const srgb = child(solidFill, "srgbClr");
      if (srgb?.attrs["val"]) fillColor = "#" + srgb.attrs["val"];
      const schemeClr = child(solidFill, "schemeClr");
      if (schemeClr?.attrs["val"]) fillColor = "theme:" + schemeClr.attrs["val"];
    }
    const ln = child(spPr, "ln");
    if (ln) {
      const w = ln.attrs["w"];
      if (w) lineWidth = Math.round(parseInt(w, 10) / EMU_PX);
      const lineFill = child(ln, "solidFill");
      if (lineFill) {
        const srgb = child(lineFill, "srgbClr");
        if (srgb?.attrs["val"]) lineColor = "#" + srgb.attrs["val"];
      }
    }
    const xfrm = findDescendant(spPr, "xfrm");
    if (xfrm?.attrs["rot"]) rotation = parseInt(xfrm.attrs["rot"], 10) / 6e4;
  }
  let text;
  let font;
  const txBody = findDescendant(sp, "txBody");
  if (txBody) {
    const bodyPr = child(txBody, "bodyPr");
    const wordArtPreset =
      (bodyPr ? child(bodyPr, "prstTxWarp") : void 0)?.attrs["prst"] ??
      bodyPr?.attrs["prstTxWarp"] ??
      "";
    const textParts = [];
    for (const p of children(txBody, "p"))
      for (const r of children(p, "r")) {
        const t = child(r, "t");
        if (t?.text) textParts.push(t.text);
        if (!font) {
          const rPr = child(r, "rPr");
          if (rPr) {
            font = {};
            if (rPr.attrs["sz"]) font.size = parseInt(rPr.attrs["sz"], 10) / 100;
            if (rPr.attrs["b"] === "1") font.bold = true;
            if (rPr.attrs["i"] === "1") font.italic = true;
            const latin = findDescendant(rPr, "latin");
            if (latin?.attrs["typeface"]) font.name = latin.attrs["typeface"];
            const solidFill = child(rPr, "solidFill");
            if (solidFill) {
              const srgb = child(solidFill, "srgbClr");
              if (srgb?.attrs["val"]) font.color = "#" + srgb.attrs["val"];
            }
          }
        }
      }
    text = textParts.join("");
    if (
      wordArtPreset &&
      /* @__PURE__ */ new Set([
        "textPlain",
        "textArchUp",
        "textArchDown",
        "textCircle",
        "textWave1",
        "textWave2",
        "textInflate",
        "textDeflate",
        "textFadeUp",
        "textFadeDown",
        "textFadeLeft",
        "textFadeRight",
        "textSlantUp",
        "textSlantDown",
        "textCascadeUp",
        "textCascadeDown",
        "textChevron",
        "textRingInside",
        "textRingOutside",
        "textStop",
      ]).has(wordArtPreset) &&
      text
    ) {
      const wa = {
        text,
        from,
        to,
      };
      if (wordArtPreset !== "textPlain") wa.preset = wordArtPreset;
      if (font) wa.font = font;
      if (fillColor) wa.fillColor = fillColor;
      if (lineColor) wa.outlineColor = lineColor;
      ws.addWordArt(wa);
      return;
    }
  }
  const shape = {
    type: shapeType,
    from,
    to,
  };
  if (text) shape.text = text;
  if (fillColor) shape.fillColor = fillColor;
  if (lineColor) shape.lineColor = lineColor;
  if (lineWidth) shape.lineWidth = lineWidth;
  if (font) shape.font = font;
  if (rotation) shape.rotation = rotation;
  ws.addShape(shape);
}
function parseSparklineExtensions(extLst, ws) {
  for (const ext of children(extLst, "ext")) {
    const sparklineGroups = ext.children?.filter(
      (c) => typeof c !== "string" && localName(c.tag) === "sparklineGroups",
    );
    if (!sparklineGroups?.length) continue;
    for (const groups of sparklineGroups)
      for (const group of children(groups, "sparklineGroup")) {
        const typeAttr = group.attrs["type"] ?? "line";
        const spkType = typeAttr === "column" ? "bar" : typeAttr === "stacked" ? "stacked" : "line";
        const colorSeries = group.children?.find(
          (c) => typeof c !== "string" && localName(c.tag) === "colorSeries",
        );
        const colorHigh = group.children?.find(
          (c) => typeof c !== "string" && localName(c.tag) === "colorHigh",
        );
        const colorLow = group.children?.find(
          (c) => typeof c !== "string" && localName(c.tag) === "colorLow",
        );
        const colorFirst = group.children?.find(
          (c) => typeof c !== "string" && localName(c.tag) === "colorFirst",
        );
        const colorLast = group.children?.find(
          (c) => typeof c !== "string" && localName(c.tag) === "colorLast",
        );
        const colorNegative = group.children?.find(
          (c) => typeof c !== "string" && localName(c.tag) === "colorNegative",
        );
        const colorMarkers = group.children?.find(
          (c) => typeof c !== "string" && localName(c.tag) === "colorMarkers",
        );
        const showHigh = group.attrs["high"] === "1";
        const showLow = group.attrs["low"] === "1";
        const showFirst = group.attrs["first"] === "1";
        const showLast = group.attrs["last"] === "1";
        const showNegative = group.attrs["negative"] === "1";
        const showMarkers = group.attrs["markers"] === "1";
        const lwAttr = group.attrs["lineWeight"];
        const lineWidth = lwAttr ? parseFloat(lwAttr) : void 0;
        const minAxisType =
          group.attrs["minAxisType"] === "individual"
            ? "individual"
            : group.attrs["minAxisType"] === "custom"
              ? "custom"
              : void 0;
        const maxAxisType =
          group.attrs["maxAxisType"] === "individual"
            ? "individual"
            : group.attrs["maxAxisType"] === "custom"
              ? "custom"
              : void 0;
        const sparklinesNode = group.children?.find(
          (c) => typeof c !== "string" && localName(c.tag) === "sparklines",
        );
        if (!sparklinesNode) continue;
        const sparklineNodes = sparklinesNode.children?.filter(
          (c) => typeof c !== "string" && localName(c.tag) === "sparkline",
        );
        if (!sparklineNodes) continue;
        for (const spk of sparklineNodes) {
          const fNode = spk.children?.find(
            (c) => typeof c !== "string" && localName(c.tag) === "f",
          );
          const sqrefNode = spk.children?.find(
            (c) => typeof c !== "string" && localName(c.tag) === "sqref",
          );
          const dataRange = fNode?.text ?? "";
          const location = sqrefNode?.text ?? "";
          if (!dataRange || !location) continue;
          const sp = {
            type: spkType,
            dataRange,
            location,
          };
          if (colorSeries) {
            const rgb = colorSeries.attrs["rgb"];
            if (rgb) sp.color = "#" + rgb.slice(2);
            const theme = colorSeries.attrs["theme"];
            if (theme) sp.color = "theme:" + theme;
          }
          if (colorHigh) {
            const rgb = colorHigh.attrs["rgb"];
            if (rgb) sp.highColor = "#" + rgb.slice(2);
            const theme = colorHigh.attrs["theme"];
            if (theme) sp.highColor = "theme:" + theme;
          }
          if (colorLow) {
            const rgb = colorLow.attrs["rgb"];
            if (rgb) sp.lowColor = "#" + rgb.slice(2);
            const theme = colorLow.attrs["theme"];
            if (theme) sp.lowColor = "theme:" + theme;
          }
          if (colorFirst) {
            const rgb = colorFirst.attrs["rgb"];
            if (rgb) sp.firstColor = "#" + rgb.slice(2);
            const theme = colorFirst.attrs["theme"];
            if (theme) sp.firstColor = "theme:" + theme;
          }
          if (colorLast) {
            const rgb = colorLast.attrs["rgb"];
            if (rgb) sp.lastColor = "#" + rgb.slice(2);
            const theme = colorLast.attrs["theme"];
            if (theme) sp.lastColor = "theme:" + theme;
          }
          if (colorNegative) {
            const rgb = colorNegative.attrs["rgb"];
            if (rgb) sp.negativeColor = "#" + rgb.slice(2);
            const theme = colorNegative.attrs["theme"];
            if (theme) sp.negativeColor = "theme:" + theme;
          }
          if (colorMarkers) {
            const rgb = colorMarkers.attrs["rgb"];
            if (rgb) sp.markersColor = "#" + rgb.slice(2);
            const theme = colorMarkers.attrs["theme"];
            if (theme) sp.markersColor = "theme:" + theme;
          }
          if (showHigh) sp.showHigh = true;
          if (showLow) sp.showLow = true;
          if (showFirst) sp.showFirst = true;
          if (showLast) sp.showLast = true;
          if (showNegative) sp.showNegative = true;
          if (showMarkers) sp.showMarkers = true;
          if (lineWidth !== void 0) sp.lineWidth = lineWidth;
          if (minAxisType) sp.minAxisType = minAxisType;
          if (maxAxisType) sp.maxAxisType = maxAxisType;
          ws.addSparkline(sp);
        }
      }
  }
}
function parseDrawingAnchorFrom(anchor, tag) {
  const result = {};
  if (tag === "twoCellAnchor" || tag === "oneCellAnchor") {
    const fromNode = child(anchor, "from");
    if (fromNode)
      result.from = {
        col: parseInt(child(fromNode, "col")?.text ?? "0", 10),
        row: parseInt(child(fromNode, "row")?.text ?? "0", 10),
        colOff: parseInt(child(fromNode, "colOff")?.text ?? "0", 10),
        rowOff: parseInt(child(fromNode, "rowOff")?.text ?? "0", 10),
      };
  }
  if (tag === "oneCellAnchor") {
    const ext2 = findDescendant(anchor, "ext");
    if (ext2) {
      result.width = parseInt(ext2.attrs["cx"] ?? "0", 10);
      result.height = parseInt(ext2.attrs["cy"] ?? "0", 10);
    }
  }
  return result;
}
function parseDrawingImage(pic, anchor, tag, anchorInfo, drawRels, drawDir, getEntry, ws) {
  const blipFill = findDescendant(pic, "blipFill");
  const blip = blipFill ? findDescendant(blipFill, "blip") : void 0;
  const embedId = blip?.attrs["r:embed"] ?? blip?.attrs["embed"] ?? "";
  if (!embedId || !drawRels) return;
  const imgRel = drawRels.get(embedId);
  if (!imgRel) return;
  const imgPath = imgRel.target.startsWith("/")
    ? imgRel.target.slice(1)
    : resolvePath(drawDir, imgRel.target);
  const imgEntry = getEntry(imgPath);
  if (!imgEntry) return;
  const format = extToFormat(imgPath.split(".").pop() ?? "png");
  const altText = findDescendant(pic, "cNvPr")?.attrs["descr"] || void 0;
  const img = {
    data: imgEntry.data,
    format,
  };
  if (altText) img.altText = altText;
  if (tag === "twoCellAnchor") {
    const fromNode = child(anchor, "from");
    const toNode = child(anchor, "to");
    if (fromNode) img.from = parseAnchorPos(fromNode);
    if (toNode) img.to = parseAnchorPos(toNode);
  } else if (tag === "oneCellAnchor") {
    if (anchorInfo.from) img.from = anchorInfo.from;
    if (anchorInfo.width) img.width = Math.round(anchorInfo.width / EMU_PX);
    if (anchorInfo.height) img.height = Math.round(anchorInfo.height / EMU_PX);
  } else {
    const posNode = child(anchor, "pos");
    const ext2 = child(anchor, "ext");
    if (posNode)
      img.position = {
        x: Math.round(parseInt(posNode.attrs["x"] ?? "0", 10) / EMU_PX),
        y: Math.round(parseInt(posNode.attrs["y"] ?? "0", 10) / EMU_PX),
      };
    if (ext2) {
      img.width = Math.round(parseInt(ext2.attrs["cx"] ?? "0", 10) / EMU_PX);
      img.height = Math.round(parseInt(ext2.attrs["cy"] ?? "0", 10) / EMU_PX);
    }
  }
  ws.addImage(img);
}
function parseDrawingMathEquation(oMath, anchor, tag, anchorInfo, ws) {
  const elements = parseOmmlChildren(oMath);
  if (!elements.length) return;
  const eq = {
    elements,
    from: anchorInfo.from ?? {
      col: 0,
      row: 0,
    },
  };
  if (anchorInfo.width) eq.width = anchorInfo.width;
  if (anchorInfo.height) eq.height = anchorInfo.height;
  const firstRPr = findDescendant(oMath, "rPr");
  if (firstRPr) {
    const szAttr = firstRPr.attrs["sz"];
    if (szAttr) eq.fontSize = parseInt(szAttr, 10) / 100;
    const latin = findDescendant(firstRPr, "latin");
    if (latin?.attrs["typeface"]) eq.fontName = latin.attrs["typeface"];
  }
  ws.addMathEquation(eq);
}
function parseOmmlChildren(node) {
  const elements = [];
  for (const c of node.children) {
    if (typeof c === "string") continue;
    const el = parseOmmlElement(c);
    if (el) elements.push(el);
  }
  return elements;
}
function parseOmmlElement(node) {
  switch (localName(node.tag)) {
    case "r":
      return {
        type: "text",
        text: findDescendant(node, "t")?.text ?? "",
      };
    case "f": {
      const num = child(node, "num");
      const den = child(node, "den");
      return {
        type: "frac",
        base: num ? parseOmmlChildren(num) : [],
        argument: den ? parseOmmlChildren(den) : [],
      };
    }
    case "sSup": {
      const e = child(node, "e");
      const sup = child(node, "sup");
      return {
        type: "sup",
        base: e ? parseOmmlChildren(e) : [],
        argument: sup ? parseOmmlChildren(sup) : [],
      };
    }
    case "sSub": {
      const e = child(node, "e");
      const sub = child(node, "sub");
      return {
        type: "sub",
        base: e ? parseOmmlChildren(e) : [],
        argument: sub ? parseOmmlChildren(sub) : [],
      };
    }
    case "sSubSup": {
      const e = child(node, "e");
      const sub = child(node, "sub");
      const sup = child(node, "sup");
      return {
        type: "subSup",
        base: e ? parseOmmlChildren(e) : [],
        subscript: sub ? parseOmmlChildren(sub) : [],
        superscript: sup ? parseOmmlChildren(sup) : [],
      };
    }
    case "nary": {
      const naryPr = child(node, "naryPr");
      const chrNode = naryPr ? child(naryPr, "chr") : void 0;
      const operator = chrNode?.attrs["m:val"] ?? chrNode?.attrs["val"] ?? "∑";
      const sub = child(node, "sub");
      const sup = child(node, "sup");
      const e = child(node, "e");
      return {
        type: "nary",
        operator,
        lower: sub ? parseOmmlChildren(sub) : [],
        upper: sup ? parseOmmlChildren(sup) : [],
        body: e ? parseOmmlChildren(e) : [],
      };
    }
    case "rad": {
      const radPr = child(node, "radPr");
      const degHide = radPr ? child(radPr, "degHide") : void 0;
      const hideDegree = degHide?.attrs["m:val"] === "1" || degHide?.attrs["val"] === "1";
      const deg = child(node, "deg");
      const e = child(node, "e");
      return {
        type: "rad",
        hideDegree,
        degree: deg ? parseOmmlChildren(deg) : [],
        body: e ? parseOmmlChildren(e) : [],
      };
    }
    case "d": {
      const dPr = child(node, "dPr");
      const begChr = dPr ? child(dPr, "begChr") : void 0;
      const endChr = dPr ? child(dPr, "endChr") : void 0;
      const e = child(node, "e");
      return {
        type: "delim",
        open: begChr?.attrs["m:val"] ?? begChr?.attrs["val"] ?? "(",
        close: endChr?.attrs["m:val"] ?? endChr?.attrs["val"] ?? ")",
        body: e ? parseOmmlChildren(e) : [],
      };
    }
    case "func": {
      const fName = child(node, "fName");
      const e = child(node, "e");
      return {
        type: "func",
        base: fName ? parseOmmlChildren(fName) : [],
        argument: e ? parseOmmlChildren(e) : [],
      };
    }
    case "groupChr": {
      const groupChrPr = child(node, "groupChrPr");
      const chrNode = groupChrPr ? child(groupChrPr, "chr") : void 0;
      const e = child(node, "e");
      return {
        type: "groupChar",
        operator: chrNode?.attrs["m:val"] ?? chrNode?.attrs["val"] ?? "⏟",
        base: e ? parseOmmlChildren(e) : [],
      };
    }
    case "acc": {
      const accPr = child(node, "accPr");
      const chrNode = accPr ? child(accPr, "chr") : void 0;
      const e = child(node, "e");
      return {
        type: "accent",
        operator: chrNode?.attrs["m:val"] ?? chrNode?.attrs["val"] ?? "̂",
        base: e ? parseOmmlChildren(e) : [],
      };
    }
    case "bar": {
      const e = child(node, "e");
      return {
        type: "bar",
        base: e ? parseOmmlChildren(e) : [],
      };
    }
    case "limLow": {
      const e = child(node, "e");
      const lim = child(node, "lim");
      return {
        type: "limLow",
        base: e ? parseOmmlChildren(e) : [],
        argument: lim ? parseOmmlChildren(lim) : [],
      };
    }
    case "limUpp": {
      const e = child(node, "e");
      const lim = child(node, "lim");
      return {
        type: "limUpp",
        base: e ? parseOmmlChildren(e) : [],
        argument: lim ? parseOmmlChildren(lim) : [],
      };
    }
    case "m": {
      const rows = [];
      for (const mr of children(node, "mr")) {
        const row = [];
        for (const e of children(mr, "e")) {
          const cellElements = parseOmmlChildren(e);
          if (cellElements.length === 1) row.push(cellElements[0]);
          else
            row.push({
              type: "text",
              text: cellElements.map((el) => el.text ?? "").join(""),
            });
        }
        rows.push(row);
      }
      return {
        type: "matrix",
        rows,
      };
    }
    case "eqArr": {
      const rows = [];
      for (const e of children(node, "e")) rows.push(parseOmmlChildren(e));
      return {
        type: "eqArr",
        rows,
      };
    }
    default:
      return null;
  }
}
function parseCellImagesFromRichData(vmCells, getText, getEntry, ws) {
  const metaXml = getText("xl/metadata.xml");
  if (!metaXml) return;
  const metaRoot = parseXml(metaXml);
  const valueMeta = child(metaRoot, "valueMetadata");
  if (!valueMeta) return;
  const vmBks = children(valueMeta, "bk");
  const futureMeta = metaRoot.children.find(
    (c) => typeof c !== "string" && localName(c.tag) === "futureMetadata",
  );
  const fmBks = futureMeta ? children(futureMeta, "bk") : [];
  const rvXml = getText("xl/richData/rdrichvalue.xml");
  if (!rvXml) return;
  const rvEntries = children(parseXml(rvXml), "rv");
  const rvRelRelsXml = getText("xl/richData/_rels/richValueRel.xml.rels");
  if (!rvRelRelsXml) return;
  const rvRelRels = parseRels(rvRelRelsXml);
  const rvRelXml = getText("xl/richData/richValueRel.xml");
  if (!rvRelXml) return;
  const relEntries = children(parseXml(rvRelXml), "rel");
  for (const [cellRef, vmIdx] of vmCells) {
    const vmBk = vmBks[vmIdx - 1];
    if (!vmBk) continue;
    const rc = vmBk.children.find((c) => typeof c !== "string" && localName(c.tag) === "rc");
    if (!rc) continue;
    const fmIdx = parseInt(rc.attrs["v"] ?? "-1", 10);
    if (fmIdx < 0) continue;
    const fmBk = fmBks[fmIdx];
    if (!fmBk) continue;
    const rvb = findDescendant(fmBk, "rvb");
    const rvIdx = parseInt(rvb?.attrs["i"] ?? "-1", 10);
    if (rvIdx < 0) continue;
    const rv = rvEntries[rvIdx];
    if (!rv) continue;
    const vNodes = children(rv, "v");
    const relIdx = parseInt(vNodes[0]?.text ?? "-1", 10);
    if (relIdx < 0) continue;
    const relEntry = relEntries[relIdx];
    if (!relEntry) continue;
    const rId = relEntry.attrs["r:id"] ?? "";
    const imgRel = rvRelRels.get(rId);
    if (!imgRel) continue;
    const imgPath = imgRel.target.startsWith("/")
      ? imgRel.target.slice(1)
      : `xl/richData/${imgRel.target}`;
    const imgEntry = getEntry(imgPath);
    if (!imgEntry) continue;
    const format = extToFormat(imgPath.split(".").pop() ?? "png");
    ws.addCellImage({
      data: imgEntry.data,
      format,
      cell: cellRef,
    });
  }
}
function findDescendant(node, tagName) {
  for (const c of node.children) {
    if (typeof c === "string") continue;
    if (localName(c.tag) === tagName) return c;
    const found = findDescendant(c, tagName);
    if (found) return found;
  }
}
function resolvePath(base, relative) {
  const parts = base.replace(/\/$/, "").split("/");
  for (const seg of relative.split("/"))
    if (seg === "..") parts.pop();
    else if (seg !== ".") parts.push(seg);
  return parts.join("/");
}
async function readWorkbook(data) {
  const zip = await readZip(data);
  const get = (path) => {
    return zip.get(path) ?? zip.get(path.replace(/^\//, ""));
  };
  const getText = (path) => {
    const e = get(path);
    return e ? entryText(e) : void 0;
  };
  const ctXml = getText("[Content_Types].xml") ?? "<Types/>";
  const contentTypes = parseContentTypes(ctXml);
  const workbookRels = parseRels(getText("xl/_rels/workbook.xml.rels") ?? "<Relationships/>");
  const wbXml = getText("xl/workbook.xml") ?? "<workbook/>";
  const stylesXml = getText("xl/styles.xml") ?? "<styleSheet/>";
  const styles = parseStyles(stylesXml);
  const ssXml = getText("xl/sharedStrings.xml") ?? "<sst/>";
  const sharedStrings = ssXml !== "<sst/>" ? parseSharedStrings(ssXml) : [];
  const coreXml = getText("docProps/core.xml") ?? "";
  const core = coreXml ? parseCoreXml(coreXml) : {};
  const appXml = getText("docProps/app.xml") ?? "";
  let extended = {};
  let extendedUnknownRaw = "";
  if (appXml) {
    const r = parseAppXml(appXml);
    extended = r.props;
    extendedUnknownRaw = r.unknownRaw;
  }
  const customXml = getText("docProps/custom.xml") ?? "";
  const custom = customXml ? parseCustomXml(customXml) : [];
  const wbRoot = parseXml(wbXml);
  const sheetsNode = find(wbRoot, "sheets");
  const sheetNodes = sheetsNode ? children(sheetsNode, "sheet") : [];
  const namedRanges = [];
  const definedNamesNode = find(wbRoot, "definedNames");
  if (definedNamesNode)
    for (const dn of children(definedNamesNode, "definedName")) {
      const name = dn.attrs["name"] ?? "";
      const ref = dn.text ?? "";
      if (!name || !ref) continue;
      const nr = {
        name,
        ref,
      };
      if (dn.attrs["localSheetId"] !== void 0) {
        const scopeSheet = sheetNodes[parseInt(dn.attrs["localSheetId"], 10)];
        if (scopeSheet) nr.scope = scopeSheet.attrs["name"] ?? void 0;
      }
      if (dn.attrs["comment"]) nr.comment = dn.attrs["comment"];
      namedRanges.push(nr);
    }
  const allRels = /* @__PURE__ */ new Map();
  for (const [path, entry] of zip)
    if (path.includes("_rels/")) allRels.set(path, parseRels(entryText(entry)));
  const sheets = [];
  for (const sn of sheetNodes) {
    const rId = sn.attrs["r:id"] ?? Object.values(sn.attrs).find((v) => v.startsWith("rId")) ?? "";
    const sheetId = sn.attrs["sheetId"] ?? "";
    const name = sn.attrs["name"] ?? `Sheet${sheetId}`;
    const rel = workbookRels.get(rId);
    if (!rel) continue;
    const target = rel.target.startsWith("/") ? rel.target.slice(1) : `xl/${rel.target}`;
    const sheetXml = getText(target) ?? "";
    if (!sheetXml) continue;
    const {
      ws,
      originalXml,
      unknownParts: sheetUnknown,
      tableRIds,
      legacyDrawingRId,
      ctrlPropRIds,
      drawingRId,
      vmCells,
    } = parseWorksheet(sheetXml, name, styles, sharedStrings);
    ws.sheetIndex = sheets.length + 1;
    ws.rId = rId;
    const tablePaths = [];
    const tableXmls = [];
    if (tableRIds.length) {
      const sheetFileName = target.split("/").pop() ?? "";
      const sheetDir = target.substring(0, target.lastIndexOf("/") + 1);
      const sheetRelsPath = `${sheetDir}_rels/${sheetFileName}.rels`;
      const sheetRels = allRels.get(sheetRelsPath);
      if (sheetRels) {
        for (const tblRId of tableRIds) {
          const tblRel = sheetRels.get(tblRId);
          if (!tblRel) continue;
          const tblTarget = tblRel.target.startsWith("/")
            ? tblRel.target.slice(1)
            : resolvePath(sheetDir, tblRel.target);
          const tblXml = getText(tblTarget);
          if (tblXml) {
            const table = parseTableXml(tblXml, styles);
            if (table) ws.addTable(table);
            tablePaths.push(tblTarget);
            tableXmls.push(tblXml);
          }
        }
        ws.tableRIds = tableRIds;
      }
    }
    if (legacyDrawingRId && ctrlPropRIds.length) {
      const sheetFileName = target.split("/").pop() ?? "";
      const sheetDir = target.substring(0, target.lastIndexOf("/") + 1);
      const sheetRelsPath = `${sheetDir}_rels/${sheetFileName}.rels`;
      const sheetRels = allRels.get(sheetRelsPath);
      if (sheetRels) {
        const vmlRel = sheetRels.get(legacyDrawingRId);
        const vmlPath = vmlRel
          ? vmlRel.target.startsWith("/")
            ? vmlRel.target.slice(1)
            : resolvePath(sheetDir, vmlRel.target)
          : "";
        const vmlXml = vmlPath ? getText(vmlPath) : "";
        const vmlControls = [];
        if (vmlXml) {
          const vmlRoot = parseXml(vmlXml);
          for (const shape of vmlRoot.children) {
            if (typeof shape === "string") continue;
            if (localName(shape.tag) !== "shape") continue;
            const cd = shape.children.find(
              (c) => typeof c !== "string" && localName(c.tag) === "ClientData",
            );
            if (!cd) continue;
            const objType = cd.attrs["ObjectType"] ?? "";
            if (objType === "Note" || !objType) continue;
            const idStr = (shape.attrs["id"] ?? "").replace(/\D/g, "");
            const shapeId = parseInt(idStr, 10) || 0;
            vmlControls.push({
              objectType: objType,
              shapeXml: nodeToXml(shape),
              clientData: cd,
              shapeId,
            });
          }
        }
        for (let ci = 0; ci < ctrlPropRIds.length; ci++) {
          const cpRel = sheetRels.get(ctrlPropRIds[ci]);
          if (!cpRel) continue;
          const cpXml =
            getText(
              cpRel.target.startsWith("/")
                ? cpRel.target.slice(1)
                : resolvePath(sheetDir, cpRel.target),
            ) ?? "";
          const cpRoot = cpXml ? parseXml(cpXml) : null;
          const typeName = OBJ_TYPE_TO_CTRL[cpRoot?.attrs["objectType"] ?? ""] ?? "button";
          const vml = vmlControls[ci];
          const anchor = parseVmlAnchor(vml?.clientData);
          const ctrl = {
            type: typeName,
            from: anchor.from,
            to: anchor.to,
            _ctrlPropXml: cpXml || void 0,
            _vmlShapeXml: vml?.shapeXml,
            _shapeId: vml?.shapeId,
          };
          if (cpRoot) {
            if (cpRoot.attrs["fmlaLink"]) ctrl.linkedCell = cpRoot.attrs["fmlaLink"];
            if (cpRoot.attrs["fmlaRange"]) ctrl.inputRange = cpRoot.attrs["fmlaRange"];
            if (cpRoot.attrs["checked"])
              ctrl.checked = CHECKED_REV[cpRoot.attrs["checked"]] ?? "unchecked";
            if (cpRoot.attrs["dropLines"]) ctrl.dropLines = parseInt(cpRoot.attrs["dropLines"], 10);
            if (cpRoot.attrs["min"]) ctrl.min = parseInt(cpRoot.attrs["min"], 10);
            if (cpRoot.attrs["max"]) ctrl.max = parseInt(cpRoot.attrs["max"], 10);
            if (cpRoot.attrs["inc"]) ctrl.inc = parseInt(cpRoot.attrs["inc"], 10);
            if (cpRoot.attrs["page"]) ctrl.page = parseInt(cpRoot.attrs["page"], 10);
            if (cpRoot.attrs["val"]) ctrl.val = parseInt(cpRoot.attrs["val"], 10);
            if (cpRoot.attrs["selType"])
              ctrl.selType =
                {
                  Single: "single",
                  Multi: "multi",
                  Extend: "extend",
                }[cpRoot.attrs["selType"]] ?? "single";
            if (cpRoot.attrs["noThreeD"] === "1") ctrl.noThreeD = true;
          }
          if (vml?.clientData) {
            const macroNode = vml.clientData.children.find(
              (c) => typeof c !== "string" && localName(c.tag) === "FmlaMacro",
            );
            if (macroNode) ctrl.macro = macroNode.text ?? "";
          }
          ws.addFormControl(ctrl);
        }
        ws.legacyDrawingRId = legacyDrawingRId;
        ws.ctrlPropRIds = ctrlPropRIds;
      }
    }
    if (drawingRId) {
      const sheetFileName = target.split("/").pop() ?? "";
      const sheetDir = target.substring(0, target.lastIndexOf("/") + 1);
      const sheetRelsPath = `${sheetDir}_rels/${sheetFileName}.rels`;
      const sheetRels = allRels.get(sheetRelsPath);
      if (sheetRels) {
        const drawRel = sheetRels.get(drawingRId);
        if (drawRel) {
          const drawPath = drawRel.target.startsWith("/")
            ? drawRel.target.slice(1)
            : resolvePath(sheetDir, drawRel.target);
          const drawXml = getText(drawPath);
          if (drawXml) {
            const drawDir = drawPath.substring(0, drawPath.lastIndexOf("/") + 1);
            const drawRelsPath = `${drawDir}_rels/${drawPath.split("/").pop() ?? ""}.rels`;
            parseDrawingObjects(drawXml, allRels.get(drawRelsPath), drawDir, get, ws);
          }
        }
      }
    }
    if (vmCells.size > 0) parseCellImagesFromRichData(vmCells, getText, get, ws);
    sheets.push({
      ws,
      sheetId,
      rId,
      originalXml,
      unknownParts: sheetUnknown,
      tablePaths,
      tableXmls,
    });
  }
  const connectionsXml = getText("xl/connections.xml") ?? "";
  const connections = [];
  if (connectionsXml) {
    const connRoot = parseXml(connectionsXml);
    for (const cn of children(connRoot, "connection")) {
      const id = parseInt(cn.attrs["id"] ?? "0", 10);
      const name = cn.attrs["name"] ?? "";
      const type = connTypeFromNum(parseInt(cn.attrs["type"] ?? "0", 10));
      if (!name || !type) continue;
      const conn = {
        id,
        name,
        type,
      };
      if (cn.attrs["description"]) conn.description = cn.attrs["description"];
      if (cn.attrs["refreshOnLoad"] === "1") conn.refreshOnLoad = true;
      if (cn.attrs["background"] === "1") conn.background = true;
      if (cn.attrs["saveData"] === "1") conn.saveData = true;
      if (cn.attrs["keepAlive"] === "1") conn.keepAlive = true;
      if (cn.attrs["interval"]) conn.interval = parseInt(cn.attrs["interval"], 10);
      const dbPr = child(cn, "dbPr");
      if (dbPr) {
        if (dbPr.attrs["connection"]) conn.connectionString = dbPr.attrs["connection"];
        if (dbPr.attrs["command"]) conn.command = dbPr.attrs["command"];
        if (dbPr.attrs["commandType"])
          conn.commandType = cmdTypeFromNum(parseInt(dbPr.attrs["commandType"], 10));
      }
      conn._rawXml = nodeToXml(cn);
      connections.push(conn);
    }
  }
  const powerQueries = [];
  for (const [path, entry] of zip) {
    if (!path.startsWith("customXml/item") || path.includes("Props") || path.includes("_rels"))
      continue;
    try {
      const pqs = await parseDataMashup(entry.data);
      if (pqs.length) {
        powerQueries.push(...pqs);
        break;
      }
    } catch {}
  }
  const handledPrefixes = /* @__PURE__ */ new Set([
    "xl/workbook.xml",
    "xl/styles.xml",
    "xl/sharedStrings.xml",
    "xl/worksheets/",
    "docProps/",
    "[Content_Types].xml",
    "_rels/",
    "xl/_rels/",
    "xl/connections.xml",
  ]);
  const unknownParts = /* @__PURE__ */ new Map();
  for (const [path, entry] of zip) {
    if (path.endsWith("_rels/") || path === "[Content_Types].xml") continue;
    if (![...handledPrefixes].some((p) => path.startsWith(p))) unknownParts.set(path, entry.data);
  }
  const nonPrintAreaRanges = [];
  for (const nr of namedRanges)
    if (nr.name === "_xlnm.Print_Area" && nr.scope) {
      const ws = sheets.find((s) => s.ws.name === nr.scope);
      if (ws) ws.ws.printArea = nr.ref;
    } else nonPrintAreaRanges.push(nr);
  return {
    sheets,
    styles,
    stylesXml,
    sharedStrings,
    sharedXml: ssXml,
    workbookXml: wbXml,
    workbookRels,
    contentTypes,
    contentTypesXml: ctXml,
    core,
    extended,
    extendedUnknownRaw,
    custom,
    hasCustomProps: custom.length > 0,
    namedRanges: nonPrintAreaRanges,
    connections,
    connectionsXml,
    powerQueries,
    unknownParts,
    allRels,
  };
}
function parseVmlAnchor(clientData) {
  const defaultAnchor = {
    from: {
      col: 0,
      row: 0,
    },
    to: {
      col: 2,
      row: 2,
    },
  };
  if (!clientData) return defaultAnchor;
  const anchorNode = clientData.children.find(
    (c) => typeof c !== "string" && localName(c.tag) === "Anchor",
  );
  if (!anchorNode) return defaultAnchor;
  const parts = (anchorNode.text ?? "").split(",").map((s) => parseInt(s.trim(), 10));
  if (parts.length < 8 || parts.some(isNaN)) return defaultAnchor;
  return {
    from: {
      col: parts[0],
      colOff: parts[1],
      row: parts[2],
      rowOff: parts[3],
    },
    to: {
      col: parts[4],
      colOff: parts[5],
      row: parts[6],
      rowOff: parts[7],
    },
  };
}
var CONN_TYPE_MAP = {
  1: "odbc",
  2: "dao",
  3: "file",
  4: "web",
  5: "oledb",
  6: "text",
  7: "dsp",
};
var CONN_TYPE_REV = Object.fromEntries(
  Object.entries(CONN_TYPE_MAP).map(([k, v]) => [v, Number(k)]),
);
function connTypeFromNum(n) {
  return CONN_TYPE_MAP[n];
}
function connTypeToNum(t) {
  return CONN_TYPE_REV[t];
}
var CMD_TYPE_MAP = {
  1: "sql",
  2: "table",
  3: "default",
  4: "web",
  5: "oledb",
};
var CMD_TYPE_REV = Object.fromEntries(Object.entries(CMD_TYPE_MAP).map(([k, v]) => [v, Number(k)]));
function cmdTypeFromNum(n) {
  return CMD_TYPE_MAP[n];
}
function cmdTypeToNum(t) {
  return CMD_TYPE_REV[t];
}
async function parseDataMashup(data) {
  if (data.length < 12) return [];
  const view = new DataView(data.buffer, data.byteOffset, data.byteLength);
  if (view.getUint32(0, true) !== 0) return [];
  const pkgLen = view.getUint32(4, true);
  if (pkgLen < 4 || 8 + pkgLen > data.length) return [];
  const pkgBytes = data.subarray(8, 8 + pkgLen);
  if (pkgBytes[0] !== 80 || pkgBytes[1] !== 75) return [];
  const innerZip = await readZip(pkgBytes);
  const queries = [];
  for (const [path, entry] of innerZip) {
    if (!path.includes("/Formula/") || !path.endsWith(".m")) continue;
    const formula = entryText(entry);
    if (!formula) continue;
    const pathMatch = path.match(/Formulas\/([^/]+)\//);
    const nameFromPath = pathMatch ? pathMatch[1] : void 0;
    const sharedRe = /shared\s+(?:#"([^"]+)"|(\w+))\s*=/g;
    let m;
    const foundNames = /* @__PURE__ */ new Set();
    while ((m = sharedRe.exec(formula)) !== null) {
      const qName = m[1] ?? m[2];
      foundNames.add(qName);
    }
    if (foundNames.size > 0) {
      const sectionRe = /shared\s+(?:#"([^"]+)"|(\w+))\s*=\s*([\s\S]*?)(?=,\s*shared\s|\]\s*$)/g;
      let sm;
      while ((sm = sectionRe.exec(formula)) !== null) {
        const qName = sm[1] ?? sm[2];
        const qFormula = sm[3].replace(/,\s*$/, "").trim();
        queries.push({
          name: qName,
          formula: qFormula,
        });
      }
      if (queries.length === 0)
        queries.push({
          name: nameFromPath ?? "Section1",
          formula,
        });
    } else if (nameFromPath)
      queries.push({
        name: nameFromPath,
        formula,
      });
  }
  return queries;
}
//#endregion
//#region node_modules/@node-projects/excelforge/dist/core/Workbook.js
var buildChartXml$1;
var buildCtrlPropXml$1;
var buildFormControlVmlShape$1;
var buildVmlWithControls$1;
var buildPivotTableFiles$1;
var VbaProject$1;
var Workbook = class Workbook {
  constructor() {
    this.sheets = [];
    this.namedRanges = [];
    this.connections = [];
    this.powerQueries = [];
    this.externalLinks = [];
    this.customPivotStyles = [];
    this.pivotSlicers = [];
    this.properties = {};
    this.compressionLevel = 6;
    this.coreProperties = {};
    this.extendedProperties = {};
    this.customProperties = [];
    this.isTemplate = false;
    this._dirtySheets = /* @__PURE__ */ new Set();
    this._activeTabIndex = 0;
    this._customTableStyles = /* @__PURE__ */ new Map();
  }
  markDirty(sheetIndexOrName) {
    if (typeof sheetIndexOrName === "string") {
      const idx = this.sheets.findIndex((s) => s.name === sheetIndexOrName);
      if (idx >= 0) this._dirtySheets.add(idx);
    } else this._dirtySheets.add(sheetIndexOrName);
  }
  static async fromBytes(data) {
    const wb = new Workbook();
    const result = await readWorkbook(data);
    wb._readResult = result;
    wb.coreProperties = result.core;
    wb.extendedProperties = result.extended;
    wb.customProperties = result.custom;
    wb.properties = {
      title: result.core.title,
      author: result.core.creator,
      subject: result.core.subject,
      description: result.core.description,
      keywords: result.core.keywords,
      company: result.extended.company,
      lastModifiedBy: result.core.lastModifiedBy,
      created: result.core.created,
      modified: result.core.modified,
      category: result.core.category,
      status: result.core.contentStatus,
    };
    wb.sheets = result.sheets.map((s) => s.ws);
    for (let i = 0; i < wb.sheets.length; i++) {
      const ws = wb.sheets[i];
      ws._workbook = wb;
      if (ws.view?.tabSelected) wb._activeTabIndex = i;
    }
    wb.namedRanges = result.namedRanges;
    wb.connections = result.connections;
    wb.powerQueries = result.powerQueries;
    const vbaData = result.unknownParts.get("xl/vbaProject.bin");
    if (vbaData) {
      if (!VbaProject$1)
        VbaProject$1 = (await Promise.resolve().then(() => VbaProject_exports)).VbaProject;
      try {
        wb.vbaProject = VbaProject$1.fromBytes(vbaData);
      } catch {}
    }
    const calcPrMatch = result.workbookXml.match(/<calcPr([^>]*)\/?>/);
    if (calcPrMatch) {
      const a = calcPrMatch[1];
      const cs = {};
      const modeMatch = a.match(/calcMode="([^"]+)"/);
      if (modeMatch) cs.calcMode = modeMatch[1];
      if (/fullCalcOnLoad="1"/.test(a)) cs.fullCalcOnLoad = true;
      else if (/fullCalcOnLoad="0"/.test(a)) cs.fullCalcOnLoad = false;
      if (/iterate="1"/.test(a)) cs.iterate = true;
      const icMatch = a.match(/iterateCount="(\d+)"/);
      if (icMatch) cs.iterateCount = parseInt(icMatch[1], 10);
      const idMatch = a.match(/iterateDelta="([^"]+)"/);
      if (idMatch) cs.iterateDelta = parseFloat(idMatch[1]);
      if (/fullPrecision="1"/.test(a)) cs.fullPrecision = true;
      else if (/fullPrecision="0"/.test(a)) cs.fullPrecision = false;
      if (/calcOnSave="1"/.test(a)) cs.calcOnSave = true;
      else if (/calcOnSave="0"/.test(a)) cs.calcOnSave = false;
      if (/concurrentCalc="1"/.test(a)) cs.concurrentCalc = true;
      else if (/concurrentCalc="0"/.test(a)) cs.concurrentCalc = false;
      if (Object.keys(cs).length > 0) wb.calcSettings = cs;
    }
    return wb;
  }
  static async fromBase64(b64) {
    return Workbook.fromBytes(base64ToBytes(b64));
  }
  static async fromFile(path) {
    const buf = await (await import("fs/promises")).readFile(path);
    return Workbook.fromBytes(new Uint8Array(buf));
  }
  static async fromBlob(blob) {
    return Workbook.fromBytes(new Uint8Array(await blob.arrayBuffer()));
  }
  addSheet(name, options = {}) {
    const ws = new Worksheet(name, options);
    ws.sheetIndex = this.sheets.length + 1;
    ws._workbook = this;
    this.sheets.push(ws);
    this._dirtySheets.add(ws.sheetIndex - 1);
    return ws;
  }
  getSheet(name) {
    return this.sheets.find((s) => s.name === name);
  }
  getSheetByIndex(idx) {
    return this.sheets[idx];
  }
  getSheetNames() {
    return this.sheets.map((s) => s.name);
  }
  getSheets() {
    return this.sheets;
  }
  removeSheet(name) {
    this.sheets = this.sheets.filter((s) => s.name !== name);
    return this;
  }
  setActiveSheet(sheetIndexOrName) {
    const idx =
      typeof sheetIndexOrName === "string"
        ? this.sheets.findIndex((s) => s.name === sheetIndexOrName)
        : sheetIndexOrName;
    if (idx < 0 || idx >= this.sheets.length) return;
    for (let i = 0; i < this.sheets.length; i++) {
      const ws = this.sheets[i];
      if (i === idx)
        ws.view = {
          ...(ws.view ?? {}),
          tabSelected: true,
        };
      else if (ws.view?.tabSelected)
        ws.view = {
          ...ws.view,
          tabSelected: false,
        };
    }
    this._activeTabIndex = idx;
  }
  addChartSheet(name, chart) {
    const ws = this.addSheet(name);
    ws._isChartSheet = true;
    ws.addChart(chart);
    return ws;
  }
  addDialogSheet(name) {
    const ws = this.addSheet(name);
    ws._isDialogSheet = true;
    return ws;
  }
  copySheet(sourceName, newName) {
    const src = this.getSheet(sourceName);
    if (!src) throw new Error(`Sheet "${sourceName}" not found`);
    const ws = this.addSheet(newName);
    const cells = src.readAllCells();
    for (const { row, col, cell } of cells) {
      const target = ws.getCell(row, col);
      if (cell.value != null) target.value = cell.value;
      if (cell.formula) target.formula = cell.formula;
      if (cell.arrayFormula) target.arrayFormula = cell.arrayFormula;
      if (cell.richText)
        target.richText = cell.richText.map((r) => ({
          ...r,
          font: r.font ? { ...r.font } : void 0,
        }));
      if (cell.style) target.style = { ...cell.style };
      if (cell.comment) target.comment = { ...cell.comment };
      if (cell.hyperlink) target.hyperlink = { ...cell.hyperlink };
    }
    for (const m of src.getMerges()) ws.merge(m.startRow, m.startCol, m.endRow, m.endCol);
    const range = src.getUsedRange();
    if (range)
      for (let c = range.startCol; c <= range.endCol; c++) {
        const cd = src.getColumn(c);
        if (cd) ws.setColumn(c, { ...cd });
      }
    if (src.pageSetup) ws.pageSetup = { ...src.pageSetup };
    if (src.printArea) ws.printArea = src.printArea;
    return ws;
  }
  registerTableStyle(name, def) {
    this._customTableStyles.set(name, def);
    return this;
  }
  addNamedRange(nr) {
    this.namedRanges.push(nr);
    return this;
  }
  getNamedRanges() {
    return this.namedRanges;
  }
  getNamedRange(name) {
    return this.namedRanges.find((nr) => nr.name === name);
  }
  removeNamedRange(name) {
    this.namedRanges = this.namedRanges.filter((nr) => nr.name !== name);
    return this;
  }
  addConnection(conn) {
    this.connections.push(conn);
    return this;
  }
  getConnections() {
    return this.connections;
  }
  getConnection(name) {
    return this.connections.find((c) => c.name === name);
  }
  removeConnection(name) {
    this.connections = this.connections.filter((c) => c.name !== name);
    return this;
  }
  getPowerQueries() {
    return this.powerQueries;
  }
  getPowerQuery(name) {
    return this.powerQueries.find((q) => q.name === name);
  }
  getCustomProperty(name) {
    return this.customProperties.find((p) => p.name === name);
  }
  setCustomProperty(name, value) {
    const idx = this.customProperties.findIndex((p) => p.name === name);
    if (idx >= 0)
      this.customProperties[idx] = {
        name,
        value,
      };
    else
      this.customProperties.push({
        name,
        value,
      });
    return this;
  }
  removeCustomProperty(name) {
    this.customProperties = this.customProperties.filter((p) => p.name !== name);
    return this;
  }
  addExternalLink(link) {
    this.externalLinks.push(link);
    return this;
  }
  getExternalLinks() {
    return this.externalLinks;
  }
  registerPivotStyle(style) {
    this.customPivotStyles.push(style);
    return this;
  }
  addPivotSlicer(slicer) {
    this.pivotSlicers.push(slicer);
    return this;
  }
  getPivotSlicers() {
    return this.pivotSlicers;
  }
  async build() {
    this._syncLegacyProperties();
    return this._readResult ? this._buildPatched() : this._buildFresh();
  }
  async _buildPatched() {
    const rr = this._readResult;
    const entries = [];
    const hasDirty = this._dirtySheets.size > 0;
    const styles = new StyleRegistry();
    const shared = new SharedStrings();
    const sheetXmls = /* @__PURE__ */ new Map();
    const origTableColNames = /* @__PURE__ */ new Map();
    if (hasDirty) {
      for (let i = 0; i < this.sheets.length; i++) {
        const ws = this.sheets[i];
        const tables = ws.getTables();
        for (let ti = 0; ti < tables.length; ti++) {
          const tbl = tables[ti];
          origTableColNames.set(
            `${i}:${ti}`,
            tbl.columns.map((c) => c.name),
          );
          const { startRow, startCol } = parseRange(tbl.ref);
          const usedNames = /* @__PURE__ */ new Set();
          for (let ci = 0; ci < tbl.columns.length; ci++) {
            const cell = ws.getCell(startRow, startCol + ci);
            const cellVal = cell.value;
            let strVal = null;
            if (cellVal == null) strVal = null;
            else if (cellVal instanceof Date) strVal = cellVal.toISOString();
            else if (typeof cellVal === "object" && "error" in cellVal)
              strVal = String(cellVal.error);
            else strVal = String(cellVal);
            const baseName = strVal ?? tbl.columns[ci].name;
            let finalName = baseName;
            let dup = 2;
            while (usedNames.has(finalName)) finalName = `${baseName}${dup++}`;
            usedNames.add(finalName);
            tbl.columns[ci].name = finalName;
            if (cellVal == null) cell.value = finalName;
            else if (typeof cellVal !== "string") cell.value = finalName;
            else if (cellVal !== finalName) cell.value = finalName;
          }
        }
      }
      const dxfRe = /<dxf>([\s\S]*?)<\/dxf>|<dxf\/>/g;
      const rawDxfs = [];
      let m;
      while ((m = dxfRe.exec(rr.stylesXml)) !== null) rawDxfs.push(m[1] ?? "");
      if (rawDxfs.length) styles.prependRawDxfs(rawDxfs);
      for (let i = 0; i < this.sheets.length; i++)
        sheetXmls.set(i, this.sheets[i].toXml(styles, shared));
    }
    entries.push({
      name: "docProps/core.xml",
      data: strToBytes(
        buildCoreXml({
          ...rr.core,
          ...this.coreProperties,
          modified: /* @__PURE__ */ new Date(),
        }),
      ),
    });
    entries.push({
      name: "docProps/app.xml",
      data: strToBytes(
        buildAppXml(
          {
            ...rr.extended,
            ...this.extendedProperties,
            titlesOfParts: this.sheets.map((s) => s.name),
            headingPairs: this._headingPairs(),
          },
          rr.extendedUnknownRaw,
        ),
      ),
    });
    const customProps =
      this.customProperties.length > 0
        ? this.customProperties
        : rr.hasCustomProps
          ? rr.custom
          : null;
    if (customProps && customProps.length > 0)
      entries.push({
        name: "docProps/custom.xml",
        data: strToBytes(buildCustomXml(customProps)),
      });
    entries.push({
      name: "xl/workbook.xml",
      data: strToBytes(this._patchWorkbookXml(rr.workbookXml, rr)),
    });
    const connectionsXml = this._connectionsXml(rr.connectionsXml);
    if (connectionsXml)
      entries.push({
        name: "xl/connections.xml",
        data: strToBytes(connectionsXml),
      });
    if (hasDirty) {
      entries.push({
        name: "xl/styles.xml",
        data: strToBytes(styles.toXml()),
      });
      entries.push({
        name: "xl/sharedStrings.xml",
        data: strToBytes(shared.toXml()),
      });
    } else {
      entries.push({
        name: "xl/styles.xml",
        data: strToBytes(rr.stylesXml),
      });
      entries.push({
        name: "xl/sharedStrings.xml",
        data: strToBytes(rr.sharedXml),
      });
    }
    for (let i = 0; i < this.sheets.length; i++) {
      const ws = this.sheets[i];
      const path = `xl/${ws._isChartSheet ? "chartsheets" : ws._isDialogSheet ? "dialogsheets" : "worksheets"}/sheet${i + 1}.xml`;
      if (hasDirty)
        if (ws._isChartSheet)
          entries.push({
            name: path,
            data: strToBytes(ws.toChartSheetXml()),
          });
        else if (ws._isDialogSheet)
          entries.push({
            name: path,
            data: strToBytes(ws.toDialogSheetXml(styles, shared)),
          });
        else
          entries.push({
            name: path,
            data: strToBytes(sheetXmls.get(i) ?? ""),
          });
      else
        entries.push({
          name: path,
          data: strToBytes(rr.sheets[i]?.originalXml ?? ""),
        });
    }
    const allTablePaths = /* @__PURE__ */ new Set();
    for (let i = 0; i < this.sheets.length; i++) {
      const tables = this.sheets[i].getTables();
      const paths = rr.sheets[i]?.tablePaths ?? [];
      const xmls = rr.sheets[i]?.tableXmls ?? [];
      for (let j = 0; j < tables.length; j++) {
        const tblPath = paths[j];
        if (tblPath) {
          allTablePaths.add(tblPath);
          if (j < xmls.length) {
            let xml = xmls[j];
            const origRefMatch = xml.match(/\bref="([^"]+)"/);
            if (origRefMatch && origRefMatch[1] !== tables[j].ref)
              xml = xml.replace(`ref="${origRefMatch[1]}"`, `ref="${tables[j].ref}"`);
            if (hasDirty) {
              const origNames = origTableColNames.get(`${i}:${j}`);
              if (origNames)
                for (let ci = 0; ci < tables[j].columns.length; ci++) {
                  const origName = origNames[ci];
                  const newName = tables[j].columns[ci].name;
                  if (origName && origName !== newName)
                    xml = xml.replace(
                      `name="${escapeXml(origName)}"`,
                      `name="${escapeXml(newName)}"`,
                    );
                }
            }
            entries.push({
              name: tblPath,
              data: strToBytes(xml),
            });
          } else {
            const idMatch = tblPath.match(/table(\d+)\.xml$/);
            const tableId = idMatch ? parseInt(idMatch[1], 10) : j + 1;
            entries.push({
              name: tblPath,
              data: strToBytes(buildTableXml(tables[j], tableId)),
            });
          }
        }
      }
    }
    for (const [path, data] of rr.unknownParts) {
      if (allTablePaths.has(path)) continue;
      if (path === "xl/vbaProject.bin" && this.vbaProject) continue;
      if (rr.allRels.has(path)) continue;
      if (hasDirty && path === "xl/calcChain.xml") continue;
      entries.push({
        name: path,
        data,
      });
    }
    if (this.vbaProject) {
      this._ensureVbaSheetModules();
      entries.push({
        name: "xl/vbaProject.bin",
        data: this.vbaProject.build(),
      });
    }
    entries.push({
      name: "_rels/.rels",
      data: strToBytes(this._buildRootRels(customProps != null && customProps.length > 0)),
    });
    entries.push({
      name: "xl/_rels/workbook.xml.rels",
      data: strToBytes(this._buildWorkbookRels(rr, hasDirty)),
    });
    for (const [relPath, relMap] of rr.allRels) {
      if (relPath === "xl/_rels/workbook.xml.rels" || relPath === "_rels/.rels") continue;
      entries.push({
        name: relPath,
        data: strToBytes(this._relsToXml(relMap)),
      });
    }
    entries.push({
      name: "[Content_Types].xml",
      data: strToBytes(
        this._patchContentTypes(
          rr.contentTypesXml,
          customProps != null && customProps.length > 0,
          hasDirty,
        ),
      ),
    });
    return buildZip(entries, { level: this.compressionLevel });
  }
  async _buildFresh() {
    const styles = new StyleRegistry();
    const shared = new SharedStrings();
    const entries = [];
    for (const [name, def] of this._customTableStyles) styles.registerTableStyle(name, def);
    let globalRId = 1;
    for (const ws of this.sheets) ws.rId = `rId${globalRId++}`;
    const allImages = [];
    const allCharts = [];
    const allTables = [];
    const allPivotTables = [];
    const sheetImageRIds = /* @__PURE__ */ new Map();
    const sheetChartRIds = /* @__PURE__ */ new Map();
    const sheetTableRIds = /* @__PURE__ */ new Map();
    const sheetPivotRIds = /* @__PURE__ */ new Map();
    let imgCtr = 1,
      chartCtr = 1,
      tableCtr = 1,
      vmlCtr = 1,
      pivotCtr = 1,
      pivotCacheIdCtr = 0,
      ctrlPropGlobal = 0,
      oleObjGlobal = 0;
    for (const ws of this.sheets) {
      const imgs = ws.getImages();
      const charts = ws.getCharts();
      const tables = ws.getTables();
      const imgRIds = [],
        chartRIds = [],
        tblRIds = [];
      if (
        imgs.length ||
        charts.length ||
        ws.getShapes().length ||
        ws.getWordArt().length ||
        ws.getMathEquations().length ||
        ws.getTableSlicers().length ||
        ws.getOleObjects().length
      )
        ws.drawingRId = `rId${globalRId++}`;
      const controls = ws.getFormControls();
      if (ws.getComments().length || controls.length) ws.legacyDrawingRId = `rId${globalRId++}`;
      for (const img of imgs) {
        const r = `rId${globalRId++}`;
        imgRIds.push(r);
        allImages.push({
          ws,
          img,
          ext: imageExt(img.format),
          idx: imgCtr++,
        });
      }
      for (let i = 0; i < charts.length; i++) {
        const r = `rId${globalRId++}`;
        chartRIds.push(r);
        allCharts.push({
          ws,
          chartIdx: i,
          globalIdx: chartCtr++,
        });
      }
      for (let i = 0; i < tables.length; i++) {
        const r = `rId${globalRId++}`;
        tblRIds.push(r);
        allTables.push({
          ws,
          tableIdx: i,
          globalTableId: tableCtr++,
        });
      }
      const oleRIds = [];
      const oleIconRIds = [];
      for (const ole of ws.getOleObjects()) {
        oleRIds.push(`rId${globalRId++}`);
        if (ole.iconData) oleIconRIds.push(`rId${globalRId++}`);
        else oleIconRIds.push("");
      }
      ws.oleRIds = oleRIds;
      ws.oleIconRIds = oleIconRIds;
      const ctrlPropRIds = [];
      for (let ci = 0; ci < controls.length; ci++) ctrlPropRIds.push(`rId${globalRId++}`);
      ws.ctrlPropRIds = ctrlPropRIds;
      sheetImageRIds.set(ws, imgRIds);
      sheetChartRIds.set(ws, chartRIds);
      sheetTableRIds.set(ws, tblRIds);
      ws.tableRIds = tblRIds;
      const ptRIds = [];
      for (const pt of ws.getPivotTables()) {
        const pivotRId = `rId${globalRId++}`;
        const cacheRId = `rId${globalRId++}`;
        ptRIds.push(pivotRId);
        allPivotTables.push({
          ws,
          pt,
          pivotIdx: pivotCtr++,
          cacheId: pivotCacheIdCtr++,
          pivotRId,
          cacheRId,
        });
      }
      sheetPivotRIds.set(ws, ptRIds);
    }
    const allCellImages = [];
    let cellImgCtr = imgCtr;
    let vmCounter = 1;
    for (const ws of this.sheets) {
      ws._cellImageVm = /* @__PURE__ */ new Map();
      for (const ci of ws.getCellImages()) {
        const ext = imageExt(ci.format);
        allCellImages.push({
          img: ci,
          ext,
          idx: cellImgCtr++,
        });
        ws._cellImageVm.set(ci.cell, vmCounter++);
        ws.getCellByRef(ci.cell);
      }
    }
    const hasCellImages = allCellImages.length > 0;
    const sheetSlicerMap = /* @__PURE__ */ new Map();
    const allSlicerCaches = [];
    let slicerDefCtr = 0,
      slicerCacheCtr = 0;
    for (const ws of this.sheets)
      if (ws.getTableSlicers().length)
        sheetSlicerMap.set(ws, {
          tableSlicers: [],
          pivotSlicers: [],
          slicerDefRId: "",
          slicerDefIdx: 0,
        });
    for (const ps of this.pivotSlicers)
      for (const ws of this.sheets)
        if (ws.getPivotTables().some((p) => p.name === ps.pivotTableName)) {
          if (!sheetSlicerMap.has(ws))
            sheetSlicerMap.set(ws, {
              tableSlicers: [],
              pivotSlicers: [],
              slicerDefRId: "",
              slicerDefIdx: 0,
            });
          break;
        }
    for (const [ws, info] of sheetSlicerMap) {
      if (!ws.drawingRId) ws.drawingRId = `rId${globalRId++}`;
      info.slicerDefRId = `rId${globalRId++}`;
      info.slicerDefIdx = ++slicerDefCtr;
      ws.slicerRId = info.slicerDefRId;
      const drawingInfo = [];
      for (const s of ws.getTableSlicers()) {
        const table = ws.getTables().find((t) => t.name === s.tableName);
        const tableId =
          allTables.find((t) => t.ws === ws && ws.getTables()[t.tableIdx] === table)
            ?.globalTableId ?? 1;
        const columnIndex = table
          ? (table.columns?.findIndex((c) => c.name === s.columnName) ?? 0) + 1
          : 1;
        info.tableSlicers.push({
          slicer: s,
          tableId,
          columnIndex,
        });
        drawingInfo.push({
          name: s.name,
          cell: s.cell,
        });
        allSlicerCaches.push({
          name: s.name + "_cache",
          sourceName: s.columnName,
          type: "table",
          rId: `rId${globalRId++}`,
          idx: ++slicerCacheCtr,
          tableId,
          columnIndex,
          sortOrder: s.sortOrder ?? "ascending",
        });
      }
      for (const ps of this.pivotSlicers) {
        const pt = ws.getPivotTables().find((p) => p.name === ps.pivotTableName);
        if (!pt) continue;
        const ptEntry = allPivotTables.find((p) => p.ws === ws && p.pt === pt);
        const sheetIdx = this.sheets.indexOf(ws) + 1;
        let items = [];
        const sourceWs = this.sheets.find((s) => s.name === pt.sourceSheet);
        if (sourceWs) {
          const sourceData = sourceWs.readRange(pt.sourceRef);
          const fieldIdx = (sourceData[0] ?? []).map((v) => String(v ?? "")).indexOf(ps.fieldName);
          if (fieldIdx >= 0) {
            const uniqueSet = /* @__PURE__ */ new Set();
            for (let r = 1; r < sourceData.length; r++)
              uniqueSet.add(String(sourceData[r][fieldIdx] ?? ""));
            items = [...uniqueSet];
          }
        }
        info.pivotSlicers.push({
          slicer: ps,
          pivotCacheId: ptEntry?.cacheId ?? 0,
        });
        drawingInfo.push({
          name: ps.name,
          cell: ps.cell,
        });
        allSlicerCaches.push({
          name: ps.name + "_cache",
          sourceName: ps.fieldName,
          type: "pivot",
          rId: `rId${globalRId++}`,
          idx: ++slicerCacheCtr,
          pivotTableName: ps.pivotTableName,
          pivotCacheId: ptEntry?.cacheId ?? 0,
          tabId: sheetIdx,
          items,
        });
      }
      ws._slicerDrawingInfo = drawingInfo;
    }
    const hasSlicers = sheetSlicerMap.size > 0;
    const hasCustom = this.customProperties.length > 0;
    const hasVba = !!this.vbaProject;
    const imgCTs = /* @__PURE__ */ new Set();
    for (const { ext } of [...allImages, ...allCellImages]) {
      const ct = imageContentType(ext);
      imgCTs.add(`<Default Extension="${ext}" ContentType="${ct}"/>`);
    }
    const sheetsWithComments = this.sheets.filter((ws) => ws.getComments().length);
    const vmlCT = this.sheets.filter((ws) => ws.getComments().length || ws.getFormControls().length)
      .length
      ? '<Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/>'
      : "";
    let vmlIdx = 0;
    const commentsCTs = sheetsWithComments
      .map(
        () =>
          `<Override PartName="/xl/comments${++vmlIdx}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml"/>`,
      )
      .join("");
    let ctrlPropCtr = 0;
    const ctrlPropCTs = [];
    for (const ws of this.sheets) {
      if (ws._isDialogSheet) continue;
      for (let ci = 0; ci < ws.getFormControls().length; ci++)
        ctrlPropCTs.push(
          `<Override PartName="/xl/ctrlProps/ctrlProp${++ctrlPropCtr}.xml" ContentType="application/vnd.ms-excel.controlproperties+xml"/>`,
        );
    }
    const oleCT = this.sheets.some((ws) => ws.getOleObjects().length > 0)
      ? '<Default Extension="bin" ContentType="application/vnd.openxmlformats-officedocument.oleObject"/>'
      : "";
    entries.push({
      name: "[Content_Types].xml",
      data: strToBytes(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
${vmlCT}
${oleCT}
${[...imgCTs].join("")}
<Override PartName="/xl/workbook.xml" ContentType="${hasVba ? "application/vnd.ms-excel.sheet.macroEnabled.main+xml" : this.isTemplate ? "application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml" : "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"}"/>
${hasVba ? '<Override PartName="/xl/vbaProject.bin" ContentType="application/vnd.ms-office.vbaProject"/>' : ""}
<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
<Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
${this.sheets
  .filter((ws) => !ws._isChartSheet && !ws._isDialogSheet)
  .map((ws) => {
    return `<Override PartName="/xl/worksheets/sheet${this.sheets.indexOf(ws) + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`;
  })
  .join("")}
${this.sheets
  .filter((ws) => ws._isChartSheet)
  .map((ws) => {
    return `<Override PartName="/xl/chartsheets/sheet${this.sheets.indexOf(ws) + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.chartsheet+xml"/>`;
  })
  .join("")}
${this.sheets
  .filter((ws) => ws._isDialogSheet)
  .map((ws) => {
    return `<Override PartName="/xl/dialogsheets/sheet${this.sheets.indexOf(ws) + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.dialogsheet+xml"/>`;
  })
  .join("")}
${this.sheets
  .filter((ws) => ws.drawingRId)
  .map(
    (_, i) =>
      `<Override PartName="/xl/drawings/drawing${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>`,
  )
  .join("")}
${allCharts.map(({ globalIdx }) => `<Override PartName="/xl/charts/chart${globalIdx}.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>`).join("")}
${allTables.map(({ globalTableId }) => `<Override PartName="/xl/tables/table${globalTableId}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/>`).join("")}
${allPivotTables.map((p) => `<Override PartName="/xl/pivotTables/pivotTable${p.pivotIdx}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml"/>`).join("\n")}
${allPivotTables.map((p) => `<Override PartName="/xl/pivotCache/pivotCacheDefinition${p.pivotIdx}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml"/>`).join("\n")}
${allPivotTables.map((p) => `<Override PartName="/xl/pivotCache/pivotCacheRecords${p.pivotIdx}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml"/>`).join("\n")}
${commentsCTs}
${ctrlPropCTs.join("")}
<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
${hasCustom ? '<Override PartName="/docProps/custom.xml" ContentType="application/vnd.openxmlformats-officedocument.custom-properties+xml"/>' : ""}
${this.connections.length ? '<Override PartName="/xl/connections.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.connections+xml"/>' : ""}
${this.externalLinks.map((_, i) => `<Override PartName="/xl/externalLinks/externalLink${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml"/>`).join("\n")}
${[...sheetSlicerMap.values()].map((info) => `<Override PartName="/xl/slicers/slicer${info.slicerDefIdx}.xml" ContentType="application/vnd.ms-excel.slicer+xml"/>`).join("\n")}
${allSlicerCaches.map((sc) => `<Override PartName="/xl/slicerCaches/slicerCache${sc.idx}.xml" ContentType="application/vnd.ms-excel.slicerCache+xml"/>`).join("\n")}
${this.sheets
  .flatMap((ws) => ws.getQueryTables())
  .map(
    (_, i) =>
      `<Override PartName="/xl/queryTables/queryTable${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.queryTable+xml"/>`,
  )
  .join("\n")}
${
  hasCellImages
    ? `<Override PartName="/xl/metadata.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml"/>
<Override PartName="/xl/richData/rdrichvalue.xml" ContentType="application/vnd.ms-excel.rdrichvalue+xml"/>
<Override PartName="/xl/richData/rdRichValueStructure.xml" ContentType="application/vnd.ms-excel.rdrichvaluestructure+xml"/>
<Override PartName="/xl/richData/richValueRel.xml" ContentType="application/vnd.ms-excel.richvaluerel+xml"/>
<Override PartName="/xl/richData/rdRichValueTypes.xml" ContentType="application/vnd.ms-excel.rdrichvaluetypes+xml"/>
<Override PartName="/xl/richData/rdarray.xml" ContentType="application/vnd.ms-excel.rdarray+xml"/>`
    : ""
}
</Types>`),
    });
    entries.push({
      name: "_rels/.rels",
      data: strToBytes(this._buildRootRels(hasCustom)),
    });
    entries.push({
      name: "xl/_rels/workbook.xml.rels",
      data: strToBytes(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
${this.sheets
  .map((ws, i) => {
    const type = ws._isChartSheet ? "chartsheet" : ws._isDialogSheet ? "dialogsheet" : "worksheet";
    const folder = ws._isChartSheet
      ? "chartsheets"
      : ws._isDialogSheet
        ? "dialogsheets"
        : "worksheets";
    return `<Relationship Id="${ws.rId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/${type}" Target="${folder}/sheet${i + 1}.xml"/>`;
  })
  .join("")}
<Relationship Id="rIdStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
<Relationship Id="rIdShared" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
<Relationship Id="rIdTheme" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>
${allPivotTables.map((p) => `<Relationship Id="${p.cacheRId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition" Target="pivotCache/pivotCacheDefinition${p.pivotIdx}.xml"/>`).join("\n")}
${allSlicerCaches.map((sc) => `<Relationship Id="${sc.rId}" Type="http://schemas.microsoft.com/office/2007/relationships/slicerCache" Target="slicerCaches/slicerCache${sc.idx}.xml"/>`).join("\n")}
${hasVba ? '<Relationship Id="rIdVBA" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" Target="vbaProject.bin"/>' : ""}
${this.connections.length ? '<Relationship Id="rIdConns" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/connections" Target="connections.xml"/>' : ""}
${this.externalLinks.map((_, i) => `<Relationship Id="rIdExtLink${i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink" Target="externalLinks/externalLink${i + 1}.xml"/>`).join("\n")}
${hasCellImages ? '<Relationship Id="rIdMeta" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sheetMetadata" Target="metadata.xml"/>' : ""}
${hasCellImages ? '<Relationship Id="rIdRichValueRel" Type="http://schemas.microsoft.com/office/2022/10/relationships/richValueRel" Target="richData/richValueRel.xml"/>' : ""}
${hasCellImages ? '<Relationship Id="rIdRichValue" Type="http://schemas.microsoft.com/office/2017/06/relationships/rdRichValue" Target="richData/rdrichvalue.xml"/>' : ""}
${hasCellImages ? '<Relationship Id="rIdRichValueStruct" Type="http://schemas.microsoft.com/office/2017/06/relationships/rdRichValueStructure" Target="richData/rdRichValueStructure.xml"/>' : ""}
${hasCellImages ? '<Relationship Id="rIdRichValueTypes" Type="http://schemas.microsoft.com/office/2017/06/relationships/rdRichValueTypes" Target="richData/rdRichValueTypes.xml"/>' : ""}
${hasCellImages ? '<Relationship Id="rIdRdArray" Type="http://schemas.microsoft.com/office/2017/06/relationships/rdArray" Target="richData/rdarray.xml"/>' : ""}
</Relationships>`),
    });
    if (hasVba) {
      this._ensureVbaSheetModules();
      entries.push({
        name: "xl/vbaProject.bin",
        data: this.vbaProject.build(),
      });
    }
    const wbPrAttrs = [
      this.properties.date1904 ? 'date1904="1"' : "",
      hasVba ? 'codeName="ThisWorkbook"' : "",
    ]
      .filter(Boolean)
      .join(" ");
    const date1904 = `<workbookPr${wbPrAttrs ? " " + wbPrAttrs : ""}/>`;
    const namedRangesXml = this._definedNamesXml(allSlicerCaches);
    const pivotCachesXml = allPivotTables.length
      ? `<pivotCaches>${allPivotTables.map((p) => `<pivotCache cacheId="${p.cacheId}" r:id="${p.cacheRId}"/>`).join("")}</pivotCaches>`
      : "";
    entries.push({
      name: "xl/workbook.xml",
      data: strToBytes(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
${date1904}
<bookViews><workbookView xWindow="0" yWindow="0" windowWidth="14400" windowHeight="8260"${this._activeTabIndex ? ` activeTab="${this._activeTabIndex}"` : ""}/></bookViews>
<sheets>${this.sheets.map((ws, i) => `<sheet name="${escapeXml(ws.name)}" sheetId="${i + 1}" r:id="${ws.rId}"${ws.options?.state && ws.options.state !== "visible" ? ` state="${ws.options.state}"` : ""}/>`).join("")}</sheets>
${namedRangesXml}
${this._calcPrXml()}
${pivotCachesXml}
${
  hasSlicers
    ? (() => {
        const tableSlicers = allSlicerCaches.filter((sc) => sc.type === "table");
        const pivotSlicers = allSlicerCaches.filter((sc) => sc.type === "pivot");
        let xml = "<extLst>";
        if (pivotSlicers.length)
          xml += `<ext uri="{BBE1A952-AA13-448e-AADC-164F8A28A991}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><x14:slicerCaches>${pivotSlicers.map((sc) => `<x14:slicerCache r:id="${sc.rId}"/>`).join("")}</x14:slicerCaches></ext>`;
        if (tableSlicers.length)
          xml += `<ext uri="{46BE6895-7355-4a93-B00E-2C351335B9C9}" xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><x15:slicerCaches>${tableSlicers.map((sc) => `<x14:slicerCache r:id="${sc.rId}"/>`).join("")}</x15:slicerCaches></ext>`;
        xml += "</extLst>";
        return xml;
      })()
    : ""
}
</workbook>`),
    });
    if (this.connections.length)
      entries.push({
        name: "xl/connections.xml",
        data: strToBytes(this._connectionsXml()),
      });
    for (let i = 0; i < this.sheets.length; i++) {
      const ws = this.sheets[i];
      const imgRIds = sheetImageRIds.get(ws) ?? [];
      const cRIds = sheetChartRIds.get(ws) ?? [];
      const tblEntries = allTables.filter((t) => t.ws === ws);
      const tblRIds_ = sheetTableRIds.get(ws) ?? [];
      const sheetFolder = ws._isChartSheet
        ? "chartsheets"
        : ws._isDialogSheet
          ? "dialogsheets"
          : "worksheets";
      const sheetPath = `xl/${sheetFolder}/sheet${i + 1}.xml`;
      if (ws._isChartSheet)
        entries.push({
          name: sheetPath,
          data: strToBytes(ws.toChartSheetXml()),
        });
      else if (ws._isDialogSheet)
        entries.push({
          name: sheetPath,
          data: strToBytes(ws.toDialogSheetXml(styles, shared)),
        });
      else
        entries.push({
          name: sheetPath,
          data: strToBytes(ws.toXml(styles, shared)),
        });
      const wsRels = [];
      if (ws.drawingRId) {
        const dIdx = this.sheets.filter((s, j) => j <= i && s.drawingRId).length;
        wsRels.push(
          `<Relationship Id="${ws.drawingRId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing${dIdx}.xml"/>`,
        );
      }
      for (let j = 0; j < ws.getImages().length; j++) {
        const g = allImages.filter((x) => x.ws === ws)[j];
        if (g)
          wsRels.push(
            `<Relationship Id="${imgRIds[j]}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image${g.idx}.${g.ext}"/>`,
          );
      }
      for (let j = 0; j < tblEntries.length; j++)
        wsRels.push(
          `<Relationship Id="${tblRIds_[j]}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table${tblEntries[j].globalTableId}.xml"/>`,
        );
      const ptRIds_ = sheetPivotRIds.get(ws) ?? [];
      const ptEntries = allPivotTables.filter((p) => p.ws === ws);
      for (let j = 0; j < ptEntries.length; j++)
        wsRels.push(
          `<Relationship Id="${ptRIds_[j]}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable" Target="../pivotTables/pivotTable${ptEntries[j].pivotIdx}.xml"/>`,
        );
      const slicerInfo = sheetSlicerMap.get(ws);
      if (slicerInfo)
        wsRels.push(
          `<Relationship Id="${slicerInfo.slicerDefRId}" Type="http://schemas.microsoft.com/office/2007/relationships/slicer" Target="../slicers/slicer${slicerInfo.slicerDefIdx}.xml"/>`,
        );
      const sheetComments = ws.getComments();
      const sheetControls = ws.getFormControls();
      if ((sheetComments.length || sheetControls.length) && ws.legacyDrawingRId) {
        if (!buildCtrlPropXml$1)
          ({
            buildCtrlPropXml: buildCtrlPropXml$1,
            buildFormControlVmlShape: buildFormControlVmlShape$1,
            buildVmlWithControls: buildVmlWithControls$1,
          } = await Promise.resolve().then(() => FormControlBuilder_exports));
        const vIdx = vmlCtr++;
        wsRels.push(
          `<Relationship Id="${ws.legacyDrawingRId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing" Target="../drawings/vmlDrawing${vIdx}.vml"/>`,
        );
        if (sheetComments.length) {
          const commRId = `rId${globalRId++}`;
          wsRels.push(
            `<Relationship Id="${commRId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="../comments${vIdx}.xml"/>`,
          );
          entries.push({
            name: `xl/comments${vIdx}.xml`,
            data: strToBytes(this._buildCommentsXml(sheetComments)),
          });
        }
        const commentShapes = sheetComments.map(({ row, col }, ci) => {
          const left = (col + 1) * 64;
          const top = (row - 1) * 20;
          return `<v:shape id="_x0000_s${1025 + i * 1e3 + ci}" type="#_x0000_t202" style="position:absolute;margin-left:${left}pt;margin-top:${top}pt;width:108pt;height:59.25pt;z-index:${ci + 1};visibility:hidden" fillcolor="#ffffe1" o:insetmode="auto">
<v:fill color2="#ffffe1"/>
<v:shadow color="black" obscured="t"/>
<v:path o:connecttype="none"/>
<v:textbox style="mso-direction-alt:auto"><div style="text-align:left"/></v:textbox>
<x:ClientData ObjectType="Note"><x:MoveWithCells/><x:SizeWithCells/><x:Anchor>${col + 1},15,${row - 1},10,${col + 3},15,${row + 4},4</x:Anchor><x:AutoFill>False</x:AutoFill><x:Row>${row - 1}</x:Row><x:Column>${col - 1}</x:Column></x:ClientData>
</v:shape>`;
        });
        const ctrlBaseId = 1025 + ws.sheetIndex * 1e3 + sheetComments.length;
        const controlShapes = sheetControls.map((ctrl, ci) =>
          buildFormControlVmlShape$1(ctrl, ctrlBaseId + ci),
        );
        entries.push({
          name: `xl/drawings/vmlDrawing${vIdx}.vml`,
          data: strToBytes(buildVmlWithControls$1(commentShapes, controlShapes)),
        });
        if (!ws._isDialogSheet) {
          const ctrlRIds = ws.ctrlPropRIds;
          for (let ci = 0; ci < sheetControls.length; ci++) {
            const ctrlPropIdx = ++ctrlPropGlobal;
            wsRels.push(
              `<Relationship Id="${ctrlRIds[ci]}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/ctrlProp" Target="../ctrlProps/ctrlProp${ctrlPropIdx}.xml"/>`,
            );
            entries.push({
              name: `xl/ctrlProps/ctrlProp${ctrlPropIdx}.xml`,
              data: strToBytes(buildCtrlPropXml$1(sheetControls[ci])),
            });
          }
        }
      }
      const oleObjects = ws.getOleObjects();
      for (let oi = 0; oi < oleObjects.length; oi++) {
        const ole = oleObjects[oi];
        const oleIdx = ++oleObjGlobal;
        const oleRId = ws.oleRIds[oi];
        if (ole.linkToFile)
          wsRels.push(
            `<Relationship Id="${oleRId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject" Target="${escapeXml(ole.linkPath ?? "")}" TargetMode="External"/>`,
          );
        else {
          const embName = `xl/embeddings/oleObject${oleIdx}.bin`;
          wsRels.push(
            `<Relationship Id="${oleRId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject" Target="../embeddings/oleObject${oleIdx}.bin"/>`,
          );
          const oleData = !ole.data
            ? /* @__PURE__ */ new Uint8Array(0)
            : typeof ole.data === "string"
              ? base64ToBytes(ole.data)
              : ole.data;
          entries.push({
            name: embName,
            data: oleData,
          });
        }
        if (ole.iconData) {
          const iconExt = ole.iconFormat ?? "emf";
          const iconIdx = imgCtr++;
          const iconRId = ws.oleIconRIds[oi];
          wsRels.push(
            `<Relationship Id="${iconRId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image${iconIdx}.${iconExt}"/>`,
          );
          const iconData =
            typeof ole.iconData === "string" ? base64ToBytes(ole.iconData) : ole.iconData;
          entries.push({
            name: `xl/media/image${iconIdx}.${iconExt}`,
            data: iconData,
          });
        }
      }
      if (wsRels.length)
        entries.push({
          name: `xl/${sheetFolder}/_rels/sheet${i + 1}.xml.rels`,
          data: strToBytes(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
${wsRels.join("\n")}
</Relationships>`),
        });
      if (ws.drawingRId) {
        const dIdx = this.sheets.filter((s, j) => j <= i && s.drawingRId).length;
        entries.push({
          name: `xl/drawings/drawing${dIdx}.xml`,
          data: strToBytes(ws.toDrawingXml(imgRIds, cRIds)),
        });
        const dRels = [];
        for (let j = 0; j < ws.getImages().length; j++) {
          const g = allImages.filter((x) => x.ws === ws)[j];
          if (g)
            dRels.push(
              `<Relationship Id="${imgRIds[j]}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image${g.idx}.${g.ext}"/>`,
            );
        }
        for (let j = 0; j < ws.getCharts().length; j++) {
          const g = allCharts.filter((x) => x.ws === ws)[j];
          if (g)
            dRels.push(
              `<Relationship Id="${cRIds[j]}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart${g.globalIdx}.xml"/>`,
            );
        }
        if (dRels.length)
          entries.push({
            name: `xl/drawings/_rels/drawing${dIdx}.xml.rels`,
            data: strToBytes(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
${dRels.join("\n")}
</Relationships>`),
          });
      }
    }
    for (const { img, ext, idx } of allImages)
      entries.push({
        name: `xl/media/image${idx}.${ext}`,
        data: typeof img.data === "string" ? base64ToBytes(img.data) : img.data,
      });
    if (hasCellImages) {
      const cellImgRIds = [];
      const cellImgRels = [];
      for (let ci = 0; ci < allCellImages.length; ci++) {
        const { img, ext, idx } = allCellImages[ci];
        const rId = `rId${ci + 1}`;
        cellImgRIds.push(rId);
        cellImgRels.push(
          `<Relationship Id="${rId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="/xl/media/image${idx}.${ext}"/>`,
        );
        entries.push({
          name: `xl/media/image${idx}.${ext}`,
          data: typeof img.data === "string" ? base64ToBytes(img.data) : img.data,
        });
      }
      entries.push({
        name: "xl/metadata.xml",
        data: strToBytes(buildMetadataXml(allCellImages.length)),
      });
      entries.push({
        name: "xl/richData/rdrichvalue.xml",
        data: strToBytes(buildRichValueXml(allCellImages.length)),
      });
      entries.push({
        name: "xl/richData/richValueRel.xml",
        data: strToBytes(buildRichValueRelXml(cellImgRIds)),
      });
      entries.push({
        name: "xl/richData/rdRichValueStructure.xml",
        data: strToBytes(buildRichValueStructureXml()),
      });
      entries.push({
        name: "xl/richData/rdRichValueTypes.xml",
        data: strToBytes(buildRichValueTypesXml()),
      });
      entries.push({
        name: "xl/richData/rdarray.xml",
        data: strToBytes(
          `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><arrayData xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata2" count="0"></arrayData>`,
        ),
      });
      entries.push({
        name: "xl/richData/_rels/richValueRel.xml.rels",
        data: strToBytes(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
${cellImgRels.join("\n")}
</Relationships>`),
      });
    }
    for (const { ws, chartIdx, globalIdx } of allCharts) {
      if (!buildChartXml$1)
        buildChartXml$1 = (await Promise.resolve().then(() => ChartBuilder_exports)).buildChartXml;
      entries.push({
        name: `xl/charts/chart${globalIdx}.xml`,
        data: strToBytes(buildChartXml$1(ws.getCharts()[chartIdx])),
      });
    }
    for (const { ws, tableIdx, globalTableId } of allTables)
      entries.push({
        name: `xl/tables/table${globalTableId}.xml`,
        data: strToBytes(buildTableXml(ws.getTables()[tableIdx], globalTableId)),
      });
    for (const { ws, pt, pivotIdx, cacheId: cId } of allPivotTables) {
      if (!buildPivotTableFiles$1)
        buildPivotTableFiles$1 = (await Promise.resolve().then(() => PivotTableBuilder_exports))
          .buildPivotTableFiles;
      const sourceWs = this.sheets.find((s) => s.name === pt.sourceSheet);
      const sourceData = sourceWs ? sourceWs.readRange(pt.sourceRef) : [[]];
      const { pivotTableXml, cacheDefXml, cacheRecordsXml } = buildPivotTableFiles$1(
        pt,
        sourceData,
        pivotIdx,
        cId,
      );
      const wbRel = (type, target) =>
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">\n<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/${type}" Target="${target}"/>\n</Relationships>`;
      entries.push({
        name: `xl/pivotTables/pivotTable${pivotIdx}.xml`,
        data: strToBytes(pivotTableXml),
      });
      entries.push({
        name: `xl/pivotTables/_rels/pivotTable${pivotIdx}.xml.rels`,
        data: strToBytes(
          wbRel("pivotCacheDefinition", `../pivotCache/pivotCacheDefinition${pivotIdx}.xml`),
        ),
      });
      entries.push({
        name: `xl/pivotCache/pivotCacheDefinition${pivotIdx}.xml`,
        data: strToBytes(cacheDefXml),
      });
      entries.push({
        name: `xl/pivotCache/_rels/pivotCacheDefinition${pivotIdx}.xml.rels`,
        data: strToBytes(wbRel("pivotCacheRecords", `pivotCacheRecords${pivotIdx}.xml`)),
      });
      entries.push({
        name: `xl/pivotCache/pivotCacheRecords${pivotIdx}.xml`,
        data: strToBytes(cacheRecordsXml),
      });
    }
    entries.push({
      name: "xl/styles.xml",
      data: strToBytes(styles.toXml()),
    });
    entries.push({
      name: "xl/sharedStrings.xml",
      data: strToBytes(shared.toXml()),
    });
    entries.push({
      name: "xl/theme/theme1.xml",
      data: strToBytes(this._buildThemeXml()),
    });
    for (let i = 0; i < this.externalLinks.length; i++) {
      const link = this.externalLinks[i];
      const idx = i + 1;
      const sheetsXml = link.sheets
        .map((s) => {
          const dNames =
            s.definedNames
              ?.map(
                (d) => `<definedName name="${escapeXml(d.name)}" refersTo="${escapeXml(d.ref)}"/>`,
              )
              .join("") ?? "";
          return `<sheetName val="${escapeXml(s.name)}"/>${dNames ? `<sheetDataSet>${dNames}</sheetDataSet>` : ""}`;
        })
        .join("");
      entries.push({
        name: `xl/externalLinks/externalLink${idx}.xml`,
        data: strToBytes(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<externalLink xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<externalBook r:id="rId1"><sheetNames>${sheetsXml}</sheetNames></externalBook>
</externalLink>`),
      });
      entries.push({
        name: `xl/externalLinks/_rels/externalLink${idx}.xml.rels`,
        data: strToBytes(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath" Target="${escapeXml(link.target)}" TargetMode="External"/>
</Relationships>`),
      });
    }
    for (const [ws, info] of sheetSlicerMap) {
      const allSheetSlicerItems = [];
      for (const ts of info.tableSlicers) {
        const s = ts.slicer;
        allSheetSlicerItems.push(
          `<slicer name="${escapeXml(s.name)}" cache="${escapeXml(s.name + "_cache")}" caption="${escapeXml(s.caption ?? s.columnName)}" rowHeight="241300" columnCount="${s.columnCount ?? 1}" style="${s.style ?? "SlicerStyleLight1"}"/>`,
        );
      }
      for (const ps of info.pivotSlicers) {
        const s = ps.slicer;
        allSheetSlicerItems.push(
          `<slicer name="${escapeXml(s.name)}" cache="${escapeXml(s.name + "_cache")}" caption="${escapeXml(s.caption ?? s.fieldName)}" rowHeight="241300" columnCount="${s.columnCount ?? 1}" style="${s.style ?? "SlicerStyleLight1"}"/>`,
        );
      }
      entries.push({
        name: `xl/slicers/slicer${info.slicerDefIdx}.xml`,
        data: strToBytes(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<slicers xmlns="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main" mc:Ignorable="x">
${allSheetSlicerItems.join("\n")}
</slicers>`),
      });
    }
    for (const sc of allSlicerCaches) {
      let cacheBody;
      if (sc.type === "table")
        cacheBody = `<extLst><x:ext uri="{2F2917AC-EB37-4324-AD4E-5DD8C200BD13}" xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main"><x15:tableSlicerCache tableId="${sc.tableId}" column="${sc.columnIndex}" sortOrder="${sc.sortOrder ?? "ascending"}"/></x:ext></extLst>`;
      else {
        const itemsXml = (sc.items ?? []).map((_, xi) => `<i x="${xi}" s="1"/>`).join("");
        cacheBody =
          `<pivotTables><pivotTable tabId="${sc.tabId}" name="${escapeXml(sc.pivotTableName ?? "")}"/></pivotTables>` +
          (sc.items?.length
            ? `<data><tabular pivotCacheId="${sc.pivotCacheId}" showMissing="0"><items count="${sc.items.length}">${itemsXml}</items></tabular></data>`
            : "");
      }
      entries.push({
        name: `xl/slicerCaches/slicerCache${sc.idx}.xml`,
        data: strToBytes(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<slicerCacheDefinition xmlns="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x" name="${escapeXml(sc.name)}" sourceName="${escapeXml(sc.sourceName)}">
${cacheBody}
</slicerCacheDefinition>`),
      });
    }
    const allQueryTables = this.sheets.flatMap((ws) => ws.getQueryTables());
    for (let i = 0; i < allQueryTables.length; i++) {
      const qt = allQueryTables[i];
      entries.push({
        name: `xl/queryTables/queryTable${i + 1}.xml`,
        data: strToBytes(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<queryTable xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" name="${escapeXml(qt.name)}" connectionId="${qt.connectionId}" autoFormatId="16" applyNumberFormats="0" applyBorderFormats="0" applyFontFormats="0" applyPatternFormats="0" applyAlignmentFormats="0" applyWidthHeightFormats="0">
<queryTableRefresh nextId="${(qt.columns?.length ?? 0) + 1}">
<queryTableFields count="${qt.columns?.length ?? 0}">
${(qt.columns ?? []).map((c, ci) => `<queryTableField id="${ci + 1}" name="${escapeXml(c)}" tableColumnId="${ci + 1}"/>`).join("\n")}
</queryTableFields>
</queryTableRefresh>
</queryTable>`),
      });
    }
    const cp = {
      ...this.coreProperties,
      created: this.coreProperties.created ?? /* @__PURE__ */ new Date(),
      modified: /* @__PURE__ */ new Date(),
    };
    if (!cp.creator && this.properties.author) cp.creator = this.properties.author;
    entries.push({
      name: "docProps/core.xml",
      data: strToBytes(buildCoreXml(cp)),
    });
    entries.push({
      name: "docProps/app.xml",
      data: strToBytes(
        buildAppXml({
          ...this.extendedProperties,
          application: this.extendedProperties.application ?? "ExcelForge",
          company: this.extendedProperties.company ?? this.properties.company,
          titlesOfParts: this.sheets.map((s) => s.name),
          headingPairs: this._headingPairs(),
        }),
      ),
    });
    if (hasCustom)
      entries.push({
        name: "docProps/custom.xml",
        data: strToBytes(buildCustomXml(this.customProperties)),
      });
    return buildZip(entries, { level: this.compressionLevel });
  }
  _headingPairs() {
    const normalCount = this.sheets.filter((ws) => !ws._isChartSheet && !ws._isDialogSheet).length;
    const chartCount = this.sheets.filter((ws) => ws._isChartSheet).length;
    const dialogCount = this.sheets.filter((ws) => ws._isDialogSheet).length;
    const pairs = [];
    if (normalCount)
      pairs.push({
        name: "Worksheets",
        count: normalCount,
      });
    if (chartCount)
      pairs.push({
        name: "Charts",
        count: chartCount,
      });
    if (dialogCount)
      pairs.push({
        name: "Dialogs",
        count: dialogCount,
      });
    if (!pairs.length)
      pairs.push({
        name: "Worksheets",
        count: 0,
      });
    return pairs;
  }
  _syncLegacyProperties() {
    const p = this.properties;
    if (p.title) this.coreProperties.title ??= p.title;
    if (p.author) this.coreProperties.creator ??= p.author;
    if (p.subject) this.coreProperties.subject ??= p.subject;
    if (p.description) this.coreProperties.description ??= p.description;
    if (p.keywords) this.coreProperties.keywords ??= p.keywords;
    if (p.company) this.extendedProperties.company ??= p.company;
    if (p.lastModifiedBy) this.coreProperties.lastModifiedBy ??= p.lastModifiedBy;
    if (p.created) this.coreProperties.created ??= p.created;
    if (p.category) this.coreProperties.category ??= p.category;
    if (p.status) this.coreProperties.contentStatus ??= p.status;
  }
  _ensureVbaSheetModules() {
    if (!this.vbaProject) return;
    if (
      this.vbaProject.modules.filter((m) => m.type === "document" && m.name !== "ThisWorkbook")
        .length >= this.sheets.length
    )
      return;
    for (const ws of this.sheets) {
      const sheetCodeName = ws.name.replace(/[^A-Za-z0-9_]/g, "_");
      if (!this.vbaProject.getModule(sheetCodeName))
        this.vbaProject.addModule({
          name: sheetCodeName,
          type: "document",
          code: "",
        });
    }
  }
  _patchWorkbookXml(originalXml, rr) {
    let xml = originalXml;
    const origCount = rr?.sheets.length ?? this.sheets.length;
    for (let i = 0; i < Math.min(this.sheets.length, origCount); i++)
      xml = xml.replace(
        new RegExp(`(<sheet[^>]+sheetId="${i + 1}"[^>]+)name="[^"]*"`),
        `$1name="${escapeXml(this.sheets[i].name)}"`,
      );
    if (rr && this.sheets.length > origCount) {
      let maxRId = 0;
      for (const [id] of rr.workbookRels) {
        const n = parseInt(id.replace(/\D/g, ""), 10);
        if (n > maxRId) maxRId = n;
      }
      let maxSheetId = 0;
      for (const s of rr.sheets) {
        const n = parseInt(s.sheetId, 10);
        if (n > maxSheetId) maxSheetId = n;
      }
      const newSheetTags = [];
      for (let i = origCount; i < this.sheets.length; i++) {
        const ws = this.sheets[i];
        const rId = `rId${++maxRId}`;
        const sheetId = ++maxSheetId;
        ws.rId = rId;
        ws.sheetIndex = sheetId;
        const stateAttr =
          ws.options?.state === "hidden"
            ? ' state="hidden"'
            : ws.options?.state === "veryHidden"
              ? ' state="veryHidden"'
              : "";
        newSheetTags.push(
          `<sheet name="${escapeXml(ws.name)}" sheetId="${sheetId}" r:id="${rId}"${stateAttr}/>`,
        );
      }
      xml = xml.replace("</sheets>", newSheetTags.join("") + "</sheets>");
    }
    if (this.vbaProject && !xml.includes("codeName=")) {
      xml = xml.replace("<workbookPr", '<workbookPr codeName="ThisWorkbook"');
      if (!xml.includes("<workbookPr"))
        xml = xml.replace("<bookViews", '<workbookPr codeName="ThisWorkbook"/><bookViews');
    }
    const dnXml = this._definedNamesXml();
    if (xml.includes("<definedNames"))
      xml = xml.replace(/<definedNames[\s\S]*?<\/definedNames>/, dnXml);
    else if (dnXml) xml = xml.replace("</sheets>", `</sheets>${dnXml}`);
    if (this.calcSettings) {
      const newCalcPr = this._calcPrXml();
      if (xml.includes("<calcPr"))
        xml = xml.replace(/<calcPr[^>]*\/>|<calcPr[^>]*>[\s\S]*?<\/calcPr>/, newCalcPr);
      else xml = xml.replace("</workbook>", `${newCalcPr}</workbook>`);
    }
    return xml;
  }
  _definedNamesXml(slicerCaches) {
    const printAreaNames = [];
    for (const ws of this.sheets)
      if (ws.printArea) {
        const ref = ws.printArea.includes("!") ? ws.printArea : `'${ws.name}'!${ws.printArea}`;
        printAreaNames.push({
          name: "_xlnm.Print_Area",
          ref,
          scope: ws.name,
        });
      }
    const slicerNames = (slicerCaches ?? []).map((sc) => ({
      name: sc.name,
      ref: "#N/A",
    }));
    const all = [...this.namedRanges, ...printAreaNames, ...slicerNames];
    if (!all.length) return "";
    return `<definedNames>${all
      .map((nr) => {
        let attrs = `name="${escapeXml(nr.name)}"`;
        if (nr.scope) {
          const idx = this.sheets.findIndex((s) => s.name === nr.scope);
          if (idx >= 0) attrs += ` localSheetId="${idx}"`;
        }
        if (nr.comment) attrs += ` comment="${escapeXml(nr.comment)}"`;
        return `<definedName ${attrs}>${escapeXml(nr.ref)}</definedName>`;
      })
      .join("")}</definedNames>`;
  }
  _calcPrXml() {
    const cs = this.calcSettings;
    let attrs = 'calcId="191028"';
    if (cs) {
      if (cs.calcMode === "manual") attrs += ' calcMode="manual"';
      else if (cs.calcMode === "autoNoTable") attrs += ' calcMode="autoNoTable"';
      if (cs.fullCalcOnLoad) attrs += ' fullCalcOnLoad="1"';
      if (cs.iterate) {
        attrs += ' iterate="1"';
        if (cs.iterateCount != null) attrs += ` iterateCount="${cs.iterateCount}"`;
        if (cs.iterateDelta != null) attrs += ` iterateDelta="${cs.iterateDelta}"`;
      }
      if (cs.fullPrecision === false) attrs += ' fullPrecision="0"';
      if (cs.calcOnSave === false) attrs += ' calcOnSave="0"';
      if (cs.concurrentCalc === false) attrs += ' concurrentCalc="0"';
    }
    return `<calcPr ${attrs}/>`;
  }
  _connectionsXml(originalXml) {
    if (!this.connections.length) return "";
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<connections xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">${this.connections
      .map((c) => {
        if (c._rawXml) return c._rawXml;
        let attrs = ` id="${c.id}" name="${escapeXml(c.name)}" type="${connTypeToNum(c.type)}" refreshedVersion="6"`;
        if (c.description) attrs += ` description="${escapeXml(c.description)}"`;
        if (c.refreshOnLoad) attrs += ' refreshOnLoad="1"';
        if (c.background) attrs += ' background="1"';
        if (c.saveData) attrs += ' saveData="1"';
        if (c.keepAlive) attrs += ' keepAlive="1"';
        if (c.interval) attrs += ` interval="${c.interval}"`;
        const dbPr =
          c.connectionString || c.command
            ? `<dbPr${c.connectionString ? ` connection="${escapeXml(c.connectionString)}"` : ""}${c.command ? ` command="${escapeXml(c.command)}"` : ""}${c.commandType ? ` commandType="${cmdTypeToNum(c.commandType)}"` : ""}/>`
            : "";
        return `<connection${attrs}>${dbPr}</connection>`;
      })
      .join("")}</connections>`;
  }
  _buildWorkbookRels(rr, dropCalcChain = false) {
    const rels = [...rr.workbookRels.entries()]
      .filter(([_, rel]) => !(dropCalcChain && rel.type.includes("/calcChain")))
      .map(([id, rel]) => `<Relationship Id="${id}" Type="${rel.type}" Target="${rel.target}"/>`);
    const origCount = rr.sheets.length;
    for (let i = origCount; i < this.sheets.length; i++) {
      const ws = this.sheets[i];
      const folder = ws._isChartSheet
        ? "chartsheets"
        : ws._isDialogSheet
          ? "dialogsheets"
          : "worksheets";
      const type = ws._isChartSheet
        ? "chartsheet"
        : ws._isDialogSheet
          ? "dialogsheet"
          : "worksheet";
      rels.push(
        `<Relationship Id="${ws.rId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/${type}" Target="${folder}/sheet${i + 1}.xml"/>`,
      );
    }
    if (![...rr.workbookRels.values()].some((r) => r.type.includes("/styles")))
      rels.push(
        `<Relationship Id="rIdStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>`,
      );
    if (![...rr.workbookRels.values()].some((r) => r.type.includes("/sharedStrings")))
      rels.push(
        `<Relationship Id="rIdShared" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>`,
      );
    if (
      this.vbaProject &&
      ![...rr.workbookRels.values()].some((r) => r.type.includes("vbaProject"))
    )
      rels.push(
        `<Relationship Id="rIdVBA" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" Target="vbaProject.bin"/>`,
      );
    if (
      this.connections.length &&
      ![...rr.workbookRels.values()].some((r) => r.type.includes("/connections"))
    )
      rels.push(
        `<Relationship Id="rIdConns" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/connections" Target="connections.xml"/>`,
      );
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
${rels.join("\n")}
</Relationships>`;
  }
  _relsToXml(relMap) {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
${[...relMap.entries()].map(([id, r]) => `<Relationship Id="${escapeXml(id)}" Type="${escapeXml(r.type)}" Target="${escapeXml(r.target)}"${r.targetMode ? ` TargetMode="${escapeXml(r.targetMode)}"` : ""}/>`).join("\n")}
</Relationships>`;
  }
  _buildThemeXml() {
    const t = this.theme;
    const majorFont = t?.majorFont ?? "Calibri Light";
    const minorFont = t?.minorFont ?? "Calibri";
    const colorElements = (
      t?.colors?.map((c) => {
        let hex = c.color.replace(/^#/, "");
        if (hex.length === 8) hex = hex.substring(2);
        return {
          name: c.name,
          color: hex,
        };
      }) ?? [
        {
          name: "dk1",
          color: "000000",
        },
        {
          name: "lt1",
          color: "FFFFFF",
        },
        {
          name: "dk2",
          color: "44546A",
        },
        {
          name: "lt2",
          color: "E7E6E6",
        },
        {
          name: "accent1",
          color: "4472C4",
        },
        {
          name: "accent2",
          color: "ED7D31",
        },
        {
          name: "accent3",
          color: "A5A5A5",
        },
        {
          name: "accent4",
          color: "FFC000",
        },
        {
          name: "accent5",
          color: "5B9BD5",
        },
        {
          name: "accent6",
          color: "70AD47",
        },
        {
          name: "hlink",
          color: "0563C1",
        },
        {
          name: "folHlink",
          color: "954F72",
        },
      ]
    )
      .map((c) => {
        if (c.name === "dk1" || c.name === "lt1")
          return `<a:${c.name}><a:sysClr val="${c.name === "dk1" ? "windowText" : "window"}" lastClr="${c.color}"/></a:${c.name}>`;
        return `<a:${c.name}><a:srgbClr val="${c.color}"/></a:${c.name}>`;
      })
      .join("");
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="${escapeXml(t?.name ?? "Office Theme")}">
<a:themeElements>
<a:clrScheme name="Office">${colorElements}</a:clrScheme>
<a:fontScheme name="Office">
<a:majorFont><a:latin typeface="${escapeXml(majorFont)}"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont>
<a:minorFont><a:latin typeface="${escapeXml(minorFont)}"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont>
</a:fontScheme>
<a:fmtScheme name="Office">
<a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:lumMod val="110000"/><a:satMod val="105000"/><a:tint val="67000"/></a:schemeClr></a:gs><a:gs pos="50000"><a:schemeClr val="phClr"><a:lumMod val="105000"/><a:satMod val="103000"/><a:tint val="73000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:lumMod val="105000"/><a:satMod val="109000"/><a:tint val="81000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:satMod val="103000"/><a:lumMod val="102000"/><a:tint val="94000"/></a:schemeClr></a:gs><a:gs pos="50000"><a:schemeClr val="phClr"><a:satMod val="110000"/><a:lumMod val="100000"/><a:shade val="100000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:lumMod val="99000"/><a:satMod val="120000"/><a:shade val="78000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill></a:fillStyleLst>
<a:lnStyleLst><a:ln w="6350" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln><a:ln w="12700" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln><a:ln w="19050" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln></a:lnStyleLst>
<a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst><a:outerShdw blurRad="57150" dist="19050" dir="5400000" algn="ctr" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="63000"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle></a:effectStyleLst>
<a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"><a:tint val="95000"/><a:satMod val="170000"/></a:schemeClr></a:solidFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="93000"/><a:satMod val="150000"/><a:shade val="98000"/><a:lumMod val="102000"/></a:schemeClr></a:gs><a:gs pos="50000"><a:schemeClr val="phClr"><a:tint val="98000"/><a:satMod val="130000"/><a:shade val="90000"/><a:lumMod val="103000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="63000"/><a:satMod val="120000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill></a:bgFillStyleLst>
</a:fmtScheme>
</a:themeElements>
<a:objectDefaults/>
<a:extraClrSchemeLst/>
</a:theme>`;
  }
  _buildRootRels(hasCustom) {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
${hasCustom ? `<Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties" Target="docProps/custom.xml"/>` : ""}
</Relationships>`;
  }
  _patchContentTypes(originalXml, addCustom, dropCalcChain = false) {
    let xml = originalXml;
    if (dropCalcChain) xml = xml.replace(/<Override[^>]*calcChain[^>]*\/>/g, "");
    if (!xml.includes("sharedStrings"))
      xml = xml.replace(
        "</Types>",
        `<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>\n</Types>`,
      );
    if (addCustom && !xml.includes("custom-properties"))
      xml = xml.replace(
        "</Types>",
        `<Override PartName="/docProps/custom.xml" ContentType="application/vnd.openxmlformats-officedocument.custom-properties+xml"/>\n</Types>`,
      );
    if (this.vbaProject) {
      if (!xml.includes("vbaProject.bin"))
        xml = xml.replace(
          "</Types>",
          `<Override PartName="/xl/vbaProject.bin" ContentType="application/vnd.ms-office.vbaProject"/>\n</Types>`,
        );
      xml = xml.replace(
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
        "application/vnd.ms-excel.sheet.macroEnabled.main+xml",
      );
    }
    if (this.isTemplate && !this.vbaProject)
      xml = xml.replace(
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml",
      );
    if (this.connections.length && !xml.includes("connections.xml"))
      xml = xml.replace(
        "</Types>",
        `<Override PartName="/xl/connections.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.connections+xml"/>\n</Types>`,
      );
    for (let i = 0; i < this.sheets.length; i++) {
      const ws = this.sheets[i];
      const folder = ws._isChartSheet
        ? "chartsheets"
        : ws._isDialogSheet
          ? "dialogsheets"
          : "worksheets";
      const ct = ws._isChartSheet
        ? "application/vnd.openxmlformats-officedocument.spreadsheetml.chartsheet+xml"
        : ws._isDialogSheet
          ? "application/vnd.openxmlformats-officedocument.spreadsheetml.dialogsheet+xml"
          : "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml";
      const partName = `/xl/${folder}/sheet${i + 1}.xml`;
      if (!xml.includes(partName))
        xml = xml.replace(
          "</Types>",
          `<Override PartName="${partName}" ContentType="${ct}"/>\n</Types>`,
        );
    }
    return xml;
  }
  async buildBase64() {
    const bytes = await this.build();
    let bin = "";
    for (let i = 0; i < bytes.length; i++) bin += String.fromCharCode(bytes[i]);
    return btoa(bin);
  }
  async writeFile(path) {
    const bytes = await this.build();
    await (await import("fs/promises")).writeFile(path, bytes);
  }
  _buildCommentsXml(comments) {
    const authors = [...new Set(comments.map((c) => c.comment.author ?? ""))];
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<authors>${authors.map((a) => `<author>${escapeXml(a)}</author>`).join("")}</authors>
<commentList>${comments
      .map(({ row, col, comment }) => {
        const ref = `${colIndexToLetter(col)}${row}`;
        const authorIdx = authors.indexOf(comment.author ?? "");
        let textXml;
        if (comment.richText && comment.richText.length > 0)
          textXml = comment.richText
            .map((run) => {
              let rPr = "";
              if (run.font) {
                const f = run.font;
                if (f.bold) rPr += "<b/>";
                if (f.italic) rPr += "<i/>";
                if (f.underline && f.underline !== "none")
                  rPr += `<u val="${f.underline === "single" ? "single" : f.underline}"/>`;
                if (f.strike) rPr += "<strike/>";
                if (f.size) rPr += `<sz val="${f.size}"/>`;
                if (f.color) rPr += `<color rgb="${f.color}"/>`;
                if (f.name) rPr += `<rFont val="${escapeXml(f.name)}"/>`;
                if (f.family != null) rPr += `<family val="${f.family}"/>`;
              }
              return `<r>${rPr ? `<rPr>${rPr}</rPr>` : ""}<t xml:space="preserve">${escapeXml(run.text)}</t></r>`;
            })
            .join("");
        else textXml = `<r><t>${escapeXml(comment.text)}</t></r>`;
        return `<comment ref="${ref}" authorId="${authorIdx}"><text>${textXml}</text></comment>`;
      })
      .join("")}</commentList>
</comments>`;
  }
  _buildVmlXml(comments, sheetIdx) {
    return `<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel">
<o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="1"/></o:shapelayout>
<v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe"><v:stroke joinstyle="miter"/><v:path gradientshapeok="t" o:connecttype="rect"/></v:shapetype>
${comments
  .map(({ row, col }, i) => {
    const left = (col + 1) * 64;
    const top = (row - 1) * 20;
    return `<v:shape id="_x0000_s${1025 + sheetIdx * 1e3 + i}" type="#_x0000_t202" style="position:absolute;margin-left:${left}pt;margin-top:${top}pt;width:108pt;height:59.25pt;z-index:${i + 1};visibility:hidden" fillcolor="#ffffe1" o:insetmode="auto">
<v:fill color2="#ffffe1"/>
<v:shadow color="black" obscured="t"/>
<v:path o:connecttype="none"/>
<v:textbox style="mso-direction-alt:auto"><div style="text-align:left"/></v:textbox>
<x:ClientData ObjectType="Note"><x:MoveWithCells/><x:SizeWithCells/><x:Anchor>${col + 1},15,${row - 1},10,${col + 3},15,${row + 4},4</x:Anchor><x:AutoFill>False</x:AutoFill><x:Row>${row - 1}</x:Row><x:Column>${col - 1}</x:Column></x:ClientData>
</v:shape>`;
  })
  .join("\n")}
</xml>`;
  }
  async download(filename = "workbook.xlsx") {
    const bytes = await this.build();
    const blob = new Blob([bytes.buffer], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = filename;
    a.click();
    URL.revokeObjectURL(url);
  }
};
function imageContentType(ext) {
  switch (ext) {
    case "jpg":
      return "image/jpeg";
    case "png":
      return "image/png";
    case "gif":
      return "image/gif";
    case "bmp":
      return "image/bmp";
    case "tiff":
      return "image/tiff";
    case "emf":
      return "image/x-emf";
    case "wmf":
      return "image/x-wmf";
    case "svg":
      return "image/svg+xml";
    case "ico":
      return "image/x-icon";
    case "webp":
      return "image/webp";
    default:
      return `image/${ext}`;
  }
}
function imageExt(format) {
  return format === "jpeg" ? "jpg" : format;
}
function buildMetadataXml(count) {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:xlrd="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata"><metadataTypes count="1"><metadataType name="XLRICHVALUE" minSupportedVersion="120000" copy="1" pasteAll="1" pasteValues="1" merge="1" splitFirst="1" rowColShift="1" clearFormats="1" clearComments="1" assign="1" coerce="1"/></metadataTypes><futureMetadata name="XLRICHVALUE" count="${count}">${Array.from({ length: count }, (_, i) => `<bk><extLst><ext uri="{3e2802c4-a4d2-4d8b-9148-e3be6c30e623}"><xlrd:rvb i="${i}"/></ext></extLst></bk>`).join("")}</futureMetadata><valueMetadata count="${count}">${Array.from({ length: count }, (_, i) => `<bk><rc t="1" v="${i}"/></bk>`).join("")}</valueMetadata></metadata>`;
}
function buildRichValueXml(count) {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><rvData xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata" count="${count}">${Array.from({ length: count }, (_, i) => `<rv s="0"><v>${i}</v><v>5</v><v></v></rv>`).join("")}</rvData>`;
}
function buildRichValueRelXml(rIds) {
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><richValueRels xmlns="http://schemas.microsoft.com/office/spreadsheetml/2022/richvaluerel" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
    rIds.map((id) => `<rel r:id="${id}"/>`).join("") +
    `</richValueRels>`
  );
}
function buildRichValueStructureXml() {
  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><rvStructures xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata" count="1"><s t="_localImage"><k n="_rvRel:LocalImageIdentifier" t="i"/><k n="CalcOrigin" t="i"/><k n="Text" t="s"/></s></rvStructures>';
}
function buildRichValueTypesXml() {
  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><rvTypesInfo xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata2" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x" xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><global><keyFlags><key name="_Self"><flag name="ExcludeFromFile" value="1"/><flag name="ExcludeFromCalcComparison" value="1"/></key><key name="_DisplayString"><flag name="ExcludeFromCalcComparison" value="1"/></key><key name="_Flags"><flag name="ExcludeFromCalcComparison" value="1"/></key><key name="_Format"><flag name="ExcludeFromCalcComparison" value="1"/></key><key name="_SubLabel"><flag name="ExcludeFromCalcComparison" value="1"/></key><key name="_Attribution"><flag name="ExcludeFromCalcComparison" value="1"/></key><key name="_Icon"><flag name="ExcludeFromCalcComparison" value="1"/></key><key name="_Display"><flag name="ExcludeFromCalcComparison" value="1"/></key><key name="_CanonicalPropertyNames"><flag name="ExcludeFromCalcComparison" value="1"/></key><key name="_ClassificationId"><flag name="ExcludeFromCalcComparison" value="1"/></key></keyFlags></global><types></types></rvTypesInfo>';
}
//#endregion
//#region node_modules/@node-projects/excelforge/dist/styles/builders.js
var StyleBuilder = class {
  constructor() {
    this.s = {};
  }
  font(props) {
    this.s.font = {
      ...(this.s.font ?? {}),
      ...props,
    };
    return this;
  }
  fontSize(size) {
    return this.font({ size });
  }
  fontName(name) {
    return this.font({ name });
  }
  bold(v = true) {
    return this.font({ bold: v });
  }
  italic(v = true) {
    return this.font({ italic: v });
  }
  strike(v = true) {
    return this.font({ strike: v });
  }
  underline(style = "single") {
    return this.font({ underline: style });
  }
  fontColor(color) {
    return this.font({ color });
  }
  fill(pattern, fgColor, bgColor) {
    this.s.fill = {
      type: "pattern",
      pattern,
      fgColor,
      bgColor,
    };
    return this;
  }
  bg(color) {
    return this.fill("solid", color);
  }
  border(style, color) {
    const side = {
      style,
      color,
    };
    this.s.border = {
      left: side,
      right: side,
      top: side,
      bottom: side,
    };
    return this;
  }
  borderLeft(style, color) {
    this.s.border = {
      ...(this.s.border ?? {}),
      left: {
        style,
        color,
      },
    };
    return this;
  }
  borderRight(style, color) {
    this.s.border = {
      ...(this.s.border ?? {}),
      right: {
        style,
        color,
      },
    };
    return this;
  }
  borderTop(style, color) {
    this.s.border = {
      ...(this.s.border ?? {}),
      top: {
        style,
        color,
      },
    };
    return this;
  }
  borderBottom(style, color) {
    this.s.border = {
      ...(this.s.border ?? {}),
      bottom: {
        style,
        color,
      },
    };
    return this;
  }
  align(horizontal, vertical) {
    this.s.alignment = {
      ...(this.s.alignment ?? {}),
      horizontal,
      vertical,
    };
    return this;
  }
  center() {
    return this.align("center", "center");
  }
  wrap(v = true) {
    this.s.alignment = {
      ...(this.s.alignment ?? {}),
      wrapText: v,
    };
    return this;
  }
  rotate(degrees) {
    this.s.alignment = {
      ...(this.s.alignment ?? {}),
      textRotation: degrees,
    };
    return this;
  }
  indent(n) {
    this.s.alignment = {
      ...(this.s.alignment ?? {}),
      indent: n,
    };
    return this;
  }
  shrink(v = true) {
    this.s.alignment = {
      ...(this.s.alignment ?? {}),
      shrinkToFit: v,
    };
    return this;
  }
  numFmt(formatCode) {
    this.s.numberFormat = { formatCode };
    return this;
  }
  numFmtId(id) {
    this.s.numFmtId = id;
    return this;
  }
  locked(v = true) {
    this.s.locked = v;
    return this;
  }
  hidden(v = true) {
    this.s.hidden = v;
    return this;
  }
  build() {
    return { ...this.s };
  }
};
function style() {
  return new StyleBuilder();
}
(style().bold().bg("FF4472C4").fontColor("FFFFFFFF").center().build(),
  style().bold().bg("FF595959").fontColor("FFFFFFFF").center().build(),
  style().bold().bg("FF70AD47").fontColor("FFFFFFFF").center().build(),
  style().numFmt('"$"#,##0.00').build(),
  style().numFmt("0.00%").build(),
  style().numFmt("#,##0").build(),
  style().numFmtId(14).build(),
  style().numFmtId(22).build(),
  style().bold().build(),
  style().center().build(),
  style().wrap().build(),
  style().border("thin").build(),
  style().border("thin", "FF000000").build(),
  style().bg("FFFFFF00").build(),
  style().bg("FFDCE6F1").build(),
  style().fontColor("FFFF0000").build(),
  style().fontColor("FF0000FF").build(),
  style().fontColor("FF00B050").build(),
  style().bold().border("thin").center().bg("FFD9E1F2").build(),
  style().strike().fontColor("FF808080").build());
new TextEncoder();
new Uint8Array([208, 207, 17, 224, 161, 177, 26, 225]);
new Uint8Array([254, 167, 210, 118, 59, 75, 158, 121]);
new Uint8Array([215, 170, 15, 109, 48, 97, 52, 78]);
new Uint8Array([20, 110, 11, 231, 171, 172, 208, 214]);
new Uint8Array([95, 178, 173, 1, 12, 185, 225, 246]);
new Uint8Array([160, 103, 127, 2, 178, 44, 132, 51]);
//#endregion
//#region node_modules/@node-projects/excelforge/dist/features/ChartBuilder.js
var ChartBuilder_exports = /* @__PURE__ */ __exportAll({ buildChartXml: () => buildChartXml });
var COLORS = [
  "FF4472C4",
  "FFED7D31",
  "FFA5A5A5",
  "FFFFC000",
  "FF5B9BD5",
  "FF70AD47",
  "FF264478",
  "FF9E480E",
  "FF636363",
  "FF997300",
];
var PALETTES = {
  office: [
    "4472C4",
    "ED7D31",
    "A5A5A5",
    "FFC000",
    "5B9BD5",
    "70AD47",
    "264478",
    "9E480E",
    "636363",
    "997300",
  ],
  blue: [
    "5B9BD5",
    "8FAADC",
    "B4C7E7",
    "D6DCE4",
    "4472C4",
    "2F5597",
    "1F4E79",
    "002060",
    "00B0F0",
    "0070C0",
  ],
  orange: [
    "ED7D31",
    "F4B183",
    "F8CBAD",
    "FCE4D6",
    "C55A11",
    "843C0C",
    "BF8F00",
    "806000",
    "FFD966",
    "FFC000",
  ],
  green: [
    "70AD47",
    "A9D18E",
    "C5E0B4",
    "E2F0D9",
    "548235",
    "375623",
    "92D050",
    "00B050",
    "00B0F0",
    "00FF00",
  ],
  red: [
    "FF0000",
    "FF6161",
    "FF9F9F",
    "FFD2D2",
    "C00000",
    "8B0000",
    "FFC000",
    "FF6600",
    "FF3300",
    "CC0000",
  ],
  purple: [
    "7030A0",
    "9B59B6",
    "BB8FCE",
    "D7BDE2",
    "5B2C6F",
    "4A235A",
    "8E44AD",
    "6C3483",
    "A569BD",
    "D2B4DE",
  ],
  teal: [
    "00B0F0",
    "00BCD4",
    "26C6DA",
    "80DEEA",
    "0097A7",
    "006064",
    "009688",
    "00796B",
    "4DB6AC",
    "80CBC4",
  ],
  gray: [
    "595959",
    "808080",
    "A6A6A6",
    "D9D9D9",
    "404040",
    "262626",
    "BFBFBF",
    "F2F2F2",
    "7F7F7F",
    "C0C0C0",
  ],
  gold: [
    "FFC000",
    "FFD966",
    "FFE699",
    "FFF2CC",
    "BF8F00",
    "806000",
    "ED7D31",
    "F4B183",
    "DBA13A",
    "C09100",
  ],
  blueWarm: [
    "4472C4",
    "5B9BD5",
    "9DC3E6",
    "BDD7EE",
    "DEEBF7",
    "2E75B6",
    "1F4E79",
    "2F5597",
    "8FAADC",
    "D6DCE4",
  ],
  blueGreen: [
    "00B0F0",
    "00B050",
    "70AD47",
    "5B9BD5",
    "4472C4",
    "2E75B6",
    "548235",
    "00B0F0",
    "009688",
    "00796B",
  ],
  greenYellow: [
    "70AD47",
    "92D050",
    "C9E265",
    "FFD966",
    "FFC000",
    "548235",
    "BF8F00",
    "A9D18E",
    "E2F0D9",
    "C5E0B4",
  ],
  redOrange: [
    "FF0000",
    "ED7D31",
    "FFC000",
    "FF6600",
    "C00000",
    "843C0C",
    "BF8F00",
    "FF3300",
    "F4B183",
    "FFD966",
  ],
  redViolet: [
    "FF0000",
    "7030A0",
    "ED7D31",
    "C00000",
    "5B2C6F",
    "843C0C",
    "FF6161",
    "9B59B6",
    "F4B183",
    "BB8FCE",
  ],
  yellowOrange: [
    "FFC000",
    "ED7D31",
    "FFD966",
    "F4B183",
    "BF8F00",
    "C55A11",
    "FFE699",
    "F8CBAD",
    "806000",
    "843C0C",
  ],
  slipstream: [
    "4472C4",
    "ED7D31",
    "A5A5A5",
    "FFC000",
    "5B9BD5",
    "70AD47",
    "255E91",
    "9E480E",
    "636363",
    "997300",
  ],
  marquee: [
    "ED7D31",
    "4472C4",
    "70AD47",
    "FFC000",
    "5B9BD5",
    "A5A5A5",
    "264478",
    "9E480E",
    "636363",
    "997300",
  ],
  aspect: [
    "5B9BD5",
    "4472C4",
    "ED7D31",
    "FFC000",
    "70AD47",
    "A5A5A5",
    "9E480E",
    "264478",
    "636363",
    "997300",
  ],
};
var MODERN_STYLE_MAP = {
  colorful1: 102,
  colorful2: 103,
  colorful3: 104,
  colorful4: 105,
  monochromatic1: 201,
  monochromatic2: 202,
  monochromatic3: 203,
  monochromatic4: 204,
  monochromatic5: 205,
  monochromatic6: 206,
  monochromatic7: 207,
  monochromatic8: 208,
  monochromatic9: 209,
  monochromatic10: 210,
  monochromatic11: 211,
  monochromatic12: 212,
};
function spPr(color, series) {
  const aNs = ' xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"';
  const hex = color.replace(/^FF/, "");
  let fill;
  if (series?.fillType === "gradient" && series.gradientStops?.length)
    fill = `<a:gradFill><a:gsLst>${series.gradientStops.map((s) => `<a:gs pos="${s.pos * 1e3}"><a:srgbClr val="${s.color.replace(/^FF/, "").replace(/^#/, "")}"/></a:gs>`).join("")}</a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill>`;
  else fill = `<a:solidFill><a:srgbClr val="${hex}"/></a:solidFill>`;
  const lnWidth = series?.lineWidth ? ` w="${Math.round(series.lineWidth * 12700)}"` : "";
  const ln = lnWidth
    ? `<a:ln${lnWidth}><a:solidFill><a:srgbClr val="${hex}"/></a:solidFill></a:ln>`
    : "";
  return `<c:spPr${aNs}>${fill}${ln}</c:spPr>`;
}
function dataLabelsXml(dl) {
  if (!dl) return "";
  return `<c:dLbls>${[
    dl.numFmt ? `<c:numFmt formatCode="${escapeXml(dl.numFmt)}" sourceLinked="0"/>` : "",
    dl.position ? `<c:dLblPos val="${dl.position}"/>` : "",
    `<c:showLegendKey val="0"/>`,
    `<c:showVal val="${dl.showValue ? "1" : "0"}"/>`,
    `<c:showCatName val="${dl.showCategory ? "1" : "0"}"/>`,
    `<c:showSerName val="${dl.showSeriesName ? "1" : "0"}"/>`,
    `<c:showPercent val="${dl.showPercent ? "1" : "0"}"/>`,
  ].join("")}</c:dLbls>`;
}
function seriesXml(type, series, idx, palette) {
  return series
    .map((s, i) => {
      const color = s.color
        ? s.color.startsWith("#")
          ? "FF" + s.color.slice(1)
          : s.color
        : palette
          ? "FF" + palette[i % palette.length]
          : COLORS[i % COLORS.length];
      const catXml = s.categories
        ? `<c:cat><c:strRef><c:f>${escapeXml(s.categories)}</c:f></c:strRef></c:cat>`
        : "";
      const valXml = `<c:val><c:numRef><c:f>${escapeXml(s.values)}</c:f></c:numRef></c:val>`;
      const nameXml = s.name
        ? `<c:tx><c:strRef><c:f>"${escapeXml(s.name)}"</c:f></c:strRef></c:tx>`
        : "";
      const marker =
        type.startsWith("line") || type === "scatter"
          ? `<c:marker><c:symbol val="none"/></c:marker>`
          : "";
      const dlXml = dataLabelsXml(s.dataLabels);
      return `<c:ser><c:idx val="${i}"/><c:order val="${i}"/>${nameXml}${spPr(color, s)}${marker}${dlXml}${catXml}${valXml}</c:ser>`;
    })
    .join("");
}
function axisXml(id, crossId, axis, delete_ = false) {
  if (delete_)
    return `<c:valAx><c:axId val="${id}"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="1"/><c:axPos val="b"/><c:crossAx val="${crossId}"/></c:valAx>`;
  const title = axis?.title
    ? `<c:title><c:tx><c:rich><a:bodyPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/><a:p xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:r><a:t>${escapeXml(axis.title)}</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title>`
    : "";
  const minMax = [
    axis?.min !== void 0 ? `<c:min val="${axis.min}"/>` : "",
    axis?.max !== void 0 ? `<c:max val="${axis.max}"/>` : "",
  ].join("");
  const numFmt = axis?.numFmt
    ? `<c:numFmt formatCode="${escapeXml(axis.numFmt)}" sourceLinked="0"/>`
    : "";
  return `<c:valAx>
  <c:axId val="${id}"/>
  <c:scaling><c:orientation val="minMax"/>${minMax}</c:scaling>
  <c:delete val="0"/>
  <c:axPos val="l"/>
  ${axis?.gridLines !== false ? `<c:majorGridlines/>` : ""}${title}${numFmt}
  <c:crossAx val="${crossId}"/>
</c:valAx>`;
}
function catAxisXml(id, crossId, axis) {
  const title = axis?.title
    ? `<c:title><c:tx><c:rich><a:bodyPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/><a:p xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:r><a:t>${escapeXml(axis.title)}</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title>`
    : "";
  return `<c:catAx>
  <c:axId val="${id}"/>
  <c:scaling><c:orientation val="minMax"/></c:scaling>
  <c:delete val="0"/>
  <c:axPos val="b"/>
  ${axis?.gridLines ? `<c:majorGridlines/>` : ""}${title}
  <c:crossAx val="${crossId}"/>
</c:catAx>`;
}
function legendXml(legend) {
  if (legend === false) return "";
  const raw = typeof legend === "string" ? legend : "b";
  return `<c:legend><c:legendPos val="${
    {
      bottom: "b",
      top: "t",
      left: "l",
      right: "r",
      corner: "tr",
    }[raw] ?? raw
  }"/></c:legend>`;
}
function buildChartXml(chart) {
  const type = chart.type;
  const series = chart.series;
  const palette = chart.colorPalette ? PALETTES[chart.colorPalette] : void 0;
  const title = chart.title
    ? `<c:title><c:tx><c:rich>
    <a:bodyPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/>
    <a:p xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
      <a:r><a:t>${escapeXml(chart.title)}</a:t></a:r>
    </a:p>
  </c:rich></c:tx><c:overlay val="0"/></c:title>`
    : "";
  const varyColors = chart.varyColors ? `<c:varyColors val="1"/>` : "";
  const globalDL = dataLabelsXml(chart.dataLabels);
  let plotXml = "";
  const isPie = type === "pie" || type === "doughnut";
  const isBar = type.startsWith("bar") || type.startsWith("column");
  const isLine = type.startsWith("line");
  const isScatter = type.startsWith("scatter") || type === "bubble";
  const isArea = type.startsWith("area");
  const isRadar = type.startsWith("radar");
  const grouping =
    chart.grouping ??
    (type.endsWith("Stacked100")
      ? "percentStacked"
      : type.endsWith("Stacked")
        ? "stacked"
        : isBar
          ? "clustered"
          : isLine || isArea
            ? "standard"
            : void 0);
  if (isPie) {
    const tag = type === "doughnut" ? "doughnutChart" : "pieChart";
    const hole = type === "doughnut" ? `<c:holeSize val="50"/>` : "";
    plotXml = `<c:${tag}>${varyColors}${seriesXml(type, series, 0, palette)}${globalDL}${hole}</c:${tag}>`;
  } else if (isBar)
    plotXml = `<c:barChart>
  <c:barDir val="${type.startsWith("bar") ? "bar" : "col"}"/>
  ${grouping ? `<c:grouping val="${grouping}"/>` : ""}
  ${varyColors}
  ${seriesXml(type, series, 0, palette)}
  ${globalDL}
  <c:axId val="1"/><c:axId val="2"/>
</c:barChart>`;
  else if (isLine)
    plotXml = `<c:lineChart>
  ${grouping ? `<c:grouping val="${grouping}"/>` : ""}
  ${varyColors}
  ${seriesXml(type, series, 0, palette)}
  ${globalDL}
  <c:axId val="1"/><c:axId val="2"/>
</c:lineChart>`;
  else if (isScatter)
    plotXml = `<c:scatterChart>
  <c:scatterStyle val="${type === "scatterSmooth" ? "smoothMarker" : "marker"}"/>
  ${varyColors}
  ${seriesXml(type, series, 0, palette)}
  ${globalDL}
  <c:axId val="1"/><c:axId val="2"/>
</c:scatterChart>`;
  else if (isArea)
    plotXml = `<c:areaChart>
  ${grouping ? `<c:grouping val="${grouping}"/>` : ""}
  ${varyColors}
  ${seriesXml(type, series, 0, palette)}
  ${globalDL}
  <c:axId val="1"/><c:axId val="2"/>
</c:areaChart>`;
  else if (isRadar)
    plotXml = `<c:radarChart>
  <c:radarStyle val="${type === "radarFilled" ? "filled" : "marker"}"/>
  ${varyColors}
  ${seriesXml(type, series, 0, palette)}
  ${globalDL}
  <c:axId val="1"/><c:axId val="2"/>
</c:radarChart>`;
  else
    plotXml = `<c:barChart>
  <c:barDir val="col"/>
  <c:grouping val="clustered"/>
  ${varyColors}
  ${seriesXml(type, series, 0, palette)}
  ${globalDL}
  <c:axId val="1"/><c:axId val="2"/>
</c:barChart>`;
  const axesXml = !isPie ? catAxisXml(1, 2, chart.xAxis) + axisXml(2, 1, chart.yAxis) : "";
  const legXml = legendXml(chart.legend ?? "b");
  const styleVal = chart.modernStyle
    ? Math.min(MODERN_STYLE_MAP[chart.modernStyle], 48)
    : chart.style;
  const styleXml = styleVal ? `<c:style val="${styleVal}"/>` : "";
  let chartAreaSpPr = "";
  if (chart.chartFill || chart.shadow)
    chartAreaSpPr = `<c:spPr>${chart.chartFill === "gradient" ? `<a:gradFill><a:gsLst><a:gs pos="0"><a:schemeClr val="lt1"/></a:gs><a:gs pos="100000"><a:schemeClr val="bg1"><a:lumMod val="85000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill>` : chart.chartFill === "white" ? `<a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>` : chart.chartFill === "none" ? `<a:noFill/>` : ""}${chart.shadow ? `<a:effectLst><a:outerShdw blurRad="50800" dist="38100" dir="5400000" algn="t" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="40000"/></a:srgbClr></a:outerShdw></a:effectLst>` : ""}</c:spPr>`;
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
${chart.roundedCorners ? `<c:roundedCorners val="1"/>` : ""}
${styleXml}
<c:chart>
  ${title}
  <c:autoTitleDeleted val="${chart.title ? "0" : "1"}"/>
  <c:plotArea>
    ${plotXml}
    ${axesXml}
  </c:plotArea>
  ${legXml}
  <c:plotVisOnly val="1"/>
</c:chart>
${chartAreaSpPr}
</c:chartSpace>`;
}
//#endregion
//#region node_modules/@node-projects/excelforge/dist/vba/ovba.js
function bitCount(pos) {
  if (pos <= 16) return 4;
  if (pos <= 32) return 5;
  if (pos <= 64) return 6;
  if (pos <= 128) return 7;
  if (pos <= 256) return 8;
  if (pos <= 512) return 9;
  if (pos <= 1024) return 10;
  if (pos <= 2048) return 11;
  return 12;
}
function compressOvba(src) {
  const out = [1];
  let srcPos = 0;
  while (srcPos < src.length) {
    const chunkStart = srcPos;
    const chunkEnd = Math.min(srcPos + 4096, src.length);
    const compressed = compressChunk(src, chunkStart, chunkEnd);
    if (compressed.length < 4096) {
      const header = 45056 | ((compressed.length + 2 - 3) & 4095);
      out.push(header & 255, (header >> 8) & 255);
      for (const b of compressed) out.push(b);
    } else {
      out.push(255, 63);
      for (let i = chunkStart; i < chunkEnd; i++) out.push(src[i]);
      for (let i = chunkEnd - chunkStart; i < 4096; i++) out.push(0);
    }
    srcPos = chunkEnd;
  }
  return new Uint8Array(out);
}
function compressChunk(src, start, end) {
  const out = [];
  let dp = 0;
  while (start + dp < end) {
    let flagByte = 0;
    const flagIdx = out.length;
    out.push(0);
    for (let bit = 0; bit < 8 && start + dp < end; bit++) {
      if (dp === 0) {
        out.push(src[start]);
        dp++;
        continue;
      }
      const bc = bitCount(dp);
      const maxOff = 1 << bc;
      const lenBits = 16 - bc;
      const maxLen = (1 << lenBits) - 1 + 3;
      const searchStart = Math.max(0, dp - maxOff);
      let bestOff = 0,
        bestLen = 0;
      for (let s = searchStart; s < dp; s++) {
        let len = 0;
        while (
          dp + len < end - start &&
          len < maxLen &&
          src[start + s + len] === src[start + dp + len]
        )
          len++;
        if (len > bestLen) {
          bestLen = len;
          bestOff = dp - s;
        }
      }
      if (bestLen >= 3) {
        flagByte |= 1 << bit;
        const token = ((bestOff - 1) << lenBits) | (bestLen - 3);
        out.push(token & 255, (token >> 8) & 255);
        dp += bestLen;
      } else {
        out.push(src[start + dp]);
        dp++;
      }
    }
    out[flagIdx] = flagByte;
  }
  return out;
}
function decompressOvba(data) {
  if (data.length === 0 || data[0] !== 1)
    throw new Error("Invalid OVBA compressed container (bad signature)");
  const out = [];
  let pos = 1;
  while (pos < data.length) {
    if (pos + 1 >= data.length) break;
    const header = data[pos] | (data[pos + 1] << 8);
    pos += 2;
    const chunkSize = (header & 4095) + 3;
    const isCompressed = (header >> 15) & 1;
    const chunkDataSize = chunkSize - 2;
    const chunkEnd = pos + chunkDataSize;
    if (!isCompressed) {
      for (let i = 0; i < 4096 && pos < data.length; i++) out.push(data[pos++]);
      pos = Math.max(pos, chunkEnd);
    } else {
      const chunkDecompStart = out.length;
      while (pos < chunkEnd && pos < data.length) {
        const flagByte = data[pos++];
        for (let bit = 0; bit < 8 && pos < chunkEnd && pos < data.length; bit++)
          if ((flagByte & (1 << bit)) === 0) out.push(data[pos++]);
          else {
            if (pos + 1 >= data.length) break;
            const token = data[pos] | (data[pos + 1] << 8);
            pos += 2;
            const lb = 16 - bitCount(out.length - chunkDecompStart);
            const off = (token >> lb) + 1;
            const len = (token & ((1 << lb) - 1)) + 3;
            const copyStart = out.length - off;
            for (let i = 0; i < len; i++) out.push(out[copyStart + i]);
          }
      }
    }
  }
  return new Uint8Array(out);
}
//#endregion
//#region node_modules/@node-projects/excelforge/dist/vba/cfb.js
var SIGNATURE = new Uint8Array([208, 207, 17, 224, 161, 177, 26, 225]);
var ENDOFCHAIN = 4294967294;
var FREESECT = 4294967295;
var FATSECT = 4294967293;
var NOSTREAM = 4294967295;
var SECTOR_SIZE = 512;
var MINI_SIZE = 64;
var MINI_CUTOFF = 4096;
var DIR_ENTRY_SZ = 128;
function u16(buf, off) {
  return buf[off] | (buf[off + 1] << 8);
}
function u32(buf, off) {
  return (buf[off] | (buf[off + 1] << 8) | (buf[off + 2] << 16) | (buf[off + 3] << 24)) >>> 0;
}
function setU16(buf, off, v) {
  buf[off] = v & 255;
  buf[off + 1] = (v >> 8) & 255;
}
function setU32(buf, off, v) {
  buf[off] = v & 255;
  buf[off + 1] = (v >> 8) & 255;
  buf[off + 2] = (v >> 16) & 255;
  buf[off + 3] = (v >> 24) & 255;
}
function encodeName(name) {
  if (name.length > 31)
    throw new Error(`CFB entry name too long (${name.length} chars, max 31): "${name}"`);
  const bytes = /* @__PURE__ */ new Uint8Array(64);
  let i = 0;
  for (const ch of name) {
    const code = ch.charCodeAt(0);
    bytes[i++] = code & 255;
    bytes[i++] = (code >> 8) & 255;
  }
  bytes[i++] = 0;
  bytes[i++] = 0;
  return {
    bytes,
    size: i,
  };
}
function decodeName(buf, off, nameSize) {
  let s = "";
  const len = Math.max(0, nameSize - 2);
  for (let i = 0; i < len; i += 2) {
    const code = buf[off + i] | (buf[off + i + 1] << 8);
    if (code === 0) break;
    s += String.fromCharCode(code);
  }
  return s;
}
function followChain(fat, start) {
  const chain = [];
  let sid = start;
  while (sid !== ENDOFCHAIN && sid < fat.length) {
    chain.push(sid);
    sid = fat[sid];
  }
  return chain;
}
function ceilDiv(a, b) {
  return Math.ceil(a / b);
}
function readCfb(data) {
  for (let i = 0; i < 8; i++) if (data[i] !== SIGNATURE[i]) throw new Error("Not a CFB file");
  const sectorSize = 1 << u16(data, 30);
  const miniSize = 1 << u16(data, 32);
  const fatSectors = u32(data, 44);
  const dirStart = u32(data, 48);
  const miniFatStart = u32(data, 60);
  const sectorOff = (sid) => 512 + sid * sectorSize;
  const difat = [];
  for (let i = 0; i < 109 && i < fatSectors; i++) {
    const sid = u32(data, 76 + i * 4);
    if (sid !== FREESECT) difat.push(sid);
  }
  const totalSectors = Math.floor((data.length - 512) / sectorSize);
  const fat = new Uint32Array(totalSectors);
  fat.fill(FREESECT);
  for (let fi = 0; fi < difat.length; fi++) {
    const off = sectorOff(difat[fi]);
    const entriesPerSector = sectorSize / 4;
    for (let j = 0; j < entriesPerSector && fi * entriesPerSector + j < totalSectors; j++)
      fat[fi * entriesPerSector + j] = u32(data, off + j * 4);
  }
  const dirChain = followChain(fat, dirStart);
  const dirBuf = new Uint8Array(dirChain.length * sectorSize);
  for (let i = 0; i < dirChain.length; i++)
    dirBuf.set(
      data.subarray(sectorOff(dirChain[i]), sectorOff(dirChain[i]) + sectorSize),
      i * sectorSize,
    );
  const numEntries = dirBuf.length / DIR_ENTRY_SZ;
  const entries = [];
  for (let i = 0; i < numEntries; i++) {
    const off = i * DIR_ENTRY_SZ;
    const nameSize = u16(dirBuf, off + 64);
    if (nameSize === 0) continue;
    entries.push({
      name: decodeName(dirBuf, off, nameSize),
      type: dirBuf[off + 66],
      child: u32(dirBuf, off + 76),
      left: u32(dirBuf, off + 68),
      right: u32(dirBuf, off + 72),
      startSector: u32(dirBuf, off + 116),
      size: u32(dirBuf, off + 120),
      index: i,
    });
  }
  const root = entries.find((e) => e.type === 5);
  if (!root) throw new Error("No root entry in CFB");
  const miniContainerChain = followChain(fat, root.startSector);
  const miniContainer = new Uint8Array(miniContainerChain.length * sectorSize);
  for (let i = 0; i < miniContainerChain.length; i++)
    miniContainer.set(
      data.subarray(
        sectorOff(miniContainerChain[i]),
        sectorOff(miniContainerChain[i]) + sectorSize,
      ),
      i * sectorSize,
    );
  const miniFatChain = followChain(fat, miniFatStart);
  const numMiniEntries = miniFatChain.length * (sectorSize / 4);
  const miniFat = new Uint32Array(numMiniEntries);
  miniFat.fill(FREESECT);
  for (let i = 0; i < miniFatChain.length; i++) {
    const off = sectorOff(miniFatChain[i]);
    const count = sectorSize / 4;
    for (let j = 0; j < count; j++) miniFat[i * count + j] = u32(data, off + j * 4);
  }
  const parentMap = /* @__PURE__ */ new Map();
  function collectChildren(parentIdx, parentName) {
    const parent = entries.find((e) => e.index === parentIdx);
    if (!parent || parent.child === NOSTREAM) return;
    const visit = (idx) => {
      if (idx === NOSTREAM || idx >= numEntries) return;
      const e = entries.find((en) => en.index === idx);
      if (!e) return;
      parentMap.set(idx, parentName);
      if (e.type === 1) collectChildren(idx, e.name);
      visit(e.left);
      visit(e.right);
    };
    visit(parent.child);
  }
  collectChildren(root.index, "");
  const streams = [];
  for (const entry of entries) {
    if (entry.type !== 2) continue;
    let streamData;
    if (entry.size < MINI_CUTOFF) {
      const chain = followChain(miniFat, entry.startSector);
      const buf = new Uint8Array(chain.length * miniSize);
      for (let i = 0; i < chain.length; i++) {
        const off = chain[i] * miniSize;
        buf.set(miniContainer.subarray(off, off + miniSize), i * miniSize);
      }
      streamData = buf.subarray(0, entry.size);
    } else {
      const chain = followChain(fat, entry.startSector);
      const buf = new Uint8Array(chain.length * sectorSize);
      for (let i = 0; i < chain.length; i++)
        buf.set(
          data.subarray(sectorOff(chain[i]), sectorOff(chain[i]) + sectorSize),
          i * sectorSize,
        );
      streamData = buf.subarray(0, entry.size);
    }
    streams.push({
      name: entry.name,
      data: streamData,
      storage: parentMap.get(entry.index) || void 0,
    });
  }
  return streams;
}
function buildCfb(streams) {
  const root = {
    name: "Root Entry",
    type: 5,
    children: [],
  };
  const storages = /* @__PURE__ */ new Map();
  storages.set("", root);
  for (const s of streams) {
    const parentName = s.storage ?? "";
    let parent = storages.get(parentName);
    if (!parent) {
      parent = {
        name: parentName,
        type: 1,
        children: [],
      };
      storages.set(parentName, parent);
      root.children.push(parent);
    }
    const node = {
      name: s.name,
      type: 2,
      data: s.data,
    };
    parent.children.push(node);
  }
  const flat = [];
  const queue = [root];
  while (queue.length) {
    const n = queue.shift();
    flat.push({
      node: n,
      parentIdx: -1,
      childIdx: NOSTREAM,
      leftIdx: NOSTREAM,
      rightIdx: NOSTREAM,
      miniStart: 0,
      size: n.data?.length ?? 0,
    });
    if (n.children) for (const c of n.children) queue.push(c);
  }
  for (let i = 0; i < flat.length; i++) {
    const n = flat[i].node;
    if (!n.children?.length) continue;
    const childIdxs = flat
      .map((e, idx) => ({
        e,
        idx,
      }))
      .filter(({ e }) => n.children.includes(e.node))
      .map(({ idx }) => idx);
    if (childIdxs.length === 0) continue;
    const sorted = childIdxs.sort((a, b) => {
      const na = flat[a].node.name,
        nb = flat[b].node.name;
      if (na.length !== nb.length) return na.length - nb.length;
      return na.toUpperCase() < nb.toUpperCase() ? -1 : na.toUpperCase() > nb.toUpperCase() ? 1 : 0;
    });
    const buildTree = (arr) => {
      if (arr.length === 0) return NOSTREAM;
      const mid = arr.length >> 1;
      flat[arr[mid]].leftIdx = buildTree(arr.slice(0, mid));
      flat[arr[mid]].rightIdx = buildTree(arr.slice(mid + 1));
      return arr[mid];
    };
    flat[i].childIdx = buildTree(sorted);
  }
  const miniSectorData = [];
  const miniChains = /* @__PURE__ */ new Map();
  let nextMiniSector = 0;
  for (let i = 0; i < flat.length; i++) {
    const d = flat[i].node.data;
    if (!d || d.length === 0) continue;
    const numMiniSectors = ceilDiv(d.length, MINI_SIZE);
    const chain = [];
    for (let m = 0; m < numMiniSectors; m++) {
      chain.push(nextMiniSector++);
      const chunk = d.subarray(m * MINI_SIZE, Math.min((m + 1) * MINI_SIZE, d.length));
      for (const b of chunk) miniSectorData.push(b);
      for (let p = chunk.length; p < MINI_SIZE; p++) miniSectorData.push(0);
    }
    flat[i].miniStart = chain[0];
    miniChains.set(i, chain);
  }
  const miniFatSectors = ceilDiv(Math.max(nextMiniSector, 1) * 4, SECTOR_SIZE);
  const miniFat = new Uint32Array(miniFatSectors * (SECTOR_SIZE / 4));
  miniFat.fill(FREESECT);
  for (const chain of miniChains.values())
    for (let j = 0; j < chain.length; j++)
      miniFat[chain[j]] = j + 1 < chain.length ? chain[j + 1] : ENDOFCHAIN;
  const containerData = new Uint8Array(miniSectorData);
  const containerSectors = ceilDiv(containerData.length || 1, SECTOR_SIZE);
  const numDirEntries = flat.length;
  const dirSectors = ceilDiv(numDirEntries * DIR_ENTRY_SZ, SECTOR_SIZE);
  const fatSectorIdx = 0;
  const dirStartSector = 1;
  const miniFatStartSector = dirStartSector + dirSectors;
  const containerStartSector = miniFatStartSector + miniFatSectors;
  const totalSectors = 1 + dirSectors + miniFatSectors + containerSectors;
  const fat = new Uint32Array(SECTOR_SIZE / 4);
  fat.fill(FREESECT);
  fat[fatSectorIdx] = FATSECT;
  for (let i = 0; i < dirSectors; i++)
    fat[dirStartSector + i] = i + 1 < dirSectors ? dirStartSector + i + 1 : ENDOFCHAIN;
  for (let i = 0; i < miniFatSectors; i++)
    fat[miniFatStartSector + i] = i + 1 < miniFatSectors ? miniFatStartSector + i + 1 : ENDOFCHAIN;
  for (let i = 0; i < containerSectors; i++)
    fat[containerStartSector + i] =
      i + 1 < containerSectors ? containerStartSector + i + 1 : ENDOFCHAIN;
  flat[0].miniStart = containerStartSector;
  flat[0].size = containerData.length;
  const dirBuf = new Uint8Array(dirSectors * SECTOR_SIZE);
  for (let i = 0; i < flat.length; i++) {
    const off = i * DIR_ENTRY_SZ;
    const e = flat[i];
    const { bytes, size: nameSize } = encodeName(e.node.name);
    dirBuf.set(bytes, off);
    setU16(dirBuf, off + 64, nameSize);
    dirBuf[off + 66] = e.node.type;
    dirBuf[off + 67] = 1;
    setU32(dirBuf, off + 68, e.leftIdx);
    setU32(dirBuf, off + 72, e.rightIdx);
    setU32(dirBuf, off + 76, e.childIdx);
    if (e.node.type === 5) {
      setU32(dirBuf, off + 116, containerStartSector);
      setU32(dirBuf, off + 120, containerData.length);
    } else if (e.node.type === 2 && e.node.data && e.node.data.length > 0) {
      setU32(dirBuf, off + 116, e.miniStart);
      setU32(dirBuf, off + 120, e.node.data.length);
    } else if (e.node.type === 1) {
      setU32(dirBuf, off + 116, 0);
      setU32(dirBuf, off + 120, 0);
    } else {
      setU32(dirBuf, off + 116, ENDOFCHAIN);
      setU32(dirBuf, off + 120, 0);
    }
  }
  const header = /* @__PURE__ */ new Uint8Array(512);
  header.set(SIGNATURE, 0);
  setU16(header, 24, 62);
  setU16(header, 26, 3);
  setU16(header, 28, 65534);
  setU16(header, 30, 9);
  setU16(header, 32, 6);
  setU32(header, 44, 1);
  setU32(header, 48, dirStartSector);
  setU32(header, 56, MINI_CUTOFF);
  setU32(header, 60, miniFatStartSector);
  setU32(header, 64, miniFatSectors);
  setU32(header, 68, ENDOFCHAIN);
  setU32(header, 72, 0);
  setU32(header, 76, fatSectorIdx);
  for (let i = 1; i < 109; i++) setU32(header, 76 + i * 4, FREESECT);
  const fileSize = 512 + totalSectors * SECTOR_SIZE;
  const file = new Uint8Array(fileSize);
  file.set(header, 0);
  const fatBuf = new Uint8Array(SECTOR_SIZE);
  for (let i = 0; i < fat.length; i++) setU32(fatBuf, i * 4, fat[i]);
  file.set(fatBuf, 512);
  file.set(dirBuf, 1024);
  const miniFatBuf = new Uint8Array(miniFatSectors * SECTOR_SIZE);
  for (let i = 0; i < miniFat.length; i++) setU32(miniFatBuf, i * 4, miniFat[i]);
  file.set(miniFatBuf, 512 + miniFatStartSector * SECTOR_SIZE);
  const containerBuf = new Uint8Array(containerSectors * SECTOR_SIZE);
  containerBuf.set(containerData);
  file.set(containerBuf, 512 + containerStartSector * SECTOR_SIZE);
  return file;
}
//#endregion
//#region node_modules/@node-projects/excelforge/dist/vba/VbaProject.js
var VbaProject_exports = /* @__PURE__ */ __exportAll({ VbaProject: () => VbaProject });
var enc = new TextEncoder();
var dec = new TextDecoder("utf-8");
function u16le(v) {
  return [v & 255, (v >> 8) & 255];
}
function u32le(v) {
  return [v & 255, (v >> 8) & 255, (v >> 16) & 255, (v >> 24) & 255];
}
function readU16(buf, off) {
  return buf[off] | (buf[off + 1] << 8);
}
function readU32(buf, off) {
  return (buf[off] | (buf[off + 1] << 8) | (buf[off + 2] << 16) | (buf[off + 3] << 24)) >>> 0;
}
function rec(id, data) {
  const d = data instanceof Uint8Array ? Array.from(data) : data;
  return [...u16le(id), ...u32le(d.length), ...d];
}
function toUtf16le(s) {
  const out = [];
  for (const ch of s) {
    const c = ch.charCodeAt(0);
    out.push(c & 255, (c >> 8) & 255);
  }
  return out;
}
function moduleAttributes(m) {
  const lines = [`Attribute VB_Name = "${m.name}"`];
  if (m.type === "document") {
    const baseClsid =
      m.name === "ThisWorkbook"
        ? "0{00020819-0000-0000-C000-000000000046}"
        : "0{00020820-0000-0000-C000-000000000046}";
    lines.push(
      `Attribute VB_Base = "${baseClsid}"`,
      "Attribute VB_GlobalNameSpace = False",
      "Attribute VB_Creatable = False",
      "Attribute VB_PredeclaredId = True",
      "Attribute VB_Exposed = True",
      "Attribute VB_TemplateDerived = False",
      "Attribute VB_Customizable = True",
    );
  } else if (m.type === "userform")
    lines.push(
      `Attribute VB_Base = "0{F8A47041-B2A6-11CE-8027-00AA00611080}"`,
      "Attribute VB_GlobalNameSpace = False",
      "Attribute VB_Creatable = False",
      "Attribute VB_PredeclaredId = True",
      "Attribute VB_Exposed = False",
      "Attribute VB_TemplateDerived = False",
      "Attribute VB_Customizable = False",
    );
  else if (m.type === "class")
    lines.push(
      "Attribute VB_GlobalNameSpace = False",
      "Attribute VB_Creatable = False",
      "Attribute VB_PredeclaredId = False",
      "Attribute VB_Exposed = False",
    );
  return lines.join("\r\n") + "\r\n";
}
var VbaProject = class VbaProject {
  constructor() {
    this.modules = [];
    this._dirty = false;
  }
  addModule(m) {
    this.modules.push(m);
    this._dirty = true;
    return this;
  }
  removeModule(name) {
    this.modules = this.modules.filter((m) => m.name !== name);
    this._dirty = true;
    return this;
  }
  getModule(name) {
    return this.modules.find((m) => m.name === name);
  }
  build() {
    if (this._raw && !this._dirty) return this._raw;
    const unsorted = this.modules.some((m) => m.name === "ThisWorkbook" && m.type === "document")
      ? this.modules
      : [
          {
            name: "ThisWorkbook",
            type: "document",
            code: "",
          },
          ...this.modules,
        ];
    const allModules = [
      ...unsorted.filter((m) => m.type === "document"),
      ...unsorted.filter((m) => m.type !== "document"),
    ];
    const moduleStreams = [];
    for (const m of allModules) {
      const src = moduleAttributes(m) + m.code;
      const compressed = compressOvba(enc.encode(src));
      moduleStreams.push({
        name: m.name,
        data: compressed,
        storage: "VBA",
      });
    }
    const vbaProjectStream = {
      name: "_VBA_PROJECT",
      data: new Uint8Array([204, 97, 255, 255, 0, 0, 0]),
      storage: "VBA",
    };
    const dirStream = {
      name: "dir",
      data: compressOvba(this._buildDirStream(allModules)),
      storage: "VBA",
    };
    const projectText = this._buildProjectText(allModules);
    const projectStream = {
      name: "PROJECT",
      data: enc.encode(projectText),
    };
    const wmData = this._buildProjectWm(allModules);
    const wmStream = {
      name: "PROJECTwm",
      data: new Uint8Array(wmData),
    };
    const allStreams = [vbaProjectStream, dirStream, ...moduleStreams, projectStream, wmStream];
    for (const m of allModules)
      if (m.type === "userform")
        if (m._formStreams && m._formStreams.size > 0)
          for (const [streamName, data] of m._formStreams)
            allStreams.push({
              name: streamName,
              data,
              storage: m.name,
            });
        else {
          allStreams.push({
            name: "CompObj",
            data: buildCompObj(),
            storage: m.name,
          });
          allStreams.push({
            name: "VBFrame",
            data: enc.encode(buildVBFrameText(m)),
            storage: m.name,
          });
          allStreams.push({
            name: "f",
            data: buildFormControlBinary(m),
            storage: m.name,
          });
          allStreams.push({
            name: "o",
            data: buildOleObjectBlob(m),
            storage: m.name,
          });
        }
    return buildCfb(allStreams);
  }
  _buildDirStream(modules) {
    const b = [];
    const nameBytes = enc.encode("VBAProject");
    b.push(...rec(1, u32le(1)));
    b.push(...rec(2, u32le(1033)));
    b.push(...rec(20, u32le(1033)));
    b.push(...rec(3, u16le(1252)));
    b.push(...rec(4, Array.from(nameBytes)));
    b.push(...rec(5, []));
    b.push(...rec(64, []));
    b.push(...rec(6, []));
    b.push(...rec(61, []));
    b.push(...rec(7, u32le(0)));
    b.push(...rec(8, u32le(0)));
    b.push(...u16le(9), ...u32le(4), ...u32le(1443), ...u16le(2));
    b.push(...rec(12, []));
    b.push(...rec(60, []));
    b.push(...rec(15, u16le(modules.length)));
    b.push(...rec(19, u16le(65535)));
    for (const m of modules) {
      const nameB = enc.encode(m.name);
      const nameU = toUtf16le(m.name);
      b.push(...rec(25, Array.from(nameB)));
      b.push(...rec(71, nameU));
      b.push(...rec(26, Array.from(nameB)));
      b.push(...rec(50, nameU));
      b.push(...rec(28, []));
      b.push(...rec(72, []));
      b.push(...rec(49, u32le(0)));
      b.push(...rec(30, u32le(0)));
      b.push(...rec(44, u16le(65535)));
      if (m.type === "document" || m.type === "userform") b.push(...rec(34, []));
      else b.push(...rec(33, []));
      b.push(...rec(43, []));
    }
    b.push(...rec(16, []));
    return new Uint8Array(b);
  }
  _buildProjectText(modules) {
    const lines = ['ID="{00000000-0000-0000-0000-000000000000}"'];
    const docs = modules.filter((m) => m.type === "document");
    const others = modules.filter((m) => m.type !== "document");
    for (const m of docs) lines.push(`Document=${m.name}/&H00000000`);
    for (const m of others)
      if (m.type === "class") lines.push(`Class=${m.name}`);
      else if (m.type === "userform") {
        lines.push(`Package={AC9F2F90-E877-11CE-9F68-00AA00574A4F}`);
        lines.push(`BaseClass=${m.name}`);
      } else lines.push(`Module=${m.name}`);
    lines.push(
      'Name="VBAProject"',
      "HelpContextID=0",
      'VersionCompatible32="393222000"',
      "",
      "[Host Extender Info]",
      "&H00000001={3832D640-CF90-11CF-8E43-00A0C911005A};VBE;&H00000000",
      "",
      "[Workspace]",
    );
    for (const m of [...docs, ...others]) lines.push(`${m.name}=0, 0, 0, 0, C `);
    return lines.join("\r\n") + "\r\n";
  }
  _buildProjectWm(modules) {
    const out = [];
    for (const m of modules) {
      for (const ch of m.name) out.push(ch.charCodeAt(0));
      out.push(0);
      for (const ch of m.name) {
        const c = ch.charCodeAt(0);
        out.push(c & 255, (c >> 8) & 255);
      }
      out.push(0, 0);
    }
    out.push(0);
    return out;
  }
  static fromBytes(data) {
    const vba = new VbaProject();
    vba._raw = data;
    const streams = readCfb(data);
    const streamMap = /* @__PURE__ */ new Map();
    for (const s of streams) {
      const key = s.storage ? `${s.storage}/${s.name}` : s.name;
      streamMap.set(key, s.data);
    }
    const dirCompressed = streamMap.get("VBA/dir");
    if (!dirCompressed) return vba;
    let dirData;
    try {
      dirData = decompressOvba(dirCompressed);
    } catch {
      return vba;
    }
    const classNames = /* @__PURE__ */ new Set();
    const userFormNames = /* @__PURE__ */ new Set();
    const projectData = streamMap.get("PROJECT");
    if (projectData) {
      const projectText = dec.decode(projectData);
      for (const line of projectText.split(/\r?\n/)) {
        const m = line.match(/^Class=(.+)$/);
        if (m) classNames.add(m[1]);
        const uf = line.match(/^BaseClass=(.+)$/);
        if (uf) userFormNames.add(uf[1]);
      }
    }
    const moduleInfos = parseDirStream(dirData);
    for (const info of moduleInfos) {
      const streamKey = `VBA/${info.name}`;
      const streamData = streamMap.get(streamKey);
      if (!streamData) continue;
      let code = "";
      try {
        const decompressed = decompressOvba(streamData.subarray(info.offset));
        code = dec.decode(decompressed);
      } catch {
        code = dec.decode(streamData.subarray(info.offset));
      }
      const stripped = stripAttributes(code);
      const type = userFormNames.has(info.name)
        ? "userform"
        : classNames.has(info.name)
          ? "class"
          : info.type;
      const mod = {
        name: info.name,
        type,
        code: stripped,
      };
      if (type === "userform") {
        const formStreams = /* @__PURE__ */ new Map();
        for (const [key, data] of streamMap)
          if (key.startsWith(info.name + "/")) {
            const streamName = key.substring(info.name.length + 1);
            formStreams.set(streamName, data);
          }
        if (formStreams.size > 0) mod._formStreams = formStreams;
        const designerKey = `${info.name}/f`;
        const designerData = streamMap.get(designerKey);
        if (designerData) mod.designerData = designerData;
      }
      vba.modules.push(mod);
    }
    vba._dirty = false;
    return vba;
  }
};
function parseDirStream(data) {
  const modules = [];
  let pos = 0;
  let currentModule = null;
  while (pos + 6 <= data.length) {
    const id = readU16(data, pos);
    const size = readU32(data, pos + 2);
    pos += 6;
    if (id === 9) {
      pos += 6;
      continue;
    }
    if (pos + size > data.length) break;
    const body = data.subarray(pos, pos + size);
    pos += size;
    switch (id) {
      case 25:
        currentModule = {
          name: dec.decode(body),
          type: "standard",
          offset: 0,
        };
        break;
      case 49:
        if (currentModule && body.length >= 4) currentModule.offset = readU32(body, 0);
        break;
      case 33:
        if (currentModule) currentModule.type = "standard";
        break;
      case 34:
        if (currentModule) currentModule.type = "document";
        break;
      case 43:
        if (currentModule?.name) modules.push(currentModule);
        currentModule = null;
        break;
    }
  }
  return modules;
}
function stripAttributes(code) {
  const lines = code.split(/\r?\n/);
  const start = lines.findIndex((l) => !l.startsWith("Attribute "));
  if (start < 0) return code;
  return lines.slice(start).join("\n");
}
var USERFORM_CLSID_STR = "{C62A69F0-16DC-11CE-9E98-00AA00574A4F}";
var CLSID_CACHE_INDEX = {
  Image: 12,
  Frame: 14,
  SpinButton: 16,
  CommandButton: 17,
  Label: 21,
  TextBox: 23,
  ListBox: 24,
  ComboBox: 25,
  CheckBox: 26,
  OptionButton: 27,
  ToggleButton: 28,
  ScrollBar: 47,
};
function buildCompObj() {
  return new Uint8Array([
    1, 0, 254, 255, 3, 10, 0, 0, 255, 255, 255, 255, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0,
    25, 0, 0, 0, 77, 105, 99, 114, 111, 115, 111, 102, 116, 32, 70, 111, 114, 109, 115, 32, 50, 46,
    48, 32, 70, 111, 114, 109, 0, 16, 0, 0, 0, 69, 109, 98, 101, 100, 100, 101, 100, 32, 79, 98,
    106, 101, 99, 116, 0, 0, 0, 0, 0, 244, 57, 178, 113, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0,
  ]);
}
function buildVBFrameText(m) {
  return (
    [
      "VERSION 5.00",
      `Begin ${USERFORM_CLSID_STR} ${m.name} `,
      `   Caption         =   "${m.name}"`,
      "   ClientHeight    =   3015",
      "   ClientLeft      =   120",
      "   ClientTop       =   465",
      "   ClientWidth     =   4560",
      "   StartUpPosition =   1  'CenterOwner",
      "   TypeInfoVer     =   2",
      "End",
    ].join("\r\n") + "\r\n"
  );
}
function buildFormControlBinary(m) {
  const controls = m.controls ?? [];
  const numControls = controls.length;
  const oBlobs = controls.map((ctrl) => buildSingleControlOBlob(ctrl));
  const b = [];
  b.push(0);
  b.push(4);
  const cbPos = b.length;
  b.push(0, 0);
  b.push(8, 12, 0, 12);
  b.push(...u32le(numControls + 1));
  b.push(...u32le(0));
  b.push(...u32le(32e3));
  let maxR = 200,
    maxB = 120;
  for (const c of controls) {
    const r = (c.left ?? 0) + (c.width ?? 72);
    const bt = (c.top ?? 0) + (c.height ?? 24);
    if (r > maxR) maxR = r;
    if (bt > maxB) maxB = bt;
  }
  const formWidthHimetric = (maxR + 20) * 26;
  const formHeightHimetric = (maxB + 20) * 26;
  b.push(...u32le(formWidthHimetric));
  b.push(...u32le(formHeightHimetric));
  b.push(...u32le(0));
  b.push(...u32le(0));
  const cbForm = b.length - 4;
  b[cbPos] = cbForm & 255;
  b[cbPos + 1] = (cbForm >> 8) & 255;
  b.push(...u16le(0));
  b.push(...u32le(numControls));
  const cbBytesPos = b.length;
  b.push(0, 0, 0, 0);
  const countedStart = b.length;
  for (let i = 0; i < numControls; i++) {
    b.push(0);
    b.push(1);
  }
  while ((b.length - countedStart) % 4 !== 0) b.push(0);
  for (let i = 0; i < numControls; i++) {
    const ctrl = controls[i];
    const nameBytes = enc.encode(ctrl.name);
    const cacheIdx = CLSID_CACHE_INDEX[ctrl.type] ?? CLSID_CACHE_INDEX["CommandButton"];
    b.push(0, 0);
    const cbSitePos = b.length;
    b.push(0, 0);
    b.push(229, 1, 0, 0);
    b.push(...u32le(nameBytes.length | 2147483648));
    b.push(...u32le(i + 1));
    b.push(...u32le(oBlobs[i].length));
    b.push(i & 255, (i >> 8) & 255);
    b.push(cacheIdx & 255, (cacheIdx >> 8) & 255);
    b.push(...nameBytes);
    while ((b.length - countedStart) % 4 !== 0) b.push(0);
    const left = (ctrl.left ?? 10 + i * 80) * 26;
    const top = (ctrl.top ?? 10 + i * 30) * 26;
    b.push(...u32le(left));
    b.push(...u32le(top));
    const cbSite = b.length - cbSitePos - 2;
    b[cbSitePos] = cbSite & 255;
    b[cbSitePos + 1] = (cbSite >> 8) & 255;
  }
  const cbBytes = b.length - countedStart;
  b[cbBytesPos] = cbBytes & 255;
  b[cbBytesPos + 1] = (cbBytes >> 8) & 255;
  b[cbBytesPos + 2] = (cbBytes >> 16) & 255;
  b[cbBytesPos + 3] = (cbBytes >> 24) & 255;
  return new Uint8Array(b);
}
function buildSingleControlOBlob(ctrl) {
  const cacheIdx = CLSID_CACHE_INDEX[ctrl.type] ?? CLSID_CACHE_INDEX["CommandButton"];
  if (cacheIdx === 17) return buildCommandButtonOBlob(ctrl);
  if (cacheIdx === 21) return buildLabelOBlob(ctrl);
  return buildMorphDataOBlob(ctrl);
}
function appendTextProps(b) {
  b.push(0, 2);
  const cbTPPos = b.length;
  b.push(0, 0);
  b.push(5, 0, 0, 0);
  const fontNameBytes = enc.encode("Tahoma");
  b.push(...u32le(fontNameBytes.length | 2147483648));
  b.push(160, 0, 0, 0);
  b.push(...fontNameBytes);
  while (b.length % 4 !== 0) b.push(0);
  const cbTP = b.length - cbTPPos - 2;
  b[cbTPPos] = cbTP & 255;
  b[cbTPPos + 1] = (cbTP >> 8) & 255;
}
function buildCommandButtonOBlob(ctrl) {
  const b = [];
  const hasCaption = !!ctrl.caption;
  b.push(0, 2);
  const cbPos = b.length;
  b.push(0, 0);
  let mask = 32;
  if (hasCaption) mask |= 8;
  b.push(mask & 255, (mask >> 8) & 255, (mask >> 16) & 255, (mask >> 24) & 255);
  if (hasCaption) {
    const capBytes = enc.encode(ctrl.caption);
    b.push(...u32le(capBytes.length | 2147483648));
  }
  if (hasCaption) {
    const capBytes = enc.encode(ctrl.caption);
    b.push(...capBytes);
    while (b.length % 4 !== 0) b.push(0);
  }
  const width = (ctrl.width ?? 72) * 26;
  const height = (ctrl.height ?? 24) * 26;
  b.push(...u32le(width));
  b.push(...u32le(height));
  const cb = b.length - cbPos - 2;
  b[cbPos] = cb & 255;
  b[cbPos + 1] = (cb >> 8) & 255;
  appendTextProps(b);
  return new Uint8Array(b);
}
function buildLabelOBlob(ctrl) {
  const b = [];
  const hasCaption = !!ctrl.caption;
  b.push(0, 2);
  const cbPos = b.length;
  b.push(0, 0);
  let mask = 32;
  if (hasCaption) mask |= 8;
  b.push(mask & 255, (mask >> 8) & 255, (mask >> 16) & 255, (mask >> 24) & 255);
  if (hasCaption) {
    const capBytes = enc.encode(ctrl.caption);
    b.push(...u32le(capBytes.length | 2147483648));
  }
  if (hasCaption) {
    const capBytes = enc.encode(ctrl.caption);
    b.push(...capBytes);
    while (b.length % 4 !== 0) b.push(0);
  }
  const width = (ctrl.width ?? 80) * 26;
  const height = (ctrl.height ?? 18) * 26;
  b.push(...u32le(width));
  b.push(...u32le(height));
  const cb = b.length - cbPos - 2;
  b[cbPos] = cb & 255;
  b[cbPos + 1] = (cb >> 8) & 255;
  appendTextProps(b);
  return new Uint8Array(b);
}
function buildMorphDataOBlob(ctrl) {
  const b = [];
  const hasCaption = !!ctrl.caption;
  b.push(0, 2);
  const cbPos = b.length;
  b.push(0, 0);
  let part1 = 257;
  if (hasCaption) part1 |= 1 << 23;
  part1 |= 1 << 31;
  b.push(part1 & 255, (part1 >> 8) & 255, (part1 >> 16) & 255, (part1 >>> 24) & 255);
  b.push(0, 0, 0, 0);
  b.push(27, 72, 128, 44);
  if (hasCaption) {
    const capBytes = enc.encode(ctrl.caption);
    b.push(...u32le(capBytes.length | 2147483648));
  }
  const width = (ctrl.width ?? 72) * 26;
  const height = (ctrl.height ?? 24) * 26;
  b.push(...u32le(width));
  b.push(...u32le(height));
  if (hasCaption) {
    const capBytes = enc.encode(ctrl.caption);
    b.push(...capBytes);
    while (b.length % 4 !== 0) b.push(0);
  }
  const cb = b.length - cbPos - 2;
  b[cbPos] = cb & 255;
  b[cbPos + 1] = (cb >> 8) & 255;
  appendTextProps(b);
  return new Uint8Array(b);
}
function buildOleObjectBlob(m) {
  const controls = m.controls ?? [];
  if (controls.length === 0) return /* @__PURE__ */ new Uint8Array(0);
  const parts = [];
  for (const ctrl of controls) parts.push(buildSingleControlOBlob(ctrl));
  const total = parts.reduce((s, p) => s + p.length, 0);
  const result = new Uint8Array(total);
  let off = 0;
  for (const p of parts) {
    result.set(p, off);
    off += p.length;
  }
  return result;
}
//#endregion
//#region node_modules/@node-projects/excelforge/dist/features/FormControlBuilder.js
var FormControlBuilder_exports = /* @__PURE__ */ __exportAll({
  buildCtrlPropXml: () => buildCtrlPropXml,
  buildFormControlVmlShape: () => buildFormControlVmlShape,
  buildVmlWithControls: () => buildVmlWithControls,
});
function buildCtrlPropXml(ctrl) {
  if (ctrl._ctrlPropXml) return ctrl._ctrlPropXml;
  const attrs = [`objectType="${CTRL_OBJ_TYPE[ctrl.type] ?? "Button"}"`, 'lockText="1"'];
  if (ctrl.linkedCell) attrs.push(`fmlaLink="${escapeXml(ctrl.linkedCell)}"`);
  if (ctrl.inputRange) attrs.push(`fmlaRange="${escapeXml(ctrl.inputRange)}"`);
  if (ctrl.checked !== void 0) attrs.push(`checked="${CHECKED_MAP[ctrl.checked] ?? "0"}"`);
  if (ctrl.dropLines !== void 0) attrs.push(`dropLines="${ctrl.dropLines}"`);
  if (ctrl.min !== void 0) attrs.push(`min="${ctrl.min}"`);
  if (ctrl.max !== void 0) attrs.push(`max="${ctrl.max}"`);
  if (ctrl.inc !== void 0) attrs.push(`inc="${ctrl.inc}"`);
  if (ctrl.page !== void 0) attrs.push(`page="${ctrl.page}"`);
  if (ctrl.val !== void 0) attrs.push(`val="${ctrl.val}"`);
  if (ctrl.selType)
    attrs.push(
      `selType="${
        {
          single: "Single",
          multi: "Multi",
          extend: "Extend",
        }[ctrl.selType] ?? "Single"
      }"`,
    );
  if (ctrl.noThreeD) attrs.push('noThreeD="1"');
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<formControlPr xmlns="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" ${attrs.join(" ")}/>`;
}
function vmlAnchor(from, to) {
  return `${from.col}, ${from.colOff ?? 0}, ${from.row}, ${from.rowOff ?? 0}, ${to.col}, ${to.colOff ?? 0}, ${to.row}, ${to.rowOff ?? 0}`;
}
function buildFormControlVmlShape(ctrl, shapeId) {
  if (ctrl._vmlShapeXml) return ctrl._vmlShapeXml;
  const objType = VML_OBJ_TYPE[ctrl.type] ?? "Button";
  const anchor = vmlAnchor(ctrl.from, ctrl.to);
  const cd = [];
  cd.push(`<x:Anchor>${anchor}</x:Anchor>`);
  cd.push("<x:PrintObject>False</x:PrintObject>");
  cd.push("<x:AutoFill>False</x:AutoFill>");
  if (ctrl.macro) cd.push(`<x:FmlaMacro>${escapeXml(ctrl.macro)}</x:FmlaMacro>`);
  if (ctrl.linkedCell) cd.push(`<x:FmlaLink>${escapeXml(ctrl.linkedCell)}</x:FmlaLink>`);
  if (ctrl.inputRange) cd.push(`<x:FmlaRange>${escapeXml(ctrl.inputRange)}</x:FmlaRange>`);
  if (ctrl.checked !== void 0)
    cd.push(`<x:Checked>${CHECKED_MAP[ctrl.checked] ?? "0"}</x:Checked>`);
  if (ctrl.dropLines !== void 0) cd.push(`<x:DropLines>${ctrl.dropLines}</x:DropLines>`);
  if (ctrl.dropStyle) cd.push(`<x:DropStyle>${escapeXml(ctrl.dropStyle)}</x:DropStyle>`);
  if (ctrl.min !== void 0) cd.push(`<x:Min>${ctrl.min}</x:Min>`);
  if (ctrl.max !== void 0) cd.push(`<x:Max>${ctrl.max}</x:Max>`);
  if (ctrl.inc !== void 0) cd.push(`<x:Inc>${ctrl.inc}</x:Inc>`);
  if (ctrl.page !== void 0) cd.push(`<x:Page>${ctrl.page}</x:Page>`);
  if (ctrl.val !== void 0) cd.push(`<x:Val>${ctrl.val}</x:Val>`);
  if (ctrl.selType)
    cd.push(
      `<x:Sel>${
        {
          single: "Single",
          multi: "Multi",
          extend: "Extend",
        }[ctrl.selType] ?? "Single"
      }</x:Sel>`,
    );
  if (ctrl.noThreeD) cd.push("<x:NoThreeD/>");
  if (ctrl.isDefault) cd.push("<x:Default/>");
  if (ctrl.isDismiss) cd.push("<x:Dismiss/>");
  if (ctrl.isCancel) cd.push("<x:Cancel/>");
  let textBox = "";
  if (
    ctrl.text &&
    (ctrl.type === "button" ||
      ctrl.type === "checkBox" ||
      ctrl.type === "optionButton" ||
      ctrl.type === "groupBox" ||
      ctrl.type === "label" ||
      ctrl.type === "dialog")
  ) {
    textBox = `<v:textbox style="mso-direction-alt:auto"><div style="text-align:${ctrl.type === "button" ? "center" : "left"}"><font face="Calibri" size="220" color="#000000">${escapeXml(ctrl.text)}</font></div></v:textbox>`;
    cd.push("<x:TextHAlign>Center</x:TextHAlign>");
    cd.push("<x:TextVAlign>Center</x:TextVAlign>");
  }
  const left = (ctrl.from.col * 64 + (ctrl.from.colOff ?? 0)) * 0.75;
  const top = (ctrl.from.row * 20 + (ctrl.from.rowOff ?? 0)) * 0.75;
  const width = (ctrl.to.col - ctrl.from.col) * 64 * 0.75;
  const height = (ctrl.to.row - ctrl.from.row) * 20 * 0.75;
  const shapeName = ctrl.text ?? `${ctrl.type}_${shapeId}`;
  const fillStroke = ctrl.type === "dialog" ? "" : ' filled="f" stroked="f"';
  return `<v:shape o:spid="_x0000_s${shapeId}" id="${escapeXml(shapeName)}" type="#_x0000_t201" style="position:absolute;margin-left:${left.toFixed(1)}pt;margin-top:${top.toFixed(1)}pt;width:${width.toFixed(1)}pt;height:${height.toFixed(1)}pt;z-index:${shapeId}"${fillStroke} o:insetmode="auto"><o:lock v:ext="edit" rotation="t"/>${textBox}<x:ClientData ObjectType="${objType}">${cd.join("")}</x:ClientData></v:shape>`;
}
function buildVmlWithControls(commentShapes, controlShapes) {
  return `<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel">
<o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="1"/></o:shapelayout>
<v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe"><v:stroke joinstyle="miter"/><v:path gradientshapeok="t" o:connecttype="rect"/></v:shapetype>
<v:shapetype id="_x0000_t201" coordsize="21600,21600" o:spt="201" path="m,l,21600r21600,l21600,xe"><v:stroke joinstyle="miter"/><v:path shadowok="f" o:extrusionok="f" strokeok="f" fillok="f" o:connecttype="rect"/><o:lock v:ext="edit" shapetype="t"/></v:shapetype>
${commentShapes.join("\n")}
${controlShapes.join("\n")}
</xml>`;
}
//#endregion
//#region node_modules/@node-projects/excelforge/dist/features/PivotTableBuilder.js
var PivotTableBuilder_exports = /* @__PURE__ */ __exportAll({
  buildPivotTableFiles: () => buildPivotTableFiles,
});
var FUNC_MAP = {
  sum: "sum",
  count: "count",
  average: "average",
  max: "max",
  min: "min",
  product: "product",
  countNums: "countNums",
  stdDev: "stdDev",
  stdDevp: "stdDevp",
  var: "var",
  varp: "varp",
};
function buildPivotTableFiles(pt, sourceData, pivotId, cacheId) {
  const headers = (sourceData[0] ?? []).map((v) => String(v ?? ""));
  const dataRows = sourceData.slice(1);
  const numFields = headers.length;
  const rowGT = pt.rowGrandTotals !== false;
  const colGT = pt.colGrandTotals !== false;
  const fieldIdx = new Map(headers.map((h, i) => [h, i]));
  const rowFldIdxs = pt.rowFields.map((f) => fieldIdx.get(f) ?? 0);
  const colFldIdxs = pt.colFields.map((f) => fieldIdx.get(f) ?? 0);
  const dataFldIdxs = pt.dataFields.map((df) => fieldIdx.get(df.field) ?? 0);
  const uniqueVals = Array.from({ length: numFields }, () => []);
  const uniqueMap = Array.from({ length: numFields }, () => /* @__PURE__ */ new Map());
  const isNumeric = new Array(numFields).fill(true);
  for (const row of dataRows)
    for (let fi = 0; fi < numFields; fi++) {
      const v = row[fi];
      const vs = v === null || v === void 0 ? "" : String(v);
      if (typeof v !== "number") isNumeric[fi] = false;
      if (!uniqueMap[fi].has(vs)) {
        uniqueMap[fi].set(vs, uniqueVals[fi].length);
        uniqueVals[fi].push(vs);
      }
    }
  const { row: tRow, col: tCol } = cellRefToIndices(pt.targetCell);
  const hasColFields = colFldIdxs.length > 0;
  const numColCombos = hasColFields
    ? colFldIdxs.reduce((n, fi) => n * Math.max(uniqueVals[fi].length, 1), 1)
    : 1;
  const totalRows =
    1 +
    (rowFldIdxs.length
      ? rowFldIdxs.reduce((n, fi) => n * Math.max(uniqueVals[fi].length, 1), 1)
      : 1) +
    (rowGT ? 1 : 0);
  const numDataFields = pt.dataFields.length + (pt.calculatedFields ?? []).length;
  const totalCols = rowFldIdxs.length + numColCombos * numDataFields + (colGT ? numDataFields : 0);
  const locationRef = `${indicesToCellRef(tRow, tCol)}:${indicesToCellRef(tRow + totalRows - 1, tCol + totalCols - 1)}`;
  const firstDataCol = rowFldIdxs.length + 1;
  const isAxisField = /* @__PURE__ */ new Set([...rowFldIdxs, ...colFldIdxs]);
  const calcFields = pt.calculatedFields ?? [];
  const calcFieldsXml = calcFields
    .map((cf) => {
      return `<cacheField name="${escapeXml(cf.name)}" numFmtId="0" formula="${escapeXml(cf.formula)}" databaseField="0"><sharedItems/></cacheField>`;
    })
    .join("");
  const totalFields = numFields + calcFields.length;
  const groupingMap = /* @__PURE__ */ new Map();
  for (const grp of pt.fieldGrouping ?? []) {
    const fi = fieldIdx.get(grp.field);
    if (fi !== void 0) groupingMap.set(fi, grp);
  }
  const cacheFieldsXml =
    headers
      .map((name, fi) => {
        const grp = groupingMap.get(fi);
        if (isAxisField.has(fi)) {
          const items = uniqueVals[fi].map((v) => `<s v="${escapeXml(v)}"/>`).join("");
          let groupXml = "";
          if (grp)
            if (grp.groupBy === "numeric")
              groupXml = `<fieldGroup base="${fi}"><rangePr startNum="${grp.start ?? 0}" endNum="${grp.end ?? 100}" groupInterval="${grp.interval ?? 10}"/></fieldGroup>`;
            else
              groupXml = `<fieldGroup base="${fi}"><rangePr groupBy="${
                {
                  days: "days",
                  months: "months",
                  quarters: "quarters",
                  years: "years",
                }[grp.groupBy] ?? "months"
              }"/></fieldGroup>`;
          return `<cacheField name="${escapeXml(name)}" numFmtId="0"><sharedItems count="${uniqueVals[fi].length}">${items}</sharedItems>${groupXml}</cacheField>`;
        }
        if (isNumeric[fi]) {
          const nums = dataRows.map((r) => Number(r[fi])).filter((n) => !isNaN(n));
          const min = nums.length ? Math.min(...nums) : 0;
          const max = nums.length ? Math.max(...nums) : 0;
          return `<cacheField name="${escapeXml(name)}" numFmtId="0"><sharedItems containsSemiMixedTypes="0" containsString="0" containsNumber="1" minValue="${min}" maxValue="${max}"/></cacheField>`;
        }
        return `<cacheField name="${escapeXml(name)}" numFmtId="0"><sharedItems/></cacheField>`;
      })
      .join("") + calcFieldsXml;
  const recordsXml = dataRows
    .map((row) => {
      return `<r>${headers
        .map((_, fi) => {
          const v = row[fi];
          if (isAxisField.has(fi)) {
            const vs = v === null || v === void 0 ? "" : String(v);
            return `<x v="${uniqueMap[fi].get(vs) ?? 0}"/>`;
          }
          if (typeof v === "number") return `<n v="${v}"/>`;
          if (typeof v === "boolean") return `<b v="${v ? 1 : 0}"/>`;
          return `<s v="${escapeXml(String(v ?? ""))}"/>`;
        })
        .join("")}</r>`;
    })
    .join("");
  const pivotFieldsXml =
    headers
      .map((_, fi) => {
        const isRow = rowFldIdxs.includes(fi);
        const isCol = colFldIdxs.includes(fi);
        const isData = dataFldIdxs.includes(fi);
        if (isRow) {
          const items =
            uniqueVals[fi].map((_, vi) => `<item x="${vi}"/>`).join("") + '<item t="default"/>';
          return `<pivotField axis="axisRow" showAll="0"><items count="${uniqueVals[fi].length + 1}">${items}</items></pivotField>`;
        }
        if (isCol) {
          const items =
            uniqueVals[fi].map((_, vi) => `<item x="${vi}"/>`).join("") + '<item t="default"/>';
          return `<pivotField axis="axisCol" showAll="0"><items count="${uniqueVals[fi].length + 1}">${items}</items></pivotField>`;
        }
        if (isData) return `<pivotField dataField="1" showAll="0"/>`;
        return `<pivotField showAll="0"/>`;
      })
      .join("") + calcFields.map(() => `<pivotField dataField="1" showAll="0"/>`).join("");
  let rowFieldsXml = "";
  let rowItemsXml = "";
  if (rowFldIdxs.length) {
    rowFieldsXml = `<rowFields count="${rowFldIdxs.length}">${rowFldIdxs.map((fi) => `<field x="${fi}"/>`).join("")}</rowFields>`;
    const fi = rowFldIdxs[0];
    const items = uniqueVals[fi].map((_, vi) => `<i><x v="${vi}"/></i>`).join("");
    const grand = rowGT ? '<i t="grand"><x/></i>' : "";
    rowItemsXml = `<rowItems count="${uniqueVals[fi].length + (rowGT ? 1 : 0)}">${items}${grand}</rowItems>`;
  }
  let colFieldsXml = "";
  let colItemsXml = "";
  if (hasColFields) {
    colFieldsXml = `<colFields count="${colFldIdxs.length}">${colFldIdxs.map((fi) => `<field x="${fi}"/>`).join("")}</colFields>`;
    const fi = colFldIdxs[0];
    const items = uniqueVals[fi].map((_, vi) => `<i><x v="${vi}"/></i>`).join("");
    const grand = colGT ? '<i t="grand"><x/></i>' : "";
    colItemsXml = `<colItems count="${uniqueVals[fi].length + (colGT ? 1 : 0)}">${items}${grand}</colItems>`;
  } else if (numDataFields > 1) {
    colFieldsXml = `<colFields count="1"><field x="-2"/></colFields>`;
    colItemsXml = `<colItems count="${numDataFields}">${Array.from({ length: numDataFields }, (_, di) => `<i><x v="${di}"/></i>`).join("")}</colItems>`;
  }
  const allDataFieldEntries = [];
  for (let i = 0; i < pt.dataFields.length; i++) {
    const df = pt.dataFields[i];
    const fi = dataFldIdxs[i];
    const func = FUNC_MAP[df.func ?? "sum"] ?? "sum";
    const name = df.name ?? `Sum of ${df.field}`;
    allDataFieldEntries.push(
      `<dataField name="${escapeXml(name)}" fld="${fi}" subtotal="${func}" showDataAs="normal" baseField="0" baseItem="0"/>`,
    );
  }
  for (let ci = 0; ci < calcFields.length; ci++) {
    const cf = calcFields[ci];
    const fi = numFields + ci;
    allDataFieldEntries.push(
      `<dataField name="${escapeXml(cf.name)}" fld="${fi}" subtotal="sum" showDataAs="normal" baseField="0" baseItem="0"/>`,
    );
  }
  const dataFieldsXml = `<dataFields count="${allDataFieldEntries.length}">${allDataFieldEntries.join("")}</dataFields>`;
  const gtParts = [rowGT ? "" : 'rowGrandTotals="0"', colGT ? "" : 'colGrandTotals="0"'].filter(
    Boolean,
  );
  const gtAttrStr = gtParts.length ? " " + gtParts.join(" ") : "";
  const style = escapeXml(pt.style ?? "PivotStyleMedium9");
  return {
    pivotTableXml: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" name="${escapeXml(pt.name)}" cacheId="${cacheId}" applyNumberFormats="0" applyBorderFormats="0" applyFontFormats="0" applyPatternFormats="0" applyAlignmentFormats="0" applyWidthHeightFormats="1" dataCaption="Values" updatedVersion="6" minRefreshableVersion="3" useAutoFormatting="1" itemPrintTitles="1" createdVersion="6" indent="0" outline="1" outlineData="1" multipleFieldFilters="0"${gtAttrStr}>
<location ref="${locationRef}" firstHeaderRow="1" firstDataRow="2" firstDataCol="${firstDataCol}"/>
<pivotFields count="${totalFields}">${pivotFieldsXml}</pivotFields>
${rowFieldsXml}${rowItemsXml}${colFieldsXml}${colItemsXml}${dataFieldsXml}
<pivotTableStyleInfo name="${style}" showRowHeaders="1" showColHeaders="1" showRowStripes="0" showColStripes="0" showLastColumn="1"/>
</pivotTableDefinition>`,
    cacheDefXml: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId1" refreshedBy="ExcelForge" refreshedDate="45000" createdVersion="6" refreshedVersion="6" minRefreshableVersion="3" recordCount="${dataRows.length}" saveData="1">
<cacheSource type="worksheet"><worksheetSource ref="${pt.sourceRef}" sheet="${escapeXml(pt.sourceSheet)}"/></cacheSource>
<cacheFields count="${totalFields}">${cacheFieldsXml}</cacheFields>
</pivotCacheDefinition>`,
    cacheRecordsXml: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotCacheRecords xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" count="${dataRows.length}">${recordsXml}</pivotCacheRecords>`,
  };
}
//#endregion
export { Workbook as n, VbaProject as t };
