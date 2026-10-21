import { a as e, n as t, o as n } from "./helpers-nemQBT-g.js";
var r = {
  sum: `sum`,
  count: `count`,
  average: `average`,
  max: `max`,
  min: `min`,
  product: `product`,
  countNums: `countNums`,
  stdDev: `stdDev`,
  stdDevp: `stdDevp`,
  var: `var`,
  varp: `varp`,
};
function i(i, a, o, s) {
  let c = (a[0] ?? []).map((e) => String(e ?? ``)),
    l = a.slice(1),
    u = c.length,
    d = i.rowGrandTotals !== !1,
    f = i.colGrandTotals !== !1,
    p = new Map(c.map((e, t) => [e, t])),
    m = i.rowFields.map((e) => p.get(e) ?? 0),
    h = i.colFields.map((e) => p.get(e) ?? 0),
    g = i.dataFields.map((e) => p.get(e.field) ?? 0),
    _ = Array.from({ length: u }, () => []),
    v = Array.from({ length: u }, () => new Map()),
    y = Array(u).fill(!0);
  for (let e of l)
    for (let t = 0; t < u; t++) {
      let n = e[t],
        r = n == null ? `` : String(n);
      (typeof n != `number` && (y[t] = !1),
        v[t].has(r) || (v[t].set(r, _[t].length), _[t].push(r)));
    }
  let { row: b, col: x } = t(i.targetCell),
    S = h.length > 0,
    C = S ? h.reduce((e, t) => e * Math.max(_[t].length, 1), 1) : 1,
    w = 1 + (m.length ? m.reduce((e, t) => e * Math.max(_[t].length, 1), 1) : 1) + +!!d,
    T = i.dataFields.length + (i.calculatedFields ?? []).length,
    E = m.length + C * T + (f ? T : 0),
    D = `${n(b, x)}:${n(b + w - 1, x + E - 1)}`,
    O = m.length + 1,
    k = new Set([...m, ...h]),
    A = i.calculatedFields ?? [],
    j = A.map(
      (t) =>
        `<cacheField name="${e(t.name)}" numFmtId="0" formula="${e(t.formula)}" databaseField="0"><sharedItems/></cacheField>`,
    ).join(``),
    M = u + A.length,
    N = new Map();
  for (let e of i.fieldGrouping ?? []) {
    let t = p.get(e.field);
    t !== void 0 && N.set(t, e);
  }
  let P =
      c
        .map((t, n) => {
          let r = N.get(n);
          if (k.has(n)) {
            let i = _[n].map((t) => `<s v="${e(t)}"/>`).join(``),
              a = ``;
            return (
              r &&
                (a =
                  r.groupBy === `numeric`
                    ? `<fieldGroup base="${n}"><rangePr startNum="${r.start ?? 0}" endNum="${r.end ?? 100}" groupInterval="${r.interval ?? 10}"/></fieldGroup>`
                    : `<fieldGroup base="${n}"><rangePr groupBy="${{ days: `days`, months: `months`, quarters: `quarters`, years: `years` }[r.groupBy] ?? `months`}"/></fieldGroup>`),
              `<cacheField name="${e(t)}" numFmtId="0"><sharedItems count="${_[n].length}">${i}</sharedItems>${a}</cacheField>`
            );
          }
          if (y[n]) {
            let r = l.map((e) => Number(e[n])).filter((e) => !isNaN(e)),
              i = r.length ? Math.min(...r) : 0,
              a = r.length ? Math.max(...r) : 0;
            return `<cacheField name="${e(t)}" numFmtId="0"><sharedItems containsSemiMixedTypes="0" containsString="0" containsNumber="1" minValue="${i}" maxValue="${a}"/></cacheField>`;
          }
          return `<cacheField name="${e(t)}" numFmtId="0"><sharedItems/></cacheField>`;
        })
        .join(``) + j,
    F = l
      .map(
        (t) =>
          `<r>${c
            .map((n, r) => {
              let i = t[r];
              if (k.has(r)) {
                let e = i == null ? `` : String(i);
                return `<x v="${v[r].get(e) ?? 0}"/>`;
              }
              return typeof i == `number`
                ? `<n v="${i}"/>`
                : typeof i == `boolean`
                  ? `<b v="${+!!i}"/>`
                  : `<s v="${e(String(i ?? ``))}"/>`;
            })
            .join(``)}</r>`,
      )
      .join(``),
    I =
      c
        .map((e, t) => {
          let n = m.includes(t),
            r = h.includes(t),
            i = g.includes(t);
          if (n) {
            let e = _[t].map((e, t) => `<item x="${t}"/>`).join(``) + `<item t="default"/>`;
            return `<pivotField axis="axisRow" showAll="0"><items count="${_[t].length + 1}">${e}</items></pivotField>`;
          }
          if (r) {
            let e = _[t].map((e, t) => `<item x="${t}"/>`).join(``) + `<item t="default"/>`;
            return `<pivotField axis="axisCol" showAll="0"><items count="${_[t].length + 1}">${e}</items></pivotField>`;
          }
          return i ? `<pivotField dataField="1" showAll="0"/>` : `<pivotField showAll="0"/>`;
        })
        .join(``) + A.map(() => `<pivotField dataField="1" showAll="0"/>`).join(``),
    L = ``,
    R = ``;
  if (m.length) {
    L = `<rowFields count="${m.length}">${m.map((e) => `<field x="${e}"/>`).join(``)}</rowFields>`;
    let e = m[0],
      t = _[e].map((e, t) => `<i><x v="${t}"/></i>`).join(``),
      n = d ? `<i t="grand"><x/></i>` : ``;
    R = `<rowItems count="${_[e].length + +!!d}">${t}${n}</rowItems>`;
  }
  let z = ``,
    B = ``;
  if (S) {
    z = `<colFields count="${h.length}">${h.map((e) => `<field x="${e}"/>`).join(``)}</colFields>`;
    let e = h[0],
      t = _[e].map((e, t) => `<i><x v="${t}"/></i>`).join(``),
      n = f ? `<i t="grand"><x/></i>` : ``;
    B = `<colItems count="${_[e].length + +!!f}">${t}${n}</colItems>`;
  } else
    T > 1 &&
      ((z = `<colFields count="1"><field x="-2"/></colFields>`),
      (B = `<colItems count="${T}">${Array.from({ length: T }, (e, t) => `<i><x v="${t}"/></i>`).join(``)}</colItems>`));
  let V = [];
  for (let t = 0; t < i.dataFields.length; t++) {
    let n = i.dataFields[t],
      a = g[t],
      o = r[n.func ?? `sum`] ?? `sum`,
      s = n.name ?? `Sum of ${n.field}`;
    V.push(
      `<dataField name="${e(s)}" fld="${a}" subtotal="${o}" showDataAs="normal" baseField="0" baseItem="0"/>`,
    );
  }
  for (let t = 0; t < A.length; t++) {
    let n = A[t],
      r = u + t;
    V.push(
      `<dataField name="${e(n.name)}" fld="${r}" subtotal="sum" showDataAs="normal" baseField="0" baseItem="0"/>`,
    );
  }
  let H = `<dataFields count="${V.length}">${V.join(``)}</dataFields>`,
    U = [d ? `` : `rowGrandTotals="0"`, f ? `` : `colGrandTotals="0"`].filter(Boolean),
    W = U.length ? ` ` + U.join(` `) : ``,
    G = e(i.style ?? `PivotStyleMedium9`);
  return {
    pivotTableXml: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" name="${e(i.name)}" cacheId="${s}" applyNumberFormats="0" applyBorderFormats="0" applyFontFormats="0" applyPatternFormats="0" applyAlignmentFormats="0" applyWidthHeightFormats="1" dataCaption="Values" updatedVersion="6" minRefreshableVersion="3" useAutoFormatting="1" itemPrintTitles="1" createdVersion="6" indent="0" outline="1" outlineData="1" multipleFieldFilters="0"${W}>
<location ref="${D}" firstHeaderRow="1" firstDataRow="2" firstDataCol="${O}"/>
<pivotFields count="${M}">${I}</pivotFields>
${L}${R}${z}${B}${H}
<pivotTableStyleInfo name="${G}" showRowHeaders="1" showColHeaders="1" showRowStripes="0" showColStripes="0" showLastColumn="1"/>
</pivotTableDefinition>`,
    cacheDefXml: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId1" refreshedBy="ExcelForge" refreshedDate="45000" createdVersion="6" refreshedVersion="6" minRefreshableVersion="3" recordCount="${l.length}" saveData="1">
<cacheSource type="worksheet"><worksheetSource ref="${i.sourceRef}" sheet="${e(i.sourceSheet)}"/></cacheSource>
<cacheFields count="${M}">${P}</cacheFields>
</pivotCacheDefinition>`,
    cacheRecordsXml: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotCacheRecords xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" count="${l.length}">${F}</pivotCacheRecords>`,
  };
}
export { i as buildPivotTableFiles };
