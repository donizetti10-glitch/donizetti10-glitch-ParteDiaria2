import { a as e } from "./helpers-nemQBT-g.js";
import { a as t, r as n, t as r } from "./FormControlBuilderCommon-2HvmNkMR.js";
function i(t) {
  if (t._ctrlPropXml) return t._ctrlPropXml;
  let i = [`objectType="${n[t.type] ?? `Button`}"`, `lockText="1"`];
  return (
    t.linkedCell && i.push(`fmlaLink="${e(t.linkedCell)}"`),
    t.inputRange && i.push(`fmlaRange="${e(t.inputRange)}"`),
    t.checked !== void 0 && i.push(`checked="${r[t.checked] ?? `0`}"`),
    t.dropLines !== void 0 && i.push(`dropLines="${t.dropLines}"`),
    t.min !== void 0 && i.push(`min="${t.min}"`),
    t.max !== void 0 && i.push(`max="${t.max}"`),
    t.inc !== void 0 && i.push(`inc="${t.inc}"`),
    t.page !== void 0 && i.push(`page="${t.page}"`),
    t.val !== void 0 && i.push(`val="${t.val}"`),
    t.selType &&
      i.push(
        `selType="${{ single: `Single`, multi: `Multi`, extend: `Extend` }[t.selType] ?? `Single`}"`,
      ),
    t.noThreeD && i.push(`noThreeD="1"`),
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<formControlPr xmlns="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" ${i.join(` `)}/>`
  );
}
function a(e, t) {
  return `${e.col}, ${e.colOff ?? 0}, ${e.row}, ${e.rowOff ?? 0}, ${t.col}, ${t.colOff ?? 0}, ${t.row}, ${t.rowOff ?? 0}`;
}
function o(n, i) {
  if (n._vmlShapeXml) return n._vmlShapeXml;
  let o = t[n.type] ?? `Button`,
    s = a(n.from, n.to),
    c = [];
  (c.push(`<x:Anchor>${s}</x:Anchor>`),
    c.push(`<x:PrintObject>False</x:PrintObject>`),
    c.push(`<x:AutoFill>False</x:AutoFill>`),
    n.macro && c.push(`<x:FmlaMacro>${e(n.macro)}</x:FmlaMacro>`),
    n.linkedCell && c.push(`<x:FmlaLink>${e(n.linkedCell)}</x:FmlaLink>`),
    n.inputRange && c.push(`<x:FmlaRange>${e(n.inputRange)}</x:FmlaRange>`),
    n.checked !== void 0 && c.push(`<x:Checked>${r[n.checked] ?? `0`}</x:Checked>`),
    n.dropLines !== void 0 && c.push(`<x:DropLines>${n.dropLines}</x:DropLines>`),
    n.dropStyle && c.push(`<x:DropStyle>${e(n.dropStyle)}</x:DropStyle>`),
    n.min !== void 0 && c.push(`<x:Min>${n.min}</x:Min>`),
    n.max !== void 0 && c.push(`<x:Max>${n.max}</x:Max>`),
    n.inc !== void 0 && c.push(`<x:Inc>${n.inc}</x:Inc>`),
    n.page !== void 0 && c.push(`<x:Page>${n.page}</x:Page>`),
    n.val !== void 0 && c.push(`<x:Val>${n.val}</x:Val>`),
    n.selType &&
      c.push(
        `<x:Sel>${{ single: `Single`, multi: `Multi`, extend: `Extend` }[n.selType] ?? `Single`}</x:Sel>`,
      ),
    n.noThreeD && c.push(`<x:NoThreeD/>`),
    n.isDefault && c.push(`<x:Default/>`),
    n.isDismiss && c.push(`<x:Dismiss/>`),
    n.isCancel && c.push(`<x:Cancel/>`));
  let l = ``;
  n.text &&
    (n.type === `button` ||
      n.type === `checkBox` ||
      n.type === `optionButton` ||
      n.type === `groupBox` ||
      n.type === `label` ||
      n.type === `dialog`) &&
    ((l = `<v:textbox style="mso-direction-alt:auto"><div style="text-align:${n.type === `button` ? `center` : `left`}"><font face="Calibri" size="220" color="#000000">${e(n.text)}</font></div></v:textbox>`),
    c.push(`<x:TextHAlign>Center</x:TextHAlign>`),
    c.push(`<x:TextVAlign>Center</x:TextVAlign>`));
  let u = (n.from.col * 64 + (n.from.colOff ?? 0)) * 0.75,
    d = (n.from.row * 20 + (n.from.rowOff ?? 0)) * 0.75,
    f = (n.to.col - n.from.col) * 64 * 0.75,
    p = (n.to.row - n.from.row) * 20 * 0.75,
    m = n.text ?? `${n.type}_${i}`,
    h = n.type === `dialog` ? `` : ` filled="f" stroked="f"`;
  return `<v:shape o:spid="_x0000_s${i}" id="${e(m)}" type="#_x0000_t201" style="position:absolute;margin-left:${u.toFixed(1)}pt;margin-top:${d.toFixed(1)}pt;width:${f.toFixed(1)}pt;height:${p.toFixed(1)}pt;z-index:${i}"${h} o:insetmode="auto"><o:lock v:ext="edit" rotation="t"/>${l}<x:ClientData ObjectType="${o}">${c.join(``)}</x:ClientData></v:shape>`;
}
function s(e, t) {
  return `<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel">
<o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="1"/></o:shapelayout>
<v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe"><v:stroke joinstyle="miter"/><v:path gradientshapeok="t" o:connecttype="rect"/></v:shapetype>
<v:shapetype id="_x0000_t201" coordsize="21600,21600" o:spt="201" path="m,l,21600r21600,l21600,xe"><v:stroke joinstyle="miter"/><v:path shadowok="f" o:extrusionok="f" strokeok="f" fillok="f" o:connecttype="rect"/><o:lock v:ext="edit" shapetype="t"/></v:shapetype>
${e.join(`
`)}
${t.join(`
`)}
</xml>`;
}
export { i as buildCtrlPropXml, o as buildFormControlVmlShape, s as buildVmlWithControls };
