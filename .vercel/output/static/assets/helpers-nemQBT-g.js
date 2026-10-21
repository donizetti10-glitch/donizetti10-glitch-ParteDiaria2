var e = [];
function t(t) {
  if (e[t]) return e[t];
  let n = ``,
    r = t;
  for (; r > 0;) {
    let e = (r - 1) % 26;
    ((n = String.fromCharCode(65 + e) + n), (r = Math.floor((r - 1) / 26)));
  }
  return ((e[t] = n), n);
}
function n(e) {
  let t = 0;
  for (let n = 0; n < e.length; n++) t = t * 26 + (e.charCodeAt(n) - 64);
  return t;
}
function r(e) {
  let t = e.match(/^(\$?)([A-Z]+)(\$?)(\d+)$/);
  if (!t) throw Error(`Invalid cell ref: ${e}`);
  return { col: n(t[2]), row: parseInt(t[4], 10) };
}
function i(e, n, r = !1) {
  let i = r ? `$` : ``;
  return `${i}${t(n)}${i}${e}`;
}
function a(e) {
  let [t, n] = e.split(`:`),
    i = r(t.replace(/\$/g, ``)),
    a = n ? r(n.replace(/\$/g, ``)) : i;
  return { startRow: i.row, startCol: i.col, endRow: a.row, endCol: a.col };
}
var o = { "&": `&amp;`, "<": `&lt;`, ">": `&gt;`, '"': `&quot;`, "'": `&apos;` },
  s = /[&<>"']/g;
function c(e) {
  return s.test(e) ? ((s.lastIndex = 0), e.replace(s, (e) => o[e])) : e;
}
var l = new TextEncoder();
function u(e) {
  return l.encode(e);
}
function d(e, t = !1) {
  let n = t ? Date.UTC(1904, 0, 1) : Date.UTC(1899, 11, 30);
  return (e.getTime() - n) / 864e5;
}
function f(e) {
  return Math.round((e * 914400) / 96);
}
function p(e) {
  let t = atob(e),
    n = new Uint8Array(t.length);
  for (let e = 0; e < t.length; e++) n[e] = t.charCodeAt(e);
  return n;
}
export { c as a, f as c, d as i, u as l, r as n, i as o, t as r, a as s, p as t };
