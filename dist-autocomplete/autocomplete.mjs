// src/renderer/univerAdapter.ts
var HEADER = ["\u5E8F\u53F7", "\u65E5\u671F", "\u8D27\u54C1\u540D\u79F0", "\u5355\u4F4D", "\u6570\u91CF", "\u5355\u4EF7", "\u91D1\u989D", "\u8C03\u8D27\u4EBA", "\u5907\u6CE8"];
var COL_COUNT = HEADER.length;
var EXCEL_EPOCH = Date.UTC(1899, 11, 30);
function parseDateParts(input) {
  const s = String(input ?? "").trim();
  if (!s) return null;
  const make = (y, m2, d) => {
    if (y < 100) y += y < 70 ? 2e3 : 1900;
    if (m2 < 1 || m2 > 12 || d < 1 || d > 31) return null;
    const dt = new Date(y, m2 - 1, d);
    if (dt.getFullYear() !== y || dt.getMonth() !== m2 - 1 || dt.getDate() !== d) return null;
    return { y, m: m2, d };
  };
  let m = s.match(/^(\d{4}|\d{2})\s*[-/.年]\s*(\d{1,2})\s*[-/.月]\s*(\d{1,2})\s*日?$/);
  if (m) return make(Number(m[1]), Number(m[2]), Number(m[3]));
  m = s.match(/^(\d{4})(\d{2})(\d{2})$/);
  if (m) return make(Number(m[1]), Number(m[2]), Number(m[3]));
  m = s.match(/^(\d{1,2})\s*[-/月]\s*(\d{1,2})\s*日?$/);
  if (m) return make((/* @__PURE__ */ new Date()).getFullYear(), Number(m[1]), Number(m[2]));
  if (/^\d{4,5}$/.test(s)) {
    const serial = Number(s);
    if (serial >= 1 && serial <= 6e4) {
      const dt = new Date(EXCEL_EPOCH + Math.round(serial) * 864e5);
      return make(dt.getUTCFullYear(), dt.getUTCMonth() + 1, dt.getUTCDate());
    }
  }
  return null;
}
function datePartsToSerial(p) {
  return Math.round((Date.UTC(p.y, p.m - 1, p.d) - EXCEL_EPOCH) / 864e5);
}
function dateStrToSerial(input) {
  const p = parseDateParts(input);
  return p ? datePartsToSerial(p) : null;
}

// src/renderer/lib/autocomplete.ts
var AC_MAX_SUGGESTIONS = 12;
var HEADER2 = ["\u5E8F\u53F7", "\u65E5\u671F", "\u8D27\u54C1\u540D\u79F0", "\u5355\u4F4D", "\u6570\u91CF", "\u5355\u4EF7", "\u91D1\u989D", "\u8C03\u8D27\u4EBA", "\u5907\u6CE8"];
var COL_COUNT2 = HEADER2.length;
var DATE_COL_INDEX = 1;
function toItem(raw, isDate) {
  if (isDate) {
    let num = null;
    if (typeof raw === "number") num = raw;
    else {
      if (/^-?\d+(\.\d+)?$/.test(String(raw).trim())) num = Number(String(raw).trim());
      else num = dateStrToSerial(String(raw));
    }
    if (num != null && isFinite(num)) {
      const d = serialToDateStr(num);
      return { display: d, raw: num };
    }
    return { display: String(raw), raw: String(raw) };
  }
  return { display: String(raw), raw };
}
function serialToDateStr(serial) {
  if (!isFinite(serial)) return "";
  const ms = (serial - 25569) * 86400 * 1e3;
  const d = new Date(ms);
  const y = d.getUTCFullYear();
  const m = String(d.getUTCMonth() + 1).padStart(2, "0");
  const day = String(d.getUTCDate()).padStart(2, "0");
  return `${y}-${m}-${day}`;
}
function extractColumnValues(wb, col, sheetName = "bill") {
  const sheets = wb?.sheets;
  const sheet = sheets ? sheets[sheetName] ?? Object.values(sheets)[0] : void 0;
  const cd = sheet?.cellData;
  if (!cd) return [];
  const agg = /* @__PURE__ */ new Map();
  for (const r of Object.keys(cd)) {
    const cell = cd[Number(r)]?.[col];
    if (cell && cell.v != null) {
      const s = String(cell.v).trim();
      if (s) {
        const prev = agg.get(s);
        if (prev) {
          prev.freq += 1;
          prev.rows.push(Number(r));
        } else {
          agg.set(s, { rows: [Number(r)], freq: 1 });
        }
      }
    }
  }
  const out = [];
  for (const [value, { rows, freq }] of agg.entries()) {
    for (const row of rows) {
      out.push({ value, row, freq });
    }
  }
  return out;
}
function computeSuggestions(partial, col, currentRow, columnVals, history) {
  const p = partial.trim().toLowerCase();
  if (!p) return [];
  const isDate = col === DATE_COL_INDEX;
  const colVals = columnVals.filter((x) => x.row < currentRow);
  const cands = /* @__PURE__ */ new Map();
  for (const c of colVals) {
    const it = toItem(c.value, isDate);
    const prev = cands.get(it.display);
    if (prev) {
      prev.freq += c.freq;
      prev.row = Math.max(prev.row, c.row);
    } else {
      cands.set(it.display, { ...it, freq: c.freq, row: c.row });
    }
  }
  for (const h of new Set(history)) {
    const it = toItem(h, isDate);
    const prev = cands.get(it.display);
    if (prev) prev.freq += 2;
    else cands.set(it.display, { ...it, freq: 2, row: -1 });
  }
  const all = [...cands.values()];
  let matched = all.filter((x) => x.display.toLowerCase().startsWith(p));
  if (!matched.length) matched = all.filter((x) => x.display.toLowerCase().includes(p));
  matched.sort((a, b) => {
    const ra = a.row < 0 ? -Infinity : a.row;
    const rb = b.row < 0 ? -Infinity : b.row;
    if (ra !== rb) return rb - ra;
    if (a.freq !== b.freq) return b.freq - a.freq;
    return a.display.localeCompare(b.display);
  });
  return matched.slice(0, AC_MAX_SUGGESTIONS);
}
export {
  AC_MAX_SUGGESTIONS,
  COL_COUNT2 as COL_COUNT,
  DATE_COL_INDEX,
  HEADER2 as HEADER,
  computeSuggestions,
  extractColumnValues,
  serialToDateStr,
  toItem
};
