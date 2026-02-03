/* global XLSX, JSZip, saveAs, html2pdf */

/**
 * ProcessXLS — browser-only tool:
 * - Reads .xls/.xlsx
 * - Sheet1: table with invoice data (first row = headers)
 * - Invoice: HTML template (inv.html)
 * - Act: HTML template (act.html)
 * - Replaces placeholders: {Key} / {{Key}}
 * - Exports PDFs via html2pdf and downloads (ZIP with 2 PDFs per row, or 2 combined PDFs)
 */

const ui = {
  fileInput: document.getElementById("fileInput"),
  invHtmlInput: document.getElementById("invHtmlInput"),
  actHtmlInput: document.getElementById("actHtmlInput"),
  btnLoad: document.getElementById("btnLoad"),
  btnPreview: document.getElementById("btnPreview"),
  btnRun: document.getElementById("btnRun"),
  btnReset: document.getElementById("btnReset"),
  status: document.getElementById("status"),
  preview: document.getElementById("preview"),
  nameColumn: document.getElementById("nameColumn"),
  mode: document.getElementById("mode"),
  invPrefix: document.getElementById("invPrefix"),
  invStart: document.getElementById("invStart"),
};

const state = {
  file: null,
  workbook: null,
  sheetNames: [],
  dataHeaders: [],
  dataRows: [], // array of objects {header:value}
  templateInvoice: null, // ws object
  templateAct: null, // ws object
  invHtmlFile: null,
  invHtmlText: null,
  invHtmlParsed: null, // { stylesText, bodyHtmlWithPlaceholders }
  actHtmlFile: null,
  actHtmlText: null,
  actHtmlParsed: null, // { stylesText, bodyHtmlWithPlaceholders }
};

function setStatus(lines) {
  ui.status.textContent = Array.isArray(lines) ? lines.join("\n") : String(lines ?? "");
}

function escapeFilename(name) {
  return String(name)
    .replace(/[<>:"/\\|?*\u0000-\u001F]/g, "_")
    .replace(/\s+/g, " ")
    .trim()
    .slice(0, 140);
}

function isRowEmpty(obj) {
  return Object.values(obj).every((v) => v == null || String(v).trim() === "");
}

function readFileAsArrayBuffer(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onerror = () => reject(new Error("Не удалось прочитать файл"));
    reader.onload = () => resolve(reader.result);
    reader.readAsArrayBuffer(file);
  });
}

function readFileAsText(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onerror = () => reject(new Error("Не удалось прочитать файл"));
    reader.onload = () => resolve(String(reader.result ?? ""));
    reader.readAsText(file, "utf-8");
  });
}

function normalizeHeader(h) {
  return String(h ?? "").trim();
}

function excelColName(n0) {
  // 0 -> A, 25 -> Z, 26 -> AA
  let n = n0;
  let s = "";
  while (n >= 0) {
    s = String.fromCharCode(65 + (n % 26)) + s;
    n = Math.floor(n / 26) - 1;
  }
  return s;
}

function looksLikeHeaderRow(row0) {
  const nonEmpty = (row0 || []).filter((v) => v != null && String(v).trim() !== "");
  if (nonEmpty.length < 2) return false;
  // If it contains any numbers or dates, it's likely data, not headers
  const hasNumberOrDate = nonEmpty.some((v) => typeof v === "number" || v instanceof Date);
  if (hasNumberOrDate) return false;
  // If most cells are long sentences, also likely data
  const longish = nonEmpty.filter((v) => String(v).length > 30).length;
  if (longish / nonEmpty.length > 0.6) return false;
  return true;
}

function parseDataSheet(ws) {
  // header row + data rows
  const range = ws["!ref"] ? XLSX.utils.decode_range(ws["!ref"]) : null;
  const colCount = range ? range.e.c - range.s.c + 1 : 0;
  const table = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: "" });
  if (!table || table.length < 1) {
    throw new Error("1-й лист пустой.");
  }

  // Find the first non-empty row (so we can tolerate top padding)
  let firstNonEmpty = -1;
  for (let i = 0; i < table.length; i++) {
    const r = table[i] || [];
    const any = r.some((v) => v != null && String(v).trim() !== "");
    if (any) {
      firstNonEmpty = i;
      break;
    }
  }
  if (firstNonEmpty === -1) throw new Error("1-й лист пустой.");

  const row0 = table[firstNonEmpty] || [];
  const hasHeader = looksLikeHeaderRow(row0);

  let headers = [];
  let dataStart = firstNonEmpty;
  if (hasHeader) {
    const rawHeaders = row0.map(normalizeHeader);
    const hasAnyHeader = rawHeaders.some(Boolean);
    if (!hasAnyHeader) throw new Error("Не нашёл заголовки в строке заголовков.");
    headers = rawHeaders.map((h, idx) => (h ? h : `__COL_${idx + 1}`));
    dataStart = firstNonEmpty + 1;
  } else {
    // No headers: use Excel-like letters based on !ref width (fallback to max row length)
    const width = colCount || Math.max(...table.map((r) => (r ? r.length : 0)), 0);
    headers = Array.from({ length: width }, (_, i) => excelColName(i));
    dataStart = firstNonEmpty;
  }

  const rows = [];
  for (let i = dataStart; i < table.length; i++) {
    const rowArr = table[i] || [];
    const obj = {};
    headers.forEach((h, idx) => {
      obj[h] = rowArr[idx] ?? "";
    });
    if (!isRowEmpty(obj)) rows.push(obj);
  }

  if (rows.length === 0) {
    throw new Error("На 1-м листе нет строк данных.");
  }

  return { headers, rows, hasHeader };
}

function deepCloneWorksheet(ws) {
  // worksheet is a plain object with cell addresses, !ref, !merges, etc.
  return JSON.parse(JSON.stringify(ws));
}

function formatValue(v) {
  if (v == null) return "";
  if (v instanceof Date) {
    // YYYY-MM-DD (treat Excel dates as local calendar dates)
    const y = v.getFullYear();
    const m = String(v.getMonth() + 1).padStart(2, "0");
    const d = String(v.getDate()).padStart(2, "0");
    return `${y}-${m}-${d}`;
  }
  if (typeof v === "number") return String(v);
  return String(v);
}

function fillTemplateWorksheet(wsTemplate, rowObj) {
  const ws = deepCloneWorksheet(wsTemplate);
  const placeholderDouble = /\{\{([^}]+)\}\}/g; // {{key}}
  const placeholderSingle = /\{([^{}]+)\}/g; // {key}
  const keys = Object.keys(ws);
  for (const k of keys) {
    if (k[0] === "!") continue;
    const cell = ws[k];
    if (!cell || cell.t !== "s") continue;
    const txt = cell.v;
    if (typeof txt !== "string") continue;
    if (txt.indexOf("{") === -1) continue;

    const replaceToken = (token) => {
      const header = String(token ?? "").trim();
      if (!header) return "";
      if (Object.prototype.hasOwnProperty.call(rowObj, header)) return formatValue(rowObj[header]);
      const found = Object.keys(rowObj).find((h) => h.toLowerCase() === header.toLowerCase());
      return found ? formatValue(rowObj[found]) : "";
    };

    let out = txt.replace(placeholderDouble, (_, token) => replaceToken(token));
    out = out.replace(placeholderSingle, (_, token) => replaceToken(token));
    ws[k] = { ...cell, v: out, w: out, t: "s" };
  }
  return ws;
}

function applyAliases(rowObj) {
  // If sheet has no headers, we still want meaningful keys for common placeholders.
  // Detect two known layouts for your "Список":
  // Layout v1 (old): A=описание, B=маршрут, C="а/м", D=номер авто, E=водитель, F=сумма, H=дата
  // Layout v2 (current): A=номер счёта, B=описание, C=маршрут, D="а/м", E=номер авто, F=водитель, G=сумма, I=дата

  const A = rowObj.A ?? rowObj[excelColName(0)];
  const B = rowObj.B ?? rowObj[excelColName(1)];
  const C = rowObj.C ?? rowObj[excelColName(2)];
  const D = rowObj.D ?? rowObj[excelColName(3)];
  const E = rowObj.E ?? rowObj[excelColName(4)];
  const F = rowObj.F ?? rowObj[excelColName(5)];
  const G = rowObj.G ?? rowObj[excelColName(6)];
  const H = rowObj.H ?? rowObj[excelColName(7)];
  const I = rowObj.I ?? rowObj[excelColName(8)];

  const aLooksLikeInvoiceNo =
    typeof A === "number" || (typeof A === "string" && /^\s*\d+\s*$/.test(A) && A.trim().length <= 10);
  const bLooksLikeTu = typeof B === "string" && B.toLowerCase().includes("ту по перевозке");

  const isV2 = aLooksLikeInvoiceNo && bLooksLikeTu;

  const desc = isV2 ? B : A;
  const route = isV2 ? C : B;
  const carPrefix = isV2 ? D : C;
  const plate = isV2 ? E : D;
  const driver = isV2 ? F : E;
  const amount = isV2 ? G : F;
  const date = isV2 ? I : H;
  const invoiceNo = isV2 ? A : null;

  const aliases = {
    ...(invoiceNo != null ? { "номер счёта": invoiceNo } : {}),
    маршрут: route,
    описание: desc,
    авто: `${String(carPrefix ?? "").trim()}${String(plate ?? "").trim()}`.trim(),
    "номер авто": plate,
    водитель: driver,
    сумма: amount,
    дата: date,
    "дата счёта": date,
    "Дата счёта": date, // for templates that use capitalized placeholder
    услуга: `${String(desc ?? "").trim()}${String(route ?? "").trim()}`.trim(),
  };
  return { ...aliases, ...rowObj };
}

function withComputedFields(rowObj, idx, invPrefix, invStart) {
  const n0 = Number.isFinite(invStart) ? invStart : 1;
  const num = n0 + idx;
  const padded = String(num).padStart(4, "0");
  const invoiceNo = `${invPrefix || ""}${padded}`;
  // Only auto-fill invoice number if it's missing in the data
  const hasInvoiceNo =
    Object.prototype.hasOwnProperty.call(rowObj, "номер счёта") &&
    rowObj["номер счёта"] != null &&
    String(rowObj["номер счёта"]).trim() !== "";
  const enriched = { ...rowObj, ...(hasInvoiceNo ? {} : { "номер счёта": invoiceNo }), "__row_index": idx + 1 };
  const dateRu = formatDateRu(enriched["дата счёта"] ?? enriched["Дата счёта"] ?? enriched.дата);
  const amount = enriched.сумма ?? enriched["сумма"];
  const sumFmt = formatRubAmount(amount);
  const sumWords = amountToWordsRubKop(amount);
  const desc = String(enriched["описание"] ?? "")
    .replace(/\u00A0/g, " ")
    .replace(/"/g, "")
    .replace(/\s+/g, " ")
    .trim();
  const route = String(enriched["маршрут"] ?? "").replace(/\u00A0/g, " ").replace(/\s+/g, " ").trim();
  const plate = String(enriched["номер авто"] ?? "").replace(/\u00A0/g, " ").replace(/\s+/g, " ").trim();
  const driverRaw = String(enriched["водитель"] ?? "").replace(/\u00A0/g, " ").replace(/\s+/g, " ").trim();
  const driverName = driverRaw.replace(/^вод\.?\s*/i, "").trim();

  const line1 = desc ? `${desc} "` : `"`;
  const line2 = [route, plate ? `а/м ${plate}` : "", "вод."].filter(Boolean).join(" ").trim();
  const line3 = driverName || driverRaw;
  const service = [line1, line2, line3].filter((x) => String(x).trim() !== "").join("\n");

  return {
    ...enriched,
    основание: "договор № 70",
    дата_ру: dateRu,
    сумма_формат: sumFmt,
    сумма_пропись: sumWords,
    услуга: service,
  };
}

function worksheetToRenderableHtml(ws, title) {
  // Attempt to preserve merges & basic structure; styles won't be 1:1 with Excel.
  let html = XLSX.utils.sheet_to_html(ws, { editable: false });
  // sheet_to_html returns a full HTML document; extract table body if possible
  const m = html.match(/<table[\s\S]*<\/table>/i);
  const table = m ? m[0] : "<div>Не удалось отрендерить лист в HTML</div>";
  return `<div class="sheet-wrap"><h3>${title}</h3>${table}</div>`;
}

function parseA1Range(rangeStr) {
  if (!rangeStr) return null;
  const s = String(rangeStr).trim();
  if (!s) return null;
  // accept A1 or A1:B5; also accept $A$1:$B$5
  const cleaned = s.replace(/\$/g, "");
  const parts = cleaned.split(":");
  if (parts.length === 1) {
    const c = XLSX.utils.decode_cell(parts[0]);
    return { s: c, e: c };
  }
  if (parts.length === 2) {
    const sCell = XLSX.utils.decode_cell(parts[0]);
    const eCell = XLSX.utils.decode_cell(parts[1]);
    return { s: sCell, e: eCell };
  }
  return null;
}

function parsePrintAreaFromWorkbook(wb, sheetName) {
  // Excel stores as defined name: _xlnm.Print_Area => 'Sheet'!$A$1:$I$55
  try {
    const names = wb?.Workbook?.Names;
    if (!Array.isArray(names)) return null;
    const candidates = names.filter(
      (n) =>
        (n?.Name === "_xlnm.Print_Area" || n?.Name === "Print_Area") &&
        typeof n?.Ref === "string" &&
        n.Ref.includes("!")
    );
    for (const entry of candidates) {
      // XLSX.js may provide local sheet scope via numeric Sheet index
      if (Number.isFinite(entry?.Sheet) && wb?.SheetNames?.[entry.Sheet] === sheetName) {
        const ref = String(entry.Ref).trim();
        const m = ref.match(/!([^!]+)$/);
        if (m) return m[1].replace(/\$/g, "");
      }
      // Or multiple sheet refs separated by commas
      const refs = String(entry.Ref)
        .split(",")
        .map((x) => x.trim())
        .filter(Boolean);
      for (const ref of refs) {
        const m =
          ref.match(/^'?([^']+)'?!([A-Z]+\$?\d+:\$?[A-Z]+\$?\d+)$/i) ||
          ref.match(/^([^!]+)!([A-Z]+\$?\d+:\$?[A-Z]+\$?\d+)$/i);
        if (!m) continue;
        const sn = m[1];
        const r = m[2];
        if (sn === sheetName) return r.replace(/\$/g, "");
      }
    }
    return null;
  } catch {
    return null;
  }
}

function getMergeMap(ws) {
  const merges = ws?.["!merges"] || [];
  const topLeftToSpan = new Map();
  const covered = new Set();
  for (const m of merges) {
    const key = `${m.s.r},${m.s.c}`;
    topLeftToSpan.set(key, { rowspan: m.e.r - m.s.r + 1, colspan: m.e.c - m.s.c + 1 });
    for (let r = m.s.r; r <= m.e.r; r++) {
      for (let c = m.s.c; c <= m.e.c; c++) {
        if (r === m.s.r && c === m.s.c) continue;
        covered.add(`${r},${c}`);
      }
    }
  }
  return { topLeftToSpan, covered };
}

function worksheetRangeToHtml(ws, rangeA1, title) {
  const range = parseA1Range(rangeA1) || (ws?.["!ref"] ? XLSX.utils.decode_range(ws["!ref"]) : null);
  if (!range) return `<div class="sheet-wrap"><h3>${title}</h3><div>Пустой лист</div></div>`;

  const { topLeftToSpan, covered } = getMergeMap(ws);
  const rows = [];
  const colgroup = [];
  const cols = ws?.["!cols"] || [];
  for (let c = range.s.c; c <= range.e.c; c++) {
    const col = cols[c];
    const wpx = col?.wpx;
    const wch = col?.wch;
    let style = "";
    if (Number.isFinite(wpx)) style = ` style="width:${wpx}px"`;
    else if (Number.isFinite(wch)) style = ` style="width:${Math.max(20, wch * 7)}px"`;
    colgroup.push(`<col${style}>`);
  }
  rows.push(`<div class="sheet-wrap"><h3>${title}</h3><table style="table-layout:fixed">`);
  if (colgroup.length) rows.push(`<colgroup>${colgroup.join("")}</colgroup>`);
  for (let r = range.s.r; r <= range.e.r; r++) {
    const tds = [];
    for (let c = range.s.c; c <= range.e.c; c++) {
      const posKey = `${r},${c}`;
      if (covered.has(posKey)) continue;
      const addr = XLSX.utils.encode_cell({ r, c });
      const cell = ws?.[addr];
      const raw = cell ? cell.v : "";
      const val = formatValue(raw);
      const span = topLeftToSpan.get(posKey);
      const rs = span?.rowspan ? ` rowspan="${span.rowspan}"` : "";
      const cs = span?.colspan ? ` colspan="${span.colspan}"` : "";
      tds.push(`<td${rs}${cs}>${escapeHtml(val)}</td>`);
    }
    rows.push(`<tr>${tds.join("")}</tr>`);
  }
  rows.push(`</table></div>`);
  return rows.join("");
}

function escapeHtml(s) {
  return String(s ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/\"/g, "&quot;")
    .replace(/'/g, "&#039;");
}

function formatDateRu(v) {
  // supports Date, ISO string, dd.mm.yyyy, etc.
  if (v instanceof Date) {
    const dd = String(v.getDate()).padStart(2, "0");
    const mm = String(v.getMonth() + 1).padStart(2, "0");
    const yyyy = v.getFullYear();
    return `${dd}.${mm}.${yyyy}`;
  }
  const s = String(v ?? "").trim();
  if (!s) return "";
  const iso = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (iso) return `${iso[3]}.${iso[2]}.${iso[1]}`;
  const ru = s.match(/^(\d{2})\.(\d{2})\.(\d{4})/);
  if (ru) return `${ru[1]}.${ru[2]}.${ru[3]}`;
  return s;
}

function formatRubAmount(v) {
  const n = typeof v === "number" ? v : parseFloat(String(v).replace(",", ".").replace(/\s/g, ""));
  const x = Number.isFinite(n) ? n : 0;
  const fixed = x.toFixed(2);
  const [intPart, frac] = fixed.split(".");
  const withSep = intPart.replace(/\B(?=(\d{3})+(?!\d))/g, "\u00A0");
  return `${withSep},${frac}\u00A0₽`;
}

function chooseForm(n, one, two, five) {
  const n10 = n % 10;
  const n100 = n % 100;
  if (n100 >= 11 && n100 <= 19) return five;
  if (n10 === 1) return one;
  if (n10 >= 2 && n10 <= 4) return two;
  return five;
}

function numToWordsRu(n, female = false) {
  const onesM = ["ноль", "один", "два", "три", "четыре", "пять", "шесть", "семь", "восемь", "девять"];
  const onesF = ["ноль", "одна", "две", "три", "четыре", "пять", "шесть", "семь", "восемь", "девять"];
  const teens = [
    "десять",
    "одиннадцать",
    "двенадцать",
    "тринадцать",
    "четырнадцать",
    "пятнадцать",
    "шестнадцать",
    "семнадцать",
    "восемнадцать",
    "девятнадцать",
  ];
  const tens = [
    "",
    "",
    "двадцать",
    "тридцать",
    "сорок",
    "пятьдесят",
    "шестьдесят",
    "семьдесят",
    "восемьдесят",
    "девяносто",
  ];
  const hundreds = [
    "",
    "сто",
    "двести",
    "триста",
    "четыреста",
    "пятьсот",
    "шестьсот",
    "семьсот",
    "восемьсот",
    "девятьсот",
  ];

  const words = [];
  const ones = female ? onesF : onesM;

  const nn = Math.abs(n);
  const h = Math.floor(nn / 100);
  const t = Math.floor((nn % 100) / 10);
  const o = nn % 10;
  if (h) words.push(hundreds[h]);
  if (t === 1) {
    words.push(teens[o]);
  } else {
    if (t) words.push(tens[t]);
    if (o || words.length === 0) words.push(ones[o]);
  }
  return words.join(" ");
}

function amountToWordsRubKop(amount) {
  const n = typeof amount === "number" ? amount : parseFloat(String(amount).replace(",", ".").replace(/\s/g, ""));
  const x = Number.isFinite(n) ? n : 0;
  const rub = Math.floor(x + 1e-9);
  const kop = Math.round((x - rub) * 100);

  const parts = [];
  const triads = [
    { value: Math.floor(rub / 1_000_000_000) % 1000, unit: "млрд" },
    { value: Math.floor(rub / 1_000_000) % 1000, unit: "млн" },
    { value: Math.floor(rub / 1000) % 1000, unit: "тыс" },
    { value: rub % 1000, unit: null },
  ];

  const units = {
    тыс: ["тысяча", "тысячи", "тысяч", true],
    млн: ["миллион", "миллиона", "миллионов", false],
    млрд: ["миллиард", "миллиарда", "миллиардов", false],
  };

  for (const tr of triads) {
    if (!tr.value) continue;
    const isFemale = tr.unit ? units[tr.unit][3] : false;
    parts.push(numToWordsRu(tr.value, isFemale));
    if (tr.unit) {
      const form = chooseForm(tr.value, units[tr.unit][0], units[tr.unit][1], units[tr.unit][2]);
      parts.push(form);
    }
  }
  if (parts.length === 0) parts.push("ноль");

  const rubForm = chooseForm(rub, "рубль", "рубля", "рублей");
  const kopStr = String(kop).padStart(2, "0");
  const kopForm = chooseForm(kop, "копейка", "копейки", "копеек");
  return `${parts.join(" ")} ${rubForm} ${kopStr} ${kopForm}`;
}

async function ensureInvTemplateParsed() {
  if (state.invHtmlParsed) return state.invHtmlParsed;

  let fullHtml = state.invHtmlText;
  if (!fullHtml) {
    if (state.invHtmlFile) {
      fullHtml = await readFileAsText(state.invHtmlFile);
    } else {
      try {
        const res = await fetch("./inv.html", { cache: "no-store" });
        if (!res.ok) throw new Error(`HTTP ${res.status}`);
        fullHtml = await res.text();
      } catch (e) {
        throw new Error(
          `Не смог загрузить inv.html. Запустите страницу через локальный сервер или выберите inv.html вручную. (${e.message || e})`
        );
      }
    }
    state.invHtmlText = fullHtml;
  }

  const parser = new DOMParser();
  const doc = parser.parseFromString(fullHtml, "text/html");
  const stylesText = Array.from(doc.querySelectorAll("style"))
    .map((s) => s.textContent || "")
    .join("\n");
  const bodyHtmlWithPlaceholders = doc.body ? doc.body.innerHTML : fullHtml;
  state.invHtmlParsed = { stylesText, bodyHtmlWithPlaceholders };
  return state.invHtmlParsed;
}

async function ensureActTemplateParsed() {
  if (state.actHtmlParsed) return state.actHtmlParsed;

  let fullHtml = state.actHtmlText;
  if (!fullHtml) {
    if (state.actHtmlFile) {
      fullHtml = await readFileAsText(state.actHtmlFile);
    } else {
      try {
        const res = await fetch("./act.html", { cache: "no-store" });
        if (!res.ok) throw new Error(`HTTP ${res.status}`);
        fullHtml = await res.text();
      } catch (e) {
        throw new Error(
          `Не смог загрузить act.html. Запустите страницу через локальный сервер или выберите act.html вручную. (${e.message || e})`
        );
      }
    }
    state.actHtmlText = fullHtml;
  }

  const parser = new DOMParser();
  const doc = parser.parseFromString(fullHtml, "text/html");
  const stylesText = Array.from(doc.querySelectorAll("style"))
    .map((s) => s.textContent || "")
    .join("\n");
  const bodyHtmlWithPlaceholders = doc.body ? doc.body.innerHTML : fullHtml;
  state.actHtmlParsed = { stylesText, bodyHtmlWithPlaceholders };
  return state.actHtmlParsed;
}

function fillHtmlTemplateFragment(htmlFragment, rowObj) {
  // IMPORTANT: inv.html contains normal CSS braces { ... }.
  // So we only replace tokens like {номер счёта} / {сумма_формат} (no ':' ';' or newlines).
  // Allow Cyrillic incl. Ёё and most "word-ish" tokens, but avoid CSS blocks by rejecting ':' ';' and newlines.
  const placeholderDouble = /\{\{([^}\n\r:;]+)\}\}/g;
  const placeholderSingle = /\{([^{}\n\r:;]+)\}/g;

  const replaceToken = (token) => {
    const key = String(token ?? "").trim();
    if (!key) return "";
    if (Object.prototype.hasOwnProperty.call(rowObj, key)) return escapeHtml(formatValue(rowObj[key]));
    const found = Object.keys(rowObj).find((h) => h.toLowerCase() === key.toLowerCase());
    return found ? escapeHtml(formatValue(rowObj[found])) : "";
  };

  let out = htmlFragment.replace(placeholderDouble, (_, token) => replaceToken(token));
  out = out.replace(placeholderSingle, (_, token) => replaceToken(token));
  return out;
}

function renderInvoiceHtmlFromInvTemplate(rowObj) {
  if (!state.invHtmlParsed) throw new Error("inv.html не загружен");
  const styleTag = state.invHtmlParsed.stylesText ? `<style>${state.invHtmlParsed.stylesText}</style>` : "";
  const body = fillHtmlTemplateFragment(state.invHtmlParsed.bodyHtmlWithPlaceholders, rowObj);
  // override: ensure top-left alignment inside pdf capture
  // Also slightly reduce height to avoid rounding that can create a blank 2nd page.
  const override = `<style>.sheet{margin:0 !important; position:relative; top:0; left:0; height:296.5mm !important; min-height:296.5mm !important; overflow:hidden;}</style>`;
  return `${styleTag}${override}${body}`;
}

function renderActHtmlFromActTemplate(rowObj) {
  if (!state.actHtmlParsed) throw new Error("act.html не загружен");
  const styleTag = state.actHtmlParsed.stylesText ? `<style>${state.actHtmlParsed.stylesText}</style>` : "";
  const body = fillHtmlTemplateFragment(state.actHtmlParsed.bodyHtmlWithPlaceholders, rowObj);
  const override = `<style>.sheet{margin:0 !important; position:relative; top:0; left:0; height:296.5mm !important; min-height:296.5mm !important; overflow:hidden;}</style>`;
  return `${styleTag}${override}${body}`;
}

function templateSeemsTiny(ws) {
  if (!ws || !ws["!ref"]) return true;
  try {
    const r = XLSX.utils.decode_range(ws["!ref"]);
    const rows = r.e.r - r.s.r + 1;
    const cols = r.e.c - r.s.c + 1;
    return rows * cols <= 20; // heuristic
  } catch {
    return true;
  }
}

function shouldUseFallback(templateMode, wsInvoice, wsAct) {
  if (templateMode === "fallback") return true;
  if (templateMode === "excel") return false;
  // auto
  // If user provides explicit print areas, we can still render a full A4 grid even when the sheet is "tiny".
  const hasExplicitPrintArea =
    (ui.printAreaAct && String(ui.printAreaAct.value || "").trim() !== "");
  if (hasExplicitPrintArea) return false;
  return templateSeemsTiny(wsInvoice) || templateSeemsTiny(wsAct);
}

function defaultInvoiceHtml(rowObj) {
  const no = formatValue(rowObj["номер счёта"] || "");
  const date = formatValue(rowObj["дата счёта"] || rowObj.дата || "");
  const route = formatValue(rowObj["маршрут"] || "");
  const car = formatValue(rowObj["авто"] || rowObj["номер авто"] || "");
  const driver = formatValue(rowObj["водитель"] || "");
  const amount = formatValue(rowObj["сумма"] || "");
  const service = formatValue(rowObj["услуга"] || "");

  return `
    <div>
      <h3>Счёт № ${no}</h3>
      <div style="display:flex; justify-content:space-between; gap:12px; font-size:12px; margin-bottom:10px;">
        <div><b>Дата счёта:</b> ${date}</div>
        <div><b>Основание:</b> Перевозка груза</div>
      </div>
      <table>
        <tr><th style="width:32%">Параметр</th><th>Значение</th></tr>
        <tr><td>Маршрут</td><td>${route}</td></tr>
        <tr><td>Автомобиль</td><td>${car}</td></tr>
        <tr><td>Водитель</td><td>${driver}</td></tr>
        <tr><td>Услуга</td><td>${service}</td></tr>
        <tr><td><b>Сумма</b></td><td><b>${amount}</b></td></tr>
      </table>
      <div style="margin-top:14px; font-size:12px;">
        <div>Поставщик: ____________________</div>
        <div style="margin-top:10px;">Покупатель: ____________________</div>
      </div>
    </div>
  `.trim();
}

function defaultActHtml(rowObj) {
  const no = formatValue(rowObj["номер счёта"] || "");
  const date = formatValue(rowObj["дата счёта"] || rowObj.дата || "");
  const route = formatValue(rowObj["маршрут"] || "");
  const amount = formatValue(rowObj["сумма"] || "");
  return `
    <div>
      <h3>Акт № ${no}</h3>
      <div style="display:flex; justify-content:space-between; gap:12px; font-size:12px; margin-bottom:10px;">
        <div><b>Дата:</b> ${date}</div>
        <div><b>К счёту №</b> ${no}</div>
      </div>
      <table>
        <tr><th style="width:32%">Наименование работ/услуг</th><th>Маршрут</th><th style="width:18%">Сумма</th></tr>
        <tr><td>Перевозка груза</td><td>${route}</td><td style="text-align:right">${amount}</td></tr>
      </table>
      <div style="margin-top:14px; font-size:12px;">
        <div>Исполнитель: ____________________</div>
        <div style="margin-top:10px;">Заказчик: ____________________</div>
      </div>
    </div>
  `.trim();
}

function buildTwoPageContainer(invoiceHtml, actHtml) {
  const container = document.createElement("div");
  container.style.width = "794px"; // ~A4 at 96dpi (8.27in*96)
  container.style.padding = "0";
  container.style.margin = "0";
  container.style.background = "white";
  container.style.color = "#111827";

  const pageStyle = `
    <style>
      .pdf-page {
        width: 794px;
        min-height: 1123px; /* ~A4 */
        padding: 0;
        page-break-after: always;
      }
      .pdf-page:last-child { page-break-after: auto; }
      .act-page { padding: 18px; }
      .act-page table { width: 100%; border-collapse: collapse; }
      .act-page td, .act-page th { border: 1px solid #cbd5e1; padding: 4px 6px; font-size: 11px; vertical-align: top; }
      .act-page h3 { margin: 0 0 8px; font-size: 13px; }
    </style>
  `;

  container.innerHTML = `${pageStyle}
    <div class="pdf-page invoice-page">${invoiceHtml}</div>
    <div class="pdf-page act-page">${actHtml}</div>`;
  return container;
}

async function htmlContainerToPdfBlob(container, filenameBase) {
  // html2pdf options tuned for readability
  const opt = {
    margin: 0,
    filename: `${filenameBase}.pdf`,
    image: { type: "jpeg", quality: 0.95 },
    html2canvas: { scale: 2, useCORS: true, backgroundColor: "#ffffff" },
    jsPDF: { unit: "pt", format: "a4", orientation: "portrait" },
    pagebreak: { mode: ["css", "legacy"] },
  };
  const worker = html2pdf().set(opt).from(container);
  const pdf = await worker.toPdf().get("pdf");
  return pdf.output("blob");
}

async function htmlFragmentToPdfBlob(fragmentHtml, filenameBase) {
  // Render a single A4 document from HTML fragment (expects `.sheet` root).
  const host = document.createElement("div");
  host.style.position = "fixed";
  host.style.left = "-100000px";
  host.style.top = "0";
  host.style.background = "white";
  host.innerHTML = fragmentHtml;
  document.body.appendChild(host);
  try {
    const target = host.querySelector(".sheet") || host;
    const opt = {
      margin: 0,
      filename: `${filenameBase}.pdf`,
      image: { type: "jpeg", quality: 0.98 },
      html2canvas: { scale: 2, useCORS: true, backgroundColor: "#ffffff", scrollY: 0 },
      jsPDF: { unit: "pt", format: "a4", orientation: "portrait" },
      pagebreak: { mode: ["css", "legacy"] },
    };
    const worker = html2pdf().set(opt).from(target);
    const pdf = await worker.toPdf().get("pdf");
    return pdf.output("blob");
  } finally {
    host.remove();
  }
}

function assertDeps() {
  const missing = [];
  if (!window.XLSX) missing.push("xlsx");
  if (!window.JSZip) missing.push("jszip");
  if (!window.saveAs) missing.push("file-saver");
  if (!window.html2pdf) missing.push("html2pdf.js");
  if (missing.length) {
    throw new Error(`Не загрузились библиотеки: ${missing.join(", ")}. Проверьте доступ к CDN или скачайте библиотеки локально.`);
  }
}

function resetAll() {
  state.file = null;
  state.workbook = null;
  state.sheetNames = [];
  state.dataHeaders = [];
  state.dataRows = [];
  state.templateInvoice = null;
  state.templateAct = null;
  state.invHtmlFile = null;
  state.invHtmlText = null;
  state.invHtmlParsed = null;
  state.actHtmlFile = null;
  state.actHtmlText = null;
  state.actHtmlParsed = null;
  ui.fileInput.value = "";
  if (ui.invHtmlInput) ui.invHtmlInput.value = "";
  if (ui.actHtmlInput) ui.actHtmlInput.value = "";
  if (ui.invPrefix) ui.invPrefix.value = "";
  if (ui.invStart) ui.invStart.value = "1";
  ui.preview.innerHTML = `<div class="small">Загрузите файл и нажмите «Превью 1-й строки».</div>`;
  setStatus("");
  ui.btnLoad.disabled = true;
  ui.btnPreview.disabled = true;
  ui.btnRun.disabled = true;
}

function enableAfterFileChosen(enabled) {
  ui.btnLoad.disabled = !enabled;
}

function enableAfterLoaded(enabled) {
  ui.btnPreview.disabled = !enabled;
  ui.btnRun.disabled = !enabled;
}

ui.fileInput.addEventListener("change", () => {
  const f = ui.fileInput.files && ui.fileInput.files[0];
  state.file = f || null;
  enableAfterFileChosen(Boolean(state.file));
  enableAfterLoaded(false);
  if (state.file) setStatus([`Файл: ${state.file.name}`, "Нажмите «Загрузить и проверить»."]);
});

if (ui.invHtmlInput) {
  ui.invHtmlInput.addEventListener("change", () => {
    const f = ui.invHtmlInput.files && ui.invHtmlInput.files[0];
    state.invHtmlFile = f || null;
    state.invHtmlText = null;
    state.invHtmlParsed = null;
    if (state.invHtmlFile) {
      setStatus([`HTML шаблон: ${state.invHtmlFile.name}`, "Шаблон будет использован для печати счёта (inv.html)."]);
    }
  });
}

if (ui.actHtmlInput) {
  ui.actHtmlInput.addEventListener("change", () => {
    const f = ui.actHtmlInput.files && ui.actHtmlInput.files[0];
    state.actHtmlFile = f || null;
    state.actHtmlText = null;
    state.actHtmlParsed = null;
    if (state.actHtmlFile) {
      setStatus([`HTML шаблон: ${state.actHtmlFile.name}`, "Шаблон будет использован для печати акта (act.html)."]);
    }
  });
}

ui.btnReset.addEventListener("click", () => resetAll());

ui.btnLoad.addEventListener("click", async () => {
  try {
    assertDeps();
    if (!state.file) throw new Error("Выберите файл.");
    setStatus(["Читаю файл...", ""]);
    const data = await readFileAsArrayBuffer(state.file);
    const wb = XLSX.read(data, { type: "array", cellDates: true, cellStyles: true });
    state.workbook = wb;
    state.sheetNames = wb.SheetNames || [];
    if (state.sheetNames.length < 1) throw new Error("В книге нет листов.");
    const wsData = wb.Sheets[state.sheetNames[0]];
    if (!wsData) throw new Error("Не удалось получить 1-й лист (данные).");

    const { headers, rows, hasHeader } = parseDataSheet(wsData);
    state.dataHeaders = headers;
    state.dataRows = rows.map(applyAliases);

    await ensureInvTemplateParsed();
    await ensureActTemplateParsed();

    setStatus([
      `Листы: 1) ${state.sheetNames[0]}`,
      `Лист 1: заголовки ${hasHeader ? "обнаружены" : "НЕ обнаружены (использую A,B,C...)"} `,
      `Колонки (лист 1): ${headers.join(", ")}`,
      `Строк данных: ${rows.length}`,
      "Шаблоны: inv.html (счёт) + act.html (акт).",
      "",
      "Плейсхолдеры: {ИмяПоля} или {{ИмяПоля}} (регистр не важен).",
    ]);
    enableAfterLoaded(true);
  } catch (e) {
    enableAfterLoaded(false);
    setStatus([`Ошибка: ${e.message || e}`]);
  }
});

ui.btnPreview.addEventListener("click", async () => {
  try {
    if (!state.dataRows.length) throw new Error("Сначала загрузите файл.");
    await ensureInvTemplateParsed();
    await ensureActTemplateParsed();

    const invPrefix = normalizeHeader(ui.invPrefix.value);
    const invStart = parseInt(String(ui.invStart.value || "1"), 10);
    const row0 = withComputedFields(state.dataRows[0], 0, invPrefix, Number.isFinite(invStart) ? invStart : 1);

    const actFragment = renderActHtmlFromActTemplate(row0);

    const invDoc = `<!doctype html><html lang="ru"><head><meta charset="utf-8">${
      state.invHtmlParsed?.stylesText ? `<style>${state.invHtmlParsed.stylesText}</style>` : ""
    }</head><body>${fillHtmlTemplateFragment(state.invHtmlParsed?.bodyHtmlWithPlaceholders || "", row0)}</body></html>`;
    const actDoc = `<!doctype html><html lang="ru"><head><meta charset="utf-8">${
      state.actHtmlParsed?.stylesText ? `<style>${state.actHtmlParsed.stylesText}</style>` : ""
    }</head><body>${fillHtmlTemplateFragment(state.actHtmlParsed?.bodyHtmlWithPlaceholders || "", row0)}</body></html>`;

    ui.preview.innerHTML = "";
    const mkLabel = (text) => {
      const d = document.createElement("div");
      d.className = "small";
      d.style.marginBottom = "8px";
      d.textContent = text;
      return d;
    };
    const mkIframe = () => {
      const f = document.createElement("iframe");
      f.setAttribute("sandbox", "");
      f.style.width = "100%";
      f.style.height = "520px";
      f.style.border = "1px solid rgba(148,163,184,0.25)";
      f.style.borderRadius = "12px";
      f.style.background = "#fff";
      return f;
    };

    ui.preview.appendChild(mkLabel("Счёт (inv.html)"));
    const invFrame = mkIframe();
    invFrame.srcdoc = invDoc;
    ui.preview.appendChild(invFrame);
    const spacer = document.createElement("div");
    spacer.style.height = "12px";
    ui.preview.appendChild(spacer);
    ui.preview.appendChild(mkLabel("Акт (act.html)"));
    const actFrame = mkIframe();
    actFrame.srcdoc = actDoc;
    ui.preview.appendChild(actFrame);
  } catch (e) {
    setStatus([`Ошибка: ${e.message || e}`]);
  }
});

ui.btnRun.addEventListener("click", async () => {
  try {
    if (!state.dataRows.length) throw new Error("Сначала загрузите файл.");
    assertDeps();
    await ensureInvTemplateParsed();
    await ensureActTemplateParsed();

    const mode = ui.mode.value;
    const nameCol = normalizeHeader(ui.nameColumn.value);
    const invPrefix = normalizeHeader(ui.invPrefix.value);
    const invStartParsed = parseInt(String(ui.invStart.value || "1"), 10);
    const invStart = Number.isFinite(invStartParsed) ? invStartParsed : 1;

    setStatus(["Генерирую PDF... Это может занять время, если строк много.", ""]);
    ui.btnRun.disabled = true;
    ui.btnPreview.disabled = true;
    ui.btnLoad.disabled = true;

    const total = state.dataRows.length;
    const getBaseName = (rowObj, idx) => {
      if (nameCol) {
        const key = Object.keys(rowObj).find((h) => h.toLowerCase() === nameCol.toLowerCase());
        const val = key ? rowObj[key] : "";
        if (val != null && String(val).trim() !== "") return escapeFilename(String(val));
      } else if (rowObj["номер счёта"]) {
        return escapeFilename(String(rowObj["номер счёта"]));
      }
      return `row_${String(idx + 1).padStart(4, "0")}`;
    };

    if (mode === "single") {
      // Two combined PDFs: all invoices and all acts
      const invHost = document.createElement("div");
      invHost.style.position = "fixed";
      invHost.style.left = "-100000px";
      invHost.style.top = "0";
      invHost.style.background = "white";
      const actHost = invHost.cloneNode(false);

      // Put styles once + page-breaks
      invHost.innerHTML = `${state.invHtmlParsed?.stylesText ? `<style>${state.invHtmlParsed.stylesText}</style>` : ""}<style>.sheet{margin:0 !important; page-break-after:always;}</style>`;
      actHost.innerHTML = `${state.actHtmlParsed?.stylesText ? `<style>${state.actHtmlParsed.stylesText}</style>` : ""}<style>.sheet{margin:0 !important; page-break-after:always;}</style>`;
      document.body.appendChild(invHost);
      document.body.appendChild(actHost);
      try {
        for (let i = 0; i < total; i++) {
          const baseRow = state.dataRows[i];
          const rowObj = withComputedFields(baseRow, i, invPrefix, invStart);
          const invBody = fillHtmlTemplateFragment(state.invHtmlParsed?.bodyHtmlWithPlaceholders || "", rowObj);
          const actBody = fillHtmlTemplateFragment(state.actHtmlParsed?.bodyHtmlWithPlaceholders || "", rowObj);
          const invWrap = document.createElement("div");
          invWrap.innerHTML = invBody;
          const actWrap = document.createElement("div");
          actWrap.innerHTML = actBody;
          invHost.appendChild(invWrap);
          actHost.appendChild(actWrap);
          setStatus([`HTML: ${i + 1}/${total}`, ""]);
          await new Promise((r) => setTimeout(r, 0));
        }

        const invBlob = await htmlContainerToPdfBlob(invHost, "invoices");
        saveAs(invBlob, "invoices.pdf");
        const actBlob = await htmlContainerToPdfBlob(actHost, "acts");
        saveAs(actBlob, "acts.pdf");
        setStatus([`Готово: invoices.pdf и acts.pdf (строк: ${total})`]);
      } finally {
        invHost.remove();
        actHost.remove();
      }
    } else {
      // ZIP with two PDFs per row
      const zip = new JSZip();
      for (let i = 0; i < total; i++) {
        const baseRow = state.dataRows[i];
        const rowObj = withComputedFields(baseRow, i, invPrefix, invStart);
        const base = getBaseName(rowObj, i);
        const invoiceHtml = renderInvoiceHtmlFromInvTemplate(rowObj);
        const actHtml = renderActHtmlFromActTemplate(rowObj);
        const invBlob = await htmlFragmentToPdfBlob(invoiceHtml, `invoice_${base}`);
        const actBlob = await htmlFragmentToPdfBlob(actHtml, `act_${base}`);
        zip.file(`invoice/${base}.pdf`, invBlob);
        zip.file(`act/${base}.pdf`, actBlob);
        setStatus([`PDF: ${i + 1}/${total} — ${base} (invoice+act)`, ""]);
        await new Promise((r) => setTimeout(r, 0));
      }
      setStatus(["Собираю ZIP...", ""]);
      const zipBlob = await zip.generateAsync({ type: "blob" });
      saveAs(zipBlob, "pdf_out.zip");
      setStatus([`Готово: pdf_out.zip (PDF файлов: ${total * 2})`]);
    }
  } catch (e) {
    setStatus([`Ошибка: ${e.message || e}`]);
  } finally {
    ui.btnRun.disabled = false;
    ui.btnPreview.disabled = !Boolean(state.dataRows.length);
    ui.btnLoad.disabled = !Boolean(state.file);
  }
});

// boot
resetAll();
enableAfterFileChosen(false);

