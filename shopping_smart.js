#!/usr/bin/env node
const axios = require("axios");
const dotenv = require("dotenv");
dotenv.config();

const KEY = process.env.SERPER_API_KEY;
if (!KEY || /^\s/.test(KEY) || /\s$/.test(KEY)) {
  console.error("Проверь .env: SERPER_API_KEY отсутствует или есть лишние пробелы.");
  process.exit(1);
}

function arg(name, def) {
  const i = process.argv.indexOf(`--${name}`);
  return (i !== -1 && process.argv[i + 1]) ? process.argv[i + 1] : def;
}

// CLI params
const QUERY     = arg("q", "");
const TOPN      = Math.max(1, parseInt(arg("top", "6"), 10));
const GL        = arg("country", "ru");
const HL        = arg("lang", "ru");
const TIMEOUT   = parseInt(arg("timeoutMs", "12000"), 10);
const CLEAN     = arg("clean", "1") === "1";
const OUR_PRICE = arg("our", "");
const DELTA     = parseFloat(arg("delta", "0.6"));
const MIN_SIM   = parseFloat(arg("minSim", "0"));
const WHITELIST = arg("whitelist", "")
  .split(",")
  .map(s => s.trim().toLowerCase())
  .filter(Boolean);

if (!QUERY) {
  console.log(`Пример:
  node shopping_smart.js --q "SF 1207 аккумулятор 12В 7Ач" --top 6 --country ru --lang ru --our "10566,44" --delta 0.6 --minSim 0.55 --whitelist "ЭТМ,Всеинструменты.ру,Сатро-Паладин"
`);
  process.exit(0);
}

// utils
function normalizeNumber(str) {
  if (!str) return null;
  let s = String(str).replace(/\u00A0|\u202F|\u2009/g, "");
  s = s.replace(/[^\d.,-]/g, "");
  if (s.includes(",") && s.includes(".")) { s = s.replace(/\./g, ""); s = s.replace(",", "."); }
  else if (s.includes(",")) { s = s.replace(",", "."); }
  const v = parseFloat(s);
  return Number.isFinite(v) ? v : null;
}
function guessCurrency(strA, strB="") {
  const s = (String(strA || "") + " " + String(strB || "")).toLowerCase();
  if (/[₽]|руб/.test(s)) return "RUB";
  if (/usd|\$/.test(s))  return "USD";
  if (/eur|€/ .test(s))   return "EUR";
  if (/uah|₴|грн/.test(s)) return "UAH";
  if (/kzt|₸|тнг/.test(s)) return "KZT";
  return null;
}
function effectivePrice(x) {
  if (Number.isFinite(x.total)) return { value: x.total, label: "Итоговая цена" };
  if (Number.isFinite(x.price) && Number.isFinite(x.shipping)) return { value: x.price + x.shipping, label: "Цена + доставка" };
  if (Number.isFinite(x.price)) return { value: x.price, label: "Цена" };
  return { value: null, label: "Цена (текст)" };
}
function quantiles(sorted, q) {
  const pos = (sorted.length - 1) * q;
  const base = Math.floor(pos);
  const rest = pos - base;
  if (sorted[base + 1] !== undefined) return sorted[base] + rest * (sorted[base + 1] - sorted[base]);
  return sorted[base];
}
function iqrFilter(rows) {
  const nums = rows.map(r => r._priceOut).filter(Number.isFinite).sort((a,b)=>a-b);
  if (nums.length < 4) return { kept: rows, dropped: [] };
  const q1 = quantiles(nums, 0.25);
  const q3 = quantiles(nums, 0.75);
  const IQR = q3 - q1;
  const low = q1 - 1.5 * IQR;
  const high = q3 + 1.5 * IQR;
  const kept = [];
  const dropped = [];
  for (const r of rows) {
    if (!Number.isFinite(r._priceOut)) { kept.push(r); continue; }
    if (r._priceOut < low || r._priceOut > high) dropped.push(r);
    else kept.push(r);
  }
  return { kept, dropped, q1, q3, low, high };
}
function stats(nums) {
  if (!nums.length) return null;
  const min = Math.min(...nums), max = Math.max(...nums);
  const avg = nums.reduce((a,b)=>a+b,0)/nums.length;
  return { min, max, avg };
}
function tokenize(str, pattern) {
  return str.toLowerCase().replace(pattern, ' ').trim().split(/\s+/).filter(Boolean);
}
function similarity(a, b, pattern) {
  const t1 = tokenize(a, pattern);
  const t2 = tokenize(b, pattern);
  const set1 = new Set(t1);
  const set2 = new Set(t2);
  let inter = 0;
  for (const t of set1) if (set2.has(t)) inter++;
  const union = new Set([...set1, ...set2]).size;
  return union ? inter / union : 0;
}
function simAll(a, b) { return similarity(a, b, /[^a-zа-я0-9]+/g); }
function simEN(a, b) { return similarity(a, b, /[^a-z0-9]+/g); }

// API call
async function shopping(q) {
  const url = "https://google.serper.dev/shopping";
  const { data } = await axios.post(
    url,
    { q, gl: GL, hl: HL },
    { headers: { "X-API-KEY": KEY, "Content-Type": "application/json" }, timeout: TIMEOUT }
  );
  const items = (data?.shopping || []).map((it, idx) => {
    const priceText    = it.price ?? it.priceText ?? it.extracted_price ?? it.extractedPrice ?? "";
    const shippingText = it.shipping ?? it.delivery ?? it.shipping_text ?? "";
    const totalText    = it.total_price ?? it.price_total ?? it.final_price ?? "";
    const price    = normalizeNumber(priceText);
    const shipping = normalizeNumber(shippingText);
    const total    = normalizeNumber(totalText);
    return {
      title: it.title || "",
      seller: it.source || it.vendor || "",
      url:   it.product_link || it.link || "",
      priceText, shippingText, totalText,
      price, shipping, total,
      currency: guessCurrency(priceText, totalText),
      position: idx + 1
    };
  }).filter(x => x.url);
  return items;
}

(async function main() {
  console.log(`Запрос (Shopping, 1 вызов): ${QUERY}\nТоп (как у Google): ${TOPN} | GL=${GL} | HL=${HL}`);
  let items = [];
  try {
    items = await shopping(QUERY);
  } catch (e) {
    console.error("Ошибка Shopping:", e?.response?.data || e?.message || e);
    process.exit(1);
  }
  if (!items.length) { console.log("\nНичего не найдено."); process.exit(0); }

  const our = normalizeNumber(OUR_PRICE);
  const low = our * (1 - DELTA);
  const high = our * (1 + DELTA);

  const top = items.slice(0, TOPN).map(r => {
    const p = effectivePrice(r);
    const all = simAll(r.title, QUERY);
    const en  = simEN(r.title, QUERY);
    const w   = WHITELIST.some(w => r.seller.toLowerCase().includes(w));
    return { ...r, _priceOut: p.value, _priceLabel: p.label, _simAll: all, _simEN: en, _isWhite: w };
  });

  // фильтр выбросов
  let kept = top, dropped = [];
  if (CLEAN) {
    const f = iqrFilter(top);
    kept = f.kept;
    dropped = f.dropped;
  }

  const within = kept.filter(r => Number.isFinite(r._priceOut) && r._priceOut >= low && r._priceOut <= high);
  const numeric = within.filter(r => Number.isFinite(r._priceOut));
  const s = stats(numeric.map(x => x._priceOut));
  let avgLine = "";
  if (s) {
    const deltaAvg = ((s.avg - our) / our) * 100;
    avgLine = `Средняя по фильтру: ${s.avg.toFixed(2)} ${numeric[0].currency || ""} (дельта: ${deltaAvg.toFixed(2)}%)`;
  }

  const comparator = (a, b) => {
    if (a._isWhite !== b._isWhite) return a._isWhite ? -1 : 1;
    if (b._simEN !== a._simEN) return b._simEN - a._simEN;
    if (b._simAll !== a._simAll) return b._simAll - a._simAll;
    const diffA = Math.abs(a._priceOut - our);
    const diffB = Math.abs(b._priceOut - our);
    return diffA - diffB;
  };

  const candidates = within.filter(r => r._simAll >= MIN_SIM).sort(comparator);
  const golden = candidates[0];
  const allCandidates = kept.filter(r => r._simAll >= MIN_SIM).sort(comparator);
  const fallback = allCandidates[0];

  let line = "";
  if (golden) {
    line = `ИТОГОВОЕ ПРЕДЛОЖЕНИЕ (golden middle): • ${golden.title} Магазин: ${golden.seller}${golden._isWhite ? " (whitelist)" : ""} Цена: ${golden._priceOut} ${golden.currency || ""} Совпадение: ${(golden._simAll*100).toFixed(0)}% Позиция в Google: ${golden.position} Ссылка: ${golden.url}`;
  } else {
    line = "ИТОГОВОЕ ПРЕДЛОЖЕНИЕ (golden middle): не найдено";
  }

  let fb = "";
  if (fallback && fallback !== golden) {
    fb = `Фолбэк: кандидаты без учёта ценового фильтра: ${fallback.position}. ${fallback.title} Магазин: ${fallback.seller} Совпадение: ${(fallback._simAll*100).toFixed(0)}% Цена: ${fallback._priceOut ? `${fallback._priceOut} ${fallback.currency || ""}` : fallback.priceText || "—"}`;
  }

  const out = [avgLine, line, fb].filter(Boolean).join(" ");
  console.log(out);

  if (CLEAN && dropped.length) {
    const droppedInfo = dropped.map(r => `${r.seller}: ${Number.isFinite(r._priceOut) ? r._priceOut : r.priceText || "—"}`).join(", ");
    console.log(` Отброшено как выбросы: ${droppedInfo}`);
  }
})();
