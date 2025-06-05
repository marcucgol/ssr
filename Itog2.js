'use strict';
const fs = require('fs');
const path = require('path');
const xml2js = require('xml2js');
const ExcelJS = require('exceljs');
const XLSX = require('xlsx');

/* ===================== Часть 1. Генерация Excel-файлов из .gge файлов ===================== */

/**
 * Рекурсивно ищет файлы с расширением .gge во всех папках внутри «Объекты»,
 * пропуская папки с названием OSR (результирующая папка).
 */
function getGGEFiles(dir) {
  let results = [];
  const list = fs.readdirSync(dir);
  list.forEach(file => {
    const filePath = path.join(dir, file);
    const stat = fs.statSync(filePath);
    if (stat.isDirectory()) {
      // Пропускаем папку OSR, чтобы не обрабатывать уже сгенерированные Excel-файлы
      if (file.toLowerCase() === 'osr') return;
      results = results.concat(getGGEFiles(filePath));
    } else if (stat.isFile() && file.toLowerCase().endsWith('.gge')) {
      results.push(filePath);
    }
  });
  return results;
}

// Функция для безопасного получения текста из XML-элемента
function safeGetText(value) {
  return (value !== undefined && value !== null && value !== '') ? value : '';
}

/**
 * Обрабатывает один .gge файл и создаёт из него Excel-файл.
 * Excel-файл сохраняется в единой папке OSR внутри корневой папки «Объекты».
 */
async function processFile(filePath, saveFolder) {
  const xmlData = fs.readFileSync(filePath, 'utf8');
  const parser = new xml2js.Parser({ explicitArray: false });
  let parsed;
  try {
    parsed = await parser.parseStringPromise(xmlData);
  } catch (err) {
    console.error(`Ошибка парсинга XML в файле ${filePath}:`, err);
    return;
  }

  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet('Сметные расчёты');

  // Извлекаем Object Name для колонки Name2
  const objectName2 = safeGetText(parsed?.Construction?.Object?.Name);

  // Верхняя часть (заголовок)
  const constrNum  = safeGetText(parsed?.Construction?.Num);
  const constrName = safeGetText(parsed?.Construction?.Name);
  worksheet.addRow([`Construction Num: ${constrNum}`, `Construction Name: ${constrName}`]);

  // Определяем исходную папку (Type) по пути .gge файла
  const objectsFolder = path.join(__dirname, 'Объекты');
  const relativePath  = path.relative(objectsFolder, filePath);
  const segments      = relativePath.split(path.sep);
  const originalType  = segments.length > 0 ? segments[0] : '';
  worksheet.addRow([`Type: ${originalType}`]);

  // Данные Object
  const obj      = parsed?.Construction?.Object || {};
  const objNum   = safeGetText(obj.Num);
  const objName  = safeGetText(obj.Name);
  const region   = safeGetText(obj.Region);
  const reason   = safeGetText(obj.Reason);
  const created  = safeGetText(obj.Created);
  const approved = safeGetText(obj.Approved);
  // Year и Quarter
  const yearStr     = safeGetText(obj.PriceLevel?.Year);
  const priceYear   = yearStr ? parseInt(yearStr, 10) : null;
  let quarterStr    = safeGetText(obj.PriceLevel?.Quarter);
  if (!quarterStr) {
    const monthStr = safeGetText(obj.PriceLevel?.Month);
    if (monthStr) {
      const m = parseInt(monthStr, 10);
      if (!isNaN(m)) quarterStr = Math.ceil(m / 3).toString();
    }
  }
  const priceQuarter = quarterStr ? parseInt(quarterStr, 10) : null;

  worksheet.addRow([
    `Object Num: ${objNum}`,
    `Object Name: ${objName}`,
    `Region: ${region}`,
    `Reason: ${reason}`,
    `Created: ${created}`,
    `Approved: ${approved}`,
    `PriceLevel: ${priceYear}/${priceQuarter}`
  ]);

  // Пустая строка, затем заголовок таблицы с дополнительным столбцом Name2
  worksheet.addRow([]);
  worksheet.addRow([
    'Name2',
    '№ п/п',
    'Обоснование',
    'Наименование локальных сметных расчётов',
    'строительных работ',
    'монтажных работ',
    'оборудования',
    'прочих затрат',
    'всего',
    'Year',
    'Quarter'
  ]);

  const localEstimates = parsed?.Construction?.Object?.LocalEstimate;
  if (!localEstimates) {
    console.warn(`Не найдены элементы <LocalEstimate> в файле ${filePath}`);
    return;
  }
  const estimatesArray = Array.isArray(localEstimates) ? localEstimates : [localEstimates];

  estimatesArray.forEach(est => {
    const num       = safeGetText(est.Num);
    const estReason = safeGetText(est.Reason);
    const estName   = safeGetText(est.Name);
    const building  = est.Building   ? parseFloat(est.Building)   : 0;
    const mounting  = est.Mounting   ? parseFloat(est.Mounting)   : 0;
    const equipment = est.Equipment  ? parseFloat(est.Equipment)  : 0;
    const other     = est.Other      ? parseFloat(est.Other)      : 0;
    const total     = est.Total      ? parseFloat(est.Total)      : 0;

    worksheet.addRow([
      objectName2,
      num,
      estReason,
      estName,
      building,
      mounting,
      equipment,
      other,
      total,
      priceYear,
      priceQuarter
    ]);
  });

  // Итоговая строка — тоже с Name2 на первом месте
  worksheet.addRow([
    objectName2,
    'ИТОГО:',
    '', '', '', '', '', '',
    priceYear,
    priceQuarter
  ]);

  const baseName     = path.basename(filePath, path.extname(filePath));
  const excelPath    = path.join(saveFolder, baseName + '.xlsx');
  try {
    await workbook.xlsx.writeFile(excelPath);
    console.log(`Создан Excel-файл: ${excelPath}`);
  } catch (err) {
    console.error(`Ошибка записи Excel-файла для ${filePath}:`, err);
  }
}

/**
 * Генерирует Excel-файлы из .gge в папке OSR внутри «Объекты».
 */
async function generateExcelFiles() {
  const sourceFolder = path.join(__dirname, 'Объекты');
  if (!fs.existsSync(sourceFolder)) {
    console.error('Папка "Объекты" не найдена:', sourceFolder);
    process.exit(1);
  }

  const saveFolder = path.join(sourceFolder, 'OSR');
  if (!fs.existsSync(saveFolder)) {
    fs.mkdirSync(saveFolder);
    console.log('Создана папка для сохранения результатов:', saveFolder);
  }

  const ggeFiles = getGGEFiles(sourceFolder);
  if (ggeFiles.length === 0) {
    console.log('Файлы .gge не найдены в папке "Объекты":', sourceFolder);
    return;
  }
  for (const file of ggeFiles) {
    await processFile(file, saveFolder);
  }
}

/* ===================== Часть 2. Обработка сгенерированных Excel-файлов ===================== */

function extractType(sheet) {
  for (let cell in sheet) {
    if (cell[0] === '!') continue;
    const cellObj = sheet[cell];
    if (cellObj && cellObj.v && typeof cellObj.v === 'string') {
      const text = cellObj.v;
      if (text.toLowerCase().startsWith("type:")) {
        return text.split(':')[1].trim();
      }
    }
  }
  return "";
}

function extractConstructionName(sheet) {
  for (let cell in sheet) {
    if (cell[0] === '!') continue;
    const cellObj = sheet[cell];
    if (cellObj && cellObj.v && typeof cellObj.v === 'string') {
      const text = cellObj.v;
      if (text.toLowerCase().includes("construction name:")) {
        let parts = text.split(':');
        if (parts.length > 1) {
          let after = parts.slice(1).join(':').trim();
          let quoteMatch = /«([^»]+)»/.exec(after);
          return quoteMatch ? quoteMatch[1] : after;
        }
      }
    }
  }
  return "";
}

function extractObjectName(sheet) {
  for (let cell in sheet) {
    if (cell[0] === '!') continue;
    const v = sheet[cell].v;
    if (typeof v === 'string' && v.toLowerCase().startsWith('object name:')) {
      return v.split(':')[1].trim();
    }
  }
  return "";
}

function loadNLSRMapping() {
  if (!fs.existsSync('NLSR.xlsx')) return [];
  const wb = XLSX.readFile('NLSR.xlsx');
  const sh = wb.Sheets[wb.SheetNames[0]];
  const data = XLSX.utils.sheet_to_json(sh, { header: 1, defval: "" });
  const header = data[0].map(h => h.toString().toLowerCase());
  const ni = header.indexOf('name'), ki = header.indexOf('keyword');
  const mapping = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row[ni] && row[ki]) {
      mapping.push({ name: row[ni], keyword: row[ki] });
    }
  }
  return mapping;
}

function loadTEPMapping() {
  if (!fs.existsSync('TEP.xlsx')) return {};
  const wb = XLSX.readFile('TEP.xlsx');
  const sh = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(sh, { defval: "" });
  const map = {};
  rows.forEach(r => {
    let tepVal = r["Tep"];
    if (typeof tepVal === "string") tepVal = parseFloat(tepVal.replace(",", "."));
    else tepVal = parseFloat(tepVal);
    const key = `${r["Type"]}|${r["Name"]}|${r["Num 1"]}|${r["Num 2"]}`;
    map[key] = tepVal;
  });
  return map;
}

function findHeaderRow(arr) {
  const need = ["Обоснование", "Наименование локальных сметных расчётов"];
  for (let i = 0; i < arr.length; i++) {
    if (Array.isArray(arr[i])) {
      const line = arr[i].join(" ").toLowerCase();
      if (need.every(k => line.includes(k.toLowerCase()))) return i;
    }
  }
  return 0;
}

function getTotalValue(v) {
  if (!v) return 0;
  return parseFloat(String(v).replace(",", ".")) || 0;
}

function processExcelFile(filePath, nlsrMapping) {
  const wb       = XLSX.readFile(filePath);
  const sheet    = wb.Sheets[wb.SheetNames[0]];
  const name1    = extractConstructionName(sheet);
  const type     = extractType(sheet);
  const name2    = extractObjectName(sheet);

  const arr      = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });
  const hr       = findHeaderRow(arr);
  const headers  = arr[hr];
  const data     = XLSX.utils.sheet_to_json(sheet, { header: headers, range: hr + 1, defval: "" });

  const out = [];
  data.forEach(r => {
    if ((r["Обоснование"] || "").toString().toLowerCase().includes("итого")) return;
    const ob  = r["Обоснование"] || "";
    const desc= r["Наименование локальных сметных расчётов"]||"";
    const m   = /ЛС(?:Р)?[\s-]*(\d{2})-(\d{2})-/i.exec(ob);
    const num1= m ? m[1] : "";
    const num2= m ? m[2] : "";

    let grp = "", kw = "";
    for (const map of nlsrMapping) {
      if (desc.toLowerCase().includes(map.keyword.toLowerCase())) {
        grp = map.name;
        kw  = map.keyword;
        break;
      }
    }

    out.push({
      "№ п/п": r["№ п/п"] || "",
      "Обоснование": ob,
      "Наименование локальных сметных расчётов": desc,
      "строительных работ": parseFloat(r["строительных работ"])||0,
      "монтажных работ": parseFloat(r["монтажных работ"])||0,
      "оборудования": parseFloat(r["оборудования"])||0,
      "прочих затрат": parseFloat(r["прочих затрат"])||0,
      "всего": parseFloat(r["всего"])||0,
      "Year": parseInt(r["Year"], 10)||null,
      "Quarter": parseInt(r["Quarter"],10)||null,
      "Name": name1,
      "Name2": name2,
      "Type": type,
      "НЛСР группа": grp,
      "Keyword": kw,
      "Num 1": num1,
      "Num 2": num2,
      "name_file": path.basename(filePath)
    });
  });
  return out;
}

function groupRows(rows) {
  const g = {};
  rows.forEach(r => {
    const key = `${r["Type"]}|${r["Name"]}|${r["Name2"]}|${r["Num 1"]}|${r["Num 2"]}|${r["НЛСР группа"]}|${r["Year"]}|${r["Quarter"]}`;
    if (!g[key]) {
      g[key] = {
        "Type": r["Type"],
        "Name": r["Name"],
        "Name2": r["Name2"],
        "Num 1": r["Num 1"],
        "Num 2": r["Num 2"],
        "НЛСР группа": r["НЛСР группа"],
        "Year": r["Year"],
        "Quarter": r["Quarter"],
        "всего": 0
      };
    }
    g[key]["всего"] += r["всего"];
  });
  return Object.values(g);
}

function prepareDetailedData(rows) {
  return rows.map(r => ({
    "Type": r["Type"],
    "Name": r["Name"],
    "Name2": r["Name2"],
    "Num 1": r["Num 1"],
    "Num 2": r["Num 2"],
    "Year": r["Year"],
    "Quarter": r["Quarter"],
    "всего": r["всего"]
  }));
}

function applyTEPMapping(dataArray, tepMapping) {
  dataArray.forEach(row => {
    const total     = getTotalValue(row["всего"]);
    const baseKey   = `${row["Type"]}|${row["Name"]}|${row["Num 1"]}|${row["Num 2"]}`;
    const with2Key  = `${row["Type"]}|${row["Name"]}|${row["Name2"]}|${row["Num 1"]}|${row["Num 2"]}`;
    let tep = tepMapping[with2Key];
    if (tep === undefined) tep = tepMapping[baseKey];
    if (tep !== undefined && tep !== 0) {
      row["TEP"]     = tep;
      row["Kvadrat"] = total / tep;
    } else {
      row["TEP"]     = "";
      row["Kvadrat"] = "";
    }
  });
}

/* Главная функция */
async function main() {
  console.log('--- Запуск генерации Excel-файлов из .gge ---');
  await generateExcelFiles();

  console.log('--- Запуск обработки и объединения в combined_output.xlsx ---');
  const nlsrMap = loadNLSRMapping();
  const tepMap  = loadTEPMapping();

  const osrDir = path.join(__dirname, 'Объекты', 'OSR');
  if (!fs.existsSync(osrDir)) {
    console.error('Папка OSR не найдена:', osrDir);
    return;
  }
  const files = fs.readdirSync(osrDir)
    .filter(f => f.endsWith('.xlsx') && f !== 'NLSR.xlsx' && f !== 'TEP.xlsx')
    .map(f => path.join(osrDir, f));

  let allRows = [];
  files.forEach(f => {
    try {
      const rows = processExcelFile(f, nlsrMap);
      allRows = allRows.concat(rows);
      console.log('Обработан файл:', f);
    } catch (e) {
      console.error('Ошибка при обработке файла', f, e);
    }
  });

  const fullHeaders = [
    "№ п/п","Обоснование","Наименование локальных сметных расчётов","строительных работ",
    "монтажных работ","оборудования","прочих затрат","всего",
    "Year","Quarter","Name","Name2","Type","НЛСР группа","Keyword","Num 1","Num 2","name_file"
  ];
  const ws1 = XLSX.utils.json_to_sheet(allRows, { header: fullHeaders });

  const grouped   = groupRows(allRows);
  applyTEPMapping(grouped, tepMap);
  const grpHeaders= ["Type","Name","Name2","Num 1","Num 2","НЛСР группа","Year","Quarter","всего","TEP","Kvadrat"];
  const ws2       = XLSX.utils.json_to_sheet(grouped, { header: grpHeaders });

  const detailed  = prepareDetailedData(allRows);
  applyTEPMapping(detailed, tepMap);
  const detHeaders= ["Type","Name","Name2","Num 1","Num 2","Year","Quarter","всего","TEP","Kvadrat"];
  const ws3       = XLSX.utils.json_to_sheet(detailed, { header: detHeaders });

  const outWb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(outWb, ws1, 'MainData');
  XLSX.utils.book_append_sheet(outWb, ws2, 'GroupedData');
  XLSX.utils.book_append_sheet(outWb, ws3, 'DetailedData');
  XLSX.writeFile(outWb, 'combined_output.xlsx');
  console.log('Создан combined_output.xlsx');
}

async function run() {
  try {
    await main();
  } catch (err) {
    console.error('Ошибка выполнения:', err);
  }
}

if (require.main === module) {
  run();
}

module.exports = { main };
