'use strict';
// Process multiple .gge files in a folder and create Excel outputs for each.
// Aggregates the LSR_Cur sheet of every workbook into one combined file in the
// project root.  Adapted from single-file example in the user prompt.

const fs = require('fs');
const path = require('path');
const xml2js = require('xml2js');
const XLSX = require('xlsx');

// Recursively find .gge files within a folder
function findGGEFiles(dir) {
  let files = [];
  fs.readdirSync(dir).forEach(f => {
    const p = path.join(dir, f);
    const stat = fs.statSync(p);
    if (stat.isDirectory()) files = files.concat(findGGEFiles(p));
    else if (stat.isFile() && f.toLowerCase().endsWith('.gge')) files.push(p);
  });
  return files;
}

// Helper for finding node by name
function findNode(obj, name) {
  if (!obj || typeof obj !== 'object') return null;
  if (obj[name] != null) return obj[name];
  for (const k of Object.keys(obj)) {
    const res = findNode(obj[k], name);
    if (res != null) return res;
  }
  return null;
}

// Create vertical sheet from object key/value pairs
function createVertical(obj) {
  const rows = [['Field', 'Value']];
  for (const [k, v] of Object.entries(obj)) rows.push([k, String(v)]);
  const ws = XLSX.utils.aoa_to_sheet(rows);
  const c0 = rows.map(r => r[0].length), c1 = rows.map(r => r[1].length);
  ws['!cols'] = [{ wch: Math.max(...c0) + 2 }, { wch: Math.max(...c1) + 2 }];
  ws['!rows'] = rows.map((r, i) => i === 0 ? { hpt: 20 } : { hpt: Math.max(Math.ceil(r[1].length / ws['!cols'][1].wch) * 15, 15) });
  return ws;
}

// Process a single .gge file and return workbook + LSR row
async function processFile(filePath) {
  const xml = fs.readFileSync(filePath, 'utf8');
  const parsed = await xml2js.parseStringPromise(xml, {
    explicitArray: false,
    mergeAttrs: true,
    trim: true
  });

  const root = parsed.Construction || parsed;
  const header = {
    FileNum: root.Num || '',
    FileName: root.Name || ''
  };
  const objNode = findNode(root, 'Object');
  if (!objNode) throw new Error('Узел <Object> не найден');
  Object.assign(header, {
    ObjectNum: objNode.Num || '',
    ObjectName: objNode.Name || '',
    RegionCode: objNode.Region?.Code || '',
    RegionName: objNode.Region?.Name || '',
    SubRegion: objNode.SubRegion?.Name || ''
  });
  const est = objNode.Estimate;
  if (!est) throw new Error('Узел <Estimate> не найден');
  Object.assign(header, {
    EstNum: est.Num || '',
    EstName: est.Name || '',
    EstType: est.EstimateType || '',
    IndexType: est.IndexType || '',
    EstDateYear: est.Date?.Year || '',
    EstDateMonth: est.Date?.Month || '',
    EstDateDay: est.Date?.Day || '',
    EstDateQuarter: (() => {
      const m = Number(est.Date?.Month);
      return m ? Math.floor((m - 1) / 3) + 1 : '';
    })(),
    Reason: est.Reason || ''
  });
  if (est.PriceLevelBase) {
    header.BaseYear = est.PriceLevelBase.Year || '';
    if (est.PriceLevelBase.Month) {
      header.BaseMonth = est.PriceLevelBase.Month;
      const bm = Number(est.PriceLevelBase.Month);
      header.BaseQuarter = bm ? Math.floor((bm - 1) / 3) + 1 : '';
    } else if (est.PriceLevelBase.Quarter) {
      header.BaseQuarter = est.PriceLevelBase.Quarter;
    }
  }
  if (est.PriceLevelCur) {
    header.CurYear = est.PriceLevelCur.Year || '';
    if (est.PriceLevelCur.Month) {
      header.CurMonth = est.PriceLevelCur.Month;
      const cm = Number(est.PriceLevelCur.Month);
      header.CurQuarter = cm ? Math.floor((cm - 1) / 3) + 1 : '';
    } else if (est.PriceLevelCur.Quarter) {
      header.CurQuarter = est.PriceLevelCur.Quarter;
    }
  }

  const ep = est.EstimatePrice;
  if (!ep) throw new Error('Узел <EstimatePrice> не найден');
  const summary = ep.Summary;
  if (!summary) throw new Error('Узел <Summary> не найден');
  const flatSummary = {};
  (function flatten(obj, prefix = '') {
    for (const [k, v] of Object.entries(obj)) {
      const key = prefix ? `${prefix}_${k}` : k;
      if (v != null && typeof v === 'object') flatten(v, key);
      else flatSummary[key] = String(v);
    }
  })(summary);

  const otherBlocks = Object.entries(ep).filter(([n]) => n !== 'Summary');
  const flatBlocks = {};
  otherBlocks.forEach(([name, block]) => {
    const out = {};
    (function flatten(obj, pre = '') {
      for (const [k, v] of Object.entries(obj)) {
        const key = pre ? `${pre}_${k}` : k;
        if (v != null && typeof v === 'object') flatten(v, key);
        else out[key] = String(v);
      }
    })(block);
    flatBlocks[name] = out;
  });

  const sections = est.Sections?.Section;
  const secArr = Array.isArray(sections) ? sections : [sections];
  const itemsRows = [];
  const allItemKeys = new Set();
  secArr.forEach(sec => {
    const secCode = sec.Code || '';
    const secName = sec.Name || '';
    let items = sec.Items?.Item;
    if (!items) return;
    items = Array.isArray(items) ? items : [items];
    items = items.filter(it => it.Material?.Code?.startsWith('ТЦ_'));
    items.forEach(item => {
      const flat = { SectionCode: secCode, SectionName: secName };
      (function f(obj, pre = '') {
        for (const [k, v] of Object.entries(obj)) {
          const key = pre ? `${pre}_${k}` : k;
          if (v != null && typeof v === 'object') f(v, key);
          else {
            flat[key] = String(v);
            allItemKeys.add(key);
          }
        }
      })(item);
      itemsRows.push(flat);
    });
  });

  const sumCur = itemsRows.reduce((s, r) => s + (parseFloat(r.Totals_Current) || 0), 0);
  const sumBase = itemsRows.reduce((s, r) => s + (parseFloat(r.Totals_Base) || 0), 0);
  flatSummary.Totals_Current_Items = sumCur.toFixed(2);
  flatSummary.Totals_Base_Items = sumBase.toFixed(2);

  function computeItog(prefix) {
    const F = flatSummary;
    const num = key => parseFloat(F[`${key}_${prefix}`]) || 0;
    const alt = (...keys) => {
      for (const k of keys) {
        const v = parseFloat(F[k]);
        if (!isNaN(v)) return v;
      }
      return 0;
    };
    const isCur = prefix === 'PriceCurrent';
    const M = alt(`Materials_Total_${prefix}`,
                  `Materials_${prefix}_Total`,
                  `${prefix}_Materials_Total`,
                  'Materials_Total');
    const K = alt(isCur ? 'Totals_Current_Items' : 'Totals_Base_Items',
                  `Totals_Items_${prefix}`,
                  'Totals_Items');
    const SNB = M - K;
    const P = num('Transport');
    const FT = num('Salary') - num('MachinistSalaryExtra');
    const E = num('MachinesTotal') - num('MachinistSalary');
    const PR = K + SNB + P + FT + E;
    const N = num('Overhead');
    const S = num('Profit');
    const KZ = N + S;
    return {
      'Материалы': M.toFixed(2),
      'КАЦ': K.toFixed(2),
      'СНБ': SNB.toFixed(2),
      'Перевозка': P.toFixed(2),
      'ФОТ': FT.toFixed(2),
      'ЭММ': E.toFixed(2),
      'Прямые затраты': PR.toFixed(2),
      'НР': N.toFixed(2),
      'СР': S.toFixed(2),
      'Косвенные затраты': KZ.toFixed(2)
    };
  }
  const itogC = computeItog('PriceCurrent');
  const itogB = computeItog('PriceBase');

  [
    ['Building',    'Строительные работы', 'Total_PriceCurrent'],
    ['Mounting',    'Монтажные работы',   'Total_PriceCurrent'],
    ['Equipment',   'Оборудование',       'Total_PriceCurrent'],
    ['OtherTotal',  'Прочие',             'PriceCurrent'],
    ['Total',       'Смета Total',        'PriceCurrent']
  ].forEach(([blk, label, prop]) => {
    const data = flatBlocks[blk] || {};
    itogC[label] = (parseFloat(data[prop]) || 0).toFixed(2);
  });

  itogC['Итого по смете'] = (
    parseFloat(itogC['Прямые затраты']) +
    parseFloat(itogC['Косвенные затраты']) +
    parseFloat(itogC['Оборудование'])
  ).toFixed(2);

  const wb = XLSX.utils.book_new();
  wb.SheetNames.push('Header');
  wb.Sheets['Header'] = createVertical(header);

  wb.SheetNames.push('EstimatePrice');
  wb.Sheets['EstimatePrice'] = createVertical(flatSummary);

  Object.keys(flatBlocks).forEach(blk => {
    wb.SheetNames.push(blk);
    wb.Sheets[blk] = createVertical(flatBlocks[blk]);
  });

  const headersItems = ['SectionCode', 'SectionName', ...Array.from(allItemKeys)];
  const wsItems = XLSX.utils.json_to_sheet(itemsRows, { header: headersItems });
  wsItems['!cols'] = headersItems.map(h => ({ wch: Math.max(h.length, ...itemsRows.map(r => r[h]?.length || 0)) + 2 }));
  wb.SheetNames.push('Items');
  wb.Sheets['Items'] = wsItems;

  wb.SheetNames.push('Itog_Current');
  wb.Sheets['Itog_Current'] = createVertical(itogC);

  wb.SheetNames.push('Itog_Base');
  wb.Sheets['Itog_Base'] = createVertical(itogB);

  const lsrColumns = [
    'FileNum', 'FileName', 'ObjectNum', 'ObjectName', 'RegionCode', 'RegionName',
    'EstNum', 'EstName', 'EstType', 'IndexType', 'Reason', 'CurYear', 'CurMonth', 'CurQuarter',
    'Материалы', 'КАЦ', 'СНБ', 'Перевозка', 'ФОТ', 'ЭММ', 'Прямые затраты', 'НР', 'СР',
    'Косвенные затраты', 'Строительные работы', 'Монтажные работы', 'Оборудование', 'Прочие',
    'Смета Total', 'Итого по смете'
  ];
  const lsrRow = {};
  lsrColumns.forEach(key => {
    if (header[key] != null) lsrRow[key] = header[key];
    else if (itogC[key] != null) lsrRow[key] = itogC[key];
    else lsrRow[key] = '';
  });
  const wsLSR = XLSX.utils.json_to_sheet([lsrRow], { header: lsrColumns });
  wsLSR['!cols'] = lsrColumns.map(h => ({ wch: Math.max(h.length, String(lsrRow[h]).length) + 2 }));
  wb.SheetNames.push('LSR_Cur');
  wb.Sheets['LSR_Cur'] = wsLSR;

  return { workbook: wb, lsrRow };
}

async function main() {
  const inputDir = process.argv[2] || path.join(__dirname, 'input');
  const outDir = process.argv[3] || path.join(__dirname, 'output');
  const combinedPath = path.join(__dirname, 'LSR_combined.xlsx');

  if (!fs.existsSync(outDir)) fs.mkdirSync(outDir, { recursive: true });

  const files = findGGEFiles(inputDir);
  if (files.length === 0) {
    console.error('Файлы .gge не найдены в', inputDir);
    return;
  }

  const combinedRows = [];

  for (const f of files) {
    try {
      const { workbook, lsrRow } = await processFile(f);
      const outName = path.basename(f, path.extname(f)) + '.xlsx';
      const outPath = path.join(outDir, outName);
      XLSX.writeFile(workbook, outPath);
      console.log('Создан файл', outPath);
      combinedRows.push(lsrRow);
    } catch (err) {
      console.error('Ошибка обработки', f, err.message);
    }
  }

  if (combinedRows.length > 0) {
    const columns = Object.keys(combinedRows[0]);
    const ws = XLSX.utils.json_to_sheet(combinedRows, { header: columns });
    ws['!cols'] = columns.map(h => ({ wch: Math.max(h.length, ...combinedRows.map(r => String(r[h]).length)) + 2 }));
    const wb = XLSX.utils.book_new();
    wb.SheetNames.push('LSR_Cur');
    wb.Sheets['LSR_Cur'] = ws;
    XLSX.writeFile(wb, combinedPath);
    console.log('Создан объединённый файл', combinedPath);
  }
}

if (require.main === module) {
  main().catch(e => {
    console.error('Fatal error', e);
  });
}
