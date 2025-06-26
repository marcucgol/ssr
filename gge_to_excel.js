// Script to convert .gge files to Excel and aggregate LSR_Cur info
// Usage: node gge_to_excel.js <inputDir> <outputDir>

const fs = require('fs');
const path = require('path');
const xml2js = require('xml2js');
const XLSX = require('xlsx');

(async () => {
  try {
    // Directories from command line or defaults
    const inputDir = process.argv[2] || './input';
    const outputDir = process.argv[3] || './output';
    const summaryFile = 'LSR_Cur_All.xlsx';

    // Ensure output directory exists
    fs.mkdirSync(outputDir, { recursive: true });

    function findNode(obj, name) {
      if (!obj || typeof obj !== 'object') return null;
      if (obj[name] != null) return obj[name];
      for (const k of Object.keys(obj)) {
        const res = findNode(obj[k], name);
        if (res != null) return res;
      }
      return null;
    }

    function createVertical(obj) {
      const rows = [['Field', 'Value']];
      for (const [k, v] of Object.entries(obj)) rows.push([k, String(v)]);
      const ws = XLSX.utils.aoa_to_sheet(rows);
      const c0 = rows.map(r => r[0].length), c1 = rows.map(r => r[1].length);
      ws['!cols'] = [{ wch: Math.max(...c0) + 2 }, { wch: Math.max(...c1) + 2 }];
      ws['!rows'] = rows.map((r, i) => i === 0 ? { hpt: 20 } : {
        hpt: Math.max(Math.ceil(r[1].length / ws['!cols'][1].wch) * 15, 15)
      });
      return ws;
    }

    function computeItog(flatSummary, prefix) {
      const F = flatSummary;
      const num = key => parseFloat(F[`${key}_${prefix}`]) || 0;
      const M = num('Materials_Total');
      const K = num('Totals_Items');
      const SNB = M - K;
      const P = num('Transport');
      const FT = num('Salary');
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

    const lsrColumns = [
      'FileNum', 'FileName', 'ObjectNum', 'ObjectName',
      'RegionCode', 'RegionName', 'EstNum', 'EstName',
      'EstType', 'IndexType', 'Reason', 'CurYear',
      'CurMonth', 'CurQuarter',
      'Материалы', 'КАЦ', 'СНБ', 'Перевозка', 'ФОТ', 'ЭММ',
      'Прямые затраты', 'НР', 'СР', 'Косвенные затраты',
      'Строительные работы', 'Монтажные работы', 'Оборудование',
      'Смета Total', 'Итого по смете'
    ];

    const allLsrRows = [];

    const files = fs.readdirSync(inputDir).filter(f => f.endsWith('.gge'));

    for (const fileName of files) {
      const fullPath = path.join(inputDir, fileName);
      const baseName = path.basename(fileName, '.gge');
      const xml = fs.readFileSync(fullPath, 'utf8');
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
      if (!objNode) throw new Error(`[${fileName}] Узел <Object> не найден`);
      Object.assign(header, {
        ObjectNum: objNode.Num || '',
        ObjectName: objNode.Name || '',
        RegionCode: objNode.Region?.Code || '',
        RegionName: objNode.Region?.Name || ''
      });
      const est = objNode.Estimate;
      if (!est) throw new Error(`[${fileName}] Узел <Estimate> не найден`);
      Object.assign(header, {
        EstNum: est.Num || '',
        EstName: est.Name || '',
        EstType: est.EstimateType || '',
        IndexType: est.IndexType || '',
        Reason: est.Reason || ''
      });

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
      if (!ep) throw new Error(`[${fileName}] Узел <EstimatePrice> не найден`);

      const summary = ep.Summary;
      if (!summary) throw new Error(`[${fileName}] Узел <Summary> не найден`);
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
      flatSummary.Totals_Current_Items = sumCur.toFixed(2);

      const itogC = computeItog(flatSummary, 'PriceCurrent');
      [
        ['Building', 'Строительные работы'],
        ['Mounting', 'Монтажные работы'],
        ['Equipment', 'Оборудование'],
        ['Total', 'Смета Total']
      ].forEach(([blk, label]) => {
        const data = flatBlocks[blk] || {};
        const curKey = blk === 'Total' ? 'PriceCurrent' : 'Total_PriceCurrent';
        itogC[label] = (parseFloat(data[curKey]) || 0).toFixed(2);
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
      otherBlocks.forEach(([blk]) => {
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

      const outPath = path.join(outputDir, `${baseName}.xlsx`);
      XLSX.writeFile(wb, outPath);
      console.log(`✅ Сохранён: ${outPath}`);

      const lsrRow = {};
      lsrColumns.forEach(col => {
        if (header[col] != null) lsrRow[col] = header[col];
        else if (itogC[col] != null) lsrRow[col] = itogC[col];
        else lsrRow[col] = '';
      });
      allLsrRows.push(lsrRow);
    }

    const summaryWb = XLSX.utils.book_new();
    const wsSum = XLSX.utils.json_to_sheet(allLsrRows, { header: lsrColumns });
    wsSum['!cols'] = lsrColumns.map(h => ({ wch: Math.max(h.length, ...allLsrRows.map(r => String(r[h]).length)) + 2 }));
    summaryWb.SheetNames.push('LSR_Cur');
    summaryWb.Sheets['LSR_Cur'] = wsSum;
    XLSX.writeFile(summaryWb, summaryFile);
    console.log(`✅ Итоговый файл: ${summaryFile}`);
  } catch (err) {
    console.error('❌ Ошибка:', err.message);
  }
})();

