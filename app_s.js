const express = require('express');
const path    = require('path');
const xlsx    = require('xlsx');

const app  = express();
const PORT = 2500;

// Отображаемые имена столбцов
const displayNames = {
  Type:         'Тип',
  Name:         'Название',
  Name2:        'Доп. название',
  'Num 1':      'Число 1',
  'Num 2':      'Число 2',
  'НЛСР группа':'Группа',
  Year:         'Год',
  Quarter:      'Квартал',
  всего:        'Всего',
  TEP:          'ТЭП',
  Kvadrat:      'Квадрат'
};

// Читаем Excel
const workbook  = xlsx.readFile(path.join(__dirname, 'combined_output.xlsx'));
const sheetName = 'GroupedData';
if (!workbook.Sheets[sheetName]) throw new Error(`Лист "${sheetName}" не найден`);
const sheet = workbook.Sheets[sheetName];
const rows  = xlsx.utils.sheet_to_json(sheet, { defval: '' });

// Все колонки и фильтрация пустых
const cols = [
  'Type','Name','Name2',
  'Num 1','Num 2','НЛСР группа',
  'Year','Quarter','всего','TEP','Kvadrat'
];
const data = rows.filter(r =>
  cols.every(c => (r[c] ?? '').toString().trim() !== '')
);

// Фильтруемые столбцы (без «всего», «TEП», «Kvadrat»)
const filterCols = cols.filter(c => !['всего','TEP','Kvadrat'].includes(c));
const keyMap = {};
filterCols.forEach(c => {
  const id = c.replace(/\W+/g,'_').toLowerCase();
  keyMap[id] = c;
});

// Статика для CSS
app.use('/static', express.static(path.join(__dirname, 'public')));

// API для данных
app.get('/data', (req, res) => res.json(data));

// Главная страница
app.get('/', (req, res) => {
  res.send(`<!DOCTYPE html>
<html lang="ru">
<head>
  <meta charset="UTF-8">
  <title>Фильтрация и анализ данных</title>
  <link rel="stylesheet" href="/static/style.css">
  <style>
    body { background: #f0f2f5; margin: 0; padding: 20px; font-family: Arial, sans-serif; }
    .container { background: #fff; padding: 20px; border-radius: 8px;
                 max-width: 1600px; margin: auto; box-shadow: 0 2px 8px rgba(0,0,0,0.1); }
    .header { display: flex; justify-content: space-between; align-items: center; margin-bottom: 15px; }
    .header h1 { margin: 0; }
    .stats { font-size: 1.1em; font-weight: bold; }
    .filters { display: flex; flex-wrap: wrap; gap: 10px; margin-bottom: 15px; }
    .multiselect { position: relative; flex: 1; min-width: 150px; }
    .selectBox { display: flex; justify-content: space-between; align-items: center;
                 border: 1px solid #ccc; border-radius: 4px; padding: 5px 10px;
                 background: #fff; cursor: pointer; }
    .selectBox span { white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
    .checkboxes { display: none; position: absolute; background: #fff; border: 1px solid #ccc;
                  border-radius: 4px; max-height: 200px; overflow-y: auto; width: 100%;
                  z-index: 10; padding: 5px; box-shadow: 0 2px 6px rgba(0,0,0,0.2); }
    .actions { display: flex; justify-content: space-between; margin-bottom: 5px; }
    .actions button {
      flex: 1; margin: 0 2px; padding: 4px; font-size: 0.9em;
      border: 1px solid #888; background: #eee; border-radius: 3px; cursor: pointer;
    }
    .checkboxes label { display: block; margin-bottom: 3px; }
    .table-wrapper { overflow-x: auto; }
    table { width: 100%; min-width: 1400px; border-collapse: collapse; }
    th, td { border: 1px solid #ddd; padding: 8px; text-align: left; }
    th { background: #4CAF50; color: #fff; cursor: pointer; user-select: none; }
    th.sort-asc::after { content: ' ▲'; }
    th.sort-desc::after { content: ' ▼'; }
    tr:nth-child(even) { background: #f9f9f9; }
    .arrow { font-weight: bold; margin-left: 4px; }
  </style>
</head>
<body>
  <div class="container">
    <div class="header">
      <h1>Фильтрация и анализ данных</h1>
      <div class="stats">
        Среднее ${displayNames.Kvadrat}: <span id="avgKvadrat">0.00</span>
      </div>
    </div>
    <div class="filters">
      ${Object.entries(keyMap).map(([id,key]) => `
        <div class="multiselect" id="multi-${id}">
          <div class="selectBox" onclick="toggleCheckboxes('${id}')">
            <span id="selected-${id}">— ${displayNames[key]} —</span>
          </div>
          <div class="checkboxes" id="checkboxes-${id}">
            <div class="actions">
              <button type="button" onclick="selectAll('${id}')">Выбрать все</button>
              <button type="button" onclick="clearAll('${id}')">Сбросить</button>
            </div>
          </div>
        </div>
      `).join('')}
    </div>
    <div class="table-wrapper">
      <table id="table">
        <thead>
          <tr>${cols.map(c => `<th data-col="${c}">${displayNames[c]||c}</th>`).join('')}</tr>
        </thead>
        <tbody></tbody>
      </table>
    </div>
  </div>

  <script>
    const cols = ${JSON.stringify(cols)};
    const keyMap = ${JSON.stringify(keyMap)};
    const displayNames = ${JSON.stringify(displayNames)};
    let allData = [], sortCol = null, sortDir = 1;

    document.addEventListener('DOMContentLoaded', () => {
      fetch('/data').then(r=>r.json()).then(json=>{
        allData = json;
        Object.entries(keyMap).forEach(([id, key]) => {
          const container = document.getElementById('checkboxes-' + id);
          const values = Array.from(new Set(allData.map(r=>r[key]))).sort();
          container.innerHTML += values.map(v =>
            '<label><input type="checkbox" value="' + v + '" onchange="renderTable()"> ' + v + '</label>'
          ).join('');
        });
        document.querySelectorAll('th[data-col]').forEach(th => {
          th.addEventListener('click', () => {
            const col = th.dataset.col;
            sortDir = (sortCol === col ? -sortDir : 1);
            sortCol = col;
            document.querySelectorAll('th').forEach(h => h.classList.remove('sort-asc','sort-desc'));
            th.classList.add(sortDir === 1 ? 'sort-asc' : 'sort-desc');
            renderTable();
          });
        });
        renderTable();
      });
      document.addEventListener('click', e => {
        Object.keys(keyMap).forEach(id => {
          const ms = document.getElementById('multi-' + id);
          if (!ms.contains(e.target)) {
            document.getElementById('checkboxes-' + id).style.display = 'none';
          }
        });
      });
    });

    function toggleCheckboxes(id) {
      const box = document.getElementById('checkboxes-' + id);
      box.style.display = box.style.display === 'block' ? 'none' : 'block';
    }
    function selectAll(id) {
      document.querySelectorAll('#checkboxes-' + id + ' input').forEach(i => i.checked = true);
      renderTable();
    }
    function clearAll(id) {
      document.querySelectorAll('#checkboxes-' + id + ' input').forEach(i => i.checked = false);
      renderTable();
    }

    function renderTable() {
      const selections = {};
      Object.entries(keyMap).forEach(([id, key]) => {
        const checked = Array.from(
          document.querySelectorAll('#checkboxes-' + id + ' input:checked')
        ).map(i => i.value);
        selections[key] = checked;
        const span = document.getElementById('selected-' + id);
        span.textContent = checked.length ? checked.join(', ') : '— ' + displayNames[key] + ' —';
      });

      let filtered = allData.filter(row =>
        Object.entries(selections).every(([key, vals]) =>
          !vals.length || vals.includes(String(row[key]))
        )
      );

      const kv = 'Kvadrat';
      const sum = filtered.reduce((s,r)=>s+Number(r[kv]||0),0);
      const avg = filtered.length ? sum/filtered.length : 0;
      document.getElementById('avgKvadrat').textContent =
        avg.toLocaleString('ru-RU',{minimumFractionDigits:2,maximumFractionDigits:2});

      if (sortCol) {
        filtered.sort((a,b) => {
          const va=a[sortCol], vb=b[sortCol];
          const na=parseFloat(va), nb=parseFloat(vb);
          if (!isNaN(na)&&!isNaN(nb)) return sortDir*(na-nb);
          return sortDir*String(va).localeCompare(String(vb),'ru',{numeric:true});
        });
      }

      const tbody = document.querySelector('#table tbody');
      tbody.innerHTML = '';
      filtered.forEach(row => {
        const tr = document.createElement('tr');
        cols.forEach(col => {
          const td = document.createElement('td');
          let v = row[col];
          if (v !== '' && !isNaN(v)) {
            const num = Number(v);
            const fmt = num.toLocaleString('ru-RU',{minimumFractionDigits:0,maximumFractionDigits:2});
            if (col === 'Kvadrat') {
              const arrow = num >= avg ? '▲' : '▼';
              const color = num >= avg ? 'green' : 'red';
              td.innerHTML = fmt + ' <span class="arrow" style="color:' + color + '">' + arrow + '</span>';
            } else {
              td.textContent = fmt;
            }
          } else {
            td.textContent = v;
          }
          tr.appendChild(td);
        });
        tbody.appendChild(tr);
      });
    }
  </script>
</body>
</html>`);
});

// Запуск сервера
app.listen(PORT, () => {
  console.log('Server запущен: http://localhost:' + PORT);
});
