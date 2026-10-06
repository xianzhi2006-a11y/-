# -<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>曜日別 三食集計</title>

<style>
*{box-sizing:border-box}
body{
  margin:0;
  padding:20px;
  background:#f4f6f9;
  color:#26344a;
  font:16px system-ui,sans-serif;
}
main{max-width:1400px;margin:auto}
h1{font-size:26px}
h2{font-size:19px}
section{
  background:#fff;
  padding:20px;
  margin:18px 0;
  border-radius:10px;
  border:1px solid #dde3ec;
}
button,input,select{
  font:inherit;
  font-size:14px;
  border:1px solid #bdc7d5;
  border-radius:5px;
  padding:9px;
  background:white;
}
button{cursor:pointer}
button:disabled{opacity:.45;cursor:default}
.primary{background:#792343;color:white;border-color:#792343}
.row{display:flex;gap:12px;flex-wrap:wrap;align-items:center}
label{display:flex;flex-direction:column;gap:6px;font-size:14px}
label select{max-width:260px}
.settings{display:flex;gap:14px;flex-wrap:wrap;margin:16px 0}
.days{
  display:grid;
  grid-template-columns:repeat(auto-fit,minmax(210px,1fr));
  gap:12px;
}
.day{
  background:#f4f6f9;
  padding:14px;
  display:grid;
  gap:10px;
  border-radius:7px;
}
.day select{width:100%;max-width:none}
.scroll{overflow:auto}
table{border-collapse:collapse;width:100%;font-size:14px}
th,td{
  border:1px solid #d4dce8;
  padding:10px;
  text-align:center;
  min-width:125px;
  white-space:pre-line;
}
th{background:#eaf0f7}
td:first-child,th:first-child{min-width:70px}
td:nth-child(2),th:nth-child(2){min-width:140px}
.legend{
  display:flex;
  gap:12px;
  flex-wrap:wrap;
  margin:14px 0;
  font-size:14px;
}
.legend span{display:flex;align-items:center;gap:6px}
.legend i{width:18px;height:18px;border-radius:3px}
#notice{font-size:14px;line-height:1.8}
details{margin:14px 0;font-size:14px}
summary{cursor:pointer}
#filters>div{display:flex;gap:8px;flex-wrap:wrap;margin:10px 0}
.muted{font-size:14px;color:#637187;line-height:1.7}
[hidden]{display:none!important}
</style>

<script src="https://cdn.jsdelivr.net/npm/exceljs@4.4.0/dist/exceljs.min.js"></script>
<script src="https://cdn.jsdelivr.net/npm/papaparse@5.7.0/papaparse.min.js"></script>
</head>

<body>
<main>
  <h1>曜日別 三食集計</h1>

  <div class="row">
    <button id="export" class="primary" disabled>Excelを出力</button>
    <button id="reset" disabled>リセット</button>
  </div>

  <section>
    <h2>① 回答データを読み込む</h2>

    <div class="row">
      <input id="file" type="file" accept=".csv,.xlsx"
             aria-label="回答ファイル">
      <button id="demo">サンプルで試す</button>
    </div>

    <p class="muted">
      Googleフォームの回答CSV、または回答スプレッドシートのExcelを
      選択してください。回答データはブラウザ内で処理します。
    </p>
  </section>

  <p id="notice" role="status">回答ファイルを読み込んでください。</p>

  <section id="config" hidden>
    <h2>② 項目と条件を選ぶ</h2>
    <p id="source" class="muted"></p>

    <div class="settings">
      <label>名前
        <select id="name"></select>
      </label>

      <label>学年
        <select id="grade"></select>
      </label>

      <label>ふりがな
        <select id="reading"></select>
      </label>

      <label>回答の扱い
        <select id="duplicate">
          <option value="latest">同じ学年・名前の最後の回答</option>
          <option value="all">すべての回答をまとめる</option>
        </select>
      </label>

      <label>曜日の形式
        <select id="mode">
          <option value="columns">曜日×三食の項目がある</option>
          <option value="single">曜日を1つの項目に回答</option>
          <option value="date">日付から曜日を判定</option>
        </select>
      </label>
    </div>

    <div class="settings" id="row-fields" hidden>
      <label>曜日・日付
        <select id="day"></select>
      </label>

      <label>朝食
        <select id="meal0"></select>
      </label>

      <label>昼食
        <select id="meal1"></select>
      </label>

      <label>夕食
        <select id="meal2"></select>
      </label>
    </div>

    <div class="settings">
      <label>食事区分の項目（別項目の場合）
        <select id="diet"></select>
      </label>

      <label>弁当・食堂の項目（別項目の場合）
        <select id="place"></select>
      </label>
    </div>

    <div id="daycolumns" class="days"></div>

    <details>
      <summary>名前の読みを補う</summary>
      <p>
        漢字名だけでは五十音順を確定できません。
        読みがない人は入力してください。
      </p>
      <div id="readings" class="settings"></div>
    </details>

    <details id="classification" hidden>
      <summary>未分類の回答を設定する</summary>
      <div id="category-map" class="settings"></div>
    </details>

    <div class="row">
      <b>絞り込み条件</b>
      <button id="addfilter">＋ 条件を追加</button>
    </div>

    <div id="filters"></div>

    <div class="settings">
      <label>名前を検索
        <input id="search" placeholder="名前を入力">
      </label>

      <label>集計する項目
        <select id="group"></select>
      </label>
    </div>
  </section>

  <section>
    <h2>③ 週間三食一覧</h2>
    <p id="count">未読み込み</p>

    <div id="legend" class="legend"></div>

    <div class="scroll">
      <table>
        <thead id="board-head"></thead>
        <tbody id="board">
          <tr><td>回答データを読み込んでください。</td></tr>
        </tbody>
      </table>
    </div>

    <div id="summary" class="muted"></div>
  </section>

  <p class="muted">
    学年昇順 → 同学年内はふりがなの五十音順。
    読みがない漢字名は同学年の末尾です。
    空欄・未設定・未分類は欠食に含めません。
    同じ食事の枠に異なる回答がある場合は「要確認」と表示します。
    <br>
    Excelには週間一覧・月〜日の各曜日の三食シート・
    条件別集計・対象回答・設定を保存します。
    リセットやページを閉じると回答と設定は消えます。
  </p>
</main>

<script>
const $ = id => document.getElementById(id);

const days = ['月','火','水','木','金','土','日'];
const meals = ['朝食','昼食','夕食'];

const categories = [
  ['アスリート食の弁当','E53935','FFFFFF'],
  ['アスリート食の食堂','FB8C00','172033'],
  ['普通食の弁当','1565C0','FFFFFF'],
  ['普通食の食堂','81D4FA','172033'],
  ['欠食','BDBDBD','172033']
];

let rows = [];
let headers = [];
let filters = [];
let readings = {};
let overrides = {};
let sample = false;
let version = 0;

const escapeHTML = s =>
  String(s ?? '').replace(/[&<>"']/g, c => ({
    '&':'&amp;',
    '<':'&lt;',
    '>':'&gt;',
    '"':'&quot;',
    "'":'&#39;'
  }[c]));

const normal = s =>
  String(s ?? '').normalize('NFKC').replace(/\s/g,'');

const kana = s =>
  normal(s).replace(/[ァ-ヶ]/g, c =>
    String.fromCharCode(c.charCodeAt(0) - 96)
  );

const key = (grade,name) =>
  normal(grade) + '|' + String(name).trim();

function gradeNumber(s) {
  const text = normal(s);
  const number = text.match(/\d+/);

  if (number) return Number(number[0]);

  const index = '一二三四五六七八九'.indexOf(text[0]);
  return index < 0 ? 999 : index + 1;
}

function options(value = '', blank = false) {
  return (blank ? '<option value="">未設定</option>' : '') +
    headers.map(h => `
      <option value="${escapeHTML(h)}"
        ${h === value ? 'selected' : ''}>
        ${escapeHTML(h)}
      </option>
    `).join('');
}

function config() {
  return {
    name: $('name').value,
    grade: $('grade').value,
    reading: $('reading').value,
    mode: $('mode').value,
    day: $('day').value,
    diet: $('diet').value,
    place: $('place').value,

    cols: Array.from(
      {length:21},
      (_,i) => $('col' + i).value
    ),

    mealCols: meals.map(
      (_,i) => $('meal' + i).value
    )
  };
}

function eligible(c, applyFilters = true) {
  let result = rows;

  if ($('duplicate').value === 'latest') {
    const map = new Map();

    rows.forEach(r => {
      if (String(r[c.name] ?? '').trim()) {
        map.set(key(r[c.grade], r[c.name]), r);
      }
    });

    result = [...map.values()];
  }

  return result.filter(r => {
    if (!String(r[c.name] ?? '').trim()) return false;
    if (!applyFilters) return true;

    return String(r[c.name]).includes($('search').value.trim()) &&
      filters.every(f =>
        !f.value ||
        String(r[f.field] ?? '') === JSON.parse(f.value)
      );
  });
}

function dayIndices(value, mode) {
  const s = String(value ?? '');

  if (mode === 'date') {
    const match = s.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})/);
    if (!match) return [];

    const [year,month,day] = match.slice(1).map(Number);
    const date = new Date(Date.UTC(year,month - 1,day));

    if (
      date.getUTCFullYear() !== year ||
      date.getUTCMonth() !== month - 1 ||
      date.getUTCDate() !== day
    ) return [];

    return [(date.getUTCDay() + 6) % 7];
  }

  return days.flatMap((day,index) =>
    s.replace(/曜日/g,'').includes(day) ? [index] : []
  );
}

function classify(value,diet,place) {
  const raw = String(value ?? '').trim();

  if (!raw) return '—';
  if (overrides[raw]) return overrides[raw];

  const s = normal(raw);

  if (/欠食|食べない|不要|なし|無し/.test(s)) {
    return '欠食';
  }

  const text = s + normal(diet) + normal(place);

  const type = /アスリート/.test(text)
    ? 'アスリート食'
    : /普通|通常/.test(text)
      ? '普通食'
      : '';

  const location = /弁当|持ち帰り/.test(text)
    ? '弁当'
    : /食堂/.test(text)
      ? '食堂'
      : '';

  return type && location
    ? type + 'の' + location
    : '未分類：' + raw;
}

function appearance(value) {
  const category = categories.find(c => c[0] === value);

  return category || [
    '',
    value === '—' || value === '未設定' ? 'F2F4F7' : 'FFF3CD',
    '26344A'
  ];
}

function board(data,c) {
  const map = new Map();

  data.forEach(r => {
    const name = String(r[c.name]).trim();
    const grade = String(r[c.grade] ?? '').trim();
    const id = key(grade,name);

    if (!map.has(id)) {
      map.set(id, {
        key:id,
        name,
        grade,
        reading:String(r[c.reading] ?? '').trim(),
        slots:Array.from({length:21}, () => [])
      });
    }

    const person = map.get(id);

    if (r[c.reading]) {
      person.reading = String(r[c.reading]).trim();
    }

    const add = (index,column) => {
      if (column && String(r[column] ?? '').trim()) {
        person.slots[index].push(
          classify(r[column],r[c.diet],r[c.place])
        );
      }
    };

    if (c.mode === 'columns') {
      c.cols.forEach((column,index) => add(index,column));
    } else {
      dayIndices(r[c.day],c.mode).forEach(day => {
        c.mealCols.forEach((column,meal) => {
          add(day * 3 + meal,column);
        });
      });
    }
  });

  const people = [...map.values()].map(person => {
    person.reading =
      readings[person.key] ||
      person.reading ||
      (/^[ぁ-ゖァ-ヶー\s]+$/.test(person.name) ? person.name : '');

    person.slots = person.slots.map((values,index) => {
      const column = c.mode === 'columns'
        ? c.cols[index]
        : c.mealCols[index % 3];

      if (!column) return '未設定';

      const unique = [...new Set(values)];

      if (!unique.length) return '—';

      return (unique.length > 1 ? '要確認：\n' : '') +
        unique.join('\n');
    });

    return person;
  });

  return people.sort((a,b) => {
    const gradeOrder = gradeNumber(a.grade) - gradeNumber(b.grade);
    if (gradeOrder) return gradeOrder;

    if (a.reading && !b.reading) return -1;
    if (!a.reading && b.reading) return 1;

    return kana(a.reading || a.name)
      .localeCompare(kana(b.reading || b.name),'ja') ||
      a.name.localeCompare(b.name,'ja');
  });
}

function renderFilters() {
  $('filters').innerHTML = filters.map((filter,index) => `
    <div>
      <select data-field="${index}" aria-label="条件の項目">
        ${options(filter.field)}
      </select>

      <select data-value="${index}" aria-label="条件の値">
        <option value="">すべて</option>
        ${
          [...new Set(rows.map(r => String(r[filter.field] ?? '')))]
            .sort()
            .map(value => `
              <option
                value="${escapeHTML(JSON.stringify(value))}"
                ${JSON.stringify(value) === filter.value ? 'selected' : ''}
              >
                ${escapeHTML(value || '（空欄）')}
              </option>
            `).join('')
        }
      </select>

      <button data-remove="${index}">削除</button>
    </div>
  `).join('');

  $('filters').querySelectorAll('[data-field]').forEach(element => {
    element.onchange = () => {
      filters[element.dataset.field] = {
        field:element.value,
        value:''
      };
      renderFilters();
      render();
    };
  });

  $('filters').querySelectorAll('[data-value]').forEach(element => {
    element.onchange = () => {
      filters[element.dataset.value].value = element.value;
      render();
    };
  });

  $('filters').querySelectorAll('[data-remove]').forEach(element => {
    element.onclick = () => {
      filters.splice(Number(element.dataset.remove),1);
      renderFilters();
      render();
    };
  });
}

function renderReadings() {
  const c = config();
  const people = board(eligible(c,false),c);

  $('readings').innerHTML = people.map((person,index) => `
    <label>
      ${escapeHTML(person.grade + ' ' + person.name)}
      <input
        data-reading="${index}"
        value="${escapeHTML(person.reading)}"
        placeholder="ふりがな"
      >
    </label>
  `).join('');

  $('readings').querySelectorAll('input').forEach(element => {
    element.oninput = () => {
      readings[people[Number(element.dataset.reading)].key] =
        element.value;
      render();
    };
  });
}

function render() {
  if (!rows.length) return;

  const c = config();
  const data = eligible(c);
  const people = board(data,c);

  $('row-fields').hidden = c.mode === 'columns';
  $('daycolumns').hidden = c.mode !== 'columns';
  $('export').disabled = !people.length;

  $('count').textContent =
    `${people.length}名 / ${data.length}回答` +
    (sample ? '（サンプル）' : '');

  $('board-head').innerHTML =
    '<tr><th rowspan="2">学年</th><th rowspan="2">名前</th>' +
    days.map(day => `<th colspan="3">${day}曜日</th>`).join('') +
    '</tr><tr>' +
    days.map(() =>
      meals.map(meal => `<th>${meal}</th>`).join('')
    ).join('') +
    '</tr>';

  $('board').innerHTML = people.length
    ? people.map(person => `
      <tr>
        <td>${escapeHTML(person.grade || '未設定')}</td>
        <td>${escapeHTML(person.name)}</td>
        ${
          person.slots.map(value => {
            const color = appearance(value);

            return `
              <td style="background:#${color[1]};color:#${color[2]}">
                ${escapeHTML(value)}
              </td>
            `;
          }).join('')
        }
      </tr>
    `).join('')
    : '<tr><td colspan="23">条件に一致する回答がありません。</td></tr>';

  $('legend').innerHTML = categories.map(category => `
    <span>
      <i style="background:#${category[1]}"></i>
      ${category[0]}
    </span>
  `).join('');

  const counts = new Map();

  data.forEach(row => {
    const value = String(row[$('group').value] || '（空欄）');
    counts.set(value,(counts.get(value) || 0) + 1);
  });

  $('summary').textContent = [...counts]
    .map(([value,count]) => `${value}：${count}回答`)
    .join(' / ');

  const unknown = [...new Set(
    people.flatMap(person =>
      person.slots.flatMap(value =>
        value.replace(/^要確認：\n/,'')
          .split('\n')
          .filter(v => v.startsWith('未分類：'))
          .map(v => v.slice(4))
      )
    )
  )];

  const values = [...new Set([
    ...unknown,
    ...Object.keys(overrides)
  ])];

  $('classification').hidden = !values.length;

  $('category-map').innerHTML = values.map((value,index) => `
    <label>
      ${escapeHTML(value)}
      <select data-category="${index}">
        <option value="">自動判定</option>
        ${
          categories.map(category => `
            <option ${overrides[value] === category[0] ? 'selected' : ''}>
              ${category[0]}
            </option>
          `).join('')
        }
      </select>
    </label>
  `).join('');

  $('category-map').querySelectorAll('select').forEach(element => {
    element.onchange = () => {
      const value = values[Number(element.dataset.category)];

      if (element.value) overrides[value] = element.value;
      else delete overrides[value];

      render();
    };
  });

  const notes = [];

  if (sample) notes.push('サンプルを表示中');

  const noReading = people.filter(person => !person.reading).length;
  if (noReading) notes.push(`${noReading}名の読みが未入力です`);

  if (people.some(person => person.slots.includes('未設定'))) {
    notes.push('未設定の食事項目があります');
  }

  if (unknown.length) notes.push('未分類の回答を設定してください');

  if (people.some(person =>
    person.slots.some(value => value.startsWith('要確認：'))
  )) {
    notes.push('複数回答の枠を確認してください');
  }

  $('notice').textContent = notes.length
    ? notes.join('。') + '。'
    : '学年昇順・同学年内の五十音順で表示しています。';
}

function mealMatch(header,index) {
  return index === 0
    ? /朝/.test(header)
    : index === 1
      ? /昼/.test(header)
      : /夕|晩|夜/.test(header);
}

function load(data,label,isSample = false) {
  if (!data.length) throw Error('回答データがありません。');

  headers = [...new Set(
    data.flatMap(row => Object.keys(row))
  )].filter(header => header.trim());

  if (!headers.length) throw Error('見出し行がありません。');

  version++;
  rows = data;
  filters = [];
  readings = {};
  overrides = {};
  sample = isSample;

  $('reset').disabled = false;

  const name = headers.find(header =>
    /名前|氏名|姓名|name/i.test(header) &&
    !/ふりがな|フリガナ|読み|カナ/.test(header)
  ) || headers[0];

  const grade = headers.find(header => /学年/.test(header)) || '';

  const reading = headers.find(header =>
    /ふりがな|フリガナ|読み|カナ/.test(header)
  ) || '';

  const selected = {
    name,
    grade,
    reading,
    group:grade || headers[0],
    day:headers.find(header => /曜日|日付/.test(header)) || headers[0],
    diet:headers.find(header =>
      /食事区分|食種|食事種類/.test(header)
    ) || '',
    place:headers.find(header =>
      /受取|食事場所/.test(header)
    ) || ''
  };

  Object.entries(selected).forEach(([id,value]) => {
    $(id).innerHTML = options(
      value,
      ['grade','reading','diet','place'].includes(id)
    );
  });

  meals.forEach((_,index) => {
    $('meal' + index).innerHTML = options(
      headers.find(header =>
        mealMatch(header,index) &&
        !days.some(day => header.includes(day + '曜'))
      ) || '',
      true
    );
  });

  $('daycolumns').innerHTML = days.map((day,dayIndex) => `
    <div class="day">
      <b>${day}曜日</b>
      ${
        meals.map((meal,mealIndex) => `
          <label>
            ${meal}
            <select id="col${dayIndex * 3 + mealIndex}">
              ${
                options(
                  headers.find(header =>
                    header.includes(day + '曜') &&
                    mealMatch(header,mealIndex)
                  ) || '',
                  true
                )
              }
            </select>
          </label>
        `).join('')
      }
    </div>
  `).join('');

  $('mode').value = Array.from(
    {length:21},
    (_,index) => $('col' + index).value
  ).some(Boolean)
    ? 'columns'
    : headers.some(header => /日付/.test(header))
      ? 'date'
      : 'single';

  $('source').textContent = label;
  $('search').value = '';
  $('config').hidden = false;

  Array.from({length:21},(_,index) => {
    $('col' + index).onchange = render;
  });

  renderFilters();
  renderReadings();
  render();
}

$('file').onchange = async event => {
  const file = event.target.files[0];
  if (!file) return;

  const currentVersion = ++version;

  try {
    if (file.size > 20 * 1024 * 1024) {
      throw Error('20MB以下のファイルを選んでください。');
    }

    let data;

    if (/\.csv$/i.test(file.name)) {
      if (!window.Papa) {
        throw Error(
          'CSVライブラリを読み込めません。ネット接続を確認してください。'
        );
      }

      const bytes = await file.arrayBuffer();
      let text = new TextDecoder('utf-8').decode(bytes);

      if (text.includes('�')) {
        text = new TextDecoder('shift_jis').decode(bytes);
      }

      const parsed = Papa.parse(text.replace(/^\uFEFF/,''), {
        header:true,
        skipEmptyLines:'greedy',
        transformHeader:header => header.trim()
      });

      if (parsed.errors.length) {
        throw Error('CSVの列数や引用符を確認してください。');
      }

      data = parsed.data;
    } else {
      if (!window.ExcelJS) {
        throw Error(
          'Excelライブラリを読み込めません。ネット接続を確認してください。'
        );
      }

      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.load(await file.arrayBuffer());

      const sheet = workbook.worksheets[0];
      if (!sheet) throw Error('シートがありません。');

      const names = sheet.getRow(1).values.slice(1)
        .map(value => String(value ?? '').trim());

      if (
        names.some(value => !value) ||
        new Set(names).size !== names.length
      ) {
        throw Error('1行目に重複のない項目名を入れてください。');
      }

      data = [];

      sheet.eachRow((row,rowNumber) => {
        if (rowNumber === 1) return;

        const item = {};

        names.forEach((header,index) => {
          let value = row.getCell(index + 1).value;

          if (value instanceof Date) {
            value =
              `${value.getFullYear()}-` +
              `${String(value.getMonth() + 1).padStart(2,'0')}-` +
              `${String(value.getDate()).padStart(2,'0')}`;
          } else if (value && typeof value === 'object') {
            value =
              value.text ??
              value.result ??
              value.richText?.map(part => part.text).join('') ??
              '';
          }

          item[header] = value ?? '';
        });

        data.push(item);
      });
    }

    if (currentVersion === version) load(data,file.name);
  } catch (error) {
    if (currentVersion === version) {
      $('notice').textContent =
        '読み込めませんでした：' + error.message;
    }
  }

  event.target.value = '';
};

$('demo').onclick = () => {
  const people = [
    ['山田 太郎','2年','やまだ たろう'],
    ['佐藤 健','1年','さとう けん'],
    ['鈴木 翔','3年','すずき しょう'],
    ['田中 直樹','2年','たなか なおき']
  ];

  const data = people.map((person,index) => {
    const row = {
      氏名:person[0],
      学年:person[1],
      ふりがな:person[2]
    };

    days.forEach((day,dayIndex) => {
      meals.forEach((meal,mealIndex) => {
        row[day + '曜日 ' + meal] =
          categories[(index + dayIndex * 3 + mealIndex) % 5][0];
      });
    });

    return row;
  });

  load(data,'操作確認用サンプル',true);
};

[
  'name','grade','reading','duplicate','mode','day',
  'diet','place','meal0','meal1','meal2','group'
].forEach(id => {
  $(id).onchange = () => {
    renderReadings();
    render();
  };
});

$('search').oninput = render;

$('addfilter').onclick = () => {
  filters.push({
    field:$('grade').value || headers[0],
    value:''
  });

  renderFilters();
  render();
};

function reset() {
  version++;

  rows = [];
  headers = [];
  filters = [];
  readings = {};
  overrides = {};
  sample = false;

  $('file').value = '';
  $('search').value = '';
  $('mode').value = 'columns';
  $('duplicate').value = 'latest';

  [
    'name','grade','reading','day','diet','place',
    'meal0','meal1','meal2','group','source','filters',
    'readings','category-map','daycolumns','legend',
    'summary','board-head'
  ].forEach(id => {
    $(id).innerHTML = '';
  });

  $('config').hidden = true;
  $('classification').hidden = true;

  document.querySelectorAll('details').forEach(element => {
    element.open = false;
  });

  $('board').innerHTML =
    '<tr><td>回答データを読み込んでください。</td></tr>';

  $('count').textContent = '未読み込み';
  $('notice').textContent = 'リセットしました。';
  $('export').disabled = true;
  $('reset').disabled = true;
}

$('reset').onclick = reset;

function styleSheet(sheet,headerRows) {
  sheet.eachRow(row => {
    row.eachCell(cell => {
      cell.font = {
        name:'Yu Gothic',
        size:11,
        color:{argb:'FF26344A'}
      };

      cell.alignment = {
        vertical:'middle',
        wrapText:true
      };

      cell.border = Object.fromEntries(
        ['top','bottom','left','right'].map(side => [
          side,
          {
            style:'thin',
            color:{argb:'FFD4DCE8'}
          }
        ])
      );
    });
  });

  headerRows.forEach(rowNumber => {
    const row = sheet.getRow(rowNumber);
    row.height = 26;

    row.eachCell(cell => {
      cell.fill = {
        type:'pattern',
        pattern:'solid',
        fgColor:{argb:'FF26344A'}
      };

      cell.font = {
        name:'Yu Gothic',
        size:11,
        bold:true,
        color:{argb:'FFFFFFFF'}
      };
    });
  });
}

function paint(cell,value) {
  const color = appearance(value);

  cell.fill = {
    type:'pattern',
    pattern:'solid',
    fgColor:{argb:'FF' + color[1]}
  };

  cell.font = {
    name:'Yu Gothic',
    size:11,
    color:{argb:'FF' + color[2]}
  };
}

function makeWorkbook(people,data,c) {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'Weekly Meals';

  const pageSetup = {
    orientation:'landscape',
    paperSize:9,
    fitToPage:true,
    fitToWidth:1,
    fitToHeight:0
  };

  // 週間三食一覧
  const weekly = workbook.addWorksheet('週間三食一覧', {
    views:[{state:'frozen',xSplit:2,ySplit:3}],
    pageSetup
  });

  weekly.mergeCells(1,1,1,23);
  weekly.getCell('A1').value =
    '週間三食一覧' + (sample ? '（サンプル）' : '');

  weekly.getCell('A2').value = '学年';
  weekly.getCell('B2').value = '名前';

  weekly.mergeCells('A2:A3');
  weekly.mergeCells('B2:B3');

  days.forEach((day,dayIndex) => {
    weekly.mergeCells(
      2,3 + dayIndex * 3,
      2,5 + dayIndex * 3
    );

    weekly.getCell(2,3 + dayIndex * 3).value = day + '曜日';

    meals.forEach((meal,mealIndex) => {
      weekly.getCell(3,3 + dayIndex * 3 + mealIndex).value = meal;
    });
  });

  people.forEach(person => {
    const row = weekly.addRow([
      person.grade || '未設定',
      person.name,
      ...person.slots
    ]);

    row.height = 50;
  });

  weekly.columns.forEach((column,index) => {
    column.width = index === 0 ? 9 : 22;
  });

  styleSheet(weekly,[2,3]);

  people.forEach((person,index) => {
    person.slots.forEach((value,column) => {
      paint(weekly.getCell(index + 4,column + 3),value);
    });
  });

  // 月曜日〜日曜日の各シート
  days.forEach((day,dayIndex) => {
    const sheet = workbook.addWorksheet(day + '曜日', {
      views:[{state:'frozen',xSplit:2,ySplit:3}],
      pageSetup
    });

    sheet.mergeCells('A1:E1');
    sheet.getCell('A1').value =
      day + '曜日 三食一覧' + (sample ? '（サンプル）' : '');

    sheet.mergeCells('A2:E2');
    sheet.getCell('A2').value =
      '学年昇順・同学年内は名前の五十音順';

    sheet.addRow(['学年','名前',...meals]);

    people.forEach(person => {
      const row = sheet.addRow([
        person.grade || '未設定',
        person.name,
        ...person.slots.slice(dayIndex * 3,dayIndex * 3 + 3)
      ]);

      row.height = 50;
    });

    sheet.columns = [
      {width:10},
      {width:22},
      ...meals.map(() => ({width:30}))
    ];

    sheet.autoFilter = {
      from:{row:3,column:1},
      to:{row:people.length + 3,column:5}
    };

    styleSheet(sheet,[3]);

    people.forEach((person,index) => {
      person.slots.slice(dayIndex * 3,dayIndex * 3 + 3)
        .forEach((value,column) => {
          paint(sheet.getCell(index + 4,column + 3),value);
        });
    });

    categories.forEach((category,index) => {
      const rowNumber = people.length + 6 + index;

      sheet.mergeCells(rowNumber,2,rowNumber,5);
      sheet.getCell(rowNumber,2).value = category[0];
      paint(sheet.getCell(rowNumber,2),category[0]);
    });
  });

  // 条件別集計
  const summary = workbook.addWorksheet('条件別集計');
  summary.addRow(['集計項目','値','回答数']);

  const counts = new Map();

  data.forEach(row => {
    const value = String(row[$('group').value] || '（空欄）');
    counts.set(value,(counts.get(value) || 0) + 1);
  });

  [...counts].forEach(([value,count]) => {
    summary.addRow([$('group').value,value,count]);
  });

  summary.addRow([]);

  const mealHeader = summary.addRow([
    '曜日','食事','分類','人数'
  ]);

  days.forEach((day,dayIndex) => {
    meals.forEach((meal,mealIndex) => {
      const labels = [
        ...categories.map(category => category[0]),
        '未分類','要確認','空欄','未設定'
      ];

      labels.forEach(category => {
        const count = people.filter(person => {
          const value = person.slots[dayIndex * 3 + mealIndex];

          if (category === '未分類') {
            return value.startsWith('未分類：');
          }

          if (category === '要確認') {
            return value.startsWith('要確認：');
          }

          if (category === '空欄') return value === '—';

          return value === category;
        }).length;

        summary.addRow([day + '曜日',meal,category,count]);
      });
    });
  });

  summary.columns.forEach(column => {
    column.width = 28;
  });

  styleSheet(summary,[1,mealHeader.number]);

  // 対象回答
  const raw = workbook.addWorksheet('対象回答');
  raw.addRow(headers);

  const order = new Map(
    people.map((person,index) => [person.key,index])
  );

  [...data].sort((a,b) =>
    order.get(key(a[c.grade],a[c.name])) -
    order.get(key(b[c.grade],b[c.name]))
  ).forEach(row => {
    raw.addRow(headers.map(header => String(row[header] ?? '')));
  });

  raw.columns.forEach(column => {
    column.width = 25;
  });

  styleSheet(raw,[1]);
  raw.views = [{state:'frozen',ySplit:1}];

  // 設定
  const settings = workbook.addWorksheet('設定');
  settings.addRow(['項目','設定']);

  settings.addRow(['データ',$('source').textContent]);
  settings.addRow(['名前',c.name]);
  settings.addRow(['学年',c.grade || '未設定']);
  settings.addRow(['ふりがな',c.reading || '未設定']);

  settings.addRow([
    '並び順',
    '学年昇順 → 同学年内の五十音順。読み未入力の漢字名は同学年の末尾'
  ]);

  settings.addRow([
    '回答の扱い',
    $('duplicate').selectedOptions[0].text
  ]);

  settings.addRow(['名前検索',$('search').value || 'なし']);
  settings.addRow(['曜日の形式',$('mode').selectedOptions[0].text]);
  settings.addRow(['食事区分',c.diet || '未設定']);
  settings.addRow(['弁当・食堂',c.place || '未設定']);

  filters.forEach(filter => {
    settings.addRow([
      '条件：' + filter.field,
      filter.value
        ? JSON.parse(filter.value) || '（空欄）'
        : 'すべて'
    ]);
  });

  if (c.mode === 'columns') {
    c.cols.forEach((value,index) => {
      settings.addRow([
        days[Math.floor(index / 3)] + '曜日 ' + meals[index % 3],
        value || '未設定'
      ]);
    });
  } else {
    settings.addRow(['曜日・日付',c.day]);

    c.mealCols.forEach((value,index) => {
      settings.addRow([meals[index],value || '未設定']);
    });
  }

  people.forEach(person => {
    settings.addRow([
      '読み：' + person.grade + ' ' + person.name,
      person.reading || '未入力'
    ]);
  });

  Object.entries(overrides).forEach(([value,category]) => {
    settings.addRow(['分類：' + value,category]);
  });

  settings.addRow([
    '注意',
    '空欄・未設定・未分類は欠食に数えません。要確認の枠には複数回答を残します。'
  ]);

  categories.forEach(category => {
    settings.addRow([
      '色：' + category[0],
      '#' + category[1]
    ]);
  });

  settings.columns = [{width:35},{width:85}];
  styleSheet(settings,[1]);

  return workbook;
}

async function exportExcel() {
  const currentVersion = version;
  if (!rows.length) return;

  const c = config();
  const data = eligible(c);
  const people = board(data,c);

  if (!people.length) return;

  $('export').disabled = true;

  try {
    if (!window.ExcelJS) {
      throw Error(
        'Excelライブラリを読み込めません。ネット接続を確認してください。'
      );
    }

    const workbook = makeWorkbook(people,data,c);
    const bytes = await workbook.xlsx.writeBuffer();

    if (currentVersion !== version) return;

    const url = URL.createObjectURL(new Blob([bytes], {
      type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    }));

    const link = document.createElement('a');
    link.href = url;
    link.download =
      (sample ? 'サンプル_' : '') + '曜日別三食集計.xlsx';

    link.click();
    setTimeout(() => URL.revokeObjectURL(url),10000);

    $('notice').textContent =
      'Excelを出力しました。' +
      (
        people.some(person => !person.reading)
          ? '読み未入力の漢字名は同学年の末尾です。'
          : ''
      );
  } catch (error) {
    if (currentVersion === version) {
      $('notice').textContent =
        '出力できませんでした：' + error.message;
    }
  } finally {
    $('export').disabled = !rows.length;
  }
}

$('export').onclick = exportExcel;
</script>
</body>
</html>