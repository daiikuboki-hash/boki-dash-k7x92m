// ===== 共通 =====
const $ = id => document.getElementById(id);
const LS_URL = 'boki_gas_url';
const LS_KEY = 'boki_gas_key';
let parsedFiles = []; // 読み込んだExcelの解析結果

function esc(s) {
  return String(s).replace(/[&<>"']/g, c => (
    { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]
  ));
}

// ===== 接続設定 =====
$('gas-url').value = localStorage.getItem(LS_URL) || '';
$('gas-key').value = localStorage.getItem(LS_KEY) || '';

$('save-settings').addEventListener('click', () => {
  localStorage.setItem(LS_URL, $('gas-url').value.trim());
  localStorage.setItem(LS_KEY, $('gas-key').value);
  $('settings-msg').textContent = '保存しました';
});
loadData();

// ===== Excelの解析 =====
function parseWorkbook(buffer, fileName) {
  const wb = XLSX.read(buffer, { type: 'array' });
  const ws = wb.Sheets[wb.SheetNames[0]];
  const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });

  // 2行目: 第223回 / 3行目: 試験日
  const round = String((aoa[1] && aoa[1][0]) || '').trim();
  const dateText = String((aoa[2] && aoa[2][1]) || '');
  const m = dateText.match(/(\d{4})年(\d{1,2})月(\d{1,2})日/);
  if (!m) throw new Error(fileName + ': 試験日が読み取れません');
  const examDate = `${m[1]}-${m[2].padStart(2, '0')}-${m[3].padStart(2, '0')}`;

  // 見出し行を探し、列の位置を名前から決める
  const hIdx = aoa.findIndex(r => r.includes('検定級段'));
  if (hIdx < 0) throw new Error(fileName + ': 見出し行(検定級段)が見つかりません');
  const h = aoa[hIdx];
  const col = {
    grade: h.indexOf('検定級段'),
    attendance: h.indexOf('出欠'),
    score: h.indexOf('試験結果'),
    result: h.indexOf('判定')
  };
  if (Object.values(col).some(v => v < 0)) {
    throw new Error(fileName + ': 必要な列(検定級段・出欠・試験結果・判定)が足りません');
  }
    const kanaSeiCol = h.indexOf('フリガナ（姓）');
    const kanaMeiCol = h.indexOf('フリガナ（名）');

  // 氏名・個人IDは読み取らない(4列+フリガナ2列だけ)
  const rows = [];
  for (let i = hIdx + 1; i < aoa.length; i++) {
    const r = aoa[i];
    const grade = String(r[col.grade] || '').trim();
    if (!grade) continue;
    rows.push({
      grade: grade,
      attendance: String(r[col.attendance] || '').trim(),
      score: r[col.score] === '' ? '' : Number(r[col.score]),
      result: String(r[col.result] || '').trim(),
      kanaSei: kanaSeiCol < 0 ? '' : String(r[kanaSeiCol] || '').trim(),
      kanaMei: kanaMeiCol < 0 ? '' : String(r[kanaMeiCol] || '').trim()
    });
  }
  return { fileName, round, examDate, rows };
}

// ===== ファイル選択 → プレビュー =====
$('file-input').addEventListener('change', async (e) => {
  parsedFiles = [];
  $('upload-msg').textContent = '';
  $('upload-btn').disabled = true;
  const files = Array.from(e.target.files);
  if (files.length === 0) { $('upload-preview').innerHTML = ''; return; }

  try {
    for (const f of files) {
      const buf = await f.arrayBuffer();
      parsedFiles.push(parseWorkbook(buf, f.name));
    }
    renderPreview();
    $('upload-btn').disabled = false;
  } catch (err) {
    $('upload-preview').innerHTML = '';
    $('upload-msg').textContent = '読み込みエラー: ' + err.message;
  }
});

function renderPreview() {
  let html = '<table class="preview"><tr><th>ファイル</th><th>回</th><th>試験日</th>' +
    '<th>級</th><th>人数</th><th>出席</th><th>欠席</th><th>合格</th></tr>';
  parsedFiles.forEach(p => {
    const grades = {};
    p.rows.forEach(r => {
      const g = grades[r.grade] || (grades[r.grade] = { n: 0, att: 0, abs: 0, pass: 0 });
      g.n++;
      if (r.attendance === '出席') g.att++;
      if (r.attendance === '欠席') g.abs++;
      if (r.result === '合格') g.pass++;
    });
    Object.keys(grades).forEach(name => {
      const g = grades[name];
      html += `<tr><td>${esc(p.fileName)}</td><td>${esc(p.round)}</td><td>${esc(p.examDate)}</td>` +
        `<td>${esc(name)}</td><td>${g.n}</td><td>${g.att}</td><td>${g.abs}</td><td>${g.pass}</td></tr>`;
    });
  });
  html += '</table>';
  $('upload-preview').innerHTML = html;
}

// ===== 保存(GASへ送信) =====
$('upload-btn').addEventListener('click', async () => {
  const url = localStorage.getItem(LS_URL);
  const key = localStorage.getItem(LS_KEY);
  if (!url || !key) {
    $('upload-msg').textContent = '先に「1. 接続設定」を保存してください';
    return;
  }

  $('upload-btn').disabled = true;
  let total = 0;
  try {
    for (let i = 0; i < parsedFiles.length; i++) {
      const p = parsedFiles[i];
      $('upload-msg').textContent = `送信中 ${i + 1}/${parsedFiles.length} ...`;
      const res = await fetch(url, {
        method: 'POST',
        headers: { 'Content-Type': 'text/plain;charset=utf-8' }, // 事前確認(preflight)を避けるため
        body: JSON.stringify({ key: key, examDate: p.examDate, round: p.round, rows: p.rows })
      });
      const data = await res.json();
      if (!data.ok) throw new Error(data.error);
      total += data.saved;
    }
    $('upload-msg').textContent = `${total}件を保存しました`;
    await loadData();
  } catch (err) {
    $('upload-msg').textContent = '保存エラー: ' + err.message;
  } finally {
    $('upload-btn').disabled = false;
  }
});

// ===== Step 5(デザイン改善版): データ取得と表示 =====
let allRows = [];
let charts = [];
const GRADE_ORDER = ['基礎簿記会計', '3級商業簿記', '2級商業簿記', '1級商業簿記', '1級会計'];

const pct = (a, b) => b ? (a / b * 100).toFixed(1) + '%' : '-';
const gradeOrder = g => { const i = GRADE_ORDER.indexOf(g); return i < 0 ? 99 : i; };

// 合格率は「出席者のうち合格した人」の割合
function aggregate(rows) {
  const att = rows.filter(r => r.attendance === '出席');
  return {
    total: rows.length,
    att: att.length,
    abs: rows.filter(r => r.attendance === '欠席').length,
    pass: att.filter(r => r.result === '合格').length
  };
}

function kpi(label, value, sub, hero) {
  return `<div class="kpi${hero ? ' hero' : ''}"><div class="kpi-label">${label}</div>` +
    `<div class="kpi-value">${value}</div><div class="kpi-sub">${sub}</div></div>`;
}

function seg(cls, n, total, label) {
  if (!n) return '';
  const w = n / total * 100;
  return `<div class="seg ${cls}" style="width:${w}%" title="${label} ${n}人(${w.toFixed(1)}%)">${w >= 7 ? n : ''}</div>`;
}

function barRow(name, a, cls) {
  return `<div class="bar-row ${cls || ''}">` +
    `<div class="bar-name">${esc(name)}<span class="bar-sub">対象 ${a.total}人</span></div>` +
    `<div class="bar">${seg('pass', a.pass, a.total, '合格')}` +
    `${seg('fail', a.att - a.pass, a.total, '不合格')}${seg('absent', a.abs, a.total, '欠席')}</div>` +
    `<div class="bar-rate">${pct(a.pass, a.att)}<span>合格率</span></div></div>`;
}

function tableRow(name, a, isTotal) {
  return `<tr class="${isTotal ? 'total-row' : ''}"><th scope="row">${esc(name)}</th>` +
    `<td>${a.total}人</td><td>${a.att}人</td><td>${a.abs}人</td><td>${pct(a.att, a.total)}</td>` +
    `<td>${a.pass}人</td><td class="rate">${pct(a.pass, a.att)}</td></tr>`;
}

function showLoading(text) {
  $('summary-section').innerHTML =
    `<div class="loading" role="status" aria-live="polite"><div class="spinner"></div><p>${text}</p></div>`;
  $('grade-section').innerHTML = '';
  ensureCompareBox().style.display = 'none';
}

async function loadData() {
  const url = localStorage.getItem(LS_URL);
  const key = localStorage.getItem(LS_KEY);
  if (!url || !key) {
    $('summary-section').innerHTML =
      '<h2>全体の結果</h2><p class="note">まず「1. 接続設定」を入力して、「設定を保存」を押してください。</p>';
    return;
  }
  showLoading('データを読み込み中です…(数秒かかることがあります)');
  try {
    const res = await fetch(`${url}?key=${encodeURIComponent(key)}`);
    const data = await res.json();
    if (!data.ok) throw new Error(data.error);
    allRows = data.rows;
    buildSelects();
  } catch (err) {
    $('summary-section').innerHTML =
      `<h2>全体の結果</h2><p class="msg">データを読み込めませんでした: ${esc(err.message)}</p>` +
      '<p class="note">接続設定のURLと合言葉を確認して、画面を再読み込みしてください。</p>';
  }
}

function buildSelects() {
  const map = {};
  allRows.forEach(r => { map[r.examDate] = r.round; });
  const dates = Object.keys(map).sort().reverse(); // 新しい回が先頭
  const opts = dates.map(d => `<option value="${esc(d)}">${esc(map[d])}(${esc(d)})</option>`).join('');
  const prev = $('current-select').value;
  $('current-select').innerHTML = opts;
  $('compare-select').innerHTML = '<option value="">比較しない</option>' + opts;
  if (prev && dates.includes(prev)) $('current-select').value = prev;
  render();
  fillStudentFilters();
}

function render() {
  charts.forEach(c => c.destroy());
  charts = [];

  const date = $('current-select').value;
  const rows = allRows.filter(r => r.examDate === date);
  if (!date || rows.length === 0) {
    $('summary-section').innerHTML =
      '<h2>全体の結果</h2><p class="note">データがありません。「2. Excelを読み込む」から保存してください。</p>';
    $('grade-section').innerHTML = '';
    ensureCompareBox().style.display = 'none';
    return;
  }

  const all = aggregate(rows);
  const grades = [...new Set(rows.map(r => r.grade))].sort((x, y) => gradeOrder(x) - gradeOrder(y));
  const byGrade = grades.map(g => [g, aggregate(rows.filter(r => r.grade === g))]);
  const [y, m, d] = date.split('-');

  // 全体の結果
  $('summary-section').innerHTML =
    `<h2>${esc(rows[0].round)}の結果<span class="h2-sub">試験日 ${y}年${+m}月${+d}日</span></h2>` +
    '<div class="kpis">' +
            kpi('合格率', pct(all.pass, all.att), `合格者 ÷ 出席者(出席者 ${all.att}人中)`, true) +
      kpi('合格者数', all.pass + '人', `対象者全体では ${pct(all.pass, all.total)}`) +
      kpi('出席率', pct(all.att, all.total), `出席 ${all.att}人 / 欠席 ${all.abs}人(欠席率 ${pct(all.abs, all.total)})`) +
      kpi('対象者数', all.total + '人', `${grades.length}つの級`) +
    '</div>' +
    '<h3>出席・合格の内訳</h3>' +
    '<div class="legend"><span><i class="dot pass"></i>合格</span>' +
    '<span><i class="dot fail"></i>不合格(出席者)</span><span><i class="dot absent"></i>欠席</span></div>' +
    '<p class="note">棒の長さは、対象者全体を100%とした割合です。右の合格率は、出席者に対する割合です。</p>' +
    '<div class="bars">' + barRow('全体', all, 'total') +
    byGrade.map(([g, a]) => barRow(g, a)).join('') + '</div>';

  // 級別の表
  $('grade-section').innerHTML =
    '<h2>級別の結果</h2><div class="table-wrap"><table class="data-table">' +
    '<tr><th>級</th><th>対象者</th><th>出席</th><th>欠席</th><th>出席率</th><th>合格者</th><th>合格率(出席者中)</th></tr>' +
    byGrade.map(([g, a]) => tableRow(g, a, false)).join('') +
    tableRow('全体', all, true) + '</table></div>';

  renderCompare(date);
}

$('current-select').addEventListener('change', render);

// ページを開いたときにデータを取得
loadData();

// ===== Step 6(デザイン改善版): 前回との比較 =====
Chart.defaults.font.family = '"IBM Plex Sans JP", "BIZ UDPGothic", "Yu Gothic UI", Meiryo, sans-serif';
Chart.defaults.font.weight = '700';
Chart.defaults.color = '#16202e';

// 棒の上に数値を表示する小さな部品
const valueLabels = {
  id: 'valueLabels',
  afterDatasetsDraw(chart) {
    const ctx = chart.ctx;
    ctx.save();
    ctx.font = '800 13px "IBM Plex Sans JP", sans-serif';
    ctx.fillStyle = '#16202e';
    ctx.textAlign = 'center';
    chart.data.datasets.forEach((ds, i) => {
      chart.getDatasetMeta(i).data.forEach((bar, j) => {
        const v = ds.data[j];
        if (v === null || v === undefined) return;
        ctx.fillText(v.toFixed(1) + '%', bar.x, bar.y - 6);
      });
    });
    ctx.restore();
  }
};

function metrics(rows) {
  if (!rows.length) return { pass: null, attRate: null, passRate: null };
  const a = aggregate(rows);
  return {
    pass: a.pass,
    attRate: a.total ? a.att / a.total * 100 : null,
    passRate: a.att ? a.pass / a.att * 100 : null
  };
}

function fmtVal(v, unit) {
  if (v === null || v === undefined) return '-';
  return unit === '%' ? v.toFixed(1) + '%' : v + '人';
}

// 今回 − 比較回 の差。増えるほど良い項目だけを扱う
function diffInfo(cur, prev, unit) {
  if (cur === null || prev === null) return { cls: 'flat', text: '-' };
  const d = cur - prev;
  if (Math.abs(d) < 0.05) return { cls: 'flat', text: '変化なし' };
  const abs = unit === '%' ? Math.abs(d).toFixed(1) + 'pt' : Math.abs(d) + '人';
  return d > 0
    ? { cls: 'up', text: '▲ +' + abs }
    : { cls: 'down', text: '▼ −' + abs };
}

function cmpCard(label, cv, pv, unit) {
  const d = diffInfo(cv, pv, unit);
  return `<div class="cmp-card"><div class="cmp-label">${label}</div>` +
    `<div class="cmp-main">${fmtVal(cv, unit)}</div>` +
    `<div class="cmp-prev">比較回:${fmtVal(pv, unit)}</div>` +
    `<div class="badge ${d.cls}">${d.text}</div></div>`;
}

function cmpCells(cv, pv, unit) {
  const d = diffInfo(cv, pv, unit);
  return `<td class="grp">${fmtVal(cv, unit)}</td><td>${fmtVal(pv, unit)}</td>` +
    `<td class="diff ${d.cls}">${d.text}</td>`;
}

function ensureCompareBox() {
  let box = $('compare-box');
  if (!box) {
    box = document.createElement('section');
    box.id = 'compare-box';
    box.className = 'card';
    $('summary-section').insertAdjacentElement('afterend', box);
  }
  return box;
}

function compareChart(canvasId, labels, curData, prevData) {
  const style = { borderWidth: 0, borderRadius: 2 };
  charts.push(new Chart($(canvasId), {
    type: 'bar',
    data: {
      labels: labels,
      datasets: [
        Object.assign({ label: '今回', data: curData, backgroundColor: '#14243d' }, style),
        Object.assign({ label: '比較回', data: prevData, backgroundColor: '#aeb8c6' }, style)
      ]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      layout: { padding: { top: 22 } },
      scales: { y: { min: 0, max: 100, ticks: { callback: v => v + '%' } } },
      plugins: {
        legend: { position: 'top' },
        tooltip: { callbacks: { label: c => `${c.dataset.label}: ${c.parsed.y.toFixed(1)}%` } }
      }
    },
    plugins: [valueLabels]
  }));
}

function renderCompare(curDate) {
  const box = ensureCompareBox();
  const prevDate = $('compare-select').value;
  if (!prevDate || prevDate === curDate) {
    box.style.display = 'none';
    box.innerHTML = '';
    return;
  }
  box.style.display = '';

  const cur = allRows.filter(r => r.examDate === curDate);
  const prev = allRows.filter(r => r.examDate === prevDate);
  const info = d => {
    const r = allRows.find(x => x.examDate === d);
    const [y, m, dd] = d.split('-');
    return `${r.round}(${y}年${+m}月${+dd}日)`;
  };
  const grades = [...new Set(cur.concat(prev).map(r => r.grade))]
    .sort((x, y) => gradeOrder(x) - gradeOrder(y));
  const scopes = [['全体', cur, prev]].concat(
    grades.map(g => [g, cur.filter(r => r.grade === g), prev.filter(r => r.grade === g)])
  );

  const mc = metrics(cur);
  const mp = metrics(prev);

  // 表
  let rowsHtml = '';
  scopes.forEach(([name, cr, pr], i) => {
    const c = metrics(cr);
    const p = metrics(pr);
    rowsHtml += `<tr class="${i === 0 ? 'total-row' : ''}"><th scope="row">${esc(name)}</th>` +
      cmpCells(c.passRate, p.passRate, '%') +
      cmpCells(c.attRate, p.attRate, '%') +
      cmpCells(c.pass, p.pass, '人') + '</tr>';
  });

  box.innerHTML =
    '<h2>前回との比較' +
    `<span class="h2-sub">今回 ${esc(info(curDate))} ／ 比較回 ${esc(info(prevDate))}</span></h2>` +
    '<div class="compare-cards">' +
      cmpCard('合格率(全体)', mc.passRate, mp.passRate, '%') +
      cmpCard('出席率(全体)', mc.attRate, mp.attRate, '%') +
      cmpCard('合格者数(全体)', mc.pass, mp.pass, '人') +
    '</div>' +
    '<p class="note">「pt」は割合の差(ポイント)です。合格率は出席者中、出席率は対象者全体に対する割合です。</p>' +
    '<div class="compare-charts">' +
      '<div class="chart-wrap"><h3>合格率の比較</h3><div class="chart-box"><canvas id="chart-compare-pass"></canvas></div></div>' +
      '<div class="chart-wrap"><h3>出席率の比較</h3><div class="chart-box"><canvas id="chart-compare-att"></canvas></div></div>' +
    '</div>' +
    '<h3>級ごとの比較表</h3>' +
    '<div class="table-wrap"><table class="compare-table"><thead>' +
      '<tr><th class="left" rowspan="2">区分</th>' +
      '<th colspan="3" class="grp">合格率(出席者中)</th>' +
      '<th colspan="3" class="grp">出席率(対象者中)</th>' +
      '<th colspan="3" class="grp">合格者数</th></tr>' +
      '<tr><th class="grp">今回</th><th>比較回</th><th>差</th>' +
      '<th class="grp">今回</th><th>比較回</th><th>差</th>' +
      '<th class="grp">今回</th><th>比較回</th><th>差</th></tr>' +
    '</thead><tbody>' + rowsHtml + '</tbody></table></div>';

  // 棒グラフ(全体 + 各級)
  const labels = ['全体'].concat(grades);
  const series = key => scopes.map(([, cr]) => (cr.length ? metrics(cr)[key] : null));
  const seriesPrev = key => scopes.map(([, , pr]) => (pr.length ? metrics(pr)[key] : null));
  compareChart('chart-compare-pass', labels, series('passRate'), seriesPrev('passRate'));
  compareChart('chart-compare-att', labels, series('attRate'), seriesPrev('attRate'));
}

$('compare-select').addEventListener('change', render);

// ===== Step 9-B: 学生別の成績 =====
const STATUS_LABEL = { pass: '合格', fail: '不合格', absent: '欠席' };
const STUDENT_MAX = 100;

function statusOf(r) {
  if (r.attendance === '欠席') return 'absent';
  return r.result === '合格' ? 'pass' : 'fail';
}

// 検索用に正規化(半角カナ→全角、ひらがな→カタカナ、空白を除去)
function normKana(s) {
  return String(s || '')
    .normalize('NFKC')
    .replace(/[ぁ-ゖ]/g, c => String.fromCharCode(c.charCodeAt(0) + 0x60))
    .replace(/[\s\u3000]+/g, '');
}

function ensureStudentBox() {
  let box = $('student-box');
  if (!box) {
    box = document.createElement('section');
    box.id = 'student-box';
    box.className = 'card';
    box.innerHTML =
      '<h2>学生別の成績<span class="h2-sub">フリガナで検索できます</span></h2>' +
      '<div class="filters">' +
        '<label>フリガナ(姓・名の一部)' +
          '<input type="text" id="student-q" placeholder="例:レイチャー / リック / れいちゃーりっく" autocomplete="off"></label>' +
        '<label>回<select id="student-round"></select></label>' +
        '<label>級<select id="student-grade"></select></label>' +
        '<label>判定<select id="student-judge">' +
          '<option value="">すべて</option><option value="pass">合格</option>' +
          '<option value="fail">不合格(出席)</option><option value="absent">欠席</option></select></label>' +
      '</div>' +
      '<div id="student-area"></div>';
    $('grade-section').insertAdjacentElement('afterend', box);
    $('student-q').addEventListener('input', renderStudents);
    ['student-round', 'student-grade', 'student-judge'].forEach(id =>
      $(id).addEventListener('change', renderStudents));
  }
  return box;
}

function fillStudentFilters() {
  ensureStudentBox();
  const keepRound = $('student-round').value;
  const keepGrade = $('student-grade').value;

  const map = {};
  allRows.forEach(r => { map[r.examDate] = r.round; });
  const dates = Object.keys(map).sort().reverse();
  const grades = [...new Set(allRows.map(r => r.grade))].sort((a, b) => gradeOrder(a) - gradeOrder(b));

  $('student-round').innerHTML = '<option value="">すべての回</option>' +
    dates.map(d => `<option value="${esc(d)}">${esc(map[d])}(${esc(d)})</option>`).join('');
  $('student-grade').innerHTML = '<option value="">すべての級</option>' +
    grades.map(g => `<option value="${esc(g)}">${esc(g)}</option>`).join('');

  if (dates.includes(keepRound)) $('student-round').value = keepRound;
  if (grades.includes(keepGrade)) $('student-grade').value = keepGrade;
  renderStudents();
}

function renderStudents() {
  const q = normKana($('student-q').value);
  const rd = $('student-round').value;
  const gd = $('student-grade').value;
  const jd = $('student-judge').value;

  const list = allRows.filter(r => {
    if (rd && r.examDate !== rd) return false;
    if (gd && r.grade !== gd) return false;
    if (jd && statusOf(r) !== jd) return false;
    if (q) {
      const sei = normKana(r.kanaSei);
      const mei = normKana(r.kanaMei);
      if (!(sei + mei).includes(q) && !(mei + sei).includes(q)) return false;
    }
    return true;
  }).sort((x, y) =>
    String(x.kanaSei).localeCompare(String(y.kanaSei), 'ja') ||
    String(x.kanaMei).localeCompare(String(y.kanaMei), 'ja') ||
    y.examDate.localeCompare(x.examDate) ||
    gradeOrder(x.grade) - gradeOrder(y.grade)
  );

  const missing = allRows.filter(r => !r.kanaSei && !r.kanaMei).length;
  let html = `<p class="note">該当 <strong>${list.length}</strong>件` +
    (list.length > STUDENT_MAX ? `(先頭の${STUDENT_MAX}件を表示しています。条件を絞り込んでください)` : '') + '</p>';
  if (missing) {
    html += `<p class="note">フリガナが未登録のデータが${missing}件あります。該当するExcelを保存し直すと入ります。</p>`;
  }
  if (!list.length) {
    $('student-area').innerHTML = html + '<p class="note">条件に合う学生がいません。</p>';
    return;
  }

  const rowsHtml = list.slice(0, STUDENT_MAX).map(r => {
    const s = statusOf(r);
    return `<tr><th scope="row">${esc(r.kanaSei)}</th><td>${esc(r.kanaMei)}</td>` +
      `<td>${esc(r.round)}</td><td>${esc(r.grade)}</td><td>${esc(r.attendance)}</td>` +
      `<td class="num">${s === 'absent' ? '-' : esc(r.score)}</td>` +
      `<td class="st ${s}">${STATUS_LABEL[s]}</td></tr>`;
  }).join('');

  $('student-area').innerHTML = html +
    '<div class="table-wrap"><table class="data-table student-table">' +
    '<tr><th>フリガナ(姓)</th><th>フリガナ(名)</th><th>回</th><th>級</th><th>出欠</th>' +
    '<th class="num">点数</th><th class="st">判定</th></tr>' + rowsHtml + '</table></div>';
}