// ─── Constants ────────────────────────────────────────────────
const QUARTERS = ['1-четверть', '2-четверть', '3-четверть', '4-четверть'];
const PAL = [
  '#f5d020','#3b82f6','#818cf8','#34d399','#fb923c',
  '#38bdf8','#a78bfa','#fbbf24','#60a5fa','#4ade80','#f472b6','#94a3b8'
];

const AWARD_COLS = [
  { key:'school', label:'Школьная грамота',        cls:'aw-school', col:6  },
  { key:'umc',    label:'УМЦ',                      cls:'aw-umc',    col:7  },
  { key:'daryn',  label:'РНПЦ Дарын',               cls:'aw-daryn',  col:8  },
  { key:'sardar', label:'Сарыарқа дарыны',          cls:'aw-sar',    col:9  },
  { key:'obl_q',  label:'Алғыс хаты ГОРОНО',       cls:'aw-oblq',   col:10 },
  { key:'obl_g',  label:'Почётная грамота ГОРОНО',  cls:'aw-oblg',   col:11 },
  { key:'min_q',  label:'Алғыс хаты Министерства', cls:'aw-minq',   col:12 },
  { key:'min_g',  label:'Почётная грамота МОН',     cls:'aw-ming',   col:13 },
  { key:'znak1',  label:'Нагр. знак Алтынсарин',    cls:'aw-znak1',  col:14 },
  { key:'znak2',  label:'Нагр. знак Білім беру',    cls:'aw-znak2',  col:15 },
  { key:'other',  label:'Другие',                   cls:'aw-other',  col:16 },
];

// ─── State ────────────────────────────────────────────────────
const S = {
  gr: { files: [], parsed: [], charts: [] },
  aw: { file: null, teachers: [], charts: [] },
  td: { file: null, data: null, charts: [] }
};
const multi_ = { gr: true, aw: false, td: false };

// ═══════════════════════════════════════════════════════
//  NAV
// ═══════════════════════════════════════════════════════
function showPage(id, btn) {
  document.querySelectorAll('.page').forEach(p => p.classList.remove('on'));
  document.querySelectorAll('.ntab').forEach(b => b.classList.remove('on'));
  document.getElementById('page-' + id).classList.add('on');
  btn.classList.add('on');
}

function iswitch(ns, pane, btn) {
  const root = ns === 'gr' ? 'gr-results' : ns === 'aw' ? 'aw-main' : ns === 'td' ? 'td-results' : '';
  document.querySelectorAll(`#${root} .ipane`).forEach(p => p.classList.remove('on'));
  document.querySelectorAll(`#${root} .itab`).forEach(b => b.classList.remove('on'));
  document.getElementById(`${ns}-${pane}`).classList.add('on');
  btn.classList.add('on');
}

// ═══════════════════════════════════════════════════════
//  FILE HANDLING
// ═══════════════════════════════════════════════════════
function setupDrop(ns) {
  const drop = document.getElementById(`${ns}-drop`);
  const inp  = document.getElementById(`${ns}-input`);
  drop.addEventListener('dragover', e => { e.preventDefault(); drop.classList.add('over'); });
  drop.addEventListener('dragleave', () => drop.classList.remove('over'));
  drop.addEventListener('drop', e => { e.preventDefault(); drop.classList.remove('over'); addFiles(ns, [...e.dataTransfer.files]); });
  inp.addEventListener('change', () => addFiles(ns, [...inp.files]));
}

function addFiles(ns, files) {
  files.forEach(f => {
    if (!multi_[ns] && S[ns].files && S[ns].files.find(x => x.name === f.name)) return;
    const reader = new FileReader();
    reader.onload = e => {
      if (ns === 'aw') {
        S.aw.file = { name: f.name, data: e.target.result };
      } else {
        S[ns].files.push({ name: f.name, data: e.target.result });
      }
      renderPills(ns);
      document.getElementById(`${ns}-btn`).disabled = false;
    };
    reader.readAsArrayBuffer(f);
  });
}

function removeFile(ns, name) {
  if (ns === 'aw') {
    S.aw.file = null;
    document.getElementById('aw-btn').disabled = true;
  } else {
    S[ns].files = S[ns].files.filter(f => f.name !== name);
    document.getElementById(`${ns}-btn`).disabled = S[ns].files.length === 0;
  }
  renderPills(ns);
}

function renderPills(ns) {
  const files = ns === 'aw' ? (S.aw.file ? [S.aw.file] : []) : S[ns].files;
  document.getElementById(`${ns}-pills`).innerHTML = files.map(f =>
    `<div class="pill"><span>📄 ${f.name}</span><span class="x" onclick="removeFile('${ns}','${f.name.replace(/'/g, "\\'")}')">×</span></div>`
  ).join('');
}

// ═══════════════════════════════════════════════════════
//  SHARED HELPERS
// ═══════════════════════════════════════════════════════
function qc(q)    { return q < 50 ? 'var(--red)' : q < 70 ? 'var(--yel)' : 'var(--blu)'; }
function badge(q) {
  if (q < 50) return '<span class="badge b-bad">Проблема</span>';
  if (q < 70) return '<span class="badge b-warn">Внимание</span>';
  return '<span class="badge b-ok">Норма</span>';
}
function card(lbl, val, cls, sub = '') {
  return `<div class="card"><div class="lbl">${lbl}</div><div class="val ${cls}">${val}</div>${sub ? `<div class="sub">${sub}</div>` : ''}</div>`;
}
function mkChart(id, cfg, ns) {
  const el = document.getElementById(id);
  if (!el) return;
  const c = new Chart(el, cfg);
  S[ns].charts.push(c);
  return c;
}
function destroyCharts(ns) {
  S[ns].charts.forEach(c => { try { c.destroy(); } catch (e) {} });
  S[ns].charts = [];
}
function hasValue(v) {
  if (v === null || v === undefined) return false;
  const s = String(v).trim().toLowerCase();
  return s !== '' && s !== 'nan' && s !== 'жоқ' && s !== 'жок' && s !== 'жоқ.' && s !== '-' && s !== 'нет';
}
function countYears(v) {
  if (!hasValue(v)) return 0;
  const s = String(v).replace(/[^\d,. \/]/g, '');
  const years = s.split(/[,. \/]+/).map(y => parseInt(y)).filter(y => y > 2000 && y < 2100);
  return years.length || 1;
}
function annualOrLast(st, s) {
  return st[s]?.['Годовая оценка'] || st[s]?.[QUARTERS[2]] || st[s]?.[QUARTERS[1]] || st[s]?.[QUARTERS[0]];
}

// ═══════════════════════════════════════════════════════
//  GRADES — parse & calculate
// ═══════════════════════════════════════════════════════
function parseGradeFile(name, buffer) {
  const wb  = XLSX.read(buffer, { type: 'array' });
  const ws  = wb.Sheets[wb.SheetNames[0]];
  const raw = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });
  if (!raw.length) return null;

  const headers = raw[0].map(h => h ? String(h).trim() : '');
  const PIDX = headers.findIndex(h => h.includes('Период') || h.includes('Наименование'));
  const NIDX = headers.findIndex(h => h.includes('Фамилия') || h.includes('имя'));
  if (PIDX < 0 || NIDX < 0) return null;

  const subjects = headers.slice(PIDX + 1).filter(h => h.length > 0);
  const rows = [];
  let cur = null;

  for (let i = 1; i < raw.length; i++) {
    const r = raw[i];
    if (r[0] != null && r[NIDX]) cur = String(r[NIDX]).trim();
    const period = r[PIDX] ? String(r[PIDX]).trim() : null;
    if (!period || !cur) continue;
    const row = { student: cur, quarter: period };
    subjects.forEach((s, si) => {
      const v = r[PIDX + 1 + si];
      row[s] = (v !== null && v !== undefined && v !== '') ? +v : NaN;
    });
    rows.push(row);
  }
  return { name: name.replace(/\.(xls|xlsx)$/i, '').replace(/_/g, ' ').trim(), rows, subjects };
}

function calcGradeStats(rows, subjects) {
  const st = {};
  subjects.forEach(s => {
    st[s] = {};
    [...QUARTERS, 'Годовая оценка'].forEach(q => {
      const vals = rows.filter(r => r.quarter === q && !isNaN(r[s])).map(r => r[s]);
      if (!vals.length) return;
      const mean    = vals.reduce((a, b) => a + b, 0) / vals.length;
      const quality = vals.filter(v => v >= 4).length / vals.length * 100;
      const dist    = { 2: 0, 3: 0, 4: 0, 5: 0 };
      vals.forEach(v => { if (dist[v] !== undefined) dist[v]++; });
      st[s][q] = { mean, quality, n: vals.length, dist, vals };
    });
  });
  return st;
}

// ═══════════════════════════════════════════════════════
//  GRADES — render
// ═══════════════════════════════════════════════════════
function runGrades() {
  S.gr.parsed = S.gr.files.map(f => parseGradeFile(f.name, f.data)).filter(Boolean);
  S.gr.parsed.forEach(p => { p.stats = calcGradeStats(p.rows, p.subjects); });
  destroyCharts('gr');
  grOverview(); grSubjects(); grDynamics(); grProblems(); grCompare();
  document.getElementById('gr-results').style.display = 'block';
  document.getElementById('gr-results').scrollIntoView({ behavior: 'smooth' });
}

function grOverview() {
  const C = S.gr.parsed;
  const allS = [...new Set(C.flatMap(c => c.subjects))];
  let students = 0, qS = 0, qC = 0, mS = 0, mC = 0, probs = 0, warns = 0;
  C.forEach(cls => {
    students += new Set(cls.rows.map(r => r.student)).size;
    allS.forEach(s => {
      const d = annualOrLast(cls.stats, s);
      if (!d) return;
      qS += d.quality; qC++; mS += d.mean; mC++;
      if (d.quality < 50) probs++; else if (d.quality < 70) warns++;
    });
  });
  const aq = qC ? qS / qC : 0, am = mC ? mS / mC : 0;
  const el = document.getElementById('gr-overview');
  el.innerHTML = `
    <div class="cards">
      ${card('Классов', C.length, '')}
      ${card('Учеников', students, '')}
      ${card('Средний балл', am.toFixed(2), am >= 4 ? 'g' : am >= 3.5 ? 'y' : 'r')}
      ${card('% качества', aq.toFixed(1) + '%', aq >= 70 ? 'g' : aq >= 50 ? 'y' : 'r')}
      ${card('Проблем', '<span class="r">' + probs + '</span>', '', '&lt;50%')}
      ${card('Внимание', '<span class="y">' + warns + '</span>', '', '50–70%')}
    </div>
    <div class="cgrid">
      <div class="ccrd"><h3>% качества по классам</h3><canvas id="c_gr_q"></canvas></div>
      <div class="ccrd"><h3>Распределение оценок</h3><canvas id="c_gr_d"></canvas></div>
    </div>`;

  const clsQ = C.map(cls => {
    const vs = allS.map(s => annualOrLast(cls.stats, s)?.quality).filter(v => v !== undefined);
    return vs.length ? +(vs.reduce((a, b) => a + b, 0) / vs.length).toFixed(1) : 0;
  });
  mkChart('c_gr_q', {
    type: 'bar',
    data: { labels: C.map(c => c.name), datasets: [{ data: clsQ, backgroundColor: clsQ.map(v => v < 50 ? '#f85149' : v < 70 ? '#f5d020' : '#3b82f6'), borderRadius: 6 }] },
    options: { plugins: { legend: { display: false } }, scales: { y: { max: 100, ticks: { color: '#7a90b0', callback: v => v + '%' }, grid: { color: '#1e2d45' } }, x: { ticks: { color: '#e2eaf8' }, grid: { display: false } } } }
  }, 'gr');

  const dist = { 2: 0, 3: 0, 4: 0, 5: 0 };
  C.forEach(cls => cls.rows.filter(r => QUARTERS.includes(r.quarter)).forEach(r =>
    cls.subjects.forEach(s => { const v = r[s]; if (!isNaN(v) && dist[v] !== undefined) dist[v]++; })
  ));
  const tot = Object.values(dist).reduce((a, b) => a + b, 0) || 1;
  mkChart('c_gr_d', {
    type: 'doughnut',
    data: { labels: ['2', '3', '4', '5'], datasets: [{ data: [2,3,4,5].map(g => (dist[g] / tot * 100).toFixed(1)), backgroundColor: ['#f85149','#f5d020','#3b82f6','#34d399'], borderColor: '#080c14', borderWidth: 3 }] },
    options: { plugins: { legend: { position: 'bottom', labels: { color: '#e2eaf8', padding: 14 } } } }
  }, 'gr');
}

function grSubjects() {
  const C = S.gr.parsed;
  const allS = [...new Set(C.flatMap(c => c.subjects))];
  let rows = '';
  allS.forEach(s => {
    let qS = 0, mS = 0, cnt = 0;
    const ns = [];
    C.forEach(cls => {
      const d = annualOrLast(cls.stats, s);
      if (!d) return;
      qS += d.quality; mS += d.mean; cnt++; ns.push(cls.name);
    });
    if (!cnt) return;
    const q = qS / cnt, m = mS / cnt, bc = qc(q);
    rows += `<tr>
      <td><strong>${s}</strong></td>
      <td style="color:var(--mut);font-size:12px">${ns.join(', ')}</td>
      <td><strong style="color:${bc}">${m.toFixed(2)}</strong></td>
      <td><div class="qbar"><div class="qbg"><div class="qfill" style="width:${q}%;background:${bc}"></div></div><span class="qnum" style="color:${bc}">${q.toFixed(0)}%</span></div></td>
      <td>${badge(q)}</td>
    </tr>`;
  });
  document.getElementById('gr-subjects').innerHTML = `
    <div class="stitle">Все предметы</div>
    <div class="tbl-wrap"><table class="dt">
      <thead><tr><th>Предмет</th><th>Классы</th><th>Ср. балл</th><th>% качества</th><th>Статус</th></tr></thead>
      <tbody>${rows}</tbody>
    </table></div>`;
}

function grDynamics() {
  const C = S.gr.parsed;
  document.getElementById('gr-dynamics').innerHTML = `
    <div class="cgrid one"><div class="ccrd"><h3>% качества по четвертям</h3><canvas id="c_gr_dq"></canvas></div></div>
    <div class="cgrid one" style="margin-top:16px"><div class="ccrd"><h3>Средний балл по четвертям</h3><canvas id="c_gr_dm"></canvas></div></div>`;

  const dsQ = C.map((cls, i) => ({
    label: cls.name,
    data: QUARTERS.map(q => { const vs = cls.subjects.map(s => cls.stats[s]?.[q]?.quality).filter(v => v !== undefined); return vs.length ? +(vs.reduce((a, b) => a + b, 0) / vs.length).toFixed(1) : null; }),
    borderColor: PAL[i], backgroundColor: PAL[i] + '22', fill: true, tension: .35, pointRadius: 5, borderWidth: 2.5
  }));
  mkChart('c_gr_dq', { type: 'line', data: { labels: QUARTERS, datasets: dsQ }, options: { scales: { y: { min: 0, max: 100, ticks: { color: '#7a90b0', callback: v => v + '%' }, grid: { color: '#1e2d45' } }, x: { ticks: { color: '#7a90b0' }, grid: { color: '#1e2d45' } } }, plugins: { legend: { labels: { color: '#e2eaf8' } } } } }, 'gr');

  const dsM = C.map((cls, i) => ({
    label: cls.name,
    data: QUARTERS.map(q => { const vs = cls.subjects.map(s => cls.stats[s]?.[q]?.mean).filter(v => v !== undefined); return vs.length ? +(vs.reduce((a, b) => a + b, 0) / vs.length).toFixed(2) : null; }),
    borderColor: PAL[i], backgroundColor: PAL[i] + '22', fill: true, tension: .35, pointRadius: 5, borderWidth: 2.5
  }));
  mkChart('c_gr_dm', { type: 'line', data: { labels: QUARTERS, datasets: dsM }, options: { scales: { y: { min: 1.5, max: 5.5, ticks: { color: '#7a90b0' }, grid: { color: '#1e2d45' } }, x: { ticks: { color: '#7a90b0' }, grid: { color: '#1e2d45' } } }, plugins: { legend: { labels: { color: '#e2eaf8' } } } } }, 'gr');
}

function grProblems() {
  const issues = [];
  S.gr.parsed.forEach(cls => cls.subjects.forEach(s => {
    const d = annualOrLast(cls.stats, s);
    if (!d || d.quality >= 70) return;
    issues.push({ subject: s, cls: cls.name, quality: d.quality, mean: d.mean, level: d.quality < 50 ? 'bad' : 'warn' });
  }));
  issues.sort((a, b) => a.quality - b.quality);
  const el = document.getElementById('gr-problems');
  if (!issues.length) { el.innerHTML = '<div style="color:var(--grn);padding:24px">✅ Все предметы в норме!</div>'; return; }
  el.innerHTML = `<div class="pgrid">${issues.map(i => `
    <div class="pcrd ${i.level}">
      <div class="ps">${i.subject}</div>
      <div style="color:var(--mut);font-size:12px;margin-bottom:8px">📚 ${i.cls}</div>
      <div class="pq" style="color:${qc(i.quality)}">${i.quality.toFixed(0)}%</div>
      <div style="font-size:12px;color:var(--mut)">ср. балл ${i.mean.toFixed(2)}</div>
      <div class="pr">${i.level === 'bad' ? '⚠️ Срочно: доп. занятия, индив. работа.' : '🔍 Усилить контроль, провести срезы.'}</div>
    </div>`).join('')}</div>`;
}

function grCompare() {
  const C = S.gr.parsed;
  if (C.length < 2) { document.getElementById('gr-compare').innerHTML = '<div style="color:var(--mut);padding:40px">Загрузите минимум 2 класса.</div>'; return; }
  const common = C[0].subjects.filter(s => C.every(c => c.subjects.includes(s)));
  document.getElementById('gr-compare').innerHTML = `
    <div class="cgrid">
      <div class="ccrd"><h3>% качества по предметам</h3><canvas id="c_gr_cs"></canvas></div>
      <div class="ccrd"><h3>Радар</h3><div style="max-width:380px;margin:0 auto"><canvas id="c_gr_cr"></canvas></div></div>
    </div>`;
  mkChart('c_gr_cs', { type: 'bar', data: { labels: common, datasets: C.map((cls, i) => ({ label: cls.name, data: common.map(s => { const d = annualOrLast(cls.stats, s); return d ? +d.quality.toFixed(1) : 0; }), backgroundColor: PAL[i] + 'cc', borderColor: PAL[i], borderWidth: 1, borderRadius: 4 })) }, options: { indexAxis: 'y', scales: { x: { max: 100, ticks: { color: '#7a90b0', callback: v => v + '%' }, grid: { color: '#1e2d45' } }, y: { ticks: { color: '#e2eaf8', font: { size: 10 } }, grid: { display: false } } }, plugins: { legend: { labels: { color: '#e2eaf8' } } } } }, 'gr');
  mkChart('c_gr_cr', { type: 'radar', data: { labels: common, datasets: C.map((cls, i) => ({ label: cls.name, data: common.map(s => { const d = annualOrLast(cls.stats, s); return d ? +d.quality.toFixed(1) : 0; }), borderColor: PAL[i], backgroundColor: PAL[i] + '33', pointBackgroundColor: PAL[i], borderWidth: 2 })) }, options: { scales: { r: { min: 0, max: 100, ticks: { color: '#7a90b0', backdropColor: 'transparent', stepSize: 25 }, grid: { color: '#1e2d45' }, pointLabels: { color: '#e2eaf8', font: { size: 9 } } } }, plugins: { legend: { labels: { color: '#e2eaf8' } } } } }, 'gr');
}

// ═══════════════════════════════════════════════════════
//  AWARDS — parse
// ═══════════════════════════════════════════════════════
function parseAwardsFile(buffer) {
  const wb  = XLSX.read(buffer, { type: 'array' });
  const ws  = wb.Sheets[wb.SheetNames[0]];
  const raw = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });
  const teachers = [];

  for (let i = 2; i < raw.length; i++) {
    const r    = raw[i];
    const num  = r[0];
    const name = r[1] ? String(r[1]).trim() : null;
    if (!name || !num) continue;

    const stazh     = parseFloat(r[2]) || 0;
    const stazh_org = r[5] ? parseFloat(String(r[5])) || 0 : 0;

    const awards = {};
    AWARD_COLS.forEach(ac => {
      const v = r[ac.col];
      awards[ac.key] = hasValue(v) ? String(v).trim() : null;
    });

    const totalTypes = AWARD_COLS.filter(ac => awards[ac.key] !== null).length;
    const totalCount = AWARD_COLS.reduce((sum, ac) => sum + (awards[ac.key] ? countYears(awards[ac.key]) : 0), 0);

    let level = 'none';
    if (awards.znak2 || awards.znak1)             level = 'знак';
    else if (awards.min_g)                         level = 'МОН грамота';
    else if (awards.min_q)                         level = 'МОН алғыс';
    else if (awards.obl_g)                         level = 'ГОРОНО грамота';
    else if (awards.obl_q)                         level = 'ГОРОНО алғыс';
    else if (awards.sardar || awards.daryn || awards.umc) level = 'республиканские';
    else if (awards.school)                        level = 'школьная';

    teachers.push({ num: parseInt(num), name, stazh, stazh_org, awards, totalTypes, totalCount, level });
  }
  return teachers;
}

// ═══════════════════════════════════════════════════════
//  AWARDS — render
// ═══════════════════════════════════════════════════════
function runAwards() {
  if (!S.aw.file) return;
  S.aw.teachers = parseAwardsFile(S.aw.file.data);
  destroyCharts('aw');
  awOverview(); awStats(); awStazh(); awTop();
  document.getElementById('aw-results').style.display = 'block';
  document.getElementById('aw-main').style.display    = 'block';
  document.getElementById('aw-drill').style.display   = 'none';
  document.getElementById('aw-results').scrollIntoView({ behavior: 'smooth' });
}

function awOverview() {
  const T          = S.aw.teachers;
  const total      = T.length;
  const withAwards = T.filter(t => t.totalTypes > 0).length;
  const withSign   = T.filter(t => t.awards.znak1 || t.awards.znak2).length;
  const withMin    = T.filter(t => t.awards.min_g || t.awards.min_q).length;
  const avgStazh   = T.reduce((a, b) => a + b.stazh, 0) / total;

  const el = document.getElementById('aw-overview');
  el.innerHTML = `
    <div class="cards">
      ${card('Педагогов', total, '')}
      ${card('Имеют награды', withAwards, 'g', Math.round(withAwards / total * 100) + '% состава')}
      ${card('Ср. стаж', avgStazh.toFixed(1) + ' лет', 'b', '')}
      ${card('Нагр. знаки', withSign, 'p', '')}
      ${card('Награды МОН', withMin, 'y', '')}
      ${card('Без наград', total - withAwards, 'r', Math.round((total - withAwards) / total * 100) + '% состава')}
    </div>
    <div class="stitle">Все педагоги — нажмите для подробной информации</div>
    <div class="tch-grid" id="aw-cards"></div>`;

  const grid = document.getElementById('aw-cards');
  T.forEach((t, i) => {
    const chipHtml = AWARD_COLS.filter(ac => t.awards[ac.key])
      .map(ac => `<span class="award-chip ${ac.cls}">${ac.label.length > 20 ? ac.label.slice(0, 18) + '…' : ac.label}</span>`)
      .join('');
    const div = document.createElement('div');
    div.className = 'tch-card';
    div.style.borderLeftColor = t.totalTypes === 0 ? 'var(--brd)' : t.totalTypes >= 5 ? 'var(--acc)' : t.totalTypes >= 3 ? 'var(--grn)' : 'var(--blu)';
    div.innerHTML = `
      <div class="tname">${t.name}</div>
      <div class="tmeta">Стаж: ${t.stazh} лет · в организации: ${t.stazh_org || '—'} лет</div>
      <div class="trow"><span>Типов наград</span><strong>${t.totalTypes}</strong></div>
      <div class="trow"><span>Всего грамот/знаков</span><strong>${t.totalCount}</strong></div>
      <div class="trow"><span>Высшая награда</span><strong style="color:var(--acc)">${t.level === 'none' ? '—' : t.level}</strong></div>
      <div class="taw">${chipHtml || '<span style="color:var(--mut);font-size:12px">Наград нет</span>'}</div>`;
    div.onclick = () => awDrill(i);
    grid.appendChild(div);
  });
}

function awStats() {
  const T = S.aw.teachers;
  document.getElementById('aw-stats').innerHTML = `
    <div class="cgrid">
      <div class="ccrd"><h3>Охват по типам наград</h3><canvas id="c_aw_types"></canvas></div>
      <div class="ccrd"><h3>Педагоги по количеству типов наград</h3><canvas id="c_aw_cnt"></canvas></div>
    </div>
    <div class="cgrid one" style="margin-top:16px">
      <div class="ccrd"><h3>Все типы наград — сколько педагогов имеют</h3><canvas id="c_aw_bar"></canvas></div>
    </div>`;

  const typeCounts = AWARD_COLS
    .map(ac => ({ label: ac.label, count: T.filter(t => t.awards[ac.key]).length }))
    .sort((a, b) => b.count - a.count);

  mkChart('c_aw_bar', {
    type: 'bar',
    data: { labels: typeCounts.map(t => t.label), datasets: [{ data: typeCounts.map(t => t.count), backgroundColor: typeCounts.map((_, i) => PAL[i % PAL.length]), borderRadius: 6 }] },
    options: { indexAxis: 'y', plugins: { legend: { display: false } }, scales: { x: { ticks: { color: '#7a90b0' }, grid: { color: '#1e2d45' } }, y: { ticks: { color: '#e2eaf8', font: { size: 11 } }, grid: { display: false } } } }
  }, 'aw');

  const withAw = T.filter(t => t.totalTypes > 0).length;
  mkChart('c_aw_types', {
    type: 'doughnut',
    data: { labels: ['Имеют награды', 'Без наград'], datasets: [{ data: [withAw, T.length - withAw], backgroundColor: ['#3b82f6', '#1e2d45'], borderColor: '#080c14', borderWidth: 3 }] },
    options: { plugins: { legend: { position: 'bottom', labels: { color: '#e2eaf8', padding: 14 } } } }
  }, 'aw');

  const dist = {};
  T.forEach(t => { dist[t.totalTypes] = (dist[t.totalTypes] || 0) + 1; });
  const maxT     = Math.max(...Object.keys(dist).map(Number));
  const distData = Array.from({ length: maxT + 1 }, (_, i) => dist[i] || 0);
  mkChart('c_aw_cnt', {
    type: 'bar',
    data: { labels: distData.map((_, i) => i === 0 ? '0 наград' : i + ' тип(ов)'), datasets: [{ data: distData, backgroundColor: distData.map((_, i) => i === 0 ? '#f85149' : i < 3 ? '#f5d020' : '#3b82f6'), borderRadius: 6 }] },
    options: { plugins: { legend: { display: false } }, scales: { y: { ticks: { color: '#7a90b0' }, grid: { color: '#1e2d45' } }, x: { ticks: { color: '#7a90b0' }, grid: { display: false } } } }
  }, 'aw');
}

function awStazh() {
  const T        = S.aw.teachers;
  const brackets = [
    { l: '0–5',   min: 0,  max: 5   },
    { l: '6–10',  min: 6,  max: 10  },
    { l: '11–15', min: 11, max: 15  },
    { l: '16–20', min: 16, max: 20  },
    { l: '21–30', min: 21, max: 30  },
    { l: '31+',   min: 31, max: 999 }
  ];
  const groups = brackets.map(b => ({ ...b, teachers: T.filter(t => t.stazh >= b.min && t.stazh <= b.max) }));

  document.getElementById('aw-stazh').innerHTML = `
    <div class="cgrid">
      <div class="ccrd"><h3>Состав по стажу</h3><canvas id="c_aw_stazh"></canvas></div>
      <div class="ccrd"><h3>Ср. кол-во наград по группам стажа</h3><canvas id="c_aw_stavg"></canvas></div>
    </div>
    <div class="stitle" style="margin-top:24px">Педагоги по группам стажа</div>
    ${groups.filter(g => g.teachers.length).map(g => `
      <div style="margin-bottom:20px">
        <div style="font-family:'Unbounded',sans-serif;font-size:12px;color:var(--acc);margin-bottom:10px">
          Стаж ${g.l} лет — ${g.teachers.length} чел.
        </div>
        <div class="tbl-wrap"><table class="dt">
          <thead><tr><th>ФИО</th><th>Стаж</th><th>Наград</th><th>Высшая</th></tr></thead>
          <tbody>${g.teachers.map(t => `<tr>
            <td style="cursor:pointer;color:var(--blu)" onclick="awDrill(${S.aw.teachers.indexOf(t)})">${t.name}</td>
            <td>${t.stazh}</td>
            <td><strong>${t.totalTypes}</strong></td>
            <td style="color:var(--acc)">${t.level === 'none' ? '—' : t.level}</td>
          </tr>`).join('')}</tbody>
        </table></div>
      </div>`).join('')}`;

  mkChart('c_aw_stazh', { type: 'bar', data: { labels: groups.map(g => g.l), datasets: [{ data: groups.map(g => g.teachers.length), backgroundColor: PAL, borderRadius: 6 }] }, options: { plugins: { legend: { display: false } }, scales: { y: { ticks: { color: '#7a90b0' }, grid: { color: '#1e2d45' } }, x: { ticks: { color: '#7a90b0' }, grid: { display: false } } } } }, 'aw');
  mkChart('c_aw_stavg', { type: 'bar', data: { labels: groups.map(g => g.l), datasets: [{ label: 'Ср. типов наград', data: groups.map(g => g.teachers.length ? +(g.teachers.reduce((a, b) => a + b.totalTypes, 0) / g.teachers.length).toFixed(1) : 0), backgroundColor: '#3b82f6', borderRadius: 6 }] }, options: { plugins: { legend: { display: false } }, scales: { y: { ticks: { color: '#7a90b0' }, grid: { color: '#1e2d45' } }, x: { ticks: { color: '#7a90b0' }, grid: { display: false } } } } }, 'aw');
}

function awTop() {
  const T    = [...S.aw.teachers].sort((a, b) => b.totalTypes - a.totalTypes || b.totalCount - a.totalCount);
  let rows = '';
  T.forEach((t, i) => {
    const medal    = i === 0 ? '🥇' : i === 1 ? '🥈' : i === 2 ? '🥉' : `<span style="color:var(--mut)">${i + 1}</span>`;
    const chipHtml = AWARD_COLS.filter(ac => t.awards[ac.key])
      .map(ac => `<span class="award-chip ${ac.cls}" title="${t.awards[ac.key]}">${ac.label.length > 16 ? ac.label.slice(0, 14) + '…' : ac.label}</span>`)
      .join('');
    rows += `<tr style="cursor:pointer" onclick="awDrill(${S.aw.teachers.indexOf(t)})">
      <td style="font-size:15px">${medal}</td>
      <td><strong style="color:var(--blu)">${t.name}</strong></td>
      <td>${t.stazh}</td>
      <td style="font-weight:700;color:var(--acc)">${t.totalTypes}</td>
      <td>${t.totalCount}</td>
      <td style="color:var(--acc);font-size:12px">${t.level === 'none' ? '—' : t.level}</td>
      <td>${chipHtml}</td>
    </tr>`;
  });
  document.getElementById('aw-top').innerHTML = `
    <div class="stitle">Рейтинг педагогов по наградам</div>
    <div class="tbl-wrap"><table class="dt">
      <thead><tr><th>#</th><th>ФИО</th><th>Стаж</th><th>Типов наград</th><th>Всего</th><th>Высшая</th><th>Награды</th></tr></thead>
      <tbody>${rows}</tbody>
    </table></div>`;
}

function awDrill(idx) {
  const t = S.aw.teachers[idx];
  document.getElementById('aw-drill').style.display = 'block';
  document.getElementById('aw-main').style.display  = 'none';

  const awardRows = AWARD_COLS.map(ac => {
    const v = t.awards[ac.key];
    if (!v) return '';
    return `<tr>
      <td><span class="award-chip ${ac.cls}">${ac.label}</span></td>
      <td style="color:var(--txt)">${v}</td>
      <td style="color:var(--mut);font-size:12px">${countYears(v)} раз</td>
    </tr>`;
  }).join('');

  document.getElementById('aw-drill').innerHTML = `
    <button class="back-btn" onclick="awBack()">← Все педагоги</button>
    <div style="font-family:'Unbounded',sans-serif;font-size:22px;font-weight:800;margin-bottom:4px">${t.name}</div>
    <div style="color:var(--mut);font-size:13px;margin-bottom:24px">Общий стаж: ${t.stazh} лет · В организации: ${t.stazh_org || '—'} лет</div>
    <div class="cards">
      ${card('Типов наград', t.totalTypes, t.totalTypes >= 5 ? 'g' : t.totalTypes >= 2 ? 'y' : 'r')}
      ${card('Всего грамот', t.totalCount, 'b')}
      ${card('Высшая награда', t.level === 'none' ? '—' : t.level, 'p', '')}
    </div>
    ${awardRows ? `
      <div class="stitle">Полный список наград</div>
      <div class="tbl-wrap"><table class="dt">
        <thead><tr><th>Тип награды</th><th>Годы / информация</th><th>Количество</th></tr></thead>
        <tbody>${awardRows}</tbody>
      </table></div>`
    : '<div style="color:var(--mut);padding:20px">У данного педагога наград не зафиксировано.</div>'}`;
}

function awBack() {
  destroyCharts('aw');
  awOverview(); awStats(); awStazh(); awTop();
  document.getElementById('aw-drill').style.display = 'none';
  document.getElementById('aw-main').style.display  = 'block';
}

// ═══════════════════════════════════════════════════════
//  INIT
// ═══════════════════════════════════════════════════════
setupDrop('gr');
setupDrop('aw');

// ═══════════════════════════════════════════════════════
//  RATINGS — teachers & students
// ═══════════════════════════════════════════════════════

S.rt = { teacherFile: null, studentFile: null, teachers: [], students: [], charts: [] };

// Extra drop setup for ratings page
function setupRatingDrops() {
  // Teacher rating drop
  const rtDrop = document.getElementById('rt-drop');
  const rtInp  = document.getElementById('rt-input');
  rtDrop.addEventListener('dragover', e => { e.preventDefault(); rtDrop.classList.add('over'); });
  rtDrop.addEventListener('dragleave', () => rtDrop.classList.remove('over'));
  rtDrop.addEventListener('drop', e => { e.preventDefault(); rtDrop.classList.remove('over'); loadRatingFile('rt', [...e.dataTransfer.files][0]); });
  rtInp.addEventListener('change', () => loadRatingFile('rt', rtInp.files[0]));

  // Student rating drop
  const rsDrop = document.getElementById('rs-drop');
  const rsInp  = document.getElementById('rs-input');
  rsDrop.addEventListener('dragover', e => { e.preventDefault(); rsDrop.classList.add('over'); });
  rsDrop.addEventListener('dragleave', () => rsDrop.classList.remove('over'));
  rsDrop.addEventListener('drop', e => { e.preventDefault(); rsDrop.classList.remove('over'); loadRatingFile('rs', [...e.dataTransfer.files][0]); });
  rsInp.addEventListener('change', () => loadRatingFile('rs', rsInp.files[0]));
}

function loadRatingFile(ns, file) {
  if (!file) return;
  const reader = new FileReader();
  reader.onload = e => {
    if (ns === 'rt') {
      S.rt.teacherFile = { name: file.name, data: e.target.result };
      document.getElementById('rt-pills').innerHTML = `<div class="pill"><span>📄 ${file.name}</span><span class="x" onclick="S.rt.teacherFile=null;document.getElementById('rt-pills').innerHTML=''">×</span></div>`;
    } else {
      S.rt.studentFile = { name: file.name, data: e.target.result };
      document.getElementById('rs-pills').innerHTML = `<div class="pill"><span>📄 ${file.name}</span><span class="x" onclick="S.rt.studentFile=null;document.getElementById('rs-pills').innerHTML=''">×</span></div>`;
      parseStudentRatings(e.target.result);
    }
    if (S.rt.teacherFile) document.getElementById('rt-btn').disabled = false;
  };
  reader.readAsArrayBuffer(file);
}

// ── Parse teacher rating file ─────────────────────────
function parseTeacherRatings(buffer) {
  const wb  = XLSX.read(buffer, { type: 'array' });
  const ws  = wb.Sheets[wb.SheetNames[0]];
  const raw = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });

  const teachers = [];
  // Find data rows: col 33+ has name (col 34) and total (col 38)
  // Teacher data rows identified where col 33 has a name and col 38 has a score
  for (let i = 0; i < raw.length; i++) {
    const r = raw[i];
    // Check right-side summary block: col 33=num, col 34=name, col 35=Q1, col 36=Q2, col 37=Q3/Q4, col 38=total
    const name = r[34] ? String(r[34]).trim() : null;
    const total = r[38] !== null && r[38] !== undefined ? +r[38] : null;
    if (!name || !name.match(/[А-ЯЁа-яёA-Za-z]/) || total === null || isNaN(total)) continue;
    if (name === 'Аты- жөні' || name.length < 3) continue;

    const q1 = r[35] !== null && r[35] !== undefined ? +r[35] : null;
    const q2 = r[36] !== null && r[36] !== undefined ? +r[36] : null;
    const q34 = r[37] !== null && r[37] !== undefined ? +r[37] : null;

    // Also collect criteria scores from left block (cols 2–31)
    const criteriaScores = [];
    for (let c = 2; c <= 31; c++) {
      const v = r[c];
      if (v !== null && v !== undefined && !isNaN(+v)) criteriaScores.push({ col: c, score: +v });
    }

    teachers.push({ name, q1, q2, q34, total, criteriaScores });
  }
  return teachers;
}

// ── Parse student rating file ──────────────────────────
function parseStudentRatings(buffer) {
  const wb  = XLSX.read(buffer, { type: 'array' });
  const ws  = wb.Sheets[wb.SheetNames[0]];
  const raw = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });

  if (!raw.length) return;

  // Row 0 = headers (col 0=№, col 1=name, col 2..79=criteria, col 80=total, col 81=name again)
  const headers = raw[0].map(h => h ? String(h).trim() : '');

  const students = [];
  for (let i = 1; i < raw.length; i++) {
    const r = raw[i];
    const num  = r[0];
    const name = r[1] ? String(r[1]).trim() : (r[81] ? String(r[81]).trim() : null);
    if (!name || isNaN(+num)) continue;

    const total = r[80] !== null && r[80] !== undefined ? +r[80] : null;

    // Gather non-null criteria with their labels
    const scores = [];
    for (let c = 2; c <= 79; c++) {
      const v = r[c];
      if (v !== null && v !== undefined && !isNaN(+v) && +v !== 0) {
        scores.push({ label: headers[c] || `Кол ${c}`, value: +v });
      }
    }

    students.push({ num: +num, name, total: total || 0, scores });
  }

  S.rt.students = students.sort((a, b) => (b.total || 0) - (a.total || 0));
  updateStudentSelect();
}

function updateStudentSelect() {
  const sel = document.getElementById('student-select');
  sel.innerHTML = '<option value="">— выберите ученика —</option>' +
    S.rt.students.map((s, i) =>
      `<option value="${i}">${s.name} (${s.total} баллов)</option>`
    ).join('');
}

// ── Run ratings ────────────────────────────────────────
function runRatings() {
  if (S.rt.teacherFile) {
    S.rt.teachers = parseTeacherRatings(S.rt.teacherFile.data);
  }
  S.rt.charts.forEach(c => { try { c.destroy(); } catch(e) {} });
  S.rt.charts = [];

  renderTeacherRating();
  renderStudentRating();

  document.getElementById('rt-results').style.display = 'block';
  document.getElementById('rt-results').scrollIntoView({ behavior: 'smooth' });
}

// ── Teacher rating table + chart ──────────────────────
function renderTeacherRating() {
  const T = [...S.rt.teachers].sort((a, b) => (b.total || 0) - (a.total || 0));
  const max = T.length ? Math.max(...T.map(t => t.total || 0)) : 100;

  let rows = '';
  T.forEach((t, i) => {
    const medal = i === 0 ? '🥇' : i === 1 ? '🥈' : i === 2 ? '🥉' : `<span style="color:var(--mut)">${i+1}</span>`;
    const pct   = max ? (t.total / max * 100) : 0;
    const bc    = pct >= 70 ? 'var(--blu)' : pct >= 40 ? 'var(--yel)' : 'var(--red)';
    const qs = [
      t.q1  !== null ? `<span style="color:var(--mut);font-size:11px">I: </span><strong>${t.q1}</strong>` : '—',
      t.q2  !== null ? `<span style="color:var(--mut);font-size:11px">II: </span><strong>${t.q2}</strong>` : '—',
      t.q34 !== null ? `<span style="color:var(--mut);font-size:11px">III-IV: </span><strong>${t.q34}</strong>` : '—',
    ].join(' &nbsp; ');
    rows += `<tr>
      <td style="font-size:15px">${medal}</td>
      <td><strong style="color:var(--txt)">${t.name}</strong></td>
      <td style="font-size:12px">${qs}</td>
      <td>
        <div class="qbar">
          <div class="qbg"><div class="qfill" style="width:${pct}%;background:${bc}"></div></div>
          <span class="qnum" style="color:${bc}">${t.total}</span>
        </div>
      </td>
    </tr>`;
  });

  document.getElementById('rt-teachers').innerHTML = `
    <div class="cards" style="margin-bottom:20px">
      ${card('Учителей', T.length, '')}
      ${card('Лидер', T[0] ? T[0].name.split(' ')[0] : '—', 'b', T[0] ? T[0].total + ' баллов' : '')}
      ${card('Средний балл', T.length ? (T.reduce((a,b)=>a+(b.total||0),0)/T.length).toFixed(1) : '—', 'y', '')}
    </div>
    <div class="cgrid" style="margin-bottom:24px">
      <div class="ccrd" style="grid-column:1/-1"><h3>Итоговый рейтинг учителей</h3><canvas id="c_rt_bar"></canvas></div>
    </div>
    <div class="stitle">Таблица рейтинга</div>
    <div class="tbl-wrap"><table class="dt">
      <thead><tr><th>#</th><th>Учитель</th><th>По четвертям</th><th>Итого баллов</th></tr></thead>
      <tbody>${rows}</tbody>
    </table></div>`;

  const c = new Chart(document.getElementById('c_rt_bar'), {
    type: 'bar',
    data: {
      labels: T.map(t => t.name),
      datasets: [{
        data: T.map(t => t.total || 0),
        backgroundColor: T.map((t,i) => PAL[i % PAL.length]),
        borderRadius: 6
      }]
    },
    options: {
      plugins: { legend: { display: false } },
      scales: {
        y: { ticks: { color: '#7a90b0' }, grid: { color: '#1e2d45' } },
        x: { ticks: { color: '#e2eaf8', font: { size: 11 } }, grid: { display: false } }
      }
    }
  });
  S.rt.charts.push(c);
}

// ── Student rating table ───────────────────────────────
function renderStudentRating() {
  const students = S.rt.students;
  if (!students.length) {
    document.getElementById('rt-students').innerHTML =
      '<div style="color:var(--mut);padding:40px">Загрузите файл рейтинга учеников выше.</div>';
    return;
  }

  const max = Math.max(...students.map(s => s.total || 0)) || 1;
  let rows = '';
  students.forEach((s, i) => {
    const medal = i === 0 ? '🥇' : i === 1 ? '🥈' : i === 2 ? '🥉' : `<span style="color:var(--mut)">${i+1}</span>`;
    const pct = s.total / max * 100;
    const bc  = pct >= 70 ? 'var(--blu)' : pct >= 40 ? 'var(--yel)' : 'var(--red)';
    rows += `<tr style="cursor:pointer" onclick="selectStudentAndSwitch(${S.rt.students.indexOf(s)})">
      <td style="font-size:15px">${medal}</td>
      <td><strong style="color:var(--blu)">${s.name}</strong></td>
      <td>
        <div class="qbar">
          <div class="qbg"><div class="qfill" style="width:${pct}%;background:${bc}"></div></div>
          <span class="qnum" style="color:${bc}">${s.total}</span>
        </div>
      </td>
    </tr>`;
  });

  document.getElementById('rt-students').innerHTML = `
    <div class="cards" style="margin-bottom:20px">
      ${card('Учеников', students.length, '')}
      ${card('Лидер', students[0] ? students[0].name.split(' ')[0] : '—', 'b', students[0] ? students[0].total + ' баллов' : '')}
      ${card('Средний балл', (students.reduce((a,b)=>a+(b.total||0),0)/students.length).toFixed(1), 'y', '')}
    </div>
    <div class="stitle">Таблица рейтинга учеников</div>
    <div class="tbl-wrap"><table class="dt">
      <thead><tr><th>#</th><th>Ученик</th><th>Итого баллов</th></tr></thead>
      <tbody>${rows}</tbody>
    </table></div>`;
}

// ── Student card (drilldown) ───────────────────────────
function selectStudentAndSwitch(idx) {
  document.getElementById('student-select').value = String(idx);
  // Switch to student card tab
  const tabs   = document.querySelectorAll('#rt-results .itab');
  const panes  = document.querySelectorAll('#rt-results .ipane');
  tabs.forEach(b => b.classList.remove('on'));
  panes.forEach(p => p.classList.remove('on'));
  tabs[2].classList.add('on');
  document.getElementById('rt-student-card').classList.add('on');
  showStudentCard(String(idx));
}

function showStudentCard(idxStr) {
  const idx = parseInt(idxStr);
  const body = document.getElementById('student-card-body');
  if (isNaN(idx) || !S.rt.students[idx]) { body.innerHTML = ''; return; }

  const s   = S.rt.students[idx];
  const all = S.rt.students;
  const rank = all.findIndex(x => x.name === s.name) + 1;
  const max  = Math.max(...all.map(x => x.total || 0)) || 1;
  const pct  = (s.total / max * 100).toFixed(0);
  const bc   = pct >= 70 ? 'var(--blu)' : pct >= 40 ? 'var(--yel)' : 'var(--red)';

  // Positive vs negative scores
  const plus  = s.scores.filter(sc => sc.value > 0);
  const minus = s.scores.filter(sc => sc.value < 0);

  const plusRows  = plus.map(sc => `<tr><td style="font-size:12px;color:var(--txt)">${sc.label}</td><td style="color:var(--grn);font-weight:700;text-align:right">+${sc.value}</td></tr>`).join('');
  const minusRows = minus.map(sc => `<tr><td style="font-size:12px;color:var(--txt)">${sc.label}</td><td style="color:var(--red);font-weight:700;text-align:right">${sc.value}</td></tr>`).join('');

  body.innerHTML = `
    <div style="font-family:'Unbounded',sans-serif;font-size:22px;font-weight:800;margin-bottom:4px">${s.name}</div>
    <div style="color:var(--mut);font-size:13px;margin-bottom:24px">Рейтинг: #${rank} из ${all.length}</div>
    <div class="cards">
      ${card('Итоговый балл', s.total, 'b', '')}
      ${card('Место в классе', '#' + rank, rank === 1 ? 'g' : rank <= 3 ? 'y' : '', 'из ' + all.length)}
      ${card('Бонусные баллы', '+' + plus.reduce((a,sc)=>a+sc.value,0), 'g', '')}
      ${card('Штрафные баллы', minus.reduce((a,sc)=>a+sc.value,0), 'r', '')}
    </div>

    <div style="margin:20px 0 8px">
      <div class="qbar" style="max-width:500px">
        <div class="qbg" style="height:10px">
          <div class="qfill" style="width:${pct}%;background:${bc};height:10px"></div>
        </div>
        <span style="color:${bc};font-weight:700;font-size:16px;margin-left:12px">${s.total} / ${max} баллов</span>
      </div>
    </div>

    <div class="cgrid" style="margin-top:24px">
      ${plusRows ? `
      <div>
        <div class="stitle">✅ Бонусные баллы</div>
        <div class="tbl-wrap"><table class="dt">
          <thead><tr><th>Критерий</th><th style="text-align:right">Баллы</th></tr></thead>
          <tbody>${plusRows}</tbody>
        </table></div>
      </div>` : ''}
      ${minusRows ? `
      <div>
        <div class="stitle">⚠️ Штрафные баллы</div>
        <div class="tbl-wrap"><table class="dt">
          <thead><tr><th>Критерий</th><th style="text-align:right">Баллы</th></tr></thead>
          <tbody>${minusRows}</tbody>
        </table></div>
      </div>` : ''}
    </div>`;
}

// Override iswitch for ratings (different root)
const _origIswitch = iswitch;
window.iswitch = function(ns, pane, btn) {
  if (ns === 'rt') {
    document.querySelectorAll('#rt-results .ipane').forEach(p => p.classList.remove('on'));
    document.querySelectorAll('#rt-results .itab').forEach(b => b.classList.remove('on'));
    document.getElementById(`rt-${pane}`).classList.add('on');
    btn.classList.add('on');
  } else {
    _origIswitch(ns, pane, btn);
  }
};

// Init
setupRatingDrops();

// ═══════════════════════════════════════════════════════
//  ADVANCED ANALYTICS — новая расширенная аналитика
// ═══════════════════════════════════════════════════════

// State for analytics
S.an = { file: null, data: null, charts: [] };
multi_.an = false;

// Setup analytics page
setupDrop('an');

function runAnalytics() {
  if (!S.an.file) return;
  S.an.data = parseGradeFile(S.an.file.name, S.an.file.data);
  if (!S.an.data) {
    alert('Ошибка парсинга файла');
    return;
  }
  S.an.data.stats = calcGradeStats(S.an.data.rows, S.an.data.subjects);
  destroyCharts('an');
  
  anForecast();
  anDistribution();
  anCorrelations();
  anInsights();
  anExport();
  
  document.getElementById('an-results').style.display = 'block';
  document.getElementById('an-results').scrollIntoView({ behavior: 'smooth' });
}

// Прогнозирование успеваемости
function anForecast() {
  const data = S.an.data;
  if (!data) return;
  
  const el = document.getElementById('an-forecast');
  
  // Рассчитываем тренды для каждого предмета
  const trends = [];
  data.subjects.forEach(subject => {
    const quarterlyData = QUARTERS.map(q => {
      const stat = data.stats[subject]?.[q];
      return stat ? stat.quality : null;
    }).filter(v => v !== null);
    
    if (quarterlyData.length >= 2) {
      // Простая линейная регрессия для прогноза
      const n = quarterlyData.length;
      const sumX = (n * (n + 1)) / 2;
      const sumY = quarterlyData.reduce((a, b) => a + b, 0);
      const sumXY = quarterlyData.reduce((acc, y, i) => acc + (i + 1) * y, 0);
      const sumX2 = (n * (n + 1) * (2 * n + 1)) / 6;
      
      const slope = (n * sumXY - sumX * sumY) / (n * sumX2 - sumX * sumX);
      const intercept = (sumY - slope * sumX) / n;
      
      // Прогноз на следующую четверть
      const nextQuarter = intercept + slope * (n + 1);
      const trend = slope > 0 ? 'рост' : slope < 0 ? 'падение' : 'стабильно';
      const trendColor = slope > 0 ? 'g' : slope < 0 ? 'r' : 'y';
      
      trends.push({
        subject,
        current: quarterlyData[quarterlyData.length - 1],
        forecast: Math.max(0, Math.min(100, nextQuarter)),
        slope,
        trend,
        trendColor
      });
    }
  });
  
  // Сортируем по изменению (худшие первыми)
  trends.sort((a, b) => a.slope - b.slope);
  
  let html = `
    <div class="stitle">Прогноз успеваемости на следующую четверть</div>
    <div class="panel-sub" style="margin-bottom:20px">
      На основе анализа динамики за предыдущие четверти. Красный цвет — ожидается снижение качества знаний.
    </div>
    <div class="cards">
      ${card('Предметов на росте', trends.filter(t => t.slope > 0).length, 'g')}
      ${card('Стабильных', trends.filter(t => t.slope >= -0.5 && t.slope <= 0.5).length, 'y')}
      ${card('На падении', trends.filter(t => t.slope < 0).length, 'r')}
    </div>
    <div class="tbl-wrap"><table class="dt">
      <thead><tr>
        <th>Предмет</th>
        <th>Текущий %</th>
        <th>Прогноз %</th>
        <th>Изменение</th>
        <th>Тренд</th>
        <th>Рекомендации</th>
      </tr></thead>
      <tbody>`;
  
  trends.forEach(t => {
    const change = t.forecast - t.current;
    const changeStr = (change > 0 ? '+' : '') + change.toFixed(1) + '%';
    const changeColor = change > 0 ? 'var(--grn)' : change < 0 ? 'var(--red)' : 'var(--yel)';
    
    let recommendation = '';
    if (t.slope < -2) {
      recommendation = '⚠️ Критично: срочные меры';
    } else if (t.slope < 0) {
      recommendation = '🔍 Усилить контроль';
    } else if (t.slope > 2) {
      recommendation = '✅ Продолжать в том же духе';
    } else {
      recommendation = '📊 Мониторинг';
    }
    
    html += `<tr>
      <td><strong>${t.subject}</strong></td>
      <td>${t.current.toFixed(1)}%</td>
      <td><strong class="${t.trendColor}">${t.forecast.toFixed(1)}%</strong></td>
      <td style="color:${changeColor}">${changeStr}</td>
      <td><span class="badge b-${t.trendColor === 'g' ? 'ok' : t.trendColor === 'r' ? 'bad' : 'warn'}">${t.trend}</span></td>
      <td style="font-size:12px">${recommendation}</td>
    </tr>`;
  });
  
  html += `</tbody></table></div>`;
  
  // График прогноза
  html += `<div class="cgrid one" style="margin-top:24px">
    <div class="ccrd"><h3>Топ-10 предметов: текущий и прогнозируемый %</h3><canvas id="c_an_forecast"></canvas></div>
  </div>`;
  
  el.innerHTML = html;
  
  // Создаем график
  const top10 = trends.slice(0, 10);
  mkChart('c_an_forecast', {
    type: 'bar',
    data: {
      labels: top10.map(t => t.subject),
      datasets: [
        {
          label: 'Текущий %',
          data: top10.map(t => t.current),
          backgroundColor: '#3b82f6aa',
          borderColor: '#3b82f6',
          borderWidth: 1
        },
        {
          label: 'Прогноз %',
          data: top10.map(t => t.forecast),
          backgroundColor: top10.map(t => t.trendColor === 'g' ? '#34d399aa' : t.trendColor === 'r' ? '#f85149aa' : '#f5d020aa'),
          borderColor: top10.map(t => t.trendColor === 'g' ? '#34d399' : t.trendColor === 'r' ? '#f85149' : '#f5d020'),
          borderWidth: 1
        }
      ]
    },
    options: {
      indexAxis: 'y',
      scales: {
        x: { max: 100, ticks: { color: '#7a90b0', callback: v => v + '%' }, grid: { color: '#1e2d45' } },
        y: { ticks: { color: '#e2eaf8', font: { size: 10 } }, grid: { display: false } }
      },
      plugins: { legend: { labels: { color: '#e2eaf8' } } }
    }
  }, 'an');
}

// Распределение оценок
function anDistribution() {
  const data = S.an.data;
  if (!data) return;
  
  const el = document.getElementById('an-distribution');
  
  // Общее распределение
  const totalDist = { 2: 0, 3: 0, 4: 0, 5: 0 };
  data.rows.forEach(r => {
    if (!QUARTERS.includes(r.quarter)) return;
    data.subjects.forEach(s => {
      const v = r[s];
      if (!isNaN(v) && totalDist[v] !== undefined) totalDist[v]++;
    });
  });
  
  const total = Object.values(totalDist).reduce((a, b) => a + b, 0) || 1;
  
  // Распределение по ученикам
  const students = [...new Set(data.rows.map(r => r.student))];
  const studentStats = students.map(student => {
    const studentRows = data.rows.filter(r => r.student === student && QUARTERS.includes(r.quarter));
    const grades = [];
    studentRows.forEach(r => {
      data.subjects.forEach(s => {
        const v = r[s];
        if (!isNaN(v)) grades.push(v);
      });
    });
    
    const avg = grades.length ? grades.reduce((a, b) => a + b, 0) / grades.length : 0;
    const quality = grades.filter(g => g >= 4).length / grades.length * 100 || 0;
    
    return { student, avg, quality, count: grades.length };
  });
  
  studentStats.sort((a, b) => b.quality - a.quality);
  
  let html = `
    <div class="stitle">Распределение оценок и успеваемости</div>
    <div class="cards">
      ${card('Оценок "5"', totalDist[5], 'g', (totalDist[5] / total * 100).toFixed(1) + '%')}
      ${card('Оценок "4"', totalDist[4], 'b', (totalDist[4] / total * 100).toFixed(1) + '%')}
      ${card('Оценок "3"', totalDist[3], 'y', (totalDist[3] / total * 100).toFixed(1) + '%')}
      ${card('Оценок "2"', totalDist[2], 'r', (totalDist[2] / total * 100).toFixed(1) + '%')}
      ${card('Средний балл', ((2 * totalDist[2] + 3 * totalDist[3] + 4 * totalDist[4] + 5 * totalDist[5]) / total).toFixed(2), '', 'по всем оценкам')}
      ${card('% качества', ((totalDist[4] + totalDist[5]) / total * 100).toFixed(1) + '%', '')}
    </div>
    
    <div class="cgrid">
      <div class="ccrd"><h3>Общее распределение оценок</h3><canvas id="c_an_dist_pie"></canvas></div>
      <div class="ccrd"><h3>Распределение учеников по среднему баллу</h3><canvas id="c_an_dist_hist"></canvas></div>
    </div>
    
    <div class="stitle" style="margin-top:28px">Рейтинг учеников по % качества</div>
    <div class="tbl-wrap"><table class="dt">
      <thead><tr>
        <th>№</th>
        <th>Ученик</th>
        <th>Средний балл</th>
        <th>% качества</th>
        <th>Всего оценок</th>
        <th>Статус</th>
      </tr></thead>
      <tbody>`;
  
  studentStats.slice(0, 20).forEach((s, i) => {
    html += `<tr>
      <td><strong>${i + 1}</strong></td>
      <td>${s.student}</td>
      <td><strong style="color:${s.avg >= 4 ? 'var(--grn)' : s.avg >= 3.5 ? 'var(--blu)' : s.avg >= 3 ? 'var(--yel)' : 'var(--red)'}">${s.avg.toFixed(2)}</strong></td>
      <td><strong style="color:${qc(s.quality)}">${s.quality.toFixed(1)}%</strong></td>
      <td>${s.count}</td>
      <td>${badge(s.quality)}</td>
    </tr>`;
  });
  
  html += `</tbody></table></div>`;
  
  el.innerHTML = html;
  
  // Pie chart
  mkChart('c_an_dist_pie', {
    type: 'doughnut',
    data: {
      labels: ['2', '3', '4', '5'],
      datasets: [{
        data: [2, 3, 4, 5].map(g => totalDist[g]),
        backgroundColor: ['#f85149', '#f5d020', '#3b82f6', '#34d399'],
        borderColor: '#080c14',
        borderWidth: 3
      }]
    },
    options: {
      plugins: {
        legend: { position: 'bottom', labels: { color: '#e2eaf8', padding: 14 } }
      }
    }
  }, 'an');
  
  // Histogram
  const histogram = { '2.0-2.5': 0, '2.5-3.0': 0, '3.0-3.5': 0, '3.5-4.0': 0, '4.0-4.5': 0, '4.5-5.0': 0 };
  studentStats.forEach(s => {
    if (s.avg < 2.5) histogram['2.0-2.5']++;
    else if (s.avg < 3.0) histogram['2.5-3.0']++;
    else if (s.avg < 3.5) histogram['3.0-3.5']++;
    else if (s.avg < 4.0) histogram['3.5-4.0']++;
    else if (s.avg < 4.5) histogram['4.0-4.5']++;
    else histogram['4.5-5.0']++;
  });
  
  mkChart('c_an_dist_hist', {
    type: 'bar',
    data: {
      labels: Object.keys(histogram),
      datasets: [{
        label: 'Количество учеников',
        data: Object.values(histogram),
        backgroundColor: ['#f85149', '#f5d020', '#f5d020', '#3b82f6', '#3b82f6', '#34d399'],
        borderColor: ['#f85149', '#f5d020', '#f5d020', '#3b82f6', '#3b82f6', '#34d399'],
        borderWidth: 1,
        borderRadius: 6
      }]
    },
    options: {
      plugins: { legend: { display: false } },
      scales: {
        y: { ticks: { color: '#7a90b0' }, grid: { color: '#1e2d45' } },
        x: { ticks: { color: '#e2eaf8' }, grid: { display: false } }
      }
    }
  }, 'an');
}

// Корреляционный анализ
function anCorrelations() {
  const data = S.an.data;
  if (!data) return;
  
  const el = document.getElementById('an-correlations');
  
  // Вычисляем корреляцию между предметами
  const correlations = [];
  
  for (let i = 0; i < data.subjects.length; i++) {
    for (let j = i + 1; j < data.subjects.length; j++) {
      const subj1 = data.subjects[i];
      const subj2 = data.subjects[j];
      
      // Собираем парные оценки
      const pairs = [];
      data.rows.forEach(r => {
        if (!QUARTERS.includes(r.quarter)) return;
        const v1 = r[subj1];
        const v2 = r[subj2];
        if (!isNaN(v1) && !isNaN(v2)) {
          pairs.push([v1, v2]);
        }
      });
      
      if (pairs.length < 5) continue;
      
      // Вычисляем коэффициент корреляции Пирсона
      const n = pairs.length;
      const sumX = pairs.reduce((a, b) => a + b[0], 0);
      const sumY = pairs.reduce((a, b) => a + b[1], 0);
      const sumXY = pairs.reduce((a, b) => a + b[0] * b[1], 0);
      const sumX2 = pairs.reduce((a, b) => a + b[0] * b[0], 0);
      const sumY2 = pairs.reduce((a, b) => a + b[1] * b[1], 0);
      
      const numerator = n * sumXY - sumX * sumY;
      const denominator = Math.sqrt((n * sumX2 - sumX * sumX) * (n * sumY2 - sumY * sumY));
      
      if (denominator === 0) continue;
      
      const correlation = numerator / denominator;
      
      correlations.push({
        subj1,
        subj2,
        correlation,
        strength: Math.abs(correlation) > 0.7 ? 'сильная' : Math.abs(correlation) > 0.4 ? 'средняя' : 'слабая',
        type: correlation > 0 ? 'положительная' : 'отрицательная'
      });
    }
  }
  
  correlations.sort((a, b) => Math.abs(b.correlation) - Math.abs(a.correlation));
  
  let html = `
    <div class="stitle">Корреляционный анализ между предметами</div>
    <div class="panel-sub" style="margin-bottom:20px">
      Анализ взаимосвязей между оценками по разным предметам. Сильная положительная корреляция означает, что ученики, успевающие по одному предмету, как правило, успевают и по другому.
    </div>
    
    <div class="cards">
      ${card('Сильных связей', correlations.filter(c => Math.abs(c.correlation) > 0.7).length, 'g')}
      ${card('Средних связей', correlations.filter(c => Math.abs(c.correlation) > 0.4 && Math.abs(c.correlation) <= 0.7).length, 'y')}
      ${card('Слабых связей', correlations.filter(c => Math.abs(c.correlation) <= 0.4).length, 'b')}
    </div>
    
    <div class="tbl-wrap"><table class="dt">
      <thead><tr>
        <th>Предмет 1</th>
        <th>Предмет 2</th>
        <th>Корреляция</th>
        <th>Сила связи</th>
        <th>Тип</th>
        <th>Интерпретация</th>
      </tr></thead>
      <tbody>`;
  
  correlations.slice(0, 15).forEach(c => {
    const corrColor = Math.abs(c.correlation) > 0.7 ? 'var(--grn)' : Math.abs(c.correlation) > 0.4 ? 'var(--yel)' : 'var(--mut)';
    
    let interpretation = '';
    if (Math.abs(c.correlation) > 0.7) {
      interpretation = c.correlation > 0 ? '✅ Успехи связаны' : '⚠️ Обратная зависимость';
    } else if (Math.abs(c.correlation) > 0.4) {
      interpretation = '📊 Умеренная связь';
    } else {
      interpretation = '➖ Слабая связь';
    }
    
    html += `<tr>
      <td><strong>${c.subj1}</strong></td>
      <td><strong>${c.subj2}</strong></td>
      <td><strong style="color:${corrColor}">${c.correlation.toFixed(3)}</strong></td>
      <td><span class="badge b-${Math.abs(c.correlation) > 0.7 ? 'ok' : Math.abs(c.correlation) > 0.4 ? 'warn' : 'blue'}">${c.strength}</span></td>
      <td>${c.type}</td>
      <td style="font-size:12px">${interpretation}</td>
    </tr>`;
  });
  
  html += `</tbody></table></div>`;
  
  el.innerHTML = html;
}

// Инсайты и рекомендации
function anInsights() {
  const data = S.an.data;
  if (!data) return;
  
  const el = document.getElementById('an-insights');
  
  const insights = [];
  
  // Анализ по предметам
  data.subjects.forEach(subject => {
    const stat = annualOrLast(data.stats, subject);
    if (!stat) return;
    
    if (stat.quality < 40) {
      insights.push({
        level: 'critical',
        title: `Критическая ситуация: ${subject}`,
        description: `Качество знаний всего ${stat.quality.toFixed(1)}%. Требуется срочное вмешательство.`,
        actions: [
          'Провести дополнительные занятия',
          'Индивидуальная работа с отстающими',
          'Пересмотреть методику преподавания',
          'Организовать взаимопомощь среди учеников'
        ]
      });
    } else if (stat.quality < 50) {
      insights.push({
        level: 'warning',
        title: `Требует внимания: ${subject}`,
        description: `Качество знаний ${stat.quality.toFixed(1)}%. Необходимы меры по улучшению.`,
        actions: [
          'Усилить контроль текущей успеваемости',
          'Провести контрольные срезы',
          'Дополнительные консультации'
        ]
      });
    } else if (stat.quality > 85) {
      insights.push({
        level: 'success',
        title: `Отличные результаты: ${subject}`,
        description: `Качество знаний ${stat.quality.toFixed(1)}%. Продолжайте в том же духе!`,
        actions: [
          'Поддерживать текущий уровень',
          'Использовать опыт для других предметов',
          'Углубленное изучение темы'
        ]
      });
    }
  });
  
  // Анализ динамики
  const quarterStats = QUARTERS.map(q => {
    const qualities = data.subjects.map(s => data.stats[s]?.[q]?.quality).filter(v => v !== undefined);
    return qualities.length ? qualities.reduce((a, b) => a + b, 0) / qualities.length : null;
  }).filter(v => v !== null);
  
  if (quarterStats.length >= 2) {
    const trend = quarterStats[quarterStats.length - 1] - quarterStats[0];
    if (trend > 5) {
      insights.push({
        level: 'success',
        title: 'Положительная динамика',
        description: `Качество знаний выросло на ${trend.toFixed(1)}% за отчетный период.`,
        actions: [
          'Проанализировать успешные методики',
          'Распространить положительный опыт',
          'Продолжать мониторинг'
        ]
      });
    } else if (trend < -5) {
      insights.push({
        level: 'warning',
        title: 'Отрицательная динамика',
        description: `Качество знаний снизилось на ${Math.abs(trend).toFixed(1)}% за отчетный период.`,
        actions: [
          'Выявить причины снижения',
          'Провести педсовет',
          'Разработать план улучшения',
          'Усилить административный контроль'
        ]
      });
    }
  }
  
  let html = `
    <div class="stitle">Аналитические инсайты и рекомендации</div>
    <div class="panel-sub" style="margin-bottom:20px">
      На основе анализа данных система выявила ключевые точки внимания и сформировала рекомендации.
    </div>
    
    <div class="cards">
      ${card('Критичных', insights.filter(i => i.level === 'critical').length, 'r')}
      ${card('Требуют внимания', insights.filter(i => i.level === 'warning').length, 'y')}
      ${card('Положительных', insights.filter(i => i.level === 'success').length, 'g')}
    </div>`;
  
  if (insights.length === 0) {
    html += `<div style="padding:40px;text-align:center;color:var(--mut)">
      <div style="font-size:48px;margin-bottom:16px">📊</div>
      <div>Критических инсайтов не обнаружено. Продолжайте мониторинг.</div>
    </div>`;
  } else {
    html += `<div class="pgrid">`;
    
    insights.forEach(insight => {
      const levelClass = insight.level === 'critical' ? 'bad' : insight.level === 'warning' ? 'warn' : '';
      const emoji = insight.level === 'critical' ? '🚨' : insight.level === 'warning' ? '⚠️' : '✅';
      
      html += `<div class="pcrd ${levelClass}">
        <div style="font-size:28px;margin-bottom:8px">${emoji}</div>
        <div class="ps">${insight.title}</div>
        <div style="color:var(--mut);font-size:13px;margin:10px 0;line-height:1.6">${insight.description}</div>
        <div class="pr">
          <strong style="display:block;margin-bottom:6px">Рекомендуемые действия:</strong>
          ${insight.actions.map(a => `<div style="margin:4px 0">• ${a}</div>`).join('')}
        </div>
      </div>`;
    });
    
    html += `</div>`;
  }
  
  el.innerHTML = html;
}

// Экспорт данных
function anExport() {
  const el = document.getElementById('an-export');
  
  let html = `
    <div class="stitle">Экспорт данных</div>
    <div class="panel-sub" style="margin-bottom:20px">
      Скачайте аналитические отчеты в различных форматах.
    </div>
    
    <div class="panel">
      <div class="panel-title">Доступные отчеты</div>
      <div style="display:grid;grid-template-columns:repeat(auto-fill,minmax(280px,1fr));gap:16px;margin-top:20px">
        
        <div style="background:var(--sur2);border:1px solid var(--brd);border-radius:12px;padding:20px">
          <div style="font-weight:600;margin-bottom:8px">📊 Общий отчет</div>
          <div style="color:var(--mut);font-size:12px;margin-bottom:14px">Сводная аналитика по всем показателям</div>
          <button class="btn" style="width:100%;justify-content:center" onclick="alert('Функция в разработке')">Скачать PDF</button>
        </div>
        
        <div style="background:var(--sur2);border:1px solid var(--brd);border-radius:12px;padding:20px">
          <div style="font-weight:600;margin-bottom:8px">📈 Детальная статистика</div>
          <div style="color:var(--mut);font-size:12px;margin-bottom:14px">Все данные в табличном виде</div>
          <button class="btn" style="width:100%;justify-content:center" onclick="exportToExcel()">Скачать Excel</button>
        </div>
        
        <div style="background:var(--sur2);border:1px solid var(--brd);border-radius:12px;padding:20px">
          <div style="font-weight:600;margin-bottom:8px">🎯 Рекомендации</div>
          <div style="color:var(--mut);font-size:12px;margin-bottom:14px">Инсайты и план действий</div>
          <button class="btn" style="width:100%;justify-content:center" onclick="alert('Функция в разработке')">Скачать Word</button>
        </div>
        
        <div style="background:var(--sur2);border:1px solid var(--brd);border-radius:12px;padding:20px">
          <div style="font-weight:600;margin-bottom:8px">📋 Данные JSON</div>
          <div style="color:var(--mut);font-size:12px;margin-bottom:14px">Сырые данные для интеграций</div>
          <button class="btn" style="width:100%;justify-content:center" onclick="exportToJSON()">Скачать JSON</button>
        </div>
        
      </div>
    </div>`;
  
  el.innerHTML = html;
}

function exportToJSON() {
  const data = S.an.data;
  if (!data) return;
  
  const exportData = {
    name: data.name,
    subjects: data.subjects,
    students: [...new Set(data.rows.map(r => r.student))],
    stats: data.stats,
    rows: data.rows,
    exportDate: new Date().toISOString()
  };
  
  const blob = new Blob([JSON.stringify(exportData, null, 2)], { type: 'application/json' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = `eduscope_analytics_${Date.now()}.json`;
  a.click();
  URL.revokeObjectURL(url);
}

function exportToExcel() {
  const data = S.an.data;
  if (!data) return;
  
  // Подготовка данных для Excel
  const wsData = [
    ['Ученик', 'Четверть', ...data.subjects]
  ];
  
  data.rows.forEach(row => {
    const rowData = [row.student, row.quarter];
    data.subjects.forEach(s => {
      rowData.push(isNaN(row[s]) ? '' : row[s]);
    });
    wsData.push(rowData);
  });
  
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet(wsData);
  XLSX.utils.book_append_sheet(wb, ws, 'Данные');
  XLSX.writeFile(wb, `eduscope_data_${Date.now()}.xlsx`);
}

// Тепловая карта для страницы классов
(function() {
  const _origRunGrades = window.runGrades;
  window.runGrades = function() {
    _origRunGrades();
    grHeatmap();
  };
})();

function grHeatmap() {
  const C = S.gr.parsed;
  if (C.length === 0) return;
  
  const el = document.getElementById('gr-heatmap');
  
  // Собираем данные для тепловой карты
  const common = C[0].subjects.filter(s => C.every(c => c.subjects.includes(s)));
  
  let html = `
    <div class="stitle">Тепловая карта успеваемости</div>
    <div class="panel-sub" style="margin-bottom:20px">
      Визуализация качества знаний по предметам и классам. Чем ярче цвет, тем выше качество.
    </div>
    
    <div class="tbl-wrap">
      <table class="dt" style="text-align:center">
        <thead>
          <tr>
            <th style="text-align:left">Предмет</th>
            ${C.map(cls => `<th>${cls.name}</th>`).join('')}
          </tr>
        </thead>
        <tbody>`;
  
  common.forEach(subject => {
    html += `<tr><td style="text-align:left"><strong>${subject}</strong></td>`;
    C.forEach(cls => {
      const stat = annualOrLast(cls.stats, subject);
      const quality = stat ? stat.quality : 0;
      const color = quality >= 80 ? '#34d399' : quality >= 70 ? '#3b82f6' : quality >= 50 ? '#f5d020' : '#f85149';
      const bgOpacity = quality / 100 * 0.3 + 0.1;
      
      html += `<td style="background:${color}${Math.floor(bgOpacity * 255).toString(16).padStart(2, '0')};color:${color};font-weight:700">
        ${quality.toFixed(0)}%
      </td>`;
    });
    html += `</tr>`;
  });
  
  html += `</tbody></table></div>`;
  
  el.innerHTML = html;
}

// Корреляция стажа и наград для учителей
(function() {
  const _origRunAwards = window.runAwards;
  window.runAwards = function() {
    _origRunAwards();
    awCorrelation();
  };
})();

function awCorrelation() {
  const T = S.aw.teachers;
  
  const el = document.getElementById('aw-correlation');
  
  if (T.length < 5) {
    el.innerHTML = '<div style="padding:40px;color:var(--mut)">Недостаточно данных для корреляционного анализа (минимум 5 учителей)</div>';
    return;
  }
  
  // Вычисляем корреляцию между стажем и количеством наград
  const pairs = T.map(t => [t.stazh, t.totalCount]);
  const n = pairs.length;
  const sumX = pairs.reduce((a, b) => a + b[0], 0);
  const sumY = pairs.reduce((a, b) => a + b[1], 0);
  const sumXY = pairs.reduce((a, b) => a + b[0] * b[1], 0);
  const sumX2 = pairs.reduce((a, b) => a + b[0] * b[0], 0);
  const sumY2 = pairs.reduce((a, b) => a + b[1] * b[1], 0);
  
  const numerator = n * sumXY - sumX * sumY;
  const denominator = Math.sqrt((n * sumX2 - sumX * sumX) * (n * sumY2 - sumY * sumY));
  const correlation = denominator === 0 ? 0 : numerator / denominator;
  
  let html = `
    <div class="stitle">Корреляционный анализ: стаж и награды</div>
    <div class="panel-sub" style="margin-bottom:20px">
      Анализ взаимосвязи между стажем работы и количеством полученных наград.
    </div>
    
    <div class="cards">
      ${card('Коэффициент корреляции', correlation.toFixed(3), Math.abs(correlation) > 0.7 ? 'g' : Math.abs(correlation) > 0.4 ? 'y' : 'b')}
      ${card('Сила связи', Math.abs(correlation) > 0.7 ? 'Сильная' : Math.abs(correlation) > 0.4 ? 'Средняя' : 'Слабая', '')}
      ${card('Тип связи', correlation > 0 ? 'Положительная' : 'Отрицательная', '')}
    </div>
    
    <div class="cgrid one">
      <div class="ccrd">
        <h3>Диаграмма рассеяния: стаж vs награды</h3>
        <canvas id="c_aw_scatter"></canvas>
      </div>
    </div>
    
    <div style="margin-top:24px;padding:20px;background:var(--sur2);border-radius:12px;border:1px solid var(--brd)">
      <div style="font-weight:600;margin-bottom:12px">Интерпретация:</div>
      <div style="color:var(--mut);line-height:1.8">`;
  
  if (Math.abs(correlation) > 0.7) {
    if (correlation > 0) {
      html += `✅ Сильная положительная корреляция. Педагоги с большим стажем, как правило, имеют больше наград. 
      Это указывает на эффективную систему поощрения опытных учителей.`;
    } else {
      html += `⚠️ Сильная отрицательная корреляция. Это необычный результат, требующий дополнительного анализа.`;
    }
  } else if (Math.abs(correlation) > 0.4) {
    html += `📊 Умеренная связь между стажем и наградами. Награды получают как опытные, так и молодые педагоги.`;
  } else {
    html += `➖ Слабая связь между стажем и наградами. Награды распределяются относительно равномерно независимо от стажа.`;
  }
  
  html += `</div></div>`;
  
  el.innerHTML = html;
  
  // Создаем scatter plot
  mkChart('c_aw_scatter', {
    type: 'scatter',
    data: {
      datasets: [{
        label: 'Учителя',
        data: T.map(t => ({ x: t.stazh, y: t.totalCount })),
        backgroundColor: '#3b82f6aa',
        borderColor: '#3b82f6',
        pointRadius: 6,
        pointHoverRadius: 8
      }]
    },
    options: {
      scales: {
        x: { title: { display: true, text: 'Стаж (лет)', color: '#e2eaf8' }, ticks: { color: '#7a90b0' }, grid: { color: '#1e2d45' } },
        y: { title: { display: true, text: 'Количество наград', color: '#e2eaf8' }, ticks: { color: '#7a90b0' }, grid: { color: '#1e2d45' } }
      },
      plugins: { legend: { labels: { color: '#e2eaf8' } } }
    }
  }, 'aw');
}

// Тренды рейтингов
(function() {
  const _origRunRatings = window.runRatings;
  window.runRatings = function() {
    _origRunRatings();
    rtTrends();
  };
})();

function rtTrends() {
  const el = document.getElementById('rt-trends');
  
  // Проверяем наличие данных
  if (!S.rt.teachers || S.rt.teachers.length === 0) {
    el.innerHTML = '<div style="padding:40px;color:var(--mut)">Загрузите данные рейтинга учителей</div>';
    return;
  }
  
  let html = `
    <div class="stitle">Анализ трендов рейтинга</div>
    <div class="panel-sub" style="margin-bottom:20px">
      Динамика изменения рейтинговых показателей по четвертям.
    </div>`;
  
  // Если есть данные по четвертям, строим графики
  const teachersWithQuarters = S.rt.teachers.filter(t => t.quarters && t.quarters.length > 1);
  
  if (teachersWithQuarters.length > 0) {
    html += `<div class="cgrid one">
      <div class="ccrd">
        <h3>Топ-5 учителей: динамика рейтинга</h3>
        <canvas id="c_rt_trend"></canvas>
      </div>
    </div>`;
    
    el.innerHTML = html;
    
    // Берем топ-5 учителей по последнему рейтингу
    const top5 = teachersWithQuarters
      .sort((a, b) => {
        const aLast = a.quarters[a.quarters.length - 1].total;
        const bLast = b.quarters[b.quarters.length - 1].total;
        return bLast - aLast;
      })
      .slice(0, 5);
    
    mkChart('c_rt_trend', {
      type: 'line',
      data: {
        labels: top5[0].quarters.map((_, i) => `${i + 1} четверть`),
        datasets: top5.map((t, i) => ({
          label: t.name,
          data: t.quarters.map(q => q.total),
          borderColor: PAL[i],
          backgroundColor: PAL[i] + '22',
          fill: true,
          tension: 0.35,
          pointRadius: 5,
          borderWidth: 2.5
        }))
      },
      options: {
        scales: {
          y: { ticks: { color: '#7a90b0' }, grid: { color: '#1e2d45' } },
          x: { ticks: { color: '#7a90b0' }, grid: { color: '#1e2d45' } }
        },
        plugins: { legend: { labels: { color: '#e2eaf8' } } }
      }
    }, 'rt');
  } else {
    html += '<div style="padding:40px;color:var(--mut)">Недостаточно данных для построения трендов</div>';
    el.innerHTML = html;
  }
}

// ═══════════════════════════════════════════════════════
//  TARDINESS — анализ опозданий
// ═══════════════════════════════════════════════════════

// Setup tardiness page
setupDrop('td');

function parseTardinessFile(buffer) {
  const wb = XLSX.read(buffer, { type: 'array' });
  const ws = wb.Sheets[wb.SheetNames[0]];
  const raw = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });
  if (!raw.length) return null;

  const headers = raw[0].map(h => h ? String(h).trim() : '');
  const data = [];
  
  for (let i = 1; i < raw.length; i++) {
    const row = raw[i];
    if (!row || !row.length) continue;
    
    const entry = {};
    headers.forEach((h, idx) => {
      if (h.toLowerCase().includes('дата') || h.toLowerCase().includes('date')) entry.date = row[idx];
      else if (h.toLowerCase().includes('время') || h.toLowerCase().includes('time')) entry.time = row[idx];
      else if (h.toLowerCase().includes('имя') || h.toLowerCase().includes('ученик') || h.toLowerCase().includes('фамилия')) entry.name = row[idx];
      else if (h.toLowerCase().includes('класс') || h.toLowerCase().includes('class')) entry.class = row[idx];
      else if (h.toLowerCase().includes('минут') || h.toLowerCase().includes('minute')) entry.minutes = row[idx];
    });
    
    if (entry.name && entry.class) data.push(entry);
  }
  
  return data;
}

function runTardiness() {
  if (!S.td.file) return;
  
  S.td.data = parseTardinessFile(S.td.file.data);
  if (!S.td.data || !S.td.data.length) {
    alert('Не удалось распарсить файл');
    return;
  }
  
  destroyCharts('td');
  tdOverview(); tdStudents(); tdClasses(); tdTrends(); tdTop();
  document.getElementById('td-results').style.display = 'block';
  document.getElementById('td-results').scrollIntoView({ behavior: 'smooth' });
}

function tdOverview() {
  const data = S.td.data;
  const totalIncidents = data.length;
  const uniqueStudents = [...new Set(data.map(d => d.name))].length;
  const uniqueClasses = [...new Set(data.map(d => d.class))].length;
  const avgMinutes = data.reduce((sum, d) => sum + (parseInt(d.minutes) || 0), 0) / totalIncidents || 0;
  const morningIncidents = data.filter(d => {
    const hour = parseInt(String(d.time).split(':')[0]);
    return hour < 12 && hour >= 6;
  }).length;
  
  const el = document.getElementById('td-overview');
  el.innerHTML = `
    <div class="cards">
      ${card('Всего опозданий', totalIncidents, '')}
      ${card('Уникальных учеников', uniqueStudents, '')}
      ${card('Классов', uniqueClasses, '')}
      ${card('Среднее опоздание', avgMinutes.toFixed(0) + ' мин', 'ora')}
    </div>
    <div class="cgrid">
      <div class="ccrd"><h3>Распределение по времени суток</h3><canvas id="c_td_time"></canvas></div>
      <div class="ccrd"><h3>Опоздания по классам</h3><canvas id="c_td_class"></canvas></div>
    </div>
    <div class="cgrid one" style="margin-top:16px">
      <div class="ccrd"><h3>Распределение по причинам (минуты опоздания)</h3><canvas id="c_td_minutes"></canvas></div>
    </div>`;

  // Time of day distribution
  const timeGroups = { '06-09': 0, '09-12': 0, '12-15': 0, '15-18': 0, '18-21': 0, '21-06': 0 };
  data.forEach(d => {
    const hour = parseInt(String(d.time).split(':')[0]) || 0;
    if (hour >= 6 && hour < 9) timeGroups['06-09']++;
    else if (hour >= 9 && hour < 12) timeGroups['09-12']++;
    else if (hour >= 12 && hour < 15) timeGroups['12-15']++;
    else if (hour >= 15 && hour < 18) timeGroups['15-18']++;
    else if (hour >= 18 && hour < 21) timeGroups['18-21']++;
    else timeGroups['21-06']++;
  });

  mkChart('c_td_time', {
    type: 'doughnut',
    data: {
      labels: Object.keys(timeGroups),
      datasets: [{
        data: Object.values(timeGroups),
        backgroundColor: ['#3b82f6', '#fb923c', '#f5d020', '#818cf8', '#f85149', '#3fb950'],
        borderColor: '#080c14',
        borderWidth: 2
      }]
    },
    options: { plugins: { legend: { position: 'bottom', labels: { color: '#e2eaf8' } } } }
  }, 'td');

  // Class distribution
  const classDist = {};
  data.forEach(d => {
    const cls = String(d.class).trim();
    classDist[cls] = (classDist[cls] || 0) + 1;
  });
  const classes = Object.keys(classDist).sort();
  const classValues = classes.map(c => classDist[c]);

  mkChart('c_td_class', {
    type: 'bar',
    data: {
      labels: classes,
      datasets: [{
        data: classValues,
        backgroundColor: '#fb923c',
        borderRadius: 6
      }]
    },
    options: {
      plugins: { legend: { display: false } },
      scales: {
        y: { ticks: { color: '#7a90b0' }, grid: { color: '#1e2d45' } },
        x: { ticks: { color: '#e2eaf8' }, grid: { display: false } }
      }
    }
  }, 'td');

  // Minutes distribution
  const minuteRanges = { '1-5': 0, '6-10': 0, '11-20': 0, '21-30': 0, '30+': 0 };
  data.forEach(d => {
    const mins = parseInt(d.minutes) || 0;
    if (mins <= 5) minuteRanges['1-5']++;
    else if (mins <= 10) minuteRanges['6-10']++;
    else if (mins <= 20) minuteRanges['11-20']++;
    else if (mins <= 30) minuteRanges['21-30']++;
    else minuteRanges['30+']++;
  });

  mkChart('c_td_minutes', {
    type: 'bar',
    data: {
      labels: Object.keys(minuteRanges),
      datasets: [{
        label: 'Количество опозданий',
        data: Object.values(minuteRanges),
        backgroundColor: '#818cf8',
        borderRadius: 6
      }]
    },
    options: {
      plugins: { legend: { labels: { color: '#e2eaf8' } } },
      scales: {
        y: { ticks: { color: '#7a90b0' }, grid: { color: '#1e2d45' } },
        x: { ticks: { color: '#e2eaf8' }, grid: { display: false } }
      }
    }
  }, 'td');
}

function tdStudents() {
  const data = S.td.data;
  const studentStats = {};
  
  data.forEach(d => {
    if (!studentStats[d.name]) {
      studentStats[d.name] = { name: d.name, class: d.class, count: 0, totalMinutes: 0, days: [] };
    }
    studentStats[d.name].count++;
    studentStats[d.name].totalMinutes += parseInt(d.minutes) || 0;
  });
  
  const students = Object.values(studentStats).sort((a, b) => b.count - a.count);
  
  let rows = '';
  students.forEach((s, idx) => {
    const avgMin = (s.totalMinutes / s.count).toFixed(1);
    const status = s.count === 1 ? '<span class="badge b-ok">Редко</span>' :
                   s.count <= 3 ? '<span class="badge b-warn">Часто</span>' :
                   '<span class="badge b-bad">Систематически</span>';
    
    rows += `<tr>
      <td style="color:var(--acc)">${idx + 1}</td>
      <td>${s.name}</td>
      <td>${s.class}</td>
      <td><strong>${s.count}</strong></td>
      <td>${avgMin} мин</td>
      <td>${status}</td>
    </tr>`;
  });
  
  document.getElementById('td-students').innerHTML = `
    <div class="stitle">Аналитика по ученикам</div>
    <div class="tbl-wrap"><table class="dt">
      <thead><tr><th>#</th><th>Ученик</th><th>Класс</th><th>Опозданий</th><th>Среднее</th><th>Статус</th></tr></thead>
      <tbody>${rows}</tbody>
    </table></div>`;
}

function tdClasses() {
  const data = S.td.data;
  const classStats = {};
  
  data.forEach(d => {
    if (!classStats[d.class]) {
      classStats[d.class] = { class: d.class, count: 0, students: new Set(), avgMin: 0 };
    }
    classStats[d.class].count++;
    classStats[d.class].students.add(d.name);
    classStats[d.class].avgMin += parseInt(d.minutes) || 0;
  });
  
  const classes = Object.values(classStats)
    .map(c => ({ ...c, students: c.students.size, avgMin: (c.avgMin / c.count).toFixed(1) }))
    .sort((a, b) => b.count - a.count);
  
  let rows = '';
  classes.forEach(c => {
    const severity = c.count < 5 ? '<span class="badge b-ok">Норма</span>' :
                     c.count < 15 ? '<span class="badge b-warn">Внимание</span>' :
                     '<span class="badge b-bad">Критично</span>';
    
    rows += `<tr>
      <td style="color:var(--ora);font-weight:600">${c.class}</td>
      <td><strong>${c.count}</strong></td>
      <td>${c.students}</td>
      <td>${c.avgMin} мин</td>
      <td>${severity}</td>
    </tr>`;
  });
  
  document.getElementById('td-classes').innerHTML = `
    <div class="stitle">Аналитика по классам</div>
    <div class="tbl-wrap"><table class="dt">
      <thead><tr><th>Класс</th><th>Всего опозданий</th><th>Уникальных учеников</th><th>Среднее (мин)</th><th>Статус</th></tr></thead>
      <tbody>${rows}</tbody>
    </table></div>`;
}

function tdTrends() {
  const data = S.td.data;
  
  // Group by date
  const dateGroups = {};
  data.forEach(d => {
    const dateStr = String(d.date).split(' ')[0] || 'Неизвестно';
    dateGroups[dateStr] = (dateGroups[dateStr] || 0) + 1;
  });
  
  const sortedDates = Object.keys(dateGroups).sort().slice(-30); // Last 30 days
  const values = sortedDates.map(d => dateGroups[d]);
  
  document.getElementById('td-trends').innerHTML = `
    <div class="stitle">Тренды опозданий</div>
    <div class="cgrid one">
      <div class="ccrd"><h3>Опоздания по дням (последние 30 дней)</h3><canvas id="c_td_trends"></canvas></div>
    </div>`;
  
  mkChart('c_td_trends', {
    type: 'line',
    data: {
      labels: sortedDates,
      datasets: [{
        label: 'Количество опозданий',
        data: values,
        borderColor: '#fb923c',
        backgroundColor: '#fb923c22',
        fill: true,
        tension: 0.35,
        pointRadius: 4,
        borderWidth: 2.5
      }]
    },
    options: {
      plugins: { legend: { labels: { color: '#e2eaf8' } } },
      scales: {
        y: { ticks: { color: '#7a90b0' }, grid: { color: '#1e2d45' } },
        x: { ticks: { color: '#7a90b0' }, grid: { color: '#1e2d45' } }
      }
    }
  }, 'td');
}

function tdTop() {
  const data = S.td.data;
  const studentStats = {};
  
  data.forEach(d => {
    if (!studentStats[d.name]) {
      studentStats[d.name] = { name: d.name, class: d.class, count: 0 };
    }
    studentStats[d.name].count++;
  });
  
  const top = Object.values(studentStats).sort((a, b) => b.count - a.count).slice(0, 10);
  
  let html = '<div class="stitle">Топ-10 нарушителей</div><div class="pgrid">';
  top.forEach((student, idx) => {
    const percentage = ((student.count / data.length) * 100).toFixed(1);
    html += `
      <div class="pcrd bad">
        <div class="ps">${idx + 1}. ${student.name}</div>
        <div class="pq">${student.count}</div>
        <div class="pr">Класс: <strong>${student.class}</strong><br>Доля: <strong>${percentage}%</strong> от всех опозданий</div>
      </div>`;
  });
  html += '</div>';
  
  document.getElementById('td-top').innerHTML = html;
}
