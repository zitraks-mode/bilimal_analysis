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
  aw: { file: null, teachers: [], charts: [] }
};
const multi_ = { gr: true, aw: false };

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
  const root = ns === 'gr' ? 'gr-results' : ns === 'aw' ? 'aw-main' : '';
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
