/* Дашборд весового контроля NTE200 — приложение.
   Работает поверх партиций (борт × месяц) из agg.js. Всё считается на клиенте,
   поэтому загруженные .xlsx обрабатываются той же логикой, что и встроенные данные. */
(function () {
  'use strict';
  var NOMINAL = 180, P110 = 198, P120 = 216;
  var MONTHS_RU = { '01':'янв','02':'фев','03':'мар','04':'апр','05':'май','06':'июн','07':'июл','08':'авг','09':'сен','10':'окт','11':'ноя','12':'дек' };
  var LS_KEY = 'nte200_weighing_v1';
  var Z_COLORS = ['var(--under)','var(--target)','var(--accept)','var(--crit)'];
  var Z_NAMES = ['Недогруз (<180 т)','Целевая (180–198 т)','Перегруз доп. (198–216 т)','Перегруз крит. (>216 т)'];

  var EMBEDDED = JSON.parse(document.getElementById('embedded-data').textContent);

  var state = {
    partitions: EMBEDDED.partitions,   // активный набор
    source: 'Демо-данные парка',
    store: null,                        // AGG store при загрузке .xlsx (для дедупликации между файлами)
    files: [],                          // имена загруженных файлов
    selMonths: null,                    // Set выбранных ym; null = все
    selTrucks: null                     // Set выбранных бортов; null = все
  };

  // ---- восстановление из localStorage ----
  try {
    var saved = localStorage.getItem(LS_KEY);
    if (saved) {
      var obj = JSON.parse(saved);
      if (obj && obj.partitions && obj.partitions.length) {
        state.partitions = obj.partitions;
        state.source = obj.source || 'Сохранённые данные';
        state.files = obj.files || [];
      }
    }
  } catch (e) {}

  // ---------- helpers ----------
  function el(id) { return document.getElementById(id); }
  function fmtInt(n) { return Math.round(n).toLocaleString('ru-RU'); }
  function fmt1(n) { return Number(n).toLocaleString('ru-RU', { minimumFractionDigits: 1, maximumFractionDigits: 1 }); }
  function fmt2(n) { return Number(n).toLocaleString('ru-RU', { minimumFractionDigits: 2, maximumFractionDigits: 2 }); }
  function ymLabel(ym) { var p = ym.split('-'); return MONTHS_RU[p[1]] + ' ' + p[0]; }
  function svg(tag, a) { var e = document.createElementNS('http://www.w3.org/2000/svg', tag); for (var k in a) e.setAttribute(k, a[k]); return e; }

  var tip = el('tip');
  function showTip(html, e) { tip.innerHTML = html; tip.classList.add('on'); moveTip(e); }
  function moveTip(e) {
    var pad = 14, tw = 230, th = 80, x = e.clientX + pad, y = e.clientY + pad;
    if (x + tw > innerWidth) x = e.clientX - tw - pad;
    if (y + th > innerHeight) y = e.clientY - th - pad;
    tip.style.left = x + 'px'; tip.style.top = y + 'px';
  }
  function hideTip() { tip.classList.remove('on'); }

  // ---------- выборка партиций по фильтрам ----------
  function activeParts() {
    return state.partitions.filter(function (p) {
      if (state.selTrucks && !state.selTrucks.has(p.t)) return false;
      if (state.selMonths && !state.selMonths.has(p.m)) return false;
      return true;
    });
  }

  // сложение партиций -> сводка
  function summarize(parts) {
    var c = 0, sp = 0, z = [0, 0, 0, 0], h = {}, ct = [0, 0, 0, 0, 0, 0],
        dl = 0, de = 0, ls = 0, es = 0, lf = 0, rf = 0, lr = 0, rr = 0, cb = 0, f = 0;
    parts.forEach(function (p) {
      c += p.c; sp += p.sp;
      for (var i = 0; i < 4; i++) z[i] += p.z[i];
      for (var b in p.h) h[b] = (h[b] || 0) + p.h[b];
      for (var j = 0; j < 6; j++) ct[j] += p.ct[j];
      dl += p.dl; de += p.de; ls += p.ls; es += p.es;
      lf += (p.lf || 0); rf += (p.rf || 0); lr += p.lr; rr += p.rr; cb += (p.cb || 0); f += p.f;
    });
    return { c: c, sp: sp, z: z, h: h, ct: ct, dl: dl, de: de, ls: ls, es: es, lf: lf, rf: rf, lr: lr, rr: rr, cb: cb, f: f };
  }

  function allTrucks() {
    var s = {}; state.partitions.forEach(function (p) { s[p.t] = 1; });
    return Object.keys(s).sort(function (a, b) { return +a - +b; });
  }
  function allMonths() {
    var s = {}; state.partitions.forEach(function (p) { s[p.m] = 1; });
    return Object.keys(s).sort();
  }

  // ---------- render: filters ----------
  function renderFilters() {
    var ms = el('month-scroll'); ms.innerHTML = '';
    allMonths().forEach(function (ym) {
      var b = document.createElement('button');
      b.className = 'chip' + ((!state.selMonths || state.selMonths.has(ym)) ? ' on' : '');
      b.textContent = ymLabel(ym);
      b.onclick = function () { toggle('selMonths', ym, allMonths()); };
      ms.appendChild(b);
    });
    var ts = el('truck-scroll'); ts.innerHTML = '';
    allTrucks().forEach(function (t) {
      var b = document.createElement('button');
      b.className = 'chip truck' + ((!state.selTrucks || state.selTrucks.has(t)) ? ' on' : '');
      b.textContent = t;
      b.onclick = function () { toggle('selTrucks', t, allTrucks()); };
      ts.appendChild(b);
    });
  }
  function toggle(field, val, universe) {
    var cur = state[field];
    if (!cur) { cur = new Set(universe); }       // всё выбрано -> материализуем
    if (cur.has(val)) cur.delete(val); else cur.add(val);
    if (cur.size === 0) cur = new Set(universe); // не даём пусто -> сброс на всё
    state[field] = (cur.size === universe.length) ? null : cur;
    renderFilters(); renderAll();
  }

  // ---------- render: status ----------
  function renderStatus() {
    var parts = state.partitions;
    var trucks = allTrucks().length, cyc = parts.reduce(function (s, p) { return s + p.c; }, 0);
    var months = allMonths();
    var period = months.length ? (ymLabel(months[0]) + ' – ' + ymLabel(months[months.length - 1])) : '—';
    el('status').innerHTML =
      '<span><span class="k">Бортов:</span> <b>' + trucks + '</b></span>' +
      '<span><span class="k">Уникальных циклов:</span> <b>' + fmtInt(cyc) + '</b></span>' +
      '<span><span class="k">Период:</span> <b>' + period + '</b></span>' +
      '<span class="src"><span class="dot"></span>' + state.source + '</span>';
  }

  // ---------- histogram ----------
  function renderHist(sum) {
    var s = el('hist'); s.innerHTML = '';
    var W = 720, H = 340, m = { t: 26, r: 12, b: 40, l: 44 };
    var pw = W - m.l - m.r, ph = H - m.t - m.b;
    var BIN = 2, LO = 150, HI = 224;
    var bins = [];
    for (var x = LO; x < HI; x += BIN) bins.push(x);
    var counts = bins.map(function () { return 0; });
    var total = 0, sumw = 0, sumw2 = 0;
    for (var b in sum.h) {
      var t = +b + 0.5, cnt = sum.h[b];
      total += cnt; sumw += t * cnt; sumw2 += t * t * cnt;
      if (t < LO) { counts[0] += cnt; continue; }
      if (t >= HI) { counts[counts.length - 1] += cnt; continue; }
      counts[Math.floor((t - LO) / BIN)] += cnt;
    }
    if (total === 0) { s.appendChild(svg('text', { x: W / 2, y: H / 2, 'text-anchor': 'middle', class: 'ax' })).textContent = 'Нет данных'; return; }
    var mean = sumw / total, variance = Math.max(1, sumw2 / total - mean * mean), sd = Math.sqrt(variance);
    var maxPct = Math.max.apply(null, counts.map(function (c) { return c / total * 100; }));
    var yMax = Math.ceil(maxPct / 3) * 3 || 3;
    var xOf = function (v) { return m.l + (v - LO) / (HI - LO) * pw; };
    var barW = pw / bins.length;
    var yOf = function (pct) { return m.t + ph - pct / yMax * ph; };
    var g = svg('g', {}); s.appendChild(g);

    // gridlines + y labels (%)
    for (var i = 0; i <= 3; i++) {
      var pv = yMax * i / 3, yy = yOf(pv);
      g.appendChild(svg('line', { x1: m.l, x2: W - m.r, y1: yy, y2: yy, class: 'gl' }));
      var tl = svg('text', { x: m.l - 6, y: yy + 3, 'text-anchor': 'end', class: 'ax' }); tl.textContent = Math.round(pv) + '%'; g.appendChild(tl);
    }
    // zone boundary lines (180,198,216)
    [[NOMINAL, 'ном. 180'], [P110, '110% · 198'], [P120, '120% · 216']].forEach(function (zb) {
      if (zb[0] < LO || zb[0] > HI) return;
      var zx = xOf(zb[0]);
      var ln = svg('line', { x1: zx, x2: zx, y1: m.t, y2: m.t + ph, class: 'zline', stroke: 'var(--ink-mute)' }); g.appendChild(ln);
      var lb = svg('text', { x: zx, y: m.t - 5, 'text-anchor': 'middle', class: 'axb' }); lb.textContent = zb[1]; g.appendChild(lb);
    });
    // bars
    var zoneOf = function (t) { return t < NOMINAL ? 0 : t <= P110 ? 1 : t <= P120 ? 2 : 3; };
    bins.forEach(function (bx, i) {
      var pct = counts[i] / total * 100; if (counts[i] === 0) return;
      var center = bx + BIN / 2, col = Z_COLORS[zoneOf(center)];
      var yy = yOf(pct), hh = m.t + ph - yy;
      var rect = svg('rect', { x: xOf(bx) + 1, y: yy, width: Math.max(barW - 2, 1), height: Math.max(hh, 1), rx: 2, fill: col });
      rect.addEventListener('mousemove', function (e) { showTip('<b>' + bx + '–' + (bx + BIN) + ' т</b><br>' + fmtInt(counts[i]) + ' циклов · ' + fmt1(pct) + '%<br>' + Z_NAMES[zoneOf(center)], e); });
      rect.addEventListener('mouseleave', hideTip);
      g.appendChild(rect);
    });
    // normal curve
    var pdf = function (v) { return Math.exp(-(v - mean) * (v - mean) / (2 * variance)) / (sd * Math.sqrt(2 * Math.PI)); };
    var peakPct = pdf(mean) * BIN * 100;
    var scale = maxPct / peakPct;
    var path = '';
    for (var v = LO; v <= HI; v += 1) {
      var yv = yOf(Math.min(pdf(v) * BIN * 100 * scale, yMax));
      path += (v === LO ? 'M' : 'L') + xOf(v).toFixed(1) + ',' + yv.toFixed(1) + ' ';
    }
    g.appendChild(svg('path', { d: path, fill: 'none', stroke: 'var(--ink)', 'stroke-width': 1.6, 'stroke-opacity': .55, 'stroke-linejoin': 'round' }));
    // x axis labels
    g.appendChild(svg('line', { x1: m.l, x2: W - m.r, y1: m.t + ph, y2: m.t + ph, class: 'gl' }));
    for (var xv = 150; xv <= 220; xv += 10) {
      var tx = svg('text', { x: xOf(xv), y: m.t + ph + 15, 'text-anchor': 'middle', class: 'ax' }); tx.textContent = xv; g.appendChild(tx);
    }
    var xl = svg('text', { x: m.l + pw / 2, y: H - 4, 'text-anchor': 'middle', class: 'ax' }); xl.textContent = 'Масса груза, т'; g.appendChild(xl);

    // legend under hist
    var lg = el('hist-legend'); lg.innerHTML = '';
    Z_NAMES.forEach(function (nm, i) {
      var cnt = sum.z[i], pc = cnt / sum.c * 100;
      var sw = document.createElement('span'); sw.className = 'legend'; sw.style.display = 'contents';
      var d1 = document.createElement('div'); d1.style.display = 'flex'; d1.style.alignItems = 'center'; d1.style.gap = '7px';
      d1.innerHTML = '<span style="width:11px;height:11px;border-radius:3px;background:' + Z_COLORS[i] + ';flex:none;"></span><span style="font-size:11px;color:var(--ink-2)">' + nm + '</span>';
      var d2 = document.createElement('div'); d2.style.fontFamily = 'var(--mono)'; d2.style.fontWeight = '700'; d2.style.fontSize = '11.5px';
      d2.textContent = fmt1(pc) + '%  ·  ' + fmtInt(cnt);
      lg.appendChild(d1); lg.appendChild(d2);
    });
  }

  // ---------- policy 10/10/20 ----------
  function renderPolicy(sum) {
    var mean = sum.sp / sum.c;
    var over110 = (sum.z[2] + sum.z[3]) / sum.c * 100;
    var over120 = sum.z[3] / sum.c * 100;
    var checks = [
      { ok: mean <= NOMINAL, t1: 'Средняя загрузка ≤ номинала', t2: 'среднее по распределению ≤ 180 т', num: fmt1(mean) + ' т' },
      { ok: over110 <= 10, t1: 'Не более 10% загрузок > 110%', t2: 'доля циклов свыше 198 т', num: fmt1(over110) + ' %' },
      { ok: over120 === 0, t1: 'Ни одной загрузки > 120%', t2: 'доля циклов свыше 216 т (' + fmtInt(sum.z[3]) + ' шт.)', num: fmt2(over120) + ' %' }
    ];
    var passN = checks.filter(function (c) { return c.ok; }).length;
    var allBadge = el('policy-all');
    allBadge.className = 'badge-all ' + (passN === 3 ? 'pass' : 'fail');
    allBadge.textContent = passN === 3 ? 'СООТВЕТСТВУЕТ' : passN + '/3 · ' + (3 - passN) + ' наруш.';
    var checkSvg = '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="3"><path d="M5 13l4 4L19 7"/></svg>';
    var crossSvg = '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="3"><path d="M6 6l12 12M18 6L6 18"/></svg>';
    el('policy-rows').innerHTML = checks.map(function (c) {
      return '<div class="prow"><div class="ico ' + (c.ok ? 'pass' : 'fail') + '">' + (c.ok ? checkSvg : crossSvg) + '</div>' +
        '<div class="txt"><span class="t1">' + c.t1 + '</span><span class="t2">' + c.t2 + '</span></div>' +
        '<span class="num ' + (c.ok ? 'pass' : 'fail') + '">' + c.num + '</span></div>';
    }).join('');
  }

  // ---------- donut ----------
  function renderDonut(sum) {
    var s = el('donut'); s.innerHTML = '';
    var cx = 64, cy = 64, r = 52, rin = 33, tot = sum.c;
    var ang = -Math.PI / 2;
    var g = svg('g', {}); s.appendChild(g);
    sum.z.forEach(function (v, i) {
      if (v === 0) return;
      var a2 = ang + v / tot * Math.PI * 2;
      var large = (a2 - ang) > Math.PI ? 1 : 0;
      var x1 = cx + r * Math.cos(ang), y1 = cy + r * Math.sin(ang), x2 = cx + r * Math.cos(a2), y2 = cy + r * Math.sin(a2);
      var xi1 = cx + rin * Math.cos(a2), yi1 = cy + rin * Math.sin(a2), xi2 = cx + rin * Math.cos(ang), yi2 = cy + rin * Math.sin(ang);
      var d = 'M' + x1 + ',' + y1 + ' A' + r + ',' + r + ' 0 ' + large + ' 1 ' + x2 + ',' + y2 +
        ' L' + xi1 + ',' + yi1 + ' A' + rin + ',' + rin + ' 0 ' + large + ' 0 ' + xi2 + ',' + yi2 + ' Z';
      var pth = svg('path', { d: d, fill: Z_COLORS[i], stroke: 'var(--surface)', 'stroke-width': 1.5 });
      pth.addEventListener('mousemove', function (e) { showTip('<b>' + fmt1(v / tot * 100) + '%</b> · ' + fmtInt(v) + ' циклов<br>' + Z_NAMES[i], e); });
      pth.addEventListener('mouseleave', hideTip);
      g.appendChild(pth);
      ang = a2;
    });
    var tgt = sum.z[1] / tot * 100;
    var c1 = svg('text', { x: cx, y: cy - 2, 'text-anchor': 'middle', 'font-family': 'var(--mono)', 'font-weight': '800', 'font-size': '18', fill: 'var(--ink)' }); c1.textContent = fmt1(tgt) + '%'; g.appendChild(c1);
    var c2 = svg('text', { x: cx, y: cy + 12, 'text-anchor': 'middle', 'font-size': '8.5', fill: 'var(--ink-mute)' }); c2.textContent = 'целевая'; g.appendChild(c2);

    var lg = el('donut-legend'); lg.innerHTML = '';
    sum.z.forEach(function (v, i) {
      var row = document.createElement('div'); row.className = 'item';
      row.innerHTML = '<span class="sw" style="background:' + Z_COLORS[i] + '"></span><span class="nm">' + Z_NAMES[i].replace(/ \(.*/, '') + '</span><span class="pc">' + fmt1(v / tot * 100) + '%</span>';
      lg.appendChild(row);
    });
  }

  // ---------- KPIs ----------
  function renderKPIs(sum) {
    el('k-mean').innerHTML = fmt1(sum.sp / sum.c) + '<span class="u">т</span>';
    el('k-cycles').textContent = fmtInt(sum.c);
    el('k-tonnage').innerHTML = fmt2(sum.sp / 1e6) + '<span class="u">млн т</span>';
  }

  // ---------- ПРОИЗВОДИТЕЛЬНОСТЬ: вспомогательные ----------
  // фазы цикла (секунды-суммы в ct): 0 стоянка гружёный, 1 ход гружёный, 2 ход порожний,
  // 3 стоянка порожний (очередь), 4 погрузка, 5 разгрузка
  var CYC_LABELS = ['Стоянка гружёным', 'Ход гружёным', 'Ход порожним', 'Ожидание у экскаватора', 'Погрузка', 'Разгрузка'];
  var CYC_COLORS = ['var(--ink-mute)', 'var(--target)', 'var(--under)', 'var(--crit)', 'var(--gold)', 'var(--accept)'];
  var CYC_ORDER = [4, 0, 1, 5, 2, 3]; // порядок операций в реальном цикле
  // группы: продуктивное (ход) / операции (погр+разгр) / ожидание (очередь+стоянка гружёным)
  function cycGroups(ctMin) {
    return [
      { name: 'Продуктивное (движение)', color: 'var(--target)', idx: [1, 2] },
      { name: 'Операции (погрузка/разгрузка)', color: 'var(--gold)', idx: [4, 5] },
      { name: 'Ожидание (очередь + простой)', color: 'var(--crit)', idx: [3, 0] }
    ].map(function (g) {
      var m = g.idx.reduce(function (s, i) { return s + ctMin[i]; }, 0);
      return { name: g.name, color: g.color, min: m };
    });
  }

  function renderPerfKPIs(sum) {
    var c = sum.c;
    var ctMin = sum.ct.map(function (v) { return v / c / 60; });
    var cycMin = ctMin.reduce(function (s, v) { return s + v; }, 0);
    var mean = sum.sp / c;
    var prod = cycMin ? mean / (cycMin / 60) : 0;
    var perShift = cycMin ? (12 * 60 / cycMin) : 0; // рейсов за 12-часовую смену (теор., без простоев смены)
    var queue = ctMin[3];
    var tiles = [
      { l: 'Производительность', v: fmt1(prod), u: 'т/ч', big: true },
      { l: 'Время цикла', v: fmt1(cycMin), u: 'мин', big: true },
      { l: 'Рейсов за 12 ч', v: fmt1(perShift), u: 'ед.' },
      { l: 'Масса за цикл', v: fmt1(mean), u: 'т' },
      { l: 'Очередь у экск.', v: fmt1(queue), u: 'мин' }
    ];
    el('perf-kpis').innerHTML = tiles.map(function (t) {
      return '<div class="kpi"><div class="lbl">' + t.l + '</div><div class="val' + (t.big ? ' big' : '') + '">' + t.v + '<span class="u">' + t.u + '</span></div></div>';
    }).join('');
  }

  // горизонтальная лента фаз цикла
  function renderCycleBar(sum) {
    var s = el('cycle-bar'); s.innerHTML = '';
    var c = sum.c, ctMin = sum.ct.map(function (v) { return v / c / 60; });
    var total = ctMin.reduce(function (a, b) { return a + b; }, 0);
    var W = 560, H = 44, x = 0, barH = 30;
    var g = svg('g', {}); s.appendChild(g);
    CYC_ORDER.forEach(function (i) {
      var w = ctMin[i] / total * W;
      var r = svg('rect', { x: x, y: 7, width: Math.max(w - 1.5, 0), height: barH, rx: 3, fill: CYC_COLORS[i] });
      r.addEventListener('mousemove', function (e) { showTip('<b>' + CYC_LABELS[i] + '</b><br>' + fmt1(ctMin[i]) + ' мин · ' + Math.round(ctMin[i] / total * 100) + '% цикла', e); });
      r.addEventListener('mouseleave', hideTip);
      g.appendChild(r); x += w;
    });
  }

  function renderCycleTable(sum) {
    var c = sum.c, ctMin = sum.ct.map(function (v) { return v / c / 60; });
    var total = ctMin.reduce(function (a, b) { return a + b; }, 0);
    var rows = CYC_ORDER.map(function (i) {
      return '<div class="cyc-row"><span class="sw" style="background:' + CYC_COLORS[i] + '"></span>' +
        '<span class="nm">' + CYC_LABELS[i] + '</span>' +
        '<span class="mn">' + fmt1(ctMin[i]) + ' мин</span>' +
        '<span class="pc">' + Math.round(ctMin[i] / total * 100) + '%</span></div>';
    }).join('');
    var groups = cycGroups(ctMin).map(function (g) {
      return '<div class="cyc-group"><span class="gname"><span class="sw" style="background:' + g.color + '"></span>' + g.name + '</span>' +
        '<span class="gmn">' + fmt1(g.min) + ' мин</span><span class="gpc">' + Math.round(g.min / total * 100) + '%</span></div>';
    }).join('');
    el('cycle-table').innerHTML = rows + groups +
      '<div class="cyc-group" style="border-top-width:1px;"><span class="gname">Итого цикл</span><span class="gmn">' + fmt1(total) + ' мин</span><span class="gpc"></span></div>';

    var queue = ctMin[3] + ctMin[0];
    var queuePct = Math.round(queue / total * 100);
    el('cycle-callout').innerHTML =
      '<svg class="ic" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="9"/><path d="M12 7v5l3 2"/></svg>' +
      '<div>Ожидание составляет <b>' + fmt1(queue) + ' мин (' + queuePct + '% цикла)</b>, из них очередь у экскаватора — <b>' + fmt1(ctMin[3]) + ' мин</b>. ' +
      'Это главный резерв производительности: сокращение простоя в очереди напрямую поднимает т/ч без роста нагрузки на технику (лучшая практика — согласование числа самосвалов с производительностью экскаватора).</div>';
  }

  // донат: продуктивное / операции / ожидание
  function renderUtilDonut(sum) {
    var s = el('util-donut'); s.innerHTML = '';
    var c = sum.c, ctMin = sum.ct.map(function (v) { return v / c / 60; });
    var total = ctMin.reduce(function (a, b) { return a + b; }, 0);
    var groups = cycGroups(ctMin);
    var cx = 60, cy = 60, r = 48, rin = 30, ang = -Math.PI / 2;
    var g = svg('g', {}); s.appendChild(g);
    groups.forEach(function (gr) {
      if (gr.min <= 0) return;
      var a2 = ang + gr.min / total * Math.PI * 2, large = (a2 - ang) > Math.PI ? 1 : 0;
      var x1 = cx + r * Math.cos(ang), y1 = cy + r * Math.sin(ang), x2 = cx + r * Math.cos(a2), y2 = cy + r * Math.sin(a2);
      var xi1 = cx + rin * Math.cos(a2), yi1 = cy + rin * Math.sin(a2), xi2 = cx + rin * Math.cos(ang), yi2 = cy + rin * Math.sin(ang);
      var d = 'M' + x1 + ',' + y1 + ' A' + r + ',' + r + ' 0 ' + large + ' 1 ' + x2 + ',' + y2 + ' L' + xi1 + ',' + yi1 + ' A' + rin + ',' + rin + ' 0 ' + large + ' 0 ' + xi2 + ',' + yi2 + ' Z';
      var p = svg('path', { d: d, fill: gr.color, stroke: 'var(--surface)', 'stroke-width': 1.5 });
      p.addEventListener('mousemove', function (e) { showTip('<b>' + fmt1(gr.min / total * 100) + '%</b> · ' + fmt1(gr.min) + ' мин<br>' + gr.name, e); });
      p.addEventListener('mouseleave', hideTip);
      g.appendChild(p); ang = a2;
    });
    var prodPct = (groups[0].min) / total * 100;
    g.appendChild(svg('text', { x: cx, y: cy - 1, 'text-anchor': 'middle', 'font-family': 'var(--mono)', 'font-weight': '800', 'font-size': '17', fill: 'var(--ink)' })).textContent = Math.round(prodPct) + '%';
    g.appendChild(svg('text', { x: cx, y: cy + 12, 'text-anchor': 'middle', 'font-size': '8', fill: 'var(--ink-mute)' })).textContent = 'движение';
    el('util-legend').innerHTML = groups.map(function (gr) {
      return '<div class="item"><span class="sw" style="background:' + gr.color + '"></span><span class="nm">' + gr.name.replace(/ \(.*/, '') + '</span><span class="pc">' + Math.round(gr.min / total * 100) + '%</span></div>';
    }).join('');
  }

  // скорости: средняя ходовая (дистанция/время) и максимальная
  function renderSpeeds(sum) {
    var c = sum.c;
    var avgLoaded = sum.ct[1] > 0 ? sum.dl / sum.ct[1] * 3.6 : 0; // м / с → км/ч
    var avgEmpty = sum.ct[2] > 0 ? sum.de / sum.ct[2] * 3.6 : 0;
    var maxLoaded = sum.ls / c, maxEmpty = sum.es / c;
    var mx = Math.max(maxLoaded, maxEmpty, 1);
    var bars = [
      { nm: 'Гружёный ср.', v: avgLoaded, col: 'var(--target)' },
      { nm: 'Гружёный макс', v: maxLoaded, col: 'var(--target)', faint: true },
      { nm: 'Порожний ср.', v: avgEmpty, col: 'var(--under)' },
      { nm: 'Порожний макс', v: maxEmpty, col: 'var(--under)', faint: true }
    ];
    el('speedbars').innerHTML = bars.map(function (b) {
      return '<div class="sb"><span class="nm">' + b.nm + '</span>' +
        '<span class="track"><span style="width:' + (b.v / mx * 100).toFixed(0) + '%;background:' + b.col + (b.faint ? ';opacity:.45' : '') + '"></span></span>' +
        '<span class="vv">' + fmt1(b.v) + ' км/ч</span></div>';
    }).join('');
  }

  // TKPH по 4 колёсам
  function renderTKPH(sum) {
    var c = sum.c;
    var vals = { 'Лев. перед.': sum.lf / c, 'Прав. перед.': sum.rf / c, 'Лев. задн.': sum.lr / c, 'Прав. задн.': sum.rr / c };
    var arr = Object.keys(vals).map(function (k) { return { k: k, v: vals[k] }; });
    var mxv = Math.max.apply(null, arr.map(function (a) { return a.v; }));
    el('tkph-grid').innerHTML = arr.map(function (a) {
      var cls = a.v === mxv ? 'max' : a.v > mxv * 0.9 ? 'hot' : '';
      return '<div class="tkph-cell ' + cls + '"><span class="pos">' + a.k + '</span><span class="v">' + fmt1(a.v) + '</span></div>';
    }).join('');
    var rear = (vals['Лев. задн.'] + vals['Прав. задн.']) / 2, front = (vals['Лев. перед.'] + vals['Прав. перед.']) / 2;
    var ratio = front > 0 ? rear / front : 0;
    el('tkph-callout').innerHTML =
      '<svg class="ic" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="9"/><path d="M12 8v4M12 16h.01"/></svg>' +
      '<div>Задняя ось нагружена в <b>' + fmt1(ratio) + '×</b> сильнее передней — норма для самосвала (груз над задней осью). Наибольший TKPH — <b>' + fmt1(mxv) + '</b> (правое заднее колесо): именно оно ближе всего к температурному пределу шины, контролировать в первую очередь.</div>';
  }

  function renderHaul(sum) {
    var c = sum.c;
    var dl = sum.dl / c / 1000, de = sum.de / c / 1000;
    var cbPct = 0; // доля циклов с налипанием
    var carry = sum.cb || 0;
    cbPct = carry / c * 100;
    el('haul-metrics').innerHTML =
      '<div class="m"><span class="l">Плечо гружёный</span><span class="v">' + fmt1(dl) + ' <span style="font-size:9.5px;color:var(--ink-mute);font-weight:600">км</span></span></div>' +
      '<div class="m"><span class="l">Плечо порожний</span><span class="v">' + fmt1(de) + ' <span style="font-size:9.5px;color:var(--ink-mute);font-weight:600">км</span></span></div>' +
      '<div class="m"><span class="l">Ход за цикл</span><span class="v">' + fmt1(dl + de) + ' <span style="font-size:9.5px;color:var(--ink-mute);font-weight:600">км</span></span></div>' +
      '<div class="m"><span class="l">Налипание</span><span class="v">' + fmt2(cbPct) + ' <span style="font-size:9.5px;color:var(--ink-mute);font-weight:600">% цикл.</span></span></div>';
    var mx = Math.max(dl, de, 1);
    el('distbars').innerHTML =
      '<div class="sb"><span class="nm">Гружёный</span><span class="track"><span style="width:' + (dl / mx * 100).toFixed(0) + '%;background:var(--target)"></span></span><span class="vv">' + fmt1(dl) + ' км</span></div>' +
      '<div class="sb"><span class="nm">Порожний</span><span class="track"><span style="width:' + (de / mx * 100).toFixed(0) + '%;background:var(--under)"></span></span><span class="vv">' + fmt1(de) + ' км</span></div>';
  }

  // per-truck ранжирование (общий помощник)
  function truckAgg() {
    var byTruck = {};
    activeParts().forEach(function (p) {
      var a = byTruck[p.t] || (byTruck[p.t] = { c: 0, sp: 0, ct: [0, 0, 0, 0, 0, 0] });
      a.c += p.c; a.sp += p.sp;
      for (var i = 0; i < 6; i++) a.ct[i] += p.ct[i];
    });
    return Object.keys(byTruck).map(function (t) {
      var a = byTruck[t], cyc = a.ct.reduce(function (s, v) { return s + v; }, 0) / a.c / 60;
      return { t: t, c: a.c, prod: cyc ? (a.sp / a.c) / (cyc / 60) : 0, queue: a.ct[3] / a.c / 60, cycle: cyc };
    }).filter(function (r) { return r.c >= 200; });
  }
  function rankList(elId, rows, valFn, unit, colFn) {
    var mx = rows.length ? Math.max.apply(null, rows.map(valFn)) : 1;
    el(elId).innerHTML = rows.map(function (r) {
      var v = valFn(r), col = colFn(r);
      return '<div class="r" data-t="' + r.t + '"><span class="id">№' + r.t + '</span>' +
        '<span class="bar"><span style="width:' + (v / mx * 100).toFixed(0) + '%;background:' + col + '"></span></span>' +
        '<span class="v" style="color:' + col + '">' + fmt1(v) + unit + '</span></div>';
    }).join('') || '<div style="font-size:11px;color:var(--ink-mute)">Нет бортов с достаточной выборкой</div>';
    Array.prototype.forEach.call(el(elId).querySelectorAll('.r'), function (row) {
      row.onclick = function () { state.selTrucks = new Set([row.dataset.t]); renderFilters(); renderAll(); };
    });
  }
  function renderPerfCompare() {
    var rows = truckAgg();
    rankList('prodtop', rows.slice().sort(function (a, b) { return b.prod - a.prod; }).slice(0, 7), function (r) { return r.prod; }, ' т/ч', function () { return 'var(--target)'; });
    rankList('queuetop', rows.slice().sort(function (a, b) { return b.queue - a.queue; }).slice(0, 7), function (r) { return r.queue; }, ' мин', function (r) { return r.queue > 8 ? 'var(--bad)' : r.queue > 6 ? 'var(--accept)' : 'var(--ink-mute)'; });
  }

  // ---------- monthly trend ----------
  function renderTrend() {
    var s = el('trend'); s.innerHTML = '';
    var W = 640, H = 150, m = { t: 12, r: 30, b: 20, l: 30 };
    var pw = W - m.l - m.r, ph = H - m.t - m.b;
    var byM = {};
    activeParts().forEach(function (p) {
      var a = byM[p.m] || (byM[p.m] = { c: 0, sp: 0 }); a.c += p.c; a.sp += p.sp;
    });
    var months = Object.keys(byM).sort();
    if (!months.length) return;
    var data = months.map(function (ym) { return { ym: ym, c: byM[ym].c, load: byM[ym].sp / byM[ym].c / NOMINAL * 100 }; });
    var maxC = Math.max.apply(null, data.map(function (d) { return d.c; }));
    var loadMin = 95, loadMax = 115;
    data.forEach(function (d) { loadMin = Math.min(loadMin, d.load); loadMax = Math.max(loadMax, d.load); });
    loadMin = Math.floor(loadMin - 1); loadMax = Math.ceil(loadMax + 1);
    var g = svg('g', {}); s.appendChild(g);
    var bw = pw / data.length;
    // bars (cycles)
    data.forEach(function (d, i) {
      var hh = d.c / maxC * ph * 0.9;
      var bx = m.l + i * bw + bw * 0.2, by = m.t + ph - hh;
      var rect = svg('rect', { x: bx, y: by, width: bw * 0.6, height: Math.max(hh, 1), rx: 2, fill: 'var(--gold)', 'fill-opacity': .5 });
      rect.addEventListener('mousemove', function (e) { showTip('<b>' + ymLabel(d.ym) + '</b><br>' + fmtInt(d.c) + ' циклов · загрузка ' + fmt1(d.load) + '%', e); });
      rect.addEventListener('mouseleave', hideTip);
      g.appendChild(rect);
    });
    // 100% baseline
    var y100 = m.t + ph - (100 - loadMin) / (loadMax - loadMin) * ph;
    g.appendChild(svg('line', { x1: m.l, x2: W - m.r, y1: y100, y2: y100, class: 'zline', stroke: 'var(--ink-mute)' }));
    // load line
    var xC = function (i) { return m.l + i * bw + bw / 2; };
    var yL = function (v) { return m.t + ph - (v - loadMin) / (loadMax - loadMin) * ph; };
    var path = '';
    data.forEach(function (d, i) { path += (i === 0 ? 'M' : 'L') + xC(i) + ',' + yL(d.load) + ' '; });
    g.appendChild(svg('path', { d: path, fill: 'none', stroke: 'var(--target)', 'stroke-width': 2, 'stroke-linejoin': 'round', 'stroke-linecap': 'round' }));
    data.forEach(function (d, i) {
      var c = svg('circle', { cx: xC(i), cy: yL(d.load), r: 3.5, fill: 'var(--target)', stroke: 'var(--surface)', 'stroke-width': 1.5 });
      c.addEventListener('mousemove', function (e) { showTip('<b>' + ymLabel(d.ym) + '</b><br>средняя загрузка ' + fmt1(d.load) + '% · ' + fmtInt(d.c) + ' циклов', e); });
      c.addEventListener('mouseleave', hideTip);
      g.appendChild(c);
      if (data.length <= 12) {
        var tx = svg('text', { x: xC(i), y: H - 5, 'text-anchor': 'middle', class: 'ax' }); tx.textContent = MONTHS_RU[d.ym.split('-')[1]]; g.appendChild(tx);
      }
    });
  }

  // ---------- foot ----------
  function renderFoot() {
    el('foot').innerHTML = '<b>Политика 10/10/20 (Caterpillar):</b> средняя загрузка не выше номинала (180 т); не более 10% загрузок превышают 110% номинала (198 т); ни одна загрузка не превышает 120% (216 т). ' +
      'Номинал 180 т. Циклы дедуплицированы по ключу «борт + дата/время + № цикла». ' +
      'Кнопка «Загрузить .xlsx» пересчитывает всё в браузере из выбранных весовых файлов; «Сохранить» выгружает обработанный набор в JSON; данные автосохраняются локально в этом браузере.';
  }

  // ---------- master render ----------
  function renderAll() {
    var parts = activeParts();
    var sum = summarize(parts);
    if (sum.c === 0) { sum = summarize(state.partitions); }
    renderStatus();
    // вкладка «Загрузка»
    renderKPIs(sum); renderHist(sum); renderPolicy(sum); renderDonut(sum); renderTrend();
    // вкладка «Производительность»
    renderPerfKPIs(sum); renderCycleBar(sum); renderCycleTable(sum); renderUtilDonut(sum);
    renderSpeeds(sum); renderTKPH(sum); renderHaul(sum); renderPerfCompare();
    renderFoot();
  }

  // ---------- persistence ----------
  function persist() {
    try { localStorage.setItem(LS_KEY, JSON.stringify({ partitions: state.partitions, source: state.source, files: state.files })); } catch (e) {}
  }

  // ---------- data ingestion ----------
  var busy = el('busy');
  function ingestFiles(fileList) {
    if (!fileList || !fileList.length) return;
    busy.classList.add('on');
    if (!state.store) { state.store = AGG.makeStore(); state.files = []; }
    var arr = Array.prototype.slice.call(fileList);
    var done = 0, addedTotal = 0;
    setTimeout(function () {
      arr.forEach(function (file) {
        var reader = new FileReader();
        reader.onload = function (ev) {
          try {
            var wb = XLSX.read(new Uint8Array(ev.target.result), { type: 'array' });
            var got = 0;
            wb.SheetNames.forEach(function (sn) {
              if (got) return;
              var rows = XLSX.utils.sheet_to_json(wb.Sheets[sn], { header: 1, blankrows: false, raw: true });
              for (var i = 0; i < Math.min(rows.length, 6); i++) {
                if (rows[i] && rows[i].indexOf && rows[i].indexOf('TruckId') !== -1) { got = AGG.addRows2D(state.store, rows) + 1; break; }
              }
            });
            if (got) { addedTotal += (got - 1); state.files.push(file.name); }
          } catch (e) { console.error('parse error', file.name, e); }
          done++;
          if (done === arr.length) finishIngest(addedTotal);
        };
        reader.onerror = function () { done++; if (done === arr.length) finishIngest(addedTotal); };
        reader.readAsArrayBuffer(file);
      });
    }, 30);
  }
  function finishIngest(added) {
    state.partitions = AGG.finalize(state.store);
    state.source = 'Загружено: ' + state.files.length + ' файл(ов)';
    state.selMonths = null; state.selTrucks = null;
    persist();
    busy.classList.remove('on');
    renderFilters(); renderAll();
    if (!added) alert('В выбранных файлах не найдено новых уникальных циклов (возможно, эти данные уже загружены).');
  }

  function saveJSON() {
    var blob = new Blob([JSON.stringify({ nominal: NOMINAL, generated: new Date().toISOString().slice(0, 10), source: state.source, files: state.files, partitions: state.partitions })], { type: 'application/json' });
    var a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = 'weighing_dashboard_' + new Date().toISOString().slice(0, 10) + '.json';
    a.click(); setTimeout(function () { URL.revokeObjectURL(a.href); }, 1000);
  }
  function openJSON(file) {
    var reader = new FileReader();
    reader.onload = function (ev) {
      try {
        var obj = JSON.parse(ev.target.result);
        if (!obj.partitions || !obj.partitions.length) throw new Error('нет партиций');
        state.partitions = obj.partitions; state.store = null; state.files = obj.files || [];
        state.source = obj.source || 'Открыт файл данных';
        state.selMonths = null; state.selTrucks = null;
        persist(); renderFilters(); renderAll();
      } catch (e) { alert('Не удалось открыть файл: ' + e.message); }
    };
    reader.readAsText(file);
  }
  function resetData() {
    state.partitions = EMBEDDED.partitions; state.store = null; state.files = [];
    state.source = 'Демо-данные парка'; state.selMonths = null; state.selTrucks = null;
    try { localStorage.removeItem(LS_KEY); } catch (e) {}
    renderFilters(); renderAll();
  }

  // ---------- theme ----------
  function toggleTheme() {
    var cur = document.documentElement.getAttribute('data-theme');
    var next = cur === 'dark' ? 'light' : cur === 'light' ? 'dark' : (matchMedia('(prefers-color-scheme: dark)').matches ? 'light' : 'dark');
    document.documentElement.setAttribute('data-theme', next);
    renderAll();
  }

  function on(id, ev, fn) { var e = el(id); if (e) e.addEventListener(ev, fn); }

  // ---------- РЕНДЕР ПЕРВЫМ (чтобы сбой навешивания слушателей не оставлял пустой экран) ----------
  try {
    renderFilters();
    renderAll();
  } catch (err) {
    var st = el('status');
    if (st) st.innerHTML = '<span style="color:var(--bad)">Ошибка отрисовки: ' + (err && err.message ? err.message : err) + '</span>';
  }

  // ---------- wire up (после рендера, с защитой) ----------
  try {
    on('btn-load', 'click', function () { el('file-xlsx').click(); });
    on('file-xlsx', 'change', function (e) { ingestFiles(e.target.files); e.target.value = ''; });
    on('btn-save', 'click', saveJSON);
    on('btn-open', 'click', function () { el('file-json').click(); });
    on('file-json', 'change', function (e) { if (e.target.files[0]) openJSON(e.target.files[0]); e.target.value = ''; });
    on('btn-reset', 'click', resetData);
    on('btn-theme', 'click', toggleTheme);
    on('month-all', 'click', function () { state.selMonths = null; renderFilters(); renderAll(); });
    on('truck-all', 'click', function () { state.selTrucks = null; renderFilters(); renderAll(); });

    window.addEventListener('mousemove', function (e) { if (tip.classList.contains('on')) moveTip(e); });
    var rz; window.addEventListener('resize', function () { clearTimeout(rz); rz = setTimeout(renderAll, 150); });

    var mq = window.matchMedia ? window.matchMedia('(prefers-color-scheme: dark)') : null;
    if (mq) {
      if (mq.addEventListener) mq.addEventListener('change', renderAll);
      else if (mq.addListener) mq.addListener(renderAll); // старые Safari/iOS
    }
  } catch (err) { /* слушатели не критичны для отображения */ }
})();
