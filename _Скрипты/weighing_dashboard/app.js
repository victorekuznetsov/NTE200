/* Дашборд весового контроля NTE200 — приложение.
   Работает поверх партиций (борт × месяц) из agg.js. Всё считается на клиенте,
   поэтому загруженные .xlsx обрабатываются той же логикой, что и встроенные данные. */
(function () {
  'use strict';
  var NOMINAL = 173.5, P110 = 190.85, P120 = 208.2;
  var TIRE_TKPH = 1000, TIRE_MAXZONE = 950; // предел ТКВЧ шины (предварительно) и порог max-зоны
  var MONTHS_RU = { '01':'янв','02':'фев','03':'мар','04':'апр','05':'май','06':'июн','07':'июл','08':'авг','09':'сен','10':'окт','11':'ноя','12':'дек' };
  var LS_KEY = 'nte200_weighing_v1';
  var Z_COLORS = ['var(--under)','var(--target)','var(--accept)','var(--crit)'];
  var Z_NAMES = ['Недогруз (<173,5 т)','Целевая (173,5–190,9 т)','Перегруз доп. (190,9–208,2 т)','Перегруз крит. (>208,2 т)'];

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
        dl = 0, de = 0, ls = 0, es = 0, lf = 0, rf = 0, lr = 0, rr = 0, cb = 0, f = 0, tk = 0, bk = 0, thr = {};
    parts.forEach(function (p) {
      c += p.c; sp += p.sp;
      for (var i = 0; i < 4; i++) z[i] += p.z[i];
      for (var b in p.h) h[b] = (h[b] || 0) + p.h[b];
      for (var j = 0; j < 6; j++) ct[j] += p.ct[j];
      dl += p.dl; de += p.de; ls += p.ls; es += p.es;
      lf += (p.lf || 0); rf += (p.rf || 0); lr += p.lr; rr += p.rr; cb += (p.cb || 0); f += p.f;
      tk += (p.tk || 0); bk += (p.bk || 0);
      if (p.thr) for (var tb in p.thr) thr[tb] = (thr[tb] || 0) + p.thr[tb];
    });
    return { c: c, sp: sp, z: z, h: h, ct: ct, dl: dl, de: de, ls: ls, es: es, lf: lf, rf: rf, lr: lr, rr: rr, cb: cb, f: f, tk: tk, bk: bk, thr: thr };
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
    var BIN = 1, LO = 145, HI = 224;
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
    // zone boundary lines (173.5 / 190.85 / 208.2)
    [[NOMINAL, 'ном. 173,5'], [P110, '110% · 190,9'], [P120, '120% · 208,2']].forEach(function (zb) {
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
      { ok: mean <= NOMINAL, t1: 'Средняя загрузка ≤ номинала', t2: 'среднее по распределению ≤ 173,5 т', num: fmt1(mean) + ' т' },
      { ok: over110 <= 10, t1: 'Не более 10% загрузок > 110%', t2: 'доля циклов свыше 190,9 т', num: fmt1(over110) + ' %' },
      { ok: over120 === 0, t1: 'Ни одной загрузки > 120%', t2: 'доля циклов свыше 208,2 т (' + fmtInt(sum.z[3]) + ' шт.)', num: fmt2(over120) + ' %' }
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

  // TKPH (ТКВЧ) по 4 колёсам — сырые значения, тонно-км/ч
  function renderTKPH(sum) {
    var c = sum.c;
    var vals = { 'Лев. перед.': sum.lf / c, 'Прав. перед.': sum.rf / c, 'Лев. задн.': sum.lr / c, 'Прав. задн.': sum.rr / c };
    var arr = Object.keys(vals).map(function (k) { return { k: k, v: vals[k] }; });
    var mxv = Math.max.apply(null, arr.map(function (a) { return a.v; }));
    el('tkph-grid').innerHTML = arr.map(function (a) {
      var cls = a.v > TIRE_TKPH ? 'max' : a.v > TIRE_MAXZONE ? 'hot' : '';
      return '<div class="tkph-cell ' + cls + '"><span class="pos">' + a.k + '</span><span class="v">' + fmtInt(a.v) + '</span></div>';
    }).join('');
    var rear = (vals['Лев. задн.'] + vals['Прав. задн.']) / 2, front = (vals['Лев. перед.'] + vals['Прав. перед.']) / 2;
    var ratio = front > 0 ? rear / front : 0;
    var overRear = rear > TIRE_TKPH;
    el('tkph-callout').innerHTML =
      '<svg class="ic" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="9"/><path d="M12 8v4M12 16h.01"/></svg>' +
      '<div>Средняя нагрузка задней оси <b>' + fmtInt(rear) + '</b> т·км/ч — в <b>' + fmt1(ratio) + '×</b> выше передней. ' +
      (overRear ? 'Это <b style="color:var(--bad)">выше предела шины ' + TIRE_TKPH + '</b> — задние шины работают в зоне перегрева, контроль в первую очередь.' :
        'Предел шины ' + TIRE_TKPH + ' т·км/ч.') +
      ' <span style="color:var(--ink-mute)">Предел ТКВЧ 1000 задан предварительно — уточнить по паспорту шины NTE200.</span></div>';
  }

  // Средняя нагрузка на шину (ТКВЧ) по месяцам, с пределом шины
  function renderTkphTrend() {
    var s = el('tkph-trend'); if (!s) return; s.innerHTML = '';
    var byM = {};
    activeParts().forEach(function (p) {
      var a = byM[p.m] || (byM[p.m] = { c: 0, lf: 0, rf: 0, lr: 0, rr: 0 });
      a.c += p.c; a.lf += p.lf; a.rf += p.rf; a.lr += p.lr; a.rr += p.rr;
    });
    var data = Object.keys(byM).sort().map(function (ym) {
      var a = byM[ym];
      return { ym: ym, front: (a.lf + a.rf) / 2 / a.c, rear: (a.lr + a.rr) / 2 / a.c };
    });
    if (!data.length) return;
    var W = 560, H = 210, m = { t: 16, r: 12, b: 24, l: 40 };
    var pw = W - m.l - m.r, ph = H - m.t - m.b;
    var vmax = Math.max(TIRE_TKPH * 1.15, Math.max.apply(null, data.map(function (d) { return Math.max(d.front, d.rear); })) * 1.1);
    var g = svg('g', {}); s.appendChild(g);
    for (var i = 0; i <= 4; i++) { var yy = m.t + ph - i / 4 * ph, val = vmax * i / 4; g.appendChild(svg('line', { x1: m.l, x2: W - m.r, y1: yy, y2: yy, class: 'gl' })); var tl = svg('text', { x: m.l - 5, y: yy + 3, 'text-anchor': 'end', class: 'ax' }); tl.textContent = Math.round(val); g.appendChild(tl); }
    var xC = function (i) { return m.l + (data.length === 1 ? pw / 2 : i / (data.length - 1) * pw); };
    var yV = function (v) { return m.t + ph - v / vmax * ph; };
    // предел шины
    var yl = yV(TIRE_TKPH);
    g.appendChild(svg('line', { x1: m.l, x2: W - m.r, y1: yl, y2: yl, stroke: 'var(--crit)', 'stroke-width': 1.4, 'stroke-dasharray': '5 3' }));
    var ll = svg('text', { x: W - m.r, y: yl - 4, 'text-anchor': 'end', class: 'axb', fill: 'var(--crit)' }); ll.textContent = 'предел шины ' + TIRE_TKPH; g.appendChild(ll);
    [['rear', 'var(--crit)', 'задняя ось'], ['front', 'var(--under)', 'передняя ось']].forEach(function (ser) {
      var path = ''; data.forEach(function (d, i) { path += (i === 0 ? 'M' : 'L') + xC(i).toFixed(1) + ',' + yV(d[ser[0]]).toFixed(1) + ' '; });
      g.appendChild(svg('path', { d: path, fill: 'none', stroke: ser[1], 'stroke-width': 2, 'stroke-linejoin': 'round', 'stroke-linecap': 'round' }));
      data.forEach(function (d, i) {
        var cc = svg('circle', { cx: xC(i), cy: yV(d[ser[0]]), r: 2.8, fill: ser[1], stroke: 'var(--surface)', 'stroke-width': 1.2 });
        cc.addEventListener('mousemove', function (e) { showTip('<b>' + ymLabel(d.ym) + '</b><br>' + ser[2] + ': ' + fmtInt(d[ser[0]]) + ' т·км/ч', e); });
        cc.addEventListener('mouseleave', hideTip); g.appendChild(cc);
      });
    });
    data.forEach(function (d, i) { if (i % Math.ceil(data.length / 10 || 1) === 0) { var tx = svg('text', { x: xC(i), y: H - 6, 'text-anchor': 'middle', class: 'ax' }); tx.textContent = MONTHS_RU[d.ym.split('-')[1]]; g.appendChild(tx); } });
  }

  // Распределение ТКВЧ переднего правого колеса, с max-зоной
  function renderTkphDist(sum) {
    var s = el('tkph-dist'); if (!s) return; s.innerHTML = '';
    var thr = sum.thr || {};
    var keys = Object.keys(thr).map(Number).sort(function (a, b) { return a - b; });
    if (!keys.length) return;
    var total = keys.reduce(function (t, k) { return t + thr[k]; }, 0);
    var W = 560, H = 210, m = { t: 14, r: 12, b: 26, l: 34 };
    var pw = W - m.l - m.r, ph = H - m.t - m.b;
    var LO = 0, HI = Math.max(1250, keys[keys.length - 1] + 50);
    var maxPct = Math.max.apply(null, keys.map(function (k) { return thr[k] / total * 100; }));
    var yMax = Math.ceil(maxPct / 2) * 2 || 2;
    var xOf = function (v) { return m.l + (v - LO) / (HI - LO) * pw; };
    var yOf = function (pct) { return m.t + ph - pct / yMax * ph; };
    var barW = pw / ((HI - LO) / 50);
    var g = svg('g', {}); s.appendChild(g);
    for (var i = 0; i <= 2; i++) { var yy = yOf(yMax * i / 2); g.appendChild(svg('line', { x1: m.l, x2: W - m.r, y1: yy, y2: yy, class: 'gl' })); var tl = svg('text', { x: m.l - 5, y: yy + 3, 'text-anchor': 'end', class: 'ax' }); tl.textContent = Math.round(yMax * i / 2) + '%'; g.appendChild(tl); }
    var over = 0;
    keys.forEach(function (k) {
      var pct = thr[k] / total * 100, hh = m.t + ph - yOf(pct);
      var isMax = k >= TIRE_MAXZONE; if (isMax) over += thr[k];
      var r = svg('rect', { x: xOf(k) + 0.5, y: yOf(pct), width: Math.max(barW - 1, 1), height: Math.max(hh, 0.5), rx: 1.5, fill: isMax ? 'var(--crit)' : 'var(--target)' });
      r.addEventListener('mousemove', function (e) { showTip('<b>' + k + '–' + (k + 50) + ' т·км/ч</b><br>' + fmt1(pct) + '% · ' + fmtInt(thr[k]) + ' циклов' + (isMax ? '<br>max-зона (>' + TIRE_MAXZONE + ')' : ''), e); });
      r.addEventListener('mouseleave', hideTip); g.appendChild(r);
    });
    // предел шины
    var lx = xOf(TIRE_TKPH);
    g.appendChild(svg('line', { x1: lx, x2: lx, y1: m.t, y2: m.t + ph, stroke: 'var(--crit)', 'stroke-width': 1.3, 'stroke-dasharray': '4 3' }));
    var lb = svg('text', { x: lx, y: m.t + 8, 'text-anchor': 'middle', class: 'axb', fill: 'var(--crit)' }); lb.textContent = '' + TIRE_TKPH; g.appendChild(lb);
    for (var xv = 0; xv <= HI; xv += 250) { var tx = svg('text', { x: xOf(xv), y: H - 6, 'text-anchor': 'middle', class: 'ax' }); tx.textContent = xv; g.appendChild(tx); }
    var cap = el('tkph-dist-cap'); if (cap) cap.innerHTML = 'переднее правое колесо · в max-зоне (>' + TIRE_MAXZONE + ' т·км/ч): <b>' + fmt1(over / total * 100) + '%</b> циклов';
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

  // ---------- ВКЛАДКА 3: ГРУЗООБОРОТ И СКОРОСТИ ----------
  function renderFlowKPIs(sum) {
    var c = sum.c;
    var cycH = sum.ct.reduce(function (s, v) { return s + v; }, 0) / 3600; // всего цикло-часов
    var tkmH = cycH ? sum.tk / cycH : 0;
    var tiles = [
      { l: 'Грузооборот', v: fmt1(sum.tk / 1e6), u: 'млн т·км', big: true },
      { l: 'Ton·km/ч', v: fmtInt(tkmH), u: 'т·км/ч', big: true },
      { l: 'Ковшей на погрузку', v: fmt1(sum.bk / c), u: 'ковш.' },
      { l: 'Ср. время погрузки', v: fmt1(sum.ct[4] / c / 60), u: 'мин' },
      { l: 'Ср. скорость гружёным', v: fmt1(sum.ct[1] > 0 ? sum.dl / sum.ct[1] * 3.6 : 0), u: 'км/ч' }
    ];
    el('flow-kpis').innerHTML = tiles.map(function (t) {
      return '<div class="kpi"><div class="lbl">' + t.l + '</div><div class="val' + (t.big ? ' big' : '') + '">' + t.v + '<span class="u">' + t.u + '</span></div></div>';
    }).join('');
  }

  function monthAgg() {
    var byM = {};
    activeParts().forEach(function (p) {
      var a = byM[p.m] || (byM[p.m] = { c: 0, sp: 0, tk: 0, dl: 0, de: 0, ct: [0, 0, 0, 0, 0, 0], z: [0, 0, 0, 0] });
      a.c += p.c; a.sp += p.sp; a.tk += (p.tk || 0); a.dl += p.dl; a.de += p.de;
      for (var i = 0; i < 6; i++) a.ct[i] += p.ct[i];
      for (var j = 0; j < 4; j++) a.z[j] += p.z[j];
    });
    return Object.keys(byM).sort().map(function (ym) { var a = byM[ym]; a.ym = ym; return a; });
  }

  // грузооборот по месяцам (bars) + погрузок (line)
  function renderFlowTrend() {
    var s = el('flow-trend'); s.innerHTML = '';
    var data = monthAgg(); if (!data.length) return;
    var W = 640, H = 220, m = { t: 22, r: 40, b: 34, l: 46 };
    var pw = W - m.l - m.r, ph = H - m.t - m.b;
    var maxTk = Math.max.apply(null, data.map(function (d) { return d.tk; })) || 1;
    var maxC = Math.max.apply(null, data.map(function (d) { return d.c; })) || 1;
    var bw = pw / data.length, g = svg('g', {}); s.appendChild(g);
    // y grid (тыс. т·км)
    for (var i = 0; i <= 3; i++) {
      var yy = m.t + ph - i / 3 * ph, val = maxTk * i / 3 / 1000;
      g.appendChild(svg('line', { x1: m.l, x2: W - m.r, y1: yy, y2: yy, class: 'gl' }));
      var tl = svg('text', { x: m.l - 5, y: yy + 3, 'text-anchor': 'end', class: 'ax' }); tl.textContent = Math.round(val); g.appendChild(tl);
    }
    var ytl = svg('text', { x: m.l - 5, y: m.t - 8, 'text-anchor': 'end', class: 'ax' }); ytl.textContent = 'тыс.т·км'; g.appendChild(ytl);
    data.forEach(function (d, i) {
      var hh = d.tk / maxTk * ph, bx = m.l + i * bw + bw * 0.18, by = m.t + ph - hh;
      var r = svg('rect', { x: bx, y: by, width: bw * 0.64, height: Math.max(hh, 1), rx: 2, fill: 'var(--gold)', 'fill-opacity': .55 });
      r.addEventListener('mousemove', function (e) { showTip('<b>' + ymLabel(d.ym) + '</b><br>' + fmtInt(d.tk / 1000) + ' тыс. т·км<br>' + fmtInt(d.c) + ' погрузок', e); });
      r.addEventListener('mouseleave', hideTip); g.appendChild(r);
    });
    var xC = function (i) { return m.l + i * bw + bw / 2; };
    var yR = function (v) { return m.t + ph - v / maxC * ph * 0.92; };
    var path = ''; data.forEach(function (d, i) { path += (i === 0 ? 'M' : 'L') + xC(i) + ',' + yR(d.c) + ' '; });
    g.appendChild(svg('path', { d: path, fill: 'none', stroke: 'var(--ink)', 'stroke-width': 1.8, 'stroke-opacity': .6, 'stroke-linejoin': 'round' }));
    data.forEach(function (d, i) {
      g.appendChild(svg('circle', { cx: xC(i), cy: yR(d.c), r: 2.6, fill: 'var(--ink)', 'fill-opacity': .6 }));
      if (i % Math.ceil(data.length / 12 || 1) === 0) { var tx = svg('text', { x: xC(i), y: H - 6, 'text-anchor': 'middle', class: 'ax' }); tx.textContent = MONTHS_RU[d.ym.split('-')[1]]; g.appendChild(tx); }
    });
  }

  // средняя скорость (гружёный/порожний) по месяцам
  function renderSpeedTrend() {
    var s = el('speed-trend'); s.innerHTML = '';
    var data = monthAgg(); if (!data.length) return;
    var W = 480, H = 220, m = { t: 14, r: 12, b: 34, l: 34 };
    var pw = W - m.l - m.r, ph = H - m.t - m.b;
    var series = data.map(function (d) { return { ym: d.ym, ld: d.ct[1] > 0 ? d.dl / d.ct[1] * 3.6 : 0, em: d.ct[2] > 0 ? d.de / d.ct[2] * 3.6 : 0 }; });
    var vmax = 0, vmin = 99; series.forEach(function (d) { vmax = Math.max(vmax, d.ld, d.em); vmin = Math.min(vmin, d.ld, d.em); });
    vmin = Math.floor(vmin - 2); vmax = Math.ceil(vmax + 2);
    var bw = pw / series.length, g = svg('g', {}); s.appendChild(g);
    for (var i = 0; i <= 3; i++) { var yy = m.t + ph - i / 3 * ph, val = vmin + (vmax - vmin) * i / 3; g.appendChild(svg('line', { x1: m.l, x2: W - m.r, y1: yy, y2: yy, class: 'gl' })); var tl = svg('text', { x: m.l - 5, y: yy + 3, 'text-anchor': 'end', class: 'ax' }); tl.textContent = Math.round(val); g.appendChild(tl); }
    var xC = function (i) { return m.l + i * bw + bw / 2; };
    var yV = function (v) { return m.t + ph - (v - vmin) / (vmax - vmin) * ph; };
    [['ld', 'var(--target)'], ['em', 'var(--under)']].forEach(function (ser) {
      var path = ''; series.forEach(function (d, i) { path += (i === 0 ? 'M' : 'L') + xC(i) + ',' + yV(d[ser[0]]) + ' '; });
      g.appendChild(svg('path', { d: path, fill: 'none', stroke: ser[1], 'stroke-width': 2, 'stroke-linejoin': 'round', 'stroke-linecap': 'round' }));
      series.forEach(function (d, i) {
        var c = svg('circle', { cx: xC(i), cy: yV(d[ser[0]]), r: 3, fill: ser[1], stroke: 'var(--surface)', 'stroke-width': 1.3 });
        c.addEventListener('mousemove', function (e) { showTip('<b>' + ymLabel(d.ym) + '</b><br>' + (ser[0] === 'ld' ? 'гружёный' : 'порожний') + ' ' + fmt1(d[ser[0]]) + ' км/ч', e); });
        c.addEventListener('mouseleave', hideTip); g.appendChild(c);
      });
    });
    series.forEach(function (d, i) { if (i % Math.ceil(series.length / 10 || 1) === 0) { var tx = svg('text', { x: xC(i), y: H - 6, 'text-anchor': 'middle', class: 'ax' }); tx.textContent = MONTHS_RU[d.ym.split('-')[1]]; g.appendChild(tx); } });
  }

  function renderFlowCompare() {
    var byTruck = {};
    activeParts().forEach(function (p) {
      var a = byTruck[p.t] || (byTruck[p.t] = { c: 0, dl: 0, tk: 0, ct1: 0 });
      a.c += p.c; a.dl += p.dl; a.tk += (p.tk || 0); a.ct1 += p.ct[1];
    });
    var rows = Object.keys(byTruck).map(function (t) { var a = byTruck[t]; return { t: t, c: a.c, spd: a.ct1 > 0 ? a.dl / a.ct1 * 3.6 : 0, tk: a.tk }; }).filter(function (r) { return r.c >= 200; });
    rankList('speedtop', rows.slice().sort(function (a, b) { return b.spd - a.spd; }).slice(0, 7), function (r) { return r.spd; }, ' км/ч', function () { return 'var(--target)'; });
    rankList('flowtop', rows.slice().sort(function (a, b) { return b.tk - a.tk; }).slice(0, 7), function (r) { return r.tk / 1e6; }, ' млн', function () { return 'var(--gold-deep)'; });
  }

  // ---------- ВКЛАДКА 4: СВОДНАЯ ТАБЛИЦА ----------
  var TCOLS = [
    { k: 't', n: 'Борт', num: false },
    { k: 'c', n: 'Циклы' },
    { k: 'mean', n: 'Ср.загр,т' },
    { k: 'tot', n: 'Перевезено,тыс.т' },
    { k: 'dist', n: 'Плечо,км' },
    { k: 'tk', n: 'Грузооб,тыс.т·км' },
    { k: 'prod', n: 'Произв,т/ч' },
    { k: 'tkmh', n: 'т·км/ч' },
    { k: 'bk', n: 'Ковш/погр' },
    { k: 'cyc', n: 'Цикл,мин' },
    { k: 'queue', n: 'Очередь,мин' },
    { k: 'rr', n: 'TKPH зад' },
    { k: 'ov110', n: '>110%,%' },
    { k: 'ov120', n: '>120%,%' }
  ];
  var tSortKey = 'tk', tSortDir = -1;
  function fleetTableRows() {
    var byTruck = {};
    activeParts().forEach(function (p) {
      var a = byTruck[p.t] || (byTruck[p.t] = { c: 0, sp: 0, tk: 0, bk: 0, dl: 0, rr: 0, z: [0, 0, 0, 0], ct: [0, 0, 0, 0, 0, 0] });
      a.c += p.c; a.sp += p.sp; a.tk += (p.tk || 0); a.bk += (p.bk || 0); a.dl += p.dl; a.rr += p.rr;
      for (var i = 0; i < 4; i++) a.z[i] += p.z[i];
      for (var j = 0; j < 6; j++) a.ct[j] += p.ct[j];
    });
    return Object.keys(byTruck).map(function (t) {
      var a = byTruck[t], cyc = a.ct.reduce(function (s, v) { return s + v; }, 0) / a.c / 60;
      var cycH = a.ct.reduce(function (s, v) { return s + v; }, 0) / 3600;
      return {
        t: t, c: a.c, mean: a.sp / a.c, tot: a.sp / 1000, dist: a.dl / a.c / 1000,
        tk: a.tk / 1000, prod: cyc ? (a.sp / a.c) / (cyc / 60) : 0, tkmh: cycH ? a.tk / cycH : 0,
        bk: a.bk / a.c, cyc: cyc, queue: a.ct[3] / a.c / 60, rr: a.rr / a.c,
        ov110: (a.z[2] + a.z[3]) / a.c * 100, ov120: a.z[3] / a.c * 100
      };
    });
  }
  function renderFleetTable() {
    var rows = fleetTableRows();
    el('fleet-thead').innerHTML = TCOLS.map(function (c) {
      return '<th data-k="' + c.k + '" class="' + (c.k === tSortKey ? 'act' : '') + '">' + c.n + '</th>';
    }).join('');
    rows.sort(function (a, b) { var av = a[tSortKey], bv = b[tSortKey]; if (tSortKey === 't') { av = +a.t; bv = +b.t; } return (av - bv) * tSortDir; });
    function cell(v, k) {
      if (k === 't') return '№' + v;
      if (k === 'c') return fmtInt(v);
      if (k === 'tot' || k === 'tk') return fmtInt(v);
      if (k === 'tkmh') return fmtInt(v);
      if (k === 'ov110' || k === 'ov120') return fmt1(v);
      return fmt1(v);
    }
    el('fleet-tbody').innerHTML = rows.map(function (r) {
      return '<tr data-t="' + r.t + '">' + TCOLS.map(function (c) {
        var col = '';
        if (c.k === 'ov110' && r.ov110 > 10) col = 'color:var(--bad)';
        else if (c.k === 'ov110' && r.ov110 > 6) col = 'color:var(--accept)';
        if (c.k === 'ov120' && r.ov120 > 0) col = 'color:var(--bad)';
        return '<td style="' + col + '">' + cell(r[c.k], c.k) + '</td>';
      }).join('') + '</tr>';
    }).join('');
    // footer totals / weighted means
    var sum = summarize(activeParts()), C = sum.c;
    var cyc = sum.ct.reduce(function (s, v) { return s + v; }, 0) / C / 60;
    var cycH = sum.ct.reduce(function (s, v) { return s + v; }, 0) / 3600;
    var foot = {
      t: 'Парк', c: C, mean: sum.sp / C, tot: sum.sp / 1000, dist: sum.dl / C / 1000, tk: sum.tk / 1000,
      prod: cyc ? (sum.sp / C) / (cyc / 60) : 0, tkmh: cycH ? sum.tk / cycH : 0, bk: sum.bk / C, cyc: cyc,
      queue: sum.ct[3] / C / 60, rr: sum.rr / C, ov110: (sum.z[2] + sum.z[3]) / C * 100, ov120: sum.z[3] / C * 100
    };
    el('fleet-tfoot').innerHTML = '<tr>' + TCOLS.map(function (c) { return '<td>' + cell(foot[c.k], c.k) + '</td>'; }).join('') + '</tr>';
    Array.prototype.forEach.call(el('fleet-thead').querySelectorAll('th'), function (th) {
      th.onclick = function () { var k = th.dataset.k; if (tSortKey === k) tSortDir *= -1; else { tSortKey = k; tSortDir = k === 't' ? 1 : -1; } renderFleetTable(); };
    });
    Array.prototype.forEach.call(el('fleet-tbody').querySelectorAll('tr'), function (tr) {
      tr.onclick = function () { state.selTrucks = new Set([tr.dataset.t]); renderFilters(); renderAll(); };
      tr.classList.toggle('sel', state.selTrucks && state.selTrucks.size === 1 && state.selTrucks.has(tr.dataset.t));
    });
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
    el('foot').innerHTML = '<b>Политика 10/10/20 (Caterpillar):</b> средняя загрузка не выше номинала (173,5 т); не более 10% загрузок превышают 110% номинала (190,9 т); ни одна загрузка не превышает 120% (208,2 т). ' +
      'Номинал 173,5 т. ТКВЧ — сырые значения поля весовой (т·км/ч), предел шины 1000 задан предварительно (уточнить по паспорту шины NTE200). Грузооборот = загрузка × плечо гружёного; т/ч и т·км/ч считаются по времени цикла (не по моточасам). Циклы дедуплицированы по ключу «борт + дата/время + № цикла». ' +
      'Все показатели — из бортовых весовых файлов. Разделы отчёта Северстали по <b>моточасам (наработка), топливу (л, г/т·км), blowby и кодам ошибок</b> здесь не показаны — они требуют выгрузки VHMS/сервис-метра машины, которой нет в весовых данных. ' +
      'Кнопка «Загрузить .xlsx» пересчитывает всё в браузере; «Сохранить» выгружает набор в JSON; данные автосохраняются локально.';
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
    renderSpeeds(sum); renderTKPH(sum); renderTkphTrend(); renderTkphDist(sum); renderHaul(sum); renderPerfCompare();
    // вкладка «Грузооборот и скорости»
    renderFlowKPIs(sum); renderFlowTrend(); renderSpeedTrend(); renderFlowCompare();
    // вкладка «Сводная по бортам»
    renderFleetTable();
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
    var text = JSON.stringify({ nominal: NOMINAL, generated: new Date().toISOString().slice(0, 10), source: state.source, files: state.files, partitions: state.partitions });
    var fname = 'weighing_dashboard_' + new Date().toISOString().slice(0, 10) + '.json';
    var blob = new Blob([text], { type: 'application/json' });
    // iOS Safari / WKWebView: navigator.share с файлом — самый надёжный путь «сохранить в Файлы»
    try {
      if (navigator.canShare && window.File) {
        var file = new File([blob], fname, { type: 'application/json' });
        if (navigator.canShare({ files: [file] })) { navigator.share({ files: [file], title: fname }).catch(function () { }); return; }
      }
    } catch (e) { }
    // обычная загрузка: ссылка ДОЛЖНА быть в DOM, иначе Safari игнорирует click()
    var url = URL.createObjectURL(blob);
    var a = document.createElement('a');
    a.href = url; a.download = fname; a.rel = 'noopener'; a.style.display = 'none';
    document.body.appendChild(a);
    a.click();
    setTimeout(function () { document.body.removeChild(a); URL.revokeObjectURL(url); }, 4000);
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
