/* Единый модуль агрегации данных весового устройства NTE200.
 * Работает и в Node (сборка встроенных данных), и в браузере (загрузка новых .xlsx).
 * Каждый цикл дедуплицируется по ключу «борт+год+месяц+день+час+минута+секунда+№цикла».
 * Итог — партиции по (борт, год-месяц) с аддитивными суммами: их можно свободно
 * складывать при фильтрации по борту и периоду. Номинал 180 т.
 */
(function (root, factory) {
  if (typeof module !== 'undefined' && module.exports) module.exports = factory();
  else root.AGG = factory();
})(typeof self !== 'undefined' ? self : this, function () {
  var NOMINAL = 180;
  var P110 = NOMINAL * 1.10; // 198
  var P120 = NOMINAL * 1.20; // 216

  function makeStore() {
    return { keys: Object.create(null), partitions: Object.create(null), n: 0 };
  }

  function findHeader(rows) {
    for (var i = 0; i < Math.min(rows.length, 6); i++) {
      var r = rows[i];
      if (r && r.indexOf && r.indexOf('TruckId') !== -1) return { header: r, start: i + 1 };
    }
    return null;
  }

  // rows2d: массив массивов (одна вкладка книги). Возвращает число новых уникальных циклов.
  function addRows2D(store, rows2d) {
    var hd = findHeader(rows2d);
    if (!hd) return 0;
    var idx = {};
    hd.header.forEach(function (h, i) { idx[h] = i; });
    var need = ['FinalPayload', 'LoadPercentage', 'TruckId', 'ear', 'Month', 'Day', 'Hour', 'Minute', 'Second', 'PayloadNumber'];
    for (var k = 0; k < need.length; k++) if (!(need[k] in idx)) return 0;

    var added = 0;
    for (var r = hd.start; r < rows2d.length; r++) {
      var row = rows2d[r];
      if (!row) continue;
      var fp = row[idx.FinalPayload], lp = row[idx.LoadPercentage];
      if (!fp || !lp) continue;
      var truck = String(row[idx.TruckId]);
      var y = row[idx.ear], mo = row[idx.Month], d = row[idx.Day],
          h = row[idx.Hour], mi = row[idx.Minute], s = row[idx.Second], pn = row[idx.PayloadNumber];
      var key = truck + '|' + y + '|' + mo + '|' + d + '|' + h + '|' + mi + '|' + s + '|' + pn;
      if (store.keys[key]) continue;
      store.keys[key] = 1;
      added++; store.n++;

      var ym = (2000 + Number(y)) + '-' + (mo < 10 ? '0' + mo : '' + mo);
      var pkey = truck + '|' + ym;
      var p = store.partitions[pkey];
      if (!p) {
        p = store.partitions[pkey] = {
          t: truck, m: ym, c: 0, sp: 0,
          z: [0, 0, 0, 0], h: {},
          ct: [0, 0, 0, 0, 0, 0],
          dl: 0, de: 0, ls: 0, es: 0,
          lf: 0, rf: 0, lr: 0, rr: 0, cb: 0, f: 0
        };
      }
      var pt = fp / 10; // тонны
      p.c++; p.sp += pt;
      if (pt < NOMINAL) p.z[0]++;
      else if (pt <= P110) p.z[1]++;
      else if (pt <= P120) p.z[2]++;
      else p.z[3]++;
      var bin = Math.floor(pt);
      p.h[bin] = (p.h[bin] || 0) + 1;

      p.ct[0] += num(row[idx.StopLoadedTime]);
      p.ct[1] += num(row[idx.MovingLoadedTime]);
      p.ct[2] += num(row[idx.MovingEmptyTime]);
      p.ct[3] += num(row[idx.StopEmptyTime]);
      p.ct[4] += num(row[idx.LoadingTime]);
      p.ct[5] += num(row[idx.DumpingTime]);
      p.dl += num(row[idx.DistanceLoaded]);
      p.de += num(row[idx.DistanceEmpty]);
      p.ls += num(row[idx.LdMaxSpeed]) / 1000;
      p.es += num(row[idx.EmMaxSpeed]) / 1000;
      p.lf += num(row[idx.LF_TKPH]) / 10;
      p.rf += num(row[idx.RF_TKPH]) / 10;
      p.lr += num(row[idx.LR_TKPH]) / 10;
      p.rr += num(row[idx.RR_TKPH]) / 10;
      if (num(row[idx.Carryback]) > 0) p.cb++;
      p.f += num(row[idx.FuelLevel]);
    }
    return added;
  }

  function num(v) { return (v == null || v === '' || isNaN(v)) ? 0 : Number(v); }

  function finalize(store) {
    var out = [];
    for (var k in store.partitions) out.push(store.partitions[k]);
    // округление сумм для компактности встроенных данных
    out.forEach(function (p) {
      p.sp = round(p.sp, 1);
      for (var i = 0; i < p.ct.length; i++) p.ct[i] = Math.round(p.ct[i]);
      p.dl = Math.round(p.dl); p.de = Math.round(p.de);
      p.ls = round(p.ls, 1); p.es = round(p.es, 1);
      p.lf = round(p.lf, 1); p.rf = round(p.rf, 1);
      p.lr = round(p.lr, 1); p.rr = round(p.rr, 1); p.f = round(p.f, 1);
    });
    out.sort(function (a, b) {
      if (a.t !== b.t) return Number(a.t) - Number(b.t);
      return a.m < b.m ? -1 : 1;
    });
    return out;
  }

  function round(v, d) { var m = Math.pow(10, d); return Math.round(v * m) / m; }

  return { NOMINAL: NOMINAL, P110: P110, P120: P120, makeStore: makeStore, addRows2D: addRows2D, finalize: finalize };
});
