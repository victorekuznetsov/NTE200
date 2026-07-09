/* Пересобирает embedded_data.json (встроенный набор для дашборда) из ВСЕХ весовых
 * выгрузок: файлов Весовая_*.xlsx в корне репозитория и архива _Архивы/Payload NTE200.7z.
 *
 * Использует ту же логику агрегации (agg.js), что и сам дашборд при загрузке новых
 * файлов в браузере, — поэтому встроенные данные гарантированно совпадают с тем, что
 * пользователь получил бы, загрузив те же файлы вручную. Номинал 180 т, дедупликация
 * циклов по ключу «борт + дата/время + № цикла».
 *
 * Запуск:
 *   1) распакуйте архив плоско в одну папку, напр.:
 *        7z e "../../_Архивы/Payload NTE200.7z" -o/tmp/arc_flat
 *      (или задайте ARCHIVE_DIR через переменную окружения)
 *   2) node build_embedded.js
 * Результат — embedded_data.json рядом со скриптом; его содержимое вставляется в
 * dashboard_template.html вместо плейсхолдера __EMBEDDED_DATA__ при сборке дашборда.
 * Парсер xlsx берётся из npm-пакета xlsx, а при его отсутствии — из приложенного
 * xlsx.mini.min.js.
 */
const path = require('path');
const fs = require('fs');
let XLSX;
try { XLSX = require('xlsx'); } catch (e) { XLSX = require('./xlsx.mini.min.js'); }
const AGG = require('./agg.js');

const REPO_ROOT = path.resolve(__dirname, '..', '..');
const ARCHIVE_DIR = process.env.ARCHIVE_DIR || '/tmp/arc_flat';

const store = AGG.makeStore();

function ingestDir(dir) {
  if (!fs.existsSync(dir)) return { files: 0, added: 0 };
  const files = fs.readdirSync(dir).filter(f => /^Весовая_.*\.xlsx$/i.test(f));
  let added = 0;
  for (const f of files) {
    const wb = XLSX.read(fs.readFileSync(path.join(dir, f)), { type: 'buffer' });
    for (const sn of wb.SheetNames) {
      const rows = XLSX.utils.sheet_to_json(wb.Sheets[sn], { header: 1, blankrows: false, raw: true });
      if (rows.some(r => r && r.indexOf && r.indexOf('TruckId') !== -1)) { added += AGG.addRows2D(store, rows); break; }
    }
  }
  return { files: files.length, added };
}

const root = ingestDir(REPO_ROOT);
console.log('корень:', root.files, 'файлов,', root.added, 'новых циклов');
const arc = ingestDir(ARCHIVE_DIR);
console.log('архив: ', arc.files, 'файлов,', arc.added, 'новых циклов');

const partitions = AGG.finalize(store);
const trucks = new Set(partitions.map(p => p.t)).size;
const cycles = partitions.reduce((s, p) => s + p.c, 0);
console.log('ИТОГО:', trucks, 'бортов,', cycles, 'уникальных циклов,', partitions.length, 'партиций');

fs.writeFileSync(path.join(__dirname, 'embedded_data.json'),
  JSON.stringify({ nominal: AGG.NOMINAL, generated: new Date().toISOString().slice(0, 10), partitions }));
console.log('записан embedded_data.json');
