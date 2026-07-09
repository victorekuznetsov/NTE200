# Пересобирает данные для "Дашборд весового контроля NTE200.html" из файлов Весовая_*.xlsx.
# Использование: python3 build_weighing_dashboard_data.py, затем вставить содержимое
# weighing_data.json в HTML-дашборд как значение константы DATA.
import openpyxl, glob, re, json, os, statistics
from collections import defaultdict

RATED_T = 178.0  # nominal payload, derived from FinalPayload/LoadPercentage across all trucks (median 1780 -> /10 = 178.0 t)

REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
files = sorted(glob.glob(os.path.join(REPO_ROOT, 'Весовая_*.xlsx')))

MONTH_NAMES = {12:'дек', 1:'янв', 2:'фев', 3:'мар', 4:'апр', 5:'май', 6:'июн', 7:'июл'}

trucks = {}
fleet_load_hist = defaultdict(int)
fleet_month_cycles = defaultdict(lambda: defaultdict(int))
fleet_month_loadsum = defaultdict(lambda: defaultdict(float))

BIN_WIDTH = 10  # tonnes
def bin_key(t):
    return int(t // BIN_WIDTH) * BIN_WIDTH

for f in files:
    m = re.match(r'.*Весовая_(\d+)_', f)
    truck = m.group(1)
    wb = openpyxl.load_workbook(f, data_only=True, read_only=True)
    sn = [s for s in wb.sheetnames if 'Payload' in s]
    if not sn:
        continue
    ws = wb[sn[0]]
    rows = ws.iter_rows(min_row=2, values_only=True)
    header = next(rows)
    idx = {h: i for i, h in enumerate(header)}

    n = 0
    payloads = []
    loadpcts = []
    stop_loaded = []
    moving_loaded = []
    moving_empty = []
    stop_empty = []
    loading_t = []
    dumping_t = []
    dist_empty = []
    dist_loaded = []
    ld_speed = []
    em_speed = []
    lr_tkph = []
    rr_tkph = []
    fuel = []
    hist = defaultdict(int)
    months = set()
    days_by_month = defaultdict(set)
    first_dt = None
    last_dt = None

    for r in rows:
        fp = r[idx['FinalPayload']]
        lp = r[idx['LoadPercentage']]
        if not fp or not lp:
            continue
        n += 1
        t = fp / 10.0
        payloads.append(t)
        loadpcts.append(lp)
        hist[bin_key(t)] += 1
        fleet_load_hist[bin_key(t)] += 1

        yr, mo, d, h, mi, s = r[idx['ear']], r[idx['Month']], r[idx['Day']], r[idx['Hour']], r[idx['Minute']], r[idx['Second']]
        full_year = 2000 + yr
        months.add((full_year, mo))
        days_by_month[mo].add(d)
        dt = (full_year, mo, d, h, mi, s)
        if first_dt is None or dt < first_dt:
            first_dt = dt
        if last_dt is None or dt > last_dt:
            last_dt = dt
        fleet_month_cycles[(full_year, mo)][truck] += 1
        fleet_month_loadsum[(full_year, mo)][truck] += lp

        sl = r[idx['StopLoadedTime']]; ml = r[idx['MovingLoadedTime']]
        me = r[idx['MovingEmptyTime']]; se = r[idx['StopEmptyTime']]
        lt = r[idx['LoadingTime']]; dt_ = r[idx['DumpingTime']]
        if sl is not None: stop_loaded.append(sl)
        if ml is not None: moving_loaded.append(ml)
        if me is not None: moving_empty.append(me)
        if se is not None: stop_empty.append(se)
        if lt is not None: loading_t.append(lt)
        if dt_ is not None: dumping_t.append(dt_)

        de = r[idx['DistanceEmpty']]; dl = r[idx['DistanceLoaded']]
        if de: dist_empty.append(de)
        if dl: dist_loaded.append(dl)

        lds = r[idx['LdMaxSpeed']]; ems = r[idx['EmMaxSpeed']]
        if lds: ld_speed.append(lds/1000.0)
        if ems: em_speed.append(ems/1000.0)

        lr = r[idx['LR_TKPH']]; rr = r[idx['RR_TKPH']]
        if lr is not None: lr_tkph.append(lr/10.0)
        if rr is not None: rr_tkph.append(rr/10.0)

        fl = r[idx['FuelLevel']]
        if fl is not None: fuel.append(fl)

    wb.close()
    if n == 0:
        continue

    avg = lambda L: (sum(L)/len(L)) if L else 0
    overload = sum(1 for lp in loadpcts if lp > 110)
    target = sum(1 for lp in loadpcts if 90 <= lp <= 110)
    underload = sum(1 for lp in loadpcts if lp < 90)

    cyc_stop_loaded = avg(stop_loaded)/60.0
    cyc_moving_loaded = avg(moving_loaded)/60.0
    cyc_moving_empty = avg(moving_empty)/60.0
    cyc_stop_empty = avg(stop_empty)/60.0
    cyc_loading = avg(loading_t)/60.0
    cyc_dumping = avg(dumping_t)/60.0
    total_cycle_min = cyc_stop_loaded+cyc_moving_loaded+cyc_moving_empty+cyc_stop_empty+cyc_loading+cyc_dumping

    total_tonnage = sum(payloads)
    productivity_tph = (avg(payloads) / (total_cycle_min/60.0)) if total_cycle_min else 0

    trucks[truck] = {
        'truck': truck,
        'cycles': n,
        'avg_payload_t': round(avg(payloads), 1),
        'avg_load_pct': round(avg(loadpcts), 1),
        'total_tonnage_t': round(total_tonnage, 0),
        'overload_pct': round(100*overload/n, 1),
        'target_pct': round(100*target/n, 1),
        'underload_pct': round(100*underload/n, 1),
        'hist': dict(sorted(hist.items())),
        'cycle_min': {
            'stop_loaded': round(cyc_stop_loaded, 2),
            'moving_loaded': round(cyc_moving_loaded, 2),
            'moving_empty': round(cyc_moving_empty, 2),
            'stop_empty': round(cyc_stop_empty, 2),
            'loading': round(cyc_loading, 2),
            'dumping': round(cyc_dumping, 2),
            'total': round(total_cycle_min, 2),
        },
        'dist_loaded_km': round(avg(dist_loaded)/1000.0, 2),
        'dist_empty_km': round(avg(dist_empty)/1000.0, 2),
        'ld_speed_kmh': round(avg(ld_speed), 1),
        'em_speed_kmh': round(avg(em_speed), 1),
        'lr_tkph': round(avg(lr_tkph), 1),
        'rr_tkph': round(avg(rr_tkph), 1),
        'fuel_pct': round(avg(fuel), 1),
        'productivity_tph': round(productivity_tph, 1),
        'months': sorted(f"{MONTH_NAMES.get(mo,mo)} {y}" for y, mo in months),
    }
    y_first, mo_first, d_first = first_dt[0], first_dt[1], first_dt[2]
    y_last, mo_last, d_last = last_dt[0], last_dt[1], last_dt[2]
    trucks[truck]['period_str'] = f"{d_first:02d}.{mo_first:02d}.{y_first} – {d_last:02d}.{mo_last:02d}.{y_last}"

fleet = {
    'rated_t': RATED_T,
    'trucks_n': len(trucks),
    'cycles_total': sum(t['cycles'] for t in trucks.values()),
    'tonnage_total': round(sum(t['total_tonnage_t'] for t in trucks.values()), 0),
    'load_hist': dict(sorted(fleet_load_hist.items())),
    'month_trend': [
        {
            'label': f"{MONTH_NAMES.get(mo, mo)} {y}",
            'cycles': sum(fleet_month_cycles[(y,mo)].values()),
            'avg_load_pct': round(sum(fleet_month_loadsum[(y,mo)].values()) / sum(fleet_month_cycles[(y,mo)].values()), 1) if sum(fleet_month_cycles[(y,mo)].values()) else 0
        } for (y, mo) in sorted(fleet_month_cycles.keys())
    ]
}

out = {'fleet': fleet, 'trucks': trucks}
out_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'weighing_data.json')
with open(out_path, 'w', encoding='utf-8') as fh:
    json.dump(out, fh, ensure_ascii=False, indent=1)

print('trucks:', len(trucks))
print('fleet cycles', fleet['cycles_total'], 'tonnage', fleet['tonnage_total'])
print('month trend', fleet['month_trend'])
