"""Один месяц честного теста на открытых часах АТС; без подбора по тесту."""
import csv
import hashlib
import json
import re
from collections import Counter
from datetime import datetime, date
from pathlib import Path
from time import perf_counter

BASE = Path(__file__).resolve().parent
import xlrd
import openpyxl


def read_rows(path):
    if path.suffix == '.xls':
        book = xlrd.open_workbook(str(path))
        sheet = book.sheet_by_index(0)
        raw = [sheet.row_values(i) for i in range(sheet.nrows)]
    else:
        book = openpyxl.load_workbook(path, data_only=True, read_only=True)
        raw = list(book.active.values)
        book.close()
    assert any('Кемеровская область' in str(row) for row in raw)
    result = []
    for row in raw:
        if isinstance(row[0], str) and re.fullmatch(r'\d{2}\.\d{2}\.\d{4}', row[0]):
            d = datetime.strptime(row[0], '%d.%m.%Y').date()
            h = int(row[1])
            assert 1 <= h <= 24 and h == row[1]
            result.append((d,h))
    assert len({d for d,h in result}) == len(result)
    return result


def main():
    started = perf_counter()
    train_path = BASE/'istochniki/202509_KUZBSBYT_32_calcfacthour.xls'
    test_path = BASE/'istochniki/202609_KUZBSBYT_32_calcfacthour.xlsx'
    train, test = read_rows(train_path), read_rows(test_path)
    assert all(d.year == 2025 and d.month == 9 for d,h in train)
    assert all(d.year == 2026 and d.month == 9 for d,h in test)
    # Консервативная граница доступности для сентября 2026 — конец июля.
    cutoff = date(2026,7,31)
    assert max(d for d,h in train) <= cutoff < min(d for d,h in test)
    counts = Counter(h for d,h in train)
    ranked = sorted(counts, key=lambda h: (-counts[h],h))
    assert len(ranked) >= 2
    source_rows = [dict(split=split,date=d.isoformat(),hour_ats=h)
                   for split,rows in [('train',train),('test',test)] for d,h in rows]
    daily, summaries = [], []
    for width in (1,2):
        selected = ranked[:width]
        # Одинаковый бюджет энергии разряда 1 МВт·ч в день.
        # Накопитель 1 МВт / 2 МВт·ч, нагрузка 3 МВт — учебные параметры.
        power = 1.0 / width
        hits = 0
        for d,h in test:
            hit = h in selected
            reduction = power if hit else 0.0
            assert 0 <= reduction <= 1 and width*power == 1
            hits += hit
            daily.append(dict(date=d.isoformat(),hour_ats=h,slots=width,
                              predicted_hours=';'.join(map(str,selected)),
                              discharge_mw_each=power,discharge_mwh_total=1,
                              hit=int(hit),peak_reduction_mw=reduction,
                              load_at_peak_mw=3-reduction))
        average = sum(r['peak_reduction_mw'] for r in daily if r['slots']==width)/len(test)
        # Независимая арифметическая сверка с числом попаданий.
        assert abs(average-hits/len(test)/width) < 1e-12
        summaries.append(dict(slots=width,hours=selected,test_days=len(test),hits=hits,
                              hit_rate=hits/len(test),mean_reduction_mw=average,
                              ideal_reduction_mw=1.0,share_of_ideal=average))
    for name,rows in [('chasy-pika.csv',source_rows),('daily.csv',daily),('summary.csv',summaries)]:
        with (BASE/name).open('w',encoding='utf-8-sig',newline='') as f:
            w=csv.DictWriter(f,fieldnames=list(rows[0]));w.writeheader();w.writerows(rows)
    manifest = dict(downloaded='2026-11-09',region='Кемеровская область',supplier='ПАО Кузбассэнергосбыт',
                    source='https://www.atsenergo.ru/results/market/calcfacthour',
                    files={p.name:dict(url='https://www.atsenergo.ru/dload/calcfacthour_regions/'+p.name,
                                       sha256=hashlib.sha256(p.read_bytes()).hexdigest()) for p in [train_path,test_path]},
                    train_max_date=max(d for d,h in train).isoformat(),cutoff=cutoff.isoformat(),
                    train_frequencies=dict(sorted(counts.items())),
                    test_frequencies=dict(sorted(Counter(h for d,h in test).items())),
                    assumptions=dict(load_mw=3,battery_mw=1,battery_mwh=2,daily_discharge_mwh=1,
                                     usable_energy='AC-side budget; charging/losses/cost not modelled',
                                     time='АТС hour labels retained without timezone conversion'))
    (BASE/'manifest.json').write_text(json.dumps(manifest,ensure_ascii=False,indent=2)+'\n',encoding='utf-8')
    print(json.dumps({'train_days':len(train),'results':summaries,'elapsed_s':perf_counter()-started},ensure_ascii=False))


if __name__ == '__main__':
    main()
