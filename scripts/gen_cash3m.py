#!/usr/bin/env python3
"""Regenerate the embedded CASH3M dataset (3-month cash rates) in index.html.

Usage:
    python3 scripts/gen_cash3m.py           # dry-run: print last rows and gaps
    python3 scripts/gen_cash3m.py --write   # write into index.html

Sources (monthly averages, % p.a.):
    usd: FRED IR3TIB01USM156N — OECD 3-month interbank rate, United States
         (USD LIBOR 3M historically; LIBOR's successor benchmark after 2023)
    eur: ECB FM.M.U2.EUR.RT.MM.EURIBOR3MD_.HSTA — Euribor 3-month

Isolated missing months are linearly interpolated (flagged in the output); months
after a series' last observation are carried forward from the last value.
"""
import csv, io, json, os, sys, urllib.request, calendar

REPO = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
HTML = os.path.join(REPO, 'index.html')
PREFIX = 'const CASH3M = '
START = '2003-01'

FRED = 'https://fred.stlouisfed.org/graph/fredgraph.csv?id=IR3TIB01USM156N'
ECB = ('https://data-api.ecb.europa.eu/service/data/FM/'
       'M.U2.EUR.RT.MM.EURIBOR3MD_.HSTA?format=csvdata&startPeriod=2002-01')


def fetch(url):
    req = urllib.request.Request(url, headers={'User-Agent': 'curl/8.0'})
    return urllib.request.urlopen(req, timeout=60).read().decode()


def fred_usd():
    out = {}
    for row in csv.DictReader(io.StringIO(fetch(FRED))):
        v = row['IR3TIB01USM156N'].strip()
        out[row['observation_date'][:7]] = float(v) if v not in ('', '.') else None
    return out


def ecb_eur():
    return {row['TIME_PERIOD']: float(row['OBS_VALUE'])
            for row in csv.DictReader(io.StringIO(fetch(ECB))) if row['OBS_VALUE']}


def months(start, end):
    y, m = map(int, start.split('-'))
    ey, em = map(int, end.split('-'))
    while (y, m) <= (ey, em):
        yield f'{y:04d}-{m:02d}'
        y, m = (y + 1, 1) if m == 12 else (y, m + 1)


def fill(series, keys, name):
    vals = [series.get(k) for k in keys]
    last = max(i for i, v in enumerate(vals) if v is not None)
    for i, v in enumerate(vals):
        if v is not None:
            continue
        if i > last:
            vals[i] = vals[last]
            print(f'{name} {keys[i]}: carried forward {vals[last]}')
        else:
            j = next(j for j in range(i + 1, len(vals)) if vals[j] is not None)
            p = vals[i - 1]
            vals[i] = round(p + (vals[j] - p) / (j - i + 1), 4)
            print(f'{name} {keys[i]}: interpolated {vals[i]}')
    return vals


def main():
    lines = open(HTML).readlines()
    raw = next(l for l in lines if l.startswith('let RAW = '))
    end = json.loads(raw[len('let RAW = '):].rstrip().rstrip(';'))['nav'][-1]['d'][:7]
    keys = list(months(START, end))
    usd = fill(fred_usd(), keys, 'usd')
    eur = fill(ecb_eur(), keys, 'eur')
    rows = []
    for k, u, e in zip(keys, usd, eur):
        y, m = map(int, k.split('-'))
        rows.append({'d': f'{k}-{calendar.monthrange(y, m)[1]:02d}',
                     'usd': round(u, 4), 'eur': round(e, 4)})
    print(len(rows), 'rows;', rows[0], '...', rows[-1])
    if '--write' in sys.argv:
        idx = next(i for i, l in enumerate(lines) if l.startswith(PREFIX))
        lines[idx] = PREFIX + json.dumps(rows, separators=(',', ':')) + ';\n'
        open(HTML, 'w').writelines(lines)
        print('written to', HTML)


if __name__ == '__main__':
    main()
