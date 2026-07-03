#!/usr/bin/env python3
"""Regenerate the embedded RAW dataset in index.html from the BG Fund workbook.

Usage:
    python3 scripts/gen_raw.py "/path/to/BG Fund - Historical data ....xlsx"          # dry-run: compare only
    python3 scripts/gen_raw.py "/path/to/BG Fund - Historical data ....xlsx" --write  # write into index.html

The dry-run prints, per section: row counts, any newly added dates (expected on a
monthly update), and any REVISED cells (a prior month's value changed in the new
workbook). It refuses to --write if a date or a key present in the old data would be
LOST, which would indicate a column mis-mapping rather than a genuine update.

Column mappings below are 0-indexed against each sheet's rows. If the fund provider
changes the workbook layout, re-check these against the sheet headers.
"""
import json, datetime, sys, os

import openpyxl

REPO = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
HTML = os.path.join(REPO, 'index.html')
RAW_PREFIX = 'let RAW = '


def num(c):
    if c is None or c == '' or c == '-':
        return None
    if isinstance(c, str):
        try:
            c = float(c)
        except ValueError:
            return None
    if isinstance(c, (int, float)):
        return round(float(c), 8)
    return None


def dstr(c):
    return c.strftime('%Y-%m-%d')


def build(wb, name, datecol, mapping):
    out = []
    for r in wb[name].iter_rows(values_only=True):
        if len(r) <= datecol or not isinstance(r[datecol], datetime.datetime):
            continue
        row = {'d': dstr(r[datecol])}
        for key, col in mapping:
            v = num(r[col]) if col < len(r) else None
            if v is not None:
                row[key] = v
        if len(row) > 1:
            out.append(row)
    return out


def generate(xlsx):
    wb = openpyxl.load_workbook(xlsx, read_only=True, data_only=True)
    return {
        'nav': build(wb, 'Monthly NAV', 0, [
            ('un', 1), ('um', 2), ('uy', 3), ('ub', 4),
            ('en', 7), ('em', 8), ('ey', 9), ('eb', 10),
            ('gn', 13), ('gm', 14), ('gy', 15),
            ('cn', 18), ('cm', 19), ('cy', 20),
            ('a', 31), ('v', 33)]),
        'perf': build(wb, 'Performance attribution', 1, [
            ('yr', 0), ('vs', 2), ('mc', 3), ('cb', 4), ('vt', 5), ('wa', 6),
            ('es', 7), ('ra', 8), ('ls', 9),
            ('cs', 10), ('cl', 11), ('ca', 12), ('cx', 13),
            ('ts', 14), ('t', 15)]),
        'exp': build(wb, 'Exposure (%)', 0, [
            ('cb', 1), ('cb_l', 2), ('cb_s', 3),
            ('vt', 5), ('vt_l', 6), ('vt_s', 7),
            ('wa', 8), ('wa_l', 9), ('wa_s', 10),
            ('eq', 11), ('eq_l', 12), ('eq_s', 13),
            ('ra', 14), ('ra_l', 15), ('ra_s', 16),
            ('ls', 17), ('ls_l', 18), ('ls_s', 19),
            ('cl', 20), ('cl_l', 21), ('cl_s', 22),
            ('cs', 23), ('cs_l', 24), ('cs_s', 25),
            ('cx', 26), ('cx_l', 27), ('cx_s', 28),
            ('tr', 29), ('tr_l', 30), ('tr_s', 31),
            ('qe', 32), ('qe_l', 33), ('qe_s', 34),
            ('sy', 35), ('sy_l', 36), ('sy_s', 37),
            ('ot', 38), ('ot_l', 39), ('ot_s', 40),
            ('g', 42), ('n', 43), ('l', 44), ('s', 45), ('am', 47)]),
        'ear': build(wb, 'Equity at risk', 0, [
            ('te', 16), ('vt', 18), ('et', 19), ('ct', 20), ('tt', 21)]),
        'greeks': build(wb, 'Greeks', 0, [
            ('de', 1), ('ga', 2), ('ve', 3), ('evw', 4),
            ('th', 6), ('oth', 7), ('ir', 8), ('irp', 9), ('cr', 10)]),
        'lev': build(wb, 'Leverage', 0, [
            ('cmx', 1), ('cc', 2), ('gmx', 3), ('gc', 4)]),
        'cash': build(wb, 'Cash', 0, [('u', 1), ('t', 2)]),
        'ifrs': build(wb, 'IFRS13', 0, [
            ('a1', 1), ('a2', 2), ('a3', 3), ('ce', 4), ('ta', 5),
            ('l1', 7), ('l2', 8), ('l3', 9), ('tl', 10), ('nt', 11)]),
    }


def load_existing(lines):
    for i, ln in enumerate(lines):
        if ln.startswith(RAW_PREFIX):
            return i, json.loads(ln[len(RAW_PREFIX):].rstrip().rstrip(';'))
    raise SystemExit('could not find "let RAW = ..." line in index.html')


def compare(old, new):
    """Return number of blocking problems (lost dates/keys). Prints all diffs."""
    problems = 0
    for sec in old:
        o = {r['d']: r for r in old[sec]}
        n = {r['d']: r for r in new[sec]}
        missing = sorted(set(o) - set(n))
        added = sorted(set(n) - set(o))
        if missing:
            print(f'{sec}: BLOCKING dates in old but not new: {missing}')
            problems += 1
        print(f'{sec}: old {len(o)} rows, new {len(n)} rows, added dates: {added}')
        for d in sorted(set(o) & set(n)):
            okeys, nkeys = set(o[d]), set(n[d])
            for k in okeys - nkeys:
                print(f'  BLOCKING {sec} {d}: key {k} lost (old={o[d][k]})')
                problems += 1
            for k in nkeys - okeys:
                print(f'  {sec} {d}: key {k} gained (new={n[d][k]})')
            for k in okeys & nkeys:
                ov, nv = o[d][k], n[d][k]
                if isinstance(ov, float) or isinstance(nv, float):
                    if abs(float(ov) - float(nv)) > 1e-9:
                        print(f'  REVISED {sec} {d}: {k} old={ov} new={nv}')
                elif ov != nv:
                    print(f'  REVISED {sec} {d}: {k} old={ov} new={nv}')
    return problems


def main():
    args = [a for a in sys.argv[1:] if not a.startswith('--')]
    write = '--write' in sys.argv
    if not args:
        raise SystemExit(__doc__)
    xlsx = args[0]

    new = generate(xlsx)
    lines = open(HTML).readlines()
    idx, old = load_existing(lines)
    problems = compare(old, new)
    print('PROBLEMS:', problems)

    if write:
        if problems:
            raise SystemExit('refusing to write with blocking problems (see above)')
        lines[idx] = RAW_PREFIX + json.dumps(new, separators=(',', ':')) + ';\n'
        open(HTML, 'w').writelines(lines)
        print('written to', HTML)


if __name__ == '__main__':
    main()
