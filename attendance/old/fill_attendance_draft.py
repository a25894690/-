#!/usr/bin/env python3
"""Fill 上班/下班 punch times into the 經貿店 attendance workbook.

usage: fill_attendance.py TEMPLATE.xls RAW.xls START END OUT.xlsx [--overwrite]
  START/END like 2026/11/11. Only blank 上班 cells are filled (休假/年假/病假...
  and already-filled cells are left alone) unless --overwrite is given.
Rows that need a human look are highlighted yellow and listed on stdout.
"""
import sys, collections, datetime as dt
import xlrd, openpyxl
from openpyxl.styles import PatternFill

GAP_MIN = 14 * 60          # punches closer than this belong to one shift
YELLOW = PatternFill('solid', fgColor='FFFF00')


def tm(v):
    if isinstance(v, float):
        m = round(v * 1440) % 1440
        return '%02d:%02d' % (m // 60, m % 60)
    return v


def hm(s):
    return int(s[:2]) * 60 + int(s[3:5])


def load_punches(raw_path):
    """{員工代號: set(datetime)}. Raw rows are per-day; t2/ovrt cross midnight."""
    sh = xlrd.open_workbook(raw_path).sheet_by_index(0)
    P = collections.defaultdict(set)
    for r in range(1, sh.nrows):
        v = sh.row_values(r)
        D = dt.datetime.strptime(v[1], '%Y/%m/%d')
        last = None
        for t in (tm(v[6]), tm(v[7]), tm(v[10]), tm(v[11])):   # t1 t2 ovrt1 ovrt2
            if not t:
                continue
            a = D + dt.timedelta(minutes=hm(t))
            if last and a < last:
                a += dt.timedelta(days=1)
            P[v[4]].add(a)
            last = a
    return P


def shifts(punches):
    """Group sorted punches into shifts -> {YYYY/MM/DD of first punch: (in, out, n)}"""
    g = []
    for p in sorted(punches):
        if g and (p - g[-1][-1]).total_seconds() / 60 < GAP_MIN:
            g[-1].append(p)
        else:
            g.append([p])
    return {x[0].strftime('%Y/%m/%d'):
            (x[0].strftime('%H:%M'), x[-1].strftime('%H:%M') if len(x) > 1 else '', len(x))
            for x in g}


def main():
    a = [x for x in sys.argv[1:] if not x.startswith('--')]
    overwrite = '--overwrite' in sys.argv
    tpl, raw, start, end, out = a
    P = {k: shifts(v) for k, v in load_punches(raw).items()}
    src = xlrd.open_workbook(tpl)
    wb = openpyxl.Workbook(); wb.remove(wb.active)
    review = []
    for s in src.sheets():
        ws = wb.create_sheet(s.name)
        h = s.row_values(0)
        iu = h.index('上班') if '上班' in h else None
        for r in range(s.nrows):
            row = [tm(c) for c in s.row_values(r)]
            flag = None
            d = row[0]
            if r and iu is not None and isinstance(d, str) and start <= d <= end:
                no = str(row[1]).split('.')[0].zfill(3)
                blank = row[iu] == ''
                if blank or overwrite and (':' in str(row[iu]) or blank):
                    sh = P.get(no, {}).get(d)
                    if sh:
                        row[iu], row[iu + 1] = sh[0], sh[1]
                        span = (hm(sh[1]) - hm(sh[0])) % 1440 if sh[1] else None
                        if not sh[1]:
                            flag = '只有一筆打卡'
                        elif span < 180 or span > 14 * 60:
                            flag = '班時異常 %.1fh' % (span / 60)
                    elif blank:
                        pass  # no punch: leave blank (未打卡/未排)
                if flag:
                    review.append((s.name, row[2], d, row[iu], row[iu + 1], flag))
            ws.append(row)
            if flag:
                for c in ws[r + 1]:
                    c.fill = YELLOW
    wb.save(out)
    print('saved', out)
    for x in review:
        print('REVIEW', *x)


main()
