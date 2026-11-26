# -*- coding: utf-8 -*-
"""
出勤報表自動填打卡時間（依 SKILL.md 規則）

用法：
  python attendance_build.py --template 範本.xlsx --punch 打卡1.xls 打卡2.xls --out 出勤_補上時間.xlsx

  --template  範本（.xlsx；若給 .xls 會用本機 Excel 轉成 .xlsx，需要 pywin32）
  --punch     打卡檔，一份或多份（.xls，兩種格式自動判斷）。
              兩份同一天重疊時，用「較新」的：依檔內最後一天排序，同天則依參數順序（後者較新）。
  --out       輸出 .xlsx（只允許 .xlsx）
  --report    驗證報告 txt（預設與輸出同名 _報告.txt）
  --month     範本月份，如 2026/11（預設由範本自動判斷）

需要：openpyxl、xlrd（讀 .xls，內建處理 Big5/Unicode 編碼）
"""
import argparse, collections, copy, datetime as dt, os, re, statistics, sys
import xlrd
import openpyxl
from openpyxl.styles import Font, PatternFill

# ----------------------------------------------------------------- 參數
DAY_CHANGE = dt.time(7, 0)          # 換天時間 07:00
EARLY_IN = dt.time(5, 0)            # 05:00 前的「上班」算前一天
MAX_GAP_H = 13.0                    # 一班最長（小時）
MIN_GAP_H = 1.0                     # 少於此視為班太短，不配成一班（早退/還時間的短班仍配成一班，2026-11-21 依使用者正確版修正）
DUP_MIN = 10                        # 同人數分鐘內重複打卡算一筆
OT_DROP_H = 3.0                     # 加班打卡緊接在下班後（3 小時內）視為多餘，忽略
LONE_COST = 6.0                     # 配對 DP：落單一筆的成本
LABEL_PEN = 1.5                     # 配對 DP：標籤（上班/下班）與位置不符的罰分
OUT_AFTER_H = 6.0                   # 單筆：比平常上班晚超過此小時數 → 當作下班
DEFAULT_LEN_H = 9.0
NIGHT_IN = dt.time(17, 30)          # 跨夜班：這之後的上班
NIGHT_OUT = dt.time(10, 0)          # 配隔天這之前的打卡
NIGHT_MIN_H = 4.0
NIGHT_BONUS = 8.0
YELLOW = PatternFill("solid", fgColor="FFFFFF00")
DATE_RE = re.compile(r"^\d{4}/\d{2}/\d{2}$")


# ----------------------------------------------------------------- 工具
def code_key(v):
    if v is None or v == "":
        return None
    if isinstance(v, float) and v == int(v):
        v = int(v)
    s = str(v).strip()
    return s.zfill(3) if s.isdigit() else s


def parse_tod(v):
    """'07:01' / Excel 時間小數 → dt.time；空白 → None"""
    if v is None or v == "":
        return None
    if isinstance(v, (int, float)):
        m = int(round((v % 1) * 1440))
        return dt.time(m // 60 % 24, m % 60)
    m = re.match(r"^\s*(\d{1,2}):(\d{2})", str(v))
    if not m:
        return None
    return dt.time(int(m.group(1)) % 24, int(m.group(2)))


def parse_date(v, datemode=0):
    if isinstance(v, (int, float)):
        t = xlrd.xldate_as_tuple(v, datemode)
        return dt.date(*t[:3])
    m = re.match(r"^\s*(\d{4})[/-](\d{1,2})[/-](\d{1,2})", str(v))
    return dt.date(int(m.group(1)), int(m.group(2)), int(m.group(3))) if m else None


def fmt(t):
    return t.strftime("%H:%M")


def ensure_xlsx(path, workdir):
    if path.lower().endswith(".xlsx"):
        return path
    if not path.lower().endswith(".xls"):
        sys.exit("範本必須是 .xls 或 .xlsx")
    import win32com.client
    out = os.path.join(workdir, os.path.splitext(os.path.basename(path))[0] + ".xlsx")
    xl = win32com.client.DispatchEx("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    try:
        wb = xl.Workbooks.Open(os.path.abspath(path))
        wb.SaveAs(os.path.abspath(out), 51)
        wb.Close(False)
    finally:
        xl.Quit()
    return out


# ----------------------------------------------------------------- 讀打卡檔
def read_punch_file(path):
    """回傳 (rows, last_date)。row = dict(code,name,date,t1,t2,o1,o2)"""
    b = xlrd.open_workbook(path)
    s = b.sheet_by_index(0)
    hdr = [str(x).strip() for x in s.row_values(0)]
    rows = []
    if "exp_2" in hdr:                       # 第二種：資料庫匯出
        c = {h: hdr.index(h) for h in ("exp_2", "no", "na", "t1", "t2", "ovrt1", "ovrt2")}
        for r in range(1, s.nrows):
            v = s.row_values(r)
            d = parse_date(v[c["exp_2"]], b.datemode)
            if not d:
                continue
            rows.append(dict(code=code_key(v[c["no"]]), name=str(v[c["na"]]).strip(), date=d,
                             t1=parse_tod(v[c["t1"]]), t2=parse_tod(v[c["t2"]]),
                             o1=parse_tod(v[c["ovrt1"]]), o2=parse_tod(v[c["ovrt2"]])))
    else:                                    # 第一種：出勤日期／代號／姓名／上班／下班／遲到／早退／加班上班
        c = {h: i for i, h in enumerate(hdr)}
        for r in range(1, s.nrows):
            v = s.row_values(r)
            d = parse_date(v[c["出勤日期"]], b.datemode)
            if not d:
                continue
            rows.append(dict(code=code_key(v[c["代號"]]), name=str(v[c["姓名"]]).strip(), date=d,
                             t1=parse_tod(v[c["上班"]]), t2=parse_tod(v[c["下班"]]),
                             o1=parse_tod(v[c["加班上班"]]) if "加班上班" in c else None,
                             o2=parse_tod(v[c["加班下班"]]) if "加班下班" in c else None))
    last = max((r["date"] for r in rows), default=None)
    return rows, last


# ----------------------------------------------------------------- 打卡 → 事件
def row_events(row):
    """一列 → 實際時間的打卡事件 [(datetime, hint, tod)]；hint: in/out"""
    D = row["date"]
    out, prev = [], None
    for key, hint in (("t1", "in"), ("t2", "out"), ("o1", "in"), ("o2", "out")):
        t = row[key]
        if t is None:
            continue
        d = D + dt.timedelta(days=1) if t < DAY_CHANGE else D       # 07:00 前歸隔天
        cand = dt.datetime.combine(d, t)
        while prev is not None and cand < prev:                     # 同列下班早於上班 → 隔天
            cand += dt.timedelta(days=1)
        if key in ("o1", "o2") and prev is not None and (cand - prev).total_seconds() < OT_DROP_H * 3600:
            out.append((cand, "ot"))      # 緊接在下班後的加班打卡：忽略（使用者正確版並非一律取最晚，所以不當候選）
            continue
        out.append((cand, hint))
        prev = cand
    return out


def dedupe(events):
    events.sort(key=lambda e: e[0])
    res = []
    for e in events:
        if res and (e[0] - res[-1][0]).total_seconds() <= DUP_MIN * 60:
            res[-1] = (res[-1][0], res[-1][1], e[0])      # 記下同一群最晚的一筆（下班取最晚）
            continue
        res.append((e[0], e[1], e[0]))
    return res


def pair_events(ev, Lp):
    """DP 配對：回傳 (pairs[(i,j)], lones[i])"""
    n = len(ev)
    INF = 1e18
    best = [INF] * (n + 2)
    best[n] = 0.0
    choice = [None] * (n + 1)
    for i in range(n - 1, -1, -1):
        best[i] = LONE_COST + best[i + 1]
        choice[i] = "lone"
        if i + 1 < n:
            g = (ev[i + 1][0] - ev[i][0]).total_seconds() / 3600
            if MIN_GAP_H <= g <= MAX_GAP_H:
                c = abs(g - Lp)
                if ev[i][1] == "out":
                    c += LABEL_PEN
                if ev[i + 1][1] == "in":
                    c += LABEL_PEN
                # 跨夜班優先：晚上的上班 + 隔天早上的打卡是同一班（不可把早上那筆配給下午那筆）
                if (ev[i][1] == "in" and ev[i][0].time() >= NIGHT_IN and ev[i + 1][0].time() < NIGHT_OUT
                        and ev[i + 1][0].date() > ev[i][0].date() and g >= NIGHT_MIN_H):
                    c -= NIGHT_BONUS
                if c + best[i + 2] < best[i]:
                    best[i] = c + best[i + 2]
                    choice[i] = "pair"
    pairs, lones, i = [], [], 0
    while i < n:
        if choice[i] == "pair":
            pairs.append((i, i + 1))
            i += 2
        else:
            lones.append(i)
            i += 1
    return pairs, lones


def work_date_in(t):
    return (t - dt.timedelta(days=1)).date() if t.time() < EARLY_IN else t.date()


def build_records(rows_by_person, last_date):
    """回傳 records[(code, work_date)] = dict(in,out,notes) 及 person info"""
    records = {}
    info = {}
    for code, rows in rows_by_person.items():
        # 先用「同列上下班都有、且合理」的列估計平常班長
        lens = []
        for r in rows:
            if r["t1"] and r["t2"]:
                evs = row_events(dict(r, o1=None, o2=None))
                if len(evs) == 2:
                    g = (evs[1][0] - evs[0][0]).total_seconds() / 3600
                    if 4 <= g <= 12.5:
                        lens.append(g)
        Lp = statistics.median(lens) if lens else DEFAULT_LEN_H
        events = []
        for r in rows:
            events.extend(row_events(r))
        ots = sorted(e[0] for e in events if e[1] == "ot")
        ev = dedupe([e for e in events if e[1] != "ot"])
        pairs, lones = pair_events(ev, Lp)
        ins = [ev[i][0] for i, j in pairs]
        tins = [t.hour * 60 + t.minute for t in ins]
        info[code] = dict(Lp=Lp, name=rows[0]["name"], n_events=len(ev))
        notes = []

        def rec(wd):
            return records.setdefault((code, wd), dict(inn=None, out=None, notes=[]))

        for i, j in sorted(pairs, key=lambda p: ev[p[0]][0]):
            a, b = ev[i][0], ev[j][0]
            r_ = rec(work_date_in(a))
            if r_["inn"] is None and r_["out"] is None:
                r_["inn"], r_["out"] = a, b
            else:
                r_["notes"].append("同日另有一班，忽略 %s-%s" % (fmt(a), fmt(b)))
        lone_in, lone_out = [], []
        for i in lones:
            t = ev[i][0]
            tod = t.hour * 60 + t.minute
            if tins:
                # 離「平常上班時間」近 → 上班；離「平常上班 + 班長」近 → 下班
                def cd(x):
                    return abs(((tod - x + 720) % 1440) - 720)
                in_dist = min(cd(x) for x in tins)
                out_dist = min(cd(int(x + Lp * 60)) for x in tins)
                is_out = out_dist < in_dist
            else:
                is_out = ev[i][1] == "out"
            (lone_out if is_out else lone_in).append(t)
        for t in lone_in:
            r_ = rec(work_date_in(t))
            if r_["inn"] is None:
                r_["inn"] = t
            else:
                r_["notes"].append("多餘的上班側打卡 %s 忽略" % fmt(t.time()))
        for t in lone_out:
            r_ = rec(work_date_in(t - dt.timedelta(hours=Lp)))
            if r_["out"] is None:
                r_["out"] = t
            else:
                r_["notes"].append("多餘的下班側打卡 %s 忽略" % fmt(t.time()))
    return records, info


# ----------------------------------------------------------------- 範本處理
def sheet_layout(ws):
    hdr = [c.value for c in ws[1]]
    if "日期" not in hdr or "上班" not in hdr:
        return None
    return {h: i for i, h in enumerate(hdr) if h}


def in_merged(ws, cell):
    for rng in ws.merged_cells.ranges:
        if cell.coordinate in rng:
            return True
    return False


def detect_month(wb):
    cnt = collections.Counter()
    for ws in wb:
        for r in ws.iter_rows(min_row=2, max_col=1):
            v = r[0].value
            if isinstance(v, str) and DATE_RE.match(v):
                cnt[v[:7]] += 1
    return cnt.most_common(1)[0][0]


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--template", required=True)
    ap.add_argument("--punch", nargs="+", required=True)
    ap.add_argument("--out", required=True)
    ap.add_argument("--report")
    ap.add_argument("--month")
    ap.add_argument("--from", dest="start", help="只填這天（含）之後，如 2026/11/11；之前的列完全不動")
    a = ap.parse_args()
    START = parse_date(a.start) if a.start else None
    if a.start and not START:
        sys.exit("--from 格式要像 2026/11/11")
    if not a.out.lower().endswith(".xlsx"):
        sys.exit("輸出只允許 .xlsx")
    report_path = a.report or os.path.splitext(a.out)[0] + "_報告.txt"
    rep = []

    def R(s=""):
        rep.append(s)

    # --- 讀打卡檔，較新的優先
    files = []
    for k, p in enumerate(a.punch):
        rows, last = read_punch_file(p)
        files.append((last, k, p, rows))
    files.sort(key=lambda x: (x[0], x[1]))
    by_key = {}
    LAST = max(f[0] for f in files)
    for last, k, p, rows in files:
        tmp = collections.defaultdict(list)
        for r in rows:
            if r["code"]:
                tmp[(r["code"], r["date"])].append(r)
        for key, lst in tmp.items():
            by_key[key] = lst                      # 後（較新）覆蓋前
        R("打卡檔 %s：%d 列，最後一天 %s" % (os.path.basename(p), len(rows), last))
    rows_by_person = collections.defaultdict(list)
    for (code, d), lst in by_key.items():
        rows_by_person[code].extend(lst)
    records, info = build_records(rows_by_person, LAST)

    # --- 範本
    wd_ = os.path.dirname(os.path.abspath(a.out))
    tpl = ensure_xlsx(a.template, wd_)
    wb = openpyxl.load_workbook(tpl)
    month = a.month or detect_month(wb)
    R("範本月份 %s；匯出最後一天 %s" % (month, LAST))

    # 幹部頁：日期還是上個月的區塊 → 改成本月（有代號者）
    fixed = 0
    if "幹部" in wb.sheetnames:
        ws = wb["幹部"]
        L = sheet_layout(ws)
        for r in ws.iter_rows(min_row=2):
            v, c = r[0].value, r[L["代號"]].value
            if isinstance(v, str) and DATE_RE.match(v) and v[:7] != month and code_key(c):
                r[0].value = month + v[7:]
                fixed += 1
    R("幹部頁日期改成 %s 的列數：%d" % (month, fixed))

    COVER_END = dt.datetime.combine(LAST + dt.timedelta(days=1), DAY_CHANGE)
    filled = collections.Counter()
    yellow = []
    skipped_leave = []
    used = set()
    tcodes = set()
    for ws in wb:
        L = sheet_layout(ws)
        if not L or ws.title == "範本":
            continue
        ci, cin, cout, cnm = L["代號"], L["上班"], L["下班"], L["姓名"]
        for row in ws.iter_rows(min_row=2):
            d = row[0].value
            if not (isinstance(d, str) and DATE_RE.match(d)):
                continue
            date = parse_date(d)
            if date > LAST or (START and date < START):
                continue
            code = code_key(row[ci].value)
            if not code:
                continue                      # 沒代號（陳濬豪）不處理
            tcodes.add(code)
            name = row[cnm].value
            cell_in, cell_out = row[cin], row[cout]
            rec = records.get((code, date))
            if rec:
                used.add((code, date))
            # 上班欄是文字（休假/請假…）或合併格 → 整列不動；已有時間時，空白的下班照填
            leave_row = in_merged(ws, cell_in) or (cell_in.value not in (None, "") and parse_tod(cell_in.value) is None)
            blocked_in = in_merged(ws, cell_in) or cell_in.value not in (None, "")
            blocked_out = in_merged(ws, cell_out) or cell_out.value not in (None, "")
            if leave_row and rec and (rec["inn"] or rec["out"]):
                # 上班欄已有文字（休假/請假…）或是合併格，但打卡顯示有出勤：僅供使用者確認
                skipped_leave.append("%s %s %s 上班欄為「%s」，打卡有 %s-%s%s"
                                     % (ws.title, name, d, cell_in.value,
                                        fmt(rec["inn"].time()) if rec["inn"] else "",
                                        fmt(rec["out"].time()) if rec["out"] else "",
                                        "（未動）"))
            # 寫休假但有打卡：上下班不填，打卡時間寫到備註欄（保留原備註）
            if ("備註" in L and str(cell_in.value).strip() == "休假"
                    and rec and (rec["inn"] or rec["out"])):
                if rec["inn"] and rec["out"]:
                    times = "%s-%s" % (fmt(rec["inn"].time()), fmt(rec["out"].time()))
                elif rec["inn"]:
                    times = "上班 %s" % fmt(rec["inn"].time())
                else:
                    times = "下班 %s" % fmt(rec["out"].time())
                note = row[L["備註"]]
                old = "" if note.value is None else str(note.value)
                if "寫休假，但有打卡" not in old and not in_merged(ws, note):
                    note.value = (old + " " if old else "") + times + "（寫休假，但有打卡）"
            for cell, blocked, val, label, reason_missing in (
                    (cell_in, blocked_in, rec["inn"] if rec else None, "上班", None),
                    (cell_out, blocked_out, rec["out"] if rec else None, "下班", None)):
                if blocked:
                    continue
                if leave_row:                 # 上班欄有字（休假/請假…）的列：下班也不填
                    continue
                if val is not None:
                    f = cell.font
                    cell.value = fmt(val.time())
                    cell.font = Font(name="Arial", size=f.sz or 14, bold=False, italic=f.i,
                                     color="FF000000")
                    filled[ws.title] += 1
                    continue
                if blocked_in:                # 休假/請假列：不算缺打卡，不標黃
                    continue
                # 缺：最後一天跨午夜的下班還沒到，不標黃
                if (label == "下班" and rec and rec["inn"] and date == LAST
                        and rec["inn"] + dt.timedelta(hours=info[code]["Lp"]) > COVER_END):
                    continue
                cell.fill = YELLOW
                if not rec:
                    why = "當天沒有任何打卡資料"
                elif rec["notes"]:
                    why = "；".join(rec["notes"])
                elif label == "下班":
                    why = "只有上班打卡"
                else:
                    why = "只有下班打卡"
                yellow.append((ws.title, name, d, label, why))
    # 一列只有一格有時間、另一格空白：水平置中（避免範本的跨欄置中讓時間跑到兩欄中間）
    for ws in wb:
        L = sheet_layout(ws)
        if not L or ws.title == "範本":
            continue
        for row in ws.iter_rows(min_row=2):
            d = row[0].value
            if START and not (isinstance(d, str) and DATE_RE.match(d) and parse_date(d) >= START):
                continue
            pair = (row[L["上班"]], row[L["下班"]])
            for cell, other in (pair, pair[::-1]):
                if (cell.value not in (None, "") and parse_tod(cell.value if isinstance(cell.value, str) else "00:00")
                        is not None and other.value in (None, "")):
                    al = copy.copy(cell.alignment)
                    al.horizontal = "center"
                    cell.alignment = al
    wb.save(a.out)

    total = sum(filled.values())
    R()
    R("填入格數：%d（%s）" % (total, "、".join("%s %d" % kv for kv in filled.items())))
    R("黃色格數：%d" % len(yellow))
    R()
    R("== 黃色格明細 ==")
    for y in yellow:
        R("  %s｜%s｜%s｜%s｜%s" % y)
    R()
    R("== 上班欄已有文字（休假/請假…），但打卡資料顯示有上班（未動，請自行確認） ==")
    for s in skipped_leave:
        R("  " + s)
    R()
    R("== 有打卡但範本沒有的代號（未處理） ==")
    miss = collections.defaultdict(int)
    for (code, d) in by_key:
        if code not in tcodes:
            miss[(code, by_key[(code, d)][0]["name"])] += 1
    for (c, n), k in sorted(miss.items()):
        R("  %s %s（%d 天）" % (c, n, k))
    open(report_path, "w", encoding="utf-8").write("\n".join(rep))
    print("\n".join(rep[:6]))
    print("輸出：", a.out)
    print("報告：", report_path)


if __name__ == "__main__":
    main()
