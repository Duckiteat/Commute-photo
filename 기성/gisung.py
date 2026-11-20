#!/usr/bin/env python3
"""기성 집계 — 출입기록 대신 '사진으로 받은 일일 기록'으로 하도급 기성 공수·금액을 계산한다.

원래 엑셀(기성 xlsb)의 흐름
    출입기록(게이트 타각) → 분류 테이블(출퇴근 시간 → 조출/주간/연장/야간/철야)
    → 일일·근무자별 근무표 → 월간 근무표(직종별) → 하도급 기성 내역서 → 검사원 → 신청서
을 아래처럼 바꾼다.
    사진(출력일보 서명지 · 안전모 19:00 사진) → 기록/YYYY-MM-DD.csv (Claude 가 읽어 만듦)
    → 이 스크립트 → 일일 · 근무자별 · 월간 집계 · 기성 내역 · 점검 (xlsx 한 파일)

명령
  설정    <기성.xlsb|xlsx> [-o 설정.json] [--다음회차]
          지난 기성 파일에서 계약 공수·단가·전회 누계·인원(식별표)·공휴일·가중치를 뽑는다.
          --다음회차: 이 파일의 '금회 누계'를 다음 회차의 '전회 누계'로 (회차 +1, 기간은 직접 고침).
  계산    <설정.json> <기록 폴더> [-o 결과.xlsx]
  출입기록 <기성.xlsb|xlsx> <출력 폴더>
          (게이트 기록이 있을 때만) 출입기록 시트를 일일 기록 CSV 로 바꾼다 — 검증 · 대조용.
"""
import argparse, csv, datetime as dt, json, re, sys
from collections import defaultdict, OrderedDict
from pathlib import Path

KINDS = ["조출", "주간", "연장", "야간", "철야"]
# 구분 → 칸별 '횟수' (가중치를 곱해 공수가 됨). 원래 엑셀 식별표의 코드표와 같은 결과가 나오게 맞춤.
#   주간 135 = 1 · 연장 136 = 1.5 · 야간 137 = 2 · 조출+주간 165 = 1.5 · 조출+연장 166 = 2 · 조출+야간 167 = 2.5
#   오전만 134 · 오후만 145 = 0.5 · 철야 = 2 (주말·공휴일은 가중치로 1.5배)
GUBUN = {
    "주간": {"주간": 1},
    "연장": {"주간": 1, "연장": 1},
    "야간": {"주간": 1, "야간": 1},
    "조출": {"조출": 1, "주간": 1},
    "조출+연장": {"조출": 1, "주간": 1, "연장": 1},
    "조출+야간": {"조출": 1, "주간": 1, "야간": 1},
    "철야": {"철야": 1},
    "오전": {"주간": 0.5},
    "오후": {"주간": 0.5},
}
ZERO = {"연차", "결근", "휴무", "미출근"}


# ───────────────────────── 공통 ─────────────────────────
def read_book(path):
    """xlsb 는 pyxlsb(저장된 계산값), xlsx 는 openpyxl(data_only)로 {시트: {(행,열): 값}} — 1부터 셈."""
    p = Path(path)
    out = {}
    if p.suffix.lower() == ".xlsb":
        from pyxlsb import open_workbook
        with open_workbook(str(p)) as wb:
            for name in wb.sheets:
                d = {}
                with wb.get_sheet(name) as sh:
                    for i, row in enumerate(sh.rows()):
                        for c in row:
                            if c.v not in (None, ""):
                                d[(i + 1, c.c + 1)] = c.v
                out[name] = d
    else:
        import openpyxl, warnings
        warnings.filterwarnings("ignore")
        wb = openpyxl.load_workbook(str(p), data_only=True, read_only=True)
        for ws in wb.worksheets:
            d = {}
            for row in ws.iter_rows():
                for c in row:
                    if c.value not in (None, ""):
                        d[(c.row, c.column)] = c.value
            out[ws.title] = d
    return out


def to_date(v):
    if isinstance(v, dt.datetime):
        return v.date()
    if isinstance(v, dt.date):
        return v
    if isinstance(v, (int, float)):
        return (dt.datetime(1899, 12, 30) + dt.timedelta(days=float(v))).date()
    return dt.date.fromisoformat(str(v)[:10])


def find_sheet(book, key):
    for name in book:
        if key in name.replace(" ", ""):
            return book[name]
    raise SystemExit(f"'{key}' 시트를 찾지 못함 — 시트 이름: {list(book)}")


def norm_job(s):
    s = str(s).strip()
    s = re.sub(r"Engineer\s*\(?\s*(\S+?)\s*\)?$", r"Eng. \1", s)
    s = re.sub(r"Leader\s*\(\s*(\S+?)\s*\)$", r"Leader \1", s)
    return re.sub(r"\s+", " ", s)


# ───────────────────────── 설정 뽑기 ─────────────────────────
def cmd_config(a):
    book = read_book(a.book)
    nae = find_sheet(book, "기성내역서")
    mon = find_sheet(book, "월간근무표")
    sik = find_sheet(book, "식별표")
    sin = find_sheet(book, "기성신청서")

    text = " ".join(str(v) for v in nae.values() if isinstance(v, str))
    m = re.search(r"(\d{4})\.(\d{1,2})\.(\d{1,2})\s*~\s*(\d{4})\.(\d{1,2})\.(\d{1,2})", text)
    period = [f"{m[1]}-{int(m[2]):02d}-{int(m[3]):02d}", f"{m[4]}-{int(m[5]):02d}-{int(m[6]):02d}"] if m else ["", ""]
    m = re.search(r"제\s*\(\s*(\d+)\s*\)\s*회", " ".join(str(v) for v in sin.values() if isinstance(v, str)))
    round_no = int(m[1]) if m else None

    # 가중치: 월간 근무표 '조출 주간 연장 야간 철야' 아래 줄
    weights = {"평일": {}, "주말": {}}
    for (r, c), v in mon.items():
        if v == "조출" and mon.get((r, c + 4)) == "철야" and isinstance(mon.get((r + 1, c)), (int, float)):
            key = "평일" if not weights["평일"] else "주말"
            if weights[key]:
                continue
            weights[key] = {k: mon[(r + 1, c + i)] for i, k in enumerate(KINDS)}
            if len(weights["평일"]) and len(weights["주말"]):
                break

    # 직종: 내역서 '직종' 열 아래 (계약 공수 · 단가 · 전회 · 금회 누계)
    hdr = next((r, c) for (r, c), v in nae.items() if v == "직종")
    jobs = []
    r = hdr[0] + 1
    while (r, hdr[1]) in nae and r < hdr[0] + 30:
        name = nae[(r, hdr[1])]
        if not isinstance(nae.get((r, hdr[1] + 1)), (int, float)):
            r += 1
            continue
        g = lambda k: nae.get((r, hdr[1] + k)) or 0
        prev_q, prev_amt = (g(10), g(11)) if a.next_round else (g(4), g(5))
        jobs.append(OrderedDict(직종=norm_job(name), 계약공수=g(1), 단가=g(2),
                                전회공수=prev_q, 전회금액=prev_amt))
        r += 1
    stamp = next((nae.get((r, 5)) for (r, c), v in nae.items() if v == "인지세"), 0)

    # 인원: 식별표 '이름 직책 공종 기성 (열1=투입 끝)'
    hr, hc = next((r, c) for (r, c), v in sik.items() if v == "이름" and sik.get((r, c + 3)) == "기성")
    people = []
    year = int(period[0][:4]) if period[0] else dt.date.today().year
    r = hr + 1
    while (r, hc) in sik:
        n = str(sik[(r, hc)]).strip()
        job = sik.get((r, hc + 3))
        if n and n != "공란" and job and str(job).strip():
            p = OrderedDict(이름=n, 직종=norm_job(job))
            end = sik.get((r, hc + 4))
            mm = re.search(r"(\d{1,2})/(\d{1,2})", str(end or ""))
            if mm:
                p["투입끝"] = f"{year}-{int(mm[1]):02d}-{int(mm[2]):02d}"
            people.append(p)
        r += 1

    # 공휴일: 식별표 '날짜 공휴일'
    holidays = {}
    hr2, hc2 = next((r, c) for (r, c), v in sik.items() if v == "공휴일" and sik.get((r, c - 1)) == "날짜")
    r = hr2 + 1
    while (r, hc2 - 1) in sik:
        try:
            holidays[to_date(sik[(r, hc2 - 1)]).isoformat()] = str(sik.get((r, hc2), ""))
        except Exception:
            pass
        r += 1

    # 회차별 금액: 내역서 '1회차 2회차 …' 합계 줄
    rounds = []
    heads = sorted((c, r, v) for (r, c), v in nae.items() if isinstance(v, str) and re.fullmatch(r"\d+회차", v))
    if heads:
        col0 = heads[0][0]
        sum_row = max(r for (r, c), v in nae.items() if c == col0 + 1 and isinstance(v, (int, float)))
        rounds = [nae.get((sum_row, c + 1), 0) for c, _, _ in heads]

    cfg = OrderedDict(
        현장=str(sin.get((2, 4), "")), 회차=(round_no + 1 if a.next_round and round_no else round_no),
        기간=(["", ""] if a.next_round else period), 가중치=weights, 직종=jobs, 인지세=stamp,
        인원=people, 공휴일=dict(sorted(holidays.items())), 이전회차금액=rounds,
        출처=Path(a.book).name)
    if a.next_round:
        cfg["이전회차금액"] = rounds  # 다음 회차는 이번 금회를 손으로 더해야 함(아래 안내)
    out = Path(a.out or "설정.json")
    out.write_text(json.dumps(cfg, ensure_ascii=False, indent=1), encoding="utf-8")
    print(f"{out} — 회차 {cfg['회차']} · 기간 {cfg['기간']} · 직종 {len(jobs)} · 인원 {len(people)} · 공휴일 {len(holidays)}")
    if a.next_round:
        print("다음 회차: '기간'을 채우고, 이번 회차 확정 금액을 '이전회차금액' 끝에 더할 것.")
    return 0


# ───────────────────────── 계산 ─────────────────────────
def load_records(folder, start, end):
    recs, files = [], []
    for f in sorted(Path(folder).glob("*.csv")):
        m = re.match(r"(\d{4}-\d{2}-\d{2})", f.name)
        if not m:
            continue
        d = dt.date.fromisoformat(m[1])
        if not (start <= d <= end):
            continue
        files.append(d)
        with open(f, encoding="utf-8-sig") as fh:
            for row in csv.DictReader(fh):
                if (row.get("이름") or "").strip():
                    recs.append((d, row["이름"].strip(), (row.get("구분") or "").strip(),
                                 (row.get("근거") or "").strip(), f.name))
    return recs, sorted(set(files))


def match_name(name, names):
    if name in names:
        return name, False
    near = [n for n in names if len(n) == len(name) and sum(x != y for x, y in zip(n, name)) == 1]
    return (near[0], True) if len(near) == 1 else (None, False)


def cmd_calc(a):
    cfg = json.loads(Path(a.config).read_text(encoding="utf-8"))
    start, end = (dt.date.fromisoformat(x) for x in cfg["기간"])
    W = cfg["가중치"]
    hol = {dt.date.fromisoformat(k): v for k, v in cfg["공휴일"].items()}
    people = {p["이름"]: p for p in cfg["인원"]}
    jobs = OrderedDict((j["직종"], j) for j in cfg["직종"])
    recs, file_days = load_records(a.folder, start, end)

    issues = []  # (중요도, 내용)
    def issue(level, msg):
        issues.append((level, msg))

    daily = []  # dict rows
    seen = {}
    for d, raw, gubun, why, fname in recs:
        name, fixed = match_name(raw, list(people))
        if not name:
            issue("확인", f"{d} {raw}: 인원 명단(설정 '인원')에 없음 — 직종을 몰라 공수·금액에서 빠짐 ({fname})")
            continue
        if fixed:
            issue("알림", f"{d} {raw} → {name} 으로 고쳐 읽음 ({fname})")
        if (d, name) in seen:
            issue("확인", f"{d} {name}: 같은 날 두 번 — 두 번째는 뺌 ({fname})")
            continue
        seen[(d, name)] = True
        p = people[name]
        if p.get("투입끝") and d > dt.date.fromisoformat(p["투입끝"]):
            issue("확인", f"{d} {name}: 투입 끝({p['투입끝']}) 뒤 기록 — 그대로 셈, 인원 설정 확인")
        if p.get("투입시작") and d < dt.date.fromisoformat(p["투입시작"]):
            issue("확인", f"{d} {name}: 투입 시작({p['투입시작']}) 전 기록 — 그대로 셈")
        if p["직종"] not in jobs:
            issue("중요", f"{name}: 직종 '{p['직종']}' 이 계약 직종에 없음 — 금액에서 빠짐")
        weekend = d.weekday() >= 5 or d in hol
        w = W["주말" if weekend else "평일"]
        if gubun in ZERO:
            counts = {}
        elif gubun in GUBUN:
            counts = GUBUN[gubun]
        else:
            issue("중요", f"{d} {name}: 구분 '{gubun}' 을 모름 — 0 으로 둠 ({fname})")
            counts = {}
        q = {k: counts.get(k, 0) * w[k] for k in KINDS}
        daily.append(OrderedDict(날짜=d, 요일="월화수목금토일"[d.weekday()],
                                 날구분=("공휴일:" + hol[d]) if d in hol else ("주말" if d.weekday() >= 5 else "평일"),
                                 이름=name, 직종=p["직종"], 구분=gubun, **q, 계=sum(q.values()), 근거=why))
        if weekend and sum(q.values()):
            issue("알림", f"{d} {name}: {'공휴일 ' + hol[d] if d in hol else '주말'} 근무 — 1.5배 가중치")

    # 기록 없는 평일
    d = start
    while d <= end:
        if d.weekday() < 5 and d not in hol and d not in file_days:
            issue("확인", f"{d}: 평일인데 기록 파일 없음 (출력일보 사진 누락?)")
        d += dt.timedelta(days=1)

    # 월간 집계 (직종 × 평일/주말 × 칸)
    agg = OrderedDict((j, {"평일": dict.fromkeys(KINDS, 0.0), "주말": dict.fromkeys(KINDS, 0.0)}) for j in jobs)
    per = defaultdict(lambda: {"평일": dict.fromkeys(KINDS, 0.0), "주말": dict.fromkeys(KINDS, 0.0)})
    for r in daily:
        key = "평일" if r["날구분"] == "평일" else "주말"
        for k in KINDS:
            if r["직종"] in agg:
                agg[r["직종"]][key][k] += r[k]
            per[r["이름"]][key][k] += r[k]

    # 기성 내역
    rows = []
    for jn, j in jobs.items():
        cur = sum(agg[jn]["평일"].values()) + sum(agg[jn]["주말"].values())
        contract_amt = j["계약공수"] * j["단가"]
        tot_q = j["전회공수"] + cur
        rows.append(OrderedDict(직종=jn, 계약공수=j["계약공수"], 단가=j["단가"], 계약금액=contract_amt,
                                전회공수=j["전회공수"], 전회금액=j["전회금액"], 금회공수=cur,
                                금회금액=cur * j["단가"], 누계공수=tot_q, 누계금액=int(tot_q * j["단가"]),
                                기성율=(tot_q / j["계약공수"]) if j["계약공수"] else 0,
                                잔여공수=j["계약공수"] - tot_q))
        if abs(j["전회공수"] * j["단가"] - j["전회금액"]) > 1:
            issue("확인", f"{jn}: 전회 공수×단가({j['전회공수'] * j['단가']:,.0f}) ≠ 전회 금액({j['전회금액']:,.0f})")
        if j["계약공수"] and tot_q > j["계약공수"]:
            issue("중요", f"{jn}: 누계 {tot_q:g} 공수 > 계약 {j['계약공수']:g} — {tot_q - j['계약공수']:g} 공수"
                         f"({(tot_q - j['계약공수']) * j['단가']:,.0f}원) 초과 청구")
        elif j["계약공수"] and tot_q / j["계약공수"] >= 0.9:
            issue("확인", f"{jn}: 기성율 {tot_q / j['계약공수']:.1%} — 잔여 {j['계약공수'] - tot_q:g} 공수")
    prev_total = sum(r["전회금액"] for r in rows)
    if cfg.get("이전회차금액") and abs(sum(cfg["이전회차금액"]) - prev_total) > 1:
        issue("중요", f"이전 회차 금액 합 {sum(cfg['이전회차금액']):,.0f} ≠ 직종별 전회 금액 합 {prev_total:,.0f}")

    write_xlsx(a, cfg, daily, per, agg, rows, issues)
    cur_total = sum(r["금회금액"] for r in rows)
    print(f"{cfg['회차']}회 {start}~{end}: 기록 {len(file_days)}일 · {len(daily)}건 · 금회 {sum(r['금회공수'] for r in rows):g} 공수 · {cur_total:,.0f}원")
    for lv in ("중요", "확인"):
        for l, m in issues:
            if l == lv:
                print(f"[{l}] {m}")
    n_info = sum(1 for l, _ in issues if l == "알림")
    if n_info:
        print(f"[알림] {n_info}건 — 결과 파일 '점검' 시트")
    return 0


def write_xlsx(a, cfg, daily, per, agg, rows, issues):
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    wb = openpyxl.Workbook()
    thin = Side(style="thin", color="999999")
    B = Border(left=thin, right=thin, top=thin, bottom=thin)
    HF, HP = Font(bold=True, color="FFFFFF"), PatternFill("solid", fgColor="1F4E78")
    TP = PatternFill("solid", fgColor="DDEBF7")

    def table(ws, head, data, widths=None, start=1, money=(), pct=(), qty=()):
        for j, h in enumerate(head, 1):
            c = ws.cell(start, j, h); c.font, c.fill, c.border = HF, HP, B
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        for i, r in enumerate(data, start + 1):
            for j, v in enumerate(r, 1):
                c = ws.cell(i, j, v); c.border = B
                h = head[j - 1]
                if h in money: c.number_format = "#,##0"
                elif h in pct: c.number_format = "0.0%"
                elif h in qty: c.number_format = "0.###"
                elif isinstance(v, dt.date): c.number_format = "yyyy-mm-dd"
        for j, w in enumerate(widths or [], 1):
            ws.column_dimensions[openpyxl.utils.get_column_letter(j)].width = w
        ws.freeze_panes = ws.cell(start + 1, 1)

    title = f"{cfg.get('현장', '')} 제{cfg['회차']}회 기성 · {cfg['기간'][0]} ~ {cfg['기간'][1]}"

    # 1. 기성 내역 (원래 3.하도급 기성 내역서 형태)
    ws = wb.active; ws.title = "기성 내역"
    ws["A1"] = title; ws["A1"].font = Font(bold=True, size=13)
    head = ["직종", "계약공수", "단가", "계약금액", "전회공수", "전회금액", "금회공수", "금회금액",
            "누계공수", "누계금액", "기성율", "잔여공수"]
    data = [list(r.values()) for r in rows]
    tot = ["합계", sum(r["계약공수"] for r in rows), "", sum(r["계약금액"] for r in rows) + cfg.get("인지세", 0),
           sum(r["전회공수"] for r in rows), sum(r["전회금액"] for r in rows),
           sum(r["금회공수"] for r in rows), sum(r["금회금액"] for r in rows),
           sum(r["누계공수"] for r in rows), sum(r["누계금액"] for r in rows), None,
           sum(r["잔여공수"] for r in rows)]
    tot[10] = tot[9] / tot[3] if tot[3] else 0
    data.append(["인지세", "", "", cfg.get("인지세", 0)] + [""] * 8)
    data.append(tot)
    table(ws, head, data, [16, 10, 10, 15, 10, 15, 10, 14, 10, 15, 8, 10], start=3,
          money=("단가", "계약금액", "전회금액", "금회금액", "누계금액"), pct=("기성율",),
          qty=("계약공수", "전회공수", "금회공수", "누계공수", "잔여공수"))
    for c in ws[3 + len(data)]: c.fill, c.font = TP, Font(bold=True)
    for i, r in enumerate(rows, 4):
        if r["잔여공수"] < 0:
            for c in ws[i]: c.fill = PatternFill("solid", fgColor="F8CBAD")
    n = 3 + len(data) + 2
    vat = tot[7] * 0.1
    for k, v in [("금회 공급가액", tot[7]), ("부가세 10%", vat), ("금회 합계", tot[7] + vat),
                 ("이전 회차 금액", " / ".join(f"{i}회 {x:,.0f}" for i, x in enumerate(cfg.get("이전회차금액", []), 1)))]:
        ws.cell(n, 1, k).font = Font(bold=True); c = ws.cell(n, 2, v)
        if isinstance(v, (int, float)): c.number_format = "#,##0"
        n += 1

    # 2. 월간 집계 (원래 4.월간 근무표 형태)
    ws = wb.create_sheet("월간 집계")
    head = ["구분", "단가", "총공수", "평일공수"] + [f"평일 {k}" for k in KINDS] + ["주말공수"] + [f"주말 {k}" for k in KINDS] + ["금액"]
    data = []
    for jn, g in agg.items():
        wd, we = sum(g["평일"].values()), sum(g["주말"].values())
        data.append([jn, cfg_job(cfg, jn)["단가"], wd + we, wd] + [g["평일"][k] for k in KINDS] + [we] + [g["주말"][k] for k in KINDS]
                    + [(wd + we) * cfg_job(cfg, jn)["단가"]])
    data.append(["합계", ""] + [sum(r[i] for r in data) for i in range(2, len(head))])
    table(ws, head, data, [14, 9] + [8] * 14 + [14], money=("단가", "금액"), qty=tuple(head[2:-1]))
    for c in ws[1 + len(data)]: c.fill, c.font = TP, Font(bold=True)
    r0 = len(data) + 4
    ws.cell(r0 - 1, 1, "가중치").font = Font(bold=True)
    for i, key in enumerate(("평일", "주말")):
        ws.cell(r0 + i, 1, key)
        for j, k in enumerate(KINDS):
            ws.cell(r0 + i, 2 + j, f"{k} {cfg['가중치'][key][k]:g}")

    # 3. 근무자별 (원래 6.근무자별 근무표)
    ws = wb.create_sheet("근무자별")
    s, e = (dt.date.fromisoformat(x) for x in cfg["기간"])
    days = [s + dt.timedelta(days=i) for i in range((e - s).days + 1)]
    hol = set(cfg["공휴일"])
    head = ["이름", "직종", "총공수", "평일", "주말·공휴"] + [f"{d.month}/{d.day}" for d in days]
    cell = {(r["이름"], r["날짜"]): r for r in daily}
    data = []
    for name, p in sorted(per.items(), key=lambda x: (cfg_person(cfg, x[0])["직종"], x[0])):
        wd, we = sum(p["평일"].values()), sum(p["주말"].values())
        data.append([name, cfg_person(cfg, name)["직종"], wd + we, wd, we] +
                    [cell[(name, d)]["계"] if (name, d) in cell else None for d in days])
    table(ws, head, data, [9, 12, 7, 6, 7] + [5.2] * len(days), qty=tuple(head[2:]))
    for j, d in enumerate(days, 6):
        if d.weekday() >= 5 or d.isoformat() in hol:
            ws.cell(1, j).fill = PatternFill("solid", fgColor="C00000")
    for i, row in enumerate(data, 2):
        for j, d in enumerate(days, 6):
            r = cell.get((row[0], d))
            if r and r["구분"] not in ("주간",):
                ws.cell(i, j).fill = PatternFill("solid", fgColor="FFF2CC")
                ws.cell(i, j).comment = None

    # 4. 일일 근무표 (원래 5.모든인원 일일 근무표)
    ws = wb.create_sheet("일일 근무표")
    head = ["날짜", "요일", "날구분", "이름", "직종", "구분"] + KINDS + ["계", "근거"]
    table(ws, head, [list(r.values()) for r in sorted(daily, key=lambda r: (r["날짜"], r["직종"], r["이름"]))],
          [11, 5, 12, 9, 12, 9, 6, 6, 6, 6, 6, 6, 40], qty=tuple(KINDS) + ("계",))
    ws.auto_filter.ref = ws.dimensions

    # 5. 점검
    ws = wb.create_sheet("점검")
    order = {"중요": 0, "확인": 1, "알림": 2}
    table(ws, ["중요도", "내용"], sorted(issues, key=lambda x: order[x[0]]), [8, 110])
    for i, (lv, _) in enumerate(sorted(issues, key=lambda x: order[x[0]]), 2):
        if lv == "중요":
            ws.cell(i, 1).fill = PatternFill("solid", fgColor="F8CBAD")

    out = Path(a.out or f"기성_{cfg['회차']}회.xlsx")
    wb.save(out)
    print(f"→ {out}")


def cfg_job(cfg, jn):
    return next(j for j in cfg["직종"] if j["직종"] == jn)


def cfg_person(cfg, n):
    return next(p for p in cfg["인원"] if p["이름"] == n)


# ───────────────────────── 출입기록 → 일일 기록 ─────────────────────────
T = dt.time
IN_BANDS = [(3, T(5, 1), T(7, 10)), (4, T(7, 10, 1), T(13, 10)), (5, T(13, 10, 1), T(23, 0)), (6, T(2, 0), T(5, 0, 59))]
OUT_BANDS = [(3, T(2, 0), T(11, 19, 59)), (4, T(11, 20), T(17, 19, 59)), (5, T(17, 20), T(19, 19, 59)),
             (6, T(19, 20), T(22, 59, 59)), (7, T(23, 0), T(23, 59, 59))]
CODE = {"35": "주간", "36": "연장", "37": "야간", "34": "오전", "45": "오후",
        "65": "조출", "66": "조출+연장", "67": "조출+야간"}


def band(t, bands):
    return next((c for c, x, y in bands if x <= t <= y), None)


def cmd_gate(a):
    book = read_book(a.book)
    sh = find_sheet(book, "출입기록")
    head = {v: c for (r, c), v in sh.items() if r == 1}
    ct, cn, cg = head["기계타각시간"], head["이름"], head.get("CtrlerNme")
    by = defaultdict(list)
    for (r, c), v in sh.items():
        if c == ct and r > 1:
            t = v if isinstance(v, dt.datetime) else dt.datetime.strptime(str(v)[:19], "%Y-%m-%d %H:%M:%S")
            by[(t.date(), str(sh.get((r, cn))).strip())].append((t, sh.get((r, cg), "")))
    out = Path(a.outdir); out.mkdir(parents=True, exist_ok=True)
    per_day = defaultdict(list)
    for (d, n), L in sorted(by.items()):
        L.sort()
        i, o = L[0][0].time(), L[-1][0].time()
        code = f"{band(i, IN_BANDS)}{band(o, OUT_BANDS)}"
        g = CODE.get(code, f"모름({code})")
        why = f"출입 {i:%H:%M}~{o:%H:%M} 코드 {code}"
        if g == "오후" and (L[-1][0] - L[0][0]) >= dt.timedelta(hours=9):
            g, why = "주간", why + " · 지각이나 8시간 준수(원래 엑셀 특이사항 규칙) — 확인"
        if i >= T(20, 30) or len(L) == 1:
            why += " · 철야/한쪽 타각 — 확인"
        per_day[d].append((n, g, why))
    for d, L in per_day.items():
        with open(out / f"{d}.csv", "w", encoding="utf-8", newline="") as f:
            w = csv.writer(f); w.writerow(["번호", "이름", "소속", "구분", "퇴근", "근거"])
            for k, (n, g, why) in enumerate(L, 1):
                w.writerow([k, n, "", g, "", why])
    print(f"{out}: {len(per_day)}일 · {sum(len(v) for v in per_day.values())}건")
    return 0


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = ap.add_subparsers(dest="cmd", required=True)
    p = sub.add_parser("설정"); p.add_argument("book"); p.add_argument("-o", "--out")
    p.add_argument("--다음회차", dest="next_round", action="store_true"); p.set_defaults(f=cmd_config)
    p = sub.add_parser("계산"); p.add_argument("config"); p.add_argument("folder"); p.add_argument("-o", "--out")
    p.set_defaults(f=cmd_calc)
    p = sub.add_parser("출입기록"); p.add_argument("book"); p.add_argument("outdir"); p.set_defaults(f=cmd_gate)
    a = ap.parse_args()
    return a.f(a)


if __name__ == "__main__":
    sys.exit(main())
