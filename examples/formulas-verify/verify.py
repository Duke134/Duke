"""런웨이 워크북(make_runway.py 레이아웃) 재검산.

  python verify.py runway_example.xlsx
  python verify.py runway_example.xlsx --set "입력!B6=900000000"   # 기초현금 시나리오

하는 일
  1. 입력·원자료 데이터 검사: 필수값, 구분·주기·통화·확정 값, 날짜 순서, USD 환율
  2. 수식 구조 검사: 월별현금흐름·전개표·요약·보고에 손으로 친 숫자가 없는지,
     월(열)·줄(행) 방향으로 수식 패턴이 일관된지 (openpyxl Translator)
  3. formulas 라이브러리로 전체 재계산 (Excel 불필요)
  4. 독립 전개: 원자료를 파이썬으로 직접 월별 전개해 월별현금흐름·요약과 대조 (허용오차 ±0.5원)
  5. Excel 저장값이 있으면 재계산값과 대조 (수동 계산 모드·저장 후 수정 탐지)
  6. 4필드 보고 + ✅ 검증완료 / ❌ 검증 실패 (exit 0 / 1)
"""
import argparse
import sys
from datetime import date, datetime
from pathlib import Path

from formulas import ExcelModel
from openpyxl import load_workbook
from openpyxl.formula.translate import Translator
from openpyxl.utils import get_column_letter as L

TOL = 0.5
CATS_OUT = ["급여(실지급)", "4대보험·원천세 납부", "임차료·관리비", "외주·용역비",
            "서버·SaaS·라이선스", "미국법인(GAME ROUND) 송금", "기타 운영비", "법인세·부가세 납부"]
CATS_IN = ["매출 입금", "정부지원금 입금", "투자금 입금", "이자·기타 입금"]
PERIODS = {"매월": 1, "분기": 3, "반기": 6, "연": 12, "1회": 0}
MONTHS = 18
LG_FIRST = 5
CF_FIRST, EX_FIRST = 3, 8


# ───────────────────────── 유틸 ─────────────────────────
def is_formula(v):
    return isinstance(v, str) and v.startswith("=")


def to_date(v):
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    return None


def idx(d):
    return d.year * 12 + d.month


def find_row(ws, label, col="A", start=1, end=200):
    for r in range(start, end + 1):
        if ws[f"{col}{r}"].value == label:
            return r
    return None


def fmt_won(x):
    return f"{x:,.0f}"


class Report:
    def __init__(self):
        self.data, self.struct, self.calc, self.notes = [], [], [], []


# ───────────────────────── 1) 데이터 검사 ─────────────────────────
def read_ledger(ws, cats, rep, name):
    rows = []
    r = LG_FIRST
    while r < LG_FIRST + 200:
        vals = [ws[f"{L(c)}{r}"].value for c in range(1, 10)]
        if all(v in (None, "") for v in vals):
            # 비어 있는 줄 — 이후 데이터가 있으면 '중간 빈 줄'은 허용
            r += 1
            if r > LG_FIRST + 60:
                break
            continue
        cat, item, amt, cur, per, s, e, st, memo = vals
        where = f"{name}!{r}행"
        if cat not in cats:
            rep.data.append(f"{where}: 구분 '{cat}'이(가) 목록에 없음")
        if not isinstance(amt, (int, float)) or amt <= 0:
            rep.data.append(f"{where}: 금액이 양수가 아님 ({amt!r})")
        if cur not in ("KRW", "USD"):
            rep.data.append(f"{where}: 통화는 KRW/USD ({cur!r})")
        if per not in PERIODS:
            rep.data.append(f"{where}: 주기 '{per}'이(가) 목록에 없음 (매월·분기·반기·연·1회)")
        sd, ed = to_date(s), to_date(e)
        if sd is None:
            rep.data.append(f"{where}: 시작월이 날짜가 아님 ({s!r})")
        if e not in (None, "") and ed is None:
            rep.data.append(f"{where}: 종료월이 날짜가 아님 ({e!r})")
        if sd and ed and idx(ed) < idx(sd):
            rep.data.append(f"{where}: 종료월 {ed:%Y-%m}이 시작월 {sd:%Y-%m}보다 앞섬")
        if st not in ("확정", "예정"):
            rep.data.append(f"{where}: 확정 칸은 '확정' 또는 '예정' ({st!r})")
        if per == "1회" and ed:
            rep.notes.append(f"{where}: 1회 항목에 종료월이 있음(무시됨)")
        rows.append(dict(row=r, cat=cat, item=item, amt=amt, cur=cur, per=per, s=sd, e=ed, st=st))
        r += 1
    return rows


# ───────────────────────── 2) 구조 검사 ─────────────────────────
def check_row_pattern(ws, row, c_first, c_last, rep, label, base_col=None):
    """한 행의 월 셀들이 첫 월 셀 수식의 평행이동인지 확인."""
    base_col = base_col or c_first
    base = ws.cell(row=row, column=base_col).value
    if not is_formula(base):
        rep.struct.append(f"{ws.title}!{L(base_col)}{row} ({label}): 수식이 아님 ({base!r})")
        return
    for c in range(c_first, c_last + 1):
        if c == base_col:
            continue
        actual = ws.cell(row=row, column=c).value
        expected = Translator(base, origin=f"{L(base_col)}{row}").translate_formula(f"{L(c)}{row}")
        if not is_formula(actual):
            rep.struct.append(f"{ws.title}!{L(c)}{row} ({label}): 수식 자리에 값 ({actual!r}), 기대 {expected}")
        elif actual.replace(" ", "") != expected.replace(" ", ""):
            rep.struct.append(f"{ws.title}!{L(c)}{row} ({label}): 수식 불일치 실제 {actual} / 기대 {expected}")


def check_col_pattern(ws, col, r_first, r_last, rep, label):
    """한 열의 줄들이 첫 줄 수식의 평행이동인지 확인 (전개표 행 방향)."""
    base = ws.cell(row=r_first, column=col).value
    if not is_formula(base):
        rep.struct.append(f"{ws.title}!{L(col)}{r_first} ({label}): 수식이 아님 ({base!r})")
        return
    for r in range(r_first + 1, r_last + 1):
        actual = ws.cell(row=r, column=col).value
        expected = Translator(base, origin=f"{L(col)}{r_first}").translate_formula(f"{L(col)}{r}")
        if not is_formula(actual):
            rep.struct.append(f"{ws.title}!{L(col)}{r} ({label}): 수식 자리에 값 ({actual!r})")
        elif actual.replace(" ", "") != expected.replace(" ", ""):
            rep.struct.append(f"{ws.title}!{L(col)}{r} ({label}): 수식 불일치 실제 {actual} / 기대 {expected}")


def structural_checks(wb, cf_rows, ex_blocks, rep):
    cf = wb["월별현금흐름"]
    c_last = CF_FIRST + MONTHS - 1
    for label, row in cf_rows.items():
        if label == "기초현금":
            check_row_pattern(cf, row, CF_FIRST + 1, c_last, rep, label, base_col=CF_FIRST + 1)
            if not is_formula(cf.cell(row=row, column=CF_FIRST).value):
                rep.struct.append(f"월별현금흐름!{L(CF_FIRST)}{row} (기초현금 첫 달): 수식이 아님")
        else:
            check_row_pattern(cf, row, CF_FIRST, c_last, rep, label)
    ex = wb["전개표"]
    e_last = EX_FIRST + MONTHS - 1
    for name, (first, last) in ex_blocks.items():
        for c in range(1, 8):  # A~G 보조열: 행 방향 패턴
            check_col_pattern(ex, c, first, last, rep, f"{name} {L(c)}열")
        for r in range(first, last + 1):  # 모든 줄: 열 방향 패턴 (월 셀 전체)
            check_row_pattern(ex, r, EX_FIRST, e_last, rep, f"{name} {r}행")
        check_col_pattern(ex, EX_FIRST, first, last, rep, f"{name} 첫 월")
    for sheet in ("요약", "보고"):
        ws = wb[sheet]
        for row in ws.iter_rows(min_row=1, max_row=80):
            for cell in row:
                v = cell.value
                if isinstance(v, (int, float)) and not isinstance(v, bool):
                    rep.struct.append(f"{sheet}!{cell.coordinate}: 손으로 친 숫자 ({v!r})")
    sc = wb["시나리오"]
    for r in range(13, 21):
        for col in "BCD":
            if not is_formula(sc[f"{col}{r}"].value):
                rep.struct.append(f"시나리오!{col}{r} (결과표): 수식이 아님 ({sc[f'{col}{r}'].value!r})")
    r = 24
    while r < 80:
        v = sc[f"A{r}"].value
        if is_formula(v) and "시나리오" in v:  # 블록 제목
            for k in range(1, 8):
                lb = f"시나리오 {sc[f'A{r + k}'].value}"
                if k == 5:  # 기말현금: 첫 달은 기초현금 참조라 둘째 달부터 패턴 검사
                    check_row_pattern(sc, r + k, CF_FIRST + 1, CF_FIRST + MONTHS - 1, rep, lb, base_col=CF_FIRST + 1)
                    if not is_formula(sc.cell(row=r + k, column=CF_FIRST).value):
                        rep.struct.append(f"시나리오!{L(CF_FIRST)}{r + k} ({lb} 첫 달): 수식이 아님")
                else:
                    check_row_pattern(sc, r + k, CF_FIRST, CF_FIRST + MONTHS - 1, rep, lb)
            r += 9
        else:
            r += 1


# ───────────────────────── 4) 독립 전개 ─────────────────────────
def expand(rows, months, fx):
    out = []
    for r in rows:
        if r["s"] is None or r["per"] not in PERIODS or not isinstance(r["amt"], (int, float)):
            continue
        amt = r["amt"] * fx if r["cur"] == "USD" else r["amt"]
        k = PERIODS[r["per"]]
        si = idx(r["s"])
        ei = idx(r["e"]) if r["e"] else 10 ** 9
        sched = []
        for m in months:
            hit = si <= m <= ei and ((m == si) if k == 0 else ((m - si) % k == 0))
            sched.append(amt if hit else 0)
        out.append((r["cat"], r["st"], sched))
    return out


def series_sum(rows, pred=lambda c, s: True):
    return [sum(sch[i] for c, s, sch in rows if pred(c, s)) for i in range(MONTHS)]


# ───────────────────────── 메인 ─────────────────────────
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("xlsx")
    ap.add_argument("--set", action="append", default=[], metavar="입력!셀=값")
    args = ap.parse_args()
    path = Path(args.xlsx)
    fname = path.name
    rep = Report()

    wb = load_workbook(path)
    wb_v = load_workbook(path, data_only=True)
    for need in ("입력", "원자료_유출", "원자료_유입", "전개표", "월별현금흐름", "요약", "시나리오", "보고"):
        if need not in wb.sheetnames:
            print(f"❌ 시트 '{need}' 없음. make_runway.py 레이아웃이 아닙니다.")
            return 1

    overrides = {}
    for s_ in args.set:
        ref, v = s_.split("=")
        sheet, addr = ref.split("!")
        if sheet != "입력":
            print("--set 은 입력 시트 셀만 지원합니다."); return 1
        overrides[addr.upper()] = float(v)

    def inp(addr):
        if addr in overrides:
            return overrides[addr]
        return wb["입력"][addr].value

    # 1) 입력 검사
    base = to_date(inp("B5")) if not isinstance(inp("B5"), float) else None
    cash0, scen, n_win, fx, safe_n = inp("B6"), inp("B7"), inp("B8"), inp("B9"), inp("B10")
    if base is None:
        rep.data.append("입력!B5 기준월이 비었거나 날짜가 아님")
    if not isinstance(cash0, (int, float)):
        rep.data.append("입력!B6 기초현금 미기재")
    if not isinstance(n_win, (int, float)) or not (1 <= n_win <= MONTHS):
        rep.data.append(f"입력!B8 런웨이 기준 개월수는 1~{MONTHS} ({n_win!r})")
    if not isinstance(safe_n, (int, float)) or safe_n < 0:
        rep.data.append(f"입력!B10 안전 잔액 기준 개월수 확인 ({safe_n!r})")
    out_rows = read_ledger(wb["원자료_유출"], CATS_OUT, rep, "원자료_유출")
    in_rows = read_ledger(wb["원자료_유입"], CATS_IN, rep, "원자료_유입")
    if any(r["cur"] == "USD" for r in out_rows + in_rows) and not isinstance(fx, (int, float)):
        rep.data.append("원자료에 USD 줄이 있는데 입력!B9 환율이 비어 있음")
    if not out_rows and not in_rows:
        rep.data.append("원자료가 비어 있음 (템플릿 상태)")

    # 행 위치 (라벨로 찾음)
    cf = wb["월별현금흐름"]
    labels = ["기초현금"] + CATS_IN + ["현금 유입 합계"] + CATS_OUT + ["현금 유출 합계", "순현금흐름",
              "순현금흐름(투자금 제외, 런웨이 계산용)", "확정 유입 합계", "확정 투자금 입금",
              "보수적 순현금흐름(확정 유입만, 투자금 제외)", "기말현금",
              "현금 소진 플래그(기말<0)", "런웨이 계산 포함(1=포함)", "안전 잔액선(N개월치 평균 유출)",
              "안전선 미달 플래그(기말<안전선)"]
    cf_rows = {}
    for lb in labels:
        r = find_row(cf, lb)
        if r is None:
            rep.struct.append(f"월별현금흐름: '{lb}' 행을 찾지 못함")
        else:
            cf_rows[lb] = r
    ex = wb["전개표"]
    ex_blocks = {}
    for name in ("유출", "유입"):
        hdr = find_row(ex, name)
        if hdr is None:
            rep.struct.append(f"전개표: '{name}' 블록 없음"); continue
        first = hdr + 3
        last = first
        while is_formula(ex[f"A{last + 1}"].value):
            last += 1
        ex_blocks[name] = (first, last)
    sm = wb["요약"]
    sm_rows = {lb: find_row(sm, lb) for lb in [
        "향후 N개월 평균 순소진(투자금 제외)", "런웨이(개월)", "현금 소진 예상월", "18개월 내 최저 기말현금",
        "보수적 평균 순소진(확정 유입만)", "보수적 런웨이(개월)", "향후 N개월 평균 월 유출", "안전 잔액선",
        "안전선 미달 예상월", "시나리오 런웨이(개월)"]}
    for lb, r in sm_rows.items():
        if r is None:
            rep.struct.append(f"요약: '{lb}' 행을 찾지 못함")

    # 2) 구조 검사
    if cf_rows and ex_blocks:
        structural_checks(wb, cf_rows, ex_blocks, rep)

    # 3) formulas 재계산
    xl = ExcelModel().loads(str(path)).finish()
    f_inputs = {f"'[{fname}]입력'!{a}": v for a, v in overrides.items()}
    sol = xl.calculate(inputs=f_inputs) if f_inputs else xl.calculate()

    def val(sheet, addr):
        tgt = f"]{sheet}'!{addr}".upper()
        for k in sol:
            if k.upper().endswith(tgt):
                v = sol[k].value
                try:
                    v = v[0][0]
                except Exception:
                    pass
                return float(v) if isinstance(v, (int, float)) or hasattr(v, "dtype") else v
        return None

    err_cells = []
    for k in sol:
        try:
            t = str(sol[k].value)
        except Exception:
            continue
        if "#" in t and "!" in k:
            err_cells.append(k.split("]")[-1])
    if err_cells:
        rep.struct.append(f"재계산 오류 셀 {len(err_cells)}개: {', '.join(err_cells[:6])}")

    # 4) 독립 전개 대조
    can_calc = base is not None and isinstance(cash0, (int, float)) and isinstance(n_win, (int, float)) and cf_rows and not rep.data
    f_run = f_cons = f_safe = None
    scen_summary = []
    if can_calc:
        months = [idx(base) + i for i in range(MONTHS)]
        fxv = fx if isinstance(fx, (int, float)) else 1
        EO, EI = expand(out_rows, months, fxv), expand(in_rows, months, fxv)
        inflow, outflow = series_sum(EI), series_sum(EO)
        inv = series_sum(EI, lambda c, s: c == "투자금 입금")
        conf_in = series_sum(EI, lambda c, s: s == "확정" and c != "투자금 입금")
        net = [a - b for a, b in zip(inflow, outflow)]
        opnet = [a - b for a, b in zip(net, inv)]
        cons = [a - b for a, b in zip(conf_in, outflow)]
        end, c = [], cash0
        for x in net:
            c += x; end.append(c)
        N = int(n_win)
        avg, cavg, oavg = sum(opnet[:N]) / N, sum(cons[:N]) / N, sum(outflow[:N]) / N
        safe = oavg * (safe_n or 0)
        conf_all = series_sum(EI, lambda c, s: s == "확정")
        conf_inv = series_sum(EI, lambda c, s: s == "확정" and c == "투자금 입금")
        exp_rows = {"현금 유입 합계": inflow, "현금 유출 합계": outflow, "순현금흐름": net,
                    "순현금흐름(투자금 제외, 런웨이 계산용)": opnet, "확정 유입 합계": conf_all, "확정 투자금 입금": conf_inv,
                    "보수적 순현금흐름(확정 유입만, 투자금 제외)": cons, "기말현금": end}
        for cat in CATS_IN:
            exp_rows[cat] = series_sum(EI, lambda c, s, cat=cat: c == cat)
        for cat in CATS_OUT:
            exp_rows[cat] = series_sum(EO, lambda c, s, cat=cat: c == cat)
        for lb, series in exp_rows.items():
            row = cf_rows.get(lb)
            if not row:
                continue
            for i, e in enumerate(series):
                got = val("월별현금흐름", f"{L(CF_FIRST + i)}{row}")
                if not isinstance(got, float) or abs(got - e) > TOL:
                    rep.calc.append(f"월별현금흐름!{L(CF_FIRST + i)}{row} {lb}: 시트 {got} vs 독립 {fmt_won(e)}")
                    break  # 행당 첫 불일치만
        run_e = cash0 / abs(avg) if avg < 0 else None
        cons_e = cash0 / abs(cavg) if cavg < 0 else None
        exp_sum = {"향후 N개월 평균 순소진(투자금 제외)": avg, "보수적 평균 순소진(확정 유입만)": cavg,
                   "향후 N개월 평균 월 유출": oavg, "안전 잔액선": safe, "18개월 내 최저 기말현금": min(end)}
        if run_e is not None:
            exp_sum["런웨이(개월)"] = run_e
        if cons_e is not None:
            exp_sum["보수적 런웨이(개월)"] = cons_e
        for lb, e in exp_sum.items():
            row = sm_rows.get(lb)
            got = val("요약", f"B{row}") if row else None
            if not isinstance(got, float) or abs(got - e) > (0.01 if "개월" in lb else TOL):
                rep.calc.append(f"요약!B{row} {lb}: 시트 {got} vs 독립 {e:,.2f}")
        # 소진·안전선 미달 월
        def first_month(flags):
            for i, f_ in enumerate(flags):
                if f_:
                    m = months[i]; return f"{(m - 1) // 12}-{(m - 1) % 12 + 1:02d}"
            return "18개월 내 없음"
        def sheet_month(v):
            if isinstance(v, float):
                d = date(1899, 12, 30).toordinal() + int(v)
                d = date.fromordinal(d); return f"{d.year}-{d.month:02d}"
            return v
        for lb, flags in (("현금 소진 예상월", [x < 0 for x in end]), ("안전선 미달 예상월", [x < safe for x in end])):
            row = sm_rows.get(lb)
            got = sheet_month(val("요약", f"B{row}")) if row else None
            if got != first_month(flags):
                rep.calc.append(f"요약!B{row} {lb}: 시트 {got} vs 독립 {first_month(flags)}")
        f_run, f_cons, f_safe = val("요약", f"B{sm_rows['런웨이(개월)']}"), val("요약", f"B{sm_rows['보수적 런웨이(개월)']}"), val("요약", f"B{sm_rows['안전 잔액선']}")

        # 시나리오 3열 독립 계산
        sc = wb["시나리오"]
        for pc in "BCD":
            name = sc[f"{pc}5"].value
            rate = sc[f"{pc}6"].value or 0
            incl_planned = (sc[f"{pc}7"].value or "포함") == "포함"
            extra = sc[f"{pc}8"].value or 0
            ex_start = to_date(sc[f"{pc}9"].value)
            adj = sc[f"{pc}10"].value or 0
            if not all(isinstance(x, (int, float)) for x in (rate, extra, adj)):
                rep.data.append(f"시나리오!{pc}6/8/10: 숫자가 아님"); continue
            if sc[f"{pc}7"].value not in ("포함", "제외"):
                rep.data.append(f"시나리오!{pc}7: '포함' 또는 '제외' ({sc[f'{pc}7'].value!r})"); continue
            s_in = [(inflow[i] - inv[i]) if incl_planned else (conf_all[i] - conf_inv[i]) for i in range(MONTHS)]
            s_inv = inv if incl_planned else conf_inv
            s_out = [outflow[i] * (1 + rate) + (extra if (ex_start is None or months[i] >= idx(ex_start)) else 0) for i in range(MONTHS)]
            s_net = [a - b for a, b in zip(s_in, s_out)]
            s_end, c = [], cash0 + adj
            for i in range(MONTHS):
                c += s_net[i] + s_inv[i]; s_end.append(c)
            s_avg = sum(s_net[:N]) / N
            s_run = (cash0 + adj) / abs(s_avg) if s_avg < 0 else None
            checks = {13: s_avg, 17: s_end[-1], 18: min(s_end)}
            if s_run is not None:
                checks[14] = s_run
            for row, e in checks.items():
                got = val("시나리오", f"{pc}{row}")
                if not isinstance(got, float) or abs(got - e) > (0.01 if row == 14 else TOL):
                    rep.calc.append(f"시나리오!{pc}{row} ({name}): 시트 {got} vs 독립 {e:,.2f}")
            for row, flags in ((16, [x < 0 for x in s_end]), (20, [x < safe for x in s_end])):
                got = sheet_month(val("시나리오", f"{pc}{row}"))
                if got != first_month(flags):
                    rep.calc.append(f"시나리오!{pc}{row} ({name}): 시트 {got} vs 독립 {first_month(flags)}")
            scen_summary.append((name, val("시나리오", f"{pc}14"), sheet_month(val("시나리오", f"{pc}16"))))

    # 5) Excel 저장값 대조
    cache_note = "파일에 Excel 저장값 없음(생성 직후). 재계산값만 사용"
    if cf_rows.get("기말현금"):
        cached = wb_v["월별현금흐름"].cell(row=cf_rows["기말현금"], column=CF_FIRST + MONTHS - 1).value
        recalc = val("월별현금흐름", f"{L(CF_FIRST + MONTHS - 1)}{cf_rows['기말현금']}")
        if isinstance(cached, (int, float)):
            if not overrides and isinstance(recalc, float) and abs(cached - recalc) > TOL:
                rep.calc.append(f"기말현금 마지막 달 Excel 저장값 {fmt_won(cached)} vs 재계산 {fmt_won(recalc)} (수동 계산 모드 또는 저장 후 수정 의심)")
                cache_note = "저장값과 재계산값 불일치"
            else:
                cache_note = "Excel 저장값과 재계산값 일치"

    # ───────────── 보고 ─────────────
    ok = not (rep.data or rep.struct or rep.calc)
    print(f"\n📄 {fname}" + (f"   시나리오: {', '.join(args.set)}" if args.set else ""))
    print("=" * 66)
    print("결과")
    if can_calc:
        mo = lambda x: f"{x:.1f}개월" if isinstance(x, float) else str(x)
        print(f"  런웨이(확정+예정)        {mo(f_run):>18}")
        print(f"  보수적 런웨이(확정 유입만){mo(f_cons):>17}")
        print(f"  안전 잔액선              {fmt_won(f_safe):>18} 원")
        print(f"  기말현금(마지막 달)      {fmt_won(val('월별현금흐름', f'{L(CF_FIRST + MONTHS - 1)}{cf_rows['기말현금']}')):>18} 원")
        for name, rv, sm_ in scen_summary:
            print(f"  시나리오 · {name}: 런웨이 {mo(rv)}, 소진 예상 {sm_}")
    else:
        print("  (입력·원자료 미비로 숫자 대조 생략)")
    print("계산과정")
    print(f"  원자료 유출 {len(out_rows)}줄, 유입 {len(in_rows)}줄 → 주기·시작·종료월로 {MONTHS}개월 전개 → 구분별 합산")
    print(f"  런웨이 = 기초현금 ÷ |기준월부터 {n_win}개월 평균 순소진(투자금 제외)|, 보수적은 유입 중 '확정'만")
    print(f"  안전 잔액선 = 같은 구간 평균 월 유출 × {safe_n}개월")
    print("근거")
    print(f"  {fname} 입력!B5:B10, 원자료_유출·유입 전체, 수식 원문 (formulas {__import__('formulas').__version__} 재계산 + 파이썬 독립 전개)")
    print("주의사항")
    print(f"  Excel 저장값: {cache_note}. 월별현금흐름은 행당 첫 불일치만 표시.")
    for n_ in rep.notes:
        print(f"  참고: {n_}")
    print("-" * 66)
    calc_ok = f"월별 {len(cf_rows)}행·요약·시나리오 3열 전부 일치 (±{TOL}원)" if can_calc else "⏭ 생략 (입력·원자료 오류를 먼저 고친 뒤 재실행)"
    for title, items, okmsg in (("입력·원자료 검사", rep.data, "필수값·목록·날짜 이상 없음"),
                                 ("수식 구조 검사", rep.struct, "하드코딩·수식 불일치·오류 셀 없음"),
                                 ("독립 전개 대조", rep.calc, calc_ok)):
        if items:
            print(f"{title} ❌")
            for it in items[:15]:
                print("  -", it)
            if len(items) > 15:
                print(f"  … 외 {len(items) - 15}건")
        else:
            print(f"{title} {'✅ ' if not okmsg.startswith('⏭') else ''}{okmsg}")
    print("=" * 66)
    print("✅ 검증완료" if ok else "❌ 검증 실패 — 위 항목 수정 전에는 '검증완료'를 붙이지 않음")
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
