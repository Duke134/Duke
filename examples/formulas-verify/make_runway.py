"""GXC 런웨이 시트 템플릿 생성 (숫자 미입력, 입력 칸은 노란색).

시트 구성
  입력        기준월·기초현금·시나리오 비율
  월별현금흐름 18개월 (열 = 월). B열 '기본값'을 넣으면 매월 자동 반복, 특정 월만 셀을 덧씌워 수정
  요약        런웨이(개월)·현금 소진 예상월·최저 잔액월·비용 +10% 시나리오

규칙
  - 노란 칸만 입력. 흰 칸은 전부 수식.
  - 금액 단위: 원. 월 단위, 입금·출금 기준(발생 기준 아님).
"""
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as L

FONT = "맑은 고딕"
F = Font(name=FONT, size=10)
FB = Font(name=FONT, size=10, bold=True)
FH = Font(name=FONT, size=12, bold=True)
FNOTE = Font(name=FONT, size=9, color="808080")
YELLOW = PatternFill("solid", fgColor="FFF2CC")
GRAY = PatternFill("solid", fgColor="F2F2F2")
BLUE = PatternFill("solid", fgColor="DDEBF7")
THIN = Side(style="thin", color="BFBFBF")
BOX = Border(top=THIN, bottom=THIN, left=THIN, right=THIN)
MONTHS = 18
FIRST_COL = 3                      # C열부터 월
LAST_COL = FIRST_COL + MONTHS - 1  # T열
C0, C1 = L(FIRST_COL), L(LAST_COL)
NUM = "#,##0;[Red]-#,##0"


def put(ws, addr, v, font=F, fill=None, fmt=None, align=None):
    c = ws[addr]
    c.value = v
    c.font = font
    if fill:
        c.fill = fill
    if fmt:
        c.number_format = fmt
    if align:
        c.alignment = Alignment(horizontal=align)
    return c


def build(path="runway_template.xlsx"):
    wb = Workbook()

    # ───────────── 입력 ─────────────
    ws = wb.active
    ws.title = "입력"
    put(ws, "A1", "런웨이 시트 입력", FH)
    put(ws, "A2", "노란 칸만 입력. 금액은 원 단위, 월별현금흐름 시트는 입금·출금 기준.", FNOTE)
    rows = [
        ("기준월(1일 날짜)", None, "yyyy-mm", "예: 2026-11-01. 월별현금흐름 1열의 월이 됩니다"),
        ("기초현금(기준월 초 은행잔고 합계)", None, NUM, "법인 계좌 전체 합계. 외화는 기준월 초 환율로 환산"),
        ("비용 시나리오 증감률", 0.10, "0%", "요약 시트의 '비용 +N%' 시나리오에 쓰는 비율"),
        ("런웨이 계산 기준 개월수", 6, "0", "최근 N개월 평균 순소진으로 런웨이를 계산. 기본 6"),
    ]
    put(ws, "A4", "항목", FB, GRAY); put(ws, "B4", "값", FB, GRAY); put(ws, "C4", "메모", FB, GRAY)
    for i, (k, v, fmt, memo) in enumerate(rows, start=5):
        put(ws, f"A{i}", k)
        put(ws, f"B{i}", v, fill=YELLOW, fmt=fmt, align="right").border = BOX
        put(ws, f"C{i}", memo, FNOTE)
    ws.column_dimensions["A"].width = 34
    ws.column_dimensions["B"].width = 18
    ws.column_dimensions["C"].width = 60

    # ───────────── 월별현금흐름 ─────────────
    cf = wb.create_sheet("월별현금흐름")
    put(cf, "A1", "월별 현금흐름 (18개월)", FH)
    put(cf, "A2", "B열 '기본값'은 매월 자동 반복. 특정 월만 다르면 그 달 셀에 금액을 직접 입력(덧씌움).", FNOTE)
    put(cf, "A4", "항목", FB, GRAY)
    put(cf, "B4", "기본값(매월)", FB, GRAY)
    put(cf, "A3", "월 번호", FNOTE)
    for j in range(MONTHS):
        col = L(FIRST_COL + j)
        put(cf, f"{col}3", j + 1, FNOTE, fmt="0", align="center")
        if j == 0:
            put(cf, f"{col}4", "=입력!$B$5", FB, GRAY, "yyyy-mm", "center")
        else:
            put(cf, f"{col}4", f"=EDATE({L(FIRST_COL + j - 1)}4,1)", FB, GRAY, "yyyy-mm", "center")
        cf.column_dimensions[col].width = 14

    # 행 레이아웃
    r = 5
    put(cf, f"A{r}", "기초현금", FB)
    row_begin = r
    for j in range(MONTHS):
        col = L(FIRST_COL + j)
        f = "=입력!$B$6" if j == 0 else f"={L(FIRST_COL + j - 1)}{{END}}"  # END는 아래에서 치환
        put(cf, f"{col}{r}", f, fmt=NUM)

    def block(title, items, start_row):
        """소제목 + 입력 항목들 + 소계. 반환: (소계 행, 다음 행)"""
        rr = start_row
        put(cf, f"A{rr}", title, FB, BLUE)
        for j in range(FIRST_COL - 1, LAST_COL + 1):
            cf.cell(row=rr, column=j).fill = BLUE
        rr += 1
        first_item = rr
        for name in items:
            put(cf, f"A{rr}", name)
            put(cf, f"B{rr}", None, fill=YELLOW, fmt=NUM).border = BOX
            for j in range(MONTHS):
                col = L(FIRST_COL + j)
                put(cf, f"{col}{rr}", f"=$B{rr}", fmt=NUM)
            rr += 1
        last_item = rr - 1
        put(cf, f"A{rr}", f"{title} 합계", FB)
        for j in range(MONTHS):
            col = L(FIRST_COL + j)
            put(cf, f"{col}{rr}", f"=SUM({col}{first_item}:{col}{last_item})", FB, fmt=NUM)
        return rr, rr + 2

    inflow_items = ["매출 입금", "정부지원금 입금", "투자금 입금", "이자·기타 입금"]
    outflow_items = ["급여(실지급)", "4대보험·원천세 납부", "임차료·관리비", "외주·용역비",
                     "서버·SaaS·라이선스", "미국법인(GAME ROUND) 송금", "기타 운영비", "법인세·부가세 납부"]
    in_total, r = block("현금 유입", inflow_items, row_begin + 2)
    out_total, r = block("현금 유출", outflow_items, r)

    put(cf, f"A{r}", "순현금흐름", FB)
    net_row = r
    for j in range(MONTHS):
        col = L(FIRST_COL + j)
        put(cf, f"{col}{r}", f"={col}{in_total}-{col}{out_total}", FB, fmt=NUM)
    r += 1
    inv_row = in_total - 2  # '투자금 입금' 행 (유입 항목 3번째)
    put(cf, f"A{r}", "순현금흐름(투자금 제외, 런웨이 계산용)", FB)
    op_net_row = r
    for j in range(MONTHS):
        col = L(FIRST_COL + j)
        put(cf, f"{col}{r}", f"={col}{net_row}-{col}{inv_row}", FB, fmt=NUM)
    r += 1
    put(cf, f"A{r}", "기말현금", FB)
    end_row = r
    for j in range(MONTHS):
        col = L(FIRST_COL + j)
        put(cf, f"{col}{r}", f"={col}{row_begin}+{col}{net_row}", FB, GRAY, NUM)
    r += 1
    put(cf, f"A{r}", "현금 소진 플래그(기말<0)", FNOTE)
    flag_row = r
    for j in range(MONTHS):
        col = L(FIRST_COL + j)
        put(cf, f"{col}{r}", f"=IF({col}{end_row}<0,1,0)", FNOTE, fmt="0")
    r += 1
    put(cf, f"A{r}", "런웨이 계산 포함(1=포함)", FNOTE)
    incl_row = r
    for j in range(MONTHS):
        col = L(FIRST_COL + j)
        put(cf, f"{col}{r}", f"=IF({col}3<=입력!$B$8,1,0)", FNOTE, fmt="0")

    # 기초현금 수식의 END 치환
    for j in range(1, MONTHS):
        col = L(FIRST_COL + j)
        cf[f"{col}{row_begin}"].value = cf[f"{col}{row_begin}"].value.replace("{END}", str(end_row))

    cf.column_dimensions["A"].width = 30
    cf.column_dimensions["B"].width = 16
    cf.freeze_panes = "C5"

    # ───────────── 요약 ─────────────
    sm = wb.create_sheet("요약")
    rng_net = f"월별현금흐름!{C0}{op_net_row}:{C1}{op_net_row}"
    rng_inv = f"월별현금흐름!{C0}{inv_row}:{C1}{inv_row}"
    rng_end = f"월별현금흐름!{C0}{end_row}:{C1}{end_row}"
    rng_flag = f"월별현금흐름!{C0}{flag_row}:{C1}{flag_row}"
    rng_mon = f"월별현금흐름!{C0}4:{C1}4"
    rng_incl = f"월별현금흐름!{C0}{incl_row}:{C1}{incl_row}"
    rng_out = f"월별현금흐름!{C0}{out_total}:{C1}{out_total}"
    rng_in = f"월별현금흐름!{C0}{in_total}:{C1}{in_total}"
    n = "입력!$B$8"

    put(sm, "A1", "런웨이 요약", FH)
    put(sm, "A2", "전부 수식. 입력·월별현금흐름 시트만 고치면 자동 갱신.", FNOTE)
    put(sm, "A4", "항목", FB, GRAY); put(sm, "B4", "값", FB, GRAY); put(sm, "C4", "계산", FB, GRAY)
    items = [
        ("기초현금", "=입력!B6", NUM, "입력!B6"),
        ("향후 N개월 평균 순소진(투자금 제외)", f"=SUMPRODUCT({rng_incl},{rng_net})/{n}", NUM,
         "기준월부터 N개월(입력!B8) 평균. 투자금 입금은 제외, 매출·지원금은 포함"),
        ("런웨이(개월)", "=IF(B6<0,B5/ABS(B6),\"영업 흑자(소진 없음)\")", "0.0",
         "기초현금 ÷ |평균 순소진|. 예정 투자금은 반영하지 않은 보수적 값"),
        ("현금 소진 예상월", f"=IFERROR(INDEX({rng_mon},MATCH(1,{rng_flag},0)),\"18개월 내 없음\")", "yyyy-mm",
         "기말현금이 처음 0 미만이 되는 월"),
        ("18개월 내 최저 기말현금", f"=MIN({rng_end})", NUM, ""),
        ("최저 기말현금 월", f"=INDEX({rng_mon},MATCH(B9,{rng_end},0))", "yyyy-mm", ""),
        ("18개월 유입 합계", f"=SUM({rng_in})", NUM, ""),
        ("18개월 유출 합계", f"=SUM({rng_out})", NUM, ""),
    ]
    for i, (k, f, fmt, memo) in enumerate(items, start=5):
        put(sm, f"A{i}", k)
        put(sm, f"B{i}", f, FB, fmt=fmt, align="right")
        put(sm, f"C{i}", memo, FNOTE)

    put(sm, "A14", "시나리오: 비용 증감", FB, BLUE)
    for j in range(1, 4):
        sm.cell(row=14, column=j).fill = BLUE
    put(sm, "A15", "비용 증감률")
    put(sm, "B15", "=입력!B7", FB, fmt="+0%;-0%", align="right")
    put(sm, "C15", "입력!B7", FNOTE)
    put(sm, "A16", "시나리오 평균 순현금흐름")
    put(sm, "B16", f"=(SUMPRODUCT({rng_incl},{rng_in})-SUMPRODUCT({rng_incl},{rng_inv})"
                   f"-SUMPRODUCT({rng_incl},{rng_out})*(1+B15))/{n}",
        FB, fmt=NUM, align="right")
    put(sm, "C16", "(유입 − 투자금) 평균 − 유출 평균 × (1+증감률)", FNOTE)
    put(sm, "A17", "시나리오 런웨이(개월)")
    put(sm, "B17", "=IF(B16<0,B5/ABS(B16),\"영업 흑자(소진 없음)\")", FB, fmt="0.0", align="right")
    put(sm, "A18", "런웨이 변화(개월)")
    put(sm, "B18", "=IF(AND(ISNUMBER(B7),ISNUMBER(B17)),B17-B7,\"-\")", FB, fmt="+0.0;-0.0", align="right")

    sm.column_dimensions["A"].width = 30
    sm.column_dimensions["B"].width = 20
    sm.column_dimensions["C"].width = 44

    wb.save(path)
    layout = dict(row_begin=row_begin, in_total=in_total, out_total=out_total,
                  net_row=net_row, op_net_row=op_net_row, end_row=end_row, flag_row=flag_row, incl_row=incl_row)
    print("saved", path, layout)
    return layout


if __name__ == "__main__":
    build()
