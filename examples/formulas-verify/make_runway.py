"""GXC 런웨이 시트 생성기.

  python make_runway.py            → runway_template.xlsx (빈 템플릿)
  python make_runway.py --example  → runway_example.xlsx  (가상 원자료 채운 예시본)

시트
  보고          대표님·이사님 공유용 한 장 (헤드라인·보고 문구·차트 2개)
  입력          기준월·기초현금·시나리오율·런웨이 기준 개월수·USD 환율
  원자료_유출   고정비 목록. 한 줄에 비용 하나 (급여는 급여대장 총액 한 줄)
  원자료_유입   매출·정부지원금·투자금·이자기타 목록
  전개표        원자료 각 줄 × 18개월 자동 전개 (검산용, 손대지 않음)
  월별현금흐름  전개표를 구분별로 합산. 전부 수식
  요약          런웨이(기본·보수적)·소진 예상월·최저 잔액·비용 시나리오

규칙
  - 노란 칸(입력·원자료)만 입력. 나머지는 전부 수식.
  - 금액 단위 원. 입금·출금 기준. USD는 입력!B9 환율로 환산.
"""
import sys
from datetime import date

from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.chart.legend import Legend
from openpyxl.chart.series import SeriesLabel
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.chart.text import RichText, Text
from openpyxl.chart.title import Title
from openpyxl.drawing.line import LineProperties
from openpyxl.drawing.text import CharacterProperties, Font as DFont, Paragraph, ParagraphProperties, RegularTextRun
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as L
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.pagebreak import Break

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
NUM = "#,##0;[Red]-#,##0"
C_BLUE, C_ORANGE = "2A78D6", "EB6834"
C_INK, C_INK2, C_MUTED, C_GRID = "0B0B0B", "52514E", "8A8984", "E5E4E0"

MONTHS = 18
CATS_OUT = ["급여(실지급)", "4대보험·원천세 납부", "임차료·관리비", "외주·용역비",
            "서버·SaaS·라이선스", "미국법인(GAME ROUND) 송금", "기타 운영비", "법인세·부가세 납부"]
CATS_IN = ["매출 입금", "정부지원금 입금", "투자금 입금", "이자·기타 입금"]
PERIODS = ["매월", "분기", "반기", "연", "1회"]
N_OUT, N_IN = 40, 25          # 원자료 행 수
LG_FIRST = 5                  # 원자료 데이터 시작 행 (4행 헤더)

# 전개표 열: A 구분 B 항목 C 확정 D 금액(원) E 시작idx F 종료idx G 주기개월, 월은 H부터
EX_FIRST = 8
EX_LAST = EX_FIRST + MONTHS - 1
# 월별현금흐름 열: A 항목 B 건수, 월은 C부터
CF_FIRST = 3
CF_LAST = CF_FIRST + MONTHS - 1

EXAMPLE_OUT = [
    ("급여(실지급)", "급여 실지급 총액(급여대장)", 62_000_000, "KRW", "매월", date(2026, 11, 1), None, "확정", "급여대장 2026-10 기준"),
    ("4대보험·원천세 납부", "4대보험·원천세 총액", 23_000_000, "KRW", "매월", date(2026, 11, 1), None, "확정", "급여대장 2026-10 기준"),
    ("급여(실지급)", "신규 채용 1명(개발)", 4_000_000, "KRW", "매월", date(2027, 3, 1), None, "예정", "채용 확정 시 확정으로 변경"),
    ("임차료·관리비", "신사 사무실 월세·관리비", 12_000_000, "KRW", "매월", date(2026, 11, 1), date(2027, 10, 1), "확정", "임대차계약 만료 2027-10"),
    ("외주·용역비", "세무기장", 500_000, "KRW", "매월", date(2026, 11, 1), None, "확정", ""),
    ("외주·용역비", "외부 QA 용역", 9_000_000, "KRW", "1회", date(2027, 1, 1), None, "예정", "견적 수령"),
    ("서버·SaaS·라이선스", "AWS", 4_500_000, "KRW", "매월", date(2026, 11, 1), None, "확정", ""),
    ("서버·SaaS·라이선스", "개발툴 연간 라이선스", 6_000_000, "KRW", "연", date(2027, 3, 1), None, "확정", "매년 3월 갱신"),
    ("미국법인(GAME ROUND) 송금", "운영비 송금", 8_000, "USD", "매월", date(2026, 11, 1), None, "확정", "USD, 환율은 입력!B9"),
    ("기타 운영비", "복리후생·소모품", 3_000_000, "KRW", "매월", date(2026, 11, 1), None, "확정", ""),
    ("기타 운영비", "단체보험 연납", 2_400_000, "KRW", "연", date(2027, 4, 1), None, "확정", ""),
    ("법인세·부가세 납부", "부가세(분기)", 15_000_000, "KRW", "분기", date(2027, 1, 1), None, "확정", "1·4·7·10월"),
    ("법인세·부가세 납부", "법인세", 10_000_000, "KRW", "연", date(2027, 3, 1), None, "예정", "결산 후 확정"),
]
EXAMPLE_IN = [
    ("매출 입금", "퍼블리싱 정산", 45_000_000, "KRW", "매월", date(2026, 11, 1), None, "확정", "계약 A"),
    ("매출 입금", "라이선스 계약 B", 30_000_000, "KRW", "분기", date(2026, 12, 1), None, "예정", "협상 중"),
    ("정부지원금 입금", "OO지원사업 2차 지급", 60_000_000, "KRW", "1회", date(2027, 2, 1), None, "확정", "협약서"),
    ("정부지원금 입금", "OO지원사업 잔금", 20_000_000, "KRW", "1회", date(2027, 8, 1), None, "예정", "최종평가 후"),
    ("투자금 입금", "브릿지 라운드", 500_000_000, "KRW", "1회", date(2027, 4, 1), None, "예정", "텀시트 전"),
    ("이자·기타 입금", "예금이자", 500_000, "KRW", "매월", date(2026, 11, 1), None, "확정", ""),
]


def put(ws, addr, v, font=F, fill=None, fmt=None, align=None, border=None):
    c = ws[addr]
    c.value = v
    c.font = font
    if fill:
        c.fill = fill
    if fmt:
        c.number_format = fmt
    if align:
        c.alignment = Alignment(horizontal=align)
    if border:
        c.border = border
    return c


def set_title(ch, text):
    cp = CharacterProperties(sz=1100, b=True, solidFill=C_INK, latin=DFont(typeface=FONT), ea=DFont(typeface=FONT))
    para = Paragraph(pPr=ParagraphProperties(defRPr=cp), r=[RegularTextRun(rPr=cp, t=text)])
    ch.title = Title(tx=Text(rich=RichText(p=[para])), overlay=False)


def style_axes(ch):
    ch.x_axis.delete = False
    ch.y_axis.delete = False
    ch.x_axis.number_format = "yy-mm"
    ch.x_axis.majorTickMark = "none"
    ch.y_axis.number_format = "#,##0,,"
    ch.y_axis.majorGridlines.spPr = GraphicalProperties(ln=LineProperties(solidFill=C_GRID, w=6350))
    ch.x_axis.spPr = GraphicalProperties(ln=LineProperties(solidFill=C_GRID, w=6350))
    ch.y_axis.spPr = GraphicalProperties(ln=LineProperties(noFill=True))
    ch.legend = None


# ───────────────────────── 원자료 ─────────────────────────
def ledger_sheet(wb, name, cats, n_rows, rows_example, note):
    ws = wb.create_sheet(name)
    put(ws, "A1", name.replace("_", " "), FH)
    put(ws, "A2", note, FNOTE)
    hdr = ["구분", "항목", "금액(1회분)", "통화", "주기", "시작월", "종료월", "확정", "메모·근거"]
    widths = [24, 30, 16, 8, 8, 12, 12, 8, 36]
    for i, (h, w) in enumerate(zip(hdr, widths), start=1):
        put(ws, f"{L(i)}4", h, FB, GRAY, align="center", border=BOX)
        ws.column_dimensions[L(i)].width = w
    last = LG_FIRST + n_rows - 1
    for r in range(LG_FIRST, last + 1):
        for i in range(1, 10):
            c = put(ws, f"{L(i)}{r}", None, fill=YELLOW, border=BOX)
            if i == 3:
                c.number_format = NUM
            if i in (6, 7):
                c.number_format = "yyyy-mm"
    # 드롭다운
    def dv(formula, cols):
        d = DataValidation(type="list", formula1=formula, allow_blank=True)
        ws.add_data_validation(d)
        for col in cols:
            d.add(f"{col}{LG_FIRST}:{col}{last}")
    dv('"' + ",".join(cats) + '"', ["A"])
    dv('"KRW,USD"', ["D"])
    dv('"' + ",".join(PERIODS) + '"', ["E"])
    dv('"확정,예정"', ["H"])
    for r, row in enumerate(rows_example, start=LG_FIRST):
        for i, v in enumerate(row, start=1):
            ws[f"{L(i)}{r}"].value = v
    ws.freeze_panes = "A5"
    return last


# ───────────────────────── 전개표 ─────────────────────────
def expand_block(ws, title, src, src_last, start_row):
    """원자료 src의 각 줄을 18개월로 전개. 반환: (첫 데이터행, 마지막 데이터행, 다음 행)"""
    r = start_row
    put(ws, f"A{r}", title, FB, BLUE)
    for j in range(1, EX_LAST + 1):
        ws.cell(row=r, column=j).fill = BLUE
    r += 1
    hdr_row = r
    for i, h in enumerate(["구분", "항목", "확정", "금액(원)", "시작idx", "종료idx", "주기(월)"], start=1):
        put(ws, f"{L(i)}{r}", h, FB, GRAY, align="center")
    for j in range(MONTHS):
        col = L(EX_FIRST + j)
        put(ws, f"{col}{r}", f"=월별현금흐름!{L(CF_FIRST + j)}$4", FB, GRAY, "yy-mm", "center")
    r += 1
    idx_row = r
    put(ws, f"A{r}", "월 idx(연×12+월)", FNOTE)
    for j in range(MONTHS):
        col = L(EX_FIRST + j)
        put(ws, f"{col}{r}", f"=YEAR({col}{hdr_row})*12+MONTH({col}{hdr_row})", FNOTE, fmt="0")
    r += 1
    first = r
    for k, sr in enumerate(range(LG_FIRST, src_last + 1)):
        rr = first + k
        S = f"{src}!"
        put(ws, f"A{rr}", f'=IF({S}A{sr}="","",{S}A{sr})')
        put(ws, f"B{rr}", f'=IF({S}B{sr}="","",{S}B{sr})')
        put(ws, f"C{rr}", f'=IF({S}H{sr}="","",{S}H{sr})')
        put(ws, f"D{rr}", f'=IF({S}C{sr}="",0,IF({S}D{sr}="USD",{S}C{sr}*입력!$B$9,{S}C{sr}))', fmt="#,##0;-#,##0;")
        put(ws, f"E{rr}", f'=IF({S}F{sr}="","",YEAR({S}F{sr})*12+MONTH({S}F{sr}))', FNOTE, fmt="0")
        put(ws, f"F{rr}", f'=IF({S}G{sr}="",999999,YEAR({S}G{sr})*12+MONTH({S}G{sr}))', FNOTE, fmt='[=999999]"";0')
        put(ws, f"G{rr}", f'=IF({S}E{sr}="매월",1,IF({S}E{sr}="분기",3,IF({S}E{sr}="반기",6,'
                          f'IF({S}E{sr}="연",12,IF({S}E{sr}="1회",0,"")))))', FNOTE, fmt="0")
        for j in range(MONTHS):
            col = L(EX_FIRST + j)
            m = f"{col}${idx_row}"
            f = (f'=IF(OR($E{rr}="",$G{rr}=""),0,'
                 f'IF(AND({m}>=$E{rr},{m}<=$F{rr},IF($G{rr}=0,{m}=$E{rr},MOD({m}-$E{rr},MAX($G{rr},1))=0)),$D{rr},0))')
            put(ws, f"{col}{rr}", f, fmt="#,##0;[Red]-#,##0;-")
    last = first + (src_last - LG_FIRST)
    return first, last, last + 3


def build(path, example=False):
    wb = Workbook()

    # ───── 입력 ─────
    ws = wb.active
    ws.title = "입력"
    put(ws, "A1", "런웨이 시트 입력", FH)
    put(ws, "A2", "노란 칸만 입력. 비용·수입 목록은 원자료_유출 / 원자료_유입 시트에.", FNOTE)
    rows = [
        ("기준월(1일 날짜)", date(2026, 11, 1) if example else None, "yyyy-mm", "월별현금흐름 첫 열의 월"),
        ("기초현금(기준월 초 은행잔고 합계)", 1_200_000_000 if example else None, NUM, "법인 계좌 전체 합계. 외화는 기준월 초 환율로 환산"),
        ("비용 시나리오 증감률", 0.10, "0%", "요약의 '비용 +N%' 시나리오 비율"),
        ("런웨이 계산 기준 개월수", 6, "0", "기준월부터 N개월 평균 순소진으로 런웨이 계산"),
        ("USD 환율(원/달러)", 1_400 if example else None, "#,##0", "원자료에서 통화가 USD인 줄의 환산에 사용"),
        ("안전 잔액 기준(개월치 유출)", 3, "0", "안전 잔액선 = 향후 N개월 평균 월 유출 × 이 값. 기본 3개월치"),
    ]
    put(ws, "A4", "항목", FB, GRAY); put(ws, "B4", "값", FB, GRAY); put(ws, "C4", "메모", FB, GRAY)
    for i, (k, v, fmt, memo) in enumerate(rows, start=5):
        put(ws, f"A{i}", k)
        put(ws, f"B{i}", v, fill=YELLOW, fmt=fmt, align="right", border=BOX)
        put(ws, f"C{i}", memo, FNOTE)
    ws.column_dimensions["A"].width = 34
    ws.column_dimensions["B"].width = 18
    ws.column_dimensions["C"].width = 56

    # ───── 원자료 ─────
    note_out = "한 줄에 비용 하나. 급여는 급여대장 총액 한 줄(개인별 금액은 넣지 않음). 주기: 매월·분기·반기·연은 시작월부터 반복, 1회는 시작월에만. 종료월 비우면 계속."
    note_in = "한 줄에 수입 하나. 예정 건은 '예정'으로 두고 확정되면 바꿈. 투자금은 런웨이 계산에서 제외되고 잔액 추이에만 반영."
    if example:
        note_out = "[예시 데이터, 전부 가상 숫자] " + note_out
        note_in = "[예시 데이터, 전부 가상 숫자] " + note_in
    out_last = ledger_sheet(wb, "원자료_유출", CATS_OUT, N_OUT, EXAMPLE_OUT if example else [], note_out)
    in_last = ledger_sheet(wb, "원자료_유입", CATS_IN, N_IN, EXAMPLE_IN if example else [], note_in)
    out_sheet, in_sheet = wb["원자료_유출"], wb["원자료_유입"]

    # ───── 월별현금흐름 (먼저 틀을 만들어 전개표가 헤더를 참조) ─────
    cf = wb.create_sheet("월별현금흐름")
    put(cf, "A1", "월별 현금흐름 (18개월)", FH)
    put(cf, "A2", "전부 수식. 숫자를 고치려면 원자료 시트에서. B열은 해당 구분의 원자료 건수.", FNOTE)
    put(cf, "A3", "월 번호", FNOTE)
    put(cf, "A4", "항목", FB, GRAY)
    put(cf, "B4", "원자료 건수", FB, GRAY, align="center")
    for j in range(MONTHS):
        col = L(CF_FIRST + j)
        put(cf, f"{col}3", j + 1, FNOTE, fmt="0", align="center")
        f = "=입력!$B$5" if j == 0 else f"=EDATE({L(CF_FIRST + j - 1)}4,1)"
        put(cf, f"{col}4", f, FB, GRAY, "yyyy-mm", "center")
        cf.column_dimensions[col].width = 13
    cf.column_dimensions["A"].width = 34
    cf.column_dimensions["B"].width = 11

    # ───── 전개표 ─────
    ex = wb.create_sheet("전개표")
    put(ex, "A1", "전개표 (자동)", FH)
    put(ex, "A2", "원자료 각 줄을 월별로 펼친 검산용 시트. 손대지 않음. 어느 항목이 어느 달에 얼마 잡혔는지 확인할 때 봄.", FNOTE)
    o_first, o_last, nxt = expand_block(ex, "유출", "원자료_유출", out_last, 4)
    i_first, i_last, _ = expand_block(ex, "유입", "원자료_유입", in_last, nxt)
    for col, w in zip("ABCDEFG", [24, 28, 6, 14, 9, 9, 8]):
        ex.column_dimensions[col].width = w
    for j in range(MONTHS):
        ex.column_dimensions[L(EX_FIRST + j)].width = 12
    ex.freeze_panes = "H7"

    def sumifs(block_first, block_last, excol, cat_cell=None, cat_lit=None, status=None):
        rng = f"전개표!{excol}${block_first}:{excol}${block_last}"
        crit = []
        if cat_cell:
            crit.append(f"전개표!$A${block_first}:$A${block_last},{cat_cell}")
        if cat_lit:
            crit.append(f'전개표!$A${block_first}:$A${block_last},"{cat_lit}"')
        if status:
            crit.append(f'전개표!$C${block_first}:$C${block_last},"{status}"')
        return f"=SUMIFS({rng},{','.join(crit)})"

    # ───── 월별현금흐름 본문 ─────
    r = 5
    put(cf, f"A{r}", "기초현금", FB)
    row_begin = r

    def block(title, cats, src_name, b_first, b_last, start_row):
        rr = start_row
        put(cf, f"A{rr}", title, FB, BLUE)
        for j in range(1, CF_LAST + 1):
            cf.cell(row=rr, column=j).fill = BLUE
        rr += 1
        first_item = rr
        for cat in cats:
            put(cf, f"A{rr}", cat)
            put(cf, f"B{rr}", f"=COUNTIF({src_name}!$A${LG_FIRST}:$A${LG_FIRST + 60},$A{rr})", FNOTE, fmt="0", align="center")
            for j in range(MONTHS):
                col = L(CF_FIRST + j)
                excol = L(EX_FIRST + j)
                put(cf, f"{col}{rr}", sumifs(b_first, b_last, excol, cat_cell=f"$A{rr}"), fmt=NUM)
            rr += 1
        last_item = rr - 1
        put(cf, f"A{rr}", f"{title} 합계", FB)
        for j in range(MONTHS):
            col = L(CF_FIRST + j)
            put(cf, f"{col}{rr}", f"=SUM({col}{first_item}:{col}{last_item})", FB, fmt=NUM)
        return rr, rr + 2

    in_total, r = block("현금 유입", CATS_IN, "원자료_유입", i_first, i_last, row_begin + 2)
    out_total, r = block("현금 유출", CATS_OUT, "원자료_유출", o_first, o_last, r)
    inv_row = in_total - 2  # 투자금 입금 행

    put(cf, f"A{r}", "순현금흐름", FB)
    net_row = r
    for j in range(MONTHS):
        col = L(CF_FIRST + j)
        put(cf, f"{col}{r}", f"={col}{in_total}-{col}{out_total}", FB, fmt=NUM)
    r += 1
    put(cf, f"A{r}", "순현금흐름(투자금 제외, 런웨이 계산용)", FB)
    op_net_row = r
    for j in range(MONTHS):
        col = L(CF_FIRST + j)
        put(cf, f"{col}{r}", f"={col}{net_row}-{col}{inv_row}", FB, fmt=NUM)
    r += 1
    put(cf, f"A{r}", "확정 유입 합계", F)
    conf_in_row = r
    for j in range(MONTHS):
        col = L(CF_FIRST + j)
        put(cf, f"{col}{r}", sumifs(i_first, i_last, L(EX_FIRST + j), status="확정"), fmt=NUM)
    r += 1
    put(cf, f"A{r}", "확정 투자금 입금", F)
    conf_inv_row = r
    for j in range(MONTHS):
        col = L(CF_FIRST + j)
        put(cf, f"{col}{r}", sumifs(i_first, i_last, L(EX_FIRST + j), cat_lit="투자금 입금", status="확정"), fmt=NUM)
    r += 1
    put(cf, f"A{r}", "보수적 순현금흐름(확정 유입만, 투자금 제외)", FB)
    cons_row = r
    for j in range(MONTHS):
        col = L(CF_FIRST + j)
        put(cf, f"{col}{r}", f"={col}{conf_in_row}-{col}{conf_inv_row}-{col}{out_total}", FB, fmt=NUM)
    r += 1
    put(cf, f"A{r}", "기말현금", FB)
    end_row = r
    for j in range(MONTHS):
        col = L(CF_FIRST + j)
        put(cf, f"{col}{r}", f"={col}{row_begin}+{col}{net_row}", FB, GRAY, NUM)
    for j in range(MONTHS):  # 기초현금 (기말 행 확정 후)
        col = L(CF_FIRST + j)
        f = "=입력!$B$6" if j == 0 else f"={L(CF_FIRST + j - 1)}{end_row}"
        put(cf, f"{col}{row_begin}", f, fmt=NUM)
    r += 1
    put(cf, f"A{r}", "현금 소진 플래그(기말<0)", FNOTE)
    flag_row = r
    for j in range(MONTHS):
        col = L(CF_FIRST + j)
        put(cf, f"{col}{r}", f"=IF({col}{end_row}<0,1,0)", FNOTE, fmt="0")
    r += 1
    put(cf, f"A{r}", "런웨이 계산 포함(1=포함)", FNOTE)
    incl_row = r
    for j in range(MONTHS):
        col = L(CF_FIRST + j)
        put(cf, f"{col}{r}", f"=IF({col}$3<=입력!$B$8,1,0)", FNOTE, fmt="0")
    r += 1
    put(cf, f"A{r}", "안전 잔액선(N개월치 평균 유출)", FB)
    safe_row = r
    for j in range(MONTHS):
        col = L(CF_FIRST + j)
        put(cf, f"{col}{r}", "=요약!$B$16", FB, fmt=NUM)
    r += 1
    put(cf, f"A{r}", "안전선 미달 플래그(기말<안전선)", FNOTE)
    safe_flag_row = r
    for j in range(MONTHS):
        col = L(CF_FIRST + j)
        put(cf, f"{col}{r}", f"=IF({col}{end_row}<{col}{safe_row},1,0)", FNOTE, fmt="0")
    cf.freeze_panes = "C5"

    # ───── 요약 ─────
    sm = wb.create_sheet("요약")
    C0, C1 = L(CF_FIRST), L(CF_LAST)
    rng = lambda row: f"월별현금흐름!{C0}{row}:{C1}{row}"
    n = "입력!$B$8"
    put(sm, "A1", "런웨이 요약", FH)
    put(sm, "A2", "전부 수식. 입력·원자료 시트만 고치면 자동 갱신.", FNOTE)
    put(sm, "A4", "항목", FB, GRAY); put(sm, "B4", "값", FB, GRAY); put(sm, "C4", "계산", FB, GRAY)
    items = [
        ("기초현금", "=입력!B6", NUM, "입력!B6"),
        ("향후 N개월 평균 순소진(투자금 제외)", f"=SUMPRODUCT({rng(incl_row)},{rng(op_net_row)})/{n}", NUM,
         "기준월부터 N개월(입력!B8) 평균. 확정+예정 전부, 투자금만 제외"),
        ("런웨이(개월)", '=IF(B6<0,B5/ABS(B6),"영업 흑자(소진 없음)")', "0.0", "기초현금 ÷ |평균 순소진|"),
        ("현금 소진 예상월", f'=IFERROR(INDEX({rng(4)},MATCH(1,{rng(flag_row)},0)),"18개월 내 없음")', "yyyy-mm",
         "기말현금이 처음 0 미만이 되는 월(확정+예정 전부 반영)"),
        ("18개월 내 최저 기말현금", f"=MIN({rng(end_row)})", NUM, ""),
        ("최저 기말현금 월", f"=INDEX({rng(4)},MATCH(B9,{rng(end_row)},0))", "yyyy-mm", ""),
        ("보수적 평균 순소진(확정 유입만)", f"=SUMPRODUCT({rng(incl_row)},{rng(cons_row)})/{n}", NUM,
         "유입은 '확정'만, 유출은 확정+예정 전부, 투자금 제외"),
        ("보수적 런웨이(개월)", '=IF(B11<0,B5/ABS(B11),"영업 흑자(소진 없음)")', "0.0", "예정 매출·지원금이 안 들어올 때"),
        ("18개월 유입 합계", f"=SUM({rng(in_total)})", NUM, ""),
        ("18개월 유출 합계", f"=SUM({rng(out_total)})", NUM, ""),
        ("향후 N개월 평균 월 유출", f"=SUMPRODUCT({rng(incl_row)},{rng(out_total)})/{n}", NUM, "기준월부터 N개월(입력!B8) 유출 평균"),
        ("안전 잔액선", "=B15*입력!B10", NUM, "평균 월 유출 × 안전 기준 개월수(입력!B10)"),
        ("안전선 미달 예상월", f'=IFERROR(INDEX({rng(4)},MATCH(1,{rng(safe_flag_row)},0)),"18개월 내 없음")', "yyyy-mm",
         "기말현금이 처음 안전 잔액선 아래로 내려가는 월"),
    ]
    for i, (k, f, fmt, memo) in enumerate(items, start=5):
        put(sm, f"A{i}", k)
        put(sm, f"B{i}", f, FB, fmt=fmt, align="right")
        put(sm, f"C{i}", memo, FNOTE)
    put(sm, "A19", "시나리오: 비용 증감", FB, BLUE)
    for j in range(1, 4):
        sm.cell(row=19, column=j).fill = BLUE
    put(sm, "A20", "비용 증감률"); put(sm, "B20", "=입력!B7", FB, fmt="+0%;-0%", align="right"); put(sm, "C20", "입력!B7", FNOTE)
    put(sm, "A21", "시나리오 평균 순소진")
    put(sm, "B21", f"=(SUMPRODUCT({rng(incl_row)},{rng(in_total)})-SUMPRODUCT({rng(incl_row)},{rng(inv_row)})"
                   f"-SUMPRODUCT({rng(incl_row)},{rng(out_total)})*(1+B20))/{n}", FB, fmt=NUM, align="right")
    put(sm, "C21", "(유입 − 투자금) 평균 − 유출 평균 × (1+증감률)", FNOTE)
    put(sm, "A22", "시나리오 런웨이(개월)")
    put(sm, "B22", '=IF(B21<0,B5/ABS(B21),"영업 흑자(소진 없음)")', FB, fmt="0.0", align="right")
    put(sm, "A23", "런웨이 변화(개월)")
    put(sm, "B23", '=IF(AND(ISNUMBER(B7),ISNUMBER(B22)),B22-B7,"-")', FB, fmt="+0.0;-0.0", align="right")
    sm.column_dimensions["A"].width = 36
    sm.column_dimensions["B"].width = 20
    sm.column_dimensions["C"].width = 50

    # ───── 시나리오 비교 ─────
    sc = wb.create_sheet("시나리오")
    put(sc, "A1", "시나리오 비교 (3열)", FH)
    put(sc, "A2", "노란 칸만 입력. 각 열은 원자료 전체를 기준으로 증감률·예정 유입 포함 여부·추가 비용·기초현금 조정을 적용한 결과.", FNOTE)
    SCOLS = ["B", "C", "D"]
    put(sc, "A4", "입력", FB, GRAY)
    for k, col in enumerate(SCOLS, start=1):
        put(sc, f"{col}4", f"시나리오 {k}", FB, GRAY, align="center")
    params = [
        ("시나리오 이름", None, ["기준(현재 원자료)", "보수: 예정 수입 제외·비용 +10%", "채용 2명 추가(2027-03~)"]),
        ("유출 증감률", "0%", [0, 0.10, 0]),
        ("예정 유입 (포함/제외)", None, ["포함", "제외", "포함"]),
        ("추가 월 비용(원)", NUM, [0, 0, 9_000_000]),
        ("추가 비용 시작월", "yyyy-mm", [None, None, date(2027, 3, 1)]),
        ("기초현금 조정(원)", NUM, [0, 0, 0]),
    ]
    for i, (label, fmt, vals) in enumerate(params, start=5):
        put(sc, f"A{i}", label)
        for col, v in zip(SCOLS, vals):
            if not example and i in (5, 7):
                v = v  # 이름·포함여부는 템플릿에도 기본값 유지
            elif not example:
                v = 0 if i in (6, 8, 10) else None
            put(sc, f"{col}{i}", v, fill=YELLOW, fmt=fmt, align="right", border=BOX)
    dvs = DataValidation(type="list", formula1='"포함,제외"', allow_blank=False)
    sc.add_data_validation(dvs); dvs.add("B7:D7")

    # 월별 그리드 (23행부터). 월 헤더는 월별현금흐름!4행 참조, 월 번호는 22행
    G0 = 23
    put(sc, f"A{G0 - 1}", "월 번호", FNOTE)
    put(sc, f"A{G0}", "월", FB, GRAY)
    for j in range(MONTHS):
        col = L(CF_FIRST + j)
        put(sc, f"{col}{G0 - 1}", j + 1, FNOTE, fmt="0", align="center")
        put(sc, f"{col}{G0}", f"=월별현금흐름!{col}$4", FB, GRAY, "yy-mm", "center")
        sc.column_dimensions[col].width = 15
    CFS = "월별현금흐름!"
    blocks = {}
    r = G0 + 1
    for k, pc in enumerate(SCOLS, start=1):
        put(sc, f"A{r}", f'="시나리오 {k}: "&${pc}$5', FB, BLUE)
        for j in range(1, CF_LAST + 1):
            sc.cell(row=r, column=j).fill = BLUE
        rows = {}
        names = ["유입(투자금 제외)", "투자금 입금", "유출", "순현금흐름(투자금 제외)", "기말현금", "소진 플래그", "안전선 미달 플래그"]
        for nm in names:
            r += 1
            rows[nm] = r
            put(sc, f"A{r}", nm, FB if nm in ("순현금흐름(투자금 제외)", "기말현금") else F)
        for j in range(MONTHS):
            col = L(CF_FIRST + j)
            f_in = (f'=IF(${pc}$7="포함",{CFS}{col}{in_total}-{CFS}{col}{inv_row},'
                    f'{CFS}{col}{conf_in_row}-{CFS}{col}{conf_inv_row})')
            f_inv = f'=IF(${pc}$7="포함",{CFS}{col}{inv_row},{CFS}{col}{conf_inv_row})'
            f_out = (f'={CFS}{col}{out_total}*(1+${pc}$6)'
                     f'+IF(${pc}$9="",${pc}$8,IF(YEAR({col}${G0})*12+MONTH({col}${G0})>=YEAR(${pc}$9)*12+MONTH(${pc}$9),${pc}$8,0))')
            f_net = f'={col}{rows["유입(투자금 제외)"]}-{col}{rows["유출"]}'
            prev = f"입력!$B$6+${pc}$10" if j == 0 else f"{L(CF_FIRST + j - 1)}{rows['기말현금']}"
            f_end = f'={prev}+{col}{rows["순현금흐름(투자금 제외)"]}+{col}{rows["투자금 입금"]}'
            put(sc, f"{col}{rows['유입(투자금 제외)']}", f_in, fmt=NUM)
            put(sc, f"{col}{rows['투자금 입금']}", f_inv, fmt=NUM)
            put(sc, f"{col}{rows['유출']}", f_out, fmt=NUM)
            put(sc, f"{col}{rows['순현금흐름(투자금 제외)']}", f_net, FB, fmt=NUM)
            put(sc, f"{col}{rows['기말현금']}", f_end, FB, GRAY, NUM)
            put(sc, f"{col}{rows['소진 플래그']}", f'=IF({col}{rows["기말현금"]}<0,1,0)', FNOTE, fmt="0")
            put(sc, f"{col}{rows['안전선 미달 플래그']}", f'=IF({col}{rows["기말현금"]}<요약!$B$16,1,0)', FNOTE, fmt="0")
        blocks[pc] = rows
        r += 2

    # 결과 표 (12~20행)
    put(sc, "A12", "결과", FB, GRAY)
    for k, col in enumerate(SCOLS, start=1):
        put(sc, f"{col}12", f"={col}5", FB, GRAY, align="center")
    mon_rng = f"{L(CF_FIRST)}{G0}:{L(CF_LAST)}{G0}"
    incl = f"{CFS}{L(CF_FIRST)}{incl_row}:{L(CF_LAST)}{incl_row}"
    res = [
        ("향후 N개월 평균 순소진", NUM, lambda pc, rw: f"=SUMPRODUCT({incl},{L(CF_FIRST)}{rw['순현금흐름(투자금 제외)']}:{L(CF_LAST)}{rw['순현금흐름(투자금 제외)']})/입력!$B$8"),
        ("런웨이(개월)", "0.0", lambda pc, rw: f'=IF({pc}13<0,(입력!$B$6+{pc}10)/ABS({pc}13),"영업 흑자")'),
        ("기준(시나리오 1) 대비 차이", "+0.0;-0.0", lambda pc, rw: f'=IF(AND(ISNUMBER({pc}14),ISNUMBER($B14)),{pc}14-$B14,"-")'),
        ("현금 소진 예상월", "yyyy-mm", lambda pc, rw: f'=IFERROR(INDEX({mon_rng},MATCH(1,{L(CF_FIRST)}{rw["소진 플래그"]}:{L(CF_LAST)}{rw["소진 플래그"]},0)),"18개월 내 없음")'),
        ("18개월 후 기말현금", NUM, lambda pc, rw: f"={L(CF_LAST)}{rw['기말현금']}"),
        ("최저 기말현금", NUM, lambda pc, rw: f"=MIN({L(CF_FIRST)}{rw['기말현금']}:{L(CF_LAST)}{rw['기말현금']})"),
        ("최저 기말현금 월", "yyyy-mm", lambda pc, rw: f"=INDEX({mon_rng},MATCH({pc}18,{L(CF_FIRST)}{rw['기말현금']}:{L(CF_LAST)}{rw['기말현금']},0))"),
        ("안전선 미달 예상월", "yyyy-mm", lambda pc, rw: f'=IFERROR(INDEX({mon_rng},MATCH(1,{L(CF_FIRST)}{rw["안전선 미달 플래그"]}:{L(CF_LAST)}{rw["안전선 미달 플래그"]},0)),"18개월 내 없음")'),
    ]
    for i, (label, fmt, fn) in enumerate(res, start=13):
        put(sc, f"A{i}", label, FB if "런웨이" in label else F)
        for pc in SCOLS:
            put(sc, f"{pc}{i}", fn(pc, blocks[pc]), FB if "런웨이" in label else F, fmt=fmt, align="right")
    sc.column_dimensions["A"].width = 30
    for col in SCOLS:
        sc.column_dimensions[col].width = 26
    sc.freeze_panes = "B5"
    sc_rows = dict(params=(5, 10), results=(13, 20), grid0=G0, blocks={pc: dict(v) for pc, v in blocks.items()})

    # ───── 보고 ─────
    rp = wb.create_sheet("보고")
    rp.sheet_view.showGridLines = False
    for col, w in zip("ABCDEFGHI", [2, 24, 3, 24, 3, 24, 3, 24, 2]):
        rp.column_dimensions[col].width = w
    rp.row_dimensions[2].height = 30
    put(rp, "B2", "현금 런웨이 보고", Font(name=FONT, size=18, bold=True, color=C_INK))
    put(rp, "B3", '="기준월 "&YEAR(입력!B5)&"-"&RIGHT("0"&MONTH(입력!B5),2)&" · 기초현금 "&TEXT(입력!B6/1000000,"#,##0")'
                  '&"백만원 · 입금·출금 기준"', Font(name=FONT, size=10, color=C_MUTED))
    mon = lambda ref: f'IF(ISNUMBER({ref}),YEAR({ref})&"-"&RIGHT("0"&MONTH({ref}),2),{ref})'
    mo = lambda ref: f'IF(ISNUMBER({ref}),TEXT({ref},"0.0")&"개월",{ref})'
    tiles = [
        ("B", "런웨이 (확정+예정)", f"={mo('요약!B7')}", None),
        ("D", "보수적 런웨이 (확정 유입만)", f"={mo('요약!B12')}", None),
        ("F", "현금 소진 예상월", f"={mon('요약!B8')}", None),
        ("H", "18개월 내 최저 잔액 (백만원)", "=요약!B9", "#,##0,,;[Red]-#,##0,,"),
    ]
    rp.row_dimensions[5].height = 18
    rp.row_dimensions[6].height = 34
    for col, label, f, fmt in tiles:
        put(rp, f"{col}5", label, Font(name=FONT, size=10, color=C_INK2))
        put(rp, f"{col}6", f, Font(name=FONT, size=20, bold=True, color=C_INK), fmt=fmt, align="left",
            border=Border(bottom=Side(style="thin", color=C_GRID)))
    put(rp, "B7", '="비용 "&IF(요약!B20>=0,"+","")&TEXT(요약!B20,"0%")&" 시 런웨이 "&' + mo("요약!B22") +
                  '&IF(ISNUMBER(요약!B23)," ("&IF(요약!B23>=0,"+","")&TEXT(요약!B23,"0.0")&"개월)","")'
                  '&" · 안전 잔액선 "&TEXT(요약!B16/1000000,"#,##0")&"백만원("&입력!B10&"개월치 유출), 미달 예상 "&' + mon("요약!B17"),
        Font(name=FONT, size=10, color=C_INK2))
    put(rp, "B9", "보고 문구", Font(name=FONT, size=10, color=C_INK2))
    put(rp, "B10", '="런웨이 "&' + mo("요약!B7") + '&", 예정 수입 제외 시 "&' + mo("요약!B12") +
                   '&". 현금 소진 예상 "&' + mon("요약!B8") +
                   '&". 월평균 순소진 "&TEXT(ABS(요약!B6)/1000000,"#,##0")&"백만원."',
        Font(name=FONT, size=11, color=C_INK))
    rp.merge_cells("B10:H10")
    rp["B10"].alignment = Alignment(wrap_text=True, vertical="top")
    rp.row_dimensions[10].height = 34

    cats = Reference(cf, min_col=CF_FIRST, max_col=CF_LAST, min_row=4)
    ln = LineChart()
    set_title(ln, "기말현금 추이 (백만원)")
    ln.style = 2
    ln.height, ln.width = 7.2, 19.5
    ln.add_data(Reference(cf, min_col=CF_FIRST, max_col=CF_LAST, min_row=end_row), from_rows=True, titles_from_data=False)
    ln.set_categories(cats)
    s1 = ln.series[0]
    s1.graphicalProperties.line.solidFill = C_BLUE
    s1.graphicalProperties.line.width = 28575
    s1.marker.symbol = "circle"
    s1.marker.size = 5
    s1.marker.graphicalProperties = GraphicalProperties(solidFill=C_BLUE)
    s1.marker.graphicalProperties.line.solidFill = "FFFFFF"
    s1.smooth = False
    # 안전 잔액선: 주황 점선, 마커 없음
    ln.add_data(Reference(cf, min_col=CF_FIRST, max_col=CF_LAST, min_row=safe_row), from_rows=True, titles_from_data=False)
    s2 = ln.series[1]
    s2.graphicalProperties.line.solidFill = C_ORANGE
    s2.graphicalProperties.line.width = 19050
    s2.graphicalProperties.line.dashStyle = "dash"
    s2.marker.symbol = "none"
    s2.smooth = False
    ln.series[0].tx = SeriesLabel(v="기말현금")
    ln.series[1].tx = SeriesLabel(v="안전 잔액선")
    style_axes(ln)
    ln.legend = Legend()
    ln.legend.position = "t"
    rp.add_chart(ln, "B12")

    bar = BarChart()
    bar.type = "col"
    bar.grouping = "clustered"
    set_title(bar, "월별 현금 유입·유출 (백만원)")
    bar.style = 2
    bar.height, bar.width = 7.2, 19.5
    bar.gapWidth = 80
    bar.overlap = -10
    bar.add_data(Reference(cf, min_col=CF_FIRST, max_col=CF_LAST, min_row=in_total), from_rows=True, titles_from_data=False)
    bar.add_data(Reference(cf, min_col=CF_FIRST, max_col=CF_LAST, min_row=out_total), from_rows=True, titles_from_data=False)
    bar.set_categories(cats)
    bar.series[0].tx = SeriesLabel(v="유입")
    bar.series[1].tx = SeriesLabel(v="유출")
    for sr, color in zip(bar.series, (C_BLUE, C_ORANGE)):
        sr.graphicalProperties.solidFill = color
        sr.graphicalProperties.line.noFill = True
    style_axes(bar)
    bar.legend = Legend()
    bar.legend.position = "t"
    rp.add_chart(bar, "B27")
    put(rp, "B42", "읽는 법: 런웨이는 투자금을 뺀 영업 순소진 기준. 보수적 런웨이는 '예정' 표시된 매출·지원금도 뺀 값. "
                   "기말현금 추이에는 예정 투자금이 반영됨. 안전 잔액선은 향후 N개월 평균 월 유출 × 안전 기준 개월수.",
        Font(name=FONT, size=9, color=C_MUTED))
    rp.merge_cells("B42:H42")
    rp["B42"].alignment = Alignment(wrap_text=True, vertical="top")
    rp.row_dimensions[42].height = 26

    # 2쪽: 시나리오 비교표
    T0 = 45
    rp.row_dimensions[T0].height = 30
    put(rp, f"B{T0}", "시나리오 비교", Font(name=FONT, size=18, bold=True, color=C_INK))
    put(rp, f"B{T0 + 1}", "각 열은 원자료 전체에 증감률·예정 유입 포함 여부·추가 비용을 적용한 결과. 상세는 시나리오 시트.",
        Font(name=FONT, size=10, color=C_MUTED))
    hdr_r = T0 + 3
    put(rp, f"B{hdr_r}", "항목", Font(name=FONT, size=10, bold=True, color=C_INK2),
        border=Border(bottom=Side(style="thin", color=C_INK2)))
    for col, pc in zip("DFH", SCOLS):
        put(rp, f"{col}{hdr_r}", f"=시나리오!{pc}5", Font(name=FONT, size=10, bold=True, color=C_INK2), align="right",
            border=Border(bottom=Side(style="thin", color=C_INK2)))
        rp[f"{col}{hdr_r}"].alignment = Alignment(horizontal="right", wrap_text=True, vertical="bottom")
    rp.row_dimensions[hdr_r].height = 30
    lines = [
        ("유출 증감률", lambda pc: f"=시나리오!{pc}6", "+0%;-0%;0%"),
        ("예정 유입", lambda pc: f"=시나리오!{pc}7", None),
        ("추가 월 비용 (백만원)", lambda pc: f"=시나리오!{pc}8", "#,##0,,;-#,##0,,;-"),
        ("런웨이 (개월)", lambda pc: f"=시나리오!{pc}14", "0.0"),
        ("기준 대비 (개월)", lambda pc: f"=시나리오!{pc}15", "+0.0;-0.0;0.0"),
        ("현금 소진 예상월", lambda pc: f"=시나리오!{pc}16", "yyyy-mm"),
        ("18개월 후 기말현금 (백만원)", lambda pc: f"=시나리오!{pc}17", "#,##0,,;[Red]-#,##0,,"),
        ("최저 기말현금 (백만원)", lambda pc: f"=시나리오!{pc}18", "#,##0,,;[Red]-#,##0,,"),
        ("안전선 미달 예상월", lambda pc: f"=시나리오!{pc}20", "yyyy-mm"),
    ]
    for i, (label, fn, fmt) in enumerate(lines, start=hdr_r + 1):
        bold = "런웨이" in label
        put(rp, f"B{i}", label, Font(name=FONT, size=11 if bold else 10, bold=bold, color=C_INK),
            border=Border(bottom=Side(style="thin", color=C_GRID)))
        for col, pc in zip("DFH", SCOLS):
            put(rp, f"{col}{i}", fn(pc), Font(name=FONT, size=14 if bold else 10, bold=bold, color=C_INK), fmt=fmt, align="right",
                border=Border(bottom=Side(style="thin", color=C_GRID)))
        rp.row_dimensions[i].height = 24 if bold else 18
    put(rp, f"B{hdr_r + len(lines) + 2}", "읽는 법: 런웨이는 투자금 제외 영업 순소진 기준. '예정 유입 제외'는 예정 표시된 매출·지원금·투자금을 모두 뺀 값.",
        Font(name=FONT, size=9, color=C_MUTED))
    rp.row_breaks.append(Break(id=43))
    rp.print_area = f"A1:I{hdr_r + len(lines) + 3}"
    rp.page_setup.orientation = "landscape"
    rp.page_setup.paperSize = rp.PAPERSIZE_A4
    rp.sheet_properties.pageSetUpPr.fitToPage = True
    rp.sheet_properties.pageSetUpPr.fitToPage = False
    rp.page_setup.scale = 62          # 1쪽(보고)·2쪽(시나리오 비교) 고정 배율, 43행 뒤 수동 페이지 나눔
    rp.print_options.horizontalCentered = True

    # 나머지 시트 인쇄 설정: 가로, 폭 1페이지
    for sh in (ws, out_sheet, in_sheet, ex, cf, sm, sc):
        sh.page_setup.orientation = "landscape"
        sh.page_setup.paperSize = sh.PAPERSIZE_A4
        sh.sheet_properties.pageSetUpPr.fitToPage = True
        sh.page_setup.fitToWidth = 1
        sh.page_setup.fitToHeight = 0
    wb.move_sheet("보고", offset=-7)
    wb.save(path)
    layout = dict(row_begin=row_begin, in_total=in_total, out_total=out_total, inv_row=inv_row, net_row=net_row,
                  op_net_row=op_net_row, cons_row=cons_row, end_row=end_row, flag_row=flag_row, incl_row=incl_row,
                  safe_row=safe_row, safe_flag_row=safe_flag_row, conf_in_row=conf_in_row, conf_inv_row=conf_inv_row,
                  scenario=sc_rows,
                  ex_out=(o_first, o_last), ex_in=(i_first, i_last))
    print("saved", path, layout)
    return layout


if __name__ == "__main__":
    if "--example" in sys.argv:
        build("runway_example.xlsx", example=True)
    else:
        build("runway_template.xlsx", example=False)
