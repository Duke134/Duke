"""개발·테스트용 가짜 세금계산서 PDF 6건과 홈택스 목록 엑셀을 samples/ 에 만든다. 실제 회사 정보 없음."""
from pathlib import Path
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.cidfonts import UnicodeCIDFont
from openpyxl import Workbook

OUR = ("123-45-67890", "주식회사 샘플게임즈")
PARTNERS = [("201-81-11111", "(주)판교오피스"), ("305-86-22222", "강남클라우드 주식회사"), ("110-12-33333", "유한회사 디자인랩")]
ROWS = [  # 승인번호, 작성일자, 공급자, 공급받는자, 공급가액, 세액
    ("20261002-41000001-00000001", "2026-10-02", PARTNERS[0], OUR, 2500000, 250000),
    ("20261010-41000002-00000002", "2026-10-10", PARTNERS[1], OUR, 1100000, 110000),
    ("20261015-41000003-00000003", "2026-10-15", OUR, PARTNERS[2], 8000000, 800000),
    ("20261020-41000004-00000004", "2026-10-20", PARTNERS[2], OUR, 450000, 45000),
    ("20260928-41000005-00000005", "2026-09-28", OUR, PARTNERS[0], 3000000, 300000),
]
base = Path(__file__).with_name("samples"); inbox = base / "inbox"; inbox.mkdir(parents=True, exist_ok=True)
pdfmetrics.registerFont(UnicodeCIDFont("HYGothic-Medium"))


def pdf(path, no, date, sup, buy, amt, tax):
    c = canvas.Canvas(str(path)); c.setFont("HYGothic-Medium", 11); y = 800
    for line in [f"전자세금계산서  승인번호 {no}", "[공급자]", f"등록번호 {sup[0]}", f"상호(법인명) {sup[1]}  성명 홍길동",
                 "[공급받는자]", f"등록번호 {buy[0]}", f"상호(법인명) {buy[1]}  성명 김철수",
                 f"작성일자 {date}", f"공급가액 {amt:,}  세액 {tax:,}  합계금액 {amt+tax:,}"]:
        c.drawString(50, y, line); y -= 22
    c.save()


for i, (no, date, sup, buy, amt, tax) in enumerate(ROWS, 1):
    pdf(inbox / f"download ({i}).pdf", no, date, sup, buy, amt, tax)
# 목록에 없고 승인번호도 없는 PDF 1건 -> 확인필요로 가야 함
c = canvas.Canvas(str(inbox / "scan_receipt.pdf")); c.setFont("HYGothic-Medium", 11); c.drawString(50, 800, "영수증 (스캔본, 글자 없음 가정)"); c.save()

wb = Workbook(); ws = wb.active
ws.append(["구분", "승인번호", "작성일자", "공급자 상호", "공급받는자 상호", "공급가액", "세액"])
for no, date, sup, buy, amt, tax in ROWS[:4]:  # 마지막 1건은 목록에 없음 -> PDF 본문으로 읽어야 함
    ws.append(["매출" if sup == OUR else "매입", no, date, sup[1], buy[1], amt, tax])
wb.save(base / "홈택스목록.xlsx")
print("샘플 생성 완료:", inbox)
