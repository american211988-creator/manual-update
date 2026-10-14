"""시험용 더미데이터 2차 묶음(data/test_pack2/): 형식별 신구 문서 1쌍씩.

  PDF   보험계약 유지관리 업무매뉴얼  v1.0 → v2.0  (흐름도·차트 이미지)
  Word  모집질서 준수 내규            v2.0 → v2.1  (등급 매트릭스 이미지 + 조치 기준표)
  Excel 보험료 수납 사무처리기준      v1.0 → v1.1  (조항|제목|내용 표 + 차트 이미지 + 개정이력 시트)
  Text  심사위원회 운영규정           v3.0 → v3.1

정답표는 변경 목록에서 자동 계산한다. 조문 제목만 바뀐 경우도 '개정'으로 넣었다.
사용법: PYTHONPATH=<python-docx 설치 경로> python3 tools/make_test_pack2.py   (LibreOffice 필요)
모든 회사명·금액·인명은 가상입니다.
"""
import csv
import shutil
import sys
import tempfile
from collections import Counter
from pathlib import Path

import openpyxl
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

sys.path.insert(0, str(Path(__file__).resolve().parent))
from make_test_pack import ORG, art, bar_chart, flowchart, key, make_docx, matrix, png, to_pdf  # noqa: E402

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "data" / "test_pack2"

# ------------------------------------------------------------------ 문서 내용
# mod 값: (새 문단 목록 또는 None=그대로, 비고, 새 제목 또는 None)
PDF_OLD = [
    art(1, "목적", "이 매뉴얼은 보험계약 유지관리 업무의 처리 절차를 정한다."),
    art(2, "적용범위", "이 매뉴얼은 개인보험 유효계약의 유지관리 업무에 적용한다."),
    art(3, "업무 흐름", "유지관리는 접수, 확인, 처리, 통지 순으로 진행하며 흐름은 아래 그림과 같다."),
    art(4, "주소 변경", "계약자 주소 변경은 신분 확인 후 즉시 반영한다."),
    art(5, "수익자 변경", "① 수익자 변경은 계약자 서면 동의로 처리한다.", "② 피보험자와 다른 경우 피보험자 동의를 받는다."),
    art(6, "납입주기 변경", "납입주기 변경은 월납과 연납 사이에서만 할 수 있다."),
    art(7, "감액", "보험가입금액 감액은 최초 가입금액의 50%까지 할 수 있다."),
    art(8, "실효 안내", "보험료 미납 시 납입최고기간 14일 전에 안내한다."),
    art(9, "부활 신청", "실효일부터 2년 이내에 부활을 신청할 수 있다."),
    art(10, "약관대출", "약관대출 한도는 해지환급금의 80%로 한다."),
    art(11, "중도인출", "중도인출은 연 4회까지 할 수 있다."),
    art(12, "자동이체", "자동이체 출금일은 매월 10일과 25일 중 선택한다."),
    art(13, "팩스 신청", "팩스로 접수된 변경 신청은 원본 확인 후 처리한다."),
    art(14, "처리 기한", "변경 신청은 접수일부터 3영업일 이내에 처리한다."),
    art(15, "통지", "처리 결과는 우편으로 통지한다."),
    art(16, "월별 현황", "유지관리 처리 건수는 월별로 집계하며 추이는 아래 그림과 같다."),
    art(17, "민원 대응", "유지관리 관련 민원은 7영업일 이내에 회신한다."),
    art(18, "서류 보관", "변경 신청 서류는 계약 소멸 후 5년간 보관한다."),
    art(19, "교육", "담당자는 연 1회 유지관리 실무 교육을 받는다."),
    art(20, "점검", "유지관리 처리 결과는 분기별로 표본 점검한다."),
    art(21, "예외 처리", "이 매뉴얼에 없는 사항은 팀장이 판단한다."),
    art(22, "개정", "이 매뉴얼은 연 1회 검토한다."),
]
PDF_CHG = dict(
    mod={
        "제6조": (["납입주기 변경은 월납, 분기납, 연납 사이에서 할 수 있다."], "분기납 추가", None),
        "제7조": (["보험가입금액 감액은 최초 가입금액의 70%까지 할 수 있다."], "감액 한도 상향", None),
        "제10조": (["약관대출 한도는 해지환급금의 90%로 한다."], "대출 한도 상향", None),
        "제14조": (["변경 신청은 접수일부터 2영업일 이내에 처리한다."], "처리 기한 단축", None),
        "제15조": (["처리 결과는 고객이 선택한 우편, 문자, 이메일로 통지한다."], "통지 수단 확대", "통지 방법"),
        "제17조": (None, "조문 제목만 변경", "민원 처리"),
    },
    add=[
        ("제5조", art(5, "지정대리청구인", "계약자는 지정대리청구인을 1명 지정하거나 바꿀 수 있다.", sub=2), "지정대리청구인 신설"),
        ("제12조", art(12, "간편 인증", "모바일 간편 인증으로 본인 확인을 대신할 수 있다.", sub=2), "간편 인증 도입"),
        ("제19조", art(19, "고령 고객 응대", "만 70세 이상 고객의 변경 신청은 가족 동석을 안내할 수 있다.", sub=2), "고령 고객 보호"),
    ],
    delete={"제13조": "팩스 신청 폐지", "제21조": "예외 처리 조항 삭제"},
)

WORD_OLD = [
    art(1, "목적", "이 내규는 보험 모집 과정의 질서 유지에 필요한 사항을 정함을 목적으로 한다."),
    art(2, "적용대상", "이 내규는 전속 설계사와 법인보험대리점 모집인에게 적용한다."),
    art(3, "금지 행위", "① 모집인은 보험료 대납을 하여서는 안 된다.", "② 모집인은 특별이익을 제공하여서는 안 된다."),
    art(4, "설명 의무", "모집인은 상품의 중요 사항을 고객에게 설명하고 확인서를 받는다."),
    art(5, "위반 등급", "위반 행위는 경중에 따라 경미, 중대, 중대 반복으로 구분하며 등급별 조치는 아래 그림과 같다."),
    art(6, "조치 기준", "위반 등급별 조치 기준은 다음 표와 같다."),
    art(7, "완전판매 모니터링", "신계약의 10%를 표본 추출하여 완전판매 여부를 확인한다."),
    art(8, "해피콜", "청약일부터 7일 이내에 해피콜을 실시한다."),
    art(9, "불완전판매 환수", "불완전판매로 계약이 취소되면 지급한 수수료를 전액 환수한다."),
    art(10, "자율 점검", "모집 조직은 반기 1회 자율 점검을 실시한다."),
    art(11, "신고 채널", "모집질서 위반은 준법감시팀 전화로 신고할 수 있다."),
    art(12, "신고자 보호", "신고자의 신원은 비밀로 한다."),
    art(13, "교육", "모집인은 연 2시간 이상 모집질서 교육을 이수한다."),
    art(14, "서면 서약", "신규 모집인은 위촉 시 모집질서 준수 서약서를 제출한다."),
    art(15, "외부 감사", "외부 감사 결과는 이사회에 보고한다."),
    art(16, "자료 보관", "점검 자료는 3년간 보관한다."),
    art(17, "포상", "우수 조직은 연 1회 포상한다."),
    art(18, "위임", "세부 운영 기준은 준법감시인이 정한다."),
    art(19, "해석", "이 내규의 해석은 준법감시인이 한다."),
    art(20, "시행", "이 내규는 대표이사 승인일부터 시행한다."),
]
TABLE_OLD = [["등급", "1차", "2차", "3차"], ["경미", "주의", "경고", "교육"], ["중대", "경고", "업무정지 1개월", "해촉"]]
TABLE_NEW = [["등급", "1차", "2차", "3차"], ["경미", "주의", "경고", "교육"], ["중대", "업무정지 1개월", "업무정지 3개월", "해촉"]]
WORD_CHG = dict(
    mod={
        "제6조": (None, "조치 기준표 강화(표 안 내용 변경)", None),
        "제7조": (["신계약의 20%를 표본 추출하여 완전판매 여부를 확인한다."], "모니터링 표본 확대", None),
        "제8조": (["청약일부터 3일 이내에 해피콜을 실시한다."], "해피콜 기한 단축", None),
        "제11조": (["모집질서 위반은 준법감시팀 전화와 사내 신고 앱으로 신고할 수 있다."], "신고 앱 추가", None),
        "제13조": (["모집인은 연 4시간 이상 모집질서 교육을 이수한다."], "교육 시간 확대", None),
        "제16조": (["점검 자료는 5년간 보관한다."], "보관 기간 연장", None),
    },
    add=[
        ("제3조", art(3, "비대면 모집", "비대면 모집 시 녹취 또는 화면 기록을 보관한다.", sub=2), "비대면 모집 기록"),
        ("제12조", art(12, "보복 금지", "신고를 이유로 신고자에게 불이익을 주어서는 안 된다.", sub=2), "신고자 보복 금지"),
    ],
    delete={"제14조": "서면 서약 전산화로 삭제", "제17조": "포상 조항 삭제(별도 규정)"},
)

EXCEL_OLD = [
    art(1, "목적", "이 기준은 보험료 수납 업무의 처리 기준을 정한다."),
    art(2, "수납 경로", "보험료는 자동이체, 가상계좌, 카드로 수납한다."),
    art(3, "자동이체 출금", "자동이체는 지정일에 출금하고 잔액 부족 시 3영업일 후 재출금한다."),
    art(4, "가상계좌", "가상계좌는 계약별로 1개를 발급한다."),
    art(5, "카드 납부", "① 카드 납부는 보장성 상품에 한한다.", "② 카드 납부 한도는 월 100만원으로 한다."),
    art(6, "선납", "보험료 선납은 최대 12개월까지 할 수 있다."),
    art(7, "선납 할인", "선납 보험료는 공시이율로 할인한다."),
    art(8, "과오납 반환", "과오납 보험료는 확인일부터 5영업일 이내에 반환한다."),
    art(9, "미납 안내", "미납 시 문자와 우편으로 안내한다."),
    art(10, "현금 수납", "지점 창구에서 현금 수납을 할 수 있다."),
    art(11, "영수증", "수납 영수증은 요청 시 우편으로 발송한다."),
    art(12, "수납 마감", "수납 마감 시각은 영업일 오후 4시로 한다."),
    art(13, "대사", "수납 내역은 매일 은행 입금 내역과 대사한다."),
    art(14, "차이 처리", "대사 차이는 2영업일 이내에 원인을 확인한다."),
    art(15, "월 마감", "월 수납 실적은 다음 달 3영업일까지 마감한다."),
    art(16, "수납 현황", "수납 경로별 비중은 아래 그림과 같이 관리한다."),
    art(17, "기록 보관", "수납 관련 기록은 10년간 보관한다."),
    art(18, "세부 사항", "그 밖의 세부 사항은 수납팀장이 정한다."),
]
EXCEL_CHG = dict(
    mod={
        "제2조": (["보험료는 자동이체, 가상계좌, 카드, 간편결제로 수납한다."], "간편결제 추가", None),
        "제3조": (["자동이체는 지정일에 출금하고 잔액 부족 시 2영업일 후와 5영업일 후에 재출금한다."], "재출금 횟수 확대", None),
        "제5조": (["① 카드 납부는 보장성 상품에 한한다.", "② 카드 납부 한도는 월 200만원으로 한다."], "카드 한도 상향", None),
        "제8조": (["과오납 보험료는 확인일부터 3영업일 이내에 반환한다."], "반환 기한 단축", None),
        "제12조": (["수납 마감 시각은 영업일 오후 5시로 한다."], "마감 시각 연장", None),
    },
    add=[
        ("제6조", art(6, "선납 취소", "선납 후 30일 이내에는 남은 선납 보험료를 돌려받을 수 있다.", sub=2), "선납 취소 신설"),
        ("제11조", art(11, "전자 영수증", "수납 영수증은 이메일이나 모바일 앱으로도 받을 수 있다.", sub=2), "전자 영수증 도입"),
    ],
    delete={"제10조": "현금 수납 폐지", "제18조": "위임 조항 삭제"},
)

TEXT_OLD = [
    art(1, "목적", "이 규정은 심사위원회의 구성과 운영에 관한 사항을 정한다."),
    art(2, "기능", "심사위원회는 고액 계약 인수와 이의 신청 건을 심의한다."),
    art(3, "구성", "위원회는 위원장 1명과 위원 4명으로 구성한다."),
    art(4, "위원장", "위원장은 계약심사부장이 맡는다."),
    art(5, "위원 임기", "위원의 임기는 1년으로 하며 연임할 수 있다."),
    art(6, "정기 회의", "정기 회의는 매월 첫째 주 화요일에 연다."),
    art(7, "임시 회의", "위원장은 필요한 경우 임시 회의를 소집할 수 있다."),
    art(8, "의결 정족수", "회의는 재적 위원 과반수 출석과 출석 위원 과반수 찬성으로 의결한다."),
    art(9, "심의 대상", "가입금액 10억원 이상 계약은 위원회 심의를 거친다."),
    art(10, "서면 심의", "긴급한 경우 서면으로 심의할 수 있다."),
    art(11, "회의록", "간사는 회의록을 작성하여 7일 이내에 위원에게 배포한다."),
    art(12, "제척", "위원은 본인과 이해관계가 있는 안건의 심의에서 빠진다."),
    art(13, "외부 자문", "위원회는 필요 시 외부 의사의 자문을 받을 수 있다."),
    art(14, "수당", "외부 자문위원에게는 회당 20만원의 수당을 지급한다."),
    art(15, "보고", "위원회 심의 결과는 분기별로 임원회의에 보고한다."),
    art(16, "시행", "이 규정은 위원회 의결일부터 시행한다."),
]
TEXT_CHG = dict(
    mod={
        "제3조": (["위원회는 위원장 1명과 위원 6명으로 구성한다."], "위원 증원", None),
        "제6조": (["정기 회의는 매월 첫째 주와 셋째 주 화요일에 연다."], "회의 횟수 확대", None),
        "제9조": (["가입금액 7억원 이상 계약은 위원회 심의를 거친다."], "심의 기준 하향", None),
        "제11조": (["간사는 회의록을 작성하여 3일 이내에 위원에게 배포한다."], "회의록 배포 기한 단축", None),
        "제14조": (["외부 자문위원에게는 회당 30만원의 수당을 지급한다."], "수당 인상", None),
    },
    add=[
        ("제7조", art(7, "화상 회의", "위원은 화상 회의로 출석할 수 있으며 이 경우 출석으로 본다.", sub=2), "화상 회의 인정"),
        ("제12조", art(12, "비밀 유지", "위원은 심의 과정에서 알게 된 고객 정보를 외부에 알려서는 안 된다.", sub=2), "비밀 유지 의무"),
    ],
    delete={"제10조": "서면 심의 폐지"},
)


# ------------------------------------------------------------------ 변경 적용·정답 계산
def apply(old, chg):
    new = []
    for a in old:
        k = key(a)
        if k in chg["delete"]:
            continue
        if k in chg["mod"]:
            paras, _, title = chg["mod"][k]
            a = (a[0], a[1], title or a[2], paras or a[3])
        new.append(a)
        new += [na for after, na, _ in chg["add"] if after == k]
    return new


def body(a, tables):
    t = tables.get(key(a))
    # 앱은 Word 표를 '[표] 칸 | 칸 / 칸 | 칸 [/표]'로 읽는다
    return " ".join(a[3]) + (" [표] " + " / ".join(" | ".join(row) for row in t) + " [/표]" if t else "")


def expected_rows(doc_id, title, old, new, chg, t_old, t_new):
    o, n = {key(a): a for a in old}, {key(a): a for a in new}
    order = sorted(set(o) | set(n), key=lambda k: ((o.get(k) or n.get(k))[0], (o.get(k) or n.get(k))[1]))
    rows = []
    for k in order:
        a, b = o.get(k), n.get(k)
        if a and b:
            if body(a, t_old) == body(b, t_new) and a[2] == b[2]:
                continue
            rows.append((f"{k}({b[2]})", "개정", body(a, t_old), body(b, t_new), chg["mod"][k][1]))
        elif b:
            rows.append((f"{k}({b[2]})", "신설", "", body(b, t_new), next(x[2] for x in chg["add"] if key(x[1]) == k)))
        else:
            rows.append((f"{k}({a[2]})", "삭제", body(a, t_old), "", chg["delete"][k]))
    return [(doc_id, title, *r) for r in rows]


# ------------------------------------------------------------------ 파일 만들기
def make_txt(path, title, version, eff, arts):
    lines = [f"[{ORG}] {title} ({version}, 시행일 {eff})", ""]
    for a in arts:
        lines.append(f"{key(a)}({a[2]}) {a[3][0]}")
        lines += a[3][1:]
        lines.append("")
    path.write_text("\n".join(lines), encoding="utf-8")


def make_xlsx(path, title, version, eff, arts, chart, history):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "기준"
    thin = Side(style="thin", color="999999")
    ws["A1"] = f"[{ORG}] {title} ({version}, 시행일 {eff})"
    ws["A1"].font = Font(bold=True, size=13)
    ws.merge_cells("A1:C1")
    for i, h in enumerate(["조항", "제목", "내용"], 1):
        c = ws.cell(row=3, column=i, value=h)
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="2F5597")
        c.alignment = Alignment(horizontal="center")
    for r, a in enumerate(arts, 4):
        vals = [key(a), a[2], "\n".join(a[3])]      # 여러 문단은 한 칸 안에서 줄바꿈
        for i, v in enumerate(vals, 1):
            c = ws.cell(row=r, column=i, value=v)
            c.alignment = Alignment(vertical="top", wrap_text=True)
            c.border = Border(left=thin, right=thin, top=thin, bottom=thin)
    ws.column_dimensions["A"].width = 12
    ws.column_dimensions["B"].width = 16
    ws.column_dimensions["C"].width = 70
    img = XLImage(png(chart))
    img.width, img.height = 480, 208
    ws.add_image(img, "E3")
    hs = wb.create_sheet("개정이력")                 # 두 번째 시트는 앱이 읽지 않아야 한다
    hs.append(["버전", "시행일", "주요 내용"])
    for row in history:
        hs.append(row)
    wb.save(path)


def main():
    if OUT.exists():
        shutil.rmtree(OUT)
    OUT.mkdir(parents=True)
    tmp = Path(tempfile.mkdtemp())

    pdf_imgs = {"제3조": [flowchart("[그림 1] 계약 유지관리 흐름", ["접수", "본인 확인", "처리", "결과 통지"], "#dfe9fb")],
                "제16조": [bar_chart("[그림 2] 월별 유지관리 처리 건수 (가상)", ["7월", "8월", "9월", "10월", "11월", "12월"], [1320, 1275, 1410, 1388, 1502, 1617], "#4f8a6e")]}
    word_imgs = {"제5조": [matrix("[그림 1] 위반 등급별 조치", ["경미", "중대", "중대 반복"], ["1차", "2차", "3차"],
                                [["표준", "표준", "할증"], ["할증", "보류", "거절"], ["보류", "거절", "거절"]])]}
    chart = bar_chart("[그림] 수납 경로별 비중 (가상, %)", ["자동이체", "가상계좌", "카드", "기타"], [68, 17, 12, 3], "#c4783a")

    docs = [
        dict(id="DOC-201", title="보험계약 유지관리 업무매뉴얼", type="업무매뉴얼", stem="보험계약_유지관리_업무매뉴얼", fmt="pdf",
             old=PDF_OLD, chg=PDF_CHG, v=("v1.0", "v2.0"), eff=("2026-03-01", "2026-11-01"), owner="담당자B", t=({}, {})),
        dict(id="DOC-202", title="모집질서 준수 내규", type="내규", stem="모집질서_준수_내규", fmt="docx",
             old=WORD_OLD, chg=WORD_CHG, v=("v2.0", "v2.1"), eff=("2026-01-01", "2026-11-01"), owner="담당자C",
             t=({"제6조": TABLE_OLD}, {"제6조": TABLE_NEW})),
        dict(id="DOC-203", title="보험료 수납 사무처리기준", type="사무처리기준", stem="보험료_수납_사무처리기준", fmt="xlsx",
             old=EXCEL_OLD, chg=EXCEL_CHG, v=("v1.0", "v1.1"), eff=("2026-01-01", "2026-11-01"), owner="담당자A", t=({}, {})),
        dict(id="DOC-204", title="심사위원회 운영규정", type="내규", stem="심사위원회_운영규정", fmt="txt",
             old=TEXT_OLD, chg=TEXT_CHG, v=("v3.0", "v3.1"), eff=("2025-12-01", "2026-11-01"), owner="담당자C", t=({}, {})),
    ]
    expected, meta, summary = [], [], []
    for doc in docs:
        new = apply(doc["old"], doc["chg"])
        for i, (arts, tables) in enumerate([(doc["old"], doc["t"][0]), (new, doc["t"][1])]):
            ver, eff = doc["v"][i], doc["eff"][i]
            name = f"{doc['stem']}_{ver}.{doc['fmt']}"
            if doc["fmt"] == "pdf":
                d = tmp / f"{doc['stem']}_{ver}.docx"
                make_docx(d, doc["title"], ver, eff, arts, pdf_imgs, tables, landscape=True)
                shutil.move(str(to_pdf(d, tmp)), OUT / name)
            elif doc["fmt"] == "docx":
                make_docx(OUT / name, doc["title"], ver, eff, arts, word_imgs, tables)
            elif doc["fmt"] == "xlsx":
                hist = [["v1.0", "2026-01-01", "제정"]] + ([["v1.1", "2026-11-01", "간편결제·전자영수증 도입"]] if i else [])
                make_xlsx(OUT / name, doc["title"], ver, eff, arts, chart, hist)
            else:
                make_txt(OUT / name, doc["title"], ver, eff, arts)
            meta.append([f"{doc['id']}-{i + 1}", doc["title"], doc["type"], ver, eff, doc["owner"], name])
        rows = expected_rows(doc["id"], doc["title"], doc["old"], new, doc["chg"], *doc["t"])
        expected += rows
        c = Counter(r[3] for r in rows)
        summary.append(f"  {doc['fmt'].upper():5} {doc['stem']}_{doc['v'][0]} → {doc['v'][1]}  조항 {len(doc['old'])}개, 변경 {len(rows)}건 (개정 {c['개정']} · 신설 {c['신설']} · 삭제 {c['삭제']})")
    shutil.rmtree(tmp)

    with open(OUT / "expected_changes_testpack2.csv", "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow(["change_id", "doc_id", "doc_title", "section", "change_type", "before_text", "after_text", "note"])
        for i, r in enumerate(expected, 1):
            w.writerow([f"TC2-{i:03d}", *r])
    base = list(csv.reader(open(ROOT / "data" / "test_pack" / "documents_testpack.csv", encoding="utf-8-sig")))
    with open(OUT / "documents_testpack2.csv", "w", encoding="utf-8-sig", newline="") as f:
        csv.writer(f).writerows(base + meta)

    c = Counter(r[3] for r in expected)
    (OUT / "README.txt").write_text(f"""Hana Life 시험용 더미데이터 2차 묶음 — 형식별 신구 문서 1쌍씩 (모든 회사명·금액·인명은 가상)

[문서]
{chr(10).join(summary)}

  PDF   흐름도·월별 처리 차트 이미지 포함
  Word  위반 등급 매트릭스 이미지 + 조치 기준표(표 안 내용만 바뀐 개정 1건)
  Excel '조항 | 제목 | 내용' 표 형식, 한 칸 안 여러 문단(①②), 차트 이미지, 두 번째 시트 '개정이력'(앱은 첫 시트만 읽음)
  Text  일반 텍스트

[정답표] expected_changes_testpack2.csv  총 {len(expected)}건 (개정 {c['개정']} · 신설 {c['신설']} · 삭제 {c['삭제']})
  - 조문 제목만 바뀐 개정 1건 포함: PDF 제17조(민원 대응 → 민원 처리). 내용은 같고 제목만 다릅니다.
  - PDF 제15조는 내용과 제목(통지 → 통지 방법)이 함께 바뀝니다.

[넣는 순서]
  1. '기준 데이터' 탭 → '문서 메타 CSV'에 documents_testpack2.csv, '정답표 CSV'에 expected_changes_testpack2.csv
  2. '문서함·분류' 탭에 문서 8개를 한꺼번에 끌어다 놓기
  3. '전체 신구 비교 실행' → '정답 검증' 탭에서 {len(expected)}건 일치 확인
""", encoding="utf-8")
    print(f"expected {len(expected)} rows {dict(c)}")
    print("\n".join(summary))


if __name__ == "__main__":
    main()
