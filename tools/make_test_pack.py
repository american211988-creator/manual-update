"""시험용 더미데이터 묶음(data/test_pack/)을 만든다.

- 신구 문서 3쌍: 조항 24~28개, 문서마다 변경 12~16건, 흐름도·도표 이미지 포함
  1) 언더라이팅 업무매뉴얼      v3.0 → v4.0  (.docx, 이미지 3개)
  2) 보험금 지급 사무처리기준    v2.1 → v3.0  (.pdf,  이미지 2개)
  3) 계약심사 전결 내규          v1.5 → v2.0  (.docx, 이미지 1개 + 표)
- 스캔본 예시 PDF 1개 (글자 없는 이미지 PDF → 앱이 안내 문구를 띄워야 함)
- expected_changes_testpack.csv : 정답표 (변경 목록에서 자동 계산)
- documents_testpack.csv       : 기존 샘플 + 시험 문서 메타

사용법: PYTHONPATH=<python-docx 설치 경로> python3 tools/make_test_pack.py
필요: python-docx, Pillow, LibreOffice(soffice, PDF 변환)
모든 회사명·금액·인명은 가상입니다.
"""
import csv
import io
import shutil
import subprocess
import tempfile
from pathlib import Path

from docx import Document
from docx.enum.section import WD_ORIENT
from docx.shared import Cm, Pt
from PIL import Image, ImageDraw, ImageFont

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "data" / "test_pack"
FONT = "/usr/share/fonts/truetype/wqy/wqy-zenhei.ttc"
ORG = "Hana Life 계약심사부"


# ---------------------------------------------------------------- 이미지
def font(size):
    return ImageFont.truetype(FONT, size)


def flowchart(title, steps, color):
    w, h = 1400, 300
    im = Image.new("RGB", (w, h), "white")
    d = ImageDraw.Draw(im)
    d.text((30, 20), title, fill="#222", font=font(30))
    bw, gap = (w - 60 - (len(steps) - 1) * 40) // len(steps), 40
    for i, s in enumerate(steps):
        x = 30 + i * (bw + gap)
        d.rounded_rectangle((x, 110, x + bw, 230), 18, fill=color, outline="#333", width=3)
        tw = d.textlength(s, font=font(26))
        d.text((x + (bw - tw) / 2, 152), s, fill="#111", font=font(26))
        if i < len(steps) - 1:
            d.line((x + bw + 4, 170, x + bw + gap - 6, 170), fill="#333", width=4)
            d.polygon([(x + bw + gap - 6, 170), (x + bw + gap - 18, 162), (x + bw + gap - 18, 178)], fill="#333")
    return im


def bar_chart(title, labels, values, color):
    w, h = 1200, 520
    im = Image.new("RGB", (w, h), "white")
    d = ImageDraw.Draw(im)
    d.text((30, 20), title, fill="#222", font=font(30))
    top, base, left = 90, 440, 90
    mx = max(values)
    d.line((left, base, w - 40, base), fill="#666", width=2)
    bw = (w - left - 60) // len(values)
    for i, (lb, v) in enumerate(zip(labels, values)):
        x = left + i * bw + 12
        y = base - (base - top) * v / mx
        d.rectangle((x, y, x + bw - 24, base), fill=color)
        d.text((x, y - 30), str(v), fill="#222", font=font(20))
        d.text((x, base + 12), lb, fill="#444", font=font(20))
    return im


def matrix(title, rows, cols, cells):
    w, h = 1100, 120 + 70 * (len(rows) + 1)
    im = Image.new("RGB", (w, h), "white")
    d = ImageDraw.Draw(im)
    d.text((30, 20), title, fill="#222", font=font(30))
    cw = (w - 60) // (len(cols) + 1)
    pal = {"표준": "#d9f2df", "할증": "#fbefc4", "보류": "#f9d9cf", "거절": "#e8c4c4"}
    for j, c in enumerate([""] + cols):
        d.text((30 + j * cw + 12, 92), c, fill="#222", font=font(22))
    for i, r in enumerate(rows):
        y = 140 + i * 70
        d.text((42, y + 20), r, fill="#222", font=font(22))
        for j, v in enumerate(cells[i]):
            x = 30 + (j + 1) * cw
            d.rectangle((x, y, x + cw - 6, y + 62), fill=pal.get(v, "#eee"), outline="#999")
            d.text((x + 14, y + 18), v, fill="#222", font=font(22))
    return im


def logo():
    im = Image.new("RGB", (420, 120), "white")
    d = ImageDraw.Draw(im)
    d.ellipse((10, 10, 110, 110), fill="#1f7a5c")
    d.text((38, 30), "H", fill="white", font=font(54))
    d.text((130, 34), "Hana Life", fill="#1f7a5c", font=font(44))
    return im


def png(im):
    b = io.BytesIO()
    im.save(b, "PNG")
    b.seek(0)
    return b


# ---------------------------------------------------------------- 문서 내용
# 조항: (번호, 의N, 제목, [문단...]) — 문단이 여러 개면 ①②로 시작한다.
# 변경: mod={조항키: (새 문단 목록, 비고)}, add=[(뒤에 올 조항키, 조항, 비고)], delete={조항키: 비고}

def art(n, title, *paras, sub=0):
    return (n, sub, title, list(paras))


def key(a):
    return f"제{a[0]}조" + (f"의{a[1]}" if a[1] else "")


DOC1_OLD = [
    art(1, "목적", "이 매뉴얼은 Hana Life 계약심사부의 언더라이팅 업무 절차와 판단 기준을 정함을 목적으로 한다."),
    art(2, "적용범위", "이 매뉴얼은 개인보험 신계약 및 부활 계약의 언더라이팅에 적용한다."),
    art(3, "용어의 정의", "① 언더라이팅이란 피보험자의 위험을 평가하여 인수 여부와 조건을 결정하는 업무를 말한다.", "② 표준체란 보험료 할증 없이 인수할 수 있는 피보험자를 말한다."),
    art(4, "심사 흐름", "언더라이팅은 접수, 서류 확인, 위험 평가, 인수 결정, 결과 통보 순으로 처리하며 세부 흐름은 아래 그림과 같다."),
    art(5, "심사자 자격", "언더라이팅 심사자는 사내 심사 자격 2급 이상을 보유하여야 한다."),
    art(6, "청약서 확인", "① 심사자는 청약서의 필수 기재 항목 누락 여부를 확인한다.", "② 누락 항목이 있는 경우 영업 채널에 3영업일 이내 보완을 요청한다."),
    art(7, "고지사항 검토", "심사자는 계약 전 알릴 의무 사항을 최근 5년간의 병력 기준으로 검토한다."),
    art(8, "무진단 한도", "무진단으로 인수할 수 있는 1인당 사망보험금 한도는 1억원으로 한다."),
    art(9, "건강진단 대상", "사망보험금 합계가 2억원을 초과하거나 피보험자가 만 55세 이상인 경우 건강진단을 실시한다."),
    art(10, "진단 항목", "건강진단은 혈압, 혈액, 소변 검사를 기본으로 하며 필요 시 심전도를 추가한다."),
    art(11, "위험등급 분류", "피보험자의 위험등급은 직업, 병력, 생활습관을 종합하여 아래 매트릭스에 따라 분류한다."),
    art(12, "직업 위험", "직업 위험등급은 1급부터 4급까지로 구분하며 4급은 인수를 보류한다."),
    art(13, "할증 인수", "① 표준체 기준을 초과하는 위험이 있는 경우 보험료를 할증하여 인수할 수 있다.", "② 할증률은 최대 100%로 한다."),
    art(14, "부담보 인수", "특정 신체 부위에 위험이 있는 경우 해당 부위를 1년간 부담보로 하여 인수할 수 있다."),
    art(15, "인수 거절", "위험이 인수 기준을 크게 초과하는 경우 인수를 거절하고 사유를 기록한다."),
    art(16, "처리기한", "심사자는 서류 완비일부터 5영업일 이내에 인수 여부를 결정한다."),
    art(17, "상위 심사", "가입금액 5억원 이상 계약은 수석 심사자의 2차 심사를 받아야 한다."),
    art(18, "재심사", "인수 거절 건에 대해 고객이 재심사를 요청하는 경우 1회에 한하여 재심사할 수 있다."),
    art(19, "결과 통보", "인수 결과는 영업 채널을 통해 서면으로 고객에게 통보한다."),
    art(20, "고위험 직종 확인", "고위험 직종 종사자는 재직증명서를 추가로 제출받는다."),
    art(21, "해외 체류자", "해외 체류 예정 기간이 6개월 이상인 경우 인수를 보류한다."),
    art(22, "부활 심사", "실효 후 부활 청약 건은 신계약 기준을 동일하게 적용하여 심사한다."),
    art(23, "팩스 서류", "팩스로 접수된 서류는 원본 도착 전까지 가심사로 처리한다."),
    art(24, "심사 의견서", "조건부 인수 또는 거절 시 심사 의견서를 작성하여 전산에 등록한다."),
    art(25, "통계 관리", "심사부는 월별 인수 결과 통계를 작성하여 부서장에게 보고한다."),
    art(26, "교육", "심사자는 연 1회 이상 언더라이팅 실무 교육을 이수하여야 한다."),
    art(27, "기록 보관", "심사 관련 서류는 계약 소멸 후 5년간 보관한다."),
    art(28, "매뉴얼 개정", "이 매뉴얼은 연 1회 정기 검토하고 필요 시 개정한다."),
]
DOC1_CHG = dict(
    mod={
        "제5조": (["언더라이팅 심사자는 사내 심사 자격 1급 이상을 보유하여야 한다."], "심사자 자격 상향"),
        "제6조": (["① 심사자는 청약서의 필수 기재 항목 누락 여부를 확인한다.", "② 누락 항목이 있는 경우 영업 채널에 2영업일 이내 보완을 요청한다."], "보완 요청 기한 단축"),
        "제7조": (["심사자는 계약 전 알릴 의무 사항을 최근 10년간의 병력 기준으로 검토한다."], "고지 검토 기간 확대"),
        "제8조": (["무진단으로 인수할 수 있는 1인당 사망보험금 한도는 1억 5천만원으로 한다."], "무진단 한도 상향"),
        "제9조": (["사망보험금 합계가 3억원을 초과하거나 피보험자가 만 50세 이상인 경우 건강진단을 실시한다."], "진단 기준 금액·연령 조정"),
        "제12조": (["직업 위험등급은 1급부터 5급까지로 구분하며 5급은 인수를 보류한다."], "직업 등급 세분화"),
        "제13조": (["① 표준체 기준을 초과하는 위험이 있는 경우 보험료를 할증하여 인수할 수 있다.", "② 할증률은 최대 150%로 한다."], "할증 한도 상향"),
        "제16조": (["심사자는 서류 완비일부터 3영업일 이내에 인수 여부를 결정한다."], "처리기한 단축"),
        "제17조": (["가입금액 3억원 이상 계약은 수석 심사자의 2차 심사를 받아야 한다."], "2차 심사 기준 하향"),
        "제26조": (["심사자는 연 2회 이상 언더라이팅 실무 교육을 이수하여야 한다."], "교육 횟수 확대"),
    },
    add=[
        ("제10조", art(10, "비대면 진단", "모바일 건강검진 결과는 공인 검진기관 발급분에 한하여 건강진단으로 인정한다.", sub=2), "비대면 진단 인정"),
        ("제15조", art(15, "AI 보조 심사", "① 전산 AI 보조 심사 결과는 참고용으로만 사용한다.", "② 최종 인수 결정은 심사자가 직접 한다.", sub=2), "AI 보조 심사 원칙"),
        ("제19조", art(19, "전자 통보", "고객이 동의한 경우 인수 결과를 모바일 전자문서로 통보할 수 있다.", sub=2), "전자 통보 도입"),
        ("제25조", art(25, "품질 점검", "심사부는 분기별로 심사 완료 건의 5%를 표본 추출하여 품질을 점검한다.", sub=2), "심사 품질 점검 신설"),
    ],
    delete={"제21조": "해외 체류자 보류 기준 폐지", "제23조": "팩스 접수 폐지"},
)

DOC2_OLD = [
    art(1, "목적", "이 기준은 Hana Life 보험금 지급 심사 업무의 처리 기준을 정함을 목적으로 한다."),
    art(2, "적용범위", "이 기준은 개인보험의 사망, 입원, 수술, 진단 보험금 청구 건에 적용한다."),
    art(3, "청구 접수", "보험금 청구는 창구, 우편, 모바일 앱으로 접수한다."),
    art(4, "필수 서류", "① 청구서, 신분증 사본, 진단서를 필수 서류로 한다.", "② 사망 보험금은 사망진단서와 가족관계증명서를 추가로 제출받는다."),
    art(5, "지급 절차", "보험금 지급은 접수, 서류 심사, 사고 조사, 지급 결정, 송금 순으로 처리한다."),
    art(6, "소액 청구", "청구금액 100만원 이하 건은 진단서 대신 처방전으로 갈음할 수 있다."),
    art(7, "지급 기한", "서류 접수일부터 3영업일 이내에 보험금을 지급한다."),
    art(8, "지급 지연", "조사가 필요한 경우 지급 예정일을 고객에게 안내하고 30일 이내에 지급한다."),
    art(9, "사고 조사 대상", "청구금액 3천만원 이상이거나 계약일부터 2년 이내 발생한 사고는 사고 조사 대상으로 한다."),
    art(10, "조사 위탁", "사고 조사는 회사가 지정한 손해사정법인에 위탁할 수 있다."),
    art(11, "면책 검토", "약관상 면책 사유 해당 여부를 사고 경위서와 의무기록으로 검토한다."),
    art(12, "중복 청구", "동일 사고로 중복 청구된 건은 먼저 접수된 건만 처리한다."),
    art(13, "지급 거절", "지급 거절 시 사유와 근거 약관 조항을 서면으로 통보한다."),
    art(14, "분쟁 조정", "고객이 지급 결정에 이의를 제기하면 분쟁조정위원회에 상정한다."),
    art(15, "이의 신청 기간", "고객은 결정 통보일부터 30일 이내에 이의를 신청할 수 있다."),
    art(16, "대리 청구", "피보험자가 청구할 수 없는 경우 지정대리청구인이 청구할 수 있다."),
    art(17, "현금 지급", "보험금은 원칙적으로 계좌 송금으로 지급하며 현금 지급은 50만원 이하에 한한다."),
    art(18, "지연 이자", "지급 기한을 넘긴 경우 보험계약대출이율로 계산한 지연 이자를 더하여 지급한다."),
    art(19, "월별 실적", "심사부는 월별 지급 처리 건수와 평균 처리일을 집계한다."),
    art(20, "우편 접수 대장", "우편으로 접수된 청구서는 수기 접수 대장에 기록한다."),
    art(21, "서류 보관", "보험금 청구 서류는 지급 완료 후 5년간 보관한다."),
    art(22, "고액 결재", "지급액 1억원 이상은 심사부장의 결재를 받는다."),
    art(23, "사기 의심 보고", "보험사기가 의심되는 건은 특별조사팀에 즉시 보고한다."),
    art(24, "교육", "지급 심사자는 연 1회 보험사기 예방 교육을 이수한다."),
    art(25, "기준 점검", "이 기준은 반기 1회 점검한다."),
    art(26, "시행 세칙", "이 기준의 시행에 필요한 세부 사항은 심사부장이 따로 정한다."),
]
DOC2_CHG = dict(
    mod={
        "제3조": (["보험금 청구는 창구, 우편, 모바일 앱, 홈페이지로 접수한다."], "홈페이지 접수 추가"),
        "제6조": (["청구금액 200만원 이하 건은 진단서 대신 처방전으로 갈음할 수 있다."], "소액 기준 상향"),
        "제7조": (["서류 접수일부터 2영업일 이내에 보험금을 지급한다."], "지급 기한 단축"),
        "제9조": (["청구금액 5천만원 이상이거나 계약일부터 1년 이내 발생한 사고는 사고 조사 대상으로 한다."], "조사 대상 기준 조정"),
        "제15조": (["고객은 결정 통보일부터 60일 이내에 이의를 신청할 수 있다."], "이의 신청 기간 확대"),
        "제17조": (["보험금은 계좌 송금으로만 지급한다."], "현금 지급 폐지"),
        "제22조": (["지급액 5천만원 이상은 심사부장의 결재를 받는다."], "고액 결재 기준 하향"),
        "제24조": (["지급 심사자는 연 2회 보험사기 예방 교육을 이수한다."], "교육 횟수 확대"),
    },
    add=[
        ("제7조", art(7, "자동 지급", "청구금액 30만원 이하이고 서류가 완비된 건은 전산 자동 심사로 즉시 지급할 수 있다.", sub=2), "자동 지급 도입"),
        ("제12조", art(12, "실손 중복 확인", "실손 보험금 청구 건은 타사 가입 내역을 조회하여 비례보상 여부를 확인한다.", sub=2), "실손 비례보상 확인"),
        ("제23조", art(23, "개인정보 마스킹", "심사 화면과 출력물의 주민등록번호 뒷자리는 마스킹하여 표시한다.", sub=2), "개인정보 보호 강화"),
    ],
    delete={"제10조": "조사 위탁 조항 삭제(세칙 이관)", "제20조": "수기 접수 대장 폐지", "제26조": "시행 세칙 조항 삭제"},
)

DOC3_OLD = [
    art(1, "목적", "이 내규는 Hana Life 계약심사부의 업무 전결 권한과 결재 절차를 정함을 목적으로 한다."),
    art(2, "정의", "전결이란 위임받은 범위에서 최종 결정권자를 대신하여 결재하는 것을 말한다."),
    art(3, "결재 라인", "결재는 담당자, 파트장, 팀장, 부장 순으로 진행하며 결재 라인은 아래 그림과 같다."),
    art(4, "전결권자", "전결권자는 팀장 이상으로 한다."),
    art(5, "인수 전결 한도", "인수심사 전결 한도는 다음 표와 같다."),
    art(6, "지급 전결 한도", "보험금 지급 전결 한도는 팀장 5천만원, 부장 2억원으로 한다."),
    art(7, "예외 결재", "전결 한도를 넘는 건은 담당 임원의 결재를 받는다."),
    art(8, "대결", "전결권자가 부재 중인 경우 차하위 직위자가 대결하고 사후 보고한다."),
    art(9, "사후 보고", "대결한 건은 3영업일 이내에 전결권자에게 보고한다."),
    art(10, "전자 결재", "결재는 전자결재 시스템으로 처리한다."),
    art(11, "긴급 결재", "긴급한 경우 구두 승인 후 1영업일 이내에 전자결재를 상신한다."),
    art(12, "결재 기한", "결재권자는 상신일부터 2영업일 이내에 결재한다."),
    art(13, "반려", "결재권자가 반려하는 경우 사유를 기재한다."),
    art(14, "기록 관리", "결재 문서는 10년간 보관한다."),
    art(15, "위임 제한", "전결 권한은 재위임할 수 없다."),
    art(16, "이해 상충", "결재권자는 본인 또는 친족 관련 건을 결재할 수 없다."),
    art(17, "감사", "내부감사부서는 연 1회 전결 운영 실태를 점검한다."),
    art(18, "위반 시 조치", "이 내규를 위반한 경우 인사 규정에 따라 조치한다."),
    art(19, "서면 결재", "전산 장애 시에는 서면으로 결재하고 복구 후 전산에 등록한다."),
    art(20, "통계 보고", "팀장은 분기별 전결 처리 현황을 부장에게 보고한다."),
    art(21, "교육", "신규 전결권자는 임명 후 1개월 이내에 전결 교육을 이수한다."),
    art(22, "해석", "이 내규의 해석에 이견이 있는 경우 부장이 결정한다."),
    art(23, "개정 절차", "이 내규는 부서 회의를 거쳐 부장이 개정한다."),
    art(24, "시행", "이 내규는 부장 승인일부터 시행한다."),
]
# 제5조에는 전결 한도표가 들어 있다(표 칸의 글자도 조항 내용으로 읽힌다)
TABLE_OLD = [["구분", "파트장", "팀장", "부장"], ["무진단", "5천만원", "1억원", "3억원"], ["진단", "1억원", "2억원", "5억원"]]
TABLE_NEW = [["구분", "파트장", "팀장", "부장"], ["무진단", "7천만원", "1억 5천만원", "3억원"], ["진단", "1억원", "3억원", "5억원"]]
DOC3_CHG = dict(
    mod={
        "제4조": (["전결권자는 파트장 이상으로 한다."], "전결권자 범위 확대"),
        "제5조": (["인수심사 전결 한도는 다음 표와 같다."], "전결 한도표 상향(표 안 금액 변경)"),
        "제6조": (["보험금 지급 전결 한도는 팀장 1억원, 부장 3억원으로 한다."], "지급 전결 한도 상향"),
        "제9조": (["대결한 건은 1영업일 이내에 전결권자에게 보고한다."], "사후 보고 기한 단축"),
        "제12조": (["결재권자는 상신일부터 1영업일 이내에 결재한다."], "결재 기한 단축"),
        "제17조": (["내부감사부서는 반기 1회 전결 운영 실태를 점검한다."], "점검 주기 단축"),
        "제21조": (["신규 전결권자는 임명 후 2주 이내에 전결 교육을 이수한다."], "교육 기한 단축"),
    },
    add=[
        ("제10조", art(10, "모바일 결재", "전자결재는 사내 인증을 거친 모바일 기기에서도 처리할 수 있다.", sub=2), "모바일 결재 허용"),
        ("제16조", art(16, "자기 결재 금지", "담당자가 처리한 건을 같은 사람이 결재권자로서 결재할 수 없다.", sub=2), "자기 결재 금지"),
        ("제20조", art(20, "AI 심사 결재", "AI 보조 심사를 활용한 건은 결재 문서에 활용 여부를 표시한다.", sub=2), "AI 활용 표시"),
    ],
    delete={"제11조": "구두 승인 후 결재 폐지", "제19조": "서면 결재 조항 삭제"},
)


def apply(old, chg):
    by = {key(a): a for a in old}
    new = []
    for a in old:
        k = key(a)
        if k in chg["delete"]:
            continue
        if k in chg["mod"]:
            a = (a[0], a[1], a[2], chg["mod"][k][0])
        new.append(a)
        for after, na, _ in chg["add"]:
            if after == k:
                new.append(na)
    return new


def body(a, tables):
    text = " ".join(a[3])
    t = tables.get(key(a))
    if t:   # 앱은 Word 표를 '[표] 칸 | 칸 / 칸 | 칸 [/표]'로 읽는다
        text += " [표] " + " / ".join(" | ".join(row) for row in t) + " [/표]"
    return text


def expected_rows(doc_id, title, old, new, chg, t_old, t_new):
    rows = []
    o = {key(a): a for a in old}
    n = {key(a): a for a in new}
    order = sorted(set(o) | set(n), key=lambda k: ((o.get(k) or n.get(k))[0], (o.get(k) or n.get(k))[1]))
    for k in order:
        a, b = o.get(k), n.get(k)
        if a and b:
            if body(a, t_old) == body(b, t_new):
                continue
            rows.append((f"{k}({b[2]})", "개정", body(a, t_old), body(b, t_new), chg["mod"][k][1]))
        elif b:
            note = next(x[2] for x in chg["add"] if key(x[1]) == k)
            rows.append((f"{k}({b[2]})", "신설", "", body(b, t_new), note))
        else:
            rows.append((f"{k}({a[2]})", "삭제", body(a, t_old), "", chg["delete"][k]))
    return [(doc_id, title, *r) for r in rows]


# ---------------------------------------------------------------- 문서 만들기
def make_docx(path, title, version, eff, arts, images=None, tables=None, landscape=False):
    images, tables = images or {}, tables or {}
    d = Document()
    st = d.styles["Normal"]
    st.font.name = "Malgun Gothic"
    st.font.size = Pt(10)
    if landscape:
        s = d.sections[0]
        s.orientation = WD_ORIENT.LANDSCAPE
        s.page_width, s.page_height = s.page_height, s.page_width
    for s in d.sections:
        s.left_margin = s.right_margin = Cm(1.5)
    d.add_picture(png(logo()), width=Cm(4))
    d.add_paragraph(f"[{ORG}] {title} ({version}, 시행일 {eff})")
    for a in arts:
        head = f"{key(a)}({a[2]}) "
        d.add_paragraph(head + a[3][0])
        for p in a[3][1:]:
            d.add_paragraph(p)
        t = tables.get(key(a))
        if t:
            tb = d.add_table(rows=len(t), cols=len(t[0]))
            tb.style = "Table Grid"
            for i, row in enumerate(t):
                for j, c in enumerate(row):
                    tb.cell(i, j).text = c
        for im in images.get(key(a), []):
            d.add_picture(png(im), width=Cm(15))
    d.save(path)


def to_pdf(docx_path, out_dir):
    subprocess.run(["soffice", "--headless", "--convert-to", "pdf", "--outdir", str(out_dir), str(docx_path)],
                   check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=180)
    return Path(out_dir) / (docx_path.stem + ".pdf")


def scanned_pdf(path):
    im = Image.new("RGB", (1240, 1754), "white")
    d = ImageDraw.Draw(im)
    d.text((120, 120), f"[{ORG}] 스캔본 예시 (v1.0)", fill="#333", font=font(40))
    for i in range(14):
        d.text((120, 240 + i * 90), f"제{i + 1}조(스캔 문서) 이 페이지는 이미지로만 되어 있어 글자를 읽을 수 없다.", fill="#444", font=font(30))
    im = im.rotate(0.8, fillcolor="white")
    im.save(path, "PDF", resolution=150)


def main():
    if OUT.exists():
        shutil.rmtree(OUT)
    OUT.mkdir(parents=True)

    flow1 = flowchart("[그림 1] 언더라이팅 심사 흐름", ["접수", "서류 확인", "위험 평가", "인수 결정", "결과 통보"], "#d6e6fb")
    risk = matrix("[그림 2] 위험등급 매트릭스", ["병력 없음", "경증 병력", "중증 병력"], ["직업 1~2급", "직업 3급", "직업 4~5급"],
                  [["표준", "표준", "할증"], ["표준", "할증", "보류"], ["할증", "보류", "거절"]])
    flow2 = flowchart("[그림 1] 보험금 지급 절차", ["접수", "서류 심사", "사고 조사", "지급 결정", "송금"], "#e5f3dc")
    chart = bar_chart("[그림 2] 월별 지급 처리 건수 (가상)", ["1월", "2월", "3월", "4월", "5월", "6월"], [812, 765, 903, 880, 951, 1024], "#5b8def")
    line3 = flowchart("[그림 1] 결재 라인", ["담당자", "파트장", "팀장", "부장"], "#f6e3c9")

    docs = [
        dict(id="DOC-101", title="언더라이팅 업무매뉴얼", type="업무매뉴얼", stem="언더라이팅_업무매뉴얼", fmt="docx",
             old=DOC1_OLD, chg=DOC1_CHG, v=("v3.0", "v4.0"), eff=("2026-01-01", "2026-10-01"),
             images={"제4조": [flow1], "제11조": [risk]}, t=({}, {}), owner="담당자B"),
        dict(id="DOC-102", title="보험금 지급 사무처리기준", type="사무처리기준", stem="보험금지급_사무처리기준", fmt="pdf",
             old=DOC2_OLD, chg=DOC2_CHG, v=("v2.1", "v3.0"), eff=("2026-02-01", "2026-10-01"),
             images={"제5조": [flow2], "제19조": [chart]}, t=({}, {}), owner="담당자A"),
        dict(id="DOC-103", title="계약심사 전결 내규", type="내규", stem="계약심사_전결_내규", fmt="docx",
             old=DOC3_OLD, chg=DOC3_CHG, v=("v1.5", "v2.0"), eff=("2025-07-01", "2026-10-01"),
             images={"제3조": [line3]}, t=({"제5조": TABLE_OLD}, {"제5조": TABLE_NEW}), owner="담당자C"),
    ]
    expected, meta = [], []
    tmp = Path(tempfile.mkdtemp())
    for doc in docs:
        new = apply(doc["old"], doc["chg"])
        for i, (arts, ver, eff, tables) in enumerate([(doc["old"], doc["v"][0], doc["eff"][0], doc["t"][0]),
                                                       (new, doc["v"][1], doc["eff"][1], doc["t"][1])]):
            name = f"{doc['stem']}_{ver}"
            docx_path = (tmp if doc["fmt"] == "pdf" else OUT) / f"{name}.docx"
            make_docx(docx_path, doc["title"], ver, eff, arts, doc["images"], tables, landscape=doc["fmt"] == "pdf")
            if doc["fmt"] == "pdf":
                shutil.move(str(to_pdf(docx_path, tmp)), OUT / f"{name}.pdf")
            meta.append([f"{doc['id']}-{i + 1}", doc["title"], doc["type"], ver, eff, doc["owner"], f"{name}.{doc['fmt']}"])
        expected += expected_rows(doc["id"], doc["title"], doc["old"], new, doc["chg"], *doc["t"])
    shutil.rmtree(tmp)
    scanned_pdf(OUT / "스캔본_예시_업무매뉴얼_v1.0.pdf")

    with open(OUT / "expected_changes_testpack.csv", "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow(["change_id", "doc_id", "doc_title", "section", "change_type", "before_text", "after_text", "note"])
        for i, r in enumerate(expected, 1):
            w.writerow([f"TCH-{i:03d}", *r])
    base = list(csv.reader(open(ROOT / "data" / "documents.csv", encoding="utf-8-sig")))
    with open(OUT / "documents_testpack.csv", "w", encoding="utf-8-sig", newline="") as f:
        csv.writer(f).writerows(base + meta)

    from collections import Counter
    c = Counter(r[3] for r in expected)
    per = Counter(r[0] for r in expected)
    (OUT / "README.txt").write_text(f"""Hana Life 시험용 더미데이터 묶음 (모든 회사명·금액·인명은 가상)

[문서]  신구 3쌍 + 스캔본 1개
  언더라이팅_업무매뉴얼_v3.0/v4.0.docx      조항 {len(DOC1_OLD)}개, 변경 {per['DOC-101']}건, 이미지 3개(로고·심사 흐름도·위험등급 매트릭스)
  보험금지급_사무처리기준_v2.1/v3.0.pdf     조항 {len(DOC2_OLD)}개, 변경 {per['DOC-102']}건, 이미지 3개(로고·지급 절차도·월별 처리 차트)
  계약심사_전결_내규_v1.5/v2.0.docx         조항 {len(DOC3_OLD)}개, 변경 {per['DOC-103']}건, 이미지 2개(로고·결재 라인) + 전결 한도표
  스캔본_예시_업무매뉴얼_v1.0.pdf            글자 없는 이미지 PDF → 앱이 '스캔 이미지 PDF는 지원하지 않습니다'를 띄워야 정상

[정답표] expected_changes_testpack.csv  총 {len(expected)}건 (개정 {c['개정']} · 신설 {c['신설']} · 삭제 {c['삭제']})
  - 제5조(인수 전결 한도)는 본문 문장은 같고 '표 안의 금액'만 바뀐 개정입니다. 표 칸의 글자도 조항 내용으로 읽히는지 확인하는 항목입니다.
  - 이미지 속 글자는 비교 대상이 아닙니다(그림이 바뀌어도 변경으로 잡히지 않음).

[넣는 순서]
  1. 앱 '기준 데이터' 탭 → '문서 메타 CSV'에 documents_testpack.csv, '정답표 CSV'에 expected_changes_testpack.csv
  2. '문서함·분류' 탭에 문서 7개를 한꺼번에 끌어다 놓기
  3. '전체 신구 비교 실행' → '정답 검증' 탭에서 {len(expected)}건 일치 확인
""", encoding="utf-8")
    print(f"expected {len(expected)} rows: {dict(c)} per doc {dict(per)}")


if __name__ == "__main__":
    main()
