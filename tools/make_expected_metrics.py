"""분류·변경률·알림·집계 검증용 정답(기대값)을 앱과 따로 계산해 CSV로 만든다.

- 기본 샘플(data/documents/*.txt): 파일을 직접 읽어 조항을 센다 (앱의 JS 코드와 별개로 구현)
- 시험용 1·2차 묶음: 생성 스크립트의 조항 목록과 변경 목록에서 계산한다
- 알림 대상: data/users.csv의 구독 유형으로 계산한다

출력: data/expected_metrics.csv, data/test_pack/expected_metrics_testpack.csv, data/test_pack2/expected_metrics_testpack2.csv
열: doc_id, doc_title, doc_type, old_version, new_version, effective_month, articles_old, articles_new,
    articles_total(신구 합집합), changes, revised, added, deleted, change_rate(%, 소수 첫째 자리), subscribers(| 구분)
사용법: python3 tools/make_expected_metrics.py
"""
import csv
import re
import sys
from collections import Counter
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
DATA = ROOT / "data"
sys.path.insert(0, str(ROOT / "tools"))

COLS = ["doc_id", "doc_title", "doc_type", "old_version", "new_version", "effective_month", "articles_old", "articles_new",
        "articles_total", "changes", "revised", "added", "deleted", "change_rate", "subscribers"]
ART = re.compile(r"^\s*제\s*(\d+)\s*조(?:\s*의\s*(\d+))?")


def read_csv(p):
    return list(csv.DictReader(open(p, encoding="utf-8-sig")))


USERS = read_csv(DATA / "users.csv")


def subscribers(doc_type):
    return "|".join(u["user_id"] for u in USERS if doc_type in u["subscribed_types"].split("|"))


def row(doc_id, title, doc_type, v_old, v_new, eff, keys_old, keys_new, types):
    c = Counter(types)
    total = len(set(keys_old) | set(keys_new))
    n = len(types)
    return [doc_id, title, doc_type, v_old, v_new, eff[:7], len(keys_old), len(keys_new), total, n,
            c["개정"], c["신설"], c["삭제"], f"{n / total * 100:.1f}", subscribers(doc_type)]


def write(path, rows):
    with open(path, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow(COLS)
        w.writerows(rows)
    print(f"wrote {path.relative_to(ROOT)} ({len(rows)} rows)")


def sample_rows():
    meta = read_csv(DATA / "documents.csv")
    expected = read_csv(DATA / "expected_changes.csv")
    rows = []
    for base in sorted({m["doc_id"].rsplit("-", 1)[0] for m in meta}):
        old, new = sorted([m for m in meta if m["doc_id"].startswith(base + "-")], key=lambda m: m["doc_id"])
        keys = []
        for m in (old, new):
            text = (DATA / "documents" / m["file_name"]).read_text(encoding="utf-8-sig")
            keys.append([f"{a}-{b or 0}" for a, b in (ART.match(l).groups() for l in text.splitlines() if ART.match(l))])
        types = [e["change_type"] for e in expected if e["doc_id"] == base]
        rows.append(row(base, new["title"], new["doc_type"], old["version"], new["version"], new["effective_date"], keys[0], keys[1], types))
    return rows


def pack_rows(module, docs):
    rows = []
    for d in docs:
        new = module.apply(d["old"], d["chg"])
        o = {module.key(a) for a in d["old"]}
        n = {module.key(a) for a in new}
        types = [r[3] for r in module.expected_rows(d["id"], d["title"], d["old"], new, d["chg"], *d.get("t", ({}, {})))]
        rows.append(row(d["id"], d["title"], d["type"], d["v"][0], d["v"][1], d["eff"][1], o, n, types))
    return rows


def main():
    write(DATA / "expected_metrics.csv", sample_rows())

    import make_test_pack as p1
    write(DATA / "test_pack" / "expected_metrics_testpack.csv", pack_rows(p1, [
        dict(id="DOC-101", title="언더라이팅 업무매뉴얼", type="업무매뉴얼", old=p1.DOC1_OLD, chg=p1.DOC1_CHG, v=("v3.0", "v4.0"), eff=("2026-01-01", "2026-10-01")),
        dict(id="DOC-102", title="보험금 지급 사무처리기준", type="사무처리기준", old=p1.DOC2_OLD, chg=p1.DOC2_CHG, v=("v2.1", "v3.0"), eff=("2026-02-01", "2026-10-01")),
        dict(id="DOC-103", title="계약심사 전결 내규", type="내규", old=p1.DOC3_OLD, chg=p1.DOC3_CHG, v=("v1.5", "v2.0"), eff=("2025-07-01", "2026-10-01"),
             t=({"제5조": p1.TABLE_OLD}, {"제5조": p1.TABLE_NEW})),
    ]))

    import make_test_pack2 as p2
    write(DATA / "test_pack2" / "expected_metrics_testpack2.csv", pack_rows(p2, [
        dict(id="DOC-201", title="보험계약 유지관리 업무매뉴얼", type="업무매뉴얼", old=p2.PDF_OLD, chg=p2.PDF_CHG, v=("v1.0", "v2.0"), eff=("2026-03-01", "2026-11-01")),
        dict(id="DOC-202", title="모집질서 준수 내규", type="내규", old=p2.WORD_OLD, chg=p2.WORD_CHG, v=("v2.0", "v2.1"), eff=("2026-01-01", "2026-11-01"),
             t=({"제6조": p2.TABLE_OLD}, {"제6조": p2.TABLE_NEW})),
        dict(id="DOC-203", title="보험료 수납 사무처리기준", type="사무처리기준", old=p2.EXCEL_OLD, chg=p2.EXCEL_CHG, v=("v1.0", "v1.1"), eff=("2026-01-01", "2026-11-01")),
        dict(id="DOC-204", title="심사위원회 운영규정", type="내규", old=p2.TEXT_OLD, chg=p2.TEXT_CHG, v=("v3.0", "v3.1"), eff=("2025-12-01", "2026-11-01")),
    ]))


if __name__ == "__main__":
    main()
