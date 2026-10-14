"""data/ 폴더의 더미 데이터를 app/src.html 에 넣어 public/index.html 을 만든다.

사용법: python3 tools/build.py
- public/index.html : 브라우저·Vercel에서 바로 여는 완성본 (오프라인 동작, 엑셀 내보내기만 CDN 필요)
"""
import json
import shutil
import zipfile
from datetime import datetime, timedelta, timezone
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
DATA = ROOT / "data"


def read(p: Path) -> str:
    return p.read_text(encoding="utf-8-sig")


def main() -> None:
    sample = {
        "files": [
            {"name": p.name, "text": read(p)}
            for p in sorted((DATA / "documents").glob("*.txt"))
        ],
        "csv": {
            name: read(DATA / f"{name}.csv")
            for name in ("documents", "users", "classification_keywords", "expected_changes", "change_history", "expected_metrics")
        },
        "mails": [
            {"name": p.name, "text": read(p)}
            for p in sorted((DATA / "mail_samples").glob("*.txt"))
        ],
    }
    # 엑셀 내려받기는 받은 양식(changes_template.xlsx)의 서식을 그대로 쓴다
    with zipfile.ZipFile(DATA / "changes_template.xlsx") as z:
        sample["xlsxTemplate"] = {
            "styles": z.read("xl/styles.xml").decode("utf-8"),
            "theme": z.read("xl/theme/theme1.xml").decode("utf-8"),
            "sheet": z.read("xl/worksheets/sheet1.xml").decode("utf-8"),
        }
    # 운영 검증용 시험 자료: public/testdata/ 에 복사하고 목록(manifest.json)을 만든다 → 앱의 '검증 실행'이 불러 쓴다
    testdata = ROOT / "public" / "testdata"
    if testdata.exists():
        shutil.rmtree(testdata)
    sets = []
    for key, label, folder in (("pack1", "시험용 1차", "test_pack"), ("pack2", "시험용 2차", "test_pack2")):
        src_dir = DATA / folder
        if not src_dir.exists():
            continue
        shutil.copytree(src_dir, testdata / folder)
        names = sorted(p.name for p in src_dir.iterdir())
        pick = lambda prefix: next(n for n in names if n.startswith(prefix))
        sets.append({"key": key, "label": label, "dir": f"testdata/{folder}/",
                     "files": [n for n in names if n.rsplit(".", 1)[-1] in ("txt", "docx", "pdf", "xlsx") and n != "README.txt"],
                     "meta": pick("documents_"), "expected": pick("expected_changes_"), "metrics": pick("expected_metrics_")})
    (testdata / "manifest.json").write_text(json.dumps({"sets": sets}, ensure_ascii=False, indent=1), encoding="utf-8")
    kst = datetime.now(timezone(timedelta(hours=9)))
    sample["build"] = kst.strftime("%Y-%m-%d %H:%M")

    # 검증 리포트 페이지: app/report.html(본문) → public/report.html
    rp = ROOT / "app" / "report.html"
    if rp.exists():
        (ROOT / "public" / "report.html").write_text(
            '<!doctype html><html lang="ko"><head><meta charset="utf-8">'
            '<meta name="viewport" content="width=device-width,initial-scale=1,viewport-fit=cover"></head><body>'
            + read(rp) + "</body></html>", encoding="utf-8")

    payload = json.dumps(sample, ensure_ascii=False).replace("</", "<\\/")
    src = read(ROOT / "app" / "src.html")
    out = src.replace("/*__SAMPLE_DATA__*/null", payload)
    if out == src:
        raise SystemExit("placeholder not found in app/src.html")
    (ROOT / "public" / "index.html").write_text(out, encoding="utf-8")
    print(f"wrote public/index.html ({len(out):,} bytes)")


if __name__ == "__main__":
    main()
