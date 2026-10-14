"""data/ 폴더의 더미 데이터를 app/src.html 에 넣어 public/index.html 을 만든다.

사용법: python3 tools/build.py
- public/index.html : 브라우저·Vercel에서 바로 여는 완성본 (오프라인 동작, 엑셀 내보내기만 CDN 필요)
"""
import json
import zipfile
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
            for name in ("documents", "users", "classification_keywords", "expected_changes", "change_history")
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
    payload = json.dumps(sample, ensure_ascii=False).replace("</", "<\\/")
    src = read(ROOT / "app" / "src.html")
    out = src.replace("/*__SAMPLE_DATA__*/null", payload)
    if out == src:
        raise SystemExit("placeholder not found in app/src.html")
    (ROOT / "public" / "index.html").write_text(out, encoding="utf-8")
    print(f"wrote public/index.html ({len(out):,} bytes)")


if __name__ == "__main__":
    main()
