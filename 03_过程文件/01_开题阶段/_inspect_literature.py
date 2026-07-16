from __future__ import annotations

import hashlib
import json
from pathlib import Path

from pypdf import PdfReader


ROOT = Path(r"D:\毕业论文")
SOURCE = ROOT / "02_参考文献原件" / "04_开题报告文献"
CACHE = ROOT / "03_过程文件" / "01_开题阶段" / "文献阅读卡" / "全文提取"
CACHE.mkdir(parents=True, exist_ok=True)

records = []
for pdf in sorted(SOURCE.glob("0[12]_*/*.pdf")):
    raw = pdf.read_bytes()
    reader = PdfReader(str(pdf))
    page_texts = []
    errors = []
    for idx, page in enumerate(reader.pages, 1):
        try:
            page_texts.append(page.extract_text() or "")
        except Exception as exc:  # retain the PDF even if one page has broken text encoding
            page_texts.append("")
            errors.append(f"p{idx}: {exc}")
    text = "\n\n".join(page_texts)
    out = CACHE / f"{pdf.stem}.txt"
    out.write_text(text, encoding="utf-8")
    records.append(
        {
            "file": str(pdf.relative_to(ROOT)),
            "bytes": len(raw),
            "sha256": hashlib.sha256(raw).hexdigest(),
            "pages": len(reader.pages),
            "text_chars": len(text),
            "title_metadata": str((reader.metadata or {}).get("/Title", "")),
            "author_metadata": str((reader.metadata or {}).get("/Author", "")),
            "extract_errors": errors,
            "text_file": str(out.relative_to(ROOT)),
        }
    )

summary = CACHE.parent / "文献PDF核验结果.json"
summary.write_text(json.dumps(records, ensure_ascii=False, indent=2), encoding="utf-8")
print(json.dumps(records, ensure_ascii=False, indent=2))
