from __future__ import annotations

import math
from pathlib import Path

import pypdfium2 as pdfium
from PIL import Image, ImageDraw, ImageFont


ROOT = Path(r"D:\毕业论文")
SOURCE = ROOT / "02_参考文献原件" / "04_开题报告文献"
OUT = ROOT / "03_过程文件" / "05_PDF渲染缓存" / "开题报告15篇文献"
OUT.mkdir(parents=True, exist_ok=True)

thumb_w = 430
gap = 18
label_h = 28
cols = 3

for pdf_path in sorted(SOURCE.glob("0[12]_*/*.pdf")):
    pdf = pdfium.PdfDocument(str(pdf_path))
    thumbs = []
    for i in range(len(pdf)):
        page = pdf[i]
        bitmap = page.render(scale=0.9)
        image = bitmap.to_pil().convert("RGB")
        h = round(image.height * thumb_w / image.width)
        image = image.resize((thumb_w, h), Image.Resampling.LANCZOS)
        cell = Image.new("RGB", (thumb_w, h + label_h), "white")
        cell.paste(image, (0, label_h))
        draw = ImageDraw.Draw(cell)
        draw.text((8, 5), f"Page {i + 1}", fill="black")
        thumbs.append(cell)
    rows = math.ceil(len(thumbs) / cols)
    max_h = max(x.height for x in thumbs)
    sheet = Image.new(
        "RGB",
        (cols * thumb_w + (cols + 1) * gap, rows * max_h + (rows + 1) * gap),
        "#d9dde3",
    )
    for i, image in enumerate(thumbs):
        x = gap + (i % cols) * (thumb_w + gap)
        y = gap + (i // cols) * (max_h + gap)
        sheet.paste(image, (x, y))
    sheet.save(OUT / f"{pdf_path.stem}_contact.png", optimize=True)

print(f"rendered {len(list(OUT.glob('*_contact.png')))} contact sheets to {OUT}")
