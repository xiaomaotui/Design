from pathlib import Path
import pypdfium2 as pdfium
from PIL import Image, ImageDraw

pdf_path = Path(r"D:\毕业论文\04_最终成品\01_开题报告\张淑鑫_浙江平湖油库工艺设计_开题报告.pdf")
out = Path(r"D:\毕业论文\03_过程文件\05_PDF渲染缓存\开题报告成品")
out.mkdir(parents=True, exist_ok=True)
pdf = pdfium.PdfDocument(str(pdf_path))
thumbs = []
for i in range(len(pdf)):
    img = pdf[i].render(scale=1.6).to_pil().convert("RGB")
    path = out / f"page-{i+1:02d}.png"
    img.save(path)
    t = img.copy(); t.thumbnail((420, 594))
    thumbs.append(t)

cols = 3; gap = 24; label_h = 28
rows = (len(thumbs) + cols - 1) // cols
sheet = Image.new("RGB", (cols * 420 + (cols + 1) * gap, rows * (594 + label_h) + (rows + 1) * gap), "#d9d9d9")
d = ImageDraw.Draw(sheet)
for i, t in enumerate(thumbs):
    x = gap + (i % cols) * (420 + gap)
    y = gap + (i // cols) * (594 + label_h + gap)
    sheet.paste(t, (x + (420-t.width)//2, y))
    d.text((x, y + 598), f"Page {i+1}", fill="black")
sheet.save(out / "contact-sheet.png")
print(f"pages={len(pdf)} sheet={out / 'contact-sheet.png'}")
