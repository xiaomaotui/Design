from pathlib import Path
import sys
from PIL import Image, ImageDraw

source = Path(sys.argv[1]) if len(sys.argv) > 1 else Path(r"D:\毕业论文\tmp\render_class2_draft_v2")
pages = sorted(source.glob("page-*.png"))
out = source / "contact_sheets"
out.mkdir(exist_ok=True)

for start in range(0, len(pages), 4):
    batch = pages[start:start + 4]
    opened = [Image.open(p).convert("RGB") for p in batch]
    thumb_w = 1050
    thumbs = []
    for image in opened:
        h = round(image.height * thumb_w / image.width)
        thumbs.append(image.resize((thumb_w, h)))
    thumb_h = max(i.height for i in thumbs)
    sheet = Image.new("RGB", (thumb_w * 2 + 60, thumb_h * 2 + 100), "white")
    draw = ImageDraw.Draw(sheet)
    for j, image in enumerate(thumbs):
        x = 20 + (j % 2) * (thumb_w + 20)
        y = 35 + (j // 2) * (thumb_h + 40)
        sheet.paste(image, (x, y))
        draw.text((x, y - 24), pages[start + j].stem, fill="black")
    sheet.save(out / f"sheet-{start + 1:02d}-{start + len(batch):02d}.jpg", quality=88)

print(f"sheets={len(list(out.glob('sheet-*.jpg')))}")
