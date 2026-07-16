from pathlib import Path
import pypdfium2 as pdfium
from PIL import Image, ImageDraw

base = Path(r"D:\毕业论文\03_过程文件\01_开题阶段\老师修改版对比")
for stem in ["原版", "老师修改版"]:
    pdf = pdfium.PdfDocument(str(base / f"{stem}.pdf"))
    out = base / f"{stem}_渲染"
    out.mkdir(parents=True, exist_ok=True)
    thumbs=[]
    for i in range(len(pdf)):
        img = pdf[i].render(scale=1.25).to_pil().convert("RGB")
        img.save(out / f"page-{i+1:02d}.png")
        t=img.copy(); t.thumbnail((260,368)); thumbs.append(t)
    cols=5 if len(pdf)>20 else 3; gap=16; w=260; h=392
    rows=(len(thumbs)+cols-1)//cols
    sheet=Image.new("RGB",(cols*w+(cols+1)*gap,rows*h+(rows+1)*gap),"#d8d8d8")
    d=ImageDraw.Draw(sheet)
    for i,t in enumerate(thumbs):
        x=gap+(i%cols)*(w+gap); y=gap+(i//cols)*(h+gap)
        sheet.paste(t,(x+(w-t.width)//2,y)); d.text((x,y+370),f"Page {i+1}",fill="black")
    sheet.save(base/f"{stem}_contact-sheet.png")
    print(stem,len(pdf),base/f"{stem}_contact-sheet.png")
