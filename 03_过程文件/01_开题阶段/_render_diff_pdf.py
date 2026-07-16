from pathlib import Path
import pypdfium2 as pdfium
from PIL import Image, ImageDraw

base=Path(r"D:\毕业论文\03_过程文件\01_开题阶段\老师修改版对比")
pdf_path=base/"差异对照版.pdf"
out=base/"差异对照版_渲染"; out.mkdir(parents=True,exist_ok=True)
pdf=pdfium.PdfDocument(str(pdf_path)); thumbs=[]
for i in range(len(pdf)):
    img=pdf[i].render(scale=1.6).to_pil().convert("RGB")
    img.save(out/f"page-{i+1:02d}.png")
    t=img.copy(); t.thumbnail((380,492)); thumbs.append(t)
cols=3; gap=20; w=380; h=522; rows=(len(thumbs)+cols-1)//cols
sheet=Image.new("RGB",(cols*w+(cols+1)*gap,rows*h+(rows+1)*gap),"#d8d8d8")
d=ImageDraw.Draw(sheet)
for i,t in enumerate(thumbs):
    x=gap+(i%cols)*(w+gap); y=gap+(i//cols)*(h+gap)
    sheet.paste(t,(x+(w-t.width)//2,y)); d.text((x,y+496),f"Page {i+1}",fill="black")
sheet.save(base/"差异对照版_contact-sheet.png")
print("pages",len(pdf),base/"差异对照版_contact-sheet.png")
