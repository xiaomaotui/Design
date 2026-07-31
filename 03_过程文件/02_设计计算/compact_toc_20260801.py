from pathlib import Path

from docx import Document
from docx.shared import Pt


path = Path(r"D:\毕业论文\04_最终成品\01_毕业设计说明书\张淑鑫_浙江平湖二级油库工艺设计_毕业设计初稿_绪论与研究动态补充版_2026-08-01.docx")
document = Document(path)
document.styles["toc 1"].paragraph_format.space_before = Pt(0)
document.styles["toc 1"].paragraph_format.line_spacing = 1.5
document.save(path)
