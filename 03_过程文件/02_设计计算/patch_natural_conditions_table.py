from copy import deepcopy
from pathlib import Path

from docx import Document
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt


DOCX_PATH = Path(
    r"D:\毕业论文\04_最终成品\01_毕业设计说明书"
    r"\张淑鑫_浙江平湖二级油库工艺设计_毕业设计初稿_计算与规范完善版_2026-07-28.docx"
)
OUTPUT_PATH = DOCX_PATH


ROWS = [
    ("温度", "年平均气温", "16.3 ℃", "16.3 ℃", "政府环评PDF第166页（报告书第162页）", "油品物性、设备环境条件"),
    ("温度", "最热月月平均气温", "26.9 ℃", "26.9 ℃", "政府环评PDF第166页（原称“月平均最高气温”）", "夏季运行温度边界"),
    ("温度", "最冷月月平均气温", "5.8 ℃", "5.8 ℃", "政府环评PDF第166页（原称“月平均最低气温”）", "冬季运行温度边界"),
    ("温度", "绝对最高气温", "39.1 ℃", "39.1 ℃", "政府环评PDF第166页（原称“历年极端最高气温”）", "材料、仪表耐温"),
    ("温度", "绝对最低气温", "-9.3 ℃", "-9.3 ℃", "政府环评PDF第166页（原称“历年极端最低气温”）", "防冻与低温适用性"),
    (
        "温度",
        "最低平均温度",
        "5.8 ℃",
        "5.8 ℃",
        "政府环评PDF第166页；公开资料未另列独立指标，按“月平均最低气温”采用",
        "最低平均运行温度",
    ),
    ("降雨量", "全年平均降雨量", "1269.7 mm", "1269.7 mm", "政府环评PDF第166页（原称“年平均降水量”）", "场地排水系统"),
    ("降雨量", "日最大降雨量", "276.4 mm", "276.4 mm", "政府环评PDF第166页（原称“一日最大降水量”）", "事故水池雨水分量"),
    ("降雨量", "平均降雨天数", "140.6 d", "140.6 d", "政府环评PDF第166页（原称“年平均降水日数”）", "排水与码头作业组织"),
    ("主导风向和风速", "主导风向", "E～SE，累计频率30%", "E～SE", "政府环评PDF第166页", "总图和管理区方位"),
    ("主导风向和风速", "年平均风速", "3.2 m/s", "3.2 m/s", "政府环评PDF第167页（报告书第163页）", "运行环境与通风"),
    ("主导风向和风速", "最大风速", "31.7 m/s", "31.7 m/s", "政府环评PDF第167页（原称“极大风速”）", "抗风与码头停工条件"),
    ("风荷载", "乍浦50年重现期基本风压", "0.45 kN/m²", "0.45 kN/m²", "《浙江省基本风压资料》PDF第43页（资料第37页）", "储罐抗风稳定性"),
    ("湿度", "年平均相对湿度", "80%", "80%", "政府环评PDF第167页（报告书第163页）", "沿海防腐与电气选型"),
    ("天气", "年平均雾日", "35 d", "35 d", "政府环评PDF第167页（报告书第163页）", "码头可用时间"),
    ("天气", "年平均雷暴日", "28 d", "28 d", "政府环评PDF第168页（报告书第164页）", "防雷和作业管理"),
    ("地震", "地震基本烈度", "Ⅵ度", "Ⅵ度", "政府环评PDF第177页（报告书第173页）", "储罐、管线抗震"),
]


def set_cell_text(cell, text, bold=False):
    cell.text = ""
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.line_spacing = 1
    run = p.add_run(text)
    run.bold = bold
    run.font.name = "宋体"
    run._element.rPr.rFonts.set(qn("w:eastAsia"), "宋体")
    run.font.size = Pt(9)
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER


def main():
    document = Document(DOCX_PATH)

    # Remove a historical equation fragment accidentally embedded in the
    # cover-title paragraph while retaining the visible title text “设计”.
    cover = document.paragraphs[0]
    for math_node in list(cover._p.findall(".//" + qn("m:oMath"))):
        parent = math_node.getparent()
        parent.remove(math_node)

    # An empty Heading 1 immediately before “1 设计总论” inherits
    # pageBreakBefore and produces a completely blank page after the TOC.
    for index, paragraph in enumerate(document.paragraphs):
        if paragraph.text.strip() != "1 设计总论" or index == 0:
            continue
        previous = document.paragraphs[index - 1]
        if not previous.text.strip() and previous.style.name == "Heading 1":
            previous._p.getparent().remove(previous._p)
        break

    target = None
    for table in document.tables:
        if not table.rows:
            continue
        header = [cell.text.strip() for cell in table.rows[0].cells]
        if header[:3] == ["类别", "参数", "原始统计值"]:
            target = table
            break
    if target is None:
        raise RuntimeError("未找到自然条件表")

    header_tr = target.rows[0]._tr
    row_template = deepcopy(target.rows[1]._tr)
    for tr in list(target._tbl.tr_lst)[1:]:
        target._tbl.remove(tr)

    for values in ROWS:
        new_tr = deepcopy(row_template)
        target._tbl.append(new_tr)
        row = target.rows[-1]
        for cell, value in zip(row.cells, values):
            set_cell_text(cell, value)

    for cell in target.rows[0].cells:
        original = cell.text.strip()
        set_cell_text(cell, original, bold=True)

    # Keep the header repeated if Word splits the expanded table over pages.
    tr_pr = header_tr.get_or_add_trPr()
    if tr_pr.find(qn("w:tblHeader")) is None:
        tbl_header = OxmlElement("w:tblHeader")
        tbl_header.set(qn("w:val"), "true")
        tr_pr.append(tbl_header)

    document.save(OUTPUT_PATH)
    print(f"updated: {OUTPUT_PATH}")
    print(f"natural-condition rows: {len(ROWS)}")


if __name__ == "__main__":
    main()
