from __future__ import annotations

import re
import sys
from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from latex2mathml.converter import convert as latex_to_mathml
from lxml import etree


XSL_PATH = Path(r"C:\Program Files\Microsoft Office\root\Office16\MML2OMML.XSL")
MARKER_RE = re.compile(r"^\[\[NUMEQ\|(.+)\]\]$")


def matching_paren(text: str, start: int) -> int:
    depth = 0
    for index in range(start, len(text)):
        if text[index] == "(":
            depth += 1
        elif text[index] == ")":
            depth -= 1
            if depth == 0:
                return index
    return -1


def replace_balanced_fractions(text: str) -> str:
    pos = 0
    output: list[str] = []
    while pos < len(text):
        if text[pos] == "(":
            end_num = matching_paren(text, pos)
            if end_num > pos and text[end_num + 1 : end_num + 3] == "/(":
                den_start = end_num + 2
                end_den = matching_paren(text, den_start)
                if end_den > den_start:
                    num = replace_balanced_fractions(text[pos + 1 : end_num])
                    den = replace_balanced_fractions(text[den_start + 1 : end_den])
                    output.append(r"\frac{" + num + "}{" + den + "}")
                    pos = end_den + 1
                    continue
        output.append(text[pos])
        pos += 1
    return "".join(output)


def linear_to_latex(text: str) -> str:
    text = replace_balanced_fractions(text)
    text = re.sub(r"√\(([^()]*)\)", r"\\sqrt{\1}", text)
    text = re.sub(r"_\(([^()]*)\)", r"_{\1}", text)
    text = re.sub(r"_([A-Za-z0-9Σ]+)", r"_{\1}", text)
    text = re.sub(r"\^\(([^()]*)\)", r"^{\1}", text)
    text = re.sub(r"\^([0-9.]+)", r"^{\1}", text)
    text = text.replace("×", r"\times ")
    text = text.replace("π", r"\pi ")
    text = text.replace("ρ", r"\rho ")
    text = text.replace("η", r"\eta ")
    text = text.replace("λ", r"\lambda ")
    text = text.replace("ν", r"\nu ")
    text = text.replace("ψ", r"\psi ")
    text = text.replace("μ", r"\mu ")
    text = text.replace("ω", r"\omega ")
    text = text.replace("Δ", r"\Delta ")
    text = text.replace("Σ", r"\Sigma ")
    text = text.replace("⌈", r"\lceil ")
    text = text.replace("⌉", r"\rceil ")
    text = text.replace("′", r"^{\prime}")
    return text


def main(path: Path) -> None:
    document = Document(path)
    transform = etree.XSLT(etree.parse(str(XSL_PATH)))
    converted = 0
    failures: list[str] = []

    for paragraph in document.paragraphs:
        match = MARKER_RE.match(paragraph.text.strip())
        if not match:
            continue
        linear = match.group(1)
        latex = linear_to_latex(linear)
        try:
            mathml = etree.fromstring(latex_to_mathml(latex).encode("utf-8"))
            omml = transform(mathml).getroot()
        except Exception as exc:
            failures.append(f"{linear}: {exc}")
            continue

        p = paragraph._p
        for child in list(p):
            if child.tag.endswith("}pPr"):
                continue
            p.remove(child)
        p.append(etree.fromstring(etree.tostring(omml)))
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        paragraph.paragraph_format.first_line_indent = 0
        paragraph.paragraph_format.left_indent = 0
        paragraph.paragraph_format.right_indent = 0
        converted += 1

    if failures:
        raise RuntimeError("Equation conversion failures:\n" + "\n".join(failures))
    document.save(path)
    print(f"converted={converted}")


if __name__ == "__main__":
    if len(sys.argv) != 2:
        raise SystemExit("usage: convert_markers_to_omml.py DOCX")
    main(Path(sys.argv[1]))
