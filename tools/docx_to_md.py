"""
文件用途：将早期 Word 初稿转换为 Markdown，曾用于建立协作主稿。
输入：根目录“医影知微_科技说明_初版.docx”，并检查选用的 assets/ 问答截图。
输出：覆盖同名 .md；按 Word 正文顺序转换段落、标题、表格和两张图片，并提取流程图。
依赖：python-docx；路径以脚本所在位置确定项目根目录。
当前状态：一次性初始化工具；后续以 Markdown 为主，不应再用旧 Word 覆盖协作稿。
"""

from pathlib import Path
from zipfile import ZipFile

from docx import Document
from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P
from docx.table import Table
from docx.text.paragraph import Paragraph


ROOT = Path(__file__).resolve().parent.parent
SOURCE = ROOT / "医影知微_科技说明_初版.docx"
TARGET = ROOT / "医影知微_科技说明_初版.md"
FLOW = ROOT / "assets" / "tech_flow.png"
SCREENSHOT = ROOT / "assets" / "b95945ac89bf5186170b732fca68081c.jpg"


def cell_text(cell):
    return cell.text.strip().replace("\n", "<br>").replace("|", r"\|")


def table_markdown(table):
    rows = [[cell_text(cell) for cell in row.cells] for row in table.rows]
    header = "| " + " | ".join(rows[0]) + " |"
    rule = "| " + " | ".join("---" for _ in rows[0]) + " |"
    body = ["| " + " | ".join(row) + " |" for row in rows[1:]]
    return "\n".join([header, rule, *body])


def main():
    if not SCREENSHOT.is_file():
        raise FileNotFoundError(SCREENSHOT)
    with ZipFile(SOURCE) as archive:
        flow_bytes = archive.read("word/media/image1.png")
    if FLOW.exists() and FLOW.read_bytes() != flow_bytes:
        raise ValueError(f"Existing flow image differs: {FLOW}")
    if not FLOW.exists():
        FLOW.write_bytes(flow_bytes)

    doc = Document(SOURCE)
    blocks = ["<!-- 协作主稿；最终提交时由本 Markdown 导出 DOCX。 -->"]
    image_links = [
        "![作品工作流程示意](assets/tech_flow.png)",
        "![影像问答界面局部运行截图](assets/b95945ac89bf5186170b732fca68081c.jpg)",
    ]
    image_index = 0
    title_index = 0
    for element in doc.element.body.iterchildren():
        if isinstance(element, CT_Tbl):
            blocks.append(table_markdown(Table(element, doc)))
            continue
        if not isinstance(element, CT_P):
            continue
        paragraph = Paragraph(element, doc)
        if element.xpath(".//w:drawing"):
            blocks.append(image_links[image_index])
            image_index += 1
            continue
        content = paragraph.text.strip()
        if not content:
            continue
        style = paragraph.style.name
        if style == "Title":
            blocks.append(f"# {content}" if title_index == 0 else f"**{content}**")
            title_index += 1
        elif style == "Section":
            blocks.append(f"## {content}")
        elif style == "Subsection":
            blocks.append(f"### {content}")
        elif style == "FigureCaption":
            blocks.append(f"*{content}*")
        else:
            blocks.append(content)

    if image_index != len(image_links):
        raise ValueError(f"Expected {len(image_links)} images, found {image_index}")
    markdown = "\n\n".join(blocks) + "\n"
    TARGET.write_text(markdown, encoding="utf-8")
    print(f"created={TARGET}")
    print(f"paragraphs={len(doc.paragraphs)} tables={len(doc.tables)} images={image_index}")
    print(f"markdown_characters={len(markdown)}")


if __name__ == "__main__":
    main()
