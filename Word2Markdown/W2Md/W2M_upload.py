import os
import re
import docx2txt
from docx import Document
from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P
from docx.table import Table
from docx.text.paragraph import Paragraph
from pathlib import Path
from tkinter import Tk, filedialog
from typing import List, Tuple

def extract_images(docx_path: str, images_dir: str) -> List[str]:
    """
    提取文档中的图片
    """
    text = docx2txt.process(docx_path, images_dir)
    image_files = os.listdir(images_dir)
    image_files.sort()
    return image_files

def convert_table_to_markdown(table: Table) -> str:
    """
    将表格转换为Markdown格式
    """
    rows = table.rows
    table_md = []
    for i, row in enumerate(rows):
        cells = row.cells
        row_md = "| " + " | ".join(cell.text.strip() for cell in cells) + " |"
        table_md.append(row_md)
        if i == 0:  # 添加表头分隔符
            separator = "| " + " | ".join("---" for _ in cells) + " |"
            table_md.append(separator)
    return "\n".join(table_md)

def convert_docx_to_markdown(file_path: str) -> None:
    """
    将docx文档转换为Markdown格式，并保存图片
    """
    doc = Document(file_path)

    # 提取文件名并去除扩展名
    file_name = os.path.basename(file_path)
    file_base = os.path.splitext(file_name)[0]

    # 为Markdown文件和图片创建目录
    output_dir = f"./generate_data/{file_base}"
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    # Markdown文件路径
    md_file_path = os.path.join(output_dir, "output.md")

    # 初始化Markdown内容
    md_content = []

    # 为图片创建目录
    images_dir = Path(output_dir) / "Images"
    images_dir.mkdir(parents=True, exist_ok=True)

    # 提取图片
    image_files = extract_images(file_path, str(images_dir))

    def handle_heading(text: str, level: int) -> str:
        """
        处理不同级别的标题
        """
        if level > 6:
            level = 6
        return f"{'#' * level} {text}"

    def contains_image(paragraph: Paragraph) -> bool:
        """
        检查段落中是否包含图片
        """
        for run in paragraph.runs:
            if "graphicData" in run._element.xml:
                return True
        return False

    image_index = 0
    content_started = False
    first_secondary_found = False
    primary_title = []

    for element in doc.element.body:
        if isinstance(element, CT_Tbl):
            table = Table(element, doc)
            table_md = convert_table_to_markdown(table)
            md_content.append(table_md)
        elif isinstance(element, CT_P):
            para = Paragraph(element, doc)
            text = para.text.strip()

            # 跳过空段落
            if not text and not contains_image(para):
                continue

            # 跳过第一页内容
            if not content_started:
                if re.match(r'^\d+\s*[^.\d].*$', text):
                    content_started = True
                else:
                    continue

            # 过滤包含日期格式或包含“-”的标题
            if re.match(r'^\d{4}-\d{2}-\d{2}', text) or '-' in text:
                continue

            # 如果找到第一个二级标题，将前面的文本视为一级标题
            if re.match(r'^\d+\s*[^.\d].*$', text):
                if not first_secondary_found:
                    first_secondary_found = True
                    primary_title_text = ' '.join(primary_title).strip()
                    # 只保留空格后面的文字
                    if primary_title_text and ' ' in primary_title_text:
                        primary_title_text = primary_title_text.split(' ', 1)[1].strip()
                    if primary_title_text:
                        print(f"Primary title matched: {primary_title_text}")  # 调试输出
                        md_content.append(handle_heading(primary_title_text, 1))
                level = 2
                md_content.append(handle_heading(text, level))
            elif re.match(r'^\d+\.\d+[^.]*$', text):
                match = re.match(r'^(\d+\.\d+)(.*)$', text)
                if match:
                    level = 3
                    md_content.append(handle_heading(match.group(1), level))
                    remaining_text = match.group(2).strip()
                    if remaining_text:
                        md_content.append(remaining_text)
            elif re.match(r'^\d+\.\d+\.\d+[^.]*$', text):
                match = re.match(r'^(\d+\.\d+\.\d+)(.*)$', text)
                if match:
                    level = 4
                    md_content.append(handle_heading(match.group(1), level))
                    remaining_text = match.group(2).strip()
                    if remaining_text:
                        md_content.append(remaining_text)
            elif re.match(r'^\d+\.\d+\.\d+\.\d+[^.]*$', text):
                match = re.match(r'^(\d+\.\d+\.\d+\.\d+)(.*)$', text)
                if match:
                    level = 5
                    md_content.append(handle_heading(match.group(1), level))
                    remaining_text = match.group(2).strip()
                    if remaining_text:
                        md_content.append(remaining_text)
            elif re.match(r'^\d+\.\d+\.\d+\.\d+\.\d+[^.]*$', text):
                match = re.match(r'^(\d+\.\d+\.\d+\.\d+\.\d+)(.*)$', text)
                if match:
                    level = 6
                    md_content.append(handle_heading(match.group(1), level))
                    remaining_text = match.group(2).strip()
                    if remaining_text:
                        md_content.append(remaining_text)
            elif re.match(r'^\d+\.\d+\.\d+\.\d+\.\d+\.\d+[^.]*$', text):
                match = re.match(r'^(\d+\.\d+\.\d+\.\d+\.\d+\.\d+)(.*)$', text)
                if match:
                    level = 7
                    md_content.append(handle_heading(match.group(1), level))
                    remaining_text = match.group(2).strip()
                    if remaining_text:
                        md_content.append(remaining_text)
            else:
                if not first_secondary_found:
                    primary_title.append(text)
                else:
                    md_content.append(text)

            # 如果段落中包含图片，插入图片
            if contains_image(para) and image_index < len(image_files):
                image_filename = image_files[image_index]
                md_content.append(f"![image_{image_index + 1}](./Images/{image_filename})")
                image_index += 1

    # 如果尚未添加一级标题，则添加一级标题
    if primary_title and not first_secondary_found:
        primary_title_text = ' '.join(primary_title).strip()
        # 只保留空格后面的文字
        if primary_title_text and ' ' in primary_title_text:
            primary_title_text = primary_title_text.split(' ', 1)[1].strip()
        if primary_title_text:
            print(f"Primary title matched: {primary_title_text}")  # 调试输出
            md_content.append(handle_heading(primary_title_text, 1))
    elif not primary_title:
        print("No primary title matched")  # 调试输出

    # 将Markdown内容写入文件
    with open(md_file_path, 'w', encoding='utf-8') as md_file:
        md_file.write('\n\n'.join(md_content))


def select_files() -> Tuple[str, ...]:
    """
    弹出文件选择对话框，允许用户选择多个DOCX文件
    """
    root = Tk()
    root.withdraw()  # 隐藏主窗口
    file_paths = filedialog.askopenfilenames(
        title="选择DOCX文件",
        filetypes=(("DOCX文件", "*.docx"), ("所有文件", "*.*"))
    )
    return file_paths


if __name__ == "__main__":
    file_paths = select_files()
    if file_paths:
        for file_path in file_paths:
            convert_docx_to_markdown(file_path)
    else:
        print("未选择任何文件。")
