# -*- coding: utf-8 -*-
import os
import re
import uuid
import json
import base64
from typing import List, Optional, Union
from urllib.parse import urlparse, quote, unquote
from datetime import datetime, timedelta

import requests
from fastapi import FastAPI, HTTPException, Request
from fastapi.responses import FileResponse, JSONResponse, Response
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel

from docx import Document
from docx.text.paragraph import Paragraph
from docx.table import Table
from docx.oxml.text.paragraph import CT_P
from docx.oxml.table import CT_Tc, CT_Tbl
from docx.oxml.shared import qn
from docx.oxml import OxmlElement
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL

HOST = os.environ.get("HOST", "0.0.0.0")
PORT = int(os.environ.get("PORT", 8000))
BASE_URL = os.environ.get("BASE_URL", f"http://localhost:{PORT}")

file_store = {}

app = FastAPI(
    title="制式文档生成API",
    description="Word模板占位符替换服务",
    version="2.0.0-md-table-generic"
)

# 全局异常捕获
@app.exception_handler(Exception)
async def global_exception_handler(request: Request, exc: Exception):
    print("❌ 全局错误：", exc)
    return JSONResponse(
        status_code=500,
        content={"detail": f"服务器错误: {str(exc)}"}
    )

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


class GenerateRequest(BaseModel):
    template_file_url: str = "https://raw.githubusercontent.com/3293203411/docx-generator/main/gaoqilixiangmoban.docx"
    text_keys: Union[List[str], str]
    text_values: Union[List[str], str]
    filename: Optional[str] = None


def parse_json_param(param):
    if param is None:
        return []
    if isinstance(param, list):
        return [str(item) for item in param]
    if isinstance(param, str):
        try:
            parsed = json.loads(param)
            if isinstance(parsed, list):
                return [str(item) for item in parsed]
            return [str(parsed)]
        except json.JSONDecodeError:
            return [param]
    return [str(param)]


def download_file(url: str) -> bytes:
    try:
        headers = {'User-Agent': 'Mozilla/5.0'}
        response = requests.get(url, headers=headers, timeout=60)
        response.raise_for_status()
        return response.content
    except requests.exceptions.RequestException as e:
        raise HTTPException(status_code=400, detail=f"下载模板文件失败: {str(e)}")


def merge_runs_in_paragraph(paragraph):
    if not paragraph.runs:
        return {"text": "", "runs": [], "formatting": None}
    full_text = ""
    for run in paragraph.runs:
        full_text += run.text if run.text else ""
    first_run = paragraph.runs[0]
    formatting = {
        "bold": first_run.bold,
        "italic": first_run.italic,
        "underline": first_run.underline,
        "font.name": first_run.font.name,
        "font.size": first_run.font.size,
        "font.color.rgb": first_run.font.color.rgb if first_run.font.color and first_run.font.color.rgb else None,
    }
    return {"text": full_text, "runs": list(paragraph.runs), "formatting": formatting}


def replace_placeholders_in_text(text, keys, values):
    result = text
    for key, value in zip(keys, values):
        pattern = r'\{\{\s*' + re.escape(key) + r'\s*\}\}'
        result = re.sub(pattern, str(value), result)
    return result


def apply_formatting_to_run(run, formatting):
    if formatting is None:
        return
    if formatting.get("bold") is not None:
        run.bold = formatting["bold"]
    if formatting.get("italic") is not None:
        run.italic = formatting["italic"]
    if formatting.get("underline") is not None:
        run.underline = formatting["underline"]
    if formatting.get("font.name"):
        run.font.name = formatting["font.name"]
    if formatting.get("font.size"):
        run.font.size = formatting["font.size"]
    if formatting.get("font.color.rgb"):
        run.font.color.rgb = formatting["font.color.rgb"]


def parse_colored_segments(value):
    segments = []
    pattern = re.compile(r'<red>(.*?)</red>', re.DOTALL)
    last_end = 0
    for match in pattern.finditer(value):
        if match.start() > last_end:
            segments.append((value[last_end:match.start()], False))
        segments.append((match.group(1), True))
        last_end = match.end()
    if last_end < len(value):
        segments.append((value[last_end:], False))
    if not segments:
        segments = [(value, False)]
    return segments


def has_color_tag(text):
    return '<red>' in text and '</red>' in text


def set_cell_border(cell):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBorders = OxmlElement('w:tcBorders')
    for border_name in ['top', 'left', 'bottom', 'right']:
        border = OxmlElement(f'w:{border_name}')
        border.set(qn('w:val'), 'single')
        border.set(qn('w:sz'), '4')
        border.set(qn('w:space'), '0')
        border.set(qn('w:color'), '000000')
        tcBorders.append(border)
    tcPr.append(tcBorders)


def set_cell_background(cell, color):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), color)
    tcPr.append(shd)


def set_cell_text(cell, text, bold=False, font_size=10.5, align=WD_ALIGN_PARAGRAPH.CENTER):
    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    para = cell.paragraphs[0]
    para.alignment = align
    run = para.add_run(text)
    run.bold = bold
    run.font.size = Pt(font_size)
    run.font.name = '宋体'
    run._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')


# ======================= 🆕 通用 Markdown 表格解析引擎 =======================
def is_markdown_table(text):
    """判断字符串是否符合标准 Markdown 表格特征"""
    if not isinstance(text, str) or '|' not in text:
        return False
    lines = [l.strip() for l in text.strip().splitlines() if l.strip()]
    if len(lines) < 3:
        return False
    # 验证第二行是否为分隔行（仅允许 | - : 空格）
    sep_line = lines[1]
    if '|' not in sep_line:
        return False
    return all(c in '|-: ' for c in sep_line)


def parse_markdown_table(text):
    """解析 Markdown 表格，返回 (表头列表, 对齐方式列表, 数据行列表)"""
    lines = [l.strip() for l in text.strip().splitlines() if l.strip()]
    
    def split_row(line):
        # 去除首尾多余的 | 后分割
        return [cell.strip() for cell in line.strip().strip('|').split('|')]

    headers = split_row(lines[0])
    sep_parts = split_row(lines[1])
    
    # 解析对齐方式
    alignments = []
    for part in sep_parts:
        part = part.strip()
        if part.startswith(':') and part.endswith(':'):
            alignments.append(WD_ALIGN_PARAGRAPH.CENTER)
        elif part.endswith(':'):
            alignments.append(WD_ALIGN_PARAGRAPH.RIGHT)
        else:
            alignments.append(WD_ALIGN_PARAGRAPH.LEFT)
            
    data_rows = [split_row(line) for line in lines[2:]]
    return headers, alignments, data_rows


def replace_placeholder_with_markdown_table(doc, placeholder, md_text):
    """查找占位符段落，替换为 Markdown 渲染的 Word 表格"""
    target_para = None
    for element in doc.element.body:
        if isinstance(element, CT_P):
            para = Paragraph(element, doc)
            if placeholder in "".join(run.text for run in para.runs):
                target_para = para
                break
    if target_para is None:
        return False

    headers, alignments, data_rows = parse_markdown_table(md_text)
    if not headers or not data_rows:
        return False

    num_cols = len(headers)
    table = doc.add_table(rows=len(data_rows) + 1, cols=num_cols)
    table.style = 'Table Grid'

    # 1. 填充表头
    for c_idx, header_text in enumerate(headers):
        cell = table.cell(0, c_idx)
        set_cell_border(cell)
        set_cell_background(cell, 'E6E6E6')
        align = alignments[c_idx] if c_idx < len(alignments) else WD_ALIGN_PARAGRAPH.LEFT
        set_cell_text(cell, header_text, bold=True, align=align)

    # 2. 填充数据行（自动补齐/截断，防止越界）
    for r_idx, row_data in enumerate(data_rows, start=1):
        row_data = (row_data + [''] * num_cols)[:num_cols]
        for c_idx, cell_text in enumerate(row_data):
            cell = table.cell(r_idx, c_idx)
            set_cell_border(cell)
            if r_idx % 2 == 0:  # 斑马纹
                set_cell_background(cell, 'F5F5F5')
            align = alignments[c_idx] if c_idx < len(alignments) else WD_ALIGN_PARAGRAPH.LEFT
            set_cell_text(cell, cell_text, bold=False, align=align)

    # 3. 插入表格并移除原占位符
    target_para._element.addprevious(table._tbl)
    target_para._element.getparent().remove(target_para._element)
    return True
# ================================================================================


def build_research_table(doc, table_data):
    headers = ['序号', '原计划研发内容', '实际完成情况', '完成度', '未完成原因及说明']
    table = doc.add_table(rows=1 + len(table_data), cols=len(headers))
    table.style = 'Table Grid'
    col_widths = [Inches(0.8), Inches(2.5), Inches(1.5), Inches(1.0), Inches(2.0)]
    for col_idx, width in enumerate(col_widths):
        for row in table.rows:
            row.cells[col_idx].width = width
    for col_idx, header in enumerate(headers):
        cell = table.cell(0, col_idx)
        set_cell_background(cell, 'D9E1F2')
        set_cell_border(cell)
        set_cell_text(cell, header, bold=True)
    for row_idx, row_data in enumerate(table_data):
        for col_idx, value in enumerate(row_data):
            cell = table.cell(row_idx + 1, col_idx)
            set_cell_border(cell)
            if row_idx % 2 == 1:
                set_cell_background(cell, 'F2F2F2')
            set_cell_text(cell, value)
    return table


def replace_placeholder_with_table(doc, placeholder, table_data):
    target_para = None
    body = doc.element.body
    for element in body:
        if isinstance(element, CT_P):
            para = Paragraph(element, doc)
            full_text = "".join(run.text for run in para.runs)
            if placeholder in full_text:
                target_para = para
                break
    if target_para is None:
        return False
    tmp_doc = Document()
    tmp_table = build_research_table(tmp_doc, table_data)
    tbl_element = tmp_table._tbl
    target_para._element.addprevious(tbl_element)
    parent = target_para._element.getparent()
    parent.remove(target_para._element)
    return True


def process_paragraph(paragraph, keys, values):
    if not paragraph.runs:
        return
    para_info = merge_runs_in_paragraph(paragraph)
    original_text = para_info["text"]
    if not original_text:
        return
    has_placeholder = False
    for key in keys:
        if re.search(r'\{\{\s*' + re.escape(key) + r'\s*\}\}', original_text):
            has_placeholder = True
            break
    if not has_placeholder:
        return
    new_text = replace_placeholders_in_text(original_text, keys, values)
    if new_text == original_text:
        return

    if not has_color_tag(new_text):
        for run in paragraph.runs:
            run.text = ""
        lines = new_text.split('\n')
        current_run = paragraph.runs[0]
        for i, line in enumerate(lines):
            if i == 0:
                current_run.text = line
            else:
                current_run.add_break()
                current_run.text += line
        apply_formatting_to_run(current_run, para_info["formatting"])
    else:
        formatting = para_info["formatting"]
        p_elem = paragraph._p
        for run in paragraph.runs:
            p_elem.remove(run._r)
        lines = new_text.split('\n')
        for line_idx, line in enumerate(lines):
            segments = parse_colored_segments(line)
            for seg_text, is_red in segments:
                if not seg_text:
                    continue
                new_run = OxmlElement('w:r')
                rPr = OxmlElement('w:rPr')
                if formatting.get("font.name"):
                    rFonts = OxmlElement('w:rFonts')
                    rFonts.set(qn('w:ascii'), formatting["font.name"])
                    rFonts.set(qn('w:hAnsi'), formatting["font.name"])
                    rFonts.set(qn('w:eastAsia'), formatting["font.name"])
                    rPr.append(rFonts)
                if formatting.get("bold"):
                    rPr.append(OxmlElement('w:b'))
                if formatting.get("italic"):
                    rPr.append(OxmlElement('w:i'))
                if formatting.get("underline"):
                    u = OxmlElement('w:u')
                    u.set(qn('w:val'), 'single')
                    rPr.append(u)
                if formatting.get("font.size"):
                    try:
                        sz_val = str(int(formatting["font.size"].pt * 2))
                        sz = OxmlElement('w:sz')
                        sz.set(qn('w:val'), sz_val)
                        rPr.append(sz)
                        szCs = OxmlElement('w:szCs')
                        szCs.set(qn('w:val'), sz_val)
                        rPr.append(szCs)
                    except Exception:
                        pass
                color_elem = OxmlElement('w:color')
                if is_red:
                    color_elem.set(qn('w:val'), 'FF0000')
                elif formatting.get("font.color.rgb"):
                    color_elem.set(qn('w:val'), str(formatting["font.color.rgb"]))
                else:
                    color_elem.set(qn('w:val'), 'auto')
                rPr.append(color_elem)
                new_run.append(rPr)
                t_elem = OxmlElement('w:t')
                t_elem.text = seg_text
                if seg_text.startswith(' ') or seg_text.endswith(' '):
                    t_elem.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
                new_run.append(t_elem)
                p_elem.append(new_run)
            if line_idx < len(lines) - 1:
                br_run = OxmlElement('w:r')
                br = OxmlElement('w:br')
                br_run.append(br)
                p_elem.append(br_run)


def process_table(table, keys, values):
    for row in table.rows:
        for cell in row.cells:
            for paragraph in cell.paragraphs:
                process_paragraph(paragraph, keys, values)


def process_document(doc, keys, values):
    processed_indices = set()

    # 1. 🆕 通用 Markdown 表格自动识别（任意 Key 均可触发）
    for i, (k, v) in enumerate(zip(keys, values)):
        if is_markdown_table(v):
            placeholder = f"{{{{{k}}}}}"
            if replace_placeholder_with_markdown_table(doc, placeholder, v):
                processed_indices.add(i)

    # 2. 兼容旧版 JSON 表格逻辑（研发内容完成情况）
    table_placeholder = '研发内容完成情况'
    table_data_key_idx = next((i for i, k in enumerate(keys) if k == table_placeholder), None)
    if table_data_key_idx is not None and table_data_key_idx not in processed_indices:
        raw_value = values[table_data_key_idx]
        try:
            table_data = json.loads(raw_value)
            if isinstance(table_data, list):
                replace_placeholder_with_table(doc, f'{{{{{table_placeholder}}}}}', table_data)
                processed_indices.add(table_data_key_idx)
        except Exception:
            pass

    # 3. 过滤已处理的 Key/Value，执行剩余文本替换
    new_keys = [k for i, k in enumerate(keys) if i not in processed_indices]
    new_values = [v for i, v in enumerate(values) if i not in processed_indices]

    for element in doc.element.body:
        if isinstance(element, CT_P):
            process_paragraph(Paragraph(element, doc), new_keys, new_values)
    for table in doc.tables:
        process_table(table, new_keys, new_values)


def generate_filename(original_url, custom_filename=None):
    if custom_filename:
        if not custom_filename.lower().endswith('.docx'):
            custom_filename += '.docx'
        return custom_filename
    parsed = urlparse(original_url)
    basename = os.path.basename(parsed.path)
    if basename and basename.lower().endswith('.docx'):
        unique_id = uuid.uuid4().hex[:8]
        name, ext = os.path.splitext(basename)
        return f"{name}_{unique_id}{ext}"
    return f"generated_{uuid.uuid4().hex[:8]}.docx"


@app.get("/")
async def root():
    return {"service": "制式文档生成API", "version": "2.0.0-md-table-generic", "status": "running"}


@app.get("/health")
async def health_check():
    return {"status": "healthy"}


@app.post("/generate")
async def generate_document(request: GenerateRequest):
    try:
        keys = parse_json_param(request.text_keys)
        values = parse_json_param(request.text_values)
        if len(keys) != len(values):
            raise HTTPException(
                status_code=400,
                detail=f"text_keys数量({len(keys)})与text_values数量({len(values)})不匹配"
            )
        if not keys:
            raise HTTPException(status_code=400, detail="text_keys不能为空")

        file_content = download_file(request.template_file_url)

        import tempfile
        temp_input = tempfile.NamedTemporaryFile(delete=False, suffix='.docx')
        temp_input.write(file_content)
        temp_input.close()

        try:
            doc = Document(temp_input.name)
            process_document(doc, keys, values)

            output_filename = generate_filename(request.template_file_url, request.filename)
            encoded_filename = quote(output_filename, encoding="utf-8")

            import io
            buffer = io.BytesIO()
            doc.save(buffer)
            file_bytes = buffer.getvalue()

            file_store[encoded_filename] = {
                "data": file_bytes,
                "expires": datetime.now() + timedelta(hours=2),
                "real_name": output_filename
            }

            download_url = f"{BASE_URL}/download/{encoded_filename}"

            return JSONResponse({
                "success": True,
                "message": "文档生成成功1",
                "full_download_url": download_url,
                "filename": output_filename,
                "replaced_count": len(keys)
            })
        finally:
            os.unlink(temp_input.name)

    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"生成文档失败: {str(e)}")


@app.get("/download/{filename}")
async def download_file_endpoint(filename: str):
    filename = os.path.basename(filename)
    if filename in file_store:
        item = file_store[filename]
        if datetime.now() < item["expires"]:
            real_name = item["real_name"]
            content_disposition = f"attachment; filename*=utf-8''{quote(real_name)}"
            return Response(
                content=item["data"],
                media_type='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
                headers={"Content-Disposition": content_disposition}
            )
        else:
            del file_store[filename]
            raise HTTPException(status_code=404, detail="文件已过期")
    raise HTTPException(status_code=404, detail="文件不存在")


def start_server():
    import uvicorn
    uvicorn.run(app, host=HOST, port=PORT)


if __name__ == "__main__":
    start_server()
