# -*- coding: utf-8 -*-
"""
制式文档生成API服务
使用FastAPI + python-docx实现Word模板替换功能
"""

import os
import re
import uuid
import json
import base64
from typing import List, Optional, Union
from urllib.parse import urlparse
from datetime import datetime, timedelta

import requests
from fastapi import FastAPI, HTTPException, Query
from fastapi.responses import FileResponse, JSONResponse
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
from docx import Document
from docx.text.paragraph import Paragraph
from docx.table import Table
from docx.oxml.text.paragraph import CT_P
from docx.oxml.table import CT_Tc
from docx.oxml.shared import qn

# 配置
HOST = os.environ.get("HOST", "0.0.0.0")
PORT = int(os.environ.get("PORT", "8000"))
BASE_URL = os.environ.get("BASE_URL", f"http://localhost:{PORT}")

# 内存文件存储（防止磁盘文件丢失）
file_store = {}

app = FastAPI(
    title="制式文档生成API",
    description="Word模板占位符替换服务",
    version="1.2.0"
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


class GenerateRequest(BaseModel):
    template_file_url: str
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
        from docx.shared import RGBColor
        run.font.color.rgb = formatting["font.color.rgb"]


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
    if paragraph.runs:
        first_run = paragraph.runs[0]
        first_run.text = new_text
        for run in paragraph.runs[1:]:
            run.text = ""
        apply_formatting_to_run(first_run, para_info["formatting"])


def process_table(table, keys, values):
    for row in table.rows:
        for cell in row.cells:
            for paragraph in cell.paragraphs:
                process_paragraph(paragraph, keys, values)


def process_document(doc, keys, values):
    for element in doc.element.body:
        if isinstance(element, CT_P):
            paragraph = Paragraph(element, doc)
            process_paragraph(paragraph, keys, values)
    for table in doc.tables:
        process_table(table, keys, values)


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
    return {"service": "制式文档生成API", "version": "1.2.0", "status": "running"}


@app.get("/health")
async def health_check():
    return {"status": "healthy"}


@app.post("/generate")
async def generate_document(request: GenerateRequest):
    try:
        keys = parse_json_param(request.text_keys)
        values = parse_json_param(request.text_values)

        if len(keys) != len(values):
            raise HTTPException(status_code=400, detail=f"text_keys数量({len(keys)})与text_values数量({len(values)})不匹配")
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

            import io
            buffer = io.BytesIO()
            doc.save(buffer)
            file_bytes = buffer.getvalue()

            # 存到内存，2小时内可下载
            file_store[output_filename] = {
                "data": file_bytes,
                "expires": datetime.now() + timedelta(hours=2)
            }

            download_url = f"{BASE_URL}/download/{output_filename}"

            return JSONResponse({
                "success": True,
                "message": "文档生成成功",
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
            from fastapi.responses import Response
            return Response(
                content=item["data"],
                media_type='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
                headers={"Content-Disposition": f"attachment; filename={filename}"}
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
