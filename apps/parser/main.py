"""
Math Demo Parser Service
Python 解析服务 - 核心解析层
"""

from fastapi import FastAPI, File, UploadFile, HTTPException, Form
from fastapi.responses import FileResponse, JSONResponse
from pathlib import Path
import shutil
import uuid

from app.core.parser import parse_docx
from app.core.splitter import split_questions
from app.core.docx_to_html import convert_docx_to_html
from app.core.exporter import export_questions as export_questions_to_word

app = FastAPI(
    title="Math Demo Parser Service",
    description="数学试卷解析服务",
    version="1.0.0"
)

# 确保目录存在（使用统一的数据目录）
BASE_DIR = Path(__file__).parent.parent.parent / "data"
UPLOAD_DIR = BASE_DIR / "uploads"
EXPORT_DIR = BASE_DIR / "exports"
IMAGES_DIR = BASE_DIR / "images"
UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
EXPORT_DIR.mkdir(parents=True, exist_ok=True)
IMAGES_DIR.mkdir(parents=True, exist_ok=True)


@app.get("/health")
async def health_check():
    """健康检查"""
    return {"status": "healthy", "service": "parser"}


@app.post("/parse")
async def parse_document(file: UploadFile = File(...), mode: str = Form("questions")):
    """
    解析 Word 文档
    
    - **file**: docx 格式的试卷文件
    - **mode**: 解析模式（questions=题目拆分, html=整体HTML）
    """
    if not file.filename.endswith('.docx'):
        raise HTTPException(status_code=400, detail="只支持 .docx 格式的文件")
    
    file_id = str(uuid.uuid4())
    file_path = UPLOAD_DIR / f"{file_id}_{file.filename}"
    
    try:
        # 保存上传的文件
        with open(file_path, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)
        
        # 根据模式选择解析方式
        if mode == "html":
            # 新模式：整体返回 HTML
            html_content = convert_docx_to_html(str(file_path))
            
            return {
                "success": True,
                "file_id": file_id,
                "mode": "html",
                "html": html_content,
                "message": "HTML 转换完成"
            }
        else:
            # 原有模式：题目拆分
            # 创建图片输出目录
            image_dir = IMAGES_DIR / file_id
            image_dir.mkdir(parents=True, exist_ok=True)
            
            # 解析文档
            result = parse_docx(
                str(file_path),
                extract_images=True,
                image_output_dir=str(image_dir),
                file_id=file_id
            )
            
            # 拆分题目
            questions = split_questions(
                result['paragraphs'],
                file_id=file_id,
                formula_render_plan=result.get('formula_render_plan'),
            )
            
            # 转换为字典格式
            questions_data = []
            for q in questions:
                question_dict = {
                    "id": q.id,
                    "number": q.number,
                    "type": q.type.value,
                    "type_name": q.type_name,
                    "content": q.content,
                    "content_html": q.content_html,
                    "images": q.images,
                    "latex_formulas": q.latex_formulas,
                    "score": q.score,
                    "difficulty": q.difficulty,
                }
                
                if q.options:
                    question_dict["options"] = [
                        {
                            "label": o.label,
                            "content": o.content,
                            "content_html": o.content_html if o.content_html else f'<span class="option-label">{o.label}.</span> {o.content}',
                            "is_latex": o.is_latex
                        }
                        for o in q.options
                    ]
                
                if q.answer:
                    question_dict["answer"] = {
                        "content": q.answer,
                        "content_html": f'<div class="answer"><strong>答案：</strong>{q.answer}</div>'
                    }
                
                if q.analysis:
                    question_dict["analysis"] = {
                        "content": q.analysis,
                        "content_html": f'<div class="analysis"><strong>解析：</strong>{q.analysis}</div>'
                    }
                
                questions_data.append(question_dict)
            
            return {
                "success": True,
                "file_id": file_id,
                "mode": "questions",
                "questions": questions_data,
                "total_questions": len(questions_data),
                "formula_render_summary": result.get("formula_render_summary", {}),
                "formula_asset_debug": result.get("formula_asset_debug", {}),
                "formula_render_plan": result.get("formula_render_plan", []),
            }
        
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"解析失败：{str(e)}")
    finally:
        file.file.close()


@app.post("/parse/v2")
async def parse_document_v2(file: UploadFile = File(...), mode: str = Form("questions"), use_pandoc: bool = Form(True)):
    """
    解析 Word 文档（增强版，支持高质量公式解析）
    
    - **file**: docx 格式的试卷文件
    - **mode**: 解析模式（questions=题目拆分，html=整体 HTML）
    - **use_pandoc**: 是否使用 Pandoc 方案，默认 true
    """
    if not file.filename.endswith('.docx'):
        raise HTTPException(status_code=400, detail="只支持 .docx 格式的文件")
    
    file_id = str(uuid.uuid4())
    file_path = UPLOAD_DIR / f"{file_id}_{file.filename}"
    
    try:
        # 保存上传的文件
        with open(file_path, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)
        
        # 根据模式选择解析方式
        if mode == "html":
            # 使用修复后的解析器（避免公式重叠）
            from app.core.unified_parser import parse_docx_unified
            
            success, result, error = parse_docx_unified(
                str(file_path),
                extract_images=True,
                image_output_dir=str(IMAGES_DIR / file_id),
                file_id=file_id
            )
            
            if not success:
                raise Exception(f"解析失败：{error}")
            
            html_content = result.get('html_content', '')
            
            return {
                "success": True,
                "file_id": file_id,
                "mode": "html",
                "html": html_content,
                "method": "unified_parser",
                "message": "HTML 转换完成"
            }
        else:
            # 题目拆分模式固定走结构化解析链路：
            # parse_docx -> split_questions
            # 原因：splitter 依赖 content_items 的顺序与类型来生成 content_html/option_html，
            # unified(Pandoc) 的段落结构会导致题干或选项公式图片缺失。
            image_dir = IMAGES_DIR / file_id
            image_dir.mkdir(parents=True, exist_ok=True)
            parse_result = parse_docx(
                str(file_path),
                extract_images=True,
                image_output_dir=str(image_dir),
                file_id=file_id
            )
            # 为兼容前端展示保留 method 字段，显式标记为稳定拆题链路
            parse_method = "standard_questions"

            # 拆分题目
            questions = split_questions(
                parse_result['paragraphs'],
                file_id=file_id,
                formula_render_plan=parse_result.get('formula_render_plan'),
            )
            
            # 转换为字典格式
            questions_data = []
            for q in questions:
                question_dict = {
                    'id': q.id,
                    'number': q.number,
                    'type': q.type.value,
                    'type_name': q.type_name,
                    'content': q.content,
                    'content_html': q.content_html,
                    'images': q.images,
                    'latex_formulas': q.latex_formulas,
                    'score': q.score,
                    'difficulty': q.difficulty,
                    'answer': q.answer,
                    'analysis': q.analysis,
                }

                if q.options:
                    question_dict['options'] = [
                        {
                            'label': o.label,
                            'content': o.content,
                            'content_html': o.content_html,
                            'is_latex': o.is_latex
                        }
                        for o in q.options
                    ]
                else:
                    question_dict['options'] = []

                questions_data.append(question_dict)
            
            return {
                "success": True,
                "file_id": file_id,
                "mode": "questions",
                "questions": questions_data,
                "total_questions": len(questions_data),
                "method": parse_method,
                "formula_render_summary": parse_result.get("formula_render_summary", {}),
                "formula_asset_debug": parse_result.get("formula_asset_debug", {}),
                "formula_render_plan": parse_result.get("formula_render_plan", []),
            }
        
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"解析失败：{str(e)}")
    finally:
        file.file.close()


@app.get("/images/{file_id}/{filename}")
async def get_image(file_id: str, filename: str):
    """
    获取图片

    - **file_id**: 文件ID
    - **filename**: 图片文件名
    """
    # 图片存储在 IMAGES_DIR / file_id / 目录下
    image_dir = IMAGES_DIR / file_id
    image_path = image_dir / filename

    # 如果文件名不包含 file_id 前缀，尝试添加前缀
    if not image_path.exists() and not filename.startswith(f"{file_id}_"):
        image_path = image_dir / f"{file_id}_{filename}"

    # 如果还找不到，尝试不带前缀的原始文件名
    if not image_path.exists() and filename.startswith(f"{file_id}_"):
        # 去掉前缀再试
        original_name = filename[len(f"{file_id}_"):]
        image_path = image_dir / original_name

    # 兼容 Pandoc 提取目录结构（如 images/{file_id}/media/...）
    if not image_path.exists():
        # 先在子目录按原文件名递归查找
        recursive_matches = [p for p in image_dir.rglob(filename) if p.is_file()]
        if recursive_matches:
            image_path = recursive_matches[0]

    if not image_path.exists():
        # 再尝试按去前缀后的文件名递归查找
        if filename.startswith(f"{file_id}_"):
            original_name = filename[len(f"{file_id}_"):]
            recursive_matches = [p for p in image_dir.rglob(original_name) if p.is_file()]
            if recursive_matches:
                image_path = recursive_matches[0]

    if not image_path.exists():
        raise HTTPException(status_code=404, detail=f"图片不存在: {image_path}")

    return FileResponse(path=image_path)


@app.post("/export")
async def export_questions(data: dict):
    """
    导出题目到 Word
    
    - **file_id**: 文件ID
    - **question_ids**: 题目ID列表
    - **options**: 导出选项
    """
    file_id = data.get("file_id")
    question_ids = data.get("question_ids") or []
    options = data.get("options") or {}
    title = data.get("title") or options.get("title") or "导出的题目"
    watermark = options.get("watermark")

    include_answer = bool(options.get("include_answer", False))
    include_analysis = bool(options.get("include_analysis", False))

    if not file_id or not isinstance(question_ids, list) or len(question_ids) == 0:
        raise HTTPException(status_code=400, detail="参数错误: 需要 file_id 和 question_ids")

    # 找到上传的原始 docx 文件
    candidates = sorted(list(UPLOAD_DIR.glob(f"{file_id}_*.docx")))
    if not candidates:
        raise HTTPException(status_code=404, detail="未找到原始上传文件")
    docx_path = candidates[0]

    # 重新解析并拆题，保证 question_id 可匹配
    image_dir = IMAGES_DIR / file_id
    image_dir.mkdir(parents=True, exist_ok=True)
    result = parse_docx(
        str(docx_path),
        extract_images=True,
        image_output_dir=str(image_dir),
        file_id=file_id,
    )
    questions = split_questions(
        result["paragraphs"],
        file_id=file_id,
        formula_render_plan=result.get("formula_render_plan"),
    )
    selected_questions = [q for q in questions if q.id in set(question_ids)]

    if not selected_questions:
        raise HTTPException(status_code=400, detail="未找到需要导出的题目")

    EXPORT_DIR.mkdir(parents=True, exist_ok=True)
    out_path = EXPORT_DIR / f"export_{file_id}_{uuid.uuid4().hex}.docx"

    export_questions_to_word(
        selected_questions,
        str(out_path),
        title=title,
        watermark=watermark,
        image_dir=str(image_dir),
        include_answer=include_answer,
        include_analysis=include_analysis,
    )

    return FileResponse(
        path=str(out_path),
        filename=out_path.name,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )


if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)
