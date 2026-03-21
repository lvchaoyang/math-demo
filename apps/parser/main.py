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
            questions = split_questions(result['paragraphs'], file_id=file_id)
            
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
                "total_questions": len(questions_data)
            }
        
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"解析失败: {str(e)}")
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
    # TODO: 实现导出功能
    raise HTTPException(status_code=501, detail="导出功能尚未实现")


if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8001)
