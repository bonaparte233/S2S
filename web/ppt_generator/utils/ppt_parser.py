"""
PPT 模板解析工具

功能：
1. 提取 PPT 元素坐标信息
2. 自动过滤背景元素
3. 判断元素是否已命名
"""

import re
from pathlib import Path
from typing import Dict, List
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE


def is_generic_name(name: str) -> bool:
    """判断是否为通用名称（未命名）"""
    generic_pattern = re.compile(
        r'^(图片|文本框|矩形|圆角|任意|椭圆|线条|组合|对象|'
        r'table|textbox|picture|group|rectangle|oval|line|object)\s*\d*$',
        re.IGNORECASE
    )
    return bool(generic_pattern.match(name))


def extract_shapes_info(pptx_path: Path) -> Dict:
    """
    提取所有元素的坐标信息，自动过滤背景元素
    
    Returns:
        {
            'pages': [
                {
                    'page_num': 1,
                    'shapes': [
                        {
                            'shape_id': 0,
                            'name': 'TextBox 1',
                            'type': 'text',
                            'left': 914400,  # EMU
                            'top': 1828800,
                            'width': 4572000,
                            'height': 914400,
                            'text_sample': '人工智能导论',
                            'char_count': 6,
                            'is_named': False,
                            'is_hidden': False
                        }
                    ]
                }
            ]
        }
    """
    prs = Presentation(str(pptx_path))
    pages = []
    
    for slide_num, slide in enumerate(prs.slides, 1):
        shapes_info = []
        
        for shape_id, shape in enumerate(slide.shapes):
            # 自动过滤：跳过母版占位符
            if shape.is_placeholder:
                continue
            
            # 自动过滤：跳过明显的背景元素
            if shape.name.startswith(('背景', '装饰', 'Background', 'Decoration')):
                continue
            
            # 判断元素类型
            if shape.has_text_frame:
                shape_type = 'text'
            elif shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                shape_type = 'image'
            elif shape.shape_type == MSO_SHAPE_TYPE.GROUP:
                shape_type = 'group'
            else:
                shape_type = 'shape'
            
            # 只保留文本框和图片
            if shape_type not in ('text', 'image'):
                continue
            
            # 判断是否已命名（非通用名称）
            is_named = not is_generic_name(shape.name)
            
            info = {
                'shape_id': shape_id,
                'name': shape.name,
                'type': shape_type,
                'left': shape.left,  # EMU 单位
                'top': shape.top,
                'width': shape.width,
                'height': shape.height,
                'is_named': is_named,
                'is_hidden': False,  # 用户可以手动隐藏
            }
            
            # 如果是文本框，提取示例文本
            if shape.has_text_frame:
                info['text_sample'] = shape.text[:100] if shape.text else ''
                info['char_count'] = len(shape.text) if shape.text else 0
            
            shapes_info.append(info)
        
        pages.append({
            'page_num': slide_num,
            'shapes': shapes_info
        })
    
    return {'pages': pages}


def update_shape_name(pptx_path: Path, page_num: int, shape_id: int, new_name: str) -> None:
    """
    更新 PPT 中指定元素的名称
    
    Args:
        pptx_path: PPT 文件路径
        page_num: 页码（从 1 开始）
        shape_id: 元素 ID
        new_name: 新名称
    """
    prs = Presentation(str(pptx_path))
    
    if page_num < 1 or page_num > len(prs.slides):
        raise ValueError(f"页码 {page_num} 超出范围")
    
    slide = prs.slides[page_num - 1]
    
    if shape_id < 0 or shape_id >= len(slide.shapes):
        raise ValueError(f"元素 ID {shape_id} 超出范围")
    
    shape = slide.shapes[shape_id]
    shape.name = new_name
    
    # 保存修改
    prs.save(str(pptx_path))

