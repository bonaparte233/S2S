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
        r"^(图片|文本框|矩形|圆角|任意|椭圆|线条|组合|对象|"
        r"table|textbox|picture|group|rectangle|oval|line|object)\s*\d*$",
        re.IGNORECASE,
    )
    return bool(generic_pattern.match(name))


def is_background_shape(shape, slide_width: int, slide_height: int) -> bool:
    """
    判断是否为背景元素

    判断条件：
    1. 名称以背景/装饰开头
    2. 面积超过幻灯片面积的 80%（可能是背景矩形）
    3. 文本框但没有实际文本内容
    """
    # 名称检查
    if shape.name.startswith(("背景", "装饰", "Background", "Decoration")):
        return True

    # 面积检查：超过幻灯片 80% 面积的可能是背景
    shape_area = shape.width * shape.height
    slide_area = slide_width * slide_height
    if shape_area > slide_area * 0.8:
        return True

    # 文本框检查：有文本框架但没有实际文本
    if shape.has_text_frame:
        text = shape.text.strip() if shape.text else ""
        # 如果是"文本框"类型但没有文本，很可能是装饰性背景
        if not text and is_generic_name(shape.name):
            # 检查形状类型，圆角矩形等可能是背景
            try:
                if (
                    hasattr(shape, "auto_shape_type")
                    and shape.auto_shape_type is not None
                ):
                    # 有自动形状类型说明是形状而不是纯文本框
                    return True
            except:
                pass

    return False


def extract_shapes_info(pptx_path: Path) -> Dict:
    """
    提取所有元素的坐标信息，自动过滤背景元素

    Returns:
        {
            'slide_width': 12192000,  # EMU 单位
            'slide_height': 6858000,
            'pages': [
                {
                    'page_num': 1,
                    'shapes': [
                        {
                            'shape_id': 0,
                            'shape_index': 3,  # 在 slide.shapes 中的实际索引
                            'name': 'TextBox 1',
                            'type': 'text',
                            'left': 914400,  # EMU
                            'top': 1828800,
                            'width': 4572000,
                            'height': 914400,
                            'text_sample': '人工智能导论',
                            'char_count': 6,
                            'is_named': False,
                            'is_hidden': False,
                            'z_order': 3  # 层级顺序（越大越在上层）
                        }
                    ]
                }
            ]
        }
    """
    prs = Presentation(str(pptx_path))
    pages = []

    # 获取幻灯片尺寸
    slide_width = prs.slide_width
    slide_height = prs.slide_height

    for slide_num, slide in enumerate(prs.slides, 1):
        shapes_info = []

        for shape_index, shape in enumerate(slide.shapes):
            # 自动过滤：跳过母版占位符
            if shape.is_placeholder:
                continue

            # 自动过滤：使用更智能的背景检测
            if is_background_shape(shape, slide_width, slide_height):
                continue

            # 判断元素类型
            if shape.has_text_frame:
                shape_type = "text"
            elif shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                shape_type = "image"
            elif shape.shape_type == MSO_SHAPE_TYPE.GROUP:
                shape_type = "group"
            else:
                shape_type = "shape"

            # 只保留文本框和图片
            if shape_type not in ("text", "image"):
                continue

            # 判断是否已命名（非通用名称）
            is_named = not is_generic_name(shape.name)

            info = {
                "shape_id": shape.shape_id,  # 使用 shape 的真实 ID
                "shape_index": shape_index,  # 在 slide.shapes 中的索引
                "name": shape.name,
                "type": shape_type,
                "left": shape.left,  # EMU 单位
                "top": shape.top,
                "width": shape.width,
                "height": shape.height,
                "is_named": is_named,
                "is_hidden": False,  # 用户可以手动隐藏
                "z_order": shape_index,  # 层级顺序
            }

            # 如果是文本框，提取示例文本
            if shape.has_text_frame:
                info["text_sample"] = shape.text[:100] if shape.text else ""
                info["char_count"] = len(shape.text) if shape.text else 0

            shapes_info.append(info)

        # 按 z_order 排序，确保上层元素在后面
        shapes_info.sort(key=lambda x: x["z_order"])

        pages.append({"page_num": slide_num, "shapes": shapes_info})

    return {
        "slide_width": slide_width,
        "slide_height": slide_height,
        "pages": pages,
    }


def update_shape_name(
    pptx_path: Path, page_num: int, shape_index: int, new_name: str
) -> None:
    """
    更新 PPT 中指定元素的名称

    Args:
        pptx_path: PPT 文件路径
        page_num: 页码（从 1 开始）
        shape_index: 元素在 slide.shapes 中的索引
        new_name: 新名称
    """
    prs = Presentation(str(pptx_path))

    if page_num < 1 or page_num > len(prs.slides):
        raise ValueError(f"页码 {page_num} 超出范围")

    slide = prs.slides[page_num - 1]

    if shape_index < 0 or shape_index >= len(slide.shapes):
        raise ValueError(f"元素索引 {shape_index} 超出范围")

    shape = slide.shapes[shape_index]
    shape.name = new_name

    # 保存修改
    prs.save(str(pptx_path))
