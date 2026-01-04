"""图片处理工具。"""

from pathlib import Path
from typing import Optional

from PIL import Image


def safe_remove_shape(shape) -> None:
    """安全移除形状元素。

    Args:
        shape: 要移除的形状对象
    """
    element_parent = shape.element.getparent()
    if element_parent is not None:
        element_parent.remove(shape.element)


def replace_picture(
    slide,
    shape,
    image_path: Optional[str]
) -> None:
    """将图片占位符替换为本地图片，保持位置大小比率。

    如果没有提供图片路径或图片文件不存在，删除该形状。

    Args:
        slide: 幻灯片对象
        shape: 图片占位符形状
        image_path: 图片文件路径
    """
    if not image_path:
        print(f"ℹ️  图片位置 [{shape.name}] 未提供图片，删除形状")
        safe_remove_shape(shape)
        return

    image_path = Path(image_path)
    if not image_path.is_file():
        print(f"⚠️  图片文件不可用：{image_path}，删除形状")
        safe_remove_shape(shape)
        return

    left, top, width, height = shape.left, shape.top, shape.width, shape.height
    name = shape.name

    try:
        with Image.open(image_path) as img:
            img_w, img_h = img.size
    except Exception as e:
        print(f"⚠️  无法读取图片 {image_path}：{e}，删除形状")
        safe_remove_shape(shape)
        return

    # 防止除零错误
    if img_h == 0 or height == 0:
        print(f"⚠️  图片或形状高度为 0，无法计算比例，删除形状")
        safe_remove_shape(shape)
        return

    # 计算保持比例的新尺寸
    img_ratio = img_w / img_h
    box_ratio = width / height
    if img_ratio > box_ratio:
        new_width = width
        new_height = width / img_ratio if img_ratio != 0 else height
    else:
        new_height = height
        new_width = height * img_ratio

    new_left = int(left + (width - new_width) / 2)
    new_top = int(top + (height - new_height) / 2)
    new_width = int(new_width)
    new_height = int(new_height)

    parent_shapes = getattr(shape, "_parent", None)
    safe_remove_shape(shape)

    if parent_shapes is None or not hasattr(parent_shapes, "add_picture"):
        parent_shapes = slide.shapes

    new_pic = parent_shapes.add_picture(
        str(image_path), new_left, new_top, width=new_width, height=new_height
    )
    new_pic.name = name
