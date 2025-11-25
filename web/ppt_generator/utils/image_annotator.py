"""
PPT 截图标注工具

功能：
1. 在 PPT 截图上绘制编号圆圈
2. 支持不同状态的颜色标注
"""

import platform
import subprocess
from pathlib import Path
from typing import List, Dict
from PIL import Image, ImageDraw, ImageFont
from pdf2image import convert_from_path


def get_soffice_path():
    """获取 LibreOffice soffice 可执行文件路径（跨平台）"""
    system = platform.system()

    if system == "Darwin":  # macOS
        return "/Applications/LibreOffice.app/Contents/MacOS/soffice"
    elif system == "Windows":
        # 常见安装路径
        paths = [
            r"C:\Program Files\LibreOffice\program\soffice.exe",
            r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
        ]
        for path in paths:
            if Path(path).exists():
                return path
        raise FileNotFoundError(
            "LibreOffice not found. Please install LibreOffice:\n"
            "Windows: choco install libreoffice\n"
            "macOS: brew install --cask libreoffice\n"
            "Linux: sudo apt-get install libreoffice"
        )
    else:  # Linux
        return "soffice"


def convert_ppt_to_pdf(pptx_path: Path, output_dir: Path) -> Path:
    """
    使用 LibreOffice 将 PPT 转换为 PDF（跨平台）

    Args:
        pptx_path: PPT 文件路径
        output_dir: 输出目录

    Returns:
        生成的 PDF 文件路径
    """
    soffice = get_soffice_path()

    # 确保输出目录存在
    output_dir.mkdir(parents=True, exist_ok=True)

    # 调用 LibreOffice 转换
    subprocess.run(
        [
            soffice,
            "--headless",
            "--convert-to",
            "pdf",
            "--outdir",
            str(output_dir),
            str(pptx_path),
        ],
        check=True,
        capture_output=True,
        text=True,
    )

    # 返回生成的 PDF 路径
    pdf_name = pptx_path.stem + ".pdf"
    return output_dir / pdf_name


def convert_pdf_to_images(
    pdf_path: Path, output_dir: Path, dpi: int = 300
) -> List[Path]:
    """
    将 PDF 转换为图片

    Args:
        pdf_path: PDF 文件路径
        output_dir: 输出目录
        dpi: 图片分辨率

    Returns:
        生成的图片路径列表
    """
    # 确保输出目录存在
    output_dir.mkdir(parents=True, exist_ok=True)

    # 转换 PDF 为图片
    images = convert_from_path(pdf_path, dpi=dpi)

    # 保存图片
    image_paths = []
    for i, image in enumerate(images, start=1):
        image_path = output_dir / f"page_{i}.png"
        image.save(image_path, "PNG")
        image_paths.append(image_path)

    return image_paths


def annotate_screenshot(
    image_path: Path, shapes_info: List[Dict], dpi: int = 300
) -> Path:
    """
    在 PPT 截图上标注元素编号

    Args:
        image_path: 截图路径
        shapes_info: 元素信息列表（已过滤背景元素）
        dpi: 图片 DPI

    Returns:
        标注后的图片路径
    """
    img = Image.open(image_path)
    draw = ImageDraw.Draw(img)

    # 加载字体
    try:
        # macOS 系统字体
        font = ImageFont.truetype("/System/Library/Fonts/Helvetica.ttc", 28)
    except:
        try:
            # Linux 系统字体
            font = ImageFont.truetype(
                "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf", 28
            )
        except:
            # 使用默认字体
            font = ImageFont.load_default()

    EMU_PER_INCH = 914400

    for idx, shape in enumerate(shapes_info, 1):
        # 跳过隐藏元素
        if shape.get("is_hidden"):
            continue

        # 计算坐标（EMU 转像素）
        x = int(shape["left"] / EMU_PER_INCH * dpi)
        y = int(shape["top"] / EMU_PER_INCH * dpi)

        # 确定颜色（根据是否已命名）
        is_named = shape.get("is_named", False)
        circle_color = "#007BFF" if is_named else "#FFC107"

        # 绘制圆形背景
        radius = 22
        draw.ellipse(
            [x - radius, y - radius, x + radius, y + radius],
            fill=circle_color,
            outline="#FFFFFF",
            width=3,
        )

        # 绘制编号
        text = str(idx)

        # 计算文本位置（居中）
        try:
            bbox = draw.textbbox((0, 0), text, font=font)
            text_width = bbox[2] - bbox[0]
            text_height = bbox[3] - bbox[1]
        except:
            # 如果 textbbox 不可用，使用估算
            text_width = len(text) * 14
            text_height = 20

        draw.text(
            (x - text_width // 2, y - text_height // 2 - 2),
            text,
            fill="#FFFFFF",
            font=font,
        )

    # 保存标注后的图片
    annotated_path = image_path.parent / f"{image_path.stem}_annotated.png"
    img.save(annotated_path)
    return annotated_path


def convert_pdf_to_images(
    pdf_path: Path, output_dir: Path, dpi: int = 300
) -> List[Path]:
    """
    将 PDF 转换为图片

    Args:
        pdf_path: PDF 文件路径
        output_dir: 输出目录
        dpi: 分辨率

    Returns:
        图片路径列表
    """
    from pdf2image import convert_from_path

    output_dir.mkdir(parents=True, exist_ok=True)

    images = convert_from_path(
        str(pdf_path),
        dpi=dpi,
        fmt="png",
        output_folder=str(output_dir),
        paths_only=False,  # 返回 PIL Image 对象
    )

    # 保存图片
    image_paths = []
    for i, img in enumerate(images, 1):
        img_path = output_dir / f"page_{i}.png"
        img.save(img_path)
        image_paths.append(img_path)

    return image_paths
