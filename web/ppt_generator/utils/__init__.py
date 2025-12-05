"""
PPT Generator utilities
"""

from .ppt_parser import extract_shapes_info, update_shape_name, is_generic_name
from .image_annotator import (
    annotate_screenshot,
    convert_pdf_to_images,
    convert_ppt_to_pdf,
    convert_ppt_to_images,
)

__all__ = [
    "extract_shapes_info",
    "update_shape_name",
    "is_generic_name",
    "annotate_screenshot",
    "convert_pdf_to_images",
    "convert_ppt_to_pdf",
    "convert_ppt_to_images",
]
