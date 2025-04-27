
"""
PowerPoint shape management functions.

This module provides functions for adding and manipulating shapes in PowerPoint
presentations.
"""
from typing import Any, Dict, List, Optional, Tuple

from pptx.enum.shapes import MSO_SHAPE
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor


# Define shape type mapping
SHAPE_TYPE_MAP = {
    'rectangle': MSO_SHAPE.RECTANGLE,
    'rounded_rectangle': MSO_SHAPE.ROUNDED_RECTANGLE,
    'oval': MSO_SHAPE.OVAL,
    'diamond': MSO_SHAPE.DIAMOND,
    'triangle': MSO_SHAPE.ISOSCELES_TRIANGLE,
    'isosceles_triangle': MSO_SHAPE.ISOSCELES_TRIANGLE,
    'right_triangle': MSO_SHAPE.RIGHT_TRIANGLE,
    'pentagon': MSO_SHAPE.PENTAGON,
    'hexagon': MSO_SHAPE.HEXAGON,
    'heptagon': MSO_SHAPE.HEPTAGON,
    'octagon': MSO_SHAPE.OCTAGON,
    'star': MSO_SHAPE.STAR_5_POINTS,
    'arrow': MSO_SHAPE.ARROW,
    'cloud': MSO_SHAPE.CLOUD,
    'heart': MSO_SHAPE.HEART,
    'lightning_bolt': MSO_SHAPE.LIGHTNING_BOLT,
    'sun': MSO_SHAPE.SUN,
    'moon': MSO_SHAPE.MOON,
    'smiley_face': MSO_SHAPE.SMILEY_FACE,
    'no_symbol': MSO_SHAPE.NO_SYMBOL,
    'flowchart_process': MSO_SHAPE.FLOWCHART_PROCESS,
    'flowchart_decision': MSO_SHAPE.FLOWCHART_DECISION,
    'flowchart_data': MSO_SHAPE.FLOWCHART_DATA,
    'flowchart_document': MSO_SHAPE.FLOWCHART_DOCUMENT,
    'flowchart_predefined_process': MSO_SHAPE.FLOWCHART_PREDEFINED_PROCESS,
    'flowchart_internal_storage': MSO_SHAPE.FLOWCHART_INTERNAL_STORAGE,
    'flowchart_connector': MSO_SHAPE.FLOWCHART_CONNECTOR
}


def add_shape(
    slide: Any, 
    shape_type: str, 
    left: float, 
    top: float, 
    width: float, 
    height: float
) -> Any:
    """
    Add an auto shape to a slide.
    
    Args:
        slide: The slide object
        shape_type: Shape type string (e.g., 'rectangle', 'oval', 'triangle')
        left: Left position in inches
        top: Top position in inches
        width: Width in inches
        height: Height in inches
        
    Returns:
        The created shape
        
    Raises:
        ValueError: If the shape type is invalid or if the shape cannot be added
    """
    shape_type_lower = str(shape_type).lower()
    
    if shape_type_lower not in SHAPE_TYPE_MAP:
        available_shapes = ', '.join(sorted(SHAPE_TYPE_MAP.keys()))
        raise ValueError(f"Unsupported shape type: '{shape_type}'. Available shape types: {available_shapes}")
    
    try:
        shape_enum = SHAPE_TYPE_MAP[shape_type_lower]
        shape = slide.shapes.add_shape(
            shape_enum, Inches(left), Inches(top), Inches(width), Inches(height)
        )
        return shape
    except Exception as e:
        raise ValueError(f"Failed to add shape '{shape_type}': {str(e)}")


def format_shape(
    shape: Any, 
    fill_color: Optional[Tuple[int, int, int]] = None, 
    line_color: Optional[Tuple[int, int, int]] = None, 
    line_width: Optional[float] = None
) -> None:
    """
    Format a shape.
    
    Args:
        shape: The shape object
        fill_color: RGB color tuple for fill (r, g, b)
        line_color: RGB color tuple for outline (r, g, b)
        line_width: Line width in points
        
    Raises:
        ValueError: If the shape cannot be formatted
    """
    try:
        if fill_color is not None:
            r, g, b = fill_color
            shape.fill.solid()
            shape.fill.fore_color.rgb = RGBColor(r, g, b)
        
        if line_color is not None:
            r, g, b = line_color
            shape.line.color.rgb = RGBColor(r, g, b)
        
        if line_width is not None:
            shape.line.width = Pt(line_width)
    except Exception as e:
        raise ValueError(f"Failed to format shape: {str(e)}")
