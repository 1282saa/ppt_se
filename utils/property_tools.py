
"""
PowerPoint document properties management functions.

This module provides functions for setting and getting core document properties
in PowerPoint presentations.
"""
from typing import Dict, Optional

from pptx import Presentation


def set_core_properties(
    presentation: Presentation, 
    title: Optional[str] = None, 
    subject: Optional[str] = None,
    author: Optional[str] = None, 
    keywords: Optional[str] = None, 
    comments: Optional[str] = None
) -> None:
    """ 
    Set core document properties.
    
    Args:
        presentation: The Presentation object
        title: Document title
        subject: Document subject
        author: Document author
        keywords: Document keywords
        comments: Document comments
        
    Raises:
        ValueError: If the properties cannot be set
    """
    try:
        core_props = presentation.core_properties
        
        if title is not None:
            core_props.title = title
            
        if subject is not None:
            core_props.subject = subject
            
        if author is not None:
            core_props.author = author
            
        if keywords is not None:
            core_props.keywords = keywords
            
        if comments is not None:
            core_props.comments = comments
    except Exception as e:
        raise ValueError(f"Failed to set core properties: {str(e)}")


def get_core_properties(presentation: Presentation) -> Dict[str, str]:
    """
    Get core document properties.
    
    Args:
        presentation: The Presentation object
        
    Returns:
        Dictionary of core properties
        
    Raises:
        ValueError: If the properties cannot be retrieved
    """
    try:
        core_props = presentation.core_properties
        
        return {
            'title': core_props.title,
            'subject': core_props.subject,
            'author': core_props.author,
            'keywords': core_props.keywords,
            'comments': core_props.comments,
            'category': core_props.category,
            'created': core_props.created,
            'modified': core_props.modified,
            'last_modified_by': core_props.last_modified_by
        }
    except Exception as e:
        raise ValueError(f"Failed to get core properties: {str(e)}")
