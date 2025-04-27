#!/usr/bin/env python
"""
PowerPoint MCP Server - 설정 로더

이 파일은 PowerPoint MCP 서버의 설정을 로드하는 모듈입니다.
디자인 시스템 구성 파일과 기타 설정을 로드하는 기능을 제공합니다.
"""
import json
import os
from pathlib import Path
from typing import Dict, Any, Optional


def load_design_config(config_path: str) -> Dict[str, Any]:
    """
    디자인 시스템 구성 파일을 로드합니다.
    
    Args:
        config_path: 디자인 시스템 구성 파일 경로
        
    Returns:
        설정 데이터를 담은 딕셔너리
        
    Raises:
        FileNotFoundError: 파일을 찾을 수 없는 경우
        ValueError: 파일 형식이 유효하지 않은 경우
    """
    try:
        with open(config_path, 'r', encoding='utf-8') as f:
            config = json.load(f)
        return config
    except FileNotFoundError:
        raise FileNotFoundError(f"디자인 시스템 구성 파일을 찾을 수 없습니다: {config_path}")
    except json.JSONDecodeError:
        raise ValueError(f"디자인 시스템 구성 파일이 유효한 JSON 형식이 아닙니다: {config_path}")


def load_slide_content(content_path: str) -> Dict[str, Any]:
    """
    슬라이드 내용 파일을 로드합니다.
    
    Args:
        content_path: 슬라이드 내용 파일 경로
        
    Returns:
        슬라이드 내용 데이터를 담은 딕셔너리
        
    Raises:
        FileNotFoundError: 파일을 찾을 수 없는 경우
        ValueError: 파일 형식이 유효하지 않은 경우
    """
    try:
        with open(content_path, 'r', encoding='utf-8') as f:
            content = json.load(f)
        return content
    except FileNotFoundError:
        raise FileNotFoundError(f"슬라이드 내용 파일을 찾을 수 없습니다: {content_path}")
    except json.JSONDecodeError:
        raise ValueError(f"슬라이드 내용 파일이 유효한 JSON 형식이 아닙니다: {content_path}")


def get_output_path(content_path: str, output_path: Optional[str] = None) -> str:
    """
    출력 파일 경로를 결정합니다.
    
    Args:
        content_path: 슬라이드 내용 파일 경로
        output_path: 사용자가 지정한 출력 경로 (기본값: None)
        
    Returns:
        결정된 출력 파일 경로
    """
    if output_path:
        return output_path
    
    # 출력 경로가 지정되지 않은 경우, 슬라이드 내용 파일 이름 기반으로 생성
    base_name = Path(content_path).stem
    return os.path.join('output', f"{base_name}_generated.pptx")


def ensure_output_directory(output_path: str) -> None:
    """
    출력 디렉토리가 존재하는지 확인하고, 없으면 생성합니다.
    
    Args:
        output_path: 출력 파일 경로
    """
    output_dir = os.path.dirname(output_path)
    if output_dir and not os.path.exists(output_dir):
        os.makedirs(output_dir) 