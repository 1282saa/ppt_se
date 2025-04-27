#!/usr/bin/env python
"""
PowerPoint MCP Server - 메인 진입점

이 파일은 PowerPoint MCP 서버를 시작하는 메인 진입점입니다.
config_loader.py를 통해 설정을 로드하고, ppt_generator.py를 통해 PPT 생성 기능을 제공합니다.
"""
import argparse
import os
import sys
from config_loader import load_design_config
from ppt_generator import PPTGenerator

def main():
    """메인 함수"""
    parser = argparse.ArgumentParser(description='PowerPoint MCP 서버')
    parser.add_argument('--host', default='127.0.0.1', help='호스트 주소 (기본값: 127.0.0.1)')
    parser.add_argument('--port', type=int, default=8000, help='포트 번호 (기본값: 8000)')
    parser.add_argument('--design', '-d', default='data/design_system.json', 
                        help='디자인 시스템 구성 파일 (기본값: data/design_system.json)')
    parser.add_argument('--content', '-c', default='data/slide_content.json',
                        help='슬라이드 내용 파일 (기본값: data/slide_content.json)')
    parser.add_argument('--output', '-o', default=None,
                        help='출력 파일 경로 (기본값: 슬라이드내용_generated.pptx)')
    
    args = parser.parse_args()
    
    # 설정 로드
    design_config = load_design_config(args.design)
    
    # MCP 서버 시작 코드 추가 예정
    # 현재는 기본 PPT 생성 기능만 구현
    generator = PPTGenerator(args.design, args.content, args.output)
    output_path = generator.generate()
    
    print(f"\n프레젠테이션이 성공적으로 생성되었습니다: {output_path}")
    
    # 여기에 MCP 서버 시작 코드 추가

if __name__ == "__main__":
    main() 