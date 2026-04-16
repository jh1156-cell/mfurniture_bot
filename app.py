import streamlit as st
from pptx import Presentation
import requests
from bs4 import BeautifulSoup
import re
import io
import copy
import os
from PIL import Image
import numpy as np
import easyocr  # 이미지 속 글자 인식을 위한 라이브러리

# OCR 판독기 초기화 (한국어, 영어 지원)
@st.cache_resource
def load_ocr():
    return easyocr.Reader(['ko', 'en'])

reader = load_ocr()

# ---------------------------------------------------------
# 1. 크롤링 및 이미지 속 소재(Material) 추출
# ---------------------------------------------------------
def scrape_product_info(url):
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/112.0.0.0 Safari/537.36'
    }
    try:
        res = requests.get(url, headers=headers, timeout=10)
        res.raise_for_status()
        soup = BeautifulSoup(res.text, 'html.parser')

        # 기본 정보 추출
        title_meta = soup.find('meta', property='og:title')
        clean_name = title_meta['content'].replace(' - (주)엠퍼니처', '').strip() if title_meta else "상품명 없음"
        
        img_meta = soup.find('meta', property='og:image')
        img_url = img_meta['content'] if img_meta else None
        if img_url and img_url.startswith('//'): img_url = 'https:' + img_url

        # [OCR] 이미지에서 소재 정보 읽기
        material = "상세페이지 참조"
        if img_url:
            img_res = requests.get(img_url)
            img_data = Image.open(io.BytesIO(img_res.content))
            # 이미지를 넘파이 배열로 변환하여 OCR 실행
            img_np = np.array(img_data)
            ocr_result = reader.readtext(img_np, detail=0)
            all_text = " ".join(ocr_result)
            
            # OCR 결과 중 소재 관련 키워드 추출
            mat_match = re.search(r'(?:소재|재질|Material|Materials)\s*[:]?\s*([^/|▼|■|●]+)', all_text)
            if mat_match:
                material = mat_match.group(1).strip()[:30]

        # 사이즈 정보
        page_text = soup.get_text()
        size = "상세페이지 참조"
        size_match = re.search(r'[Ww]\s*\d+.*?([Hh]\s*\d+)', page_text)
        if size_match: size = size_match.group(0).strip()

        return {"name": clean_name, "img_url": img_url, "size": size, "material": material}
    except:
        return {"name": "오류", "img_url": None, "size": "-", "material": "-"}

# ---------------------------------------------------------
# 2. PPT 텍스트 배치 로직 (Size 바로 위 빈칸 공략)
# ---------------------------------------------------------
def process_slide_content(prs, slide, page_items, start_idx):
    prs_height = prs.slide_height
    mid_point = prs_height / 2
    
    for i, item in enumerate(page_items):
        is_top = (i == 0)
        current_num = f"{start_idx + i + 1:02d}"
        
        # 슬라이드 내 모든 도형 검사
        for shape in slide.shapes:
            if not shape.has_text_frame: continue
            
            # 영역 판별
            shape_is_top = shape.top < mid_point
            if is_top != shape_is_top: continue

            # 텍스트 치환
            txt = shape.text_frame.text
            
            # 1. 상품명/번호/사이즈 교체
            if "FA 루츠" in txt or "M 카라" in txt:
                shape.text_frame.text = item['name']
            elif txt.strip() == "01":
                shape.text_frame.text = current_num
            elif "W" in txt and "H" in txt:
                # 사이즈 정보 교체
                original_size_text = txt
                shape.text_frame.text = item['size']
                
                # [핵심] 사이즈 바로 위 도형을 찾아 소재 입력
                # 사이즈 도형의 바로 위(y축)에 있는 가장 가까운 빈 도형을 찾습니다.
                find_and_fill_material(slide, shape, item['material'])

def find_and_fill_material(slide, size_shape, material_text):
    # 사이즈 도형보다 약간 위에 위치한 도형들 찾기
    potential_shapes = []
    for s in slide.shapes:
        if s.has_text_frame:
            # 사이즈 도형의 top 값보다 작고(위쪽), x축 위치가 비슷한 도형
            if 0 < (size_shape.top - s.top) < (slide.parent.slide_height * 0.1):
                if abs(s.left - size_shape.left) < (size_shape.width * 0.5):
                    potential_shapes.append(s)
    
    # 가장 가까운 도형에 소재 텍스트 삽입
    if potential_shapes:
        target = min(potential_shapes, key=lambda x: size_shape.top - x.top)
        target.text_frame.text = material_text

# ---------------------------------------------------------
# (중략: duplicate_slide, delete_bottom_half 등 기존 함수 동일 사용)
# ---------------------------------------------------------

# Streamlit UI 부분에서 process_slide_content 호출 시 
# 위에서 수정한 로직이 적용됩니다.
