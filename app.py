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
import easyocr

# OCR 판독기 캐싱 (처음 실행 시 로딩 후 재사용)
@st.cache_resource
def load_ocr():
    return easyocr.Reader(['ko', 'en'])

reader = load_ocr()

# ---------------------------------------------------------
# 1. 정보 추출 및 OCR (이미지 속 소재 읽기)
# ---------------------------------------------------------
def scrape_product_info(url):
    headers = {'User-Agent': 'Mozilla/5.0'}
    try:
        res = requests.get(url, headers=headers, timeout=10)
        res.raise_for_status()
        soup = BeautifulSoup(res.text, 'html.parser')

        # 상품명
        title_meta = soup.find('meta', property='og:title')
        clean_name = title_meta['content'].replace(' - (주)엠퍼니처', '').strip() if title_meta else "상품명 없음"
        
        # 이미지 URL
        img_meta = soup.find('meta', property='og:image')
        img_url = img_meta['content'] if img_meta else None
        if img_url and img_url.startswith('//'): img_url = 'https:' + img_url

        # [OCR] 이미지 분석으로 소재(Material) 찾기
        material = "상세페이지 참조"
        if img_url:
            try:
                img_res = requests.get(img_url)
                img_data = Image.open(io.BytesIO(img_res.content))
                img_np = np.array(img_data)
                ocr_result = reader.readtext(img_np, detail=0)
                all_text = " ".join(ocr_result)
                
                # '소재' 키워드 뒤의 텍스트 추출 (최대 30자)
                mat_match = re.search(r'(?:소재|재질|Material|Materials)\s*[:]?\s*([^/|▼|■|●]+)', all_text)
                if mat_match:
                    material = mat_match.group(1).strip()[:30]
            except: pass

        # 사이즈 정보 (텍스트 영역)
        page_text = soup.get_text()
        size = "상세페이지 참조"
        size_match = re.search(r'[Ww]\s*\d+.*?([Hh]\s*\d+)', page_text)
        if size_match: size = size_match.group(0).strip()

        return {"name": clean_name, "img_url": img_url, "size": size, "material": material}
    except:
        return {"name": "오류", "img_url": None, "size": "-", "material": "-"}

# ---------------------------------------------------------
# 2. PPT 배치 로직 (번호 순서 및 사이즈 위 칸 배치)
# ---------------------------------------------------------
def process_slide_content(prs, slide, page_items, start_idx):
    mid_point = prs.slide_height / 2
    
    for i, item in enumerate(page_items):
        is_top = (i == 0)
        current_num = f"{start_idx + i + 1:02d}"
        
        for shape in slide.shapes:
            if not shape.has_text_frame: continue
            
            shape_is_top = shape.top < mid_point
            if is_top != shape_is_top: continue

            txt = shape.text_frame.text
            
            # 텍스트 교체 (번호, 이름)
            if txt.strip() == "01":
                shape.text_frame.text = current_num
            elif "FA 루츠" in txt or "M 카라" in txt:
                shape.text_frame.text = item['name']
            elif "W" in txt and "H" in txt:
                # 사이즈 교체 후, 바로 위 빈 칸에 소재 입력
                shape.text_frame.text = item['size']
                find_and_fill_material(slide, shape, item['material'], prs.slide_height)

def find_and_fill_material(slide, size_shape, material_text, slide_height):
    potential_shapes = []
    for s in slide.shapes:
        if s.has_text_frame:
            # 사이즈 도형보다 위쪽에 위치한 가까운 도형 탐색
            if 0 < (size_shape.top - s.top) < (slide_height * 0.12):
                if abs(s.left - size_shape.left) < (size_shape.width * 0.5):
                    potential_shapes.append(s)
    
    if potential_shapes:
        target = min(potential_shapes, key=lambda x: size_shape.top - x.top)
        target.text_frame.text = material_text

# ---------------------------------------------------------
# 3. 이미지 및 슬라이드 유틸리티
# ---------------------------------------------------------
def replace_image_fit(slide, old_pic_shape, new_img_url):
    if not new_img_url: return
    try:
        res = requests.get(new_img_url, stream=True)
        img_bytes = io.BytesIO(res.content)
        with Image.open(img_bytes) as img:
            iw, ih = img.size
        img_bytes.seek(0)
        tx, ty, tw, th = old_pic_shape.left, old_pic_shape.top, old_pic_shape.width, old_pic_shape.height
        iasp, tasp = iw/ih, tw/th
        if iasp > tasp:
            nw, nh = tw, int(tw/iasp)
        else:
            nh, nw = th, int(th*iasp)
        ox, oy = tx + (tw-nw)/2, ty + (th-nh)/2
        slide.shapes.add_picture(img_bytes, ox, oy, nw, nh)
        old_pic_shape._element.getparent().remove(old_pic_shape._element)
    except: pass

def duplicate_slide(prs, source_slide):
    layout = prs.slide_layouts[0]
    new_slide = prs.slides.add_slide(layout)
    for shape in source_slide.shapes:
        new_el = copy.deepcopy(shape._element)
        new_slide.shapes._spTree.append(new_el)
    return new_slide

# ---------------------------------------------------------
# 4. UI 및 실행
# ---------------------------------------------------------
st.set_page_config(page_title="제안서 생성기", layout="wide")
st.title("🪑 매직퍼니처 자동 제안서 시스템")

TEMPLATE_FILE = "magic_furniture_proposal (1).pptx"

col1, col2 = st.columns([2, 1])
with col1:
    proposal_title = st.text_input("제안서 제목 (예: 가평 펜션 제안서)")
    links_input = st.text_area("가구 링크 (한 줄에 하나씩)", height=250)
    links = [l.strip() for l in links_input.split('\n') if
