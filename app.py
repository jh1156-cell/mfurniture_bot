import streamlit as st
from pptx import Presentation
from pptx.util import Inches
import requests
from bs4 import BeautifulSoup
import re
import io
import copy
import os
from PIL import Image

# ---------------------------------------------------------
# 1. 크롤링 함수: 링크에서 제품명, 이미지, 사이즈 추출
# ---------------------------------------------------------
def scrape_product_info(url):
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
    }
    try:
        res = requests.get(url, headers=headers)
        res.raise_for_status()
        soup = BeautifulSoup(res.text, 'html.parser')

        # 제품명 추출
        title_meta = soup.find('meta', property='og:title')
        name = title_meta['content'] if title_meta else "상품명 인식 실패"

        # 이미지 추출
        img_meta = soup.find('meta', property='og:image')
        img_url = img_meta['content'] if img_meta else None

        # 사이즈 추출
        size = "사이즈 정보 없음"
        text_content = soup.get_text()
        size_match = re.search(r'[Ww]\s*\d+.*?([Hh]\s*\d+|[Aa][Hh]\s*\d+)', text_content)
        if size_match:
            size = size_match.group(0).strip()

        return {"name": name, "img_url": img_url, "size": size}
    except Exception as e:
        return {"name": "오류 발생", "img_url": None, "size": f"크롤링 실패: {e}"}

# ---------------------------------------------------------
# 2. 텍스트 교체 함수 (서식 유지)
# ---------------------------------------------------------
def replace_text(slide, search_text, replace_text):
    for shape in slide.shapes:
        if not shape.has_text_frame:
            continue
        for paragraph in shape.text_frame.paragraphs:
            for run in paragraph.runs:
                if search_text in run.text:
                    run.text = run.text.replace(search_text, str(replace_text))

# ---------------------------------------------------------
# 3. 이미지 Fit 교체 함수
# ---------------------------------------------------------
def replace_image_fit(slide, old_pic_shape, new_img_url):
    try:
        res = requests.get(new_img_url)
        img_bytes = io.BytesIO(res.content)
        img = Image.open(img_bytes)
        img_w, img_h = img.size

        target_x = old_pic_shape.left
        target_y = old_pic_shape.top
        target_w = old_pic_shape.width
        target_h = old_pic_shape.height

        aspect_img = img_w / img_h
        aspect_target = target_w / target_h

        if aspect_img > aspect_target:
            new_w = target_w
            new_h = int(target_w / aspect_img)
        else:
            new_h = target_h
            new_w = int(target_h * aspect_img)

        offset_x = target_x + (target_w - new_w) / 2
        offset_y = target_y + (target_h - new_h) / 2

        slide.shapes.add_picture(img_bytes, offset_x, offset_y, new_w, new_h)
        sp = old_pic_shape._element
        sp.getparent().remove(sp)
    except:
        pass

# ---------------------------------------------------------
# 4. 슬라이드 복제 함수 (IndexError 방지 버전)
# ---------------------------------------------------------
def duplicate_slide(prs, source_slide):
    try:
        layout = prs.slide_layouts[0] 
    except:
        layout = prs.slide_layouts[-1]
        
    new_slide = prs.slides.add_slide(layout)
    for shape in source_slide.shapes:
        new_el = copy.deepcopy(shape._element)
        new_slide.shapes._spTree.append(new_el)
    return new_slide

# ---------------------------------------------------------
# 5. 하단 영역 삭제 함수 (홀수 페이지 처리)
# ---------------------------------------------------------
def delete_bottom_half(slide, prs_height):
    shapes_to_delete = []
    for shape in slide.shapes:
        if shape.top >= (prs_height / 2):
            shapes_to_delete.append(shape)
    for shape in shapes_to_delete:
        sp = shape._element
        sp.getparent().remove(sp)

# ---------------------------------------------------------
# Streamlit 웹 UI
# ---------------------------------------------------------
st.set_page_config(page_title="제안서 자동 생성기", layout="wide")
st.title("🪑 매직퍼니처 제안서 자동 생성기")

# GitHub에 함께 올린 템플릿 파일 이름 (파일명이 다르면 아래 이름을 수정하세요)
TEMPLATE_FILE = "magic_furniture_proposal (1).pptx"

if not os.path.exists(TEMPLATE_FILE):
    st.error(f"⚠️ 템플릿 파일('{TEMPLATE_FILE}')을 찾을 수 없습니다. GitHub 레포지토리에 파일이 있는지 확인해주세요.")
else:
    links_input = st.text_area("가구 상세페이지 링크를 입력하세요 (엔터로 구분)", height=150)
    links = [link.strip() for link in links_input.split('\n') if link.strip()]

    if st.button("🚀 제안서 생성하기"):
        if not links:
            st.error("링크를 입력해주세요.")
        else:
            with st.spinner("제안서를 생성 중입니다..."):
                product_data = [scrape_product_info(link) for link in links]
                
                prs = Presentation(TEMPLATE_FILE)
                # 두 번째 슬라이드(index 1)를 템플릿으로 사용
                template_slide = prs.slides[1]
                prs_height = prs.slide_height
                
                total_items = len(product_data)
                
                for i in range(0, total_items, 2):
                    current_slide = duplicate_slide(prs, template_slide)
                    
                    # 상단 아이템
                    item1 = product_data[i]
                    page_num = (i // 2) + 1
                    replace_text(current_slide, "FA 루츠 암체어", item1['name'])
                    replace_text(current_slide, "W560 × D520 × H750 × SH440 × AH650", item1['size'])
                    replace_text(current_slide, "01", f"{i + 1:02d}")
                    replace_text(current_slide, "1", str(page_num))

                    # 하단 아이템 또는 삭제
                    if i + 1 < total_items:
                        item2 = product_data[i+1]
                        replace_text(current_slide, "M 카라 테이블", item2['name'])
                        replace_text(current_slide, "W2600 × D900 × H730", item2['size'])
                        replace_text(current_slide, "02", f"{i + 2:02d}")
                    else:
                        delete_bottom_half(current_slide, prs_height)

                    # 이미지 처리
                    pics = [s for s in current_slide.shapes if s.shape_type == 13]
                    pics.sort(key=lambda x: x.top)

                    if len(pics) >= 1 and item1['img_url']:
                        replace_image_fit(current_slide, pics[0], item1['img_url'])
                    if len(pics) >= 2 and (i + 1 < total_items) and item2['img_url']:
                        replace_image_fit(current_slide, pics[1], item2['img_url'])

                # 원본 템플릿 슬라이드 삭제(선택 사항)
                # prs.slides._sldIdLst.remove(prs.slides._sldIdLst[1]) 

                output = io.BytesIO()
                prs.save(output)
                output.seek(0)
                
                st.success("🎉 생성 완료!")
                st.download_button("📥 제안서 다운로드", output, "매직퍼니처_제안서.pptx")
