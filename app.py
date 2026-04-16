import streamlit as st
from pptx import Presentation
import requests
from bs4 import BeautifulSoup
import re
import io
import copy
import os
from PIL import Image

# ---------------------------------------------------------
# 1. 크롤링 및 데이터 정제 (상품명에서 회사명 제거)
# ---------------------------------------------------------
def scrape_product_info(url):
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
    }
    try:
        res = requests.get(url, headers=headers)
        res.raise_for_status()
        soup = BeautifulSoup(res.text, 'html.parser')

        # 제품명 추출 및 정제 (' - (주)엠퍼니처' 제거)
        title_meta = soup.find('meta', property='og:title')
        raw_name = title_meta['content'] if title_meta else "상품명 인식 실패"
        clean_name = raw_name.replace(' - (주)엠퍼니처', '').strip()

        # 이미지 추출 (원본 이미지 주소 확보)
        img_meta = soup.find('meta', property='og:image')
        img_url = img_meta['content'] if img_meta else None
        if img_url and not img_url.startswith('http'):
            img_url = 'https:' + img_url

        # 사이즈 추출 (W, D, H 패턴 탐색)
        size = "사이즈 정보 없음"
        text_content = soup.get_text()
        size_match = re.search(r'[Ww]\s*\d+.*?([Hh]\s*\d+)', text_content)
        if size_match:
            size = size_match.group(0).strip()

        return {"name": clean_name, "img_url": img_url, "size": size}
    except Exception as e:
        return {"name": "오류 발생", "img_url": None, "size": f"연결 실패: {e}"}

# ---------------------------------------------------------
# 2. 텍스트 교체 (2026 고정 및 부분 일치 방지)
# ---------------------------------------------------------
def replace_text(slide, search_text, replace_text):
    for shape in slide.shapes:
        if not shape.has_text_frame:
            continue
        # "2026"이라는 텍스트가 포함된 도형은 수정하지 않음 (고정값 보호)
        if "2026" in shape.text:
            continue
            
        for paragraph in shape.text_frame.paragraphs:
            for run in paragraph.runs:
                if search_text in run.text:
                    run.text = run.text.replace(search_text, str(replace_text))

# ---------------------------------------------------------
# 3. 이미지 삽입 (에러 방지 및 Fit 처리)
# ---------------------------------------------------------
def replace_image_fit(slide, old_pic_shape, new_img_url):
    if not new_img_url:
        return
    try:
        res = requests.get(new_img_url, stream=True)
        img_bytes = io.BytesIO(res.content)
        
        # 이미지 크기 측정
        with Image.open(img_bytes) as img:
            img_w, img_h = img.size
        
        # 기준 틀 위치
        tx, ty, tw, th = old_pic_shape.left, old_pic_shape.top, old_pic_shape.width, old_pic_shape.height
        
        # 비율 계산 (Fit)
        img_aspect = img_w / img_h
        target_aspect = tw / th
        
        if img_aspect > target_aspect:
            nw = tw
            nh = int(tw / img_aspect)
        else:
            nh = th
            nw = int(th * img_aspect)
            
        ox = tx + (tw - nw) / 2
        oy = ty + (th - nh) / 2
        
        # 이미지 삽입 후 기존 틀 제거
        slide.shapes.add_picture(img_bytes, ox, oy, nw, nh)
        sp = old_pic_shape._element
        sp.getparent().remove(sp)
    except:
        pass

# ---------------------------------------------------------
# 4. 슬라이드 복제 및 하단 삭제
# ---------------------------------------------------------
def duplicate_slide(prs, source_slide):
    layout = prs.slide_layouts[0]
    new_slide = prs.slides.add_slide(layout)
    for shape in source_slide.shapes:
        new_el = copy.deepcopy(shape._element)
        new_slide.shapes._spTree.append(new_el)
    return new_slide

def delete_bottom_half(slide, prs_height):
    # 슬라이드 아래쪽 50% 지점보다 낮은 개체들 삭제
    to_delete = [s for s in slide.shapes if s.top > (prs_height * 0.55)]
    for s in to_delete:
        sp = s._element
        sp.getparent().remove(sp)

# ---------------------------------------------------------
# Streamlit 실행부
# ---------------------------------------------------------
st.set_page_config(page_title="제안서 생성기", layout="wide")
st.title("🪑 가구 제안서 자동 생성 시스템")

TEMPLATE_FILE = "magic_furniture_proposal (1).pptx"

if not os.path.exists(TEMPLATE_FILE):
    st.error(f"GitHub에 '{TEMPLATE_FILE}' 파일이 업로드되어 있어야 합니다.")
else:
    links_input = st.text_area("가구 페이지 링크 입력 (줄바꿈으로 구분)", height=200)
    links = [l.strip() for l in links_input.split('\n') if l.strip()]

    if st.button("🚀 제안서 생성"):
        if not links:
            st.warning("링크를 입력해 주세요.")
        else:
            with st.spinner("이미지를 불러오고 페이지를 구성 중입니다..."):
                product_data = [scrape_product_info(link) for link in links]
                
                prs = Presentation(TEMPLATE_FILE)
                # 2페이지(index 1)를 복사할 원본으로 설정
                source_slide = prs.slides[1]
                prs_height = prs.slide_height
                
                # 가구 데이터 처리
                for i in range(0, len(product_data), 2):
                    current_slide = duplicate_slide(prs, source_slide)
                    
                    # [공통] 오른쪽 하단 페이지 번호 (슬라이드 마스터 효과)
                    # 템플릿의 '2'를 현재 슬라이드 번호(표지 포함)로 변경
                    current_page_idx = len(prs.slides)
                    replace_text(current_slide, "2", str(current_page_idx))
                    
                    # --- 1번째 가구 (상단) ---
                    item1 = product_data[i]
                    replace_text(current_slide, "FA 루츠 암체어", item1['name'])
                    replace_text(current_slide, "W560 × D520 × H750 × SH440 × AH650", item1['size'])
                    # PRODUCT 번호 (01, 02, 03... 순차적으로 증가)
                    replace_text(current_slide, "01", f"{i + 1:02d}")

                    # --- 2번째 가구 (하단) ---
                    if i + 1 < len(product_data):
                        item2 = product_data[i+1]
                        replace_text(current_slide, "M 카라 테이블", item2['name'])
                        replace_text(current_slide, "W2600 × D900 × H730", item2['size'])
                        # 하단 PRODUCT 번호는 상단 다음 번호
                        replace_text(current_slide, "02", f"{i + 2:02d}")
                    else:
                        # 홀수일 경우 하단 삭제
                        delete_bottom_half(current_slide, prs_height)

                    # 이미지 배치 (순서대로 상단, 하단 틀 찾기)
                    pics = [s for s in current_slide.shapes if s.shape_type == 13]
                    pics.sort(key=lambda x: x.top)

                    if len(pics) >= 1:
                        replace_image_fit(current_slide, pics[0], item1['img_url'])
                    if len(pics) >= 2 and (i + 1 < len(product_data)):
                        replace_image_fit(current_slide, pics[1], item2['img_url'])

                # 작업 완료 후, 원본 예시 템플릿(2페이지) 삭제
                # (이렇게 해야 생성된 결과물만 남습니다)
                xml_slides = prs.slides._sldIdLst
                del xml_slides[1]

                # 최종 저장
                output = io.BytesIO()
                prs.save(output)
                output.seek(0)
                
                st.success("🎉 제안서 작성이 완료되었습니다!")
                st.download_button("📥 완성본 다운로드", output, "Final_Proposal.pptx")
