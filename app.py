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
# 1. 크롤링 및 정보 추출 (소재 추출 로직 추가)
# ---------------------------------------------------------
def scrape_product_info(url):
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/112.0.0.0 Safari/537.36',
        'Referer': 'https://magicfn.com/'
    }
    try:
        res = requests.get(url, headers=headers, timeout=10)
        res.raise_for_status()
        soup = BeautifulSoup(res.text, 'html.parser')

        # 상품명
        title_meta = soup.find('meta', property='og:title')
        raw_name = title_meta['content'] if title_meta else "상품명 인식 실패"
        clean_name = raw_name.replace(' - (주)엠퍼니처', '').strip()

        # 가구 종류
        category = "ITEM"
        if any(keyword in clean_name for keyword in ['체어', '의자', 'Chair', 'CHAIR']):
            category = "CHAIR"
        elif any(keyword in clean_name for keyword in ['테이블', '식탁', '데스크', 'Table', 'TABLE']):
            category = "TABLE"
        elif any(keyword in clean_name for keyword in ['소파', '쇼파', 'Sofa', 'SOFA']):
            category = "SOFA"

        # [NEW] 소재(Material) 추출: 텍스트 영역에서 '소재' 키워드 탐색
        page_text = soup.get_text(separator=' ', strip=True)
        material = "상세페이지 참조"
        # '소재 : 원목' 혹은 '소재 원목' 형태 탐색
        mat_match = re.search(r'(?:소재|재질|Material)\s*[:]?\s*([^/|\n|▼|■]+)', page_text)
        if mat_match:
            material = mat_match.group(1).strip()[:30] # 너무 길면 30자에서 자름

        # 이미지 및 사이즈
        img_meta = soup.find('meta', property='og:image')
        img_url = img_meta['content'] if img_meta else None
        if img_url and img_url.startswith('//'): img_url = 'https:' + img_url

        size = "상세페이지 참조"
        size_match = re.search(r'[Ww]\s*\d+.*?([Hh]\s*\d+)', page_text)
        if size_match: size = size_match.group(0).strip()

        return {"name": clean_name, "category": category, "img_url": img_url, "size": size, "material": material}
    except Exception:
        return {"name": "오류 발생", "category": "ITEM", "img_url": None, "size": "-", "material": "-"}

# ---------------------------------------------------------
# 2. 텍스트 치환 (번호 꼬임 해결 & 영역 독립 처리)
# ---------------------------------------------------------
def replace_text_in_shape(shape, search_text, replace_text):
    if not shape.has_text_frame: return
    for paragraph in shape.text_frame.paragraphs:
        for run in paragraph.runs:
            if search_text in run.text:
                run.text = run.text.replace(search_text, str(replace_text))

def process_slide_content(prs, slide, page_items, start_idx, counts):
    prs_height = prs.slide_height
    mid_point = prs_height / 2
    
    for i, item in enumerate(page_items):
        is_top_data = (i == 0)
        current_num = f"{start_idx + i + 1:02d}"
        counts[item['category']] += 1
        cat_label = f"{item['category']} {counts[item['category']]:02d}"
        
        for shape in slide.shapes:
            if not shape.has_text_frame: continue
            
            # 위치 판별 (현재 shape이 상단인지 하단인지)
            shape_is_top = shape.top < mid_point
            if is_top_data != shape_is_top: continue # 내 영역 데이터가 아니면 무시

            # 상단 영역 교체
            if is_top_data:
                replace_text_in_shape(shape, "CHAIR 01", cat_label)
                replace_text_in_shape(shape, "FA 루츠 암체어", item['name'])
                replace_text_in_shape(shape, "W560 × D520 × H750 × SH440 × AH650", item['size'])
                replace_text_in_shape(shape, "상세페이지 참조", item['material']) # 템플릿의 METERIAL 내용
                replace_text_in_shape(shape, "01", current_num) # PRODUCT 번호
            # 하단 영역 교체
            else:
                replace_text_in_shape(shape, "TABLE 01", cat_label)
                replace_text_in_shape(shape, "M 카라 테이블", item['name'])
                replace_text_in_shape(shape, "W2600 × D900 × H730", item['size'])
                replace_text_in_shape(shape, "상세페이지 참조", item['material'])
                replace_text_in_shape(shape, "01", current_num)

# ---------------------------------------------------------
# 3. 이미지 및 슬라이드 제어
# ---------------------------------------------------------
def replace_image_fit(slide, old_pic_shape, new_img_url):
    if not new_img_url: return
    try:
        headers = {'User-Agent': 'Mozilla/5.0'}
        img_res = requests.get(new_img_url, headers=headers, stream=True, timeout=10)
        img_res.raise_for_status()
        img_bytes = io.BytesIO(img_res.content)
        with Image.open(img_bytes) as img:
            img_w, img_h = img.size
        img_bytes.seek(0)
        tx, ty, tw, th = old_pic_shape.left, old_pic_shape.top, old_pic_shape.width, old_pic_shape.height
        img_aspect, target_aspect = img_w / img_h, tw / th
        if img_aspect > target_aspect:
            nw, nh = tw, int(tw / img_aspect)
        else:
            nh, nw = th, int(th * img_aspect)
        ox, oy = tx + (tw - nw) / 2, ty + (th - nh) / 2
        slide.shapes.add_picture(img_bytes, ox, oy, nw, nh)
        sp = old_pic_shape._element
        sp.getparent().remove(sp)
    except: pass

def duplicate_slide(prs, source_slide):
    layout = prs.slide_layouts[0]
    new_slide = prs.slides.add_slide(layout)
    for shape in source_slide.shapes:
        new_el = copy.deepcopy(shape._element)
        new_slide.shapes._spTree.append(new_el)
    return new_slide

def delete_bottom_half(prs, slide):
    to_delete = [s for s in slide.shapes if s.top > (prs.slide_height * 0.55)]
    for s in to_delete:
        sp = s._element
        sp.getparent().remove(sp)

# ---------------------------------------------------------
# 4. Streamlit UI
# ---------------------------------------------------------
st.set_page_config(page_title="제안서 생성기", layout="wide")
st.title("🪑 매직퍼니처 제안서 자동 생성 시스템")

if 'history' not in st.session_state: st.session_state.history = []
TEMPLATE_FILE = "magic_furniture_proposal (1).pptx"

if not os.path.exists(TEMPLATE_FILE):
    st.error("템플릿 파일이 없습니다.")
else:
    col1, col2 = st.columns([2, 1])
    with col1:
        proposal_title = st.text_input("1. 제안서 제목")
        links_input = st.text_area("2. 가구 링크 (줄바꿈 구분)", height=250)
        links = [l.strip() for l in links_input.split('\n') if l.strip()]

    with col2:
        st.subheader("📜 최근 작업 기록")
        if st.session_state.history:
            for idx, hist in enumerate(st.session_state.history):
                with st.expander(f"✅ {hist['title']}", expanded=(idx==0)):
                    st.code("\n".join(hist['links']), language=None)

    if st.button("🚀 제안서 생성하기", use_container_width=True):
        if not proposal_title or not links:
            st.warning("정보를 입력하세요.")
        else:
            with st.spinner("정보 추출 및 제안서 생성 중..."):
                product_data = [scrape_product_info(link) for link in links]
                prs = Presentation(TEMPLATE_FILE)
                source_slide = prs.slides[1]
                counts = {"CHAIR": 0, "TABLE": 0, "SOFA": 0, "ITEM": 0}
                
                for i in range(0, len(product_data), 2):
                    current_slide = duplicate_slide(prs, source_slide)
                    # 페이지 번호
                    for shape in current_slide.shapes:
                        if shape.has_text_frame and shape.text.strip() == "2":
                            shape.text_frame.text = str(len(prs.slides))

                    page_items = product_data[i:i+2]
                    process_slide_content(prs, current_slide, page_items, i, counts)
                    
                    if len(page_items) < 2:
                        delete_bottom_half(prs, current_slide)

                    pics = [s for s in current_slide.shapes if s.shape_type == 13]
                    pics.sort(key=lambda x: x.top)
                    for p_idx, item in enumerate(page_items):
                        if p_idx < len(pics):
                            replace_image_fit(current_slide, pics[p_idx], item['img_url'])

                del prs.slides._sldIdLst[1]
                st.session_state.history.insert(0, {"title": proposal_title, "links": links})
                st.session_state.history = st.session_state.history[:3]
                output = io.BytesIO()
                prs.save(output)
                output.seek(0)
                st.success("완료!")
                st.download_button("📥 다운로드", output, f"{proposal_title}.pptx", use_container_width=True)
