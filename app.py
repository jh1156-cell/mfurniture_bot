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
# 1. 크롤링 함수 (이미지 주소 및 텍스트 정제)
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

        # 제품명 정제
        title_meta = soup.find('meta', property='og:title')
        raw_name = title_meta['content'] if title_meta else "상품명 인식 실패"
        clean_name = raw_name.replace(' - (주)엠퍼니처', '').strip()

        # 이미지 URL (원본 소스 확보)
        img_meta = soup.find('meta', property='og:image')
        img_url = img_meta['content'] if img_meta else None
        if img_url and img_url.startswith('//'):
            img_url = 'https:' + img_url

        # 사이즈 추출
        size = "사이즈 정보 없음"
        text_content = soup.get_text()
        size_match = re.search(r'[Ww]\s*\d+.*?([Hh]\s*\d+)', text_content)
        if size_match:
            size = size_match.group(0).strip()

        return {"name": clean_name, "img_url": img_url, "size": size}
    except Exception as e:
        return {"name": "오류 발생", "img_url": None, "size": f"연결 실패"}

# ---------------------------------------------------------
# 2. 이미지 삽입 (에러 방지 강화)
# ---------------------------------------------------------
def replace_image_fit(slide, old_pic_shape, new_img_url):
    if not new_img_url:
        return
    try:
        headers = {'User-Agent': 'Mozilla/5.0'}
        img_res = requests.get(new_img_url, headers=headers, stream=True, timeout=10)
        img_res.raise_for_status()
        img_bytes = io.BytesIO(img_res.content)
        
        with Image.open(img_bytes) as img:
            img_w, img_h = img.size
            img_format = img.format
        
        img_bytes.seek(0)
        tx, ty, tw, th = old_pic_shape.left, old_pic_shape.top, old_pic_shape.width, old_pic_shape.height
        
        img_aspect = img_w / img_h
        target_aspect = tw / th
        
        if img_aspect > target_aspect:
            nw, nh = tw, int(tw / img_aspect)
        else:
            nh, nw = th, int(th * img_aspect)
            
        ox, oy = tx + (tw - nw) / 2, ty + (th - nh) / 2
        slide.shapes.add_picture(img_bytes, ox, oy, nw, nh)
        
        # 기존 개체 제거
        sp = old_pic_shape._element
        sp.getparent().remove(sp)
    except:
        pass

# [중간 생략: replace_text, duplicate_slide, delete_bottom_half 함수는 이전과 동일]
def replace_text(slide, search_text, replace_text):
    for shape in slide.shapes:
        if not shape.has_text_frame: continue
        if "2026" in shape.text: continue
        for paragraph in shape.text_frame.paragraphs:
            for run in paragraph.runs:
                if search_text in run.text:
                    run.text = run.text.replace(search_text, str(replace_text))

def duplicate_slide(prs, source_slide):
    layout = prs.slide_layouts[0]
    new_slide = prs.slides.add_slide(layout)
    for shape in source_slide.shapes:
        new_el = copy.deepcopy(shape._element)
        new_slide.shapes._spTree.append(new_el)
    return new_slide

def delete_bottom_half(slide, prs_height):
    to_delete = [s for s in slide.shapes if s.top > (prs_height * 0.55)]
    for s in to_delete:
        sp = s._element
        sp.getparent().remove(sp)

# ---------------------------------------------------------
# 3. Streamlit UI (기록 가독성 및 제목 입력 강화)
# ---------------------------------------------------------
st.set_page_config(page_title="제안서 생성기", layout="wide")
st.title("🪑 매직퍼니처 제안서 자동 생성 시스템")

if 'history' not in st.session_state: st.session_state.history = []
TEMPLATE_FILE = "magic_furniture_proposal (1).pptx"

if not os.path.exists(TEMPLATE_FILE):
    st.error(f"GitHub에 '{TEMPLATE_FILE}' 파일이 업로드되어 있어야 합니다.")
else:
    col1, col2 = st.columns([2, 1])
    
    with col1:
        proposal_title = st.text_input("1. 제안서 제목", placeholder="파일명으로 사용됩니다")
        links_input = st.text_area("2. 가구 상세페이지 링크 (줄바꿈으로 구분)", height=250)
        links = [l.strip() for l in links_input.split('\n') if l.strip()]

    with col2:
        st.subheader("📜 최근 작업 기록 (최근 2개)")
        if not st.session_state.history:
            st.info("아직 생성 기록이 없습니다.")
        else:
            for idx, hist in enumerate(st.session_state.history):
                with st.expander(f"✅ {hist['title']}", expanded=True):
                    # 링크가 잘리지 않도록 코드 블록 형태로 표시
                    all_links_str = "\n".join(hist['links'])
                    st.text_area(f"기록 {idx+1} 링크 리스트", value=all_links_str, height=100, disabled=True)
                    st.caption(f"생성된 가구 수: {len(hist['links'])}개")

    if st.button("🚀 제안서 생성하기", use_container_width=True):
        if not proposal_title or not links:
            st.warning("제목과 링크를 모두 입력해 주세요.")
        else:
            with st.spinner("데이터 수집 및 슬라이드 구성 중..."):
                product_data = [scrape_product_info(link) for link in links]
                prs = Presentation(TEMPLATE_FILE)
                source_slide = prs.slides[1]
                prs_height = prs.slide_height
                
                for i in range(0, len(product_data), 2):
                    current_slide = duplicate_slide(prs, source_slide)
                    current_page_idx = len(prs.slides)
                    replace_text(current_slide, "2", str(current_page_idx))
                    
                    # 상단 가구
                    item1 = product_data[i]
                    replace_text(current_slide, "FA 루츠 암체어", item1['name'])
                    replace_text(current_slide, "W560 × D520 × H750 × SH440 × AH650", item1['size'])
                    replace_text(current_slide, "01", f"{i + 1:02d}")

                    # 하단 가구
                    if i + 1 < len(product_data):
                        item2 = product_data[i+1]
                        replace_text(current_slide, "M 카라 테이블", item2['name'])
                        replace_text(current_slide, "W2600 × D900 × H730", item2['size'])
                        replace_text(current_slide, "02", f"{i + 2:02d}")
                    else:
                        delete_bottom_half(current_slide, prs_height)

                    # 이미지 매칭 (도형 타입 13 = Picture)
                    pics = [s for s in current_slide.shapes if s.shape_type == 13]
                    pics.sort(key=lambda x: x.top)

                    if len(pics) >= 1: replace_image_fit(current_slide, pics[0], item1['img_url'])
                    if len(pics) >= 2 and (i+1 < len(product_data)):
                        replace_image_fit(current_slide, pics[1], item2['img_url'])

                # 템플릿 슬라이드 제거 및 기록 업데이트
                del prs.slides._sldIdLst[1]
                st.session_state.history.insert(0, {"title": proposal_title, "links": links})
                st.session_state.history = st.session_state.history[:2]

                output = io.BytesIO()
                prs.save(output)
                output.seek(0)
                st.success(f"🎉 '{proposal_title}' 생성 완료!")
                st.download_button("📥 PPT 파일 다운로드", output, f"{proposal_title}.pptx", use_container_width=True)
