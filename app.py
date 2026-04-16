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
# 1. 크롤링 및 가구 종류 판별 함수
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

        title_meta = soup.find('meta', property='og:title')
        raw_name = title_meta['content'] if title_meta else "상품명 인식 실패"
        clean_name = raw_name.replace(' - (주)엠퍼니처', '').strip()

        # 가구 종류 판별
        category = "ITEM"
        if any(keyword in clean_name for keyword in ['체어', '의자', 'Chair', 'CHAIR']):
            category = "CHAIR"
        elif any(keyword in clean_name for keyword in ['테이블', '식탁', '데스크', 'Table', 'TABLE']):
            category = "TABLE"
        elif any(keyword in clean_name for keyword in ['소파', '쇼파', 'Sofa', 'SOFA']):
            category = "SOFA"

        img_meta = soup.find('meta', property='og:image')
        img_url = img_meta['content'] if img_meta else None
        if img_url and img_url.startswith('//'):
            img_url = 'https:' + img_url

        size = "사이즈 정보 없음"
        text_content = soup.get_text()
        size_match = re.search(r'[Ww]\s*\d+.*?([Hh]\s*\d+)', text_content)
        if size_match:
            size = size_match.group(0).strip()

        return {"name": clean_name, "category": category, "img_url": img_url, "size": size}
    except Exception:
        return {"name": "오류 발생", "category": "ITEM", "img_url": None, "size": "연결 실패"}

# ---------------------------------------------------------
# 2. 위치 기반 텍스트 교체 함수
# ---------------------------------------------------------
def replace_text_by_position(slide, search_text, replace_text, position='all'):
    prs_height = slide.parent.slide_height
    mid_point = prs_height / 2
    for shape in slide.shapes:
        if not shape.has_text_frame: continue
        if "2026" in shape.text: continue
        if position == 'top' and shape.top > mid_point: continue
        if position == 'bottom' and shape.top <= mid_point: continue
        for paragraph in shape.text_frame.paragraphs:
            for run in paragraph.runs:
                if search_text in run.text:
                    run.text = run.text.replace(search_text, str(replace_text))

# [이미지 삽입, 슬라이드 복제, 하단 삭제 함수는 이전과 동일]
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
        img_aspect = img_w / img_h
        target_aspect = tw / th
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

def delete_bottom_half(slide, prs_height):
    to_delete = [s for s in slide.shapes if s.top > (prs_height * 0.55)]
    for s in to_delete:
        sp = s._element
        sp.getparent().remove(sp)

# ---------------------------------------------------------
# 3. UI 및 메인 로직
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
        links_input = st.text_area("2. 가구 링크 (줄바꿈 구분)", height=250)
        links = [l.strip() for l in links_input.split('\n') if l.strip()]

    with col2:
        st.subheader("📜 최근 작업 기록 (최근 3개)")
        if not st.session_state.history:
            st.info("기록이 없습니다.")
        else:
            for idx, hist in enumerate(st.session_state.history):
                with st.expander(f"✅ {hist['title']}", expanded=(idx==0)):
                    all_links_str = "\n".join(hist['links'])
                    # 1. 링크 보여주기
                    st.text_area(f"링크 리스트 {idx+1}", value=all_links_str, height=100, disabled=True, key=f"hist_area_{idx}")
                    
                    # 2. 복사 버튼 추가 (클릭 시 클립보드로 복사)
                    if st.button(f"📋 링크 전체 복사 (기록 {idx+1})", key=f"copy_btn_{idx}"):
                        st.write_to_clipboard(all_links_str) # Streamlit 최신 기능
                        st.toast("클립보드에 복사되었습니다!")

    if st.button("🚀 제안서 생성하기", use_container_width=True):
        if not proposal_title or not links:
            st.warning("제목과 링크를 입력해 주세요.")
        else:
            with st.spinner("제안서를 구성 중입니다..."):
                product_data = [scrape_product_info(link) for link in links]
                prs = Presentation(TEMPLATE_FILE)
                source_slide = prs.slides[1]
                prs_height = prs.slide_height
                
                counts = {"CHAIR": 0, "TABLE": 0, "SOFA": 0, "ITEM": 0}
                
                for i in range(0, len(product_data), 2):
                    current_slide = duplicate_slide(prs, source_slide)
                    current_page_idx = len(prs.slides)
                    replace_text_by_position(current_slide, "2", str(current_page_idx))
                    
                    # 상단
                    item1 = product_data[i]
                    counts[item1['category']] += 1
                    cat_num1 = f"{item1['category']} {counts[item1['category']]:02d}"
                    replace_text_by_position(current_slide, "CHAIR 01", cat_num1, 'top')
                    replace_text_by_position(current_slide, "FA 루츠 암체어", item1['name'], 'top')
                    replace_text_by_position(current_slide, "W560 × D520 × H750 × SH440 × AH650", item1['size'], 'top')
                    replace_text_by_position(current_slide, "01", f"{i + 1:02d}", 'top')

                    # 하단
                    if i + 1 < len(product_data):
                        item2 = product_data[i+1]
                        counts[item2['category']] += 1
                        cat_num2 = f"{item2['category']} {counts[item2['category']]:02d}"
                        replace_text_by_position(current_slide, "TABLE 01", cat_num2, 'bottom')
                        replace_text_by_position(current_slide, "M 카라 테이블", item2['name'], 'bottom')
                        replace_text_by_position(current_slide, "W2600 × D900 × H730", item2['size'], 'bottom')
                        replace_text_by_position(current_slide, "02", f"{i + 2:02d}", 'bottom')
                    else:
                        delete_bottom_half(current_slide, prs_height)

                    pics = [s for s in current_slide.shapes if s.shape_type == 13]
                    pics.sort(key=lambda x: x.top)
                    if len(pics) >= 1: replace_image_fit(current_slide, pics[0], item1['img_url'])
                    if len(pics) >= 2 and (i + 1 < len(product_data)):
                        replace_image_fit(current_slide, pics[1], item2['img_url'])

                del prs.slides._sldIdLst[1]
                st.session_state.history.insert(0, {"title": proposal_title, "links": links})
                st.session_state.history = st.session_state.history[:3]

                output = io.BytesIO()
                prs.save(output)
                output.seek(0)
                st.success("🎉 생성 완료!")
                st.download_button("📥 다운로드", output, f"{proposal_title}.pptx", use_container_width=True)
