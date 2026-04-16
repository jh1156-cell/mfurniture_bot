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
# 1. 크롤링 및 데이터 정제 (User-Agent 강화)
# ---------------------------------------------------------
def scrape_product_info(url):
    # 일반 브라우저처럼 보이기 위한 헤더 설정 강화
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/110.0.0.0 Safari/537.36',
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

        # 이미지 URL 추출 및 보정
        img_meta = soup.find('meta', property='og:image')
        img_url = img_meta['content'] if img_meta else None
        if img_url and not img_url.startswith('http'):
            img_url = 'https:' + img_url

        # 사이즈 추출
        size = "사이즈 정보 없음"
        text_content = soup.get_text()
        size_match = re.search(r'[Ww]\s*\d+.*?([Hh]\s*\d+)', text_content)
        if size_match:
            size = size_match.group(0).strip()

        return {"name": clean_name, "img_url": img_url, "size": size}
    except Exception as e:
        return {"name": "오류 발생", "img_url": None, "size": f"연결 실패: {e}"}

# ---------------------------------------------------------
# 2. 이미지 삽입 함수 (가장 중요한 수정 부분)
# ---------------------------------------------------------
def replace_image_fit(slide, old_pic_shape, new_img_url):
    if not new_img_url:
        return
    try:
        # 이미지 다운로드 시에도 헤더 포함 (엑스 표시 방지)
        headers = {'User-Agent': 'Mozilla/5.0'}
        img_res = requests.get(new_img_url, headers=headers, stream=True, timeout=10)
        img_res.raise_for_status()
        img_bytes = io.BytesIO(img_res.content)
        
        # 이미지 유효성 확인
        with Image.open(img_bytes) as img:
            img.verify() # 파일이 깨졌는지 확인
        
        img_bytes.seek(0) # verify 후 포인터 초기화
        with Image.open(img_bytes) as img:
            img_w, img_h = img.size
        
        # 기존 틀 정보
        tx, ty, tw, th = old_pic_shape.left, old_pic_shape.top, old_pic_shape.width, old_pic_shape.height
        
        # 비율 계산 (Fit)
        img_aspect = img_w / img_h
        target_aspect = tw / th
        
        if img_aspect > target_aspect:
            nw, nh = tw, int(tw / img_aspect)
        else:
            nh, nw = th, int(th * img_aspect)
            
        ox, oy = tx + (tw - nw) / 2, ty + (th - nh) / 2
        
        # 새 이미지 삽입
        slide.shapes.add_picture(img_bytes, ox, oy, nw, nh)
        
        # 기존 '엑스' 또는 '빈칸' 틀 제거
        sp = old_pic_shape._element
        sp.getparent().remove(sp)
    except Exception as e:
        st.error(f"이미지 삽입 실패: {e}")

# ---------------------------------------------------------
# [이전 기능들과 동일한 함수들: replace_text, duplicate_slide, delete_bottom_half 생략]
# ---------------------------------------------------------
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
# Streamlit 실행부
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
        proposal_title = st.text_input("1. 제안서 제목", placeholder="파일명 입력")
        links_input = st.text_area("2. 가구 링크 입력", height=200)
        links = [l.strip() for l in links_input.split('\n') if l.strip()]

    with col2:
        st.subheader("📜 최근 작업 기록")
        for idx, hist in enumerate(st.session_state.history):
            with st.expander(f"{hist['title']}"):
                for l in hist['links']: st.write(f"• {l[:30]}...")

    if st.button("🚀 제안서 생성"):
        if not proposal_title or not links:
            st.warning("제목과 링크를 모두 입력해 주세요.")
        else:
            with st.spinner("이미지를 매칭하는 중..."):
                product_data = [scrape_product_info(link) for link in links]
                prs = Presentation(TEMPLATE_FILE)
                source_slide = prs.slides[1]
                prs_height = prs.slide_height
                
                for i in range(0, len(product_data), 2):
                    current_slide = duplicate_slide(prs, source_slide)
                    
                    # 페이지 번호 및 텍스트 교체
                    current_page_idx = len(prs.slides)
                    replace_text(current_slide, "2", str(current_page_idx))
                    
                    # 상단 가구
                    item1 = product_data[i]
                    replace_text(current_slide, "FA 루츠 암체어", item1['name'])
                    replace_text(current_slide, "W560 × D520 × H750 × SH440 × AH650", item1['size'])
                    replace_text(current_slide, "01", f"{i + 1:02d}")

                    # 하단 가구 처리
                    if i + 1 < len(product_data):
                        item2 = product_data[i+1]
                        replace_text(current_slide, "M 카라 테이블", item2['name'])
                        replace_text(current_slide, "W2600 × D900 × H730", item2['size'])
                        replace_text(current_slide, "02", f"{i + 2:02d}")
                    else:
                        delete_bottom_half(current_slide, prs_height)

                    # --- 이미지 교체 핵심 로직 ---
                    # 슬라이드 내 모든 이미지 개체(Type 13)를 찾음
                    pics = [s for s in current_slide.shapes if s.shape_type == 13]
                    # 위에서 아래 순서로 정렬 (상단 가구 틀이 먼저 오게)
                    pics.sort(key=lambda x: x.top)

                    if len(pics) >= 1:
                        replace_image_fit(current_slide, pics[0], item1['img_url'])
                    if len(pics) >= 2 and (i + 1 < len(product_data)):
                        replace_image_fit(current_slide, pics[1], item2['img_url'])

                # 원본 템플릿 삭제 및 기록 저장
                del prs.slides._sldIdLst[1]
                st.session_state.history.insert(0, {"title": proposal_title, "links": links})
                st.session_state.history = st.session_state.history[:2]

                output = io.BytesIO()
                prs.save(output)
                output.seek(0)
                st.success("🎉 생성 완료!")
                st.download_button("📥 다운로드", output, f"{proposal_title}.pptx")
