from __future__ import annotations

from datetime import datetime
from pathlib import Path
import tempfile

import streamlit as st

from ppt_utils import build_presentation
from scraper import parse_links, scrape_product

BASE_DIR = Path(__file__).resolve().parent
TEMPLATE_PATH = BASE_DIR / "template.pptx"

st.set_page_config(page_title="Magic Furniture Proposal Generator", layout="centered")
st.title("Magic Furniture 링크 기반 PPT 자동화")
st.caption("링크만 넣으면 승인된 제안서 템플릿 형식으로 PPT를 생성합니다.")

sample_text = """https://magicfn.com/product/fa-%EB%A3%A8%EC%B8%A0-%EC%95%94%EC%B2%B4%EC%96%B4/21601/category/24/display/1/
https://magicfn.com/product/%EC%A3%BC%EB%AC%B8%EC%A0%9C%EC%9E%91-m-%EC%B9%B4%EB%9D%BC-%ED%85%8C%EC%9D%B4%EB%B8%94/21138/category/12/display/1/
"""

raw_links = st.text_area("상품 링크 붙여넣기", value=sample_text, height=220, help="한 줄에 하나씩 넣으면 됩니다.")
col1, col2 = st.columns(2)
year = col1.text_input("연도", value=str(datetime.today().year))
month = col2.text_input("월", value=f"{datetime.today().month:02d}")
output_name = st.text_input("파일명", value="magic_furniture_proposal")

if st.button("PPT 생성", type="primary"):
    links = parse_links(raw_links)
    if not links:
        st.error("유효한 링크를 한 개 이상 넣어주세요.")
        st.stop()

    if not TEMPLATE_PATH.exists():
        st.error("template.pptx 파일을 찾을 수 없습니다. app.py와 같은 폴더에 있어야 합니다.")
        st.stop()

    products = []
    progress = st.progress(0)
    status = st.empty()

    for idx, link in enumerate(links, start=1):
        status.write(f"수집 중: {idx}/{len(links)}")
        try:
            info = scrape_product(link)
            products.append(info.to_dict())
        except Exception as e:
            st.warning(f"링크 수집 실패: {link}\n\n{e}")
        progress.progress(idx / len(links))

    if not products:
        st.error("수집된 상품이 없습니다. 링크를 확인해주세요.")
        st.stop()

    with tempfile.TemporaryDirectory() as tmpdir:
        out_path = Path(tmpdir) / f"{output_name}.pptx"
        build_presentation(
            template_path=str(TEMPLATE_PATH),
            output_path=str(out_path),
            products=products,
            year=year,
            month=month,
        )
        data = out_path.read_bytes()

    st.success(f"완료: {len(products)}개 상품으로 PPT 생성")
    st.download_button(
        label="PPT 다운로드",
        data=data,
        file_name=f"{output_name}.pptx",
        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
    )

    with st.expander("수집 결과 확인"):
        st.json(products)
