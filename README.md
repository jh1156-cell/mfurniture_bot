# Magic Furniture 링크 기반 PPT 자동화

## 실행
```bash
pip install -r requirements.txt
streamlit run app.py
```

## 포함 파일
- `app.py` : Streamlit UI
- `scraper.py` : 상품 링크에서 이름/사이즈/이미지 추출
- `ppt_utils.py` : 최종 승인 레이아웃 그대로 PPT 생성
- `template.pptx` : 사용자가 승인한 템플릿

## 특징
- 상품 링크만 입력하면 2개씩 한 장에 자동 배치
- `Magic Furniture` / `Magic-furniture` 표기 고정
- 모델명은 상품명과 동일하게 자동 입력
- 사진은 잘라내지 않고 원본 비율 그대로 박스 안에 맞춤
- 표지는 현재 연/월로 자동 갱신

## 주의
- Pretendard 폰트는 템플릿에 반영된 스타일을 그대로 따릅니다.
- 웹사이트 구조가 바뀌면 `scraper.py`의 추출 규칙을 조금 수정해야 할 수 있습니다.
