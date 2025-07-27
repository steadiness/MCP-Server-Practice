# PowerPoint MCP Server Manual

## 시작하기
* 📝 [새로운 프레젠테이션을 만드는 방법](#새로운-프레젠테이션을-만드는-방법)
* 🖍️ [프레젠테이션 스타일 살펴보기](#프레젠테이션-스타일-살펴보기)
  * [스타일 사용하기](#스타일-사용하기)
  * [텍스트 이상의 기능](#텍스트-이상의-기능) - 프레젠테이션은 이미지, 차트, 도형을 비롯하여 사실상 모든 멀티미디어 요소를 포함할 수 있습니다.
  * [아이디어 시각화하기](#아이디어-시각화하기) - 다이어그램과 플로우차트도 가능합니다!
* ℹ️ [프레젠테이션 파악하기](#프레젠테이션-파악하기) - 정보 패널에서 슬라이드 수, 읽는 시간과 같은 통계와 프레젠테이션에 대한 개요를 확인할 수 있습니다

---

## 새로운 프레젠테이션을 만드는 방법

### 기본 프레젠테이션 생성
```python
# 빈 프레젠테이션 생성
mcp_ppt_create_presentation()

# 제목이 있는 프레젠테이션 생성
mcp_ppt_create_presentation(title="내 첫 번째 프레젠테이션")
```

### 템플릿에서 프레젠테이션 생성
```python
# 템플릿 파일에서 생성
mcp_ppt_create_presentation_from_template(template_path="templates/template.pptx")

# 자동 생성 (주제 기반)
mcp_ppt_auto_generate_presentation(
    topic="2024년 사업 계획",
    slide_count=8,
    presentation_type="business",
    color_scheme="modern_blue"
)
```

---

## 프레젠테이션 스타일 살펴보기

### 스타일 사용하기

#### 색상 스키마 적용
```python
# 다양한 색상 스키마 지원
- modern_blue: 현대적인 파란색 테마
- corporate_gray: 기업용 회색 테마  
- elegant_green: 우아한 녹색 테마
- warm_red: 따뜻한 빨간색 테마
```

#### 전문적인 디자인 적용
```python
mcp_ppt_apply_professional_design(
    operation="apply_theme",
    color_scheme="modern_blue",
    enhance_title=True,
    enhance_content=True,
    enhance_shapes=True,
    enhance_charts=True
)
```

### 텍스트 이상의 기능

#### 이미지 추가
```python
# 파일에서 이미지 추가
mcp_ppt_manage_image(
    slide_index=0,
    operation="add",
    image_source="path/to/image.jpg",
    left=1, top=1, width=4, height=3
)

# 이미지 효과 적용
mcp_ppt_apply_picture_effects(
    slide_index=0,
    shape_index=0,
    effects={
        "brightness": 1.2,
        "contrast": 1.1,
        "saturation": 0.9,
        "filter_type": "sepia"
    }
)
```

#### 차트 생성
```python
mcp_ppt_add_chart(
    slide_index=1,
    chart_type="bar",
    left=1, top=1, width=6, height=4,
    categories=["Q1", "Q2", "Q3", "Q4"],
    series_names=["매출", "비용"],
    series_values=[[100, 120, 140, 160], [80, 90, 100, 110]],
    title="분기별 실적",
    color_scheme="modern_blue"
)
```

#### 도형 추가
```python
mcp_ppt_add_shape(
    slide_index=2,
    shape_type="rectangle",
    left=2, top=2, width=3, height=2,
    fill_color=[52, 152, 219],
    text="중요 포인트",
    font_size=14,
    font_color=[255, 255, 255]
)
```

#### 테이블 생성
```python
mcp_ppt_add_table(
    slide_index=3,
    rows=4, cols=3,
    left=1, top=1, width=6, height=3,
    data=[
        ["항목", "2023", "2024"],
        ["매출", "1,000", "1,200"],
        ["비용", "800", "900"],
        ["이익", "200", "300"]
    ],
    header_bg_color=[52, 152, 219],
    header_font_size=12
)
```

### 아이디어 시각화하기

#### 연결선 추가
```python
mcp_ppt_add_connector(
    slide_index=4,
    connector_type="elbow",
    start_x=1, start_y=2,
    end_x=5, end_y=4,
    line_width=2,
    color=[52, 152, 219]
)
```

#### 템플릿 기반 슬라이드 생성
```python
# 사용 가능한 템플릿
- title_slide: 제목 슬라이드
- text_with_image: 텍스트와 이미지
- bullet_points: 글머리 기호 목록
- comparison: 비교 슬라이드
- process_flow: 프로세스 플로우
- data_visualization: 데이터 시각화

mcp_ppt_create_slide_from_template(
    template_id="process_flow",
    content_mapping={
        "title": "개발 프로세스",
        "step1": "요구사항 분석",
        "step2": "설계",
        "step3": "구현",
        "step4": "테스트"
    },
    color_scheme="modern_blue"
)
```

---

## 프레젠테이션 파악하기

### 프레젠테이션 정보 확인
```python
# 기본 정보 조회
mcp_ppt_get_presentation_info()

# 슬라이드별 정보
mcp_ppt_get_slide_info(slide_index=0)

# 전체 텍스트 추출
mcp_ppt_extract_presentation_text(include_slide_info=True)
```

### 텍스트 최적화
```python
mcp_ppt_optimize_slide_text(
    slide_index=0,
    auto_resize=True,
    auto_wrap=True,
    optimize_spacing=True,
    min_font_size=8,
    max_font_size=36
)
```

### 하이퍼링크 관리
```python
# 하이퍼링크 추가
mcp_ppt_manage_hyperlinks(
    operation="add",
    slide_index=0,
    shape_index=0,
    text="자세히 보기",
    url="https://example.com"
)
```

---

## 고급 기능

### 슬라이드 전환 효과
```python
mcp_ppt_manage_slide_transitions(
    slide_index=0,
    operation="set",
    transition_type="fade",
    duration=1.5
)
```

### 차트 데이터 업데이트
```python
mcp_ppt_update_chart_data(
    slide_index=1,
    shape_index=0,
    categories=["Jan", "Feb", "Mar", "Apr"],
    series_data=[
        {"name": "매출", "values": [100, 120, 140, 160]},
        {"name": "비용", "values": [80, 90, 100, 110]}
    ]
)
```

### 슬라이드 마스터 관리
```python
mcp_ppt_manage_slide_masters(
    operation="list"
)
```

---

## 설정 및 구성

### 환경 변수
- `PPT_TEMPLATE_PATH`: 템플릿 파일 경로
- `PYTHONPATH`: Python 경로 설정

### 지원 파일 형식
- 입력: PPTX, PPT
- 출력: PPTX
- 이미지: JPG, PNG, GIF, BMP
- 템플릿: PPTX

---

## 팁과 모범 사례

1. **일관성 유지**: 동일한 색상 스키마를 전체 프레젠테이션에 적용
2. **가독성**: 적절한 폰트 크기와 여백 사용
3. **시각적 계층**: 제목, 부제목, 본문의 명확한 구분
4. **간결함**: 슬라이드당 핵심 메시지 하나씩
5. **브랜딩**: 회사 로고와 색상 활용

---

## 문제 해결

### 일반적인 문제
- **이미지가 표시되지 않음**: 파일 경로 확인
- **폰트 문제**: 시스템에 설치된 폰트 사용
- **색상 차이**: RGB 값으로 정확한 색상 지정

### 지원 및 문의
프로젝트 GitHub 저장소에서 이슈를 등록하거나 문서를 참조하세요.
