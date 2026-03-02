---
name: lg-pptx
description: 사내 보고서·발표자료 작성 시간을 획기적으로 줄여주는 PPTX 자동 생성 스킬입니다. 에이투지체 폰트와 사내 디자인 컨벤션(액센트 블록, L-브래킷 장식, 컬러 시스템)을 자동 적용하여 직원들이 콘텐츠에만 집중할 수 있게 합니다. "PPT 만들어줘", "발표자료 만들어줘", "보고서 PPT", "주간보고 슬라이드", "제안서 만들어줘" 같은 요청에 트리거됩니다. 한국어와 영어 프레젠테이션 모두 지원합니다.
---

# 사내 PPTX 자동 생성 스킬 (에이투지체)

사내 디자인 컨벤션을 따르는 프레젠테이션을 자동 생성하여, 직원들의 보고서 작성 시간을 줄이고 일하는 방식을 혁신하는 스킬입니다.

## Quick Reference

| 작업 | 방법 |
|------|------|
| 새 프레젠테이션 생성 | `scripts/lg_pptx_builder.py` 임포트 후 `LGPresentation` 클래스 사용 |
| 디자인 토큰 확인 | `references/design-tokens.md` 참조 |
| 커스텀 레이아웃 | 빌더 메서드 조합으로 자유 구성 |

## 사전 준비: 에이투지체 폰트 설치

이 스킬은 **에이투지체(A2Z)** 폰트가 시스템에 설치되어 있어야 정상 렌더링됩니다.
폰트가 없으면 맑은 고딕으로 대체됩니다.

**다운로드 링크:**
- 공식 배포: https://freesentation.blog/a2z
- 눈누 폰트 플랫폼: https://noonnu.cc/font_page/1778

**설치 방법:**
1. 위 링크에서 폰트 파일(.ttf/.otf)을 다운로드
2. macOS: 다운로드한 파일을 더블클릭 → "서체 설치" 버튼 클릭
3. Windows: 다운로드한 파일 우클릭 → "설치" 또는 "모든 사용자용으로 설치"

**설치 확인 (터미널):**
```bash
# macOS/Linux
fc-list | grep -i "에이투지체"
# 출력 예: /Users/.../에이투지체 4 Regular.ttf: 에이투지체 4 Regular:style=Regular
```

> 라이선스: 에이투지체는 OFL(Open Font License)로 상업적·개인적 용도 모두 무료 사용 가능합니다.

## Workflow

### Step 1: 의존성 확인

```bash
pip install python-pptx Pillow
```

### Step 2: 프레젠테이션 생성

빌더 스크립트를 사용해 프레젠테이션을 생성합니다. 아래는 기본 패턴입니다:

```python
import sys
sys.path.insert(0, '<skill-scripts-path>')
from lg_pptx_builder import LGPresentation

# 프레젠테이션 생성
prs = LGPresentation(
    font_name="에이투지체",     # 기본 폰트 (한글+영문 모두)
    font_name_latin="에이투지체",  # 라틴 폰트도 동일
    logo_path=None              # LG 로고 이미지 경로 (선택)
)

# 슬라이드 추가
prs.add_cover("프로젝트 제목", subtitle="팀명", date="2025.01.01")
prs.add_toc([
    ("Summary", []),
    ("시스템 소개", ["항목 1", "항목 2", "항목 3"]),
    ("첨부자료", [])
])
prs.add_content("1.1 시스템 개요", section="Summary", bullets=["내용1", "내용2"])
prs.add_roadmap(
    title="[SPC] 로드맵",
    section="로드맵",
    subtitle="설명 텍스트",
    years=["(2025) Phase 1", "(2026) Phase 2", "(2027) Phase 3"],
    roadmap_items={...},
    table_data={...}
)
prs.save("output.pptx")
```

### Step 3: QA 검증

생성된 PPTX를 검증합니다:
1. `markitdown` 으로 텍스트 추출하여 내용 확인
2. 가능하면 `soffice` → `pdftoppm` 으로 이미지 변환하여 시각적 확인

## Design System Overview

LG 그룹 프레젠테이션의 핵심 디자인 요소입니다. 상세 토큰은 `references/design-tokens.md`를 참조하세요.

### Color Palette

| 역할 | 색상 | HEX | 용도 |
|------|------|-----|------|
| Primary | LG RED | `#A50034` | 브래킷, 액센트 바, TOC 번호, 강조 |
| Text Primary | Black | `#000000` | 제목, 본문 |
| Text Secondary | Dark Gray | `#333333` | 부제목, 보조 텍스트 |
| Text Tertiary | Medium Gray | `#666666` | 섹션명, 캡션 |
| Background | White | `#FFFFFF` | 슬라이드 배경 |
| Surface | Light Gray | `#F2F2F2` | 콘텐츠 박스 배경 |
| Header Bar | Charcoal | `#3C3C3C` | 타임라인 헤더, 테이블 헤더 |
| Accent Green | Green | `#2E7D32` | 미래/계획 항목 |
| Accent Orange | Orange | `#D4760A` | 하이라이트 항목 |

### Typography

- **폰트**: 에이투지체 — 한글+영문 모두 동일 폰트 적용
- 시스템에 폰트가 없는 경우 "맑은 고딕" 또는 "Malgun Gothic" 사용
- East Asian 폰트 설정 필수 (python-pptx에서 `<a:ea>` XML 요소 직접 설정)

| 용도 | 크기 | 굵기 | 색상 |
|------|------|------|------|
| 표지 제목 | 32pt | Bold | Black |
| 표지 부제 | 14pt | Regular | Dark Gray |
| 섹션 제목 | 24pt | Bold | Black |
| 본문 제목 | 18pt | Bold | Black |
| 본문 | 12pt | Regular | Black/Dark Gray |
| 표 헤더 | 10pt | Bold | White |
| 표 본문 | 10pt | Regular | Black |
| 캡션/주석 | 9pt | Regular | Medium Gray |

### Slide Types

#### 1. Cover (표지)
- 흰 배경
- **좌상단 L-브래킷**: LG RED, 두께 ~0.4cm, 팔 길이 ~2.5cm
- **우하단 L-브래킷**: LG RED, 180도 회전 (대칭)
- 중앙: 제목 (Bold, 32pt, Black)
- 하단 중앙: 부제 + 날짜 (14pt, Dark Gray)
- 우하단 (브래킷 안): LG 로고 (선택)

#### 2. Table of Contents (목차)
- 흰 배경
- 상단: 얇은 회색 가로선
- 좌측: "Contents" 텍스트 (28pt, Black) + 짧은 빨간 밑줄 바 (~3cm)
- 아래: 회색 구분선
- 목차 항목: 로마 숫자 (LG RED, Bold) + 항목명
- 하위 항목: 들여쓰기 + dash prefix (Black)

#### 3. Content (내용 슬라이드)
- 흰 배경
- **좌상단 액센트 블록**: LG RED 작은 사각형 (0.5cm x 1.4cm), 제목 왼쪽에 위치
- 제목: 액센트 블록 오른쪽에 번호+제목 (Bold, 24pt, Black) — 예: "1.1 공정 DX 서비스"
- 우상단: 섹션명 (12pt, Medium Gray) + 빨간 원형 인디케이터
- 본문 영역: 자유 구성 (텍스트, 박스, 표 등)

#### 4. Roadmap (로드맵)
- Content 슬라이드 기본 구조 유지 (좌측 레드 바, 제목, 섹션명)
- 부제: 설명 텍스트 (14pt, Dark Gray)
- **타임라인 헤더**: 다크 차콜 (#3C3C3C) 쉐브론/화살표, 연도 + 설명 (White, Bold)
- **좌측 라벨 블록**: 다크 레드 (#A50034) 세로 블록, 라벨 텍스트 (White, Bold)
- **콘텐츠 그리드**: 연도별 컬럼, 라이트 그레이 배경 셀
- 텍스트 색상: 검정(기본), 초록(미래 계획), 주황(하이라이트)
- **하단 테이블** (선택): 비교 표

#### 5. Comparison Table (비교 테이블)
- Content 슬라이드 기본 구조
- 다크 차콜 헤더 행 (White Bold 텍스트)
- 얇은 회색 테두리
- 교대 행 배경 (White / Light Gray)

#### 6. Summary Matrix (요약 매트릭스)
- Content 슬라이드 기본 구조 (액센트 블록 + 제목)
- **좌측 2열**: 카테고리(세로 병합) + 서브라벨 (회색 배경)
- **상단 헤더**: 다크 차콜 배경, 흰색 텍스트 (계열사/항목명)
- 셀 내용: 좌측 정렬, 연도별 불릿 텍스트
- 얇은 회색 테두리

#### 7. Two Column (2단 레이아웃)
- Content 슬라이드 기본 구조
- 좌우 2개 컬럼, 각각 제목 + 불릿 포인트
- 비교, Before/After, 장단점 분석에 적합

#### 8. KPI Cards (핵심 지표)
- Content 슬라이드 기본 구조
- 가로 나란히 배치된 카드들 (Light Gray 배경)
- 큰 숫자 (40pt, 색상 커스텀 가능) + 라벨 텍스트
- 경영진 보고, 성과 요약에 적합

#### 9. Timeline (타임라인)
- Content 슬라이드 기본 구조
- 가로 타임라인 선 + 빨간 원형 마커
- 날짜 (상단, LG RED) + 제목/설명 (하단)
- 프로젝트 일정, 마일스톤 표현

#### 10. Process Flow (프로세스 흐름)
- Content 슬라이드 기본 구조
- 가로 정렬된 단계 박스 (차콜 헤더 + 라이트 그레이 본문)
- 단계 사이 화살표 연결
- 워크플로우, 시스템 아키텍처 개요

#### 11. SWOT Analysis (SWOT 분석)
- Content 슬라이드 기본 구조
- 2x2 그리드: 강점(RED), 약점(CHARCOAL), 기회(GREEN), 위협(ORANGE)
- 각 사분면: 컬러 헤더 + 라이트 그레이 본문 + 불릿

## Key Design Rules

1. **일관성**: 모든 내용 슬라이드에 좌상단 빨간 액센트 블록 적용
2. **여백**: 상하 1.2cm, 좌측 1.8cm (액센트 블록 이후), 우측 0.8cm 여백 유지
3. **색상 절제**: LG RED는 강조에만 사용, 과용 금지
4. **계층 구조**: 제목 → 부제 → 본문 크기/굵기 차이로 시각적 계층 표현
5. **박스 스타일**: 둥근 모서리 없음 (직각 사각형), 라이트 그레이 배경
6. **라벨 뱃지**: 카테고리 라벨은 LG RED 배경 + White 텍스트의 작은 사각형
7. **표 스타일**: 헤더는 다크 차콜 배경, 테두리는 얇은 회색

## Builder API Reference

`scripts/lg_pptx_builder.py`의 `LGPresentation` 클래스는 다음 메서드를 제공합니다:

### 생성자
```python
LGPresentation(font_name="에이투지체", font_name_latin="에이투지체", logo_path=None)
```

### 슬라이드 메서드

| 메서드 | 설명 |
|--------|------|
| `add_cover(title, subtitle, date, logo_path)` | 표지 슬라이드 |
| `add_toc(items)` | 목차 슬라이드 |
| `add_section_divider(number, title)` | 섹션 구분 슬라이드 |
| `add_content(title, section, body, bullets)` | 일반 내용 슬라이드 |
| `add_roadmap(title, section, subtitle, years, roadmap_items, table_data)` | 로드맵 슬라이드 |
| `add_table(title, section, headers, rows)` | 테이블 슬라이드 |
| `add_summary_matrix(title, section, headers, row_groups)` | 요약 매트릭스 (카테고리 병합 테이블) |
| `add_two_column(title, section, left_title, left_bullets, right_title, right_bullets)` | 2단 레이아웃 |
| `add_kpi_cards(title, section, kpis)` | KPI/핵심 지표 카드 |
| `add_timeline(title, section, milestones)` | 타임라인 슬라이드 |
| `add_process_flow(title, section, steps)` | 프로세스 흐름도 |
| `add_swot(title, section, strengths, weaknesses, opportunities, threats)` | SWOT 분석 |
| `add_architecture(title, section, subtitle, columns, rows)` | 멀티컬럼 아키텍처/시스템 구조도 |
| `add_strategy_pillars(title, section, subtitle, pillars)` | 전략 필러 (3~5개 수직 컬럼) |
| `add_risk_matrix(title, section, subtitle, risks)` | 3x3 리스크 평가 매트릭스 |
| `add_financial_summary(title, section, subtitle, categories)` | 투자/예산 요약 테이블 (소계+합계) |
| `add_milestone_tracker(title, section, subtitle, phases)` | 마일스톤 추적 (상태별 색상) |
| `add_comparison_cards(title, section, subtitle, cards)` | 솔루션/옵션 비교 카드 |
| `save(filename)` | PPTX 파일 저장 |

### 헬퍼 메서드 (내부)

| 메서드 | 설명 |
|--------|------|
| `_set_font(run)` | Latin + EA 폰트 동시 설정 |
| `_add_l_bracket(slide, corner, arm_len, thickness, color)` | L-브래킷 장식 |
| `_add_accent_bar(slide)` | 좌측 빨간 액센트 바 |
| `_add_section_indicator(slide, section_name)` | 우상단 섹션명 + 빨간 점 |
| `_add_slide_title(slide, title)` | 제목 텍스트 추가 |

## Common Patterns

### 내용이 많은 슬라이드

콘텐츠가 한 슬라이드에 다 안 들어갈 때:
- 같은 섹션 제목으로 여러 슬라이드 분할
- 제목에 "(1/2)", "(2/2)" 등 페이지 표시
- 각 슬라이드에 동일한 액센트 바 + 섹션 인디케이터 유지

### 다이어그램/아키텍처 슬라이드

복잡한 다이어그램은 python-pptx의 기본 도형으로 구성:
- 사각형 (`add_shape`) + 텍스트로 블록 구성
- 화살표/커넥터로 연결
- 카테고리 라벨: LG RED 배경 소형 사각형
- 콘텐츠 블록: Light Gray 배경 사각형

### 로드맵 구성

로드맵 슬라이드의 `roadmap_items` 파라미터 구조:
```python
roadmap_items = {
    "label": "시스템 로드맵",  # 좌측 라벨
    "rows": [
        {
            "items_by_year": [
                # Year 1 items
                [{"text": "항목 1", "tag": "LGES", "tag_color": "#1565C0"}],
                # Year 2 items
                [{"text": "항목 2", "color": "green"}],
                # Year 3 items
                [{"text": "항목 3", "color": "orange"}]
            ]
        }
    ]
}
```
