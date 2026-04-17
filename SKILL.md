---
name: lg-pptx
description: 사내 보고서·발표자료 작성 시간을 획기적으로 줄여주는 PPTX 자동 생성 스킬입니다. 에이투지체 폰트와 사내 디자인 컨벤션(액센트 블록, L-브래킷 장식, 컬러 시스템)을 자동 적용하여 직원들이 콘텐츠에만 집중할 수 있게 합니다. 33종 슬라이드 타입(표지, 목차, 내용, 로드맵, 테이블, SWOT, KPI, 타임라인, 프로세스, 간트차트, 조직도, 피라미드, 포지셔닝맵, 키워드강조, 스윔레인, 제언, 멀티컬럼, 카드그리드, 사이드바, 타이틀컬럼, 아이콘프로세스 등)과 matplotlib 차트/이미지 삽입을 지원합니다. **플래닝 모드 내장**: 5장 이상 요청 시 자동으로 디스커버리 인터뷰(5가지 강제 질문) → 슬라이드 아키텍처 설계 → 콘텐츠 개발 → 7-Pass 디자인 리뷰 → 생성 순서로 진행하여 알찬 내용의 PPT를 만듭니다. "PPT 만들어줘", "발표자료 만들어줘", "보고서 PPT", "주간보고 슬라이드", "제안서 만들어줘", "간트차트 PPT", "조직도 슬라이드", "차트 넣어줘", "이미지 PPT", "SWOT 분석 PPT", "KPI 대시보드", "프로젝트 일정표", "비교 분석 자료", "전략 보고서", "리스크 매트릭스", "스윔레인", "업무 프로세스 다이어그램", "역할별 흐름도", "제언 슬라이드", "PPT 구성안", "슬라이드 기획" 같은 요청에 트리거됩니다. 한국어와 영어 프레젠테이션 모두 지원합니다.
---

# 사내 PPTX 자동 생성 스킬 (에이투지체)

사내 디자인 컨벤션을 따르는 프레젠테이션을 자동 생성하여, 직원들의 보고서 작성 시간을 줄이고 일하는 방식을 혁신하는 스킬입니다.

## Quick Reference

| 작업 | 방법 |
|------|------|
| 새 프레젠테이션 생성 | `scripts/lg_pptx_builder.py` 임포트 후 `LGPresentation` 클래스 사용 |
| 디자인 토큰 확인 | `references/design-tokens.md` 참조 |
| 커스텀 레이아웃 | 빌더 메서드 조합으로 자유 구성 |

## 멀티 테마 지원 (LG / 한화)

이 스킬은 **LG 테마**(기본)와 **한화 테마**를 모두 지원합니다.

```python
# LG 테마 (기본 — 기존과 동일)
prs = LGPresentation()  # theme="lg" 기본값

# 한화 테마
prs = LGPresentation(theme="hanwha")
```

### 테마별 차이 요약

| 요소 | LG (`theme="lg"`) | 한화 (`theme="hanwha"`) |
|------|---------------------|-------------------------|
| Primary 색상 | LG RED `#A50034` | Hanwha Orange `#F37321` |
| 폰트 | 에이투지체 | 한화고딕 / 한화체 |
| 표지 장식 | L-브래킷 | 하단 오렌지 바 |
| 테이블 헤더 | Charcoal `#3C3C3C` | Navy `#1D1E37` |
| 뱃지/라벨 | Dark Red `#C00000` | Light Navy `#353968` |
| 배경 서피스 | Light Gray `#F2F2F2` | Warm Grey `#EFEEE8` |

### 한화 테마 사용 예시

```python
import sys
sys.path.insert(0, '<skill-scripts-path>')
from lg_pptx_builder import LGPresentation

prs = LGPresentation(theme="hanwha", logo_path="logo.png")
prs.add_cover("프로젝트 제목", subtitle="조직/부서명", date="2026.04")
prs.add_toc([("개요", []), ("설계", ["시스템 구성", "데이터 흐름"])])
prs.add_content("1.1 시스템 개요", section="개요", chapter="I. 개요",
                bullets=["시스템 항목 1", "시스템 항목 2"])
prs.save("output_hanwha.pptx")
```

### 한화 테마 폰트 요구사항

한화 테마 사용 시 **한화고딕 / 한화체** 폰트가 필요합니다:
- `한화 B_OTF` (Bold) — 제목용
- `한화고딕 R_OTF` (Regular) — 본문용
- `한화고딕 B_OTF` (Bold) — 강조용
- `한화고딕 L_OTF` (Light) — 보조 텍스트

폰트 미설치 시 맑은 고딕으로 대체됩니다.

---

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

## Planning Mode (기획 모드)

**PPT를 만들기 전에, PPT를 설계한다.** 5장 이상의 프레젠테이션 요청 시 자동으로 기획 모드가 활성화됩니다.

> 상세 프레임워크: `references/planning-mode.md` 참조

### 기획 모드 흐름

```
Phase 1: 디스커버리 인터뷰 (5가지 강제 질문)
    ↓ 목적·청중·메시지·소스·스타일 확정
Phase 2: 슬라이드 아키텍처 (구조표 작성)
    ↓ 슬라이드별 제목·타입·메서드·밀도·내용 매핑
Phase 3: 콘텐츠 개발 (실제 텍스트 채우기)
    ↓ 불릿·테이블·수치 등 구체적 내용 완성
Phase 4: 디자인 리뷰 (7-Pass 검증)
    ↓ 정보구조·시각다양성·충실도·청중적합성·메서드적합성·흐름·실행가능성
Phase 5: 생성 및 QA
    ↓ 코드 생성 → 실행 → markitdown 검증
```

### Phase 1: 디스커버리 — 5가지 강제 질문

| # | 질문 | Push 예시 |
|---|------|-----------|
| Q1 | 청중은 누구이고, 발표 후 어떤 행동을 기대하나요? | "경영진 보고" → "어떤 결정을 내리길 원하는지 구체적으로" |
| Q2 | 이 PPT를 한 문장으로 요약하면? | "AI 도입 필요" → "왜 지금이고, 안 하면 어떤 리스크?" |
| Q3 | 이미 있는 자료가 있나요? | "없다" → "반드시 포함할 수치/사실은?" |
| Q4 | 분량, 발표 시간, 필수 섹션은? | "적당히" → "10분=12장, 30분=25장, 읽기용=30장+" |
| Q5 | 비주얼 임팩트 vs 데이터 중심 vs 균형? | "깔끔하게" → "(A) 인포그래픽 (B) 빽빽 (C) 균형" |

**스마트 라우팅**: 이미 알고 있는 정보의 질문은 생략. 전부 명확하면 Phase 2로 직행.
**인터뷰 규칙**: 한 번에 한 질문, 금지어("interesting", "여러 방법"), 항상 추천안 먼저 제시.

### Phase 2: 슬라이드 아키텍처 — 구조표

```markdown
| # | 제목 | 타입 | 메서드 | Density | 핵심 내용 |
|---|------|------|--------|---------|-----------|
| 1 | 프로젝트명 | 표지 | add_cover | - | 제목, 팀명, 날짜 |
| 2 | Contents | 목차 | add_toc | - | 4개 섹션 |
| 3 | I. 배경 | 섹션구분 | add_section_divider | - | 섹션 1 |
| 4 | 1.1 현황 | 내용 | add_content | normal | 현황 불릿 5개 |
| 5 | 1.2 비교 | 2단 | add_two_column | normal | As-Is vs To-Be |
```

**설계 원칙:**
- **내러티브 흐름**: 문제-해결, 현황-전략, 보고형, 교육형, 비교형 중 선택
- **레이아웃 다양성**: 같은 레이아웃 3장 연속 금지, 불릿 40% 이하, 시각적 슬라이드 30% 이상
- **밀도 일관성**: 같은 섹션 내 동일 density 유지

### Phase 4: 디자인 리뷰 — 7-Pass 체크리스트

| Pass | 검증 항목 | 핵심 기준 |
|------|-----------|-----------|
| 1 | 정보 아키텍처 | 논리적 흐름, 1슬라이드=1메시지 |
| 2 | 시각적 다양성 | 3연속 금지, 시각 슬라이드 30%+ |
| 3 | 콘텐츠 충실도 | 빈 슬라이드 없음, 구체적 수치 |
| 4 | 청중 적합성 | 용어 수준, "so what?" 테스트 |
| 5 | 메서드 적합성 | 내용↔메서드 최적 매칭 |
| 6 | 흐름과 전환 | 섹션 구분, chapter 일관성 |
| 7 | 실행 가능성 | 메서드 지원 여부, 파라미터 준비 |

### 메서드 매핑 가이드 (Quick Reference)

| 표현하고 싶은 것 | 추천 메서드 | density |
|------------------|-------------|---------|
| 핵심 수치/KPI | `add_kpi_cards` | spacious |
| 장단점/비교 | `add_two_column`, `add_comparison_cards` | normal |
| 3~4개 전략 축 | `add_titled_columns`, `add_strategy_pillars` | normal |
| 상세 데이터 테이블 | `add_table` | compact/dense |
| 일정/로드맵 | `add_roadmap`, `add_gantt_chart`, `add_timeline` | normal |
| 프로세스/워크플로우 | `add_process_flow`, `add_icon_process`, `add_swimlane` | normal |
| 조직/계층 구조 | `add_org_chart`, `add_pyramid` | normal |
| 분석 (SWOT/리스크) | `add_swot`, `add_risk_matrix` | normal |
| 카드형 항목 | `add_card_grid`, `add_multi_column` | normal |
| 메인+보조 패널 | `add_content_sidebar` | normal/compact |
| 핵심 키워드/비전 | `add_keyword_highlight` | spacious |
| 시장 포지셔닝 | `add_positioning_map` | spacious |
| 재무/투자 요약 | `add_financial_summary` | compact |
| 권고/제언 | `add_recommendation` | normal |

### 약식 모드

5장 이하 또는 특정 슬라이드 타입만 요청 시: 인터뷰 생략, 바로 생성.

---

## Workflow

### Step 1: 의존성 확인

```bash
pip install python-pptx Pillow matplotlib
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
prs.add_content("1.1 시스템 개요", section="Summary", chapter="I. Summary", bullets=["내용1", "내용2"])
prs.add_roadmap(
    title="[프로젝트명] 로드맵",
    section="로드맵",
    chapter="II. 로드맵",
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
| Primary | LG RED | `#A50034` | 브래킷, 액센트 바, 챕터명, 구분선 |
| Badge Red | DARK RED | `#C00000` | 서브헤더 뱃지, 라벨 뱃지 |
| Text Primary | Black | `#000000` | 제목, 본문 |
| Text Secondary | Dark Gray | `#333333` | 부제목, 보조 텍스트 |
| Text Tertiary | Medium Gray | `#666666` | 태그명, 캡션 |
| Background | White | `#FFFFFF` | 슬라이드 배경, 콘텐츠 박스 |
| Surface | Light Gray | `#F2F2F2` | 보조 배경, 테이블 교대행 |
| Header Bar | Charcoal | `#3C3C3C` | 타임라인 헤더, 테이블 헤더 |
| Accent Green | Green | `#2E7D32` | 미래/계획 항목 |
| Accent Orange | Orange | `#D4760A` | 하이라이트 항목 |
| Accent Blue | Blue | `#1565C0` | 보조 강조 |

### Typography

- **폰트**: 에이투지체 — 한글+영문 모두 동일 폰트 적용
- 시스템에 폰트가 없는 경우 "맑은 고딕" 또는 "Malgun Gothic" 사용
- East Asian 폰트 설정 필수 (python-pptx에서 `<a:ea>` XML 요소 직접 설정)

| 용도 | Spacious | Normal | Compact | Dense | 굵기 | 색상 |
|------|----------|--------|---------|-------|------|------|
| 표지 제목 | 32pt | 32pt | 32pt | 32pt | Bold | Black |
| 표지 부제 | 14pt | 14pt | 14pt | 14pt | Regular | Dark Gray |
| 섹션 제목 | 28pt | 28pt | 28pt | 28pt | Bold | Black |
| 본문 제목 | 16pt | 16pt | 16pt | 16pt | Bold | Black |
| 본문/불릿 | 16pt | 12pt | 10pt | 9pt | Regular | Black/Dark Gray |
| 표 헤더 | 12pt | 10pt | 9pt | 8pt | Bold | White |
| 표 본문 | 11pt | 9pt | 8pt | 7pt | Regular | Black |
| 캡션/주석 | 10pt | 9pt | 8pt | 7pt | Regular | Medium Gray |

#### 밀도(density) 선택 가이드

| 밀도 | 용도 | 예시 |
|------|------|------|
| `"spacious"` | 인포그래픽, KPI, 비전/키워드 | 불릿 3개 이하, 핵심 메시지, 큰 숫자/차트 |
| `"normal"` (기본) | 일반 발표자료, 교육자료 | 불릿 5개 이하, 3~4열 테이블 |
| `"compact"` | 컨설팅 보고서, 주간보고 | 불릿 10개, 5~6열 테이블, 셀당 1~2줄 |
| `"dense"` | 제안서, 구축범위, 투입인력표 | 불릿 15개+, 8~9열 테이블, 셀당 2~3줄 |

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

#### 3. Content (내용 슬라이드) — L-Style Chrome
- 흰 배경, 모든 내용 슬라이드에 공통 적용되는 L-Style 크롬:
  - **챕터명** (좌상단): LG RED, 10pt Bold — 예: "I. 기본 슬라이드"
  - **제목** (좌측): 24pt Bold, Black — 예: "1.1 시스템 개요"
  - **액센트 바**: LG RED 직사각형 (0.4cm x 1.0cm), 제목 왼쪽 (0.67, 2.30)
  - **부제목**: 16pt SemiBold, Dark Gray
  - **빨간 구분선**: LG RED 가로선 (y=4.87cm), 헤더와 콘텐츠 영역 분리
  - **태그** (우상단): 11pt, Medium Gray, 우측 정렬 + 빨간 원형 인디케이터
- 콘텐츠 영역: y≈5.62cm부터 시작, **둥근 모서리 사각형**(roundRect) + 흰 배경
- `chapter` 파라미터로 챕터명 설정

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

#### 12. Gantt Chart (간트 차트)
- Content 슬라이드 기본 구조
- 좌측: 태스크명 + 담당자 목록, 우측: 월별 타임라인 그리드
- 바 차트: 태스크별 컬러 바 (시작~종료), 완료율 표시
- 마일스톤 마커 (빨간 다이아몬드)
- 프로젝트 일정 관리, WBS 표현에 적합

#### 13. Org Chart (조직도)
- Content 슬라이드 기본 구조
- 최상위 노드 (LG RED 배경) → 하위 노드 (차콜 배경) 계층 구조
- 수직 커넥터 라인으로 연결
- 각 노드: 직책/이름 + 부서/역할 텍스트
- 조직 구조, 보고 체계 시각화

#### 14. Pyramid (피라미드 다이어그램)
- Content 슬라이드 기본 구조
- 위에서 아래로 넓어지는 사다리꼴 레이어
- 각 레이어: 컬러 배경 + 제목/설명 텍스트
- 전략 계층, 가치 피라미드, 우선순위 표현

#### 15. Positioning Map (포지셔닝 맵)
- Content 슬라이드 기본 구조
- X/Y 축 교차 2D 맵 + 사분면 라벨
- 항목별 원형 마커 (크기/색상 커스텀)
- 경쟁사 분석, 시장 포지셔닝, BCG 매트릭스

#### 16. Keyword Highlight (키워드 강조)
- Content 슬라이드 기본 구조
- 태그 클라우드 스타일: 키워드별 크기/색상/굵기 차등
- 하단 설명 텍스트
- 핵심 메시지, 비전/미션, 키워드 요약

#### 17. Swimlane (스윔레인 프로세스)
- Content 슬라이드 기본 구조
- 역할/부서별 수평 레인 (수영장 레인 스타일)
- 레인 내 프로세스 단계 박스 배치 (라운드 사각형)
- 같은 레인 이동: 수평 화살표, 다른 레인 이동: L자형 커넥터
- 레인별 고유 색상 (RED, CHARCOAL, BLUE, GREEN, ORANGE, PURPLE)
- 업무 프로세스, R&R 구분, 시스템 간 연동 흐름 표현

#### 18. Recommendation (제언)
- Content 슬라이드 기본 구조
- 번호 원형 (LG RED) + 제목 + 상세 설명
- 클로징 직전 배치, 핵심 제언/권고사항 정리
- 문자열 리스트 또는 {title, detail} 딕셔너리 리스트 지원

#### 19. Multi-Column (다중 컬럼 레이아웃)
- Content 슬라이드 기본 구조
- N개 컬럼 레이아웃 (1~4+) with 유연한 비율 지정
- 각 컬럼: 라운드 사각형 박스 (LIGHT_GRAY) + 타이틀 필 + 불릿/본문
- 비율 예: [7,3] (70/30), [3,7] (30/70), [1,1] (50/50), [1,1,1,1] (4등분)

#### 20. Card Grid (카드 그리드)
- Content 슬라이드 기본 구조
- 2~4열 카드 그리드, 각 카드에 타이틀 필/아이콘/아이템/푸터
- 원형 아이콘 (선택), 스택 아이템 필 (BORDER_GRAY 테두리)
- 제품 비교, 기능 소개, 팀 소개에 적합

#### 21. Content Sidebar (콘텐츠+사이드바)
- Content 슬라이드 기본 구조
- 메인 콘텐츠 + N개 스택 사이드바 패널
- 사이드바 위치: 좌/우 선택, 비율 조정 가능
- 대시보드, 상세+요약, KPI+설명 조합에 적합

#### 22. Titled Columns (타이틀바+컬럼)
- Content 슬라이드 기본 구조
- 전체 너비 타이틀 바 + 하단 N개 컬럼
- 각 컬럼: 타이틀 필 + 아이콘 + 불릿 + 하단 아이템 필
- 전략 필러, 비교 분석, 기능 분류에 적합

#### 23. Icon Process (아이콘 프로세스 흐름)
- Content 슬라이드 기본 구조
- 상단: N개 아이콘 원 + 화살표 연결 (프로세스 스텝)
- 하단: 2개 상세 패널 (타이틀 필 + 항목)
- 워크플로우, 파이프라인, 변환 프로세스 표현

#### 24. Chart Slide (차트 이미지)
- Content 슬라이드 기본 구조
- 외부 차트 이미지 파일(.png/.jpg) 중앙 배치
- 하단 캡션 텍스트
- matplotlib/seaborn 등으로 생성한 차트 삽입

#### 18. Image Slide (이미지 배치)
- Content 슬라이드 기본 구조
- 1~4장 이미지 자동 레이아웃 (1장=전체, 2장=좌우, 3장+=그리드)
- 이미지별 캡션 텍스트
- 스크린샷, 사진, 다이어그램 배치

#### 19. Matplotlib Chart (직접 삽입)
- Content 슬라이드 기본 구조
- matplotlib Figure 객체를 직접 전달 → 자동 렌더링
- 임시 파일 생성 후 삽입, 자동 정리
- 데이터 분석 결과 즉시 슬라이드화

## Key Design Rules (L-Style)

1. **L-Style 크롬**: 모든 내용 슬라이드에 챕터명 → 제목 → 부제목 → 빨간 구분선 → 콘텐츠 순서
2. **챕터명 필수**: `chapter` 파라미터로 좌상단 빨간 챕터명 표시 (10pt, RED, Bold)
3. **콘텐츠 박스 스타일**: 둥근 모서리(roundRect) + **테두리 없음** + 음영(LIGHT_GRAY) + 그림자(card shadow)로 입체감 표현. 컬러 테두리 사용하지 않음
4. **색상 체계**: LG RED(#A50034)는 크롬 요소, DARK RED(#C00000)는 서브헤더 뱃지
5. **계층 구조**: 챕터명(10pt) → 제목(24pt) → 부제목(16pt) → 본문(12pt) 계층
6. **빨간 구분선**: y=4.87cm 위치에 가로 RED 라인으로 헤더-콘텐츠 분리
7. **라벨 뱃지**: 카테고리 라벨은 DARK RED(#C00000) 배경 + White 텍스트의 둥근 사각형
8. **표 스타일**: 헤더는 다크 차콜 배경, 테두리는 얇은 회색
9. **차트/이미지**: matplotlib Figure 직접 삽입 또는 이미지 파일 배치 지원

## 자유 레이아웃 가이드 (Free-form Layout)

**원칙: 내용에 맞는 레이아웃을 설계한다. 기존 메서드에 내용을 끼워 맞추지 않는다.**

### 레이아웃 결정 프로세스

1. **내용 분석**: 슬라이드에 들어갈 정보의 구조와 양을 먼저 파악
2. **레이아웃 선택**: 기존 고수준 메서드가 딱 맞으면 사용, 아니면 자유 배치
3. **글자 초과 방지**: 컨테이너 크기를 내용에 맞게 조절 (고정 크기에 내용을 억지로 넣지 않음)

### 자유 배치 방법 (add_blank_content 활용)

```python
# 1. 크롬만 있는 빈 슬라이드 생성 — (slide, y_pos) 튜플 반환
slide, y_pos = prs.add_blank_content(
    title="슬라이드 제목", section="섹션명", chapter="챕터명",
    subtitle="부제목"
)

# 2. y_pos 아래 영역에 자유롭게 도형/텍스트 배치
# 저수준 메서드 사용:
prs.add_box(slide, left, top, width, height, text="내용", bg_color=..., shadow=True)
prs.add_label_badge(slide, left, top, "라벨")
prs._add_textbox(slide, left, top, width, height, text="텍스트", size=Pt(12))

# 3. 또는 python-pptx 직접 사용:
from pptx.util import Cm, Pt
from pptx.enum.shapes import MSO_SHAPE
box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Cm(1), y_pos, Cm(15), Cm(8))
```

### 콘텐츠 영역 좌표 참고

| 항목 | 값 | 설명 |
|------|-----|------|
| 콘텐츠 시작 Y | `y_pos` (≈ Cm(5.62)) | `add_blank_content` 반환값 사용 |
| 좌측 여백 | Cm(1.33) | 콘텐츠 좌측 시작 |
| 슬라이드 폭 | Inches(13.333) ≈ Cm(33.87) | 16:9 기준 |
| 슬라이드 높이 | Inches(7.5) ≈ Cm(19.05) | 16:9 기준 |
| 사용 가능 높이 | 약 Cm(12.5) | y_pos ~ 하단 여백 |
| 사용 가능 폭 | 약 Cm(31.5) | 좌우 여백 제외 |

### 컨테이너 스타일 규칙

- **테두리**: 없음 (투명). 컬러 테두리로 박스를 구분하지 않음
- **배경**: LIGHT_GRAY 음영으로 영역 구분
- **입체감**: `shadow_type="card"` 그림자로 깊이감 표현
- **모서리**: roundRect 사용 (val 8000)

## Builder API Reference

`scripts/lg_pptx_builder.py`의 `LGPresentation` 클래스는 다음 메서드를 제공합니다:

### 생성자
```python
LGPresentation(font_name=None, font_name_latin=None, logo_path=None, theme="lg")
# theme: "lg" (기본) 또는 "hanwha"
# font_name: None이면 테마 기본값 사용 (LG=에이투지체, 한화=한화고딕)
```

### 슬라이드 메서드

| 메서드 | 설명 |
|--------|------|
| `add_cover(title, subtitle, date, logo_path)` | 표지 슬라이드 |
| `add_toc(items)` | 목차 슬라이드 |
| `add_section_divider(number, title)` | 섹션 구분 슬라이드 |
| `add_content(title, section, chapter, body, bullets, density)` | 일반 내용 슬라이드 |
| `add_roadmap(title, section, chapter, subtitle, years, roadmap_items, table_data)` | 로드맵 슬라이드 |
| `add_table(title, section, chapter, headers, rows, density, merge_column, col_widths)` | 테이블 슬라이드 |
| `add_summary_matrix(title, section, chapter, headers, row_groups)` | 요약 매트릭스 (카테고리 병합 테이블) |
| `add_two_column(title, section, left_title, left_bullets, right_title, right_bullets)` | 2단 레이아웃 |
| `add_multi_column(title, section, columns, ratio, density)` | N-컬럼 레이아웃 (유연한 비율) |
| `add_card_grid(title, section, cards, cols, show_icon, density)` | 카드 그리드 (아이콘+아이템) |
| `add_content_sidebar(title, section, main_body/bullets, sidebar_items, sidebar_position, sidebar_ratio, density)` | 콘텐츠+사이드바 |
| `add_titled_columns(title, section, bar_title, columns, show_bar, density)` | 타이틀바+컬럼 |
| `add_icon_process(title, section, steps, bottom_sections, density)` | 아이콘 프로세스 흐름 |
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
| `add_gantt_chart(title, section, subtitle, tasks, start_date, months)` | 간트 차트 (프로젝트 일정) |
| `add_org_chart(title, section, subtitle, org_data)` | 조직도 (계층 구조) |
| `add_pyramid(title, section, subtitle, levels)` | 피라미드 다이어그램 |
| `add_positioning_map(title, section, subtitle, x_label, y_label, items, quadrant_labels)` | 2D 포지셔닝 맵 |
| `add_keyword_highlight(title, section, subtitle, keywords, description)` | 키워드 강조/태그 클라우드 |
| `add_swimlane(title, section, subtitle, lanes, steps, connections)` | 스윔레인 프로세스 다이어그램 |
| `add_recommendation(title, section, subtitle, recommendations)` | 제언/권고사항 (클로징 전) |
| `add_chart_slide(title, section, subtitle, chart_path, caption)` | 차트 이미지 삽입 |
| `add_image_slide(title, section, subtitle, images)` | 이미지 배치 (1~4장 자동 레이아웃) |
| `add_matplotlib_chart(title, section, subtitle, fig, caption)` | matplotlib Figure 직접 삽입 |
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

### 고밀도 테이블 (제안서/구축범위 스타일)

컨설팅 자료나 제안서처럼 한 슬라이드에 빡빡하게 정보를 넣어야 할 때 `density` 파라미터를 사용합니다.

```python
# 구축범위 테이블 — 대분류 셀 병합 + 고밀도
prs.add_table(
    title="Phase 1 구축 범위", section="구축범위",
    chapter="VI. 구축 범위",
    density="dense",          # 7pt 본문, 최소 여백
    merge_column=0,           # 대분류 열 자동 병합
    row_alignment=PP_ALIGN.LEFT,
    col_widths=[5, 6, 16, 5],
    headers=["대분류", "중분류", "세부 내용", "산출물"],
    rows=[
        ["백엔드", "플랫폼 구축", "인프라 설치, 마스터/워커 배치", "서버 구성도"],
        ["백엔드", "워크플로우", "이벤트 기반 실시간 처리, 자동 재학습", "프로그램 설계서"],
        ["백엔드", "CI/CD 자동화", "버전관리·태깅·롤백, A/B 배포", "프로그램 설계서"],
        ["프론트엔드", "대시보드", "차트 라이브러리 + 데이터 그리드 UI", "화면설계서"],
        ["프론트엔드", "관리 화면", "목록/상세/배포, 시뮬레이터", "화면설계서"],
    ]
)
```

### 투입인력표 (9열 컴팩트 테이블)

```python
prs.add_table(
    title="4.2 투입 인력 상세", section="투자비용",
    density="compact",        # 8pt 본문, 적당한 여백
    col_widths=[1.5, 4, 1.5, 2, 2, 3, 2, 3.5, 12],
    headers=["#", "역할", "인원", "소속", "레벨", "단가(천원)", "M/M", "인건비(천원)", "역할 정의"],
    rows=[
        ["1", "PM", "1", "SI사", "L-4.5", "37,979", "3.2", "121,533", "프로젝트 총괄, 고객 커뮤니케이션"],
        ["2", "PL/분석설계", "1", "SI사", "L-3.0", "28,728", "6.0", "172,368", "요건 정의, 분석 및 설계"],
        # ...
    ]
)
```

### 고밀도 불릿 (15개+ 항목)

```python
prs.add_content(
    title="세부 기능 목록", section="기능",
    chapter="III. 기능 정의",
    density="dense",          # 9pt 불릿, 최소 간격
    bullets=[
        "설비 및 설비그룹 관리",
        "가상계측 모델 등록·검증·배포",
        "이벤트 기반 실시간 예측",
        "자동 재학습 스케줄 관리 및 트리거",
        # ... 15개 이상 항목
    ]
)
```

### 다중 컬럼 레이아웃 (비율 지정)

```python
# 70/30 2-column layout
prs.add_multi_column(
    title="시스템 구성", section="아키텍처", chapter="II. 아키텍처",
    ratio=[7, 3],
    columns=[
        {"title": "메인 시스템", "bullets": ["MLOps 파이프라인", "실시간 예측 엔진", "모델 저장소"]},
        {"title": "보조 시스템", "bullets": ["모니터링", "로깅"]},
    ]
)

# 4-column equal layout
prs.add_multi_column(
    title="4대 핵심 전략", section="전략", chapter="I. 전략",
    columns=[
        {"title": "데이터", "body": "데이터 품질 관리 체계 수립"},
        {"title": "플랫폼", "body": "클라우드 네이티브 전환"},
        {"title": "분석", "body": "AI/ML 모델 고도화"},
        {"title": "조직", "body": "DX 전문인력 양성"},
    ]
)
```

### 카드 그리드

```python
prs.add_card_grid(
    title="핵심 역량", section="역량", chapter="III. 역량",
    cols=3, show_icon=True,
    cards=[
        {"title": "AI/ML", "icon_text": "AI", "items": ["예측 모델", "이상 탐지", "자동 분류"]},
        {"title": "Data", "icon_text": "DB", "items": ["실시간 수집", "ETL 파이프라인"], "footer": "24/7 운영"},
        {"title": "Cloud", "icon_text": "CL", "items": ["K8s 오케스트레이션", "CI/CD 자동화"]},
    ]
)
```

### 콘텐츠+사이드바

```python
prs.add_content_sidebar(
    title="프로젝트 현황", section="현황", chapter="I. 현황",
    main_bullets=["Sprint 12 완료", "모델 정확도 95.2% 달성", "API 응답시간 < 200ms"],
    sidebar_items=[
        {"title": "일정", "body": "2025.Q3 완료 예정"},
        {"title": "리스크", "bullets": ["데이터 품질 이슈", "인력 부족"]},
        {"title": "예산", "body": "집행률 72%"},
    ],
    sidebar_position="right", sidebar_ratio=0.48
)
```

### 타이틀바+컬럼

```python
prs.add_titled_columns(
    title="추진 전략", section="전략", chapter="II. 전략",
    bar_title="디지털 트랜스포메이션 3대 추진 방향",
    columns=[
        {"title": "자동화", "icon_text": "AT", "bullets": ["RPA 도입", "워크플로우 자동화"]},
        {"title": "지능화", "icon_text": "AI", "bullets": ["예측 분석", "자연어 처리"]},
        {"title": "최적화", "icon_text": "OP", "bullets": ["비용 절감", "프로세스 개선"]},
    ]
)
```

### 아이콘 프로세스 흐름

```python
prs.add_icon_process(
    title="데이터 파이프라인", section="파이프라인", chapter="IV. 파이프라인",
    steps=[
        {"title": "수집", "icon_text": "1", "description": "센서 데이터 실시간 수집"},
        {"title": "전처리", "icon_text": "2", "description": "클렌징 및 피처 추출"},
        {"title": "학습", "icon_text": "3", "description": "모델 학습 및 검증"},
        {"title": "배포", "icon_text": "4", "description": "실시간 추론 서비스 배포"},
    ],
    bottom_sections=[
        {"title": "입력 데이터", "items": [{"label": "센서", "value": "1,200개"}, {"label": "주기", "value": "1초"}]},
        {"title": "출력 결과", "items": [{"label": "예측값", "value": "두께/품질"}, {"label": "정확도", "value": "95%+"}]},
    ]
)
```

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
                [{"text": "항목 1", "tag": "계열사A", "tag_color": "#1565C0"}],
                # Year 2 items
                [{"text": "항목 2", "color": "green"}],
                # Year 3 items
                [{"text": "항목 3", "color": "orange"}]
            ]
        }
    ]
}
```

### 간트 차트 구성

```python
prs.add_gantt_chart(
    title="프로젝트 일정표", section="일정",
    subtitle="2025년 상반기",
    start_date="2025-01",
    months=6,
    tasks=[
        {"name": "요구사항 분석", "owner": "김팀장", "start": 0, "duration": 2, "progress": 100, "color": "#A50034"},
        {"name": "설계", "owner": "이과장", "start": 1, "duration": 3, "progress": 60},
        {"name": "개발", "owner": "박대리", "start": 3, "duration": 4, "progress": 0},
        {"name": "Go-Live", "owner": "", "start": 5, "duration": 0, "milestone": True},
    ]
)
```

### 조직도 구성

```python
prs.add_org_chart(
    title="조직 구조", section="조직",
    org_data={
        "name": "CEO", "title": "대표이사",
        "children": [
            {"name": "CTO", "title": "기술총괄", "children": [
                {"name": "개발팀", "title": "팀장 김OO"},
                {"name": "인프라팀", "title": "팀장 이OO"},
            ]},
            {"name": "CFO", "title": "재무총괄"},
        ]
    }
)
```

### 피라미드 다이어그램

```python
prs.add_pyramid(
    title="전략 계층 구조", section="전략",
    levels=[
        {"label": "비전", "description": "글로벌 No.1", "color": "#A50034"},
        {"label": "전략", "description": "디지털 전환 가속화"},
        {"label": "실행 과제", "description": "AI·데이터·클라우드 역량 강화"},
        {"label": "기반", "description": "인재·문화·인프라"},
    ]
)
```

### 포지셔닝 맵

```python
prs.add_positioning_map(
    title="경쟁사 포지셔닝", section="분석",
    x_label="가격 경쟁력", y_label="기술력",
    quadrant_labels=["고가·고기술", "저가·고기술", "고가·저기술", "저가·저기술"],
    items=[
        {"label": "자사", "x": 0.7, "y": 0.8, "size": 1.5, "color": "#A50034"},
        {"label": "경쟁사A", "x": 0.3, "y": 0.6, "size": 1.0},
        {"label": "경쟁사B", "x": 0.5, "y": 0.4, "size": 1.2, "color": "#2E7D32"},
    ]
)
```

### 키워드 강조

```python
prs.add_keyword_highlight(
    title="핵심 키워드", section="요약",
    description="2025년 전략 방향의 핵심 키워드입니다.",
    keywords=[
        {"text": "디지털 전환", "size": 36, "color": "#A50034", "bold": True},
        {"text": "AI", "size": 32, "color": "#1565C0", "bold": True},
        {"text": "클라우드", "size": 28},
        {"text": "자동화", "size": 24, "color": "#2E7D32"},
        {"text": "데이터", "size": 26, "bold": True},
    ]
)
```

### matplotlib 차트 삽입

```python
import matplotlib.pyplot as plt

# 방법 1: Figure 객체 직접 전달
fig, ax = plt.subplots(figsize=(8, 4))
ax.bar(["Q1", "Q2", "Q3", "Q4"], [120, 145, 160, 180])
ax.set_title("분기별 매출")
prs.add_matplotlib_chart(title="매출 추이", section="실적", fig=fig, caption="단위: 억원")
plt.close(fig)

# 방법 2: 이미지 파일로 저장 후 삽입
fig.savefig("/tmp/chart.png", dpi=150, bbox_inches="tight")
prs.add_chart_slide(title="매출 차트", section="실적", chart_path="/tmp/chart.png", caption="2025년 실적")
```

### 이미지 배치

```python
prs.add_image_slide(
    title="현장 사진", section="현황",
    images=[
        {"path": "/tmp/photo1.jpg", "caption": "공장 전경"},
        {"path": "/tmp/photo2.jpg", "caption": "생산라인"},
    ]
)
```

### 스윔레인 프로세스

```python
prs.add_swimlane(
    title="도입 업무 프로세스", section="프로세스",
    lanes=["고객사", "PM", "개발팀", "QA팀"],
    steps=[
        {"lane": 0, "col": 0, "text": "요구사항 전달", "color": "#A50034"},
        {"lane": 1, "col": 1, "text": "분석/설계"},
        {"lane": 2, "col": 2, "text": "개발", "color": "#1565C0"},
        {"lane": 3, "col": 3, "text": "테스트", "color": "#2E7D32"},
        {"lane": 0, "col": 4, "text": "최종 승인", "color": "#A50034"},
    ],
    connections=[(0,1), (1,2), (2,3), (3,4)],
)
```

### 제언 (클로징 전 권고사항)

```python
# 방법 1: 간단한 문자열 리스트
prs.add_recommendation(
    title="제언", section="제언",
    recommendations=["데이터 품질 확보 최우선", "단계적 도입 전략 수립"]
)

# 방법 2: 제목 + 상세 설명
prs.add_recommendation(
    title="제언", section="제언",
    recommendations=[
        {"title": "데이터 품질 확보", "detail": "입력 데이터 정합성 검증 체계를 초기부터 확립"},
        {"title": "단계적 도입", "detail": "파일럿 → 확대 적용 → 전사 확산 순서로 진행"},
    ]
)
```
