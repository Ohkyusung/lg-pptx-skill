# LG Group Design Tokens Reference

## Color Palette

### Primary Colors

| Token | Name | HEX | RGB | 용도 |
|-------|------|-----|-----|------|
| `LG_RED` | LG Red | `#A50034` | `(165, 0, 52)` | 브랜드 컬러. L-브래킷, 액센트 바, TOC 번호, 카테고리 라벨, 강조 |
| `BLACK` | Black | `#000000` | `(0, 0, 0)` | 제목, 본문 텍스트 |
| `WHITE` | White | `#FFFFFF` | `(255, 255, 255)` | 슬라이드 배경 |

### Secondary Colors (텍스트)

| Token | Name | HEX | RGB | 용도 |
|-------|------|-----|-----|------|
| `DARK_GRAY` | Dark Gray | `#333333` | `(51, 51, 51)` | 부제목, 보조 텍스트 |
| `MEDIUM_GRAY` | Medium Gray | `#666666` | `(102, 102, 102)` | 섹션명, 캡션, 비활성 텍스트 |

### Surface Colors (배경)

| Token | Name | HEX | RGB | 용도 |
|-------|------|-----|-----|------|
| `LIGHT_GRAY` | Light Gray | `#F2F2F2` | `(242, 242, 242)` | 콘텐츠 박스 배경, 테이블 교대행 |
| `CHARCOAL` | Charcoal | `#3C3C3C` | `(60, 60, 60)` | 타임라인 헤더, 테이블 헤더 |
| `BORDER_GRAY` | Border Gray | `#CCCCCC` | `(204, 204, 204)` | 구분선, 테두리 |

### Accent Colors (강조)

| Token | Name | HEX | RGB | 용도 |
|-------|------|-----|-----|------|
| `GREEN` | Accent Green | `#2E7D32` | `(46, 125, 50)` | 미래 계획 항목, 신규 기능 |
| `ORANGE` | Accent Orange | `#D4760A` | `(212, 118, 10)` | 하이라이트 항목, 주의 사항 |

### Affiliate Tag Colors (계열사 뱃지)

| 계열사 | HEX | 용도 |
|--------|-----|------|
| 계열사A | `#1565C0` | 파란색 텍스트 태그 |
| 계열사B | `#D4760A` | 주황색 텍스트 태그 |
| 계열사C | `#2E7D32` | 초록색 텍스트 태그 |

## Typography

### Font Family

| 우선순위 | 폰트명 | 비고 |
|----------|--------|------|
| 1 | 에이투지체 (A2z) | 기본 폰트 (한글+영문) |
| 2 | 맑은 고딕 (Malgun Gothic) | 폴백 폰트 |
| 3 | Arial | 최종 폴백 (영문) |

### Type Scale

| 용도 | Size | Weight | Line Height | 색상 |
|------|------|--------|-------------|------|
| 표지 제목 | 32pt | Bold | 1.3 | Black |
| 표지 부제 | 14pt | Regular | 1.4 | Dark Gray |
| TOC 타이틀 | 28pt | Regular | 1.3 | Black |
| TOC 항목 | 16pt | Bold | 1.3 | LG Red |
| TOC 하위항목 | 13pt | Regular | 1.3 | Dark Gray |
| 섹션 제목 | 24pt | Bold | 1.2 | Black |
| 본문 제목 | 18pt | Bold | 1.3 | Black |
| 본문 | 12pt | Regular | 1.5 | Black |
| 본문 (소) | 11pt | Regular | 1.5 | Black/Dark Gray |
| 표 헤더 | 10pt | Bold | 1.2 | White (on Charcoal) |
| 표 본문 | 10pt | Regular | 1.3 | Black |
| 캡션/주석 | 9pt | Regular | 1.3 | Medium Gray |
| 라벨 뱃지 | 9pt | Bold | 1.2 | White (on LG Red) |

### Korean Font Setup (python-pptx)

python-pptx에서 한글 폰트를 올바르게 적용하려면 Latin과 East Asian 폰트를 모두 설정해야 합니다:

```python
# font.name은 <a:latin> 요소만 설정 (영문에만 적용)
# 한글에는 <a:ea> 요소를 직접 XML로 설정해야 함

from pptx.oxml.ns import qn

def set_korean_font(run, font_name):
    rPr = run._r.get_or_add_rPr()
    # East Asian font
    for ea in rPr.findall(qn('a:ea')):
        rPr.remove(ea)
    ea = OxmlElement('a:ea')
    ea.set('typeface', font_name)
    rPr.append(ea)
    # Complex Script font
    for cs in rPr.findall(qn('a:cs')):
        rPr.remove(cs)
    cs = OxmlElement('a:cs')
    cs.set('typeface', font_name)
    rPr.append(cs)
```

---

## Hanwha Theme Design Tokens

### Color Palette (한화 Brand System Design Guide 기반)

#### Primary Colors

| Token | Name | HEX | RGB | 용도 |
|-------|------|-----|-----|------|
| `RED` (= ORANGE) | Hanwha Orange | `#F37321` | `(243, 115, 33)` | 브랜드 컬러. 액센트 바, 챕터명, 하단 바, 강조 |
| `BLACK` | Navy Black | `#1A1A2E` | `(26, 26, 46)` | 제목, 본문 텍스트 |
| `WHITE` | White | `#FFFFFF` | `(255, 255, 255)` | 슬라이드 배경 |

#### Secondary Colors

| Token | Name | HEX | RGB | 용도 |
|-------|------|-----|-----|------|
| `DARK_RED` (= LIGHT_NAVY) | Light Navy | `#353968` | `(53, 57, 104)` | 뱃지, 라벨 배경 |
| `DARK_GRAY` | Dark Gray | `#333333` | `(51, 51, 51)` | 부제목, 보조 텍스트 |
| `MEDIUM_GRAY` | Medium Gray | `#666666` | `(102, 102, 102)` | 섹션명, 캡션 |

#### Surface Colors

| Token | Name | HEX | RGB | 용도 |
|-------|------|-----|-----|------|
| `LIGHT_GRAY` | Warm Grey | `#EFEEE8` | `(239, 238, 232)` | 콘텐츠 박스 배경 |
| `CHARCOAL` (= NAVY) | Navy | `#1D1E37` | `(29, 30, 55)` | 테이블 헤더, 다크 배경 |
| `BORDER_GRAY` (= SAND) | Sand | `#C7BB9F` | `(199, 187, 159)` | 구분선, 테두리 |

#### Accent Colors

| Token | Name | HEX | RGB | 용도 |
|-------|------|-----|-----|------|
| `ORANGE_70` | Hanwha Orange 70% | `#F89B6C` | `(248, 155, 108)` | 보조 강조 |
| `ORANGE_50` | Hanwha Orange 50% | `#FBB584` | `(251, 181, 132)` | 연한 강조 |
| `GREEN` | Green | `#2E7D32` | `(46, 125, 50)` | 미래 계획 항목 |
| `BLUE` (= TURQUOISE) | Turquoise | `#5C7687` | `(92, 118, 135)` | 보조 강조 |

#### Background Colors (참고)

| Name | HEX | RGB | 용도 |
|------|-----|-----|------|
| Sand | `#C7BB9F` | `(199, 187, 159)` | 따뜻한 배경 |
| Turquoise | `#5C7687` | `(92, 118, 135)` | 차분한 배경 |
| Navy | `#1D1E37` | `(29, 30, 55)` | 다크 배경 |

### Hanwha Typography

#### Font Family

| 우선순위 | 폰트명 | 비고 |
|----------|--------|------|
| 1 (제목) | 한화 B_OTF | Bold — 표지, 섹션 제목 |
| 2 (본문) | 한화고딕 R_OTF | Regular — 일반 텍스트 |
| 3 (강조) | 한화고딕 B_OTF | Bold — 테이블 헤더, 강조 |
| 4 (보조) | 한화고딕 L_OTF | Light — 캡션, 보조 텍스트 |
| 5 (폴백) | 맑은 고딕 | 폴백 |

### Hanwha Chrome (LG L-Style과의 차이)

| 요소 | LG | 한화 |
|------|-----|------|
| 표지 장식 | L-브래킷 (RED) | 하단 오렌지 바 |
| 액센트 바 두께 | 0.40cm | 0.30cm (더 얇음) |
| 구분선 높이 | 0.05cm | 0.03cm (더 얇음) |
| 헤더 바 색상 | Charcoal (#3C3C3C) | Navy (#1D1E37) |

---

## Layout Dimensions

### Slide Size
- **비율**: 16:9 와이드스크린
- **크기**: 13.333" x 7.5" (33.867cm x 19.05cm)

### Margins & Spacing

| 영역 | 값 | 비고 |
|------|------|------|
| 좌측 마진 (표지) | 0.8cm | L-브래킷 시작점 |
| 좌측 마진 (내용) | 1.5cm | 액센트 바 뒤 |
| 우측 마진 | 1.0cm | |
| 상단 마진 | 1.2cm | |
| 하단 마진 | 1.0cm | |
| 콘텐츠 시작점 | 2.0cm (좌) | 액센트 바 + 여백 |
| 제목-본문 간격 | 1.5cm | |
| 불릿 항목 간격 | 6pt (space_after) | |
| 표 셀 내부 패딩 | 0.2cm (좌우), 0.1cm (상하) | |

### Component Dimensions

| 컴포넌트 | 폭 | 높이 | 비고 |
|----------|------|------|------|
| 액센트 바 (좌) | 0.3cm | 슬라이드 전체 | x=0 |
| L-브래킷 팔 길이 | 2.5cm | - | |
| L-브래킷 두께 | 0.4cm | - | |
| 섹션 인디케이터 점 | 0.35cm | 0.35cm | 원형 |
| 타임라인 헤더 높이 | 1.2cm | - | |
| 라벨 뱃지 | 4.0cm | 0.8cm | 가변 폭 |
| 로드맵 좌측 라벨 | 2.5cm | 가변 | |
| 테이블 행 높이 | ~1.0cm | - | |

## Slide Type Specifications

### Cover (표지)
```
┌──────────────────────────────┐
│ ┌──┐                        │
│ │  │                        │
│ │                            │
│                              │
│      [제목 - 32pt Bold]      │
│                              │
│                              │
│      [부제 - 14pt]           │
│      [날짜 - 14pt Bold]   ┌─┤
│                        LG │ │
│                            └─┤
└──────────────────────────────┘
```

### TOC (목차)
```
┌──────────────────────────────┐
│  ─────────────────────────── │ ← 상단 회색선
│                              │
│  Contents                    │
│  ───                         │ ← 빨간 밑줄 바
│  ─────────────────────────── │ ← 회색 구분선
│                              │
│    I.  Summary               │ ← 빨간색
│    II. 시스템 소개           │ ← 빨간색
│        - 항목 1              │ ← 검정 들여쓰기
│        - 항목 2              │
│    III. 첨부자료             │
│                              │
└──────────────────────────────┘
```

### Content (내용)
```
┌──────────────────────────────┐
│▌ 1.1 제목               섹션 ●│ ← 좌: 레드바, 우: 섹션+점
│▌                              │
│▌ [본문 내용 영역]             │
│▌  • 불릿 항목 1               │
│▌  • 불릿 항목 2               │
│▌  • 불릿 항목 3               │
│▌                              │
│▌                              │
│▌                              │
└──────────────────────────────┘
```

### Roadmap (로드맵)
```
┌──────────────────────────────┐
│▌ [프로젝트] 로드맵       섹션 ●│
│▌ 설명 텍스트                  │
│▌                              │
│▌     ◄ 2025 ►◄ 2026 ►◄ 2027 ►│ ← 차콜 쉐브론
│▌ ┌──┐┌────┐┌────┐┌────┐      │
│▌ │ 라││    ││    ││    │      │
│▌ │ 벨││ 셀 ││ 셀 ││ 셀 │     │
│▌ └──┘└────┘└────┘└────┘      │
│▌                              │
│▌ ■ 계열사별 현황              │
│▌ ┌────┬────┬────┬────┐       │
│▌ │헤더│헤더│헤더│헤더│        │ ← 차콜 헤더
│▌ ├────┼────┼────┼────┤       │
│▌ │데이│터  │    │    │       │
│▌ └────┴────┴────┴────┘       │
└──────────────────────────────┘
```
