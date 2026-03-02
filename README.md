# lg-pptx

LG 그룹 디자인 컨벤션을 따르는 PPTX 자동 생성 스킬 (에이투지체 폰트)

## Features

- LG RED 액센트 블록, L-브래킷 장식, 섹션 인디케이터 자동 적용
- 20가지 슬라이드 템플릿: Cover, TOC, Content, Roadmap, Table, KPI Cards, Timeline, Process Flow, SWOT, Architecture, Strategy Pillars, Financial Summary 등
- 에이투지체 폰트 + 한국어/영어 완벽 지원
- `python-pptx` 기반, 16:9 와이드스크린

## Install

```bash
npx skills add <owner>/lg-pptx-skill
```

## Requirements

```bash
pip install python-pptx Pillow
```

### Font

- **에이투지체 (A2Z)**: https://freesentation.blog/a2z
- 라이선스: SIL OFL (무료)
- 폴백: 맑은 고딕 (Malgun Gothic)

## Quick Start

```python
from lg_pptx_builder import LGPresentation

prs = LGPresentation()
prs.add_cover("프로젝트 제목", subtitle="팀명", date="2025.03")
prs.add_content("시스템 개요", section="Summary", bullets=["내용1", "내용2"])
prs.save("output.pptx")
```

## Color Palette

| Color | HEX | Usage |
|-------|-----|-------|
| LG RED | `#A50034` | Brand accent |
| Black | `#000000` | Title text |
| Dark Gray | `#333333` | Subtitle |
| Charcoal | `#3C3C3C` | Table headers |
| Light Gray | `#F2F2F2` | Card backgrounds |

## Slide Types

Cover, TOC, Section Divider, Content, Roadmap, Table, Summary Matrix, Two Column, KPI Cards, Timeline, Process Flow, SWOT, Architecture, Strategy Pillars, Risk Matrix, Financial Summary, Milestone Tracker, Comparison Cards

## License

MIT
