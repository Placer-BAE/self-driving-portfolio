"""
기초데이터 및 21개 투자안 후보군 상세 장표 생성 스크립트
- 16:9 와이드스크린 (13.333" x 7.5")
- 맥킨지/BCG 스타일 엄격한 개조식 명사형 종결
- 슬라이드 1: [상세 1] 기초데이터 체계 및 18개 대표 ETF 자산 유니버스
- 슬라이드 2: [상세 2] 논문 Exhibit 5 기반 4대 계열 21개 투자안 후보군 상세
- 슬라이드 3: [1장 통합 요약] 기초데이터 & 21개 후보군 원페이지 총괄 맵
"""

import os
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

OUT_FILE = r"C:\Users\bgh68\.gemini\antigravity\scratch\self-driving-portfolio\기초데이터_및_21개_투자안_후보군_상세.pptx"

NAVY = RGBColor(5, 28, 44)
BLUE = RGBColor(0, 82, 204)
BLUE_BG = RGBColor(239, 246, 255)
GRAY_BG = RGBColor(246, 248, 250)
LINE = RGBColor(214, 220, 228)
DARK = RGBColor(15, 23, 42)
BODY = RGBColor(71, 85, 105)
MUTED = RGBColor(148, 163, 184)
WHITE = RGBColor(255, 255, 255)
GREEN = RGBColor(5, 150, 105)
AMBER = RGBColor(217, 119, 6)
RED = RGBColor(185, 28, 28)
FONT = "Malgun Gothic"

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)
BLANK = prs.slide_layouts[6]

def box(s, x, y, w, h, fill=None, line=None, shape=MSO_SHAPE.RECTANGLE, lw=1):
    sh = s.shapes.add_shape(shape, Inches(x), Inches(y), Inches(w), Inches(h))
    if fill is None:
        sh.fill.background()
    else:
        sh.fill.solid()
        sh.fill.fore_color.rgb = fill
    if line is None:
        sh.line.fill.background()
    else:
        sh.line.color.rgb = line
        sh.line.width = Pt(lw)
    sh.shadow.inherit = False
    return sh

def text(s, x, y, w, h, t, size=14, bold=False, color=DARK, align=PP_ALIGN.LEFT,
         anchor=MSO_ANCHOR.TOP, spacing=4):
    tb = s.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    lines = t if isinstance(t, list) else [t]
    for i, ln in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        p.space_after = Pt(spacing)
        r = p.add_run()
        r.text = ln
        r.font.name = FONT
        r.font.size = Pt(size)
        r.font.bold = bold
        r.font.color.rgb = color
    return tb

def header(s, section, title, page):
    box(s, 0, 0, 13.333, 0.08, fill=NAVY)
    text(s, 0.6, 0.35, 10, 0.3, section, 11, True, BLUE)
    text(s, 0.6, 0.65, 12, 0.6, title, 24, True, NAVY)
    box(s, 0.6, 1.35, 12.13, 0.015, fill=LINE)
    text(s, 12.2, 7.0, 0.6, 0.3, str(page), 10, color=MUTED, align=PP_ALIGN.RIGHT)

def card(s, x, y, w, h, title, items, accent=BLUE, fill=GRAY_BG, title_size=16, item_size=12.5, spacing=6):
    box(s, x, y, w, h, fill=fill)
    box(s, x, y, w, 0.06, fill=accent)
    text(s, x + 0.2, y + 0.2, w - 0.4, 0.4, title, title_size, True, NAVY)
    text(s, x + 0.2, y + 0.72, w - 0.4, h - 0.8, ["·  " + i for i in items], item_size, color=BODY, spacing=spacing)

# ================================================================
# 슬라이드 1: 기초데이터 체계 및 18개 대표 ETF 유니버스 (Exhibit 11)
# ================================================================
s1 = prs.slides.add_slide(BLANK)
header(s1, "SYSTEM ARCHITECTURE | DATA LAYER", "기초데이터 체계 및 18개 대표 ETF 유니버스 매핑 (Exhibit 11)", 1)

# 좌측 카드: 데이터 수집 규격 & 전처리
card(s1, 0.6, 1.6, 3.8, 4.15, "1. 데이터 원천 및 수집 규격", [
    "데이터 소스: 야후 파이낸스(Yahoo Finance) 공식 실데이터",
    "분석 기간: 2013년 7월 ~ 2026년 9월 (총 159개월 장기 시계열)",
    "데이터 유형: 배당 재투자 반영 월간 수정 총수익률(Adj Close)",
    "거시 지표: FRED 실질 GDP(성장), CPI(물가), 10Y-2Y 금리차 연동",
    "무위험 수익률: 미국 1-3개월 T-Bill(BIL) 기반 3.7% 설정",
    "상용 유료 벤더 배제: 누구나 100% 동일하게 검증/재현 가능"
], accent=BLUE, title_size=16, item_size=12, spacing=7)

# 중앙 카드: 18개 대표 ETF 자산군 매핑
card(s1, 4.65, 1.6, 4.6, 4.15, "2. 18개 ETF 자산 유니버스 (Exhibit 11)", [
    "주식군 (6개 ETF / 4대 권역):",
    "  - 미국 대형주(SPY), 성장주(VUG), 가치주(VTV), 중소형주(IWM)",
    "  - 선진국 EAFE 주식(EFA), 신흥국 주식(EEM)",
    "채권군 (8개 ETF / 만기 및 신용 분산):",
    "  - 미국 단기국채(SHY), 중기국채(IEF), 장기국채(TLT)",
    "  - 투자등급 회사채(LQD), 미국 하이일드 채권(HYG)",
    "  - 국제 국채(BNDX), 국제 회사채(PICB), 신흥국 달러채(EMB)",
    "대체/실물 (3개): 금(GLD), 원자재(GSG), 미국 리츠(VNQ)",
    "현금성 자산 (1개): 초단기 T-Bill(BIL, 유동성 버퍼)"
], accent=NAVY, title_size=16, item_size=11.5, spacing=4)

# 우측 카드: 수리 엔진 및 공분산 정규화
card(s1, 9.5, 1.6, 3.23, 4.15, "3. 수리 통계 및 정규화", [
    "공분산 추정 엔진:",
    "  - Ledoit-Wolf OAS 수축 기법",
    "  - 표본 공분산의 노이즈 원천 제거",
    "  - 행렬 역전(Inversion) 수치 안정성 확보",
    "기대수익률(CMA) 모델링:",
    "  - 10년 자본시장 가정(CMA) 기반",
    "  - 거시 국면(Late-cycle) 충격 반영",
    "벤치마크(BM) 설정:",
    "  - 글로벌 전통 60/40 포트폴리오",
    "  - SPY 60% + IEF 40% 기준"
], accent=BLUE, title_size=16, item_size=11.5, spacing=5)

# 하단 요약 배너
box(s1, 0.6, 5.95, 12.13, 0.85, fill=NAVY)
text(s1, 0.9, 5.95, 11.5, 0.85, [
    "·  실무 정합성: 유동성이 풍부한 18개 대표 ETF로 글로벌 기관 포트폴리오의 실질 익스포저를 100% 복제",
    "·  무결성 원칙: 결측치 보정 및 이상치 정제 완료, Python 과학 연산(NumPy/SciPy)으로 수치 환각 0% 보장"
], 12, color=WHITE, spacing=4, anchor=MSO_ANCHOR.MIDDLE)


# ================================================================
# 슬라이드 2: 4대 계열 21개 투자안 후보군 상세 체계 (Exhibit 5)
# ================================================================
s2 = prs.slides.add_slide(BLANK)
header(s2, "PORTFOLIO CONSTRUCTION | 21 CANDIDATE MODELS", "4대 계열 21개 포트폴리오 최적화 후보군 상세 (Exhibit 5)", 2)

families = [
    ("Family A : 휴리스틱 (5개)", "매개변수 추정 오차 배제\n단순 규칙 기반 배분", [
        "1. 동일가중 (1/N Equal Weight)",
        "2. 시가총액 가중 (Market-Cap Weight)",
        "3. 역변동성 (Inverse Volatility, 1/σ)",
        "4. 역분산 (Inverse Variance, 1/σ²)",
        "5. 타깃 변동성 (Volatility Targeting)"
    ], BLUE),
    ("Family B : 수익률 최적화 (5개)", "기대수익률(μ) 극대화\n위험-수익 트레이드오프", [
        "1. 평균-분산 최적화 (MVO / Markowitz)",
        "2. 블랙-리터만 (Black-Litterman)",
        "3. 최대 샤프 (Max Sharpe Ratio)",
        "4. 로버스트 평균-분산 (Robust MVO)",
        "5. 리샘플링 프론티어 (Resampled)"
    ], NAVY),
    ("Family C : 위험 구조화 (5개)", "μ 추정 오류 원천 배제\n공분산(Σ) 기반 리스크 균등화", [
        "1. 위험균형 (Risk Parity / ERC)",
        "2. 최대 분산 (Max Diversification)",
        "3. 최소 분산 (Minimum Variance)",
        "4. 계층적 위험균형 (HRP / Tree ML)",
        "5. 꼬리위험 균형 (Tail Risk Parity)"
    ], BLUE),
    ("Family D : 비전통 모델 (6개)", "비선형 리스크 제어 및\n거버넌스 직교 헤지 탐색", [
        "1. Two-Phase Allocation (TPA)",
        "2. 조건부 가치위험 (CVaR 최적화)",
        "3. 소르티노 위험 (Downside Risk)",
        "4. 최대 엔트로피 (Max Entropy)",
        "5. 적대적 분산기 (Adversarial Div)",
        "6. 레짐 스위칭 (Regime-Switching)"
    ], NAVY)
]

for i, (title, subtitle, models, accent) in enumerate(families):
    x = 0.6 + i * 3.08
    box(s2, x, 1.6, 2.88, 4.15, fill=GRAY_BG)
    box(s2, x, 1.6, 2.88, 0.06, fill=accent)
    text(s2, x + 0.18, 1.78, 2.52, 0.35, title, 14.5, True, NAVY)
    text(s2, x + 0.18, 2.18, 2.52, 0.55, subtitle.split("\n"), 11, False, BLUE, spacing=2)
    box(s2, x + 0.18, 2.8, 2.52, 0.015, fill=LINE)
    text(s2, x + 0.18, 2.95, 2.52, 2.65, models, 12, False, DARK, spacing=9)

# 하단 거버넌스 연결 배너
box(s2, 0.6, 5.95, 12.13, 0.85, fill=BLUE_BG, line=BLUE)
text(s2, 0.9, 5.95, 11.5, 0.85, [
    "·  후보군 다양성: 전통적 금융공학(Family A, B)과 최신 머신러닝/위험제어(Family C, D) 21개 모델이 동일 조건에서 공정 경쟁",
    "·  거버넌스 연계: 21개 모델의 제안은 CRO 독립 리스크 심사(45%)와 심사역 간 피어 리뷰 수정 보르다 투표(55%)로 종합 평가됨"
], 12, color=DARK, spacing=4, anchor=MSO_ANCHOR.MIDDLE)


# ================================================================
# 슬라이드 3: [1장 통합 요약] 기초데이터 & 21개 투자안 총괄 맵 (One-Page Executive Summary)
# ================================================================
s3 = prs.slides.add_slide(BLANK)
header(s3, "EXECUTIVE SUMMARY | DATA & MODELS", "기초데이터 체계 및 21개 투자안 후보군 총괄 맵 (One-Page Summary)", 3)

# 좌측 영역 (5.9" 폭): 기초데이터 및 18개 자산군
box(s3, 0.6, 1.6, 5.85, 4.2, fill=GRAY_BG)
box(s3, 0.6, 1.6, 5.85, 0.06, fill=BLUE)
text(s3, 0.85, 1.8, 5.35, 0.4, "Part 1. 기초데이터 체계 및 18개 ETF 유니버스", 16, True, NAVY)

text(s3, 0.85, 2.3, 5.35, 0.35, "■ 데이터 원천 및 규격", 12.5, True, BLUE)
text(s3, 0.85, 2.65, 5.35, 0.8, [
    "· 야후 파이낸스 실데이터 159개월 (2013.07 ~ 2026.09, 배당재투자 수정 총수익률)",
    "· 거시 지표: FRED 실질 GDP(성장), CPI(물가), 장단기 금리차 연동",
    "· 수리 엔진: Ledoit-Wolf OAS 공분산 수축 및 10년 CMA 기대수익률"
], 11.5, color=BODY, spacing=3)

text(s3, 0.85, 3.55, 5.35, 0.35, "■ 18개 대표 ETF 유동성 대용치 (Exhibit 11)", 12.5, True, BLUE)
text(s3, 0.85, 3.9, 5.35, 1.7, [
    "· 주식군(6): 미국 대형주(SPY), 성장(VUG), 가치(VTV), 중소형(IWM), 선진국(EFA), 신흥국(EEM)",
    "· 채권군(8): 미국 단기채(SHY), 중기채(IEF), 장기채(TLT), 회사채(LQD), 하이일드(HYG), 국제채(BNDX, PICB, EMB)",
    "· 실물자산(3): 금(GLD), 원자재(GSG), 미국 리츠(VNQ)",
    "· 현금성(1): 미국 초단기 T-Bill(BIL, 무위험 수익률 대용)"
], 11.5, color=DARK, spacing=4)

# 우측 영역 (6.0" 폭): 21개 모델 후보군 체계
box(s3, 6.75, 1.6, 5.98, 4.2, fill=GRAY_BG)
box(s3, 6.75, 1.6, 5.98, 0.06, fill=NAVY)
text(s3, 7.0, 1.8, 5.48, 0.4, "Part 2. 논문 Exhibit 5 기반 4대 계열 21개 모델 후보군", 16, True, NAVY)

families_compact = [
    ("Family A : 휴리스틱 (5개)", "동일가중(1/N), 시가총액 가중, 역변동성(1/σ), 역분산(1/σ²), 타깃 변동성"),
    ("Family B : 수익률 최적화 (5개)", "평균-분산(MVO), 블랙-리터만(BL), 최대 샤프, 로버스트 MVO, 리샘플링 프론티어"),
    ("Family C : 위험 구조화 (5개)", "위험균형(Risk Parity), 최대 분산, 최소 분산, 계층적 위험균형(HRP), 꼬리위험 균형"),
    ("Family D : 비전통 모델 (6개)", "Two-Phase Allocation(TPA), CVaR 최적화, 소르티노 위험, 최대 엔트로피, 적대적 분산기, 레짐 스위칭")
]

for idx, (f_name, f_desc) in enumerate(families_compact):
    y_pos = 2.3 + idx * 0.82
    box(s3, 7.0, y_pos, 5.48, 0.72, fill=WHITE, line=LINE)
    text(s3, 7.15, y_pos + 0.08, 5.18, 0.28, f_name, 12, True, NAVY)
    text(s3, 7.15, y_pos + 0.35, 5.18, 0.32, f_desc, 11, color=BODY)

# 하단 연결 배너
box(s3, 0.6, 5.95, 12.13, 0.85, fill=NAVY)
text(s3, 0.9, 5.95, 11.5, 0.85, [
    "·  거버넌스 평가 결론: CRO 정량 평가(45%) + 심사역 피어 리뷰 보르다 투표(55%)를 합산하여 종합 순위 산출",
    "·  IPS 규정 준수: 종합 점수가 우수해도 사전 변동성 [8~12%] 밴드를 충족한 모델(TPA, 시총가중)만 최종 앙상블로 진입"
], 12, color=WHITE, spacing=4, anchor=MSO_ANCHOR.MIDDLE)

prs.save(OUT_FILE)
print("Saved Dedicated Slides PPTX successfully to:", OUT_FILE)
