"""
중간발표용 PPT 생성 (웹페이지 실제 캡쳐본 4종 삽입 및 실화면 수치 100% 동기화)
- 16:9 와이드스크린 (13.333" x 7.5")
- 맥킨지/BCG 스타일 간결한 개조식 단어 위주
- 웹페이지 캡쳐 4종 최적 배치
"""
import os
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

OUT = r"C:\Users\bgh68\.gemini\antigravity\scratch\self-driving-portfolio\The_Self_Driving_Portfolio_중간발표_최종본.pptx"
OUT_LEGACY = r"C:\Users\bgh68\.gemini\antigravity\scratch\self-driving-portfolio\The_Self_Driving_Portfolio_중간발표_간결본.pptx"

IMG_DIR = r"C:\Users\bgh68\.gemini\antigravity\brain\362c0791-8de4-4bb2-af50-e049d89fa8db\.user_uploaded"
IMG1 = os.path.join(IMG_DIR, "media_1791889036964.png") # 메인 대시보드 + 스테퍼 + KPI
IMG2 = os.path.join(IMG_DIR, "media_1791889094889.png") # 대자산군 도넛 + 60/40 비교
IMG3 = os.path.join(IMG_DIR, "media_1791889106342.png") # 18개 자산군 비중 표
IMG4 = os.path.join(IMG_DIR, "media_1791889122698.png") # 전략 심의 보르다 투표

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

def card(s, x, y, w, h, title, items, accent=BLUE, fill=GRAY_BG, title_size=16, item_size=13):
    box(s, x, y, w, h, fill=fill)
    box(s, x, y, w, 0.06, fill=accent)
    text(s, x + 0.22, y + 0.25, w - 0.44, 0.45, title, title_size, True, NAVY)
    text(s, x + 0.22, y + 0.85, w - 0.44, h - 0.95, ["·  " + i for i in items], item_size, color=BODY, spacing=7)

def add_image_card(s, x, y, w, h, img_path):
    # 사진 감싸는 얇은 테두리 카드
    box(s, x - 0.04, y - 0.04, w + 0.08, h + 0.08, fill=WHITE, line=LINE, lw=1)
    s.shapes.add_picture(img_path, Inches(x), Inches(y), Inches(w), Inches(h))

# ================================================================ 1. 표지
s = prs.slides.add_slide(BLANK)
box(s, 0, 0, 13.333, 7.5, fill=NAVY)
box(s, 0, 0, 0.25, 7.5, fill=BLUE)
text(s, 1.0, 2.2, 11, 0.5, "중간발표", 16, True, RGBColor(147, 197, 253))
text(s, 1.0, 2.7, 11, 1.0, "The Self-Driving Portfolio", 44, True, WHITE)
text(s, 1.0, 3.7, 11, 0.6, "멀티 에이전트 기반 기관 전략적 자산배분(SAA) 시스템 구현", 22, False, RGBColor(226, 232, 240))
box(s, 1.0, 4.6, 1.2, 0.03, fill=BLUE)
text(s, 1.0, 4.85, 11, 0.4, "Base Paper  |  Andrew Ang, Nazym Azimbayev, Andrey Kim (2026)", 13, color=MUTED)
text(s, 1.0, 5.25, 11, 0.4, "Live Demo   |  self-driving-portfolio.onrender.com", 13, color=RGBColor(147, 197, 253))

# ================================================================ 2. 목차
s = prs.slides.add_slide(BLANK)
header(s, "AGENDA", "목차", 2)
agenda = [("01", "논문 배경", "기관 SAA의 구조적 한계 · 6단계 파이프라인 · 4대 설계 원칙"),
          ("02", "주제 선정 이유", "시의성 · 실무성 · 구현 가능성"),
          ("03", "웹 시스템 구현 및 실화면", "메인 대시보드 · 자산 배분 · 18개 ETF 포트폴리오 · 보르다 투표"),
          ("04", "중간 성과 및 향후 계획", "진행 현황 · 주요 이슈 · 기말까지 로드맵")]
for i, (n, t, d) in enumerate(agenda):
    y = 1.9 + i * 1.25
    text(s, 1.2, y, 1.2, 0.8, n, 36, True, BLUE)
    text(s, 2.6, y + 0.05, 5.5, 0.5, t, 22, True, NAVY)
    text(s, 2.6, y + 0.6, 8.5, 0.4, d, 14, color=BODY)
    if i < 3:
        box(s, 2.6, y + 1.1, 9.5, 0.01, fill=LINE)

# ================================================================ 3. 논문 배경: 한계
s = prs.slides.add_slide(BLANK)
header(s, "01  논문 배경", "기관 SAA의 구조적 한계", 3)
problems = [("연 1회", "경직된 주기", ["시장 국면 변화 대응 지연", "형식적 재검토"]),
            ("사일로", "부서 간 단절", ["부서별 전제 불일치", "수작업 취합 비용"]),
            ("사후 검증", "리스크 통제", ["확정 후 리스크 검토", "수정 비용 과다"]),
            ("기록 부재", "감사 추적", ["조정 사유 불투명", "재현 · 검증 곤란"])]
for i, (big, sub, items) in enumerate(problems):
    x = 0.6 + i * 3.08
    box(s, x, 1.8, 2.88, 3.4, fill=GRAY_BG)
    text(s, x + 0.25, 2.05, 2.4, 0.7, big, 28, True, RED)
    text(s, x + 0.25, 2.8, 2.4, 0.4, sub, 15, True, NAVY)
    text(s, x + 0.25, 3.4, 2.45, 1.7, ["·  " + it for it in items], 13, color=BODY, spacing=8)
box(s, 0.6, 5.55, 12.13, 1.05, fill=NAVY)
text(s, 0.95, 5.55, 11.5, 1.05, "논문 제안  →  약 50개 AI 에이전트가 SAA 전 과정을 분업 · 상호 검증 · 감사 기록",
     18, True, WHITE, anchor=MSO_ANCHOR.MIDDLE)

# ================================================================ 4. 6단계 파이프라인
s = prs.slides.add_slide(BLANK)
header(s, "01  논문 배경", "6단계 에이전트 파이프라인", 4)
stages = [("거시 국면", "Macro (1)", "후기 사이클 진단"),
          ("자산군 CMA", "CMA (18)", "18개 자산 10년 예측"),
          ("공분산 행렬", "Covariance (1)", "Ledoit-Wolf / OAS 수축"),
          ("포트폴리오 구성", "PC (21)", "21개 최적화 모델 경쟁"),
          ("전략 심의 · 투표", "CRO + 21 심사역", "수정 보르다 카운트 투표"),
          ("최종 앙상블", "CIO (1)", "역추적오차(TE) 결합")]
w = 2.0
for i, (t, en, d) in enumerate(stages):
    x = 0.6 + i * 2.03
    fill = BLUE if i == 4 else NAVY
    box(s, x, 2.0, w + 0.15, 1.3, fill=fill, shape=MSO_SHAPE.CHEVRON if i else MSO_SHAPE.PENTAGON)
    text(s, x + 0.35, 2.0, w - 0.45, 1.3, str(i + 1), 26, True, WHITE, PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE)
    text(s, x, 3.6, w, 0.4, t, 15, True, NAVY, PP_ALIGN.CENTER)
    text(s, x, 4.05, w, 0.35, en, 11.5, True, BLUE, PP_ALIGN.CENTER)
    text(s, x, 4.45, w, 0.4, d, 11.5, color=BODY, align=PP_ALIGN.CENTER)
box(s, 6.7, 5.25, 4.1, 0.55, fill=BLUE_BG, line=BLUE)
text(s, 6.7, 5.25, 4.1, 0.55, "↺  CRO IPS 위반 시 4단계 재최적화", 13, True, BLUE, PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE)
text(s, 0.6, 6.3, 12, 0.4, "입력: 18개 자산군 시계열 · 거시지표   |   출력: 최종 SAA 비중 · 이사회 메모", 13, color=BODY)

# ================================================================ 5. 4대 설계 원칙
s = prs.slides.add_slide(BLANK)
header(s, "01  논문 배경", "4대 설계 원칙", 5)
principles = [("전문화", "Specialization", "거시 · 자산군 · 최적화 · 리스크 에이전트 분업"),
              ("도구 결정론", "Tool Determinism", "LLM = 정성적 해석 · 토론\nPython = 수학 행렬 계산"),
              ("적대적 검증", "Adversarial Critique", "CRO 거부권(Veto) → 미달 시 배제"),
              ("감사 추적", "Audit Trail", "보르다 투표 · 소수의견 · 조정 사유 기록")]
for i, (k, en, d) in enumerate(principles):
    x = 0.6 + i * 3.08
    box(s, x, 1.9, 2.88, 3.9, fill=BLUE_BG if i == 1 else GRAY_BG)
    text(s, x + 0.25, 2.2, 0.8, 0.6, f"0{i+1}", 18, True, BLUE)
    text(s, x + 0.25, 2.85, 2.4, 0.6, k, 24, True, NAVY)
    text(s, x + 0.25, 3.5, 2.4, 0.4, en, 12, True, BLUE)
    box(s, x + 0.25, 4.0, 0.6, 0.02, fill=BLUE)
    text(s, x + 0.25, 4.2, 2.4, 1.4, d.split("\n"), 13.5, color=BODY, spacing=6)
text(s, 0.6, 6.3, 12, 0.4, "핵심  →  LLM 수치 환각 차단 (Tool Determinism) + 의사결정 투명성", 15, True, NAVY)

# ================================================================ 6. 주제 선정 이유
s = prs.slides.add_slide(BLANK)
header(s, "02  주제 선정 이유", "왜 이 논문인가", 6)
reasons = [("시의성", ["2026년 4월 최신 연구", "Agentic AI × 자산운용 접목", "기관 도입 초기 단계 탐색"]),
           ("실무성", ["연기금 · 국부펀드 실제 SAA 구조", "위원회 · IPS 규정 등 거버넌스 반영", "금융공학 이론과 실무의 결합"]),
           ("구현 가능성", ["명확한 수식 (OAS · BL · CVaR)", "18개 ETF 실제 데이터로 재현", "웹 시연으로 즉시 검증 가능"])]
for i, (t, items) in enumerate(reasons):
    x = 0.6 + i * 4.1
    card(s, x, 1.8, 3.9, 3.0, t, items, title_size=22, item_size=15)
box(s, 0.6, 5.3, 12.13, 0.9, fill=NAVY)
text(s, 0.95, 5.3, 11.5, 0.9, "목표  →  논문의 50개 에이전트 구조를 실제 작동하는 웹 시스템으로 재현",
     18, True, WHITE, anchor=MSO_ANCHOR.MIDDLE)

# ================================================================ 7. 시스템 구성도
s = prs.slides.add_slide(BLANK)
header(s, "03  웹 시스템 구현", "시스템 아키텍처 구성", 7)
blocks = [("데이터 계층", ["18개 ETF 월간 수익률", "거시지표 (성장·물가·금리)", "야후 파이낸스 실데이터"]),
          ("수리 엔진 (백엔드)", ["Python · FastAPI", "6단계 파이프라인 오케스트레이션", "NumPy · SciPy 결정론적 계산"]),
          ("대시보드 (프론트엔드)", ["React · TypeScript · Tailwind", "자산배분 · 투표 · 성과 진단 탭", "IPS 실시간 재실행"]),
          ("배포 및 인프라", ["Docker 멀티 스테이지 빌드", "Render 클라우드 24/7 서비스", "GitHub CI/CD 연동"])]
for i, (t, items) in enumerate(blocks):
    x = 0.6 + i * 3.14
    card(s, x, 1.9, 2.7, 3.0, t, items, accent=BLUE if i != 1 else NAVY, title_size=19, item_size=13.5)
    if i < 3:
        text(s, x + 2.72, 3.0, 0.4, 0.6, "→", 26, True, BLUE, PP_ALIGN.CENTER)
text(s, 0.6, 5.4, 12, 0.4, "핵심 구현 원칙", 14, True, NAVY)
text(s, 0.6, 5.8, 12, 0.8, ["·  수치 계산: 결정론적 코드 (NumPy/SciPy, 재현 가능, 환각 0%)",
                            "·  해설 및 토론: LLM 하이브리드 모드 (API 키 없이도 시뮬레이션 지원)"], 13, color=BODY)

# ================================================================ 8. [캡쳐 1] 메인 대시보드
s = prs.slides.add_slide(BLANK)
header(s, "03  웹 시스템 구현", "메인 대시보드 및 6단계 파이프라인 (실화면)", 8)
# 좌측 요약 카드
card(s, 0.6, 1.6, 4.2, 5.0, "화면 구성 및 핵심 지표", [
    "거버넌스 L3: 조건부 자율 SAA 운용",
    "IPS 적격 심사: 21개 후보 중 2개만 통과",
    "예상 명목수익률: 5.75% (실질 +3.3%)",
    "사전 연간변동성: 8.85% (밴드 8~12% 충족)",
    "샤프 지수: 0.231 (vs 60/40 0.219 대비 우수)",
    "예상 최대낙폭: -18.59% (하방 방어력 우수)",
    "트래킹 에러: 0.77% (한도 < 6% 준수)",
    "유효 자산수: 10.2개 (다변화 확보)"
], title_size=18, item_size=12.5)
# 우측 캡쳐 이미지 (가로 7.6" 높이 5.0")
add_image_card(s, 5.1, 1.6, 7.6, 5.0, IMG1)

# ================================================================ 9. [캡쳐 2] 대자산군 배분 및 벤치마크
s = prs.slides.add_slide(BLANK)
header(s, "03  웹 시스템 구현", "대자산군 배분 및 60/40 벤치마크 비교 (실화면)", 9)
card(s, 0.6, 1.6, 4.2, 5.0, "자산 배분 및 벤치마크 진단", [
    "주식군 (Equity): 56.49%",
    "채권군 (Fixed Income): 40.71%",
    "실물자산 (Real Assets): 1.65%",
    "현금 버퍼 (Cash): 1.17%",
    "스태그플레이션 방어: 주식 언더웨이트",
    "채권 듀레이션 및 현금 버퍼 확보",
    "최대 낙폭 개선: -34.3% → -18.59%",
    "사전 트래킹 에러: 0.77% (IPS 한도 준수)"
], title_size=18, item_size=12.5)
add_image_card(s, 5.1, 1.6, 7.6, 5.0, IMG2)

# ================================================================ 10. [캡쳐 3] 18개 자산군별 비중
s = prs.slides.add_slide(BLANK)
header(s, "03  웹 시스템 구현", "18개 자산군별 최종 권고 비중 (Exhibit 11)", 10)
card(s, 0.6, 1.6, 4.2, 5.0, "포트폴리오 비중 특징", [
    "미국 대형주 (SPY): 28.56% (위험 기여 45.2%)",
    "미국 중기국채 (IEF): 16.57% (핵심 안전판)",
    "미국 성장주/가치주 (VUG/VTV): 각 9.41%",
    "미국 중소형주 (IWM): 5.59%",
    "미국 단기/장기국채 (SHY/TLT): 각 5.09%",
    "투자등급 회사채 (LQD): 5.09%",
    "해외 국채 (BNDX): 3.56%  |  HYG: 2.79%",
    "선진국/신흥국 주식 (EFA/EEM): 각 1.76%"
], title_size=18, item_size=12.5)
add_image_card(s, 5.1, 1.6, 7.6, 5.0, IMG3)

# ================================================================ 11. [캡쳐 4] 전략 심의 보르다 투표
s = prs.slides.add_slide(BLANK)
header(s, "03  웹 시스템 구현", "전략 심의: 수정 보르다 투표 결과 (Exhibit 9)", 11)
card(s, 0.6, 1.6, 4.2, 5.0, "투표 결과 및 적격 심사", [
    "종합 점수: 동료 투표 55% + CRO 정량 45%",
    "1위 최대 분산 (변동성 2.7% → 8% 미달 탈락)",
    "2위 블랙-리터만 (변동성 6.9% → 미달 탈락)",
    "3위 위험균형 (변동성 6.5% → 미달 탈락)",
    "4위 TPA (변동성 8.3% → IPS 적격 통과)",
    "6위 시가총액 가중 (변동성 9.2% → 적격 통과)",
    "적대적 분산 에이전트: 반대 관점 제기",
    "결론: 적격 통과한 2개 모델만 앙상블 결합"
], title_size=18, item_size=12.5)
add_image_card(s, 5.1, 1.6, 7.6, 5.0, IMG4)

# ================================================================ 12. 현재 이슈 및 과제
s = prs.slides.add_slide(BLANK)
header(s, "04  향후 계획", "현재 이슈 및 검토 과제", 12)
issues = [("IPS 변동성 밴드 재검토", ["기본 변동성 밴드 [8~12%]", "21개 모델 중 적격 2개뿐", "→ 저변동성 우수 모델 배제 완화 검토"]),
          ("데이터 연속성 및 확장", ["현재: 18개 ETF 월간 수익률", "거시지표 주기 동기화 필요", "→ 실시간 FRED API 연동"]),
          ("LLM 실 API 추론 평가", ["현재: 시뮬레이션 + 템플릿 결합", "실제 LLM 호출 시 비용·지연 발생", "→ 핵심 에이전트 선별 호출"]),
          ("백테스트 강건성 분석", ["단일 기간(2013-2026) 검증", "거래비용·슬리피지 미반영", "→ 서브 구간별 안정성 검증"])]
for i, (t, items) in enumerate(issues):
    x = 0.6 + i * 3.08
    card(s, x, 1.8, 2.88, 4.3, t, items, accent=AMBER, title_size=17, item_size=13)

# ================================================================ 13. 기말 로드맵
s = prs.slides.add_slide(BLANK)
header(s, "04  향후 계획", "기말까지 로드맵", 13)
phases = [("Phase 1", "1~2주", "통합 및 안정화", ["팀원 PR 병합 완료", "IPS 밴드 파라미터 튜닝", "배포 도메인 단일화"]),
          ("Phase 2", "3~4주", "데이터 고도화", ["FRED 거시지표 API 연동", "18개 ETF 공분산 재추정", "자산 유니버스 확장 검토"]),
          ("Phase 3", "5~6주", "LLM 실연동 검증", ["거시 · CIO 실 API 연동", "에이전트 추론 정성 평가", "시뮬레이션 모드와 비교"]),
          ("Phase 4", "7~8주", "최종 검증 및 보고", ["거래비용 반영 백테스트", "60/40 · 동일가중 심층 비교", "기말 보고서 및 발표 준비"])]
box(s, 0.9, 2.35, 11.6, 0.04, fill=LINE)
for i, (ph, wk, t, items) in enumerate(phases):
    x = 0.6 + i * 3.08
    text(s, x + 0.25, 1.65, 2.6, 0.3, ph, 12, True, BLUE)
    text(s, x + 1.2, 1.6, 1.6, 0.4, wk, 16, True, NAVY)
    box(s, x + 0.25, 2.17, 0.4, 0.4, fill=BLUE if i == 0 else NAVY, shape=MSO_SHAPE.OVAL)
    box(s, x, 2.9, 2.88, 3.0, fill=BLUE_BG if i == 0 else GRAY_BG)
    text(s, x + 0.25, 3.15, 2.4, 0.5, t, 20, True, NAVY)
    text(s, x + 0.25, 3.9, 2.45, 2.0, ["·  " + it for it in items], 14, color=BODY, spacing=12)
text(s, 0.6, 6.55, 12, 0.4, "※ 주차는 중간발표 이후 기준 일정", 11, color=MUTED)

# ================================================================ 14. 기말 산출물 + Q&A
s = prs.slides.add_slide(BLANK)
header(s, "04  향후 계획", "기말 목표 산출물 및 결언", 14)
outs = [("완성형 자율 SAA 시스템", "실데이터 · 실 API 연동 및 고도화 완료"),
        ("24/7 글로벌 라이브 데모", "Render 클라우드 상시 운영 (onrender.com)"),
        ("심층 실증 분석 보고서", "시장 국면별 백테스트 및 벤치마크 대비 우수성 입증"),
        ("코드베이스 및 문서화", "GitHub 저장소 완비 및 상세 아키텍처 가이드")]
for i, (k, v) in enumerate(outs):
    x = 0.6 + (i % 2) * 6.15
    y = 1.8 + (i // 2) * 1.75
    box(s, x, y, 5.95, 1.5, fill=GRAY_BG)
    box(s, x, y, 0.08, 1.5, fill=BLUE)
    text(s, x + 0.4, y + 0.3, 5.3, 0.5, k, 20, True, NAVY)
    text(s, x + 0.4, y + 0.85, 5.3, 0.4, v, 14, color=BODY)
box(s, 0.6, 5.6, 12.13, 1.1, fill=NAVY)
text(s, 0.6, 5.6, 12.13, 1.1, "Q & A", 30, True, WHITE, PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE)

prs.save(OUT)
print("Saved PPT to:", OUT)

try:
    prs.save(OUT_LEGACY)
    print("Saved PPT to:", OUT_LEGACY)
except Exception as e:
    print("Note: Legacy file locked, saved to primary:", OUT)
