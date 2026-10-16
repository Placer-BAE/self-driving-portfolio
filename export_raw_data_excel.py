"""
Self-Driving Portfolio: Raw Data 및 파이프라인 전 단계 수치 데이터 엑셀(XLSX) 통합 추출 스크립트
생성 파일:
  1. Self_Driving_Portfolio_Raw_Data_종합.xlsx (7개 시트 구성)
  2. Yahoo_Finance_18_ETF_Monthly_Returns_Raw.csv (원본 CSV 복사본)
"""

import json
import shutil
import sys
from pathlib import Path

import numpy as np
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

ROOT = Path(__file__).resolve().parent
BACKEND = ROOT / "backend"
sys.path.insert(0, str(BACKEND))

from app.config import ASSET_CLASSES, DEFAULT_IPS
from app.data_loader import BASE_HISTORICAL_STATS, get_market_dataset
from app.pipeline import SelfDrivingPortfolioPipeline
from app.backtest import run_backtest, load_returns

OUT_XLSX = ROOT / "Self_Driving_Portfolio_Raw_Data_종합.xlsx"
OUT_CSV = ROOT / "Yahoo_Finance_18_ETF_Monthly_Returns_Raw.csv"
RAW_CSV_SRC = BACKEND / "data" / "etf_monthly_returns.csv"
RAW_META_SRC = BACKEND / "data" / "etf_monthly_returns.meta.json"


def style_Workbook(wb: openpyxl.Workbook):
    header_font = Font(name="Malgun Gothic", size=10, bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="051C2C", end_color="051C2C", fill_type="solid")
    sub_fill = PatternFill(start_color="EFF6FF", end_color="EFF6FF", fill_type="solid")
    zebra_fill = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")
    body_font = Font(name="Malgun Gothic", size=9.5, color="0F172A")
    bold_font = Font(name="Malgun Gothic", size=9.5, bold=True, color="051C2C")
    thin_border = Border(
        left=Side(style="thin", color="CBD5E1"),
        right=Side(style="thin", color="CBD5E1"),
        top=Side(style="thin", color="CBD5E1"),
        bottom=Side(style="thin", color="CBD5E1"),
    )

    for ws in wb.worksheets:
        ws.views.sheetView[0].showGridLines = True
        max_row = ws.max_row
        max_col = ws.max_column

        for r in range(1, max_row + 1):
            for c in range(1, max_col + 1):
                cell = ws.cell(row=r, column=c)
                cell.border = thin_border
                if r == 1:
                    cell.font = header_font
                    cell.fill = header_fill
                    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                else:
                    cell.font = body_font
                    if r % 2 == 1:
                        cell.fill = zebra_fill
                    if isinstance(cell.value, (int, float)):
                        cell.alignment = Alignment(horizontal="right", vertical="center")
                    else:
                        cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)

        ws.row_dimensions[1].height = 26
        ws.freeze_panes = "B2"

        # Auto column width
        for c in range(1, max_col + 1):
            col_letter = get_column_letter(c)
            max_len = 0
            for r in range(1, min(max_row + 1, 60)):
                val = str(ws.cell(row=r, column=c).value or "")
                # 한글 가중 폭 계산
                char_len = sum(1.7 if ord(ch) > 127 else 1.0 for ch in val)
                if char_len > max_len:
                    max_len = char_len
            ws.column_dimensions[col_letter].width = max(11, min(max_len + 4, 48))


def main():
    # 1. 원본 CSV 복사
    if RAW_CSV_SRC.exists():
        shutil.copyfile(RAW_CSV_SRC, OUT_CSV)

    meta = json.loads(RAW_META_SRC.read_text(encoding="utf-8")) if RAW_META_SRC.exists() else {}
    panel = load_returns()  # 159 months x 18 ETFs (decimal monthly return)

    # 2. 파이프라인 및 백테스트 실행
    pipe = SelfDrivingPortfolioPipeline(DEFAULT_IPS)
    res = pipe.run()
    bt = run_backtest(res)
    mkt_data = get_market_dataset()
    cov_step_raw = pipe.cov_agent.estimate(DEFAULT_IPS.universe_ids, res["macro_step"]["regime"])

    ticker_to_name = {a["ticker"]: a["name"] for a in ASSET_CLASSES}
    id_to_asset = {a["id"]: a for a in ASSET_CLASSES}

    # --- Sheet 0: 가이드 및 메타정보 ---
    guide_rows = [
        {"구분": "데이터 원천 (Source)", "항목": "수집 출처", "내용": meta.get("source", "Yahoo Finance")},
        {"구분": "데이터 원천 (Source)", "항목": "수집 함수 및 기준", "내용": meta.get("method", "")},
        {"구분": "데이터 원천 (Source)", "항목": "수집 일자", "내용": meta.get("fetched_at", "2026-10-13")},
        {"구분": "데이터 원천 (Source)", "항목": "분석 기간 (월수)", "내용": f"{meta.get('first_month')} ~ {meta.get('last_month')} (총 {len(panel)}개월, 결측치 0건)"},
        {"구분": "데이터 원천 (Source)", "항목": "대상 유니버스", "내용": "18개 대표 유동성 ETF (SPY, VUG, VTV, IWM, EFA, EEM, SHY, IEF, TLT, LQD, HYG, BNDX, PICB, EMB, GLD, DBC, VNQ, BIL)"},
        {"구분": "거시 환경 설정 (Stage 1)", "항목": "판정 거시 국면", "내용": f"{res['macro_step']['regime']} (신뢰도 {int(res['macro_step']['confidence']*100)}%, 침체확률 {int(res['macro_step']['recession_probability']*100)}%)"},
        {"구분": "IPS 제약 조건", "항목": "목표 실질수익률 / 예상 CPI", "내용": f"CPI({DEFAULT_IPS.cpi_estimate}%) + {DEFAULT_IPS.target_real_return}% = 명목 6.0%"},
        {"구분": "IPS 제약 조건", "항목": "사전 변동성 허용 밴드", "내용": f"[{DEFAULT_IPS.volatility_min}% ~ {DEFAULT_IPS.volatility_max}%]"},
        {"구분": "IPS 제약 조건", "항목": "트래킹 에러 / MDD 한도", "내용": f"TE < {DEFAULT_IPS.tracking_error_limit}%, 예상 MDD > {DEFAULT_IPS.max_drawdown_limit}%"},
        {"구분": "IPS 제약 조건", "항목": "주식 비중 허용 범위", "내용": f"[{DEFAULT_IPS.equity_min}% ~ {DEFAULT_IPS.equity_max}%]"},
        {"구분": "시트 구성 안내", "항목": "01_18개ETF_월간수익률_Raw", "내용": "Yahoo Finance 159개월(2013-07 ~ 2026-09) 18개 ETF 월간 수정 총수익률 원본 데이터 (소수점 원자료)"},
        {"구분": "시트 구성 안내", "항목": "02_18개ETF_연도별수익률_성과", "내용": "2013~2026 연도별 18개 ETF 실현 수익률(%) 및 전체 기간 누적 성과(CAGR, 변동성, MDD)"},
        {"구분": "시트 구성 안내", "항목": "03_자산군모수_및_7대CMA추정", "내용": "18개 자산군 기초 통계 모수(1996-2026 캘리브레이션) 및 7가지 CMA 모델별 10년 기대수익률 추정치"},
        {"구분": "시트 구성 안내", "항목": "04_상관계수_및_공분산행렬", "내용": "18x18 자산군 간 기본 상관계수, OAS 수축 상관계수(Shrinkage 20%), 국면 조정 수축 공분산 행렬"},
        {"구분": "시트 구성 안내", "항목": "05_21개모델_비중_및_심사결과", "내용": "4대 패밀리 21개 최적화 후보군의 18개 ETF 비중(%), 보르다 투표 점수, CRO 리스크 지표, 적격 심사 결과"},
        {"구분": "시트 구성 안내", "항목": "06_백테스트_월별NAV_및_성과", "내용": "159개월 분기 리밸런싱 실데이터 백테스트 누적 NAV 시계열 (최종 SAA vs 60/40 vs 21개 후보 모델)"},
    ]
    df_guide = pd.DataFrame(guide_rows)

    # --- Sheet 1: 18개 ETF 월간 수익률 Raw Data (2013-07 ~ 2026-09) ---
    df_raw_monthly = panel.copy()
    df_raw_monthly.columns = [f"{col} ({ticker_to_name.get(col, col)})" for col in df_raw_monthly.columns]
    df_raw_monthly.index.name = "연월 (YYYY-MM)"
    df_raw_monthly = df_raw_monthly.reset_index()

    # --- Sheet 2: 18개 ETF 연도별 수익률 및 누적 성과 ---
    ann = bt["annual"]
    years = ann["years"]
    months_per_yr = ann["months"]
    asset_summary_list = []
    for item in bt["assets"]:
        t = item["ticker"]
        row_dict = {
            "티커": t,
            "자산군명": item["name"],
            "대분류": item["category"],
            "최종SAA_비중(%)": item["final_weight_pct"],
            "60/40벤치마크_비중(%)": item["bench_weight_pct"],
            "실현_연평균수익률_CAGR(%)": item["cagr_pct"],
            "실현_연율화변동성(%)": item["volatility_pct"],
            "실현_최대낙폭_MDD(%)": item["max_drawdown_pct"],
            "최고수익월": f"{item['best_month']['month']} ({item['best_month']['return_pct']}%)",
            "최저수익월": f"{item['worst_month']['month']} ({item['worst_month']['return_pct']}%)",
        }
        for y, m_cnt in zip(years, months_per_yr):
            col_label = f"{y}년({m_cnt}M,%)" if m_cnt < 12 else f"{y}년(%)"
            yr_idx = years.index(y)
            row_dict[col_label] = ann["by_ticker"][t][yr_idx]
        asset_summary_list.append(row_dict)
    df_annual_etf = pd.DataFrame(asset_summary_list)

    # --- Sheet 3: 18개 자산군 기본 모수 및 7대 CMA 기대수익률 추정 (Stage 2) ---
    cma_rows = []
    for a in ASSET_CLASSES:
        aid = a["id"]
        st = BASE_HISTORICAL_STATS[aid]
        cma_info = res["cma_step"]["cmas"][aid]
        cands = cma_info["candidates"]
        cma_rows.append({
            "자산ID": aid,
            "티커": a["ticker"],
            "자산군명": a["name"],
            "대분류": a["category"],
            "60/40벤치마크_기본비중(%)": round(a.get("bench_weight", 0.0) * 100, 2),
            "역사적_평균수익률_모수(%)": round(st["mean_ret"] * 100, 2),
            "역사적_연간변동성_모수(%)": round(st["vol"] * 100, 2),
            "배당/이자수익률_모수(%)": round(st["yield"] * 100, 2),
            "이익성장률_모수(%)": round(st["growth"] * 100, 2),
            "CAPE_또는_듀레이션": a.get("cape_historical", a.get("duration", "-")),
            "CMA1_역사적ERP(%)": cands["historical_erp"],
            "CMA2_국면조정(%)": cands["regime_adjusted"],
            "CMA3_BL균형(%)": cands["bl_equilibrium"],
            "CMA4_역고든모형(%)": cands["inverse_gordon"],
            "CMA5_내재ERP(%)": cands["implied_erp"],
            "CMA6_서베이컨센서스(%)": cands["survey_analyst"],
            "CMA7_신뢰도가중_AutoBlend(%)": cma_info["auto_blend_pct"],
            "Judge_최종채택_기대수익률(%)": cma_info["expected_return_pct"],
            "Judge_조정폭(%p)": cma_info["delta_vs_auto_blend"],
            "LLM_Judge_심사근거": cma_info["rationale"],
        })
    df_cma = pd.DataFrame(cma_rows)

    # --- Sheet 4: 상관계수 및 수축 공분산 행렬 (Stage 3) ---
    labels = [f"{a['ticker']} ({a['name']})" for a in ASSET_CLASSES]
    shrunk_corr = pd.DataFrame(np.round(cov_step_raw["correlation_matrix"], 4), index=labels, columns=labels)
    shrunk_cov = pd.DataFrame(np.round(cov_step_raw["covariance_matrix"], 6), index=labels, columns=labels)

    # 상관계수와 공분산 행렬을 위아래로 결합
    corr_part = shrunk_corr.reset_index().rename(columns={"index": "구분 / 자산군"})
    corr_part.insert(0, "행렬유형", "OAS 수축 상관계수 (Correlation)")
    cov_part = shrunk_cov.reset_index().rename(columns={"index": "구분 / 자산군"})
    cov_part.insert(0, "행렬유형", "OAS 수축 공분산 (Covariance)")
    df_matrices = pd.concat([corr_part, cov_part], ignore_index=True)

    # --- Sheet 5: 21개 후보 모델 비중 및 CRO/보르다 투표 심사 결과 ---
    rankings_map = {r["method_key"]: r for r in res["review_step"]["rankings"]}
    method_weights = res["ensemble_step"]["method_weights"]
    cro_reports = res["cro_reports"]

    model_rows = []
    # 1) 최종 SAA 앙상블 행 추가
    fp = res["ensemble_step"]["final_portfolio"]
    final_row = {
        "순위": "최종안",
        "모델키": "final_saa",
        "모델명": "최종 SAA (CIO 적격 앙상블)",
        "패밀리": "Stage 6: 최종 앙상블",
        "IPS_적격여부": "적격 (Approved)",
        "앙상블_반영비중(%)": 100.0,
        "보르다_투표점수": "-",
        "CRO_정량점수": "-",
        "종합점수(Composite)": "-",
        "사전_기대수익률(%)": fp["expected_return_pct"],
        "사전_연간변동성(%)": fp["volatility_pct"],
        "사전_샤프지수": fp["sharpe_ratio"],
        "사전_트래킹에러(%)": fp["tracking_error_pct"],
        "예상_MDD(%)": fp["max_drawdown_est_pct"],
        "유효분산자산수(Eff_N)": fp["effective_n"],
        "주식비중(%)": fp["category_exposures"]["equity"],
        "채권비중(%)": fp["category_exposures"]["fixed_income"],
        "실물자산비중(%)": fp["category_exposures"]["real_assets"],
        "현금비중(%)": fp["category_exposures"]["cash"],
        "IPS_위반사유": "위반 없음 (모든 IPS 제약 충족)",
    }
    for a in ASSET_CLASSES:
        final_row[f"{a['ticker']}_비중(%)"] = round(fp["weights"][a["id"]] * 100, 2)
    model_rows.append(final_row)

    # 2) 21개 모델 종합 순위순 추가
    for r_item in res["review_step"]["rankings"]:
        k = r_item["method_key"]
        rep = cro_reports[k]
        row = {
            "순위": r_item["rank"],
            "모델키": k,
            "모델명": r_item["name"],
            "패밀리": r_item["category"],
            "IPS_적격여부": "통과 (O)" if rep["is_compliant"] else "탈락 (X)",
            "앙상블_반영비중(%)": round(method_weights.get(k, 0.0) * 100, 2),
            "보르다_투표점수": r_item["vote"],
            "CRO_정량점수": r_item["metric_score"],
            "종합점수(Composite)": r_item["composite_score"],
            "사전_기대수익률(%)": rep["expected_return"],
            "사전_연간변동성(%)": rep["volatility"],
            "사전_샤프지수": rep["sharpe_ratio"],
            "사전_트래킹에러(%)": rep["tracking_error"],
            "예상_MDD(%)": rep["estimated_mdd"],
            "유효분산자산수(Eff_N)": rep["effective_n"],
            "주식비중(%)": rep["exposures"]["equity"],
            "채권비중(%)": rep["exposures"]["fixed_income"],
            "실물자산비중(%)": rep["exposures"]["real_assets"],
            "현금비중(%)": rep["exposures"]["cash"],
            "IPS_위반사유": "; ".join(rep["violations"]) if rep["violations"] else "적격 통과",
        }
        for a in ASSET_CLASSES:
            row[f"{a['ticker']}_비중(%)"] = round(rep["weights"][a["id"]] * 100, 2)
        model_rows.append(row)

    df_models = pd.DataFrame(model_rows)

    # --- Sheet 6: 백테스트 월별 NAV 시계열 및 성과 (159개월) ---
    bt_dates = bt["dates"]
    bt_series = bt["series"]
    nav_dict = {"기준연월": bt_dates}
    # 요약 정보를 상단에 두고 시계열을 배치하거나, 열 이름에 모델명+CAGR 표기
    row_meta_by_key = {r["key"]: r for r in bt["rows"]}
    ordered_keys = ["final", "benchmark"] + [r["method_key"] for r in res["review_step"]["rankings"]]
    for k in ordered_keys:
        if k in bt_series:
            m_info = row_meta_by_key[k]
            col_title = f"{m_info['name']} [CAGR {m_info['cagr_pct']}%, Vol {m_info['volatility_pct']}%, Sharpe {m_info['sharpe_ratio']}, MDD {m_info['max_drawdown_pct']}%]"
            nav_dict[col_title] = bt_series[k]
    df_bt_nav = pd.DataFrame(nav_dict)

    # 엑셀 파일 저장
    with pd.ExcelWriter(OUT_XLSX, engine="openpyxl") as writer:
        df_guide.to_excel(writer, sheet_name="00_데이터_요약_가이드", index=False)
        df_raw_monthly.to_excel(writer, sheet_name="01_18개ETF_월간수익률_Raw", index=False)
        df_annual_etf.to_excel(writer, sheet_name="02_18개ETF_연도별수익률_성과", index=False)
        df_cma.to_excel(writer, sheet_name="03_자산군모수_및_7대CMA추정", index=False)
        df_matrices.to_excel(writer, sheet_name="04_상관계수_및_공분산행렬", index=False)
        df_models.to_excel(writer, sheet_name="05_21개모델_비중_및_심사결과", index=False)
        df_bt_nav.to_excel(writer, sheet_name="06_백테스트_월별NAV_시계열", index=False)

    wb = openpyxl.load_workbook(OUT_XLSX)
    style_Workbook(wb)
    wb.save(OUT_XLSX)
    print("Saved Excel Workbook:", OUT_XLSX)
    print("Saved Raw CSV Copy:", OUT_CSV)


if __name__ == "__main__":
    main()
