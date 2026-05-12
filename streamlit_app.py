import os
import sys
import subprocess
from pathlib import Path
from datetime import datetime

import streamlit as st
from dotenv import load_dotenv


# =========================================================
# 기본 설정
# =========================================================

load_dotenv()

BASE_DIR = Path(__file__).parent

REPORT_SCRIPT = BASE_DIR / "integrated_steel_report.py"

SAVE_PATH = BASE_DIR / "data" / "processed"
SAVE_PATH.mkdir(parents=True, exist_ok=True)


def get_secret(key):
    try:
        if key in st.secrets:
            return st.secrets[key]
    except Exception:
        pass

    return os.getenv(key)


def get_latest_report_file():
    files = list(SAVE_PATH.glob("integrated_steel_demand_report_*.xlsx"))

    if not files:
        return None

    latest_file = max(files, key=lambda file: file.stat().st_mtime)

    return latest_file


def run_integrated_report():
    result = subprocess.run(
        [sys.executable, str(REPORT_SCRIPT)],
        cwd=str(BASE_DIR),
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace"
    )

    return result


# =========================================================
# 화면 설정
# =========================================================

st.set_page_config(
    page_title="철근 수요 예측 리포트",
    page_icon="📊",
    layout="wide"
)

# CSS
st.markdown(
    """
    <style>
    .main-title {
        font-size: 42px;
        font-weight: 800;
        margin-bottom: 0px;
    }
    .sub-title {
        font-size: 18px;
        color: #AAB2C0;
        margin-top: 4px;
        margin-bottom: 24px;
    }
    .hero-box {
        background: linear-gradient(135deg, #102A43 0%, #1F4E78 55%, #2F5597 100%);
        padding: 30px;
        border-radius: 18px;
        color: white;
        margin-bottom: 22px;
        box-shadow: 0 8px 24px rgba(0,0,0,0.22);
    }
    .hero-small {
        font-size: 15px;
        color: #D9EAF7;
    }
    .metric-card {
        background-color: #111827;
        border: 1px solid #253044;
        border-radius: 16px;
        padding: 18px;
        min-height: 120px;
    }
    .metric-title {
        color: #AAB2C0;
        font-size: 14px;
        margin-bottom: 8px;
    }
    .metric-value {
        color: white;
        font-size: 28px;
        font-weight: 700;
    }
    .section-box {
        border: 1px solid #2D3748;
        padding: 20px;
        border-radius: 16px;
        margin-bottom: 14px;
    }
    </style>
    """,
    unsafe_allow_html=True
)

# =========================================================
# 헤더
# =========================================================

st.markdown(
    """
    <div class="hero-box">
        <div class="main-title">📊 철근 수요 예측 통합 리포트</div>
        <div class="sub-title">
            나라장터 입찰 데이터와 건설CALS 진행공사 데이터를 통합해
            지역별 철근 수요 흐름과 영업 우선순위를 자동으로 분석합니다.
        </div>
        <div class="hero-small">
            Public Procurement Intelligence · Construction Demand Signal · Sales Priority Report
        </div>
    </div>
    """,
    unsafe_allow_html=True
)

# =========================================================
# API 키 상태
# =========================================================

nara_key = get_secret("NARA_API_KEY")
cals_key = get_secret("CALS_API_KEY")

col1, col2, col3 = st.columns(3)

with col1:
    if nara_key:
        st.success("나라장터 API 연결 준비 완료")
    else:
        st.error("나라장터 API 키 없음")

with col2:
    if cals_key:
        st.success("건설CALS API 연결 준비 완료")
    else:
        st.error("건설CALS API 키 없음")

with col3:
    latest_file = get_latest_report_file()
    if latest_file:
        st.info("최근 리포트 사용 가능")
    else:
        st.warning("최근 리포트 없음")


# =========================================================
# 상단 다운로드 영역
# =========================================================

st.markdown("## 📥 리포트 다운로드")

latest_file = get_latest_report_file()

download_col1, download_col2 = st.columns([1, 2])

with download_col1:
    if latest_file:
        with open(latest_file, "rb") as file:
            st.download_button(
                label="📥 최신 엑셀 리포트 다운로드",
                data=file,
                file_name=latest_file.name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
    else:
        st.button(
            "📥 최신 리포트 없음",
            disabled=True,
            use_container_width=True
        )

with download_col2:
    if latest_file:
        st.success(f"최근 생성 파일: {latest_file.name}")
    else:
        st.info("아직 생성된 리포트가 없습니다. 아래 버튼으로 새 리포트를 생성하세요.")

st.divider()


# =========================================================
# 안내 카드
# =========================================================

st.markdown("## 🧭 분석 개요")

info_col1, info_col2, info_col3 = st.columns(3)

with info_col1:
    st.markdown(
        """
        <div class="metric-card">
            <div class="metric-title">나라장터 입찰 데이터</div>
            <div class="metric-value">선행 수요</div>
            <div class="hero-small">향후 발주 가능성이 있는 공사 입찰 흐름을 추적합니다.</div>
        </div>
        """,
        unsafe_allow_html=True
    )

with info_col2:
    st.markdown(
        """
        <div class="metric-card">
            <div class="metric-title">건설CALS 진행공사</div>
            <div class="metric-value">현재 수요</div>
            <div class="hero-small">현재 진행 중인 공사를 기반으로 지역별 수요를 확인합니다.</div>
        </div>
        """,
        unsafe_allow_html=True
    )

with info_col3:
    st.markdown(
        """
        <div class="metric-card">
            <div class="metric-title">통합 수요지수</div>
            <div class="metric-value">영업 우선순위</div>
            <div class="hero-small">입찰 흐름과 진행공사를 결합해 우선 공략 지역을 도출합니다.</div>
        </div>
        """,
        unsafe_allow_html=True
    )

st.divider()


# =========================================================
# 리포트 생성
# =========================================================

st.markdown("## 🚀 신규 리포트 생성")

st.markdown(
    """
    아래 버튼을 누르면 최신 공공데이터를 수집해 엑셀 리포트를 자동 생성합니다.

    - 나라장터 공사 입찰 데이터 수집  
    - 철근 관련 키워드 필터링  
    - 취소공고 및 수의계약 제외  
    - 입찰마감일 기준 영업긴급도 산정  
    - 건설CALS 진행공사 데이터 수집  
    - 준공예정일 기준 CALS 영업긴급도 산정  
    - 지역별·기관별·키워드별 집계  
    - 통합 철근 수요지수 계산  
    - 엑셀 요약 리포트 생성  
    """
)

st.warning("데이터 수집량에 따라 약 1~3분 정도 소요될 수 있습니다.")

if st.button("🚀 최신 통합 리포트 생성하기", type="primary", use_container_width=True):

    if not nara_key or not cals_key:
        st.error("API 키가 설정되어 있지 않습니다. .env 또는 Streamlit Secrets를 확인하세요.")
        st.stop()

    if not REPORT_SCRIPT.exists():
        st.error(f"통합 리포트 파일을 찾을 수 없습니다: {REPORT_SCRIPT}")
        st.stop()

    if REPORT_SCRIPT.stat().st_size == 0:
        st.error("integrated_steel_report.py 파일이 비어 있습니다.")
        st.stop()

    with st.spinner("공공데이터 수집 및 통합 리포트를 생성하는 중입니다..."):
        start_time = datetime.now()
        result = run_integrated_report()
        end_time = datetime.now()
        elapsed = (end_time - start_time).total_seconds()

    if result.returncode != 0:
        st.error("리포트 생성 중 오류가 발생했습니다.")

        with st.expander("오류 내용 보기"):
            st.code(result.stderr)

        with st.expander("실행 로그 보기"):
            st.code(result.stdout)

    else:
        st.success(f"리포트 생성 완료! 소요 시간: {elapsed:.1f}초")

        latest_file = get_latest_report_file()

        if latest_file is None:
            st.warning("생성된 엑셀 파일을 찾지 못했습니다.")
        else:
            st.balloons()

            st.markdown("### 📌 생성된 리포트")
            st.code(str(latest_file))

            with open(latest_file, "rb") as file:
                st.download_button(
                    label="📥 방금 생성한 엑셀 리포트 다운로드",
                    data=file,
                    file_name=latest_file.name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )

            with st.expander("실행 로그 보기"):
                st.code(result.stdout)


st.divider()

st.caption("MVP version 0.2 · 철근 수요 예측 통합 리포트 · 나라장터 × 건설CALS")
