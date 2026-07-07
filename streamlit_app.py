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


# =========================================================
# 유틸 함수
# =========================================================

def get_secret(key: str):
    """
    로컬에서는 .env에서 읽고,
    Streamlit Cloud에서는 st.secrets에서 읽습니다.
    """
    try:
        if key in st.secrets:
            return st.secrets[key]
    except Exception:
        pass

    return os.getenv(key)


def get_latest_report_file():
    """
    data/processed 폴더에서 가장 최근 생성된 통합 리포트 엑셀 파일을 찾습니다.
    """
    files = list(SAVE_PATH.glob("integrated_steel_demand_report_*.xlsx"))

    if not files:
        return None

    return max(files, key=lambda file: file.stat().st_mtime)


def run_integrated_report():
    """
    integrated_steel_report.py를 실행하여 최신 엑셀 리포트를 생성합니다.
    """
    result = subprocess.run(
        [sys.executable, str(REPORT_SCRIPT)],
        cwd=str(BASE_DIR),
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
    )

    return result


def show_download_button(report_file: Path, label: str):
    """
    엑셀 다운로드 버튼을 표시합니다.
    """
    with open(report_file, "rb") as file:
        st.download_button(
            label=label,
            data=file,
            file_name=report_file.name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )


# =========================================================
# 화면 설정
# =========================================================

st.set_page_config(
    page_title="철근 수요 예측 엑셀 리포트",
    page_icon="📥",
    layout="centered",
)

st.markdown(
    """
    <style>
    .main-title {
        font-size: 36px;
        font-weight: 800;
        margin-bottom: 8px;
    }
    .sub-title {
        font-size: 16px;
        color: #8A94A6;
        margin-bottom: 24px;
    }
    .box {
        border: 1px solid #2D3748;
        border-radius: 16px;
        padding: 20px;
        margin-bottom: 16px;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

st.markdown(
    """
    <div class="main-title">📥 철근 수요 예측 엑셀 리포트</div>
    <div class="sub-title">
        나라장터 공사입찰 데이터와 건설CALS 진행공사 데이터를 수집하여
        통합 엑셀 리포트를 생성합니다.
    </div>
    """,
    unsafe_allow_html=True,
)


# =========================================================
# API 키 확인
# =========================================================

nara_key = get_secret("NARA_API_KEY")
cals_key = get_secret("CALS_API_KEY")
archhub_key = get_secret("ARCHHUB_API_KEY")

missing_keys = []

if not nara_key:
    missing_keys.append("NARA_API_KEY")

if not cals_key:
    missing_keys.append("CALS_API_KEY")

# ARCHHUB_API_KEY는 향후 건축HUB 연동용입니다.
# 현재 integrated_steel_report.py가 건축HUB를 아직 사용하지 않는다면 없어도 실행 가능합니다.

if missing_keys:
    st.error("API 키가 설정되지 않았습니다.")
    st.code("\n".join(missing_keys))
    st.stop()


# =========================================================
# 최신 파일 다운로드
# =========================================================

latest_file = get_latest_report_file()

st.markdown("## 1. 기존 리포트 다운로드")

if latest_file:
    modified_time = datetime.fromtimestamp(latest_file.stat().st_mtime).strftime("%Y-%m-%d %H:%M:%S")
    st.success("다운로드 가능한 리포트가 있습니다.")
    st.caption(f"최근 생성 파일: {latest_file.name}")
    st.caption(f"생성/수정 시각: {modified_time}")
    show_download_button(latest_file, "📥 최신 엑셀 리포트 다운로드")
else:
    st.info("아직 생성된 엑셀 리포트가 없습니다. 아래 버튼으로 새 리포트를 생성하세요.")

st.divider()


# =========================================================
# 신규 리포트 생성
# =========================================================

st.markdown("## 2. 새 리포트 생성")

st.write(
    "아래 버튼을 누르면 최신 데이터를 수집하여 엑셀 리포트를 새로 생성합니다. "
    "데이터 수집량에 따라 약 1~3분 정도 소요될 수 있습니다."
)

if st.button("🚀 최신 데이터로 엑셀 리포트 생성", type="primary", use_container_width=True):

    if not REPORT_SCRIPT.exists():
        st.error(f"리포트 생성 파일을 찾을 수 없습니다: {REPORT_SCRIPT}")
        st.stop()

    if REPORT_SCRIPT.stat().st_size == 0:
        st.error("integrated_steel_report.py 파일이 비어 있습니다.")
        st.stop()

    with st.spinner("공공데이터 수집 및 엑셀 리포트 생성 중입니다..."):
        start_time = datetime.now()
        result = run_integrated_report()
        end_time = datetime.now()
        elapsed = (end_time - start_time).total_seconds()

    if result.returncode != 0:
        st.error("리포트 생성 중 오류가 발생했습니다.")

        with st.expander("오류 내용 보기"):
            st.code(result.stderr or "표준 오류 메시지가 없습니다.")

        with st.expander("실행 로그 보기"):
            st.code(result.stdout or "실행 로그가 없습니다.")

    else:
        new_file = get_latest_report_file()

        if new_file is None:
            st.warning("실행은 완료되었지만 생성된 엑셀 파일을 찾지 못했습니다.")
        else:
            st.success(f"리포트 생성 완료! 소요 시간: {elapsed:.1f}초")
            st.caption(f"생성 파일: {new_file.name}")
            show_download_button(new_file, "📥 방금 생성한 엑셀 리포트 다운로드")

            with st.expander("실행 로그 보기"):
                st.code(result.stdout or "실행 로그가 없습니다.")


st.divider()
st.caption("철근 수요 예측 통합 리포트 · 나라장터 × 건설CALS")
