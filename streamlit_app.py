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
    """
    로컬에서는 .env에서 읽고,
    Streamlit Cloud에서는 st.secrets에서 읽기 위한 함수
    """
    try:
        if key in st.secrets:
            return st.secrets[key]
    except Exception:
        pass

    return os.getenv(key)


# =========================================================
# 함수
# =========================================================

def get_latest_report_file():
    files = list(SAVE_PATH.glob("integrated_steel_demand_report_*.xlsx"))

    if not files:
        return None

    latest_file = max(files, key=lambda file: file.stat().st_mtime)

    return latest_file


def run_integrated_report():
    """
    로컬/배포 환경 모두에서 현재 실행 중인 Python으로
    integrated_steel_report.py 실행
    """
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

st.title("📊 철근 수요 예측 통합 리포트")
st.write("나라장터 입찰 데이터와 건설CALS 진행공사 데이터를 통합 분석합니다.")

st.divider()


# =========================================================
# API 키 상태 확인
# =========================================================

nara_key = get_secret("NARA_API_KEY")
cals_key = get_secret("CALS_API_KEY")

col1, col2 = st.columns(2)

with col1:
    if nara_key:
        st.success("나라장터 API 키 확인 완료")
    else:
        st.error("나라장터 API 키가 없습니다.")

with col2:
    if cals_key:
        st.success("건설CALS API 키 확인 완료")
    else:
        st.error("건설CALS API 키가 없습니다.")

st.divider()


# =========================================================
# 설명
# =========================================================

st.subheader("리포트 생성 내용")

st.write(
    """
    아래 버튼을 누르면 다음 작업이 자동으로 실행됩니다.

    1. 나라장터 공사 입찰 데이터 수집  
    2. 철근 관련 키워드 필터링  
    3. 취소공고 및 수의계약 제외  
    4. 건설CALS 진행공사 데이터 수집  
    5. 지역별·기관별·키워드별 집계  
    6. 통합 철근 수요지수 계산  
    7. 엑셀 리포트 생성  
    """
)

st.info("데이터 수집량에 따라 1~3분 정도 걸릴 수 있습니다.")


# =========================================================
# 리포트 생성 버튼
# =========================================================

if st.button("🚀 통합 리포트 생성하기", type="primary"):

    if not nara_key or not cals_key:
        st.error(".env 파일에 NARA_API_KEY와 CALS_API_KEY가 모두 있어야 합니다.")
        st.stop()

    if not REPORT_SCRIPT.exists():
        st.error(f"통합 리포트 파일을 찾을 수 없습니다: {REPORT_SCRIPT}")
        st.stop()

    if REPORT_SCRIPT.stat().st_size == 0:
        st.error("integrated_steel_report.py 파일이 비어 있습니다. 먼저 통합 리포트 코드를 넣어주세요.")
        st.stop()

    with st.spinner("통합 리포트를 생성하는 중입니다. 잠시만 기다려주세요..."):

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
            st.write("생성된 파일:")
            st.code(str(latest_file))

            with open(latest_file, "rb") as file:
                st.download_button(
                    label="📥 엑셀 리포트 다운로드",
                    data=file,
                    file_name=latest_file.name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            with st.expander("실행 로그 보기"):
                st.code(result.stdout)


st.divider()


# =========================================================
# 최근 리포트 다운로드
# =========================================================

st.subheader("최근 생성된 리포트")

latest_file = get_latest_report_file()

if latest_file is None:
    st.write("아직 생성된 통합 리포트가 없습니다.")
else:
    st.write(f"최근 파일: `{latest_file.name}`")

    with open(latest_file, "rb") as file:
        st.download_button(
            label="📥 최근 리포트 다운로드",
            data=file,
            file_name=latest_file.name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )


st.divider()

st.caption("MVP version 0.1 / 철근 수요 예측 통합 리포트")