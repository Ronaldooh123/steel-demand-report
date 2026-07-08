import os
import math
import time
import json
import re
import requests
import urllib3
import pandas as pd

from numbers import Number
from pathlib import Path
from datetime import datetime, timedelta
from dotenv import load_dotenv
from xlsxwriter.utility import xl_col_to_name


# =========================================================
# 기본 설정
# =========================================================

load_dotenv()

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)


def get_secret(key):
    """
    로컬에서는 .env에서 읽고,
    Streamlit Cloud에서는 st.secrets에서 읽기 위한 함수
    """
    try:
        import streamlit as st

        if key in st.secrets:
            return st.secrets[key]

    except Exception:
        pass

    return os.getenv(key)


NARA_API_KEY = get_secret("NARA_API_KEY")
CALS_API_KEY = get_secret("CALS_API_KEY")
ARCHHUB_API_KEY = get_secret("ARCHHUB_API_KEY") or get_secret("NARA_API_KEY")


BASE_DIR = Path(__file__).parent

SAVE_PATH = BASE_DIR / "data" / "processed"
SAVE_PATH.mkdir(parents=True, exist_ok=True)

CACHE_PATH = BASE_DIR / "data" / "cache"
CACHE_PATH.mkdir(parents=True, exist_ok=True)

NOW = datetime.now().strftime("%Y%m%d_%H%M%S")
NARA_LOOKBACK_DAYS = int(os.getenv("NARA_LOOKBACK_DAYS", "7"))

ARCHHUB_ENABLED = os.getenv("ARCHHUB_ENABLED", "true").lower() not in ["0", "false", "no", "n"]
ARCHHUB_LOOKBACK_DAYS = int(os.getenv("ARCHHUB_LOOKBACK_DAYS", "90"))
ARCHHUB_MODE = os.getenv("ARCHHUB_MODE", "sample")  # single 또는 sample
ARCHHUB_SIGUNGU_CD = os.getenv("ARCHHUB_SIGUNGU_CD", "11680")
ARCHHUB_BJDONG_CD = os.getenv("ARCHHUB_BJDONG_CD", "")
ARCHHUB_MAX_PAGES = int(os.getenv("ARCHHUB_MAX_PAGES", "1"))  # 0이면 전체 페이지
ARCHHUB_NUM_ROWS = int(os.getenv("ARCHHUB_NUM_ROWS", "100"))
ARCHHUB_CACHE_HOURS = int(os.getenv("ARCHHUB_CACHE_HOURS", "24"))
ARCHHUB_FORCE_REFRESH = os.getenv("ARCHHUB_FORCE_REFRESH", "false").lower() in ["1", "true", "yes", "y"]

ARCHHUB_BASIS_URL = "https://apis.data.go.kr/1613000/ArchPmsHubService/getApBasisOulnInfo"



def require_api_key(value, key_name):
    if value:
        return value

    raise ValueError(f"{key_name}를 찾을 수 없습니다. .env, Streamlit Secrets, GitHub Secrets를 확인하세요.")


ARCHHUB_SIDO_PREFIX_MAP = {
    "11": "서울",
    "26": "부산",
    "27": "대구",
    "28": "인천",
    "29": "광주",
    "30": "대전",
    "31": "울산",
    "36": "세종",
    "41": "경기",
    "42": "강원",
    "43": "충북",
    "44": "충남",
    "45": "전북",
    "46": "전남",
    "47": "경북",
    "48": "경남",
    "50": "제주",
    "51": "강원",
    "52": "전북",
}

# 1차 MVP용 대표 시군구 샘플입니다.
# 전국 전체 수집은 법정동 코드 목록 확보 후 별도 확장하는 것을 권장합니다.
ARCHHUB_SAMPLE_SIGUNGU_CODES = {
    "서울_강남구": "11680",
    "서울_송파구": "11710",
    "경기_성남분당구": "41135",
    "경기_화성시": "41590",
    "인천_연수구": "28185",
    "부산_해운대구": "26350",
    "대구_수성구": "27260",
    "대전_유성구": "30200",
    "광주_광산구": "29200",
    "울산_울주군": "31710",
    "세종": "36110",
    "충북_청주시흥덕구": "43113",
    "충남_천안서북구": "44133",
    "전북_전주덕진구": "45113",
    "전남_나주시": "46170",
    "경북_포항북구": "47113",
    "경남_창원성산구": "48123",
    "강원_원주시": "51130",
    "제주_제주시": "50110",
}


# =========================================================
# 공통 함수
# =========================================================

def classify_region_by_text(text):
    """
    수요기관/공고기관/공고명에 포함된 지명으로 광역 지역을 판정합니다.
    기존의 단순 광역명 매칭보다 시·군·구/출장소/지방청 명칭까지 폭넓게 반영합니다.
    """
    text = str(text).strip()

    if text == "":
        return "기타"

    compact_text = text.replace(" ", "")

    # 1. 광역 시·도 및 흔한 축약명
    direct_region_map = {
        "서울특별시": "서울",
        "서울시": "서울",
        "서울": "서울",

        "부산광역시": "부산",
        "부산시": "부산",
        "부산": "부산",

        "대구광역시": "대구",
        "대구시": "대구",
        "대구": "대구",

        "인천광역시": "인천",
        "인천시": "인천",
        "인천": "인천",

        "광주광역시": "광주",
        "광주광역": "광주",

        "대전광역시": "대전",
        "대전시": "대전",
        "대전": "대전",

        "울산광역시": "울산",
        "울산시": "울산",
        "울산": "울산",

        "세종특별자치시": "세종",
        "세종시": "세종",
        "세종": "세종",

        "경기도": "경기",
        "경기": "경기",

        "강원특별자치도": "강원",
        "강원도": "강원",
        "강원": "강원",

        "충청북도": "충북",
        "충북": "충북",

        "충청남도": "충남",
        "충남": "충남",

        "전북특별자치도": "전북",
        "전라북도": "전북",
        "전북": "전북",

        "전라남도": "전남",
        "전남": "전남",

        "경상북도": "경북",
        "경북": "경북",

        "경상남도": "경남",
        "경남": "경남",

        "제주특별자치도": "제주",
        "제주도": "제주",
        "제주": "제주",
    }

    for keyword, region in sorted(direct_region_map.items(), key=lambda x: len(x[0]), reverse=True):
        if keyword in text or keyword in compact_text:
            return region

    # 2. 시·군·구 / 주요 지명 / 지방청 명칭
    city_region_map = {
        # 서울
        "종로구": "서울",
        "서울중구": "서울",
        "용산구": "서울",
        "성동구": "서울",
        "광진구": "서울",
        "동대문구": "서울",
        "중랑구": "서울",
        "성북구": "서울",
        "강북구": "서울",
        "도봉구": "서울",
        "노원구": "서울",
        "은평구": "서울",
        "서대문구": "서울",
        "마포구": "서울",
        "양천구": "서울",
        "강서구": "서울",
        "구로구": "서울",
        "금천구": "서울",
        "영등포구": "서울",
        "동작구": "서울",
        "관악구": "서울",
        "서초구": "서울",
        "강남구": "서울",
        "송파구": "서울",
        "강동구": "서울",
        "서울지방": "서울",

        # 경기
        "수원시": "경기",
        "수원": "경기",
        "성남시": "경기",
        "성남": "경기",
        "고양시": "경기",
        "고양": "경기",
        "용인시": "경기",
        "용인": "경기",
        "부천시": "경기",
        "부천": "경기",
        "안산시": "경기",
        "안산": "경기",
        "안양시": "경기",
        "안양": "경기",
        "남양주시": "경기",
        "남양주": "경기",
        "화성시": "경기",
        "화성": "경기",
        "평택시": "경기",
        "평택": "경기",
        "의정부시": "경기",
        "의정부": "경기",
        "시흥시": "경기",
        "시흥": "경기",
        "파주시": "경기",
        "파주": "경기",
        "김포시": "경기",
        "김포": "경기",
        "광명시": "경기",
        "광명": "경기",
        "경기도광주시": "경기",
        "광주시": "경기",
        "군포시": "경기",
        "군포": "경기",
        "오산시": "경기",
        "오산": "경기",
        "이천시": "경기",
        "이천": "경기",
        "안성시": "경기",
        "안성": "경기",
        "의왕시": "경기",
        "의왕": "경기",
        "하남시": "경기",
        "하남": "경기",
        "여주시": "경기",
        "여주": "경기",
        "양주시": "경기",
        "양주": "경기",
        "동두천시": "경기",
        "동두천": "경기",
        "구리시": "경기",
        "구리": "경기",
        "포천시": "경기",
        "포천": "경기",
        "과천시": "경기",
        "과천": "경기",
        "가평군": "경기",
        "가평": "경기",
        "양평군": "경기",
        "양평": "경기",
        "연천군": "경기",
        "연천": "경기",

        # 강원
        "춘천시": "강원",
        "춘천": "강원",
        "원주시": "강원",
        "원주": "강원",
        "강릉시": "강원",
        "강릉": "강원",
        "동해시": "강원",
        "동해": "강원",
        "태백시": "강원",
        "태백": "강원",
        "속초시": "강원",
        "속초": "강원",
        "삼척시": "강원",
        "삼척": "강원",
        "홍천군": "강원",
        "홍천": "강원",
        "횡성군": "강원",
        "횡성": "강원",
        "영월군": "강원",
        "영월": "강원",
        "평창군": "강원",
        "평창": "강원",
        "정선군": "강원",
        "정선": "강원",
        "철원군": "강원",
        "철원": "강원",
        "화천군": "강원",
        "화천": "강원",
        "양구군": "강원",
        "양구": "강원",
        "인제군": "강원",
        "인제": "강원",
        "강원고성": "강원",
        "양양군": "강원",
        "양양": "강원",

        # 충북
        "청주시": "충북",
        "청주": "충북",
        "충주시": "충북",
        "충주": "충북",
        "제천시": "충북",
        "제천": "충북",
        "보은군": "충북",
        "보은": "충북",
        "옥천군": "충북",
        "옥천": "충북",
        "영동군": "충북",
        "영동": "충북",
        "증평군": "충북",
        "증평": "충북",
        "진천군": "충북",
        "진천": "충북",
        "괴산군": "충북",
        "괴산": "충북",
        "음성군": "충북",
        "음성": "충북",
        "단양군": "충북",
        "단양": "충북",

        # 충남
        "천안시": "충남",
        "천안": "충남",
        "공주시": "충남",
        "공주": "충남",
        "보령시": "충남",
        "보령": "충남",
        "아산시": "충남",
        "아산": "충남",
        "서산시": "충남",
        "서산": "충남",
        "논산시": "충남",
        "논산": "충남",
        "계룡시": "충남",
        "계룡": "충남",
        "당진시": "충남",
        "당진": "충남",
        "금산군": "충남",
        "금산": "충남",
        "부여군": "충남",
        "부여": "충남",
        "서천군": "충남",
        "서천": "충남",
        "청양군": "충남",
        "청양": "충남",
        "홍성군": "충남",
        "홍성": "충남",
        "예산군": "충남",
        "예산": "충남",
        "태안군": "충남",
        "태안": "충남",
        "충남지역본부": "충남",

        # 전북
        "전주시": "전북",
        "전주": "전북",
        "군산시": "전북",
        "군산": "전북",
        "익산시": "전북",
        "익산": "전북",
        "정읍시": "전북",
        "정읍": "전북",
        "남원시": "전북",
        "남원": "전북",
        "김제시": "전북",
        "김제": "전북",
        "완주군": "전북",
        "완주": "전북",
        "진안군": "전북",
        "진안": "전북",
        "무주군": "전북",
        "무주": "전북",
        "장수군": "전북",
        "장수": "전북",
        "임실군": "전북",
        "임실": "전북",
        "순창군": "전북",
        "순창": "전북",
        "고창군": "전북",
        "고창": "전북",
        "부안군": "전북",
        "부안": "전북",

        # 전남
        "목포시": "전남",
        "목포": "전남",
        "여수시": "전남",
        "여수": "전남",
        "순천시": "전남",
        "순천": "전남",
        "나주시": "전남",
        "나주": "전남",
        "광양시": "전남",
        "광양": "전남",
        "담양군": "전남",
        "담양": "전남",
        "곡성군": "전남",
        "곡성": "전남",
        "구례군": "전남",
        "구례": "전남",
        "고흥군": "전남",
        "고흥": "전남",
        "보성군": "전남",
        "보성": "전남",
        "화순군": "전남",
        "화순": "전남",
        "장흥군": "전남",
        "장흥": "전남",
        "강진군": "전남",
        "강진": "전남",
        "해남군": "전남",
        "해남": "전남",
        "영암군": "전남",
        "영암": "전남",
        "무안군": "전남",
        "무안": "전남",
        "함평군": "전남",
        "함평": "전남",
        "영광군": "전남",
        "영광": "전남",
        "장성군": "전남",
        "장성": "전남",
        "완도군": "전남",
        "완도": "전남",
        "진도군": "전남",
        "진도": "전남",
        "신안군": "전남",
        "신안": "전남",

        # 경북
        "포항시": "경북",
        "포항": "경북",
        "경주시": "경북",
        "경주": "경북",
        "김천시": "경북",
        "김천": "경북",
        "안동시": "경북",
        "안동": "경북",
        "구미시": "경북",
        "구미": "경북",
        "영주시": "경북",
        "영주": "경북",
        "영천시": "경북",
        "영천": "경북",
        "상주시": "경북",
        "상주": "경북",
        "문경시": "경북",
        "문경": "경북",
        "경산시": "경북",
        "경산": "경북",
        "의성군": "경북",
        "의성": "경북",
        "청송군": "경북",
        "청송": "경북",
        "영양군": "경북",
        "영양": "경북",
        "영덕군": "경북",
        "영덕": "경북",
        "청도군": "경북",
        "청도": "경북",
        "고령군": "경북",
        "고령": "경북",
        "성주군": "경북",
        "성주": "경북",
        "칠곡군": "경북",
        "칠곡": "경북",
        "예천군": "경북",
        "예천": "경북",
        "봉화군": "경북",
        "봉화": "경북",
        "울진군": "경북",
        "울진": "경북",
        "울릉군": "경북",
        "울릉": "경북",
        "경북지역본부": "경북",
        "경북본부": "경북",

        # 경남
        "창원시": "경남",
        "창원": "경남",
        "진주시": "경남",
        "진주": "경남",
        "통영시": "경남",
        "통영": "경남",
        "사천시": "경남",
        "사천": "경남",
        "김해시": "경남",
        "김해": "경남",
        "밀양시": "경남",
        "밀양": "경남",
        "거제시": "경남",
        "거제": "경남",
        "양산시": "경남",
        "양산": "경남",
        "의령군": "경남",
        "의령": "경남",
        "함안군": "경남",
        "함안": "경남",
        "창녕군": "경남",
        "창녕": "경남",
        "경남고성": "경남",
        "남해군": "경남",
        "남해": "경남",
        "하동군": "경남",
        "하동": "경남",
        "산청군": "경남",
        "산청": "경남",
        "함양군": "경남",
        "함양": "경남",
        "거창군": "경남",
        "거창": "경남",
        "합천군": "경남",
        "합천": "경남",
        "경남지역본부": "경남",
        "경남본부": "경남",

        # 부산
        "부산진구": "부산",
        "해운대구": "부산",
        "사하구": "부산",
        "사상구": "부산",
        "금정구": "부산",
        "연제구": "부산",
        "수영구": "부산",
        "기장군": "부산",

        # 대구
        "달성군": "대구",
        "달성": "대구",
        "군위군": "대구",
        "군위": "대구",
        "수성구": "대구",
        "달서구": "대구",

        # 인천
        "강화군": "인천",
        "강화": "인천",
        "옹진군": "인천",
        "옹진": "인천",
        "미추홀구": "인천",
        "연수구": "인천",
        "남동구": "인천",
        "부평구": "인천",
        "계양구": "인천",

        # 울산
        "울주군": "울산",
        "울주": "울산",

        # 제주
        "제주시": "제주",
        "서귀포시": "제주",
        "제주지역본부": "제주",
        "제주본부": "제주",
    }

    for keyword, region in sorted(city_region_map.items(), key=lambda x: len(x[0]), reverse=True):
        if keyword in text or keyword in compact_text:
            return region

    return "기타"


def classify_region_from_row(row):
    """
    나라장터: 수요기관 + 공고기관 + 공고명을 함께 보고 지역 판정
    """
    demand_agency = str(row.get("dminsttNm", ""))
    notice_agency = str(row.get("ntceInsttNm", ""))
    bid_name = str(row.get("bidNtceNm", ""))

    combined_text = f"{demand_agency} {notice_agency} {bid_name}"
    return classify_region_by_text(combined_text)


def get_region_source(row):
    """
    지역 판정 근거 확인용 컬럼
    """
    demand_agency = str(row.get("dminsttNm", ""))
    notice_agency = str(row.get("ntceInsttNm", ""))
    bid_name = str(row.get("bidNtceNm", ""))

    return f"수요기관:{demand_agency} / 공고기관:{notice_agency} / 공고명:{bid_name}"


def safe_to_numeric(series):
    return (
        series.astype(str)
        .str.replace(",", "", regex=False)
        .str.replace(" ", "", regex=False)
        .replace("", "0")
        .pipe(pd.to_numeric, errors="coerce")
        .fillna(0)
    )


def parse_date_safe(value):
    """
    여러 형태의 날짜 문자열을 안전하게 datetime으로 변환
    """
    if pd.isna(value):
        return pd.NaT

    value = str(value).strip()

    if value == "":
        return pd.NaT

    if len(value) == 8 and value.isdigit():
        return pd.to_datetime(value, format="%Y%m%d", errors="coerce")

    if len(value) == 12 and value.isdigit():
        return pd.to_datetime(value, format="%Y%m%d%H%M", errors="coerce")

    return pd.to_datetime(value, errors="coerce")


def first_existing_col(df, candidates):
    for col in candidates:
        if col in df.columns:
            return col

    return None


def get_series_or_default(df, candidates, default=""):
    col = first_existing_col(df, candidates)

    if col is None:
        return pd.Series([default] * len(df), index=df.index)

    return df[col]


def add_nara_sales_urgency(df):
    """
    나라장터 입찰마감일 기준 영업 긴급도 생성
    """
    today = pd.Timestamp.today().normalize()

    if "bidNtceDate" in df.columns:
        df["공고일_변환"] = df["bidNtceDate"].apply(parse_date_safe)
    else:
        df["공고일_변환"] = pd.NaT

    if "bidClseDate" in df.columns:
        df["입찰마감일_변환"] = df["bidClseDate"].apply(parse_date_safe)
    else:
        df["입찰마감일_변환"] = pd.NaT

    if "opengDate" in df.columns:
        df["개찰일_변환"] = df["opengDate"].apply(parse_date_safe)
    else:
        df["개찰일_변환"] = pd.NaT

    if "준공예정일" not in df.columns:
        df["준공예정일"] = "공고문 확인 필요"

    df["입찰마감까지_남은일수"] = (
        df["입찰마감일_변환"] - today
    ).dt.days

    def classify_urgency(days):
        if pd.isna(days):
            return "일정 미확인"
        elif days < 0:
            return "마감"
        elif days <= 2:
            return "매우 긴급"
        elif days <= 5:
            return "긴급"
        elif days <= 10:
            return "주의"
        else:
            return "여유"

    df["영업긴급도"] = df["입찰마감까지_남은일수"].apply(classify_urgency)

    return df


def add_cals_sales_urgency(df):
    """
    CALS 준공예정일 기준 영업 긴급도 생성
    """
    today = pd.Timestamp.today().normalize()

    if "ccwXpcDt" in df.columns:
        df["준공예정일_변환"] = df["ccwXpcDt"].apply(parse_date_safe)
    else:
        df["준공예정일_변환"] = pd.NaT

    df["준공예정일까지_남은일수"] = (
        df["준공예정일_변환"] - today
    ).dt.days

    def classify_cals_urgency(days):
        if pd.isna(days):
            return "일정 미확인"
        elif days < 0:
            return "기간 경과 확인"
        elif days <= 30:
            return "매우 긴급"
        elif days <= 90:
            return "긴급"
        elif days <= 180:
            return "주의"
        else:
            return "여유"

    df["영업긴급도"] = df["준공예정일까지_남은일수"].apply(classify_cals_urgency)

    return df


def classify_price_band(value):
    value = float(value or 0)

    if value >= 10_000_000_000:
        return "100억 이상"
    if value >= 5_000_000_000:
        return "50억~100억"
    if value >= 1_000_000_000:
        return "10억~50억"
    if value >= 300_000_000:
        return "3억~10억"
    if value > 0:
        return "3억 미만"
    return "가격 미확인"


def add_nara_sales_columns(df):
    """
    나라장터 원천 데이터에 영업 우선순위 판단용 컬럼을 추가합니다.
    """
    if df.empty:
        return df

    urgency_score = {
        "매우 긴급": 40,
        "긴급": 30,
        "주의": 20,
        "여유": 10,
        "마감": 0,
        "일정 미확인": 5,
    }

    price_score = pd.cut(
        df["추정가격_숫자"].fillna(0),
        bins=[-1, 0, 300_000_000, 1_000_000_000, 5_000_000_000, 10_000_000_000, float("inf")],
        labels=[0, 5, 10, 20, 30, 40],
    ).astype(int)

    keyword_score = df["매칭키워드"].apply(lambda value: 20 if str(value).strip() else 0)

    df["추정가격구간"] = df["추정가격_숫자"].apply(classify_price_band)
    df["영업우선순위점수"] = (
        df["영업긴급도"].map(urgency_score).fillna(0).astype(int)
        + price_score
        + keyword_score
    )

    def classify_priority(score):
        if score >= 80:
            return "A. 즉시 확인"
        if score >= 55:
            return "B. 우선 검토"
        if score >= 30:
            return "C. 모니터링"
        return "D. 참고"

    def recommend_action(row):
        urgency = row.get("영업긴급도", "")
        keywords = str(row.get("매칭키워드", "")).strip()
        price_band = row.get("추정가격구간", "가격 미확인")

        if urgency in ["매우 긴급", "긴급"] and keywords:
            return "공고문·규격·납기 즉시 확인"
        if keywords and price_band in ["100억 이상", "50억~100억", "10억~50억"]:
            return "수요기관 담당/현장 규모 우선 확인"
        if keywords:
            return "철근 관련성 확인 후 추적"
        return "참고 데이터"

    df["영업우선순위"] = df["영업우선순위점수"].apply(classify_priority)
    df["권장영업액션"] = df.apply(recommend_action, axis=1)

    return df


def add_cals_sales_columns(df):
    """
    CALS 진행공사에 후속 수요 관점의 영업 판단 컬럼을 추가합니다.
    """
    if df.empty:
        return df

    field_weight = {
        "도로": "토목/도로 물량 가능성 확인",
        "하천": "호안·교량·수문 구조물 확인",
        "건축": "골조·증축 여부 우선 확인",
        "철도": "토목 구조물·역사 공사 확인",
        "항만": "항만 구조물 후속 물량 확인",
        "공항": "부대 토목·건축 물량 확인",
    }

    def classify_follow_up(days):
        if pd.isna(days):
            return "일정 확인 필요"
        if days < 0:
            return "준공 경과 여부 확인"
        if days <= 90:
            return "준공 임박/후속 발주 확인"
        if days <= 180:
            return "중기 모니터링"
        return "장기 모니터링"

    def recommend_action(row):
        field = row.get("bzarNm", "")
        base = field_weight.get(field, "공종/현장 규모 확인")
        stage = row.get("후속수요관점", "")
        return f"{stage} - {base}"

    df["후속수요관점"] = df["준공예정일까지_남은일수"].apply(classify_follow_up)
    df["권장영업액션"] = df.apply(recommend_action, axis=1)

    return df


# =========================================================
# 건축HUB 공통/수집/캐시 함수
# =========================================================

def archhub_region_from_sigungu(sigungu_cd):
    sigungu_cd = str(sigungu_cd).strip()

    if len(sigungu_cd) >= 2:
        return ARCHHUB_SIDO_PREFIX_MAP.get(sigungu_cd[:2], "기타")

    return "기타"


def archhub_extract_items(data):
    response = data.get("response", {}) if isinstance(data, dict) else {}
    body = response.get("body", {}) if isinstance(response, dict) else {}
    items = body.get("items", []) if isinstance(body, dict) else []

    if isinstance(items, dict) and "item" in items:
        items = items["item"]

    if isinstance(items, dict):
        items = [items]

    if items is None:
        items = []

    return items


def archhub_get_total_count(data):
    response = data.get("response", {}) if isinstance(data, dict) else {}
    body = response.get("body", {}) if isinstance(response, dict) else {}
    total_count = body.get("totalCount", 0) if isinstance(body, dict) else 0

    try:
        return int(total_count)
    except Exception:
        return 0


def archhub_get_result_info(data):
    response = data.get("response", {}) if isinstance(data, dict) else {}
    header = response.get("header", {}) if isinstance(response, dict) else {}

    result_code = header.get("resultCode", "") if isinstance(header, dict) else ""
    result_msg = header.get("resultMsg", "") if isinstance(header, dict) else ""

    return result_code, result_msg


def archhub_is_response_ok(data):
    if data is None:
        return False

    result_code, _ = archhub_get_result_info(data)

    if result_code in ["00", "NORMAL_SERVICE"]:
        return True

    response = data.get("response", {}) if isinstance(data, dict) else {}
    body = response.get("body", {}) if isinstance(response, dict) else {}

    if isinstance(body, dict):
        if "items" in body:
            return True

        if "totalCount" in body:
            try:
                int(body.get("totalCount", 0))
                return True
            except Exception:
                pass

    return False


def archhub_response_json(response):
    try:
        return response.json()
    except Exception:
        print("[건축HUB 오류] JSON 변환 실패")
        print("응답 앞부분:")
        print((response.text or "")[:500])
        return None


def archhub_request_page(
    sigungu_cd,
    start_date,
    end_date,
    page_no=1,
    num_of_rows=100,
    bjdong_cd="",
):
    params = {
        "serviceKey": require_api_key(ARCHHUB_API_KEY, "ARCHHUB_API_KEY 또는 NARA_API_KEY"),
        "sigunguCd": sigungu_cd,
        "startDate": start_date,
        "endDate": end_date,
        "numOfRows": str(num_of_rows),
        "pageNo": str(page_no),
        "_type": "json",
    }

    if str(bjdong_cd).strip() != "":
        params["bjdongCd"] = str(bjdong_cd).strip()

    response = requests.get(
        ARCHHUB_BASIS_URL,
        params=params,
        timeout=30,
        verify=False,
    )

    bjdong_label = params.get("bjdongCd", "전체")

    print(
        f"[건축HUB] sigunguCd={sigungu_cd or '전체'} "
        f"bjdongCd={bjdong_label} page={page_no} status={response.status_code}"
    )

    data = archhub_response_json(response)

    if data is None:
        return None, response.status_code

    result_code, result_msg = archhub_get_result_info(data)
    print(f"[건축HUB] resultCode={result_code}, resultMsg={result_msg}")

    return data, response.status_code


def collect_archhub_basis_by_sigungu(
    sigungu_cd,
    start_date,
    end_date,
    bjdong_cd="",
    max_pages=1,
    num_of_rows=100,
    sleep_sec=0.2,
    fallback_bjdong_cd="10100",
):
    """
    건축HUB 기본개요를 시군구 기준으로 수집합니다.

    현재 getApBasisOulnInfo는 일부 조건에서 bjdongCd 없는 시군구 전체 조회가
    HTTP 200이지만 resultCode/items가 비어 있는 응답을 줄 수 있습니다.
    이 경우 1차 MVP에서는 fallback 법정동 코드로 재시도합니다.
    """
    data, status_code = archhub_request_page(
        sigungu_cd=sigungu_cd,
        bjdong_cd=bjdong_cd,
        start_date=start_date,
        end_date=end_date,
        page_no=1,
        num_of_rows=num_of_rows,
    )

    used_bjdong_cd = bjdong_cd

    if data is None or status_code >= 500 or not archhub_is_response_ok(data):
        if str(bjdong_cd).strip() == "" and fallback_bjdong_cd:
            print(
                f"[건축HUB 재시도] 시군구 전체 조회가 유효하지 않음 → "
                f"fallback bjdongCd={fallback_bjdong_cd}"
            )
            data, status_code = archhub_request_page(
                sigungu_cd=sigungu_cd,
                bjdong_cd=fallback_bjdong_cd,
                start_date=start_date,
                end_date=end_date,
                page_no=1,
                num_of_rows=num_of_rows,
            )
            used_bjdong_cd = fallback_bjdong_cd

    if data is None:
        return pd.DataFrame()

    if not archhub_is_response_ok(data):
        _, result_msg = archhub_get_result_info(data)
        print(f"[건축HUB 경고] API 정상 응답이 아닙니다: {result_msg}")
        return pd.DataFrame()

    total_count = archhub_get_total_count(data)
    first_items = archhub_extract_items(data)

    if total_count == 0 and len(first_items) == 0:
        if str(bjdong_cd).strip() == "" and fallback_bjdong_cd and used_bjdong_cd != fallback_bjdong_cd:
            print(
                f"[건축HUB 재시도] 시군구 전체 조회 결과 0건 → "
                f"fallback bjdongCd={fallback_bjdong_cd}"
            )
            data, status_code = archhub_request_page(
                sigungu_cd=sigungu_cd,
                bjdong_cd=fallback_bjdong_cd,
                start_date=start_date,
                end_date=end_date,
                page_no=1,
                num_of_rows=num_of_rows,
            )
            used_bjdong_cd = fallback_bjdong_cd

            if data is None or not archhub_is_response_ok(data):
                return pd.DataFrame()

            total_count = archhub_get_total_count(data)
            first_items = archhub_extract_items(data)

    total_pages = math.ceil(total_count / num_of_rows) if total_count else 1

    if max_pages is not None and max_pages > 0:
        total_pages = min(total_pages, max_pages)

    print(f"[건축HUB] totalCount={total_count}, 수집 예정 페이지={total_pages}")

    all_items = []
    all_items.extend(first_items)

    for page in range(2, total_pages + 1):
        page_data, _ = archhub_request_page(
            sigungu_cd=sigungu_cd,
            bjdong_cd=used_bjdong_cd,
            start_date=start_date,
            end_date=end_date,
            page_no=page,
            num_of_rows=num_of_rows,
        )

        if page_data is None:
            print(f"[건축HUB 경고] {page}페이지 수집 실패")
            continue

        all_items.extend(archhub_extract_items(page_data))
        time.sleep(sleep_sec)

    df = pd.DataFrame(all_items)

    if not df.empty:
        df["조회_sigunguCd"] = sigungu_cd
        df["조회_bjdongCd"] = used_bjdong_cd if str(used_bjdong_cd).strip() != "" else "전체"
        df["조회_region"] = archhub_region_from_sigungu(sigungu_cd)
        df["조회_totalCount"] = total_count

    return df


def collect_archhub_basis_sample(
    start_date,
    end_date,
    bjdong_cd="",
    max_pages=1,
    num_of_rows=100,
):
    frames = []

    for label, sigungu_cd in ARCHHUB_SAMPLE_SIGUNGU_CODES.items():
        print("\n" + "=" * 70)
        print(f"건축HUB 샘플 지역 수집: {label} / {sigungu_cd}")
        print("=" * 70)

        try:
            df_part = collect_archhub_basis_by_sigungu(
                sigungu_cd=sigungu_cd,
                bjdong_cd=bjdong_cd,
                start_date=start_date,
                end_date=end_date,
                max_pages=max_pages,
                num_of_rows=num_of_rows,
            )

            if not df_part.empty:
                df_part["샘플지역명"] = label
                frames.append(df_part)

        except Exception as exc:
            print(f"[건축HUB 경고] {label} 수집 실패: {exc}")

        time.sleep(0.2)

    if not frames:
        return pd.DataFrame()

    return pd.concat(frames, ignore_index=True)


def normalize_archhub_basis(df):
    if df.empty:
        return df

    out = pd.DataFrame()

    out["원천관리번호"] = get_series_or_default(
        df,
        ["mgmPmsrgstPk", "관리번호", "mgmNo"],
        "",
    )

    out["원천_sigunguCd"] = get_series_or_default(
        df,
        ["sigunguCd", "조회_sigunguCd"],
        "",
    )

    out["원천_bjdongCd"] = get_series_or_default(
        df,
        ["bjdongCd", "조회_bjdongCd"],
        "",
    )

    out["region"] = out["원천_sigunguCd"].apply(archhub_region_from_sigungu)

    out["대지위치"] = get_series_or_default(
        df,
        ["platPlc", "newPlatPlc", "대지위치", "siteLoc", "addr"],
        "",
    )

    out["건물명"] = get_series_or_default(df, ["bldNm", "건물명"], "")
    out["허가구분"] = get_series_or_default(df, ["archGbCdNm", "허가구분", "archGbNm"], "")
    out["주용도"] = get_series_or_default(df, ["mainPurpsCdNm", "mainPurpsNm", "mainPurps", "주용도"], "")
    out["지목"] = get_series_or_default(df, ["jimokCdNm", "지목"], "")
    out["지역지구"] = get_series_or_default(df, ["jiyukCdNm", "지역지구"], "")

    out["허가일"] = get_series_or_default(df, ["archPmsDay", "pmsDay", "허가일", "permitDay"], "").apply(parse_date_safe)
    out["착공예정일"] = get_series_or_default(df, ["stcnsSchedDay", "착공예정일"], "").apply(parse_date_safe)
    out["착공지연일"] = get_series_or_default(df, ["stcnsDelayDay", "착공지연일"], "").apply(parse_date_safe)
    out["실제착공일"] = get_series_or_default(df, ["realStcnsDay", "stcnsDay", "착공일", "startDay"], "").apply(parse_date_safe)
    out["사용승인일"] = get_series_or_default(df, ["useAprDay", "사용승인일", "useApprovalDay"], "").apply(parse_date_safe)
    out["데이터생성일"] = get_series_or_default(df, ["crtnDay", "생성일"], "").apply(parse_date_safe)

    out["대지면적"] = safe_to_numeric(get_series_or_default(df, ["platArea", "대지면적"], "0"))
    out["건축면적"] = safe_to_numeric(get_series_or_default(df, ["archArea", "건축면적"], "0"))
    out["연면적"] = safe_to_numeric(get_series_or_default(df, ["totArea", "totFlrArea", "연면적", "totalArea"], "0"))
    out["용적률산정연면적"] = safe_to_numeric(get_series_or_default(df, ["vlRatEstmTotArea", "용적률산정연면적"], "0"))
    out["건폐율"] = safe_to_numeric(get_series_or_default(df, ["bcRat", "건폐율"], "0"))
    out["용적률"] = safe_to_numeric(get_series_or_default(df, ["vlRat", "용적률"], "0"))
    out["주건축물수"] = safe_to_numeric(get_series_or_default(df, ["mainBldCnt", "주건축물수"], "0"))
    out["부속건축물동수"] = safe_to_numeric(get_series_or_default(df, ["atchBldDongCnt", "부속건축물동수"], "0"))
    out["세대수"] = safe_to_numeric(get_series_or_default(df, ["hhldCnt", "세대수"], "0"))
    out["호수"] = safe_to_numeric(get_series_or_default(df, ["hoCnt", "호수"], "0"))
    out["가구수"] = safe_to_numeric(get_series_or_default(df, ["fmlyCnt", "가구수"], "0"))
    out["주차대수"] = safe_to_numeric(get_series_or_default(df, ["totPkngCnt", "주차대수"], "0"))
    out["조회_bjdongCd"] = get_series_or_default(df, ["조회_bjdongCd"], "")
    out["조회_totalCount"] = safe_to_numeric(get_series_or_default(df, ["조회_totalCount"], "0"))

    if "샘플지역명" in df.columns:
        out["샘플지역명"] = df["샘플지역명"]

    return out


def calc_archhub_purpose_weight(value):
    text = str(value)

    if "공동주택" in text or "아파트" in text:
        return 1.30
    if "업무" in text or "오피스텔" in text:
        return 1.10
    if "교육" in text or "연구" in text:
        return 1.00
    if "공장" in text:
        return 0.90
    if "의료" in text or "판매" in text:
        return 0.90
    if "근린생활" in text:
        return 0.80
    if "창고" in text:
        return 0.60
    if "단독주택" in text:
        return 0.50
    if "가설" in text:
        return 0.20

    return 0.80


def classify_archhub_start_status(row):
    if pd.notna(row.get("사용승인일")):
        return "사용승인"
    if pd.notna(row.get("실제착공일")):
        return "실제착공"
    if pd.notna(row.get("착공예정일")):
        return "착공예정"
    if pd.notna(row.get("허가일")):
        return "허가"
    return "일정미확인"


def calc_archhub_status_weight(status):
    if status == "실제착공":
        return 1.20
    if status == "착공예정":
        return 1.10
    if status == "허가":
        return 1.00
    if status == "사용승인":
        return 0.30
    return 0.70


def add_archhub_demand_score(df):
    if df.empty:
        return df

    df = df.copy()

    df["착공상태"] = df.apply(classify_archhub_start_status, axis=1)
    df["용도가중치"] = df["주용도"].apply(calc_archhub_purpose_weight)
    df["착공상태가중치"] = df["착공상태"].apply(calc_archhub_status_weight)

    df["공동주택여부"] = df["주용도"].astype(str).str.contains("공동주택|아파트", regex=True, na=False)
    df["업무시설여부"] = df["주용도"].astype(str).str.contains("업무|오피스텔", regex=True, na=False)
    df["공장여부"] = df["주용도"].astype(str).str.contains("공장", regex=True, na=False)
    df["착공연계여부"] = df["착공상태"].isin(["실제착공", "착공예정"])
    df["대형허가여부"] = df["연면적"] >= 10000

    df["건축허가_선행수요지수"] = (
        (df["연면적"] / 10000)
        * df["용도가중치"]
        * df["착공상태가중치"]
    )

    return df


def make_archhub_region_summary(df):
    if df.empty:
        return pd.DataFrame()

    summary = (
        df.groupby("region")
        .agg(
            건축허가건수=("원천관리번호", "count"),
            건축허가_연면적합계=("연면적", "sum"),
            공동주택_연면적=("연면적", lambda s: s[df.loc[s.index, "공동주택여부"]].sum()),
            업무시설_연면적=("연면적", lambda s: s[df.loc[s.index, "업무시설여부"]].sum()),
            공장_연면적=("연면적", lambda s: s[df.loc[s.index, "공장여부"]].sum()),
            착공연계_연면적=("연면적", lambda s: s[df.loc[s.index, "착공연계여부"]].sum()),
            실제착공_연면적=("연면적", lambda s: s[df.loc[s.index, "착공상태"].eq("실제착공")].sum()),
            착공예정_연면적=("연면적", lambda s: s[df.loc[s.index, "착공상태"].eq("착공예정")].sum()),
            사용승인_연면적=("연면적", lambda s: s[df.loc[s.index, "착공상태"].eq("사용승인")].sum()),
            대형허가건수=("대형허가여부", "sum"),
            세대수합계=("세대수", "sum"),
            건축허가_선행수요지수=("건축허가_선행수요지수", "sum"),
        )
        .reset_index()
        .sort_values(by="건축허가_선행수요지수", ascending=False)
    )

    summary.insert(0, "선행수요순위", range(1, len(summary) + 1))

    return summary


def archhub_cache_files():
    return {
        "normalized": CACHE_PATH / "archhub_basis_normalized_latest.csv",
        "summary": CACHE_PATH / "archhub_basis_region_summary_latest.csv",
        "meta": CACHE_PATH / "archhub_cache_meta.json",
    }


def is_archhub_cache_fresh():
    files = archhub_cache_files()

    if ARCHHUB_FORCE_REFRESH:
        return False

    if not files["normalized"].exists() or not files["summary"].exists():
        return False

    modified_time = datetime.fromtimestamp(files["summary"].stat().st_mtime)
    age_hours = (datetime.now() - modified_time).total_seconds() / 3600

    return age_hours <= ARCHHUB_CACHE_HOURS


def load_archhub_cache():
    files = archhub_cache_files()

    if not is_archhub_cache_fresh():
        return pd.DataFrame(), pd.DataFrame()

    try:
        archhub_df = pd.read_csv(files["normalized"])
        archhub_region_summary = pd.read_csv(files["summary"])
        print(f"[건축HUB] 캐시 사용: {files['summary']}")
        return archhub_df, archhub_region_summary
    except Exception as exc:
        print(f"[건축HUB 경고] 캐시 읽기 실패: {exc}")
        return pd.DataFrame(), pd.DataFrame()


def save_archhub_cache(raw_df, archhub_df, archhub_region_summary):
    now = datetime.now().strftime("%Y%m%d_%H%M%S")

    timestamp_paths = {
        "raw": CACHE_PATH / f"archhub_basis_raw_{now}.csv",
        "normalized": CACHE_PATH / f"archhub_basis_normalized_{now}.csv",
        "summary": CACHE_PATH / f"archhub_basis_region_summary_{now}.csv",
    }

    latest_paths = {
        "raw": CACHE_PATH / "archhub_basis_raw_latest.csv",
        "normalized": CACHE_PATH / "archhub_basis_normalized_latest.csv",
        "summary": CACHE_PATH / "archhub_basis_region_summary_latest.csv",
    }

    raw_df.to_csv(timestamp_paths["raw"], index=False, encoding="utf-8-sig")
    archhub_df.to_csv(timestamp_paths["normalized"], index=False, encoding="utf-8-sig")
    archhub_region_summary.to_csv(timestamp_paths["summary"], index=False, encoding="utf-8-sig")

    raw_df.to_csv(latest_paths["raw"], index=False, encoding="utf-8-sig")
    archhub_df.to_csv(latest_paths["normalized"], index=False, encoding="utf-8-sig")
    archhub_region_summary.to_csv(latest_paths["summary"], index=False, encoding="utf-8-sig")

    meta = {
        "created_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "mode": ARCHHUB_MODE,
        "lookback_days": ARCHHUB_LOOKBACK_DAYS,
        "max_pages": ARCHHUB_MAX_PAGES,
        "raw_rows": int(len(raw_df)),
        "normalized_rows": int(len(archhub_df)),
        "summary_rows": int(len(archhub_region_summary)),
    }

    (CACHE_PATH / "archhub_cache_meta.json").write_text(
        json.dumps(meta, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )


def collect_archhub():
    print("\n")
    print("=" * 70)
    print("3. 건축HUB 건축인허가 기본개요 수집 시작")
    print("=" * 70)

    if not ARCHHUB_ENABLED:
        print("건축HUB 수집 비활성화: ARCHHUB_ENABLED=false")
        return pd.DataFrame(), pd.DataFrame()

    cached_df, cached_summary = load_archhub_cache()

    if not cached_df.empty or not cached_summary.empty:
        return cached_df, cached_summary

    end_date = datetime.now().strftime("%Y%m%d")
    start_date = (datetime.now() - timedelta(days=ARCHHUB_LOOKBACK_DAYS)).strftime("%Y%m%d")
    max_pages = None if ARCHHUB_MAX_PAGES == 0 else ARCHHUB_MAX_PAGES

    print(f"건축HUB 조회기간: {start_date} ~ {end_date}")
    print(f"건축HUB 모드: {ARCHHUB_MODE}")
    print(f"건축HUB 시군구별 최대 페이지: {'전체' if max_pages is None else max_pages}")

    if ARCHHUB_MODE == "single":
        raw_df = collect_archhub_basis_by_sigungu(
            sigungu_cd=ARCHHUB_SIGUNGU_CD,
            bjdong_cd=ARCHHUB_BJDONG_CD,
            start_date=start_date,
            end_date=end_date,
            max_pages=max_pages,
            num_of_rows=ARCHHUB_NUM_ROWS,
        )
    else:
        raw_df = collect_archhub_basis_sample(
            bjdong_cd=ARCHHUB_BJDONG_CD,
            start_date=start_date,
            end_date=end_date,
            max_pages=max_pages,
            num_of_rows=ARCHHUB_NUM_ROWS,
        )

    if raw_df.empty:
        print("건축HUB 수집 데이터가 없습니다.")
        return pd.DataFrame(), pd.DataFrame()

    archhub_df = normalize_archhub_basis(raw_df)
    archhub_df = add_archhub_demand_score(archhub_df)
    archhub_region_summary = make_archhub_region_summary(archhub_df)

    save_archhub_cache(raw_df, archhub_df, archhub_region_summary)

    print("건축HUB 원천 데이터 건수:", len(raw_df))
    print("건축HUB 정규화 데이터 건수:", len(archhub_df))
    print("건축HUB 지역 요약 건수:", len(archhub_region_summary))

    return archhub_df, archhub_region_summary


# =========================================================
# 1. 나라장터 공사입찰 수집
# =========================================================

def collect_narajangteo():
    nara_api_key = require_api_key(NARA_API_KEY, "NARA_API_KEY")

    print("\n")
    print("=" * 70)
    print("1. 나라장터 공사입찰 수집 시작")
    print("=" * 70)

    base_url = (
        "http://apis.data.go.kr/1230000/ad/"
        "BidPublicInfoService/getBidPblancListInfoCnstwk"
    )

    num_of_rows = 100
    all_items = []

    end_date = datetime.now()
    start_date = end_date - timedelta(days=NARA_LOOKBACK_DAYS)

    first_params = {
        "serviceKey": nara_api_key,
        "pageNo": "1",
        "numOfRows": str(num_of_rows),
        "inqryDiv": "1",
        "inqryBgnDt": start_date.strftime("%Y%m%d0000"),
        "inqryEndDt": end_date.strftime("%Y%m%d2359"),
        "type": "json"
    }

    response = requests.get(
        base_url,
        params=first_params,
        timeout=30,
        verify=False
    )

    print("나라장터 상태코드:", response.status_code)

    data = response.json()

    result_code = data["response"]["header"].get("resultCode")
    result_msg = data["response"]["header"].get("resultMsg")

    print("나라장터 응답코드:", result_code)
    print("나라장터 응답메시지:", result_msg)

    if result_code != "00":
        raise ValueError(f"나라장터 API 오류: {result_msg}")

    body = data["response"]["body"]
    total_count = int(body.get("totalCount", 0))
    total_pages = math.ceil(total_count / num_of_rows)

    print("나라장터 전체 입찰 건수:", total_count)
    print("나라장터 전체 페이지 수:", total_pages)

    for page in range(1, total_pages + 1):
        params = {
            "serviceKey": nara_api_key,
            "pageNo": str(page),
            "numOfRows": str(num_of_rows),
            "inqryDiv": "1",
            "inqryBgnDt": start_date.strftime("%Y%m%d0000"),
            "inqryEndDt": end_date.strftime("%Y%m%d2359"),
            "type": "json"
        }

        response = requests.get(
            base_url,
            params=params,
            timeout=30,
            verify=False
        )

        data = response.json()
        items = data["response"]["body"].get("items", [])

        if isinstance(items, dict):
            items = [items]

        all_items.extend(items)

        print(f"나라장터 {page}/{total_pages} 페이지 수집 완료 - 누적 {len(all_items)}건")
        time.sleep(0.2)

    df = pd.DataFrame(all_items)

    if df.empty:
        print("나라장터 조회 데이터가 없습니다.")
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

    selected_cols = [
        "bidNtceNo",
        "bidNtceOrd",
        "bidNtceNm",
        "dminsttNm",
        "ntceInsttNm",
        "bidNtceDate",
        "bidNtceBgn",
        "bidClseDate",
        "bidClseTm",
        "opengDate",
        "opengTm",
        "presmptPrce",
        "rgstTyNm",
        "ntceKindNm",
        "cntrctCnclsMthdNm",
        "bidMethdNm"
    ]

    selected_cols = [col for col in selected_cols if col in df.columns]
    df = df[selected_cols]
    df = df.fillna("")

    before_filter_count = len(df)

    if "ntceKindNm" in df.columns:
        df = df[df["ntceKindNm"] != "취소공고"]

    if "cntrctCnclsMthdNm" in df.columns:
        df = df[~df["cntrctCnclsMthdNm"].str.contains("수의계약", na=False)]

    after_filter_count = len(df)
    excluded_count = before_filter_count - after_filter_count

    print("나라장터 필터 전 건수:", before_filter_count)
    print("나라장터 제외 건수:", excluded_count)
    print("나라장터 필터 후 건수:", after_filter_count)

    steel_keywords = [
        "철근",
        "콘크리트",
        "보수공사",
        "토목공사",
        "포장공사",
        "조성사업",
        "재해복구사업",
        "정비사업",
        "복구",
        "시설보수",
        "개보수",
        "수해복구",
        "교량",
        "도로확장",
        "배수로",
        "우수관",
        "L형측구",
        "옹벽",
        "하수관로",
        "도로정비",
        "침수복구",
        "급경사지",
        "재포장",
        "도로개설",
        "하천정비"
    ]

    def find_matched_keywords(title):
        matched = []

        for keyword in steel_keywords:
            if keyword in str(title):
                matched.append(keyword)

        return ", ".join(matched)

    df["매칭키워드"] = df["bidNtceNm"].apply(find_matched_keywords)

    df = add_nara_sales_urgency(df)

    if "presmptPrce" in df.columns:
        df["추정가격_숫자"] = safe_to_numeric(df["presmptPrce"])
    else:
        df["추정가격_숫자"] = 0

    df = add_nara_sales_columns(df)

    # 지역 분류 추가: 수요기관 + 공고기관 + 공고명 전체를 기준으로 판정
    df["region"] = df.apply(classify_region_from_row, axis=1)
    df["지역판정기준"] = df.apply(get_region_source, axis=1)

    df_steel = df[df["매칭키워드"] != ""].copy()

    df = df.drop_duplicates()
    df_steel = df_steel.drop_duplicates()

    keyword_rows = []

    for keyword in steel_keywords:
        count = df["bidNtceNm"].str.contains(keyword, na=False).sum()

        keyword_rows.append({
            "키워드": keyword,
            "입찰건수": count
        })

    keyword_summary = pd.DataFrame(keyword_rows)
    keyword_summary = keyword_summary.sort_values(by="입찰건수", ascending=False)
    keyword_summary.insert(0, "순위", range(1, len(keyword_summary) + 1))

    if "dminsttNm" in df_steel.columns:
        agency_summary = (
            df_steel.groupby("dminsttNm")
            .size()
            .reset_index(name="철근관련입찰건수")
            .sort_values(by="철근관련입찰건수", ascending=False)
        )
        agency_summary.insert(0, "순위", range(1, len(agency_summary) + 1))
    else:
        agency_summary = pd.DataFrame(columns=["순위", "dminsttNm", "철근관련입찰건수"])

    nara_region_summary = (
        df_steel.groupby("region")
        .agg(
            나라장터_철근관련입찰건수=("bidNtceNm", "count"),
            나라장터_추정가격합계=("추정가격_숫자", "sum")
        )
        .reset_index()
        .sort_values(by="나라장터_철근관련입찰건수", ascending=False)
    )

    nara_region_summary.insert(
        0,
        "순위",
        range(1, len(nara_region_summary) + 1)
    )

    return df, df_steel, keyword_summary, agency_summary, nara_region_summary


# =========================================================
# 2. 건설CALS 진행공사 수집
# =========================================================

def collect_calspia():
    cals_api_key = require_api_key(CALS_API_KEY, "CALS_API_KEY")

    print("\n")
    print("=" * 70)
    print("2. 건설CALS 진행공사 수집 시작")
    print("=" * 70)

    # CALS는 외부 클라우드/Streamlit Cloud 환경에서 간헐적으로 연결 지연이 발생할 수 있어
    # 단일 요청 실패로 전체 리포트가 중단되지 않도록 재시도 로직을 둡니다.
    base_urls = [
        "https://www.calspia.go.kr/io/openapi/cm/selectIoCmConstructionList.do",
        # HTTPS 연결이 일시적으로 불안정할 때를 대비한 보조 후보입니다.
        # 서버 정책에 따라 HTTP가 차단될 수 있으며, 이 경우 자동으로 다음 재시도에서 무시됩니다.
        "http://www.calspia.go.kr/io/openapi/cm/selectIoCmConstructionList.do",
    ]

    num_of_rows = int(os.getenv("CALS_NUM_ROWS", "100"))
    connect_timeout = int(os.getenv("CALS_CONNECT_TIMEOUT", "15"))
    read_timeout = int(os.getenv("CALS_READ_TIMEOUT", "60"))
    max_retries = int(os.getenv("CALS_MAX_RETRIES", "3"))
    backoff_sec = float(os.getenv("CALS_BACKOFF_SEC", "2.0"))
    max_pages_limit = int(os.getenv("CALS_MAX_PAGES", "0"))  # 0이면 전체 페이지

    all_items = []

    session = requests.Session()
    session.headers.update({
        "User-Agent": (
            "Mozilla/5.0 (compatible; steel-demand-report/1.0; "
            "+https://github.com/Ronaldooh123/steel-demand-report)"
        ),
        "Accept": "application/json, text/plain, */*",
        "Connection": "close",
    })

    def mask_sensitive(text_value):
        text_value = str(text_value)
        text_value = re.sub(r"(serviceKey=)[^&\s]+", r"\1***", text_value)
        text_value = re.sub(r"(serviceKey%3D)[^&\s]+", r"\1***", text_value)
        return text_value

    def request_cals_page(page_no):
        params = {
            "serviceKey": cals_api_key,
            "type": "json",
            "pageNo": page_no,
            "numOfRows": num_of_rows,
        }

        last_error = None

        for base_url in base_urls:
            for attempt in range(1, max_retries + 1):
                try:
                    response = session.get(
                        base_url,
                        params=params,
                        timeout=(connect_timeout, read_timeout),
                        verify=False,
                    )

                    print(
                        f"건설CALS page={page_no} "
                        f"status={response.status_code} "
                        f"attempt={attempt}/{max_retries}"
                    )

                    # 서버가 일시적으로 불안정한 경우 재시도합니다.
                    if response.status_code in [429, 500, 502, 503, 504]:
                        last_error = RuntimeError(f"HTTP {response.status_code}")
                        time.sleep(backoff_sec * attempt)
                        continue

                    response.raise_for_status()

                    try:
                        data = response.json()
                    except Exception as exc:
                        preview = mask_sensitive((response.text or "")[:300])
                        raise RuntimeError(
                            f"CALS JSON 변환 실패: {type(exc).__name__}, 응답앞부분={preview}"
                        ) from exc

                    if not isinstance(data, dict) or "response" not in data:
                        preview = mask_sensitive(str(data)[:300])
                        raise RuntimeError(f"CALS 응답 구조 이상: {preview}")

                    return data

                except requests.exceptions.ConnectTimeout as exc:
                    last_error = exc
                    print(
                        f"[건설CALS 재시도] 연결 시간초과 "
                        f"page={page_no}, attempt={attempt}/{max_retries}"
                    )
                    time.sleep(backoff_sec * attempt)

                except requests.exceptions.ReadTimeout as exc:
                    last_error = exc
                    print(
                        f"[건설CALS 재시도] 응답 읽기 시간초과 "
                        f"page={page_no}, attempt={attempt}/{max_retries}"
                    )
                    time.sleep(backoff_sec * attempt)

                except requests.exceptions.RequestException as exc:
                    last_error = exc
                    print(
                        f"[건설CALS 재시도] 요청 오류 "
                        f"page={page_no}, type={type(exc).__name__}, attempt={attempt}/{max_retries}"
                    )
                    time.sleep(backoff_sec * attempt)

                except Exception as exc:
                    # JSON 파싱/응답 구조 문제는 네트워크 문제가 아닐 수 있으므로 즉시 중단합니다.
                    raise RuntimeError(mask_sensitive(str(exc))) from exc

        error_type = type(last_error).__name__ if last_error is not None else "UnknownError"
        error_msg = mask_sensitive(str(last_error))[:500] if last_error is not None else "원인 미상"
        raise RuntimeError(
            f"건설CALS API 연결 실패 page={page_no}, "
            f"error_type={error_type}, error={error_msg}"
        )

    first_data = request_cals_page(1)

    body = first_data.get("response", {}).get("body", {})
    total_count = int(body.get("totalCount", 0))
    total_pages = math.ceil(total_count / num_of_rows) if total_count else 1

    if max_pages_limit > 0:
        total_pages = min(total_pages, max_pages_limit)

    print("건설CALS 전체 공사 건수:", total_count)
    print("건설CALS 전체 페이지 수:", total_pages)

    first_items = body.get("items", [])

    if isinstance(first_items, dict):
        first_items = [first_items]

    if first_items:
        all_items.extend(first_items)

    print(f"건설CALS 1/{total_pages} 페이지 수집 완료 - 누적 {len(all_items)}건")

    for page in range(2, total_pages + 1):
        data = request_cals_page(page)
        items = data.get("response", {}).get("body", {}).get("items", [])

        if isinstance(items, dict):
            items = [items]

        all_items.extend(items)

        print(f"건설CALS {page}/{total_pages} 페이지 수집 완료 - 누적 {len(all_items)}건")
        time.sleep(0.2)

    df = pd.DataFrame(all_items)

    if df.empty:
        print("건설CALS 조회 데이터가 없습니다.")
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

    selected_cols = [
        "cwkNm",
        "ornm",
        "rutNm",
        "pdznNm",
        "bzKdNm",
        "stwrDt",
        "ccwDt",
        "ccwXpcDt",
        "ccwYnNm",
        "bzarNm",
        "cwkSctnNm"
    ]

    selected_cols = [col for col in selected_cols if col in df.columns]
    df = df[selected_cols]
    df = df.fillna("")

    before_filter_count = len(df)

    if "ccwYnNm" in df.columns:
        df = df[df["ccwYnNm"] == "진행공사"]

    after_filter_count = len(df)
    excluded_count = before_filter_count - after_filter_count

    print("건설CALS 필터 전 건수:", before_filter_count)
    print("건설CALS 제외 건수:", excluded_count)
    print("건설CALS 진행공사 건수:", after_filter_count)

    df = add_cals_sales_urgency(df)
    df = add_cals_sales_columns(df)

    df["region"] = df["ornm"].apply(classify_region_by_text)

    df = df.drop_duplicates()

    region_summary = (
        df.groupby("region")
        .size()
        .reset_index(name="건설CALS_진행공사건수")
        .sort_values(by="건설CALS_진행공사건수", ascending=False)
    )

    region_summary.insert(0, "순위", range(1, len(region_summary) + 1))

    region_field_summary = (
        df.groupby(["region", "bzarNm"])
        .size()
        .reset_index(name="공사건수")
        .sort_values(by=["region", "공사건수"], ascending=[True, False])
    )

    agency_summary = (
        df.groupby("ornm")
        .size()
        .reset_index(name="진행공사건수")
        .sort_values(by="진행공사건수", ascending=False)
    )

    agency_summary.insert(0, "순위", range(1, len(agency_summary) + 1))

    steel_weight = {
        "도로": 1.0,
        "하천": 0.8,
        "건축": 1.2,
        "철도": 1.3,
        "항만": 1.1,
        "공항": 1.1,
        "기타": 0.5,
        "": 0.5
    }

    region_field_summary["철근가중치"] = (
        region_field_summary["bzarNm"]
        .map(steel_weight)
        .fillna(0.6)
    )

    region_field_summary["건설CALS_철근수요지수"] = (
        region_field_summary["공사건수"]
        * region_field_summary["철근가중치"]
    )

    cals_demand_summary = (
        region_field_summary.groupby("region")["건설CALS_철근수요지수"]
        .sum()
        .reset_index()
        .sort_values(by="건설CALS_철근수요지수", ascending=False)
    )

    cals_demand_summary.insert(
        0,
        "수요순위",
        range(1, len(cals_demand_summary) + 1)
    )

    return df, region_summary, region_field_summary, agency_summary, cals_demand_summary


# =========================================================
# 3. 통합 수요지수 생성
# =========================================================

def make_integrated_summary(nara_region_summary, cals_demand_summary, archhub_region_summary=None):
    print("\n")
    print("=" * 70)
    print("4. 통합 수요지수 생성")
    print("=" * 70)

    if archhub_region_summary is None:
        archhub_region_summary = pd.DataFrame()

    if nara_region_summary.empty and cals_demand_summary.empty and archhub_region_summary.empty:
        return pd.DataFrame()

    nara = nara_region_summary.copy()
    cals = cals_demand_summary.copy()
    archhub = archhub_region_summary.copy()

    if not nara.empty:
        nara = nara[[
            "region",
            "나라장터_철근관련입찰건수",
            "나라장터_추정가격합계"
        ]]
    else:
        nara = pd.DataFrame(columns=[
            "region",
            "나라장터_철근관련입찰건수",
            "나라장터_추정가격합계"
        ])

    if not cals.empty:
        cals = cals[[
            "region",
            "건설CALS_철근수요지수"
        ]]
    else:
        cals = pd.DataFrame(columns=[
            "region",
            "건설CALS_철근수요지수"
        ])

    if not archhub.empty:
        if "지역" in archhub.columns and "region" not in archhub.columns:
            archhub = archhub.rename(columns={"지역": "region"})

        archhub_cols = [
            "region",
            "건축허가건수",
            "건축허가_연면적합계",
            "공동주택_연면적",
            "업무시설_연면적",
            "착공연계_연면적",
            "대형허가건수",
            "세대수합계",
            "건축허가_선행수요지수",
        ]
        archhub = archhub[[col for col in archhub_cols if col in archhub.columns]]
    else:
        archhub = pd.DataFrame(columns=[
            "region",
            "건축허가건수",
            "건축허가_연면적합계",
            "공동주택_연면적",
            "업무시설_연면적",
            "착공연계_연면적",
            "대형허가건수",
            "세대수합계",
            "건축허가_선행수요지수",
        ])

    integrated = pd.merge(cals, nara, on="region", how="outer")
    integrated = pd.merge(integrated, archhub, on="region", how="outer")
    integrated = integrated.fillna(0)

    for col in [
        "건설CALS_철근수요지수",
        "나라장터_철근관련입찰건수",
        "나라장터_추정가격합계",
        "건축허가건수",
        "건축허가_연면적합계",
        "공동주택_연면적",
        "업무시설_연면적",
        "착공연계_연면적",
        "대형허가건수",
        "세대수합계",
        "건축허가_선행수요지수",
    ]:
        if col not in integrated.columns:
            integrated[col] = 0

    integrated["통합수요지수"] = (
        integrated["건설CALS_철근수요지수"] * 1.0
        + integrated["나라장터_철근관련입찰건수"] * 1.5
        + integrated["건축허가_선행수요지수"] * 0.8
    )

    integrated["수요등급"] = pd.cut(
        integrated["통합수요지수"],
        bins=[-1, 5, 15, 30, 999999],
        labels=["낮음", "보통", "높음", "매우 높음"]
    )

    integrated["신규입찰비중"] = integrated.apply(
        lambda row: (
            row["나라장터_철근관련입찰건수"] * 1.5 / row["통합수요지수"]
            if row["통합수요지수"] > 0
            else 0
        ),
        axis=1
    )

    integrated["건축허가비중"] = integrated.apply(
        lambda row: (
            row["건축허가_선행수요지수"] * 0.8 / row["통합수요지수"]
            if row["통합수요지수"] > 0
            else 0
        ),
        axis=1
    )

    def recommend_integrated_action(row):
        grade = str(row.get("수요등급", ""))
        nara_count = float(row.get("나라장터_철근관련입찰건수", 0))
        cals_index = float(row.get("건설CALS_철근수요지수", 0))
        archhub_index = float(row.get("건축허가_선행수요지수", 0))
        arch_area = float(row.get("건축허가_연면적합계", 0))

        if grade == "매우 높음" and nara_count >= 5:
            return "최우선 공략: 신규 입찰·진행공사·건축허가 동시 확인"
        if grade in ["매우 높음", "높음"] and archhub_index >= 20 and arch_area >= 100000:
            return "건축허가 선행수요 확인: 민간·건축 현장 모니터링"
        if grade in ["매우 높음", "높음"] and cals_index >= 15:
            return "진행공사 기반 수요처/현장 추적"
        if grade in ["매우 높음", "높음"]:
            return "입찰 공고 중심 단기 모니터링"
        if nara_count > 0:
            return "철근 관련 공고 발생 시 추적"
        if archhub_index > 0:
            return "건축허가 추세 정기 모니터링"
        return "정기 모니터링"

    integrated["권장영업액션"] = integrated.apply(recommend_integrated_action, axis=1)

    integrated = integrated.sort_values(
        by="통합수요지수",
        ascending=False
    )

    integrated.insert(
        0,
        "통합순위",
        range(1, len(integrated) + 1)
    )

    return integrated


# =========================================================
# 4. 엑셀 리포트 품질 개선 함수
# =========================================================

def rename_columns_korean(df):
    """
    컬럼명을 사용자가 보기 쉬운 한글명으로 변경
    """
    column_map = {
        # 나라장터
        "bidNtceNo": "입찰공고번호",
        "bidNtceOrd": "공고차수",
        "bidNtceNm": "입찰공고명",
        "dminsttNm": "수요기관",
        "ntceInsttNm": "공고기관",
        "bidNtceDate": "공고일자",
        "bidNtceBgn": "입찰개시일시",
        "bidClseDate": "입찰마감일자",
        "bidClseTm": "입찰마감시각",
        "opengDate": "개찰일자",
        "opengTm": "개찰시각",
        "presmptPrce": "추정가격",
        "추정가격_숫자": "추정가격_숫자",
        "추정가격구간": "추정가격구간",
        "rgstTyNm": "등록유형",
        "ntceKindNm": "공고종류",
        "cntrctCnclsMthdNm": "계약체결방법",
        "bidMethdNm": "입찰방식",
        "매칭키워드": "매칭키워드",
        "영업우선순위점수": "영업우선순위점수",
        "영업우선순위": "영업우선순위",
        "권장영업액션": "권장영업액션",
        "region": "지역",
        "지역판정기준": "지역판정기준",

        # CALS
        "cwkNm": "공사명",
        "ornm": "발주기관",
        "rutNm": "노선/수계",
        "pdznNm": "권역",
        "bzKdNm": "사업종류",
        "stwrDt": "착공일",
        "ccwDt": "준공일",
        "ccwXpcDt": "준공예정일",
        "ccwYnNm": "진행상태",
        "bzarNm": "분야",
        "cwkSctnNm": "공사구간",

        # 집계
        "순위": "순위",
        "수요순위": "수요순위",
        "통합순위": "통합순위",
        "키워드": "키워드",
        "입찰건수": "입찰건수",
        "철근관련입찰건수": "철근 관련 입찰건수",
        "진행공사건수": "진행공사 건수",
        "공사건수": "공사건수",
        "건설CALS_진행공사건수": "CALS 진행공사 건수",
        "건설CALS_철근수요지수": "CALS 철근수요지수",
        "나라장터_철근관련입찰건수": "나라장터 철근관련 입찰건수",
        "나라장터_추정가격합계": "나라장터 추정가격 합계",
        "철근가중치": "철근 가중치",
        "철근수요지수": "철근 수요지수",
        "통합수요지수": "통합 수요지수",
        "수요등급": "수요등급",
        "신규입찰비중": "신규입찰비중",
        "건축허가비중": "건축허가비중",
        "후속수요관점": "후속수요관점",

        # 건축HUB
        "원천관리번호": "원천관리번호",
        "원천_sigunguCd": "시군구코드",
        "원천_bjdongCd": "법정동코드",
        "대지위치": "대지위치",
        "건물명": "건물명",
        "허가구분": "허가구분",
        "주용도": "주용도",
        "지목": "지목",
        "지역지구": "지역지구",
        "허가일": "허가일",
        "착공예정일": "착공예정일",
        "착공지연일": "착공지연일",
        "실제착공일": "실제착공일",
        "사용승인일": "사용승인일",
        "데이터생성일": "데이터생성일",
        "대지면적": "대지면적",
        "건축면적": "건축면적",
        "연면적": "연면적",
        "용적률산정연면적": "용적률산정연면적",
        "건폐율": "건폐율",
        "용적률": "용적률",
        "주건축물수": "주건축물수",
        "부속건축물동수": "부속건축물동수",
        "세대수": "세대수",
        "호수": "호수",
        "가구수": "가구수",
        "주차대수": "주차대수",
        "착공상태": "착공상태",
        "용도가중치": "용도가중치",
        "착공상태가중치": "착공상태가중치",
        "공동주택여부": "공동주택여부",
        "업무시설여부": "업무시설여부",
        "공장여부": "공장여부",
        "착공연계여부": "착공연계여부",
        "대형허가여부": "대형허가여부",
        "건축허가_선행수요지수": "건축허가 선행수요지수",
        "선행수요순위": "선행수요순위",
        "건축허가건수": "건축허가 건수",
        "건축허가_연면적합계": "건축허가 연면적 합계",
        "공동주택_연면적": "공동주택 연면적",
        "업무시설_연면적": "업무시설 연면적",
        "공장_연면적": "공장 연면적",
        "착공연계_연면적": "착공연계 연면적",
        "실제착공_연면적": "실제착공 연면적",
        "착공예정_연면적": "착공예정 연면적",
        "사용승인_연면적": "사용승인 연면적",
        "대형허가건수": "대형허가 건수",
        "세대수합계": "세대수 합계",

        # 날짜/긴급도
        "공고일_변환": "공고일",
        "입찰마감일_변환": "입찰마감일",
        "개찰일_변환": "개찰일",
        "입찰마감까지_남은일수": "입찰마감까지 남은일수",
        "준공예정일": "준공예정일",
        "준공예정일_변환": "준공예정일",
        "준공예정일까지_남은일수": "준공예정일까지 남은일수",
        "영업긴급도": "영업긴급도",
    }

    return df.rename(columns=column_map)


def safe_len(df):
    if df is None:
        return 0
    return len(df)


def get_top_region(integrated_summary):
    if integrated_summary is None or integrated_summary.empty:
        return "데이터 없음"

    if "region" in integrated_summary.columns:
        return str(integrated_summary.iloc[0]["region"])

    if "지역" in integrated_summary.columns:
        return str(integrated_summary.iloc[0]["지역"])

    return "데이터 없음"


def get_top_keyword(keyword_summary):
    if keyword_summary is None or keyword_summary.empty:
        return "데이터 없음"

    if "키워드" in keyword_summary.columns:
        return str(keyword_summary.iloc[0]["키워드"])

    return "데이터 없음"


def format_worksheet(writer, sheet_name, df, group_type):
    """
    각 데이터 시트 공통 서식 적용
    """
    workbook = writer.book
    worksheet = writer.sheets[sheet_name]

    if group_type == "nara":
        tab_color = "#FFD966"
        header_color = "#BF9000"
    elif group_type == "cals":
        tab_color = "#A9D18E"
        header_color = "#548235"
    elif group_type == "integrated":
        tab_color = "#9DC3E6"
        header_color = "#1F4E78"
    elif group_type == "archhub":
        tab_color = "#D9EAD3"
        header_color = "#38761D"
    else:
        tab_color = "#D9EAD3"
        header_color = "#666666"

    worksheet.set_tab_color(tab_color)

    header_format = workbook.add_format({
        "bold": True,
        "font_color": "white",
        "bg_color": header_color,
        "border": 1,
        "align": "center",
        "valign": "vcenter"
    })

    number_format = workbook.add_format({
        "border": 1,
        "num_format": "#,##0",
        "valign": "vcenter"
    })

    decimal_format = workbook.add_format({
        "border": 1,
        "num_format": "#,##0.0",
        "valign": "vcenter"
    })

    percent_format = workbook.add_format({
        "border": 1,
        "num_format": "0.0%",
        "valign": "vcenter"
    })

    if df is None or df.empty:
        return

    row_count = len(df)
    col_count = len(df.columns)

    for col_num, value in enumerate(df.columns):
        worksheet.write(0, col_num, value, header_format)

    worksheet.autofilter(0, 0, row_count, col_count - 1)
    worksheet.freeze_panes(1, 0)

    for idx, col in enumerate(df.columns):
        series = df.iloc[:, idx].apply(lambda x: "" if pd.isna(x) else str(x))

        max_len = max(
            series.map(len).max() if not series.empty else 0,
            len(str(col))
        )

        width = min(max(max_len + 2, 10), 38)

        if "공고명" in col or "공사명" in col or "대지위치" in col:
            width = 45
        elif "기관" in col:
            width = 28
        elif "권장영업액션" in col or "판정기준" in col:
            width = 55
        elif "가격" in col or "합계" in col:
            width = 18
        elif "일자" in col or "일시" in col:
            width = 18

        worksheet.set_column(idx, idx, width)

    worksheet.set_row(0, 24)

    for idx, col in enumerate(df.columns):
        if any(keyword in col for keyword in ["건수", "가격", "합계", "남은일수"]):
            worksheet.set_column(idx, idx, 18, number_format)

        if any(keyword in col for keyword in ["지수", "가중치"]):
            worksheet.set_column(idx, idx, 16, decimal_format)

        if "비중" in col:
            worksheet.set_column(idx, idx, 14, percent_format)

    if "영업우선순위" in df.columns:
        priority_col_idx = list(df.columns).index("영업우선순위")
        priority_col_letter = xl_col_to_name(priority_col_idx)

        worksheet.conditional_format(
            f"{priority_col_letter}2:{priority_col_letter}{row_count + 1}",
            {
                "type": "text",
                "criteria": "containing",
                "value": "A. 즉시 확인",
                "format": workbook.add_format({
                    "bg_color": "#F4CCCC",
                    "font_color": "#990000",
                    "bold": True
                })
            }
        )

        worksheet.conditional_format(
            f"{priority_col_letter}2:{priority_col_letter}{row_count + 1}",
            {
                "type": "text",
                "criteria": "containing",
                "value": "B. 우선 검토",
                "format": workbook.add_format({
                    "bg_color": "#FCE5CD",
                    "font_color": "#B45F06",
                    "bold": True
                })
            }
        )

    if "수요등급" in df.columns:
        grade_col_idx = list(df.columns).index("수요등급")
        grade_col_letter = xl_col_to_name(grade_col_idx)

        worksheet.conditional_format(
            f"{grade_col_letter}2:{grade_col_letter}{row_count + 1}",
            {
                "type": "text",
                "criteria": "containing",
                "value": "매우 높음",
                "format": workbook.add_format({
                    "bg_color": "#F4CCCC",
                    "font_color": "#990000",
                    "bold": True
                })
            }
        )

        worksheet.conditional_format(
            f"{grade_col_letter}2:{grade_col_letter}{row_count + 1}",
            {
                "type": "text",
                "criteria": "containing",
                "value": "높음",
                "format": workbook.add_format({
                    "bg_color": "#FCE5CD",
                    "font_color": "#B45F06",
                    "bold": True
                })
            }
        )

    if "영업긴급도" in df.columns:
        urgency_col_idx = list(df.columns).index("영업긴급도")
        urgency_col_letter = xl_col_to_name(urgency_col_idx)

        worksheet.conditional_format(
            f"{urgency_col_letter}2:{urgency_col_letter}{row_count + 1}",
            {
                "type": "text",
                "criteria": "containing",
                "value": "매우 긴급",
                "format": workbook.add_format({
                    "bg_color": "#F4CCCC",
                    "font_color": "#990000",
                    "bold": True
                })
            }
        )

        worksheet.conditional_format(
            f"{urgency_col_letter}2:{urgency_col_letter}{row_count + 1}",
            {
                "type": "text",
                "criteria": "containing",
                "value": "긴급",
                "format": workbook.add_format({
                    "bg_color": "#FCE5CD",
                    "font_color": "#B45F06",
                    "bold": True
                })
            }
        )

        worksheet.conditional_format(
            f"{urgency_col_letter}2:{urgency_col_letter}{row_count + 1}",
            {
                "type": "text",
                "criteria": "containing",
                "value": "주의",
                "format": workbook.add_format({
                    "bg_color": "#FFF2CC",
                    "font_color": "#7F6000"
                })
            }
        )


def write_summary_sheet(
    writer,
    nara_df,
    nara_steel_df,
    keyword_summary,
    nara_region_summary,
    cals_df,
    cals_demand_summary,
    integrated_summary,
    archhub_df=None,
    archhub_region_summary=None
):
    workbook = writer.book
    worksheet = workbook.add_worksheet("요약")
    writer.sheets["요약"] = worksheet

    worksheet.set_tab_color("#4472C4")

    title_format = workbook.add_format({
        "bold": True,
        "font_size": 20,
        "font_color": "white",
        "bg_color": "#1F4E78",
        "align": "center",
        "valign": "vcenter"
    })

    section_format = workbook.add_format({
        "bold": True,
        "font_size": 13,
        "font_color": "white",
        "bg_color": "#5B9BD5",
        "border": 1,
        "align": "center",
        "valign": "vcenter"
    })

    label_format = workbook.add_format({
        "bold": True,
        "bg_color": "#D9EAF7",
        "border": 1,
        "align": "left",
        "valign": "vcenter"
    })

    value_format = workbook.add_format({
        "border": 1,
        "align": "right",
        "valign": "vcenter",
        "num_format": "#,##0"
    })

    text_format = workbook.add_format({
        "border": 1,
        "valign": "top",
        "text_wrap": True
    })

    percent_format = workbook.add_format({
        "border": 1,
        "align": "right",
        "valign": "vcenter",
        "num_format": "0.0%"
    })

    worksheet.merge_range("A1:H2", "철근 수요 예측 통합 리포트", title_format)
    worksheet.set_row(0, 30)
    worksheet.set_row(1, 30)

    report_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    worksheet.merge_range("A4:C4", "1. 핵심 지표", section_format)

    urgent_count = 0
    priority_count = 0
    estimated_total = 0

    if nara_steel_df is not None and not nara_steel_df.empty:
        if "영업긴급도" in nara_steel_df.columns:
            urgent_count = nara_steel_df["영업긴급도"].isin(["매우 긴급", "긴급"]).sum()
        if "영업우선순위" in nara_steel_df.columns:
            priority_count = nara_steel_df["영업우선순위"].isin(["A. 즉시 확인", "B. 우선 검토"]).sum()
        if "추정가격_숫자" in nara_steel_df.columns:
            estimated_total = nara_steel_df["추정가격_숫자"].sum()

    archhub_count = safe_len(archhub_df)
    archhub_top_region = "데이터 없음"
    archhub_area_total = 0

    if archhub_region_summary is not None and not archhub_region_summary.empty:
        archhub_top_region = get_top_region(archhub_region_summary)
        if "건축허가_연면적합계" in archhub_region_summary.columns:
            archhub_area_total = archhub_region_summary["건축허가_연면적합계"].sum()

    summary_rows = [
        ["리포트 생성일", report_time],
        ["나라장터 전체 입찰공고 수", safe_len(nara_df)],
        ["나라장터 철근 관련 입찰 수", safe_len(nara_steel_df)],
        ["긴급/매우 긴급 입찰 수", int(urgent_count)],
        ["A/B 우선 검토 입찰 수", int(priority_count)],
        ["철근 관련 추정가격 합계", int(estimated_total)],
        ["건설CALS 진행공사 수", safe_len(cals_df)],
        ["건축HUB 인허가 수", int(archhub_count)],
        ["건축HUB 연면적 합계", int(archhub_area_total)],
        ["통합 수요 1위 지역", get_top_region(integrated_summary)],
        ["건축허가 선행수요 1위", archhub_top_region],
        ["최다 매칭 키워드", get_top_keyword(keyword_summary)],
    ]

    start_row = 5

    for i, row in enumerate(summary_rows):
        worksheet.write(start_row + i, 0, row[0], label_format)
        worksheet.write(start_row + i, 1, row[1], value_format if isinstance(row[1], Number) else text_format)

    worksheet.set_column("A:A", 28)
    worksheet.set_column("B:B", 24)
    worksheet.set_column("C:C", 3)

    worksheet.merge_range("D4:H4", "2. 자동 해석 요약", section_format)

    top_region = get_top_region(integrated_summary)
    top_keyword = get_top_keyword(keyword_summary)

    nara_steel_count = safe_len(nara_steel_df)
    cals_count = safe_len(cals_df)

    summary_text = (
        f"이번 리포트는 나라장터 공사입찰, 건설CALS 진행공사, 건축HUB 건축인허가 데이터를 통합해 "
        f"지역별 철근 수요 가능성과 영업 우선순위를 산정한 자료입니다.\n\n"
        f"나라장터는 단기 입찰 수요, CALS는 현재 진행공사 수요, 건축HUB는 향후 착공 가능성이 있는 "
        f"건축 부문 선행수요로 해석합니다.\n\n"
        f"현재 철근 관련 입찰은 {nara_steel_count:,}건이며, 긴급 또는 매우 긴급 공고는 "
        f"{int(urgent_count):,}건입니다. A/B 우선 검토 대상은 {int(priority_count):,}건입니다.\n\n"
        f"건축HUB 인허가 데이터는 {archhub_count:,}건이며, 표본 기준 건축허가 선행수요 1위 지역은 "
        f"'{archhub_top_region}'입니다.\n\n"
        f"통합 수요지수 기준 최우선 지역은 '{top_region}'입니다. 해당 지역은 입찰 공고, 진행공사, 건축허가 흐름을 "
        f"같이 확인하고 공고문·규격·납기·현장 규모를 우선 점검하는 것이 좋습니다."
    )

    worksheet.merge_range("D5:H12", summary_text, text_format)

    worksheet.merge_range("A13:H13", "3. 통합 철근 수요 TOP 5 지역", section_format)

    top5 = integrated_summary.head(5).copy() if integrated_summary is not None and not integrated_summary.empty else pd.DataFrame()

    if not top5.empty:
        top5_cols = [
            "통합순위",
            "region",
            "통합수요지수",
            "수요등급",
            "나라장터_철근관련입찰건수",
            "건설CALS_철근수요지수",
            "건축허가_선행수요지수",
            "건축허가_연면적합계",
            "건축허가비중",
            "권장영업액션",
        ]
        top5 = top5[[col for col in top5_cols if col in top5.columns]]
        top5 = rename_columns_korean(top5)
        headers = list(top5.columns)

        for col_idx, header in enumerate(headers[:8]):
            worksheet.write(14, col_idx, header, label_format)

        for row_idx, (_, row) in enumerate(top5.iterrows(), start=15):
            for col_idx, header in enumerate(headers[:8]):
                cell_format = value_format if isinstance(row[header], Number) else text_format
                if header in ["신규입찰비중", "건축허가비중"]:
                    cell_format = percent_format
                worksheet.write(
                    row_idx,
                    col_idx,
                    row[header],
                    cell_format
                )

    worksheet.merge_range("A23:H23", "4. 나라장터 철근 관련 키워드 TOP 10", section_format)

    top_keywords = keyword_summary.head(10).copy() if keyword_summary is not None and not keyword_summary.empty else pd.DataFrame()

    if not top_keywords.empty:
        top_keywords = rename_columns_korean(top_keywords)
        headers = list(top_keywords.columns)

        for col_idx, header in enumerate(headers):
            worksheet.write(24, col_idx, header, label_format)

        for row_idx, (_, row) in enumerate(top_keywords.iterrows(), start=25):
            for col_idx, header in enumerate(headers):
                worksheet.write(
                    row_idx,
                    col_idx,
                    row[header],
                    value_format if isinstance(row[header], Number) else text_format
                )

    worksheet.merge_range(
        "D25:H30",
        "※ 해석 기준\n"
        "- 나라장터: 향후 발주·입찰 가능성을 보여주는 선행지표\n"
        "- 건설CALS: 현재 진행공사 기반의 현재 수요지표\n"
        "- 통합 수요지수: 현재 공사 수요와 신규 입찰 흐름을 함께 반영한 참고 지표\n"
        "- 영업우선순위: 입찰마감 긴급도, 추정가격, 철근 관련 키워드를 합산한 내부 검토용 점수\n"
        "- 본 지수는 영업 우선순위 판단을 위한 보조 지표이며, 실제 철근 소요량 산정에는 공사 규모·구조·공종별 보정이 추가로 필요합니다.",
        text_format
    )

    worksheet.set_column("A:H", 18)
    worksheet.set_column("D:H", 20)
    worksheet.set_column("H:H", 42)
    worksheet.freeze_panes(4, 0)


def export_integrated_excel_report(
    excel_file,
    nara_df,
    nara_steel_df,
    keyword_summary,
    nara_agency_summary,
    nara_region_summary,
    cals_df,
    cals_region_summary,
    cals_region_field_summary,
    cals_agency_summary,
    cals_demand_summary,
    integrated_summary,
    archhub_df=None,
    archhub_region_summary=None
):
    nara_df_kr = rename_columns_korean(nara_df.copy())
    nara_steel_df_kr = rename_columns_korean(nara_steel_df.copy())
    keyword_summary_kr = rename_columns_korean(keyword_summary.copy())
    nara_agency_summary_kr = rename_columns_korean(nara_agency_summary.copy())
    nara_region_summary_kr = rename_columns_korean(nara_region_summary.copy())

    cals_df_kr = rename_columns_korean(cals_df.copy())
    cals_region_summary_kr = rename_columns_korean(cals_region_summary.copy())
    cals_region_field_summary_kr = rename_columns_korean(cals_region_field_summary.copy())
    cals_agency_summary_kr = rename_columns_korean(cals_agency_summary.copy())
    cals_demand_summary_kr = rename_columns_korean(cals_demand_summary.copy())

    if archhub_df is None:
        archhub_df = pd.DataFrame()
    if archhub_region_summary is None:
        archhub_region_summary = pd.DataFrame()

    archhub_df_kr = rename_columns_korean(archhub_df.copy())
    archhub_region_summary_kr = rename_columns_korean(archhub_region_summary.copy())

    integrated_summary_kr = rename_columns_korean(integrated_summary.copy())

    with pd.ExcelWriter(excel_file, engine="xlsxwriter") as writer:

        write_summary_sheet(
            writer,
            nara_df,
            nara_steel_df,
            keyword_summary,
            nara_region_summary,
            cals_df,
            cals_demand_summary,
            integrated_summary,
            archhub_df=archhub_df,
            archhub_region_summary=archhub_region_summary
        )

        integrated_summary_kr.to_excel(writer, sheet_name="통합_수요지수", index=False)
        archhub_region_summary_kr.to_excel(writer, sheet_name="건축HUB_선행수요", index=False)
        archhub_df_kr.to_excel(writer, sheet_name="건축HUB_원천", index=False)

        nara_steel_df_kr.to_excel(writer, sheet_name="나라장터_철근입찰", index=False)
        nara_region_summary_kr.to_excel(writer, sheet_name="나라장터_지역별", index=False)
        nara_agency_summary_kr.to_excel(writer, sheet_name="나라장터_기관별", index=False)
        keyword_summary_kr.to_excel(writer, sheet_name="나라장터_키워드", index=False)

        cals_demand_summary_kr.to_excel(writer, sheet_name="CALS_수요지수", index=False)
        cals_df_kr.to_excel(writer, sheet_name="CALS_진행공사", index=False)
        cals_region_summary_kr.to_excel(writer, sheet_name="CALS_지역별", index=False)
        cals_region_field_summary_kr.to_excel(writer, sheet_name="CALS_지역분야", index=False)
        cals_agency_summary_kr.to_excel(writer, sheet_name="CALS_기관별", index=False)

        nara_df_kr.to_excel(writer, sheet_name="나라장터_전체입찰", index=False)

        format_worksheet(writer, "통합_수요지수", integrated_summary_kr, "integrated")
        format_worksheet(writer, "건축HUB_선행수요", archhub_region_summary_kr, "archhub")
        format_worksheet(writer, "건축HUB_원천", archhub_df_kr, "archhub")
        format_worksheet(writer, "나라장터_전체입찰", nara_df_kr, "nara")
        format_worksheet(writer, "나라장터_철근입찰", nara_steel_df_kr, "nara")
        format_worksheet(writer, "나라장터_키워드", keyword_summary_kr, "nara")
        format_worksheet(writer, "나라장터_기관별", nara_agency_summary_kr, "nara")
        format_worksheet(writer, "나라장터_지역별", nara_region_summary_kr, "nara")

        format_worksheet(writer, "CALS_진행공사", cals_df_kr, "cals")
        format_worksheet(writer, "CALS_지역별", cals_region_summary_kr, "cals")
        format_worksheet(writer, "CALS_지역분야", cals_region_field_summary_kr, "cals")
        format_worksheet(writer, "CALS_기관별", cals_agency_summary_kr, "cals")
        format_worksheet(writer, "CALS_수요지수", cals_demand_summary_kr, "cals")


# =========================================================
# 5. 실행
# =========================================================

if __name__ == "__main__":

    nara_df, nara_steel_df, keyword_summary, nara_agency_summary, nara_region_summary = collect_narajangteo()

    try:
        cals_df, cals_region_summary, cals_region_field_summary, cals_agency_summary, cals_demand_summary = collect_calspia()
    except Exception as exc:
        print("[건설CALS 경고] 수집 중 오류가 발생하여 CALS 없이 리포트를 생성합니다.")
        print("[건설CALS 경고] Streamlit Cloud 또는 외부망에서 CALS 서버 접속이 지연/차단될 수 있습니다.")
        safe_error = re.sub(r"(serviceKey=)[^&\s]+", r"\1***", str(exc))
        print(f"[건설CALS 경고] 오류유형: {type(exc).__name__}")
        print(f"[건설CALS 경고] 오류요약: {safe_error[:500]}")

        cals_df = pd.DataFrame()
        cals_region_summary = pd.DataFrame(columns=[
            "순위",
            "region",
            "건설CALS_진행공사건수",
        ])
        cals_region_field_summary = pd.DataFrame(columns=[
            "region",
            "bzarNm",
            "공사건수",
            "철근가중치",
            "건설CALS_철근수요지수",
        ])
        cals_agency_summary = pd.DataFrame(columns=[
            "순위",
            "ornm",
            "진행공사건수",
        ])
        cals_demand_summary = pd.DataFrame(columns=[
            "수요순위",
            "region",
            "건설CALS_철근수요지수",
        ])

    try:
        archhub_df, archhub_region_summary = collect_archhub()
    except Exception as exc:
        print("[건축HUB 경고] 수집 중 오류가 발생하여 건축HUB 없이 리포트를 생성합니다.")
        print(exc)
        archhub_df = pd.DataFrame()
        archhub_region_summary = pd.DataFrame()

    integrated_summary = make_integrated_summary(
        nara_region_summary,
        cals_demand_summary,
        archhub_region_summary
    )

    excel_file = os.path.join(
        SAVE_PATH,
        f"integrated_steel_demand_report_{NOW}.xlsx"
    )

    export_integrated_excel_report(
        excel_file=excel_file,
        nara_df=nara_df,
        nara_steel_df=nara_steel_df,
        keyword_summary=keyword_summary,
        nara_agency_summary=nara_agency_summary,
        nara_region_summary=nara_region_summary,
        cals_df=cals_df,
        cals_region_summary=cals_region_summary,
        cals_region_field_summary=cals_region_field_summary,
        cals_agency_summary=cals_agency_summary,
        cals_demand_summary=cals_demand_summary,
        integrated_summary=integrated_summary,
        archhub_df=archhub_df,
        archhub_region_summary=archhub_region_summary
    )

    print("\n")
    print("=" * 70)
    print("통합 리포트 생성 완료")
    print("=" * 70)

    print("\n저장 위치:")
    print(excel_file)

    print("\n")
    print("=" * 70)
    print("통합 수요지수")
    print("=" * 70)

    if not integrated_summary.empty:
        print(integrated_summary.to_string(index=False))
    else:
        print("통합 수요지수 데이터가 없습니다.")
