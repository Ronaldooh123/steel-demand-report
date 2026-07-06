import os
import math
import time
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


BASE_DIR = Path(__file__).parent

SAVE_PATH = BASE_DIR / "data" / "processed"
SAVE_PATH.mkdir(parents=True, exist_ok=True)

NOW = datetime.now().strftime("%Y%m%d_%H%M%S")
NARA_LOOKBACK_DAYS = int(os.getenv("NARA_LOOKBACK_DAYS", "7"))


def require_api_key(value, key_name):
    if value:
        return value

    raise ValueError(f"{key_name}를 찾을 수 없습니다. .env, Streamlit Secrets, GitHub Secrets를 확인하세요.")


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

    base_url = "https://www.calspia.go.kr/io/openapi/cm/selectIoCmConstructionList.do"

    num_of_rows = 100
    all_items = []

    first_params = {
        "serviceKey": cals_api_key,
        "type": "json",
        "pageNo": 1,
        "numOfRows": num_of_rows
    }

    response = requests.get(
        base_url,
        params=first_params,
        timeout=30,
        verify=False
    )

    print("건설CALS 상태코드:", response.status_code)

    data = response.json()

    body = data["response"]["body"]
    total_count = int(body.get("totalCount", 0))
    total_pages = math.ceil(total_count / num_of_rows)

    print("건설CALS 전체 공사 건수:", total_count)
    print("건설CALS 전체 페이지 수:", total_pages)

    for page in range(1, total_pages + 1):
        params = {
            "serviceKey": cals_api_key,
            "type": "json",
            "pageNo": page,
            "numOfRows": num_of_rows
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

def make_integrated_summary(nara_region_summary, cals_demand_summary):
    print("\n")
    print("=" * 70)
    print("3. 통합 수요지수 생성")
    print("=" * 70)

    if nara_region_summary.empty and cals_demand_summary.empty:
        return pd.DataFrame()

    nara = nara_region_summary.copy()
    cals = cals_demand_summary.copy()

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

    integrated = pd.merge(
        cals,
        nara,
        on="region",
        how="outer"
    )

    integrated = integrated.fillna(0)

    integrated["통합수요지수"] = (
        integrated["건설CALS_철근수요지수"]
        + integrated["나라장터_철근관련입찰건수"] * 1.5
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

    def recommend_integrated_action(row):
        grade = str(row.get("수요등급", ""))
        nara_count = float(row.get("나라장터_철근관련입찰건수", 0))
        cals_index = float(row.get("건설CALS_철근수요지수", 0))

        if grade == "매우 높음" and nara_count >= 5:
            return "최우선 공략: 신규 입찰·진행공사 동시 확인"
        if grade in ["매우 높음", "높음"] and cals_index >= 15:
            return "진행공사 기반 수요처/현장 추적"
        if grade in ["매우 높음", "높음"]:
            return "입찰 공고 중심 단기 모니터링"
        if nara_count > 0:
            return "철근 관련 공고 발생 시 추적"
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
        "후속수요관점": "후속수요관점",

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

        if "공고명" in col or "공사명" in col:
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
    integrated_summary
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

    summary_rows = [
        ["리포트 생성일", report_time],
        ["나라장터 전체 입찰공고 수", safe_len(nara_df)],
        ["나라장터 철근 관련 입찰 수", safe_len(nara_steel_df)],
        ["긴급/매우 긴급 입찰 수", int(urgent_count)],
        ["A/B 우선 검토 입찰 수", int(priority_count)],
        ["철근 관련 추정가격 합계", int(estimated_total)],
        ["건설CALS 진행공사 수", safe_len(cals_df)],
        ["통합 수요 1위 지역", get_top_region(integrated_summary)],
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
        f"이번 리포트는 나라장터 공사입찰을 선행 수요, 건설CALS 진행공사를 현재 수요로 보고 "
        f"지역별 철근 수요 가능성과 영업 우선순위를 통합한 자료입니다.\n\n"
        f"현재 철근 관련 입찰은 {nara_steel_count:,}건이며, 이 중 긴급 또는 매우 긴급으로 분류된 공고는 "
        f"{int(urgent_count):,}건입니다. A/B 우선 검토 대상은 {int(priority_count):,}건입니다.\n\n"
        f"통합 수요지수 기준 최우선 지역은 '{top_region}'입니다. 해당 지역은 진행공사 규모와 신규 입찰 흐름을 함께 확인하고, "
        f"상위 수요기관·공고기관을 중심으로 공고문, 규격, 납기, 현장 규모를 우선 점검하는 것이 좋습니다.\n\n"
        f"최다 매칭 키워드는 '{top_keyword}'입니다. 키워드는 실제 철근 소요량 확정값이 아니라 탐색 신호이므로, "
        f"대형 공사·구조물 공사·토목 공사 여부를 공고문에서 추가 확인해야 합니다."
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
            "신규입찰비중",
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
                if header == "신규입찰비중":
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
    integrated_summary
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
            integrated_summary
        )

        integrated_summary_kr.to_excel(writer, sheet_name="통합_수요지수", index=False)

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

    cals_df, cals_region_summary, cals_region_field_summary, cals_agency_summary, cals_demand_summary = collect_calspia()

    integrated_summary = make_integrated_summary(
        nara_region_summary,
        cals_demand_summary
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
        integrated_summary=integrated_summary
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
