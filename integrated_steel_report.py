import os
import math
import time
import requests
import urllib3
import pandas as pd

from datetime import datetime, timedelta
from dotenv import load_dotenv

# =========================================================
# 기본 설정
# =========================================================

# =========================================================
# 기본 설정
# =========================================================

from pathlib import Path

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

if NARA_API_KEY is None:
    raise ValueError("NARA_API_KEY를 찾을 수 없습니다. .env 또는 Streamlit Secrets를 확인하세요.")

if CALS_API_KEY is None:
    raise ValueError("CALS_API_KEY를 찾을 수 없습니다. .env 또는 Streamlit Secrets를 확인하세요.")


# 현재 파일 기준 프로젝트 폴더
BASE_DIR = Path(__file__).parent

# 배포용/로컬 공통 저장 경로
SAVE_PATH = BASE_DIR / "data" / "processed"
SAVE_PATH.mkdir(parents=True, exist_ok=True)

NOW = datetime.now().strftime("%Y%m%d_%H%M%S")

# =========================================================
# 공통 함수
# =========================================================

def classify_region_by_text(text):
    text = str(text)

    if "서울" in text:
        return "서울"
    elif "부산" in text:
        return "부산"
    elif "대구" in text:
        return "대구"
    elif "인천" in text:
        return "인천"
    elif "광주" in text:
        return "광주"
    elif "대전" in text:
        return "대전"
    elif "울산" in text:
        return "울산"
    elif "세종" in text:
        return "세종"
    elif "경기" in text or "수원" in text:
        return "경기"
    elif "강원" in text or "원주" in text:
        return "강원"
    elif "충북" in text:
        return "충북"
    elif "충남" in text:
        return "충남"
    elif "전북" in text or "익산" in text:
        return "전북"
    elif "전남" in text:
        return "전남"
    elif "경북" in text:
        return "경북"
    elif "경남" in text:
        return "경남"
    elif "제주" in text:
        return "제주"
    else:
        return "기타"


def safe_to_numeric(series):
    return (
        series.astype(str)
        .str.replace(",", "", regex=False)
        .str.replace(" ", "", regex=False)
        .replace("", "0")
        .pipe(pd.to_numeric, errors="coerce")
        .fillna(0)
    )


# =========================================================
# 1. 나라장터 공사입찰 수집
# =========================================================

def collect_narajangteo():
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
    start_date = end_date - timedelta(days=7)

    first_params = {
        "serviceKey": NARA_API_KEY,
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
            "serviceKey": NARA_API_KEY,
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

    # 취소공고 제외
    if "ntceKindNm" in df.columns:
        df = df[df["ntceKindNm"] != "취소공고"]

    # 수의계약 제외
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

    if "presmptPrce" in df.columns:
        df["추정가격_숫자"] = safe_to_numeric(df["presmptPrce"])
    else:
        df["추정가격_숫자"] = 0

    df["region"] = df["dminsttNm"].apply(classify_region_by_text)

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
    print("\n")
    print("=" * 70)
    print("2. 건설CALS 진행공사 수집 시작")
    print("=" * 70)

    base_url = "https://www.calspia.go.kr/io/openapi/cm/selectIoCmConstructionList.do"

    num_of_rows = 100
    all_items = []

    first_params = {
        "serviceKey": CALS_API_KEY,
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
            "serviceKey": CALS_API_KEY,
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

    # 나라장터는 미래 수요 선행지표이므로 가중치를 조금 높게 둠
    integrated["통합수요지수"] = (
        integrated["건설CALS_철근수요지수"]
        + integrated["나라장터_철근관련입찰건수"] * 1.5
    )

    integrated["수요등급"] = pd.cut(
        integrated["통합수요지수"],
        bins=[-1, 5, 15, 30, 999999],
        labels=["낮음", "보통", "높음", "매우 높음"]
    )

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
        "rgstTyNm": "등록유형",
        "ntceKindNm": "공고종류",
        "cntrctCnclsMthdNm": "계약체결방법",
        "bidMethdNm": "입찰방식",
        "매칭키워드": "매칭키워드",
        "region": "지역",

        # CALS
        "cwkNm": "공사명",
        "ornm": "발주기관",
        "rutNm": "노선/수계",
        "pdznNm": "권역",
        "bzKdNm": "사업종류",
        "stwrDt": "착공일",
        "ccwDt": "준공일",
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
    }

    return df.rename(columns=column_map)


def safe_len(df):
    if df is None:
        return 0
    return len(df)


def safe_sum(df, col):
    if df is None or df.empty or col not in df.columns:
        return 0
    return df[col].sum()


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
        tab_color = "#FFD966"      # 노란색
        header_color = "#BF9000"
    elif group_type == "cals":
        tab_color = "#A9D18E"      # 초록색
        header_color = "#548235"
    elif group_type == "integrated":
        tab_color = "#9DC3E6"      # 파란색
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

    body_format = workbook.add_format({
        "border": 1,
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

    if df is None or df.empty:
        return

    row_count = len(df)
    col_count = len(df.columns)

    # 헤더 재작성
    for col_num, value in enumerate(df.columns):
        worksheet.write(0, col_num, value, header_format)

    # 필터 / 틀고정
    worksheet.autofilter(0, 0, row_count, col_count - 1)
    worksheet.freeze_panes(1, 0)

    # 열 너비 자동 조정
    for idx, col in enumerate(df.columns):
        series = df[col].astype(str)

        max_len = max(
            series.map(len).max() if not series.empty else 0,
            len(str(col))
        )

        width = min(max(max_len + 2, 10), 38)

        if "공고명" in col or "공사명" in col:
            width = 45
        elif "기관" in col:
            width = 28
        elif "가격" in col or "합계" in col:
            width = 18
        elif "일자" in col or "일시" in col:
            width = 18

        worksheet.set_column(idx, idx, width)

    # 행 높이
    worksheet.set_row(0, 24)

    # 숫자 컬럼 서식
    for idx, col in enumerate(df.columns):
        if any(keyword in col for keyword in ["건수", "가격", "합계"]):
            worksheet.set_column(idx, idx, 18, number_format)

        if any(keyword in col for keyword in ["지수", "가중치"]):
            worksheet.set_column(idx, idx, 16, decimal_format)

    # 수요등급 조건부 서식
    if "수요등급" in df.columns:
        grade_col_idx = list(df.columns).index("수요등급")
        grade_col_letter = chr(65 + grade_col_idx)

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
    """
    첫 번째 요약 시트 생성
    """
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

    highlight_format = workbook.add_format({
        "bold": True,
        "font_color": "#C00000",
        "bg_color": "#FFF2CC",
        "border": 1,
        "align": "center",
        "valign": "vcenter"
    })

    # 제목
    worksheet.merge_range("A1:H2", "철근 수요 예측 통합 리포트", title_format)
    worksheet.set_row(0, 30)
    worksheet.set_row(1, 30)

    report_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    # 기본 정보
    worksheet.merge_range("A4:H4", "1. 리포트 기본 정보", section_format)

    summary_rows = [
        ["리포트 생성일", report_time],
        ["나라장터 전체 입찰공고 수", safe_len(nara_df)],
        ["나라장터 철근 관련 입찰 수", safe_len(nara_steel_df)],
        ["건설CALS 진행공사 수", safe_len(cals_df)],
        ["통합 수요 1위 지역", get_top_region(integrated_summary)],
        ["최다 매칭 키워드", get_top_keyword(keyword_summary)],
    ]

    start_row = 5

    for i, row in enumerate(summary_rows):
        worksheet.write(start_row + i, 0, row[0], label_format)
        worksheet.write(start_row + i, 1, row[1], value_format if isinstance(row[1], int) else text_format)

    worksheet.set_column("A:A", 28)
    worksheet.set_column("B:B", 28)

    # 해석 요약
    worksheet.merge_range("D5:H5", "2. 자동 해석 요약", section_format)

    top_region = get_top_region(integrated_summary)
    top_keyword = get_top_keyword(keyword_summary)

    nara_steel_count = safe_len(nara_steel_df)
    cals_count = safe_len(cals_df)

    summary_text = (
        f"이번 리포트는 나라장터 공사 입찰 데이터와 건설CALS 진행공사 데이터를 통합하여 "
        f"지역별 철근 수요 가능성을 분석한 자료입니다.\n\n"
        f"나라장터 데이터는 향후 발주 가능성을 보여주는 선행지표이며, "
        f"건설CALS 데이터는 현재 진행 중인 공사 기반의 현재 수요지표로 볼 수 있습니다.\n\n"
        f"현재 기준 철근 관련 입찰은 {nara_steel_count:,}건, "
        f"건설CALS 진행공사는 {cals_count:,}건으로 집계되었습니다.\n\n"
        f"통합 수요지수 기준으로 가장 주목해야 할 지역은 '{top_region}'이며, "
        f"입찰 키워드 중에서는 '{top_keyword}' 관련 공고가 가장 많이 확인되었습니다.\n\n"
        f"영업 관점에서는 통합 수요지수가 높은 지역을 우선 모니터링하고, "
        f"해당 지역의 수요기관 및 공고기관을 중심으로 발주 예정 공사와 납품 가능성을 추적하는 전략이 필요합니다."
    )

    worksheet.merge_range("D6:H12", summary_text, text_format)

    # TOP 5 지역
    worksheet.merge_range("A13:H13", "3. 통합 철근 수요 TOP 5 지역", section_format)

    top5 = integrated_summary.head(5).copy() if integrated_summary is not None and not integrated_summary.empty else pd.DataFrame()

    if not top5.empty:
        top5 = rename_columns_korean(top5)

        headers = list(top5.columns)

        for col_idx, header in enumerate(headers[:7]):
            worksheet.write(14, col_idx, header, label_format)

        for row_idx, (_, row) in enumerate(top5.iterrows(), start=15):
            for col_idx, header in enumerate(headers[:7]):
                worksheet.write(row_idx, col_idx, row[header], value_format if isinstance(row[header], (int, float)) else text_format)

    # 키워드 TOP 10
    worksheet.merge_range("A23:H23", "4. 나라장터 철근 관련 키워드 TOP 10", section_format)

    top_keywords = keyword_summary.head(10).copy() if keyword_summary is not None and not keyword_summary.empty else pd.DataFrame()

    if not top_keywords.empty:
        top_keywords = rename_columns_korean(top_keywords)

        headers = list(top_keywords.columns)

        for col_idx, header in enumerate(headers):
            worksheet.write(24, col_idx, header, label_format)

        for row_idx, (_, row) in enumerate(top_keywords.iterrows(), start=25):
            for col_idx, header in enumerate(headers):
                worksheet.write(row_idx, col_idx, row[header], value_format if isinstance(row[header], (int, float)) else text_format)

    # 안내 문구
    worksheet.merge_range(
        "D25:H30",
        "※ 해석 기준\n"
        "- 나라장터: 향후 발주·입찰 가능성을 보여주는 선행지표\n"
        "- 건설CALS: 현재 진행공사 기반의 현재 수요지표\n"
        "- 통합 수요지수: 현재 공사 수요와 신규 입찰 흐름을 함께 반영한 참고 지표\n"
        "- 본 지수는 영업 우선순위 판단을 위한 보조 지표이며, 실제 철근 소요량 산정에는 공사 규모·구조·공종별 보정이 추가로 필요합니다.",
        text_format
    )

    worksheet.set_column("A:H", 18)
    worksheet.set_column("D:H", 20)
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
    """
    보기 좋은 통합 엑셀 리포트 생성
    """

    # 컬럼 한글화
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

        # 요약 시트
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

        # 나라장터 시트
        nara_df_kr.to_excel(writer, sheet_name="나라장터_전체입찰", index=False)
        nara_steel_df_kr.to_excel(writer, sheet_name="나라장터_철근입찰", index=False)
        keyword_summary_kr.to_excel(writer, sheet_name="나라장터_키워드", index=False)
        nara_agency_summary_kr.to_excel(writer, sheet_name="나라장터_기관별", index=False)
        nara_region_summary_kr.to_excel(writer, sheet_name="나라장터_지역별", index=False)

        # CALS 시트
        cals_df_kr.to_excel(writer, sheet_name="CALS_진행공사", index=False)
        cals_region_summary_kr.to_excel(writer, sheet_name="CALS_지역별", index=False)
        cals_region_field_summary_kr.to_excel(writer, sheet_name="CALS_지역분야", index=False)
        cals_agency_summary_kr.to_excel(writer, sheet_name="CALS_기관별", index=False)
        cals_demand_summary_kr.to_excel(writer, sheet_name="CALS_수요지수", index=False)

        # 통합 시트
        integrated_summary_kr.to_excel(writer, sheet_name="통합_수요지수", index=False)

        # 서식 적용
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

        format_worksheet(writer, "통합_수요지수", integrated_summary_kr, "integrated")


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