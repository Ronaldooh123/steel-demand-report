import os
import math
import time
import json
import re
import argparse
from pathlib import Path
from datetime import datetime, timedelta

import requests
import urllib3
import pandas as pd
from dotenv import load_dotenv


# =========================================================
# 건축HUB 건축인허가 기본개요 수집기
# - single: 특정 시군구/법정동 테스트
# - sample: 대표 시군구 샘플 테스트
# - priority: 주요 권역/도시 중심 샘플 수집
# - national: 전국 법정동코드 기준 반복 수집
#
# API: 국토교통부_건축HUB_건축인허가정보 서비스
# Endpoint: /ArchPmsHubService/getApBasisOulnInfo
# =========================================================

load_dotenv()
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

BASE_DIR = Path(__file__).parent
CACHE_DIR = BASE_DIR / "data" / "cache"
REFERENCE_DIR = BASE_DIR / "data" / "reference"
CACHE_DIR.mkdir(parents=True, exist_ok=True)
REFERENCE_DIR.mkdir(parents=True, exist_ok=True)

ARCHHUB_API_KEY = os.getenv("ARCHHUB_API_KEY") or os.getenv("NARA_API_KEY")

if not ARCHHUB_API_KEY:
    raise ValueError(
        "ARCHHUB_API_KEY 또는 NARA_API_KEY를 찾을 수 없습니다. .env 파일을 확인하세요."
    )

ARCHHUB_BASIS_URL = (
    "https://apis.data.go.kr/1613000/ArchPmsHubService/getApBasisOulnInfo"
)

DEFAULT_REFERENCE_FILE = REFERENCE_DIR / "bjdong_codes_for_archhub.csv"


SIDO_PREFIX_MAP = {
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

SAMPLE_SIGUNGU_CODES = {
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


# priority 모드에서 먼저 조회할 주요 시군구입니다.
# 철근 영업 관점에서 수도권, 광역시, 산업/택지/혁신도시 중심으로 구성했습니다.
PRIORITY_SIGUNGU_BY_REGION = {
    "서울": [
        "강남구", "서초구", "송파구", "강동구", "마포구", "영등포구", "용산구", "성동구",
        "광진구", "동작구", "강서구", "양천구", "구로구", "금천구", "노원구", "은평구",
        "중구", "종로구",
    ],
    "경기": [
        "화성시", "평택시", "용인시", "수원시", "성남시", "고양시", "남양주시", "안산시",
        "안양시", "부천시", "시흥시", "김포시", "파주시", "광명시", "하남시", "의정부시",
        "이천시", "안성시", "오산시", "군포시", "의왕시", "과천시", "구리시", "광주시", "양주시",
    ],
    "인천": ["연수구", "남동구", "서구", "미추홀구", "부평구", "계양구", "중구", "강화군"],
    "부산": ["해운대구", "강서구", "기장군", "부산진구", "동래구", "남구", "사하구", "연제구", "수영구", "사상구"],
    "대구": ["수성구", "달서구", "북구", "동구", "달성군", "군위군", "중구"],
    "대전": ["유성구", "서구", "대덕구", "동구", "중구"],
    "광주": ["광산구", "북구", "서구", "남구", "동구"],
    "울산": ["울주군", "남구", "북구", "중구", "동구"],
    "세종": ["세종특별자치시", ""],
    "충북": ["청주시", "충주시", "제천시", "진천군", "음성군", "증평군", "옥천군"],
    "충남": ["천안시", "아산시", "당진시", "서산시", "공주시", "논산시", "계룡시", "홍성군", "예산군"],
    "전북": ["전주시", "익산시", "군산시", "완주군", "김제시", "정읍시", "남원시"],
    "전남": ["나주시", "순천시", "여수시", "광양시", "목포시", "무안군", "화순군", "해남군"],
    "경북": ["포항시", "구미시", "경산시", "경주시", "안동시", "김천시", "영주시", "영천시", "칠곡군"],
    "경남": ["창원시", "김해시", "양산시", "진주시", "거제시", "사천시", "밀양시", "통영시", "함안군"],
    "강원": ["원주시", "춘천시", "강릉시", "동해시", "속초시", "삼척시", "홍천군"],
    "제주": ["제주시", "서귀포시"],
}

DEFAULT_PRIORITY_REGIONS = [
    "서울", "경기", "인천", "부산", "대구", "대전", "광주", "울산", "세종",
    "충북", "충남", "전북", "전남", "경북", "경남", "강원", "제주",
]


def region_from_sigungu(sigungu_cd):
    sigungu_cd = str(sigungu_cd).strip()
    if len(sigungu_cd) >= 2:
        return SIDO_PREFIX_MAP.get(sigungu_cd[:2], "기타")
    return "기타"


def safe_to_numeric(series):
    return (
        series.astype(str)
        .str.replace(",", "", regex=False)
        .str.replace(" ", "", regex=False)
        .replace("", "0")
        .replace("nan", "0")
        .pipe(pd.to_numeric, errors="coerce")
        .fillna(0)
    )


def parse_date_safe(value):
    if pd.isna(value):
        return pd.NaT

    value = str(value).strip()
    if value == "" or value.lower() == "nan":
        return pd.NaT

    if len(value) == 8 and value.isdigit():
        return pd.to_datetime(value, format="%Y%m%d", errors="coerce")

    if len(value) == 10 and value[:4].isdigit() and "-" in value:
        return pd.to_datetime(value, errors="coerce")

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


# =========================================================
# 법정동 기준 파일 처리
# =========================================================

def split_lawdong_code(code):
    code = str(code).strip().replace(".0", "")
    code = "".join(ch for ch in code if ch.isdigit()).zfill(10)
    return code, code[:5], code[5:]


def _clean_reference_columns(df):
    df = df.copy()
    df.columns = [str(col).strip().replace("\ufeff", "") for col in df.columns]
    return df


def _manual_parse_bjdong_reference(path, encoding):
    """
    GitHub 웹 편집/엑셀 복사 과정에서 CSV가 탭 또는 공백 정렬 텍스트처럼 저장되는 경우를 대비한 보정 파서입니다.
    최소한 법정동코드, sigunguCd, bjdongCd, 지역, 시도명, 시군구명, 읍면동명, 법정동명, 생성일자를 복원합니다.
    """
    text = Path(path).read_text(encoding=encoding, errors="replace")
    lines = [line.strip() for line in text.splitlines() if line.strip()]

    if not lines:
        return pd.DataFrame()

    # 탭 구분 파일이면 우선 탭 기준으로 복원합니다.
    if "\t" in lines[0]:
        header = [col.strip().replace("\ufeff", "") for col in lines[0].split("\t")]
        rows = []
        for line in lines[1:]:
            values = [value.strip() for value in line.split("\t")]
            if len(values) < len(header):
                values += [""] * (len(header) - len(values))
            rows.append(values[:len(header)])
        df = pd.DataFrame(rows, columns=header)
        return _clean_reference_columns(df)

    # 공백 정렬 텍스트로 저장된 경우의 최후 보정.
    # 법정동명은 공백을 포함할 수 있으므로, 앞쪽 고정 컬럼과 마지막 날짜만 확정적으로 분리합니다.
    rows = []
    for line in lines[1:]:
        parts = line.split()
        if len(parts) < 7:
            continue

        lawdong_code = parts[0]
        if not str(lawdong_code).isdigit():
            continue

        sigungu_cd = parts[1] if len(parts) > 1 else str(lawdong_code)[:5]
        bjdong_cd = parts[2] if len(parts) > 2 else str(lawdong_code)[5:]
        region = parts[3] if len(parts) > 3 else region_from_sigungu(sigungu_cd)
        sido_name = parts[4] if len(parts) > 4 else ""
        sigungu_name = parts[5] if len(parts) > 5 else ""
        eupmyeondong_name = parts[6] if len(parts) > 6 else ""

        created_day = ""
        tail = parts[7:]
        if tail and re.match(r"^\d{4}-\d{2}-\d{2}$", tail[-1]):
            created_day = tail[-1]
            tail = tail[:-1]

        # 리명은 정확 복원이 어려우므로 비워두고, 남은 문자열을 법정동명으로 둡니다.
        # 수요지수 산정에는 법정동명보다 sigunguCd/bjdongCd/지역이 중요합니다.
        lawdong_name = " ".join(tail).strip()
        if not lawdong_name:
            lawdong_name = " ".join([sido_name, sigungu_name, eupmyeondong_name]).strip()

        rows.append({
            "법정동코드": lawdong_code,
            "sigunguCd": sigungu_cd,
            "bjdongCd": bjdong_cd,
            "지역": region,
            "시도명": sido_name,
            "시군구명": sigungu_name,
            "읍면동명": eupmyeondong_name,
            "리명": "",
            "법정동명": lawdong_name,
            "생성일자": created_day,
        })

    return pd.DataFrame(rows)


def read_bjdong_reference_file(path):
    """
    기준 파일을 쉼표 CSV, 탭 TSV, 자동 구분자, 공백 정렬 텍스트 순서로 읽습니다.
    GitHub 웹에서 CSV를 수정하면서 탭/공백 형태로 바뀐 경우까지 흡수합니다.
    """
    path = Path(path)
    encodings = ["utf-8-sig", "utf-8", "cp949", "euc-kr"]
    attempts = []

    for encoding in encodings:
        for sep, engine in [(",", "c"), ("\t", "c"), (";", "c"), ("|", "c"), (None, "python")]:
            try:
                kwargs = {
                    "dtype": str,
                    "encoding": encoding,
                    "keep_default_na": False,
                }
                if sep is None:
                    kwargs["sep"] = None
                    kwargs["engine"] = "python"
                else:
                    kwargs["sep"] = sep
                    if engine == "python":
                        kwargs["engine"] = "python"

                df = pd.read_csv(path, **kwargs)
                df = _clean_reference_columns(df)

                if "법정동코드" in df.columns and len(df.columns) > 1:
                    print(f"[법정동 기준] 파일 읽기 성공: encoding={encoding}, sep={repr(sep)}")
                    return df

                attempts.append(f"encoding={encoding}, sep={repr(sep)}, columns={list(df.columns)[:3]}")
            except Exception as exc:
                attempts.append(f"encoding={encoding}, sep={repr(sep)}, error={type(exc).__name__}: {exc}")

    for encoding in encodings:
        try:
            df = _manual_parse_bjdong_reference(path, encoding)
            df = _clean_reference_columns(df)
            if "법정동코드" in df.columns and not df.empty:
                print(f"[법정동 기준] 수동 파싱 성공: encoding={encoding}")
                return df
        except Exception as exc:
            attempts.append(f"manual encoding={encoding}, error={type(exc).__name__}: {exc}")

    print("[법정동 기준 오류] 파일 읽기 시도 내역")
    for item in attempts[:20]:
        print("-", item)

    raise ValueError(
        "법정동코드 컬럼이 필요합니다. "
        "data/reference/bjdong_codes_for_archhub.csv 파일이 쉼표 CSV 또는 탭 TSV 형식인지 확인하세요."
    )


def normalize_bjdong_reference(path):
    """
    지원 파일 형식:
    1) 전처리 파일: 법정동코드,sigunguCd,bjdongCd,지역,시도명,시군구명,읍면동명,리명,법정동명
    2) 원본 파일: 법정동코드,시도명,시군구명,읍면동명,리명,순위,생성일자,삭제일자
    3) GitHub/Excel에서 생긴 탭 구분 또는 공백 정렬 텍스트
    """
    path = Path(path)
    if not path.exists():
        raise FileNotFoundError(f"법정동코드 기준 파일을 찾을 수 없습니다: {path}")

    df = read_bjdong_reference_file(path)
    df.columns = [str(col).strip().replace("\ufeff", "") for col in df.columns]

    if "법정동코드" not in df.columns:
        print("[법정동 기준 오류] 현재 컬럼:", list(df.columns))
        raise ValueError("법정동코드 컬럼이 필요합니다.")

    df = df.copy()
    df["법정동코드"] = df["법정동코드"].apply(lambda x: split_lawdong_code(x)[0])

    # 원본 파일이면 삭제일자 비어 있는 현재 유효 법정동만 사용
    if "삭제일자" in df.columns:
        df["삭제일자"] = df["삭제일자"].fillna("").astype(str).str.strip()
        df = df[df["삭제일자"].isin(["", "nan", "NaN", "None"])]

    if "sigunguCd" not in df.columns:
        df["sigunguCd"] = df["법정동코드"].str[:5]
    else:
        df["sigunguCd"] = df["sigunguCd"].apply(lambda x: str(x).strip().replace(".0", "").zfill(5))

    if "bjdongCd" not in df.columns:
        df["bjdongCd"] = df["법정동코드"].str[5:]
    else:
        df["bjdongCd"] = df["bjdongCd"].apply(lambda x: str(x).strip().replace(".0", "").zfill(5))

    # 시도/시군구 대표행 제외
    df = df[df["bjdongCd"] != "00000"]

    if "지역" not in df.columns:
        df["지역"] = df["sigunguCd"].apply(region_from_sigungu)

    for col in ["시도명", "시군구명", "읍면동명", "리명"]:
        if col not in df.columns:
            df[col] = ""
        df[col] = df[col].fillna("").astype(str).str.strip()

    if "법정동명" not in df.columns:
        df["법정동명"] = (
            df["시도명"].fillna("").astype(str).str.strip()
            + " "
            + df["시군구명"].fillna("").astype(str).str.strip()
            + " "
            + df["읍면동명"].fillna("").astype(str).str.strip()
            + " "
            + df["리명"].fillna("").astype(str).str.strip()
        ).str.replace(r"\s+", " ", regex=True).str.strip()

    # 리 단위가 존재하는 읍/면은 리 단위만 사용하고, 부모 읍/면 행은 제외
    df["읍면동그룹"] = df["법정동코드"].str[:8]
    df["리코드"] = df["법정동코드"].str[8:]
    group_has_ri = df.groupby("읍면동그룹")["리코드"].transform(lambda s: (s != "00").any())
    df = df[~((df["리코드"] == "00") & group_has_ri)]

    df = df.drop_duplicates(subset=["sigunguCd", "bjdongCd"])

    cols = [
        "법정동코드",
        "sigunguCd",
        "bjdongCd",
        "지역",
        "시도명",
        "시군구명",
        "읍면동명",
        "리명",
        "법정동명",
    ]
    if "생성일자" in df.columns:
        cols.append("생성일자")

    return df[[col for col in cols if col in df.columns]].reset_index(drop=True)

def save_normalized_reference_if_needed(ref_df):
    """
    원본 기준 파일(data/reference/bjdong_codes_for_archhub.csv)을 덮어쓰지 않도록
    정규화본은 별도 파일로 저장합니다.
    """
    out_path = REFERENCE_DIR / "bjdong_codes_for_archhub_normalized.csv"
    out_path.parent.mkdir(parents=True, exist_ok=True)
    ref_df.to_csv(out_path, index=False, encoding="utf-8-sig")
    return out_path


def select_priority_reference(ref_df, regions=None, max_per_region=50):
    """
    전국 법정동 전체 대신 주요 권역/도시 중심으로 법정동을 샘플링합니다.
    - regions가 있으면 해당 광역지역만 사용합니다.
    - max_per_region은 광역지역별 최대 법정동 수입니다.
    - 시군구 우선순위는 PRIORITY_SIGUNGU_BY_REGION을 따릅니다.
    """
    if ref_df.empty:
        return ref_df

    df = ref_df.copy()

    if regions:
        wanted_regions = [str(x).strip() for x in regions if str(x).strip()]
    else:
        wanted_regions = DEFAULT_PRIORITY_REGIONS

    df = df[df["지역"].isin(wanted_regions)].copy()

    if df.empty:
        return df

    for col in ["시군구명", "읍면동명", "리명", "법정동명"]:
        if col not in df.columns:
            df[col] = ""
        df[col] = df[col].fillna("").astype(str)

    def calc_priority_rank(row):
        region = str(row.get("지역", ""))
        sigungu_name = str(row.get("시군구명", ""))

        if region == "세종":
            return 0

        sigungu_list = PRIORITY_SIGUNGU_BY_REGION.get(region, [])
        for rank, priority_name in enumerate(sigungu_list):
            priority_name = str(priority_name).strip()
            if priority_name == "":
                continue

            # 예: 수원시장안구는 수원시로, 화성시동탄구는 화성시로 매칭합니다.
            if sigungu_name == priority_name or sigungu_name.startswith(priority_name):
                return rank

        return 999

    df["우선시군구순위"] = df.apply(calc_priority_rank, axis=1)
    df["읍면동명_길이"] = df["읍면동명"].astype(str).str.len()
    df["리명_존재"] = df["리명"].astype(str).str.strip().ne("")

    df = df.sort_values(
        by=["지역", "우선시군구순위", "sigunguCd", "리명_존재", "읍면동명_길이", "bjdongCd"],
        ascending=[True, True, True, True, False, True],
    )

    if max_per_region is not None and max_per_region > 0:
        df = df.groupby("지역", group_keys=False).head(max_per_region)

    df = df.reset_index(drop=True)
    return df


def collect_archhub_basis_priority(
    reference_file,
    start_date,
    end_date,
    regions=None,
    max_per_region=50,
    limit=None,
    start_index=0,
    max_pages=1,
    num_of_rows=100,
    sleep_sec=0.05,
    save_every=50,
    resume=False,
):
    ref_df = normalize_bjdong_reference(reference_file)
    save_normalized_reference_if_needed(ref_df)

    priority_ref_df = select_priority_reference(
        ref_df=ref_df,
        regions=regions,
        max_per_region=max_per_region,
    )

    if start_index > 0:
        priority_ref_df = priority_ref_df.iloc[start_index:].copy().reset_index(drop=True)

    if limit is not None and limit > 0:
        priority_ref_df = priority_ref_df.head(limit).copy()

    temp_reference_path = CACHE_DIR / "archhub_priority_reference_current.csv"
    priority_ref_df.to_csv(temp_reference_path, index=False, encoding="utf-8-sig")

    print("\n" + "=" * 70)
    print("건축HUB priority 법정동 선별 완료")
    print("=" * 70)
    print(f"선별 법정동 수: {len(priority_ref_df):,}개")
    print(f"지역 필터: {', '.join(regions) if regions else '기본 priority 전체'}")
    print(f"지역별 최대 법정동 수: {max_per_region}")
    print(f"임시 기준 파일: {temp_reference_path}")

    return collect_archhub_basis_national(
        reference_file=temp_reference_path,
        start_date=start_date,
        end_date=end_date,
        regions=None,
        limit=None,
        start_index=0,
        max_pages=max_pages,
        num_of_rows=num_of_rows,
        sleep_sec=sleep_sec,
        save_every=save_every,
        resume=resume,
    )


# =========================================================
# API 응답 처리
# =========================================================

def extract_items(data):
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


def get_total_count(data):
    response = data.get("response", {}) if isinstance(data, dict) else {}
    body = response.get("body", {}) if isinstance(response, dict) else {}
    total_count = body.get("totalCount", 0) if isinstance(body, dict) else 0

    try:
        return int(total_count)
    except Exception:
        return 0


def get_result_info(data):
    response = data.get("response", {}) if isinstance(data, dict) else {}
    header = response.get("header", {}) if isinstance(response, dict) else {}

    result_code = str(header.get("resultCode", ""))
    result_msg = str(header.get("resultMsg", ""))

    return result_code, result_msg


def is_response_ok(data):
    if data is None:
        return False

    result_code, _ = get_result_info(data)

    if result_code in ["00", "NORMAL_SERVICE"]:
        return True

    response = data.get("response", {}) if isinstance(data, dict) else {}
    body = response.get("body", {}) if isinstance(response, dict) else {}

    if isinstance(body, dict):
        if "items" in body:
            return True
        if "totalCount" in body:
            return True

    return False


def response_json(response):
    try:
        return response.json()
    except Exception:
        print("[건축HUB 오류] JSON 변환 실패")
        print("응답 앞부분:")
        print((response.text or "")[:500])
        return None


def request_archhub_page(
    sigungu_cd,
    start_date,
    end_date,
    page_no=1,
    num_of_rows=100,
    bjdong_cd="",
    verbose=True,
):
    params = {
        "serviceKey": ARCHHUB_API_KEY,
        "sigunguCd": str(sigungu_cd).strip(),
        "startDate": start_date,
        "endDate": end_date,
        "numOfRows": str(num_of_rows),
        "pageNo": str(page_no),
        "_type": "json",
    }

    if str(bjdong_cd).strip() != "":
        params["bjdongCd"] = str(bjdong_cd).strip().zfill(5)

    try:
        response = requests.get(
            ARCHHUB_BASIS_URL,
            params=params,
            timeout=30,
            verify=False,
        )
    except Exception as exc:
        if verbose:
            print(f"[건축HUB 오류] 요청 실패: {exc}")
        return None, 0

    if verbose:
        bjdong_label = params.get("bjdongCd", "전체")
        print(
            f"[건축HUB] sigunguCd={sigungu_cd or '전체'} "
            f"bjdongCd={bjdong_label} page={page_no} status={response.status_code}"
        )

    data = response_json(response)

    if data is None:
        return None, response.status_code

    if verbose:
        result_code, result_msg = get_result_info(data)
        print(f"[건축HUB] resultCode={result_code}, resultMsg={result_msg}")

    return data, response.status_code


# =========================================================
# 수집 로직
# =========================================================

def collect_archhub_basis_by_bjdong(
    sigungu_cd,
    bjdong_cd,
    start_date,
    end_date,
    lawdong_name="",
    lawdong_code="",
    max_pages=1,
    num_of_rows=100,
    sleep_sec=0.05,
    verbose=True,
):
    data, status_code = request_archhub_page(
        sigungu_cd=sigungu_cd,
        bjdong_cd=bjdong_cd,
        start_date=start_date,
        end_date=end_date,
        page_no=1,
        num_of_rows=num_of_rows,
        verbose=verbose,
    )

    if data is None or status_code >= 500 or not is_response_ok(data):
        return pd.DataFrame(), {
            "status": "failed",
            "http_status": status_code,
            "totalCount": 0,
            "rows": 0,
        }

    total_count = get_total_count(data)
    first_items = extract_items(data)

    total_pages = math.ceil(total_count / num_of_rows) if total_count else 1

    if max_pages is not None and max_pages > 0:
        total_pages = min(total_pages, max_pages)

    all_items = []
    all_items.extend(first_items)

    for page in range(2, total_pages + 1):
        page_data, page_status = request_archhub_page(
            sigungu_cd=sigungu_cd,
            bjdong_cd=bjdong_cd,
            start_date=start_date,
            end_date=end_date,
            page_no=page,
            num_of_rows=num_of_rows,
            verbose=verbose,
        )

        if page_data is None or not is_response_ok(page_data):
            print(f"[건축HUB 경고] {sigungu_cd}-{bjdong_cd} {page}페이지 수집 실패")
            continue

        all_items.extend(extract_items(page_data))
        time.sleep(sleep_sec)

    df = pd.DataFrame(all_items)

    if not df.empty:
        df["조회_법정동코드"] = str(lawdong_code) if lawdong_code else str(sigungu_cd) + str(bjdong_cd)
        df["조회_sigunguCd"] = str(sigungu_cd).zfill(5)
        df["조회_bjdongCd"] = str(bjdong_cd).zfill(5)
        df["조회_법정동명"] = lawdong_name
        df["조회_region"] = region_from_sigungu(sigungu_cd)
        df["조회_totalCount"] = total_count

    return df, {
        "status": "ok" if len(df) > 0 else "empty",
        "http_status": status_code,
        "totalCount": total_count,
        "rows": int(len(df)),
    }


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
    # 시군구 전체 조회를 우선 시도하고, 안 되면 테스트용 fallback 법정동으로 재시도합니다.
    data, status_code = request_archhub_page(
        sigungu_cd=sigungu_cd,
        bjdong_cd=bjdong_cd,
        start_date=start_date,
        end_date=end_date,
        page_no=1,
        num_of_rows=num_of_rows,
    )

    used_bjdong_cd = bjdong_cd

    if data is None or status_code >= 500 or not is_response_ok(data):
        if str(bjdong_cd).strip() == "" and fallback_bjdong_cd:
            print(
                f"[재시도] 시군구 전체 조회가 유효하지 않음 → "
                f"fallback bjdongCd={fallback_bjdong_cd}로 재시도"
            )
            data, status_code = request_archhub_page(
                sigungu_cd=sigungu_cd,
                bjdong_cd=fallback_bjdong_cd,
                start_date=start_date,
                end_date=end_date,
                page_no=1,
                num_of_rows=num_of_rows,
            )
            used_bjdong_cd = fallback_bjdong_cd

    if data is None or not is_response_ok(data):
        return pd.DataFrame()

    total_count = get_total_count(data)
    first_items = extract_items(data)

    if total_count == 0 and len(first_items) == 0:
        if str(bjdong_cd).strip() == "" and fallback_bjdong_cd and used_bjdong_cd != fallback_bjdong_cd:
            print(
                f"[재시도] 시군구 전체 조회 결과 0건 → "
                f"fallback bjdongCd={fallback_bjdong_cd}로 재시도"
            )
            data, status_code = request_archhub_page(
                sigungu_cd=sigungu_cd,
                bjdong_cd=fallback_bjdong_cd,
                start_date=start_date,
                end_date=end_date,
                page_no=1,
                num_of_rows=num_of_rows,
            )
            used_bjdong_cd = fallback_bjdong_cd

            if data is None or not is_response_ok(data):
                return pd.DataFrame()

            total_count = get_total_count(data)
            first_items = extract_items(data)

    total_pages = math.ceil(total_count / num_of_rows) if total_count else 1

    if max_pages is not None and max_pages > 0:
        total_pages = min(total_pages, max_pages)

    print(f"[건축HUB] totalCount={total_count}, 수집 예정 페이지={total_pages}")

    all_items = []
    all_items.extend(first_items)

    for page in range(2, total_pages + 1):
        page_data, _ = request_archhub_page(
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

        all_items.extend(extract_items(page_data))
        time.sleep(sleep_sec)

    df = pd.DataFrame(all_items)

    if not df.empty:
        df["조회_법정동코드"] = str(sigungu_cd).zfill(5) + str(used_bjdong_cd).zfill(5)
        df["조회_sigunguCd"] = str(sigungu_cd).zfill(5)
        df["조회_bjdongCd"] = used_bjdong_cd if str(used_bjdong_cd).strip() != "" else "전체"
        df["조회_법정동명"] = ""
        df["조회_region"] = region_from_sigungu(sigungu_cd)
        df["조회_totalCount"] = total_count

    return df


def collect_archhub_basis_sample(start_date, end_date, bjdong_cd="", max_pages=1, num_of_rows=100):
    frames = []

    for label, sigungu_cd in SAMPLE_SIGUNGU_CODES.items():
        print("\n" + "=" * 70)
        print(f"샘플 지역 수집: {label} / {sigungu_cd}")
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
            print(f"[경고] {label} 수집 실패: {exc}")

        time.sleep(0.2)

    if not frames:
        return pd.DataFrame()

    return pd.concat(frames, ignore_index=True)


def collect_archhub_basis_national(
    reference_file,
    start_date,
    end_date,
    regions=None,
    limit=200,
    start_index=0,
    max_pages=1,
    num_of_rows=100,
    sleep_sec=0.05,
    save_every=50,
    resume=False,
):
    ref_df = normalize_bjdong_reference(reference_file)
    save_normalized_reference_if_needed(ref_df)

    if regions:
        wanted_regions = [x.strip() for x in regions if str(x).strip()]
        ref_df = ref_df[ref_df["지역"].isin(wanted_regions)].copy()

    ref_df = ref_df.reset_index(drop=True)

    if start_index > 0:
        ref_df = ref_df.iloc[start_index:].copy().reset_index(drop=True)

    if limit is not None and limit > 0:
        ref_df = ref_df.head(limit).copy()

    progress_raw_path = CACHE_DIR / "archhub_national_raw_progress.csv"
    attempted_path = CACHE_DIR / "archhub_national_attempted_pairs.csv"

    raw_frames = []
    attempted_records = []
    skip_pairs = set()

    if resume and attempted_path.exists():
        attempted_old = pd.read_csv(attempted_path, dtype=str, encoding="utf-8-sig")
        attempted_records.extend(attempted_old.to_dict("records"))
        if not attempted_old.empty:
            attempted_old["pair"] = attempted_old["sigunguCd"].astype(str).str.zfill(5) + "-" + attempted_old["bjdongCd"].astype(str).str.zfill(5)
            skip_old = attempted_old[attempted_old["status"].isin(["ok", "empty"])]
            skip_pairs = set(skip_old["pair"].tolist())

    if resume and progress_raw_path.exists():
        progress_raw = pd.read_csv(progress_raw_path, dtype=str, encoding="utf-8-sig")
        if not progress_raw.empty:
            raw_frames.append(progress_raw)

    total_targets = len(ref_df)
    print("\n" + "=" * 70)
    print("건축HUB 전국 법정동 수집 시작")
    print("=" * 70)
    print(f"기준 파일: {reference_file}")
    print(f"조회기간: {start_date} ~ {end_date}")
    print(f"수집 대상 법정동 수: {total_targets:,}개")
    print(f"지역 필터: {', '.join(regions) if regions else '전체'}")
    print(f"limit: {'전체' if not limit else limit}")
    print(f"max_pages: {'전체' if max_pages is None else max_pages}")
    print(f"resume: {resume}")

    for i, row in ref_df.iterrows():
        sigungu_cd = str(row["sigunguCd"]).zfill(5)
        bjdong_cd = str(row["bjdongCd"]).zfill(5)
        lawdong_code = str(row["법정동코드"]).zfill(10)
        lawdong_name = str(row.get("법정동명", ""))
        region = str(row.get("지역", region_from_sigungu(sigungu_cd)))
        pair = f"{sigungu_cd}-{bjdong_cd}"

        if pair in skip_pairs:
            print(f"[{i + 1:,}/{total_targets:,}] SKIP {lawdong_name} ({pair})")
            continue

        print(f"[{i + 1:,}/{total_targets:,}] {region} {lawdong_name} ({pair})")

        try:
            part_df, meta = collect_archhub_basis_by_bjdong(
                sigungu_cd=sigungu_cd,
                bjdong_cd=bjdong_cd,
                lawdong_code=lawdong_code,
                lawdong_name=lawdong_name,
                start_date=start_date,
                end_date=end_date,
                max_pages=max_pages,
                num_of_rows=num_of_rows,
                sleep_sec=sleep_sec,
                verbose=False,
            )

            if not part_df.empty:
                raw_frames.append(part_df)

            attempted_records.append(
                {
                    "collected_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    "법정동코드": lawdong_code,
                    "sigunguCd": sigungu_cd,
                    "bjdongCd": bjdong_cd,
                    "지역": region,
                    "법정동명": lawdong_name,
                    "status": meta.get("status", "unknown"),
                    "http_status": meta.get("http_status", ""),
                    "totalCount": meta.get("totalCount", 0),
                    "rows": meta.get("rows", 0),
                }
            )

            print(
                f"    → {meta.get('status')} / totalCount={meta.get('totalCount')} / rows={meta.get('rows')}"
            )

        except Exception as exc:
            print(f"    → failed: {exc}")
            attempted_records.append(
                {
                    "collected_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    "법정동코드": lawdong_code,
                    "sigunguCd": sigungu_cd,
                    "bjdongCd": bjdong_cd,
                    "지역": region,
                    "법정동명": lawdong_name,
                    "status": "failed",
                    "http_status": "",
                    "totalCount": 0,
                    "rows": 0,
                    "error": str(exc)[:500],
                }
            )

        if save_every and (i + 1) % save_every == 0:
            save_national_progress(raw_frames, attempted_records, progress_raw_path, attempted_path)
            print(f"    [중간저장] {progress_raw_path}")

        time.sleep(sleep_sec)

    save_national_progress(raw_frames, attempted_records, progress_raw_path, attempted_path)

    if raw_frames:
        raw_df = pd.concat(raw_frames, ignore_index=True)
        raw_df = raw_df.drop_duplicates(subset=["mgmPmsrgstPk"], keep="first") if "mgmPmsrgstPk" in raw_df.columns else raw_df.drop_duplicates()
    else:
        raw_df = pd.DataFrame()

    attempted_df = pd.DataFrame(attempted_records)

    return raw_df, attempted_df, ref_df


def save_national_progress(raw_frames, attempted_records, progress_raw_path, attempted_path):
    if raw_frames:
        progress_raw_df = pd.concat(raw_frames, ignore_index=True)
        if "mgmPmsrgstPk" in progress_raw_df.columns:
            progress_raw_df = progress_raw_df.drop_duplicates(subset=["mgmPmsrgstPk"], keep="first")
        progress_raw_df.to_csv(progress_raw_path, index=False, encoding="utf-8-sig")

    if attempted_records:
        attempted_df = pd.DataFrame(attempted_records)
        attempted_df = attempted_df.drop_duplicates(subset=["sigunguCd", "bjdongCd"], keep="last")
        attempted_df.to_csv(attempted_path, index=False, encoding="utf-8-sig")


# =========================================================
# 정규화 및 선행수요지수
# =========================================================

def normalize_archhub_basis(df):
    if df.empty:
        return df

    out = pd.DataFrame()

    out["원천관리번호"] = get_series_or_default(df, ["mgmPmsrgstPk", "관리번호", "mgmNo"], "")
    out["원천_법정동코드"] = get_series_or_default(df, ["조회_법정동코드"], "")
    out["원천_sigunguCd"] = get_series_or_default(df, ["sigunguCd", "조회_sigunguCd"], "")
    out["원천_bjdongCd"] = get_series_or_default(df, ["bjdongCd", "조회_bjdongCd"], "")
    out["지역"] = out["원천_sigunguCd"].apply(region_from_sigungu)
    out["법정동명"] = get_series_or_default(df, ["조회_법정동명"], "")

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
    out["조회_totalCount"] = safe_to_numeric(get_series_or_default(df, ["조회_totalCount"], "0"))

    if "샘플지역명" in df.columns:
        out["샘플지역명"] = df["샘플지역명"]

    return out


def calc_purpose_weight(value):
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


def classify_start_status(row):
    if pd.notna(row.get("사용승인일")):
        return "사용승인"
    if pd.notna(row.get("실제착공일")):
        return "실제착공"
    if pd.notna(row.get("착공예정일")):
        return "착공예정"
    if pd.notna(row.get("허가일")):
        return "허가"
    return "일정미확인"


def calc_status_weight(status):
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

    df["착공상태"] = df.apply(classify_start_status, axis=1)
    df["용도가중치"] = df["주용도"].apply(calc_purpose_weight)
    df["착공상태가중치"] = df["착공상태"].apply(calc_status_weight)

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
        df.groupby("지역")
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


def make_archhub_sigungu_summary(df):
    if df.empty:
        return pd.DataFrame()

    temp = df.copy()
    temp["시군구코드"] = temp["원천_sigunguCd"].astype(str).str.zfill(5)

    summary = (
        temp.groupby(["지역", "시군구코드"])
        .agg(
            건축허가건수=("원천관리번호", "count"),
            건축허가_연면적합계=("연면적", "sum"),
            세대수합계=("세대수", "sum"),
            대형허가건수=("대형허가여부", "sum"),
            건축허가_선행수요지수=("건축허가_선행수요지수", "sum"),
        )
        .reset_index()
        .sort_values(by="건축허가_선행수요지수", ascending=False)
    )

    summary.insert(0, "시군구수요순위", range(1, len(summary) + 1))

    return summary


# =========================================================
# 저장
# =========================================================

def save_outputs(raw_df, norm_df, region_summary_df, sigungu_summary_df=None, attempted_df=None, ref_df=None, run_mode=""):
    now = datetime.now().strftime("%Y%m%d_%H%M%S")

    files = {
        "raw_timestamp": CACHE_DIR / f"archhub_basis_raw_{now}.csv",
        "normalized_timestamp": CACHE_DIR / f"archhub_basis_normalized_{now}.csv",
        "region_summary_timestamp": CACHE_DIR / f"archhub_basis_region_summary_{now}.csv",
        "raw_latest": CACHE_DIR / "archhub_basis_raw_latest.csv",
        "normalized_latest": CACHE_DIR / "archhub_basis_normalized_latest.csv",
        "region_summary_latest": CACHE_DIR / "archhub_basis_region_summary_latest.csv",
    }

    raw_df.to_csv(files["raw_timestamp"], index=False, encoding="utf-8-sig")
    norm_df.to_csv(files["normalized_timestamp"], index=False, encoding="utf-8-sig")
    region_summary_df.to_csv(files["region_summary_timestamp"], index=False, encoding="utf-8-sig")

    raw_df.to_csv(files["raw_latest"], index=False, encoding="utf-8-sig")
    norm_df.to_csv(files["normalized_latest"], index=False, encoding="utf-8-sig")
    region_summary_df.to_csv(files["region_summary_latest"], index=False, encoding="utf-8-sig")

    if sigungu_summary_df is not None:
        files["sigungu_summary_timestamp"] = CACHE_DIR / f"archhub_basis_sigungu_summary_{now}.csv"
        files["sigungu_summary_latest"] = CACHE_DIR / "archhub_basis_sigungu_summary_latest.csv"
        sigungu_summary_df.to_csv(files["sigungu_summary_timestamp"], index=False, encoding="utf-8-sig")
        sigungu_summary_df.to_csv(files["sigungu_summary_latest"], index=False, encoding="utf-8-sig")

    if attempted_df is not None:
        files["attempted_timestamp"] = CACHE_DIR / f"archhub_national_attempted_pairs_{now}.csv"
        files["attempted_latest"] = CACHE_DIR / "archhub_national_attempted_pairs_latest.csv"
        attempted_df.to_csv(files["attempted_timestamp"], index=False, encoding="utf-8-sig")
        attempted_df.to_csv(files["attempted_latest"], index=False, encoding="utf-8-sig")

    meta = {
        "created_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "mode": run_mode,
        "raw_rows": int(len(raw_df)),
        "normalized_rows": int(len(norm_df)),
        "region_summary_rows": int(len(region_summary_df)),
        "sigungu_summary_rows": int(len(sigungu_summary_df)) if sigungu_summary_df is not None else 0,
        "attempted_rows": int(len(attempted_df)) if attempted_df is not None else 0,
        "reference_rows": int(len(ref_df)) if ref_df is not None else 0,
        "files": {key: str(value) for key, value in files.items()},
    }

    meta_path = CACHE_DIR / "archhub_cache_meta.json"
    meta_path.write_text(json.dumps(meta, ensure_ascii=False, indent=2), encoding="utf-8")

    return files, meta_path


# =========================================================
# 실행부
# =========================================================

def main():
    parser = argparse.ArgumentParser()

    parser.add_argument(
        "--mode",
        choices=["single", "sample", "priority", "national"],
        default="single",
        help="single: 특정 시군구/법정동 / sample: 대표지역 / priority: 주요 권역 샘플 / national: 전국 법정동 반복 수집",
    )
    parser.add_argument("--sigunguCd", default="11680", help="single 모드 시군구 코드. 기본값: 서울 강남구 11680")
    parser.add_argument("--bjdongCd", default="", help="single 모드 법정동 코드. 예: 역삼동 10100")
    parser.add_argument("--reference-file", default=str(DEFAULT_REFERENCE_FILE), help="전국 법정동코드 CSV 경로")
    parser.add_argument("--regions", default="", help="priority/national 모드 지역 필터. 예: 서울,경기,인천")
    parser.add_argument("--days", type=int, default=90, help="최근 N일 조회. 기본값 90일")
    parser.add_argument("--start-date", default=None, help="조회 시작일 YYYYMMDD")
    parser.add_argument("--end-date", default=None, help="조회 종료일 YYYYMMDD")
    parser.add_argument("--max-pages", type=int, default=1, help="법정동별 최대 페이지. 0이면 전체 페이지")
    parser.add_argument("--num-rows", type=int, default=100, help="페이지당 건수. 기본값 100")
    parser.add_argument("--limit", type=int, default=None, help="priority/national 모드 전체 수집 법정동 수 제한. 0이면 전체")
    parser.add_argument("--max-per-region", type=int, default=50, help="priority 모드 광역지역별 최대 법정동 수. 기본값 50")
    parser.add_argument("--start-index", type=int, default=0, help="priority/national 모드 시작 인덱스")
    parser.add_argument("--sleep-sec", type=float, default=0.05, help="API 호출 간 대기시간")
    parser.add_argument("--save-every", type=int, default=50, help="priority/national 모드 중간저장 주기")
    parser.add_argument("--resume", action="store_true", help="priority/national 모드 중단 지점 이어받기")

    args = parser.parse_args()

    end_dt = datetime.now()
    start_dt = end_dt - timedelta(days=args.days)

    start_date = args.start_date or start_dt.strftime("%Y%m%d")
    end_date = args.end_date or end_dt.strftime("%Y%m%d")
    max_pages = None if args.max_pages == 0 else args.max_pages
    regions = [x.strip() for x in args.regions.split(",") if x.strip()] if args.regions else None

    if args.mode == "national":
        # national 모드는 실수로 전국 18,000개를 바로 때리는 것을 막기 위해 기본 200개 제한을 둡니다.
        limit = 200 if args.limit is None else (None if args.limit == 0 else args.limit)
    else:
        # priority 모드는 max-per-region으로 통제하므로 기본 limit은 적용하지 않습니다.
        limit = None if args.limit is None or args.limit == 0 else args.limit

    print("\n" + "=" * 70)
    print("건축HUB 건축인허가 기본개요 수집")
    print("=" * 70)
    print(f"조회기간: {start_date} ~ {end_date}")
    print(f"실행모드: {args.mode}")

    attempted_df = None
    ref_df = None

    if args.mode == "national":
        raw_df, attempted_df, ref_df = collect_archhub_basis_national(
            reference_file=args.reference_file,
            start_date=start_date,
            end_date=end_date,
            regions=regions,
            limit=limit,
            start_index=args.start_index,
            max_pages=max_pages,
            num_of_rows=args.num_rows,
            sleep_sec=args.sleep_sec,
            save_every=args.save_every,
            resume=args.resume,
        )
    elif args.mode == "priority":
        raw_df, attempted_df, ref_df = collect_archhub_basis_priority(
            reference_file=args.reference_file,
            start_date=start_date,
            end_date=end_date,
            regions=regions,
            max_per_region=args.max_per_region,
            limit=limit,
            start_index=args.start_index,
            max_pages=max_pages,
            num_of_rows=args.num_rows,
            sleep_sec=args.sleep_sec,
            save_every=args.save_every,
            resume=args.resume,
        )
    elif args.mode == "sample":
        raw_df = collect_archhub_basis_sample(
            start_date=start_date,
            end_date=end_date,
            bjdong_cd=args.bjdongCd,
            max_pages=max_pages,
            num_of_rows=args.num_rows,
        )
    else:
        raw_df = collect_archhub_basis_by_sigungu(
            sigungu_cd=args.sigunguCd,
            bjdong_cd=args.bjdongCd,
            start_date=start_date,
            end_date=end_date,
            max_pages=max_pages,
            num_of_rows=args.num_rows,
        )

    if raw_df.empty:
        print("\n[결과] 수집 데이터가 없습니다.")
        if attempted_df is not None and not attempted_df.empty:
            attempted_path = CACHE_DIR / "archhub_national_attempted_pairs_latest.csv"
            attempted_df.to_csv(attempted_path, index=False, encoding="utf-8-sig")
            print(f"수집 시도 이력은 저장했습니다: {attempted_path}")
        return

    print("\n" + "=" * 70)
    print("원천 컬럼 목록")
    print("=" * 70)
    for col in raw_df.columns:
        print("-", col)

    norm_df = normalize_archhub_basis(raw_df)
    norm_df = add_archhub_demand_score(norm_df)
    region_summary_df = make_archhub_region_summary(norm_df)
    sigungu_summary_df = make_archhub_sigungu_summary(norm_df)

    files, meta_path = save_outputs(
        raw_df=raw_df,
        norm_df=norm_df,
        region_summary_df=region_summary_df,
        sigungu_summary_df=sigungu_summary_df,
        attempted_df=attempted_df,
        ref_df=ref_df,
        run_mode=args.mode,
    )

    print("\n" + "=" * 70)
    print("건축HUB 수집 완료")
    print("=" * 70)
    print("원천 데이터 건수:", len(raw_df))
    print("정규화 데이터 건수:", len(norm_df))
    print("지역 요약 건수:", len(region_summary_df))
    print("시군구 요약 건수:", len(sigungu_summary_df))

    print("\n저장 파일:")
    for key, value in files.items():
        print(f"{key}: {value}")
    print("meta:", meta_path)

    print("\n" + "=" * 70)
    print("지역별 건축허가 선행수요 요약")
    print("=" * 70)
    if not region_summary_df.empty:
        print(region_summary_df.to_string(index=False))
    else:
        print("요약 데이터가 없습니다.")

    print("\n" + "=" * 70)
    print("시군구별 TOP 20")
    print("=" * 70)
    if not sigungu_summary_df.empty:
        print(sigungu_summary_df.head(20).to_string(index=False))
    else:
        print("요약 데이터가 없습니다.")


if __name__ == "__main__":
    main()
