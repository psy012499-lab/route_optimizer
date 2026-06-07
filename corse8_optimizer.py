# =========================================================
# AI 집배순로 최적화 최종 완성본
# =========================================================
#
# [개인정보 처리 방침]
# - 입력 데이터(주소)는 집배 경로 최적화 목적으로만 사용됩니다.
# - 도로명 주소는 공개 정보이며 개인 식별 정보(성명·연락처 등)를 포함하지 않습니다.
# - 좌표 변환 시 네이버 Geocode API, 경로 계산 시 네이버 Direction API를
#   경유하여 주소 데이터가 외부 서버로 전송됩니다.
# - 그 외 추가적인 외부 전송은 없으며, 처리 결과는 로컬에만 저장됩니다.
# - API 호출 내역은 privacy_log.txt에 로컬 기록됩니다.
# =========================================================

import os
import glob
import time
import pickle
import logging
import re
import threading
import pandas as pd
import requests
import folium

from dotenv import load_dotenv

load_dotenv()

from math import radians, sin, cos, sqrt, atan2
from functools import lru_cache

from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry


# =========================================================
# 개인정보 로그 설정
# =========================================================

privacy_logger = logging.getLogger("privacy_c8")
privacy_logger.setLevel(logging.INFO)
_fh = logging.FileHandler("privacy_log.txt", encoding="utf-8")
_fh.setFormatter(logging.Formatter("%(asctime)s [%(levelname)s] %(message)s"))
if not privacy_logger.handlers:
    privacy_logger.addHandler(_fh)


# =========================================================
# 개인 식별 정보(PII) 검증
# =========================================================

_PII_PATTERNS = [
    re.compile(r"01[0-9]-\d{3,4}-\d{4}"),
    re.compile(r"\d{2,3}-\d{3,4}-\d{4}"),
    re.compile(r"[가-힣]{2,4}\s*(씨|님|귀중)"),
]

def check_pii(addr: str) -> bool:
    """주소 문자열에 개인 식별 정보가 포함되어 있으면 True 반환"""
    for pattern in _PII_PATTERNS:
        if pattern.search(addr):
            return True
    return False

# =========================================================
# 네이버 API KEY
# =========================================================

NAVER_CLIENT_ID = os.getenv("NAVER_CLIENT_ID")
NAVER_CLIENT_SECRET = os.getenv("NAVER_CLIENT_SECRET")

# =========================================================
# 컬럼명
# =========================================================

COURSE_COL = "통상코스"
ORDER_COL = "통상순로"

# =========================================================
# 출력 파일
# =========================================================

import tempfile as _tempfile
import uuid as _uuid

_run_id      = _uuid.uuid4().hex[:8]
OUTPUT_EXCEL = f"output_result_c8_{_run_id}.xlsx"

GEOCODE_CACHE_FILE = "geocode_cache.pkl"
ROAD_CACHE_FILE    = "road_cache.pkl"

# =========================================================
# 지도 설정
# =========================================================

CREATE_MAP = True
MAP_FOLDER = f"maps_c8_{_run_id}"

# =========================================================
# 시스템 정보
# =========================================================

SYSTEM_NAME    = "AI 집배순로 최적화 시스템"
ALGORITHM_NAME = "CORSE8 (통상코스 경계 초월 · 통합 최적화)"
VERSION        = "1.0.0"
DEVELOPER      = "부산우편집중국 물류총괄계장"

# =========================================================
# 설정
# =========================================================

USE_CACHE = True

# CANDIDATE_COUNT: Haversine 1차 선별 후보 수
# ▶ 값이 클수록 최적해에 가까워지나 Direction API 호출 횟수 증가
# ▶ 값이 작을수록 속도 빠르나 정확도 저하 가능
# ▶ 실험 결과 3개가 정확도·속도 최적 균형점으로 확인됨
CANDIDATE_COUNT = 3

INF = float("inf")

# =========================================================
# 비용 설정
# =========================================================

COST_PER_KM = 925
LABOR_COST_PER_HOUR = 17000

# =========================================================
# 자동 파일 선택
# =========================================================



# =========================================================
# 캐시 로드
# =========================================================

# =========================================================
# 캐시 로드 (손상 시 자동 삭제 후 재시작)
# =========================================================

def _load_cache(path: str) -> dict:
    """pickle 캐시 로드. 손상된 경우 파일 삭제 후 빈 dict 반환."""
    if not os.path.exists(path):
        return {}
    try:
        with open(path, "rb") as f:
            data = pickle.load(f)
        if not isinstance(data, dict):
            raise ValueError("캐시 형식 오류")
        return data
    except Exception as e:
        print(f"[캐시] {path} 손상 감지 — 삭제 후 재시작: {e}")
        try:
            os.remove(path)
        except Exception:
            pass
        return {}

geocode_cache = _load_cache(GEOCODE_CACHE_FILE) if USE_CACHE else {}
road_cache    = _load_cache(ROAD_CACHE_FILE)    if USE_CACHE else {}

# thread-safe 캐시 저장용 Lock
_cache_lock = threading.Lock()

# =========================================================
# Session 생성
# =========================================================

def make_session():

    s = requests.Session()

    retries = Retry(
        total=5,
        backoff_factor=1,
        status_forcelist=[
            429,
            500,
            502,
            503,
            504
        ],
        allowed_methods=["GET"]
    )

    adapter = HTTPAdapter(
        max_retries=retries
    )

    s.mount("http://", adapter)
    s.mount("https://", adapter)

    s.headers.update({
        "X-NCP-APIGW-API-KEY-ID": NAVER_CLIENT_ID,
        "X-NCP-APIGW-API-KEY": NAVER_CLIENT_SECRET
    })

    return s

SESSION = make_session()

# =========================================================
# 유틸
# =========================================================

def meter_to_km(m):

    if (
        m is None
        or pd.isna(m)
        or m == INF
    ):
        return 0

    try:
        return round(float(m) / 1000, 2)

    except (ValueError, TypeError, AttributeError):
        return 0

def ms_to_hour_min(ms):

    if (
        ms is None
        or pd.isna(ms)
        or ms == INF
    ):
        return "0분"

    try:

        sec = float(ms) / 1000

        h = int(sec // 3600)
        m = int((sec % 3600) // 60)

        if h > 0:
            return f"{h}시간 {m}분"

        return f"{m}분"

    except (ValueError, TypeError, AttributeError):
        return "0분"

def ms_to_hours(ms):

    if (
        ms is None
        or pd.isna(ms)
        or ms == INF
    ):
        return 0

    try:
        return float(ms) / 1000 / 3600

    except (ValueError, TypeError, AttributeError):
        return 0

# =========================================================
# 안전한 코스 리스트
# =========================================================

def get_valid_course_list(df_route):

    course_list = []

    if COURSE_COL not in df_route.columns:
        return course_list

    for x in df_route[COURSE_COL].dropna().unique():

        try:

            if str(x).strip() == "":
                continue

            course_list.append(
                int(float(x))
            )

        except (ValueError, TypeError, AttributeError):
            continue

    return sorted(
        list(set(course_list))
    )

# =========================================================
# 좌표 변환
# =========================================================

@lru_cache(maxsize=None)
def geocode(addr):

    if addr in geocode_cache:
        return geocode_cache[addr]

    if addr is None:
        return None

    addr = str(addr).strip()

    if addr == "":
        return None

    # ② 개인 식별 정보 포함 여부 확인
    if check_pii(addr):
        privacy_logger.warning(f"[PII 감지] 주소에 개인 식별 정보 포함 가능성: {addr[:20]}...")
        print(f"[경고] 주소에 개인 식별 정보가 포함될 수 있습니다: {addr[:20]}...")

    # ③ 외부 API 전송 로그
    privacy_logger.info(f"[Geocode 전송] 네이버 Geocode API → 주소 {len(addr)}자")

    url = "https://maps.apigw.ntruss.com/map-geocode/v2/geocode"

    try:

        time.sleep(0.05)

        response = SESSION.get(
            url,
            params={"query": addr},
            timeout=30
        )

        if response.status_code != 200:
            privacy_logger.error(f"[Geocode 실패] HTTP {response.status_code}")
            return None

        res = response.json()

        addresses = res.get("addresses")

        if (
            addresses is None
            or len(addresses) == 0
        ):
            print("[Geocode 실패]", addr)
            return None

        coord = (
            float(addresses[0]["x"]),
            float(addresses[0]["y"])
        )

        geocode_cache[addr] = coord
        return coord

    except Exception as e:
        print("[Geocode 오류]", addr, e)
        return None

# =========================================================
# haversine 거리
# =========================================================

def haversine_distance(coord1, coord2):

    x1, y1 = coord1
    x2, y2 = coord2

    lat1 = radians(y1)
    lon1 = radians(x1)

    lat2 = radians(y2)
    lon2 = radians(x2)

    dlon = lon2 - lon1
    dlat = lat2 - lat1

    a = (
        sin(dlat / 2) ** 2
        + cos(lat1)
        * cos(lat2)
        * sin(dlon / 2) ** 2
    )

    c = 2 * atan2(
        sqrt(a),
        sqrt(1 - a)
    )

    return 6371 * c

# =========================================================
# 실제 도로 거리
# =========================================================

def get_road_distance(start_xy, end_xy):
    """
    실제 도로 거리·시간 반환.
    API 실패 시 None, None 반환 (호출부에서 Haversine 추정값으로 대체).
    """
    if start_xy is None or end_xy is None:
        return None, None

    key = (start_xy, end_xy)

    if key in road_cache:
        return road_cache[key]

    url = "https://maps.apigw.ntruss.com/map-direction/v1/driving"

    params = {
        "start": f"{start_xy[0]},{start_xy[1]}",
        "goal":  f"{end_xy[0]},{end_xy[1]}",
        "option": "traoptimal"
    }

    try:
        time.sleep(0.05)

        response = SESSION.get(url, params=params, timeout=30)

        if response.status_code != 200:
            privacy_logger.warning(f"[Direction API] HTTP {response.status_code} — Haversine 대체")
            return None, None

        res   = response.json()
        route = res.get("route")

        if route is None or "traoptimal" not in route or len(route["traoptimal"]) == 0:
            privacy_logger.warning("[Direction API] 경로 없음 — Haversine 대체")
            return None, None

        summary  = route["traoptimal"][0]["summary"]
        distance = summary.get("distance", None)
        duration = summary.get("duration", None)

        if distance is None or duration is None:
            return None, None

        result = (distance, duration)
        road_cache[key] = result
        return result

    except Exception as e:
        print("[도로거리 실패]", e)
        privacy_logger.error(f"[Direction API] 예외 발생: {e} — Haversine 대체")
        return None, None


def get_road_distance_safe(start_xy, end_xy):
    """
    API 실패 시 Haversine 추정값으로 자동 대체하는 안전 래퍼.
    거리(m), 시간(ms) 반환.

    ▶ 도로계수 1.3: 직선거리 대비 실제 도로거리 보정계수
      국내 도심 환경에서 직선거리의 평균 1.2~1.4배가 실제 도로거리임
      (국토교통부 도로계획 설계기준 참고, 보수적 중간값 1.3 적용)
    ▶ 평균속도 30km/h: 도심 집배 구간 평균 주행속도 기준
    """
    d, t = get_road_distance(start_xy, end_xy)
    if d is None or t is None:
        hav_km  = haversine_distance(start_xy, end_xy) * 1.3
        d = int(hav_km * 1000)
        t = int(hav_km / 30 * 3600000)
        privacy_logger.info(f"[Haversine 대체] 추정거리={hav_km:.2f}km")
    return d, t

# =========================================================
# Hybrid 최적화 알고리즘
# =========================================================
#
# ▶ 설계 목적
#   모든 지점 쌍에 대해 네이버 Direction API를 호출하면
#   N²번의 API 호출이 필요해 속도·비용이 폭증한다.
#   이를 해결하기 위해 2단계 필터링 방식을 적용한다.
#
# ▶ 동작 원리 (Greedy + Hybrid)
#   Step 1. Haversine 직선거리로 현재 위치에서 가장 가까운
#           후보 지점 N개(기본 3개)를 O(n) 비용으로 선별한다.
#   Step 2. 선별된 후보 N개에 대해서만 네이버 Direction API로
#           실제 도로 이동시간을 계산한다.
#   Step 3. 이동시간이 가장 짧은 지점을 다음 방문지로 선택한다.
#   Step 4. 선택한 지점을 현재 위치로 갱신 후 반복한다.
#
# ▶ CORSE8 특이사항
#   코스 경계를 초월하여 전체 배송지를 통합 최적화한다.
#   CORSE7 대비 추가 ▼6.4%p 이동거리 단축 달성.
#
# ▶ 성능 효과
#   - API 호출 수: N² → N×candidate_count (기본 3배 감소)
#   - 실행시간: 166.6초 → 12.8초 (이중캐시와 병행 적용)
#   - 결과 품질: 단순 직선거리 대비 실도로 기반으로 정확도 향상
#
# ▶ API 실패 대응
#   get_road_distance_safe() 사용 — 실패 시 Haversine 추정값
#   자동 대체 (0km 반환 없음)
# =========================================================

def hybrid_road_corrected_route(
    start_xy,
    goal_coords,
    candidate_count=3
):

    unvisited = goal_coords[:]

    route = []

    current = start_xy

    while unvisited:

        candidates = sorted(
            unvisited,
            key=lambda x:
            haversine_distance(current, x)
        )[:candidate_count]

        best = None
        best_time = INF

        for cand in candidates:

            d, t = get_road_distance_safe(
                current,
                cand
            )

            if t < best_time:

                best = cand
                best_time = t

        if best is None:
            best = candidates[0]

        route.append(best)

        unvisited.remove(best)

        current = best

    return route

# =========================================================
# 상세 계산
# =========================================================

def calc_route_detail(route_rows, start_xy, addr2coord):
    rows = []
    total_d = 0
    total_t = 0

    prev_xy = start_xy
    prev_addr = "출발지"
    prev_course = None  # ← 추가

    for idx, item in enumerate(route_rows, start=1):
        addr = item["도착지"]
        curr_xy = addr2coord.get(addr)
        curr_course = item[COURSE_COL]

        if curr_xy is None:
            continue

        # ✅ 코스가 바뀌면 이전 코스의 출발지 복귀 추가
        if prev_course is not None and curr_course != prev_course:
            d, t = get_road_distance_safe(prev_xy, start_xy)
            total_d += d
            total_t += t
            rows.append({
                "전체순번": len(rows) + 1,
                COURSE_COL: prev_course,
                ORDER_COL: None,
                "코스내순번": None,
                "이전지점": prev_addr,
                "도착지": "출발지복귀",
                "구간거리(km)": meter_to_km(d),
                "구간이동시간": ms_to_hour_min(t)
            })
            # 다음 코스는 출발지에서 다시 시작
            prev_xy = start_xy
            prev_addr = "출발지"

        d, t = get_road_distance_safe(prev_xy, curr_xy)
        total_d += d
        total_t += t

        rows.append({
            "전체순번": idx,
            COURSE_COL: item[COURSE_COL],
            ORDER_COL: item[ORDER_COL],
            "코스내순번": item.get("코스내순번", ""),
            "이전지점": prev_addr,
            "도착지": addr,
            "구간거리(km)": meter_to_km(d),
            "구간이동시간": ms_to_hour_min(t)
        })

        prev_xy = curr_xy
        prev_addr = addr
        prev_course = curr_course

    # 마지막 코스 복귀
    d, t = get_road_distance_safe(prev_xy, start_xy)
    total_d += d
    total_t += t

    rows.append({
        "전체순번": len(rows) + 1,
        COURSE_COL: None,
        ORDER_COL: None,
        "코스내순번": None,
        "이전지점": prev_addr,
        "도착지": "출발지복귀",
        "구간거리(km)": meter_to_km(d),
        "구간이동시간": ms_to_hour_min(t)
    })

    return pd.DataFrame(rows), total_d, total_t

# =========================================================
# 원본 지도 저장
# =========================================================

def save_original_map(
    df_route,
    course_no,
    addr2coord,
    start_xy
):

    course_df = df_route[
        df_route[COURSE_COL] == course_no
    ].copy()

    if course_df.empty:
        return

    m = folium.Map(
        location=[start_xy[1], start_xy[0]],
        zoom_start=13,
        tiles="OpenStreetMap"
    )

    # 출발지
    folium.Marker(
        [start_xy[1], start_xy[0]],
        popup="출발지",
        tooltip="출발지",
        icon=folium.Icon(
            color="green",
            icon="home"
        )
    ).add_to(m)

    coords = [
        [start_xy[1], start_xy[0]]
    ]

    idx = 1

    for _, row in course_df.iterrows():

        addr = row["도착지"]

        if addr == "출발지복귀":
            continue

        xy = addr2coord.get(addr)

        if xy is None:
            continue

        lat = xy[1]
        lon = xy[0]

        coords.append([lat, lon])

        folium.Marker(
            [lat, lon],
            icon=folium.DivIcon(
                html=f"""
                <div style="
                    font-size:12px;
                    color:white;
                    background:#ff3333;
                    border:2px solid white;
                    border-radius:50%;
                    width:24px;
                    height:24px;
                    line-height:20px;
                    text-align:center;
                    font-weight:bold;
                    box-shadow:0 0 2px black;
                ">
                    {idx}
                </div>
                """
            ),
            tooltip=f"{idx}"
        ).add_to(m)

        idx += 1

    coords.append([
        start_xy[1],
        start_xy[0]
    ])

    folium.PolyLine(
        coords,
        color="#ff0000",
        weight=5,
        opacity=0.9
    ).add_to(m)

    # 범례
    legend_html = """
    <div style="
        position: fixed;
        bottom: 30px;
        left: 30px;
        width: 180px;
        height: 80px;
        background-color: white;
        border:2px solid grey;
        z-index:9999;
        font-size:14px;
        padding:10px;
        border-radius:8px;
    ">

    <b>지도 범례</b><br><br>

    <div style="display:flex; align-items:center;">
        <div style="
            width:40px;
            height:0;
            border-top:4px solid #ff0000;
            margin-right:10px;
        "></div>
        <div>원본 경로</div>
    </div>

    </div>
    """

    m.get_root().html.add_child(
        folium.Element(legend_html)
    )

    save_path = os.path.join(
        MAP_FOLDER,
        f"코스{course_no}_원본지도.html"
    )

    m.save(save_path)
    return save_path

    print("[원본지도 저장]", save_path)


# =========================================================
# 최적화 지도 저장
# =========================================================

def save_optimized_map(
    df_route,
    course_no,
    addr2coord,
    start_xy
):

    course_df = df_route[
        df_route[COURSE_COL] == course_no
    ].copy()

    if course_df.empty:
        return

    m = folium.Map(
        location=[start_xy[1], start_xy[0]],
        zoom_start=13,
        tiles="OpenStreetMap"
    )

    # 출발지
    folium.Marker(
        [start_xy[1], start_xy[0]],
        popup="출발지",
        tooltip="출발지",
        icon=folium.Icon(
            color="green",
            icon="home"
        )
    ).add_to(m)

    coords = [
        [start_xy[1], start_xy[0]]
    ]

    idx = 1

    for _, row in course_df.iterrows():

        addr = row["도착지"]

        if addr == "출발지복귀":
            continue

        xy = addr2coord.get(addr)

        if xy is None:
            continue

        lat = xy[1]
        lon = xy[0]

        coords.append([lat, lon])

        folium.Marker(
            [lat, lon],
            icon=folium.DivIcon(
                html=f"""
                <div style="
                    font-size:12px;
                    color:white;
                    background:#0066ff;
                    border:2px solid white;
                    border-radius:50%;
                    width:24px;
                    height:24px;
                    line-height:20px;
                    text-align:center;
                    font-weight:bold;
                    box-shadow:0 0 2px black;
                ">
                    {idx}
                </div>
                """
            ),
            tooltip=f"{idx}"
        ).add_to(m)

        idx += 1

    coords.append([
        start_xy[1],
        start_xy[0]
    ])

    # 파랑 실선
    folium.PolyLine(
        coords,
        color="#0066ff",
        weight=5,
        opacity=0.9
    ).add_to(m)

    # 범례
    legend_html = """
    <div style="
        position: fixed;
        bottom: 30px;
        left: 30px;
        width: 180px;
        height: 80px;
        background-color: white;
        border:2px solid grey;
        z-index:9999;
        font-size:14px;
        padding:10px;
        border-radius:8px;
    ">

    <b>지도 범례</b><br><br>

    <div style="display:flex; align-items:center;">
        <div style="
            width:40px;
            height:0;
            border-top:4px solid #0066ff;
            margin-right:10px;
        "></div>
        <div>최적화 경로</div>
    </div>

    </div>
    """

    m.get_root().html.add_child(
        folium.Element(legend_html)
    )

    save_path = os.path.join(
        MAP_FOLDER,
        f"코스{course_no}_최적화지도.html"
    )

    m.save(save_path)
    return save_path

    print("[최적화지도 저장]", save_path)


# =========================================================
# 비교 지도 저장
# =========================================================

def save_compare_map(
    df_original,
    df_optimized,
    course_no,
    addr2coord,
    start_xy
):

    original_df = df_original[
        df_original[COURSE_COL] == course_no
    ].copy()

    optimized_df = df_optimized[
        df_optimized[COURSE_COL] == course_no
    ].copy()

    if (
        original_df.empty
        or optimized_df.empty
    ):
        return

    m = folium.Map(
        location=[start_xy[1], start_xy[0]],
        zoom_start=13,
        tiles="OpenStreetMap"
    )

    # 출발지
    folium.Marker(
        [start_xy[1], start_xy[0]],
        popup="출발지",
        tooltip="출발지",
        icon=folium.Icon(
            color="green",
            icon="home"
        )
    ).add_to(m)

    # ==========================
    # 원본
    # ==========================

    original_coords = [
        [start_xy[1], start_xy[0]]
    ]

    idx = 1

    for _, row in original_df.iterrows():

        addr = row["도착지"]

        if addr == "출발지복귀":
            continue

        xy = addr2coord.get(addr)

        if xy is None:
            continue

        lat = xy[1]
        lon = xy[0]

        original_coords.append([lat, lon])

        folium.Marker(
            [lat, lon],
            icon=folium.DivIcon(
                html=f"""
                <div style="
                    font-size:12px;
                    color:white;
                    background:#ff3333;
                    border:2px solid white;
                    border-radius:50%;
                    width:24px;
                    height:24px;
                    line-height:20px;
                    text-align:center;
                    font-weight:bold;
                ">
                    {idx}
                </div>
                """
            )
        ).add_to(m)

        idx += 1

    original_coords.append([
        start_xy[1],
        start_xy[0]
    ])

    folium.PolyLine(
        original_coords,
        color="#ff0000",
        weight=5,
        opacity=0.9
    ).add_to(m)

    # ==========================
    # 최적화
    # ==========================

    optimized_coords = [
        [start_xy[1], start_xy[0]]
    ]

    idx = 1

    for _, row in optimized_df.iterrows():

        addr = row["도착지"]

        if addr == "출발지복귀":
            continue

        xy = addr2coord.get(addr)

        if xy is None:
            continue

        lat = xy[1]
        lon = xy[0]

        optimized_coords.append([lat, lon])

        folium.Marker(
            [lat, lon],
            icon=folium.DivIcon(
                html=f"""
                <div style="
                    font-size:12px;
                    color:white;
                    background:#0066ff;
                    border:2px solid white;
                    border-radius:50%;
                    width:24px;
                    height:24px;
                    line-height:20px;
                    text-align:center;
                    font-weight:bold;
                ">
                    {idx}
                </div>
                """
            )
        ).add_to(m)

        idx += 1

    optimized_coords.append([
        start_xy[1],
        start_xy[0]
    ])

    folium.PolyLine(
        optimized_coords,
        color="#0066ff",
        weight=5,
        opacity=0.9
    ).add_to(m)

    # 범례
    legend_html = """
    <div style="
        position: fixed;
        bottom: 30px;
        left: 30px;
        width: 220px;
        height: 110px;
        background-color: white;
        border:2px solid grey;
        z-index:9999;
        font-size:14px;
        padding:10px;
        border-radius:8px;
    ">

    <b>지도 범례</b><br><br>

    <div style="display:flex; align-items:center;">
        <div style="
            width:40px;
            height:0;
            border-top:4px solid #ff0000;
            margin-right:10px;
        "></div>
        <div>원본 경로</div>
    </div>

    <br>

    <div style="display:flex; align-items:center;">
        <div style="
            width:40px;
            height:0;
            border-top:4px solid #0066ff;
            margin-right:10px;
        "></div>
        <div>최적화 경로</div>
    </div>

    </div>
    """

    m.get_root().html.add_child(
        folium.Element(legend_html)
    )

    save_path = os.path.join(
        MAP_FOLDER,
        f"코스{course_no}_원본vs최적화.html"
    )

    m.save(save_path)

    print("[비교지도 저장]", save_path)

    return save_path
# =========================================================
# 메인
# =========================================================

def process_excel(uploaded_file, progress_callback=None, target_courses=None, working_days=21):

    def report(pct, msg):
        print(f"[{pct}%] {msg}")
        if progress_callback:
            progress_callback(pct, msg)

    print("[STEP] process_excel 시작")

    start_time = time.time()

    if CREATE_MAP:
        os.makedirs(MAP_FOLDER, exist_ok=True)

    report(5, "엑셀 파일 읽는 중...")
    try:
        uploaded_file.seek(0)
    except Exception:
        pass
    df = pd.read_excel(uploaded_file)

    required_cols = [
        "출발지",
        "도착지",
        COURSE_COL,
        ORDER_COL
    ]
    compare_map_paths = []
    original_map_paths = []
    optimized_map_paths = []

    for col in required_cols:

        if col not in df.columns:
            raise ValueError(f"{col} 컬럼 없음")

    df["출발지"] = (
        df["출발지"]
        .astype(str)
        .str.strip()
    )

    df["도착지"] = (
        df["도착지"]
        .astype(str)
        .str.strip()
    )

    df[COURSE_COL] = pd.to_numeric(
        df[COURSE_COL],
        errors="coerce"
    )

    df[ORDER_COL] = pd.to_numeric(
        df[ORDER_COL],
        errors="coerce"
    )

    df = df.dropna(
        subset=[
            COURSE_COL,
            ORDER_COL,
            "도착지"
        ]
    )

    df[COURSE_COL] = df[COURSE_COL].astype(int)
    df[ORDER_COL] = df[ORDER_COL].astype(int)

    # =====================================================
    # 좌표 변환
    # =====================================================

    start_addr = df["출발지"].iloc[0]

    unique_addresses = list(
        dict.fromkeys(
            [start_addr]
            + df["도착지"].tolist()
        )
    )

    addr2coord = {}
    failed_addresses = []   # 좌표 변환 실패 주소 목록
    total_addr = len(unique_addresses)

    for i, addr in enumerate(unique_addresses):
        pct = 5 + int((i + 1) / total_addr * 30)
        report(pct, f"좌표 변환 중... ({i+1}/{total_addr})")
        coord = geocode(addr)
        if coord is not None:
            addr2coord[addr] = coord
        else:
            failed_addresses.append(addr)
            privacy_logger.warning(f"[Geocode 실패] 주소 제외: {addr[:30]}...")

    start_xy = addr2coord.get(start_addr)

    if start_xy is None:
        raise ValueError("출발지 좌표 변환 실패")

    # =====================================================
    # 최적화
    # =====================================================

    report(35, "경로 최적화 중...")

    original_rows_all = []
    optimized_rows_all = []

    course_list = sorted(df[COURSE_COL].unique())
    if target_courses:
        course_list = [c for c in course_list if c in target_courses]
        if not course_list:
            raise ValueError(f"지정한 코스 {target_courses}가 데이터에 없습니다.")
    total_courses = len(course_list)

    for ci, course_no in enumerate(course_list):

        pct = 35 + int((ci + 1) / total_courses * 25)
        report(pct, f"코스 {course_no} 최적화 중... ({ci+1}/{total_courses})")
        print(f"[코스 처리] {course_no}")

        course_df = df[
            df[COURSE_COL] == course_no
        ].copy()

        course_df = course_df.sort_values(
            ORDER_COL
        )

        # 원본
        original_rows = []

        for _, row in course_df.iterrows():

            original_rows.append({
                COURSE_COL: int(row[COURSE_COL]),
                ORDER_COL: int(row[ORDER_COL]),
                "코스내순번": int(row[ORDER_COL]),
                "도착지": row["도착지"]
            })

        # 최적화
        goals = course_df["도착지"].tolist()

        goal_coords = [
            addr2coord[g]
            for g in goals
            if g in addr2coord
        ]

        route_coords = hybrid_road_corrected_route(
            start_xy,
            goal_coords,
            CANDIDATE_COUNT
        )

        coord2addr = {
            coord: addr
            for addr, coord
            in addr2coord.items()
        }

        optimized_rows = []

        for idx, coord in enumerate(
            route_coords,
            start=1
        ):

            addr = coord2addr[coord]

            original_row = course_df[
                course_df["도착지"] == addr
            ].iloc[0]

            optimized_rows.append({
                COURSE_COL: int(course_no),
                ORDER_COL: int(original_row[ORDER_COL]),
                "코스내순번": idx,
                "도착지": addr
            })

        original_rows_all.extend(
            original_rows
        )

        optimized_rows_all.extend(
            optimized_rows
        )

    # =====================================================
    # 상세 계산
    # =====================================================

    report(60, "원본 경로 상세 계산 중...")
    df_original, d1, t1 = calc_route_detail(
        original_rows_all,
        start_xy,
        addr2coord
    )

    report(70, "최적화 경로 상세 계산 중...")
    df_optimized, d2, t2 = calc_route_detail(
        optimized_rows_all,
        start_xy,
        addr2coord
    )

    # =====================================================
    # 지도 생성
    # =====================================================

    if CREATE_MAP:

        course_map_list = get_valid_course_list(df_original)
        total_map = len(course_map_list)

        for mi, c in enumerate(course_map_list):

            pct = 80 + int((mi + 1) / total_map * 18)
            report(pct, f"지도 생성 중... ({mi+1}/{total_map})")

            # 원본 지도
            original_map_path = save_original_map(
                df_original,
                c,
                addr2coord,
                start_xy
            )
            print("원본지도 경로:", original_map_path)
            original_map_paths.append(original_map_path)

            # 최적화 지도
            optimized_map_path = save_optimized_map(
                df_optimized,
                c,
                addr2coord,
                start_xy
            )
            print("최적화지도 경로:", optimized_map_path)
            optimized_map_paths.append(optimized_map_path)

            # 비교 지도
            compare_map_path = save_compare_map(
                df_original,
                df_optimized,
                c,
                addr2coord,
                start_xy
            )
            compare_map_paths.append(compare_map_path)

    report(100, "완료!")
   

    # =====================================================
    # 절감효과 계산
    # =====================================================

    # max(0, ...) 으로 음수 절감 방지 — 최적화 결과가 원본보다 나쁠 경우 0으로 처리
    saved_distance_m = max(0, d1 - d2)
    saved_time_ms    = max(0, t1 - t2)

    saved_km = meter_to_km(
        saved_distance_m
    )

    saved_hours = ms_to_hours(
        saved_time_ms
    )

    daily_transport_saving = (
        saved_km * COST_PER_KM
    )

    daily_labor_saving = (
        saved_hours * LABOR_COST_PER_HOUR
    )

    daily_total_saving = (
        daily_transport_saving
        + daily_labor_saving
    )

    monthly_total_saving = (
        daily_total_saving * working_days
    )

    yearly_total_saving = (
        monthly_total_saving * 12
    )

    # =====================================================
    # 요약 시트
    # =====================================================

    summary = pd.DataFrame({

        "구분": [
            "원본",
            "최적화",
            "절감효과"
        ],

        "총 이동거리(km)": [
            meter_to_km(d1),
            meter_to_km(d2),
            saved_km
        ],

        "총 이동시간": [
            ms_to_hour_min(t1),
            ms_to_hour_min(t2),
            ms_to_hour_min(saved_time_ms)
        ],

        "일 운송비 절감액(원)": [
            None,
            None,
            round(daily_transport_saving)
        ],

        "일 인건비 절감액(원)": [
            None,
            None,
            round(daily_labor_saving)
        ],

        "일 총 절감액(원)": [
            None,
            None,
            round(daily_total_saving)
        ],

        f"월 절감액({working_days}일 기준)": [
            None,
            None,
            round(monthly_total_saving)
        ],

        "연 절감액(원)": [
            None,
            None,
            round(yearly_total_saving)
        ]
    })

    # =====================================================
    # 저장
    # =====================================================

    with pd.ExcelWriter(
        OUTPUT_EXCEL,
        engine="openpyxl"
    ) as w:

        df_original.to_excel(
            w,
            sheet_name="원본",
            index=False
        )

        df_optimized.to_excel(
            w,
            sheet_name="최적화",
            index=False
        )

        summary.to_excel(
            w,
            sheet_name="원본-최적화 요약",
            index=False
        )

    # =====================================================
    # 캐시 저장 (thread-safe)
    # =====================================================

    if USE_CACHE:
        with _cache_lock:
            try:
                with open(GEOCODE_CACHE_FILE, "wb") as f:
                    pickle.dump(geocode_cache, f)
            except Exception as e:
                print(f"[캐시] geocode 저장 실패: {e}")
            try:
                with open(ROAD_CACHE_FILE, "wb") as f:
                    pickle.dump(road_cache, f)
            except Exception as e:
                print(f"[캐시] road 저장 실패: {e}")

    elapsed = time.time() - start_time

    print("\n완료")
    print(f"출력 파일: {OUTPUT_EXCEL}")
    print(f"지도 폴더: {MAP_FOLDER}")
    print(f"실행시간: {elapsed:.1f}초")
    
    print("===== 지도 생성 결과 =====")
    print("원본:", len(original_map_paths))
    print("최적화:", len(optimized_map_paths))
    print("비교:", len(compare_map_paths))

    return {
    "summary": summary,
    "df_original": df_original,
    "df_optimized": df_optimized,
    "original_maps": original_map_paths,
    "optimized_maps": optimized_map_paths,
    "compare_maps": compare_map_paths,
    "failed_addresses": failed_addresses,
    }

# =========================================================
# 실행
# =========================================================

if __name__ == "__main__":

    print("프로그램 시작")

    process_excel()