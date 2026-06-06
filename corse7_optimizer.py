# =========================================================
# AI 집배순로 최적화 최종 완성본
# =========================================================

import os
import glob
import time
import pickle
import pandas as pd
import requests
import folium

from dotenv import load_dotenv
import os

load_dotenv()

from math import radians, sin, cos, sqrt, atan2
from functools import lru_cache

from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

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

OUTPUT_EXCEL = "output_result.xlsx"

GEOCODE_CACHE_FILE = "geocode_cache.pkl"
ROAD_CACHE_FILE = "road_cache.pkl"

# =========================================================
# 지도 설정
# =========================================================

CREATE_MAP = True
MAP_FOLDER = "maps"

# =========================================================
# 설정
# =========================================================

USE_CACHE = True
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

if USE_CACHE and os.path.exists(GEOCODE_CACHE_FILE):
    with open(GEOCODE_CACHE_FILE, "rb") as f:
        geocode_cache = pickle.load(f)
else:
    geocode_cache = {}

if USE_CACHE and os.path.exists(ROAD_CACHE_FILE):
    with open(ROAD_CACHE_FILE, "rb") as f:
        road_cache = pickle.load(f)
else:
    road_cache = {}

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

    except:
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

    except:
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

    except:
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

        except:
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

    url = "https://maps.apigw.ntruss.com/map-geocode/v2/geocode"

    try:

        time.sleep(0.05)

        response = SESSION.get(
            url,
            params={"query": addr},
            timeout=30
        )

        if response.status_code != 200:
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

    if start_xy is None or end_xy is None:
        return 0, 0

    key = (start_xy, end_xy)

    if key in road_cache:
        return road_cache[key]

    url = "https://maps.apigw.ntruss.com/map-direction/v1/driving"

    params = {
        "start": f"{start_xy[0]},{start_xy[1]}",
        "goal": f"{end_xy[0]},{end_xy[1]}",
        "option": "traoptimal"
    }

    try:

        time.sleep(0.05)

        response = SESSION.get(
            url,
            params=params,
            timeout=30
        )

        if response.status_code != 200:
            return 0, 0

        res = response.json()

        route = res.get("route")

        if route is None:
            return 0, 0

        if "traoptimal" not in route:
            return 0, 0

        traoptimal = route["traoptimal"]

        if len(traoptimal) == 0:
            return 0, 0

        summary = traoptimal[0]["summary"]

        distance = summary.get("distance", 0)
        duration = summary.get("duration", 0)

        result = (
            distance,
            duration
        )

        road_cache[key] = result

        return result

    except Exception as e:

        print("[도로거리 실패]", e)

        return 0, 0

# =========================================================
# Hybrid 최적화 (계산한 도로거리 캐시로 재사용)
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
            key=lambda x: haversine_distance(current, x)
        )[:candidate_count]

        best = None
        best_time = INF

        for cand in candidates:
            d, t = get_road_distance(current, cand)
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
            d, t = get_road_distance(prev_xy, start_xy)
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

        d, t = get_road_distance(prev_xy, curr_xy)
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
    d, t = get_road_distance(prev_xy, start_xy)
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
    total_addr = len(unique_addresses)

    for i, addr in enumerate(unique_addresses):
        pct = 5 + int((i + 1) / total_addr * 30)
        report(pct, f"좌표 변환 중... ({i+1}/{total_addr})")
        coord = geocode(addr)
        if coord is not None:
            addr2coord[addr] = coord

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

    saved_distance_m = d1 - d2
    saved_time_ms = t1 - t2

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
    # 캐시 저장
    # =====================================================

    if USE_CACHE:

        with open(
            GEOCODE_CACHE_FILE,
            "wb"
        ) as f:

            pickle.dump(
                geocode_cache,
                f
            )

        with open(
            ROAD_CACHE_FILE,
            "wb"
        ) as f:

            pickle.dump(
                road_cache,
                f
            )

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
    "compare_maps": compare_map_paths
    }

# =========================================================
# 실행
# =========================================================

if __name__ == "__main__":

    print("프로그램 시작")

    process_excel()