# =========================================================
# 공유 캐시 관리자 (cache_manager.py)
# corse7_optimizer.py / corse8_optimizer.py 공용
#
# ▶ 캐시 전략
#   - geocode_cache : 주소 → (경도, 위도) 좌표 변환 결과
#   - road_cache    : (출발좌표, 도착좌표) → (거리m, 시간ms) 경로 결과
#   - API 호출 후 즉시 파일에 저장 → 중간 종료 시에도 캐시 유지
#   - corse7 / corse8 동일 pkl 파일 공유 → 중복 API 호출 방지
# =========================================================

import os
import pickle
import threading

GEOCODE_CACHE_FILE = "geocode_cache.pkl"
ROAD_CACHE_FILE    = "road_cache.pkl"

_lock = threading.Lock()

# ── 내부 캐시 딕셔너리 ────────────────────────────────────

_geocode_cache: dict = {}
_road_cache: dict    = {}

# ── 통계 카운터 ──────────────────────────────────────────

_stats = {
    "geocode_hit":  0,
    "geocode_miss": 0,
    "road_hit":     0,
    "road_miss":    0,
}

# =========================================================
# 초기 로드
# =========================================================

def load_caches():
    """프로그램 시작 시 1회 호출 — pkl 파일에서 캐시 로드"""
    global _geocode_cache, _road_cache

    if os.path.exists(GEOCODE_CACHE_FILE):
        try:
            with open(GEOCODE_CACHE_FILE, "rb") as f:
                _geocode_cache = pickle.load(f)
            print(f"[캐시] geocode 캐시 로드: {len(_geocode_cache)}건")
        except Exception as e:
            print(f"[캐시] geocode 캐시 로드 실패 (새로 시작): {e}")
            _geocode_cache = {}
    else:
        _geocode_cache = {}

    if os.path.exists(ROAD_CACHE_FILE):
        try:
            with open(ROAD_CACHE_FILE, "rb") as f:
                _road_cache = pickle.load(f)
            print(f"[캐시] road 캐시 로드: {len(_road_cache)}건")
        except Exception as e:
            print(f"[캐시] road 캐시 로드 실패 (새로 시작): {e}")
            _road_cache = {}
    else:
        _road_cache = {}


def _save_geocode():
    """geocode 캐시를 파일에 즉시 저장 (thread-safe)"""
    with _lock:
        try:
            with open(GEOCODE_CACHE_FILE, "wb") as f:
                pickle.dump(_geocode_cache, f)
        except Exception as e:
            print(f"[캐시] geocode 저장 실패: {e}")


def _save_road():
    """road 캐시를 파일에 즉시 저장 (thread-safe)"""
    with _lock:
        try:
            with open(ROAD_CACHE_FILE, "wb") as f:
                pickle.dump(_road_cache, f)
        except Exception as e:
            print(f"[캐시] road 저장 실패: {e}")


# =========================================================
# Geocode 캐시 인터페이스
# =========================================================

def get_geocode(addr: str):
    """캐시에서 좌표 반환. 없으면 None."""
    if addr in _geocode_cache:
        _stats["geocode_hit"] += 1
        return _geocode_cache[addr]
    _stats["geocode_miss"] += 1
    return None


def set_geocode(addr: str, coord: tuple):
    """좌표를 캐시에 저장하고 즉시 파일에 기록."""
    _geocode_cache[addr] = coord
    _save_geocode()


# =========================================================
# Road 캐시 인터페이스
# =========================================================

def get_road(start_xy: tuple, end_xy: tuple):
    """캐시에서 경로(거리, 시간) 반환. 없으면 None."""
    key = (start_xy, end_xy)
    if key in _road_cache:
        _stats["road_hit"] += 1
        return _road_cache[key]
    _stats["road_miss"] += 1
    return None


def set_road(start_xy: tuple, end_xy: tuple, result: tuple):
    """경로 결과를 캐시에 저장하고 즉시 파일에 기록."""
    key = (start_xy, end_xy)
    _road_cache[key] = result
    _save_road()


# =========================================================
# 통계 / 정보
# =========================================================

def get_stats() -> dict:
    """현재 세션의 API 호출 통계 반환."""
    total_geo  = _stats["geocode_hit"]  + _stats["geocode_miss"]
    total_road = _stats["road_hit"]     + _stats["road_miss"]

    geo_hit_rate  = (_stats["geocode_hit"]  / total_geo  * 100) if total_geo  else 0
    road_hit_rate = (_stats["road_hit"]     / total_road * 100) if total_road else 0

    return {
        "geocode_cache_size":  len(_geocode_cache),
        "road_cache_size":     len(_road_cache),
        "geocode_hit":         _stats["geocode_hit"],
        "geocode_miss":        _stats["geocode_miss"],
        "geocode_hit_rate":    round(geo_hit_rate, 1),
        "road_hit":            _stats["road_hit"],
        "road_miss":           _stats["road_miss"],
        "road_hit_rate":       round(road_hit_rate, 1),
        "api_calls_saved":     _stats["geocode_hit"] + _stats["road_hit"],
        "api_calls_made":      _stats["geocode_miss"] + _stats["road_miss"],
    }


def reset_stats():
    """세션 통계 초기화 (캐시 내용은 유지)."""
    for k in _stats:
        _stats[k] = 0


def clear_all_caches():
    """캐시 전체 삭제 (파일 포함). 주소가 대폭 바뀔 때 사용."""
    global _geocode_cache, _road_cache
    _geocode_cache = {}
    _road_cache    = {}
    for path in (GEOCODE_CACHE_FILE, ROAD_CACHE_FILE):
        try:
            if os.path.exists(path):
                os.remove(path)
        except Exception as e:
            print(f"[캐시] 삭제 실패 {path}: {e}")
    reset_stats()
    print("[캐시] 전체 캐시 삭제 완료")


# =========================================================
# 모듈 임포트 시 자동 로드
# =========================================================

load_caches()
