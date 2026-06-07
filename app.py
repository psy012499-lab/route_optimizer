# =========================================================
# AI 집배순로 최적화 시스템 — app.py
# =========================================================
#
# ▶ 시스템 개요
#   집배원의 통상코스·통상순로 데이터를 입력받아
#   네이버 Maps AI API + Haversine 하이브리드 알고리즘으로
#   실도로 기반 최적 순로를 자동 산출하는 웹 기반 업무개선 시스템
#
# ▶ 주요 구성
#   - app.py            : Streamlit 웹 UI · 사용자 입출력 관리
#   - corse7_optimizer  : 통상코스 유지 · 코스 내 순로 최적화
#   - corse8_optimizer  : 통상코스 경계 초월 · 전체 통합 최적화
#   - cache_manager     : geocode · 도로거리 캐시 통합 관리
#
# ▶ AI 활용
#   - 네이버 Geocode API  : 주소 → 좌표 변환
#   - 네이버 Direction API: 실도로 거리·시간 계산
#   - Claude AI (Haiku)   : 최적화 결과 자연어 해석
#
# ▶ 개발자: 부산우편집중국 물류총괄계장
# ▶ 버전: 1.0.0
# =========================================================

import streamlit as st
import io
import os
import zipfile
import pandas as pd
import requests as _requests
from dotenv import load_dotenv

load_dotenv()

from corse7_optimizer import process_excel as process7
from corse8_optimizer import process_excel as process8


# =====================================================
# 설정 상수
# =====================================================

ANTHROPIC_API_URL   = "https://api.anthropic.com/v1/messages"
HAIKU_MODEL         = "claude-haiku-4-5-20251001"

DEFAULT_COST_PER_KM         = 925
DEFAULT_LABOR_COST_PER_HOUR = 17000
DEFAULT_WORKING_DAYS        = 21


# =====================================================
# Claude AI — 결과 자연어 해석
# =====================================================

def build_summary_text(summary: pd.DataFrame, algorithm_name: str,
                        cost_per_km: int, labor_cost: int, working_days: int) -> str:
    try:
        lines = [
            f"알고리즘: {algorithm_name}",
            f"운송비 단가: {cost_per_km}원/km",
            f"인건비 단가: {labor_cost}원/시간",
            f"월 근무일: {working_days}일",
        ]
        for col in summary.columns:
            if col == "구분":
                continue
            orig_val  = summary[summary["구분"] == "원본"][col].values
            opt_val   = summary[summary["구분"] == "최적화"][col].values
            saved_row = summary[summary["구분"] == "절감효과"]
            saved_val = saved_row[col].values if not saved_row.empty else summary.iloc[-1:][col].values
            if len(orig_val)  > 0 and str(orig_val[0])  not in ("None", "nan", ""):
                lines.append(f"  원본 {col}: {orig_val[0]}")
            if len(opt_val)   > 0 and str(opt_val[0])   not in ("None", "nan", ""):
                lines.append(f"  최적화 {col}: {opt_val[0]}")
            if len(saved_val) > 0 and str(saved_val[0]) not in ("None", "nan", ""):
                lines.append(f"  절감 {col}: {saved_val[0]}")
        return "\n".join(lines)
    except Exception:
        return summary.to_string(index=False)


def call_claude_interpretation(summary_text: str, algorithm_name: str) -> str:
    """
    Claude AI를 활용한 최적화 결과 자연어 해석 함수.

    ▶ 프롬프트 설계 의도
      단순 수치 나열 대신 현장 관리자의 의사결정 흐름에 맞춰
      3단계 구조로 설계:
      1) [최적화 결과 요약]  — 핵심 수치를 2~3문장으로 압축
      2) [절감 효과 분석]   — 비용·시간의 현업 의미 해석
      3) [현장 적용 권고]   — 실무 적용 시 고려사항 안내
      → 관리자가 보고서 없이 즉시 의사결정에 활용 가능하도록 설계

    ▶ 모델 선택: claude-haiku (빠른 응답 · 비용 최소화)
    """
    api_key = os.getenv("ANTHROPIC_API_KEY", "").strip()
    if not api_key:
        return None

    system = (
        "당신은 우정사업본부 물류 전문 AI 분석관입니다. "
        "집배 경로 최적화 결과를 집배원과 관리자가 쉽게 이해할 수 있도록 "
        "간결하고 명확한 한국어로 해석해 주세요. "
        "전문 용어는 풀어서 설명하고, 수치는 구체적으로 언급하며, "
        "현업 적용 시 기대되는 효과를 강조해 주세요.\n\n"
        "응답은 반드시 아래 형식을 따르세요:\n\n"
        "**[최적화 결과 요약]**\n(2~3문장으로 핵심 수치 요약)\n\n"
        "**[절감 효과 분석]**\n(비용·시간 절감 의미를 현업 관점에서 설명)\n\n"
        "**[현장 적용 권고]**\n(실무 적용 시 고려사항 또는 기대 효과 1~2가지)"
    )
    user = (
        f"다음은 AI 집배순로 최적화 시스템({algorithm_name})의 실행 결과입니다.\n\n"
        f"{summary_text}\n\n"
        "이 결과를 분석하여 우편집중국 담당자가 이해하기 쉽게 해석해 주세요."
    )
    headers = {
        "Content-Type": "application/json",
        "anthropic-version": "2023-06-01",
        "x-api-key": api_key,
    }
    payload = {
        "model": HAIKU_MODEL,
        "max_tokens": 800,
        "system": system,
        "messages": [{"role": "user", "content": user}],
    }
    try:
        resp = _requests.post(ANTHROPIC_API_URL, headers=headers, json=payload, timeout=30)
        if resp.status_code != 200:
            return f"[오류] {resp.status_code} — {resp.json().get('error', {}).get('message', resp.text)}"
        return resp.json()["content"][0]["text"]
    except Exception as e:
        return f"[오류] {e}"


# =====================================================
# 페이지 설정
# =====================================================

st.set_page_config(
    page_title="AI 집배순로 최적화",
    page_icon="🚚",
    layout="centered",
    initial_sidebar_state="collapsed",
)

st.markdown("""
<style>
html, body, [class*="css"] { font-size: 15px; }
h1 { font-size: 1.5rem !important; }
h2 { font-size: 1.2rem !important; }
h3 { font-size: 1.05rem !important; }

.metric-grid {
    display: grid;
    grid-template-columns: repeat(2, 1fr);
    gap: 8px;
    margin: 8px 0 16px;
}
.metric-card {
    background: #f0f2f6;
    border-radius: 10px;
    padding: 10px 12px;
    text-align: center;
}
.metric-label { font-size: 11px; color: #666; margin-bottom: 4px; }
.metric-value { font-size: 1.35rem; font-weight: 700; color: #1C2B4A; }
.metric-value.highlight { color: #E67E22; }
.metric-value.green     { color: #2e7d32; }

.stDataFrame { overflow-x: auto !important; }
.stDataFrame table { font-size: 12px !important; }
.stDataFrame th, .stDataFrame td { padding: 4px 8px !important; white-space: nowrap; }
.stButton > button { min-height: 48px; font-size: 15px; }
.stDownloadButton > button { min-height: 44px; font-size: 14px; }
.stTabs [data-baseweb="tab"] { font-size: 13px; padding: 6px 12px; }

.param-box {
    background: #f8f4ff;
    border-left: 4px solid #7F77DD;
    border-radius: 0 8px 8px 0;
    padding: 12px 16px;
    margin: 8px 0 12px;
    font-size: 13px;
    line-height: 1.7;
}
.param-tag {
    display: inline-block;
    background: #EEEDFE;
    color: #3C3489;
    border-radius: 6px;
    padding: 2px 8px;
    font-size: 12px;
    font-weight: 600;
    margin: 2px 3px;
}
.ai-box {
    background: #f0f7ff;
    border-left: 4px solid #1C6EBF;
    border-radius: 0 8px 8px 0;
    padding: 16px 18px;
    margin: 8px 0 16px;
    font-size: 14px;
    line-height: 1.8;
    color: #042C53;
}
</style>
""", unsafe_allow_html=True)


# =====================================================
# 샘플 데이터 생성
# =====================================================

@st.cache_data(show_spinner=False)
def make_sample_excel() -> io.BytesIO:
    # 부산 동래구 실측 집배 데이터 — 10코스·143건
    import os

    def is_bad(addr):
        if not isinstance(addr, str): return True
        addr = addr.strip()
        if addr in ["-", "", "nan"]: return True
        if "  " in addr: return True
        return False

    try:
        sample_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "input(시내통상구).xlsx")
        df = pd.read_excel(sample_path)
        df["출발지"] = df["출발지"].ffill()
        df = df[~df["도착지"].apply(is_bad)].copy()
        df["도착지"] = df["도착지"].str.strip()
        df = df[["출발지", "도착지", "통상코스", "통상순로"]].reset_index(drop=True)
    except Exception:
        # 파일 없을 경우 최소 백업 데이터
        DEPART = "부산광역시 동래구 명륜로 169"
        rows = [
            (DEPART, "부산광역시 동래구 충렬대로181번길 109", 1, 1),
            (DEPART, "부산광역시 동래구 동래로 109",          1, 2),
            (DEPART, "부산광역시 동래구 명륜로 165",          5, 1),
            (DEPART, "부산광역시 동래구 명륜로 163",          5, 2),
        ]
        df = pd.DataFrame(rows, columns=["출발지", "도착지", "통상코스", "통상순로"])

    n_courses = int(df["통상코스"].nunique())
    n_rows    = len(df)
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Sheet1")
    buf.seek(0)
    buf.name = "sample_input.xlsx"
    buf.n_courses = n_courses
    buf.n_rows    = n_rows
    return buf


# =====================================================
# 요약 카드 렌더링
# =====================================================

def render_summary_cards(summary_df: pd.DataFrame, working_days: int = 21):
    try:
        orig_row = summary_df[summary_df["구분"] == "원본"]
        opt_row  = summary_df[summary_df["구분"] == "최적화"]
        sav_row  = summary_df[summary_df["구분"] == "절감효과"]

        if orig_row.empty or opt_row.empty or sav_row.empty:
            return False

        def fv(row, col):
            if col not in row.columns: return None
            try: return float(str(row[col].values[0]).replace(",", ""))
            except: return None

        def find_col(df, keyword):
            """컬럼명에 키워드가 포함된 첫 번째 컬럼 반환"""
            cols = [c for c in df.columns if keyword in c]
            return cols[0] if cols else None

        def fmt_money(v):
            if v is None: return "-"
            if v >= 100_000_000: return f"{v/100_000_000:.1f}억원"
            if v >= 1_000_000:   return f"{v/1_000_000:.1f}백만원"
            if v >= 1_000:       return f"{v/1000:.1f}천원"
            return f"{v:,.0f}원"

        def hhmm_to_min(val):
            """'3시간 7분' 또는 '45분' 형태 문자열 → 분(int)"""
            try:
                val = str(val).strip()
                import re
                h = re.search(r"(\d+)시간", val)
                m = re.search(r"(\d+)분",  val)
                total = 0
                if h: total += int(h.group(1)) * 60
                if m: total += int(m.group(1))
                return total if total > 0 else None
            except: return None

        # ── 이동거리 절감률 & 절감량 ──────────────────────
        dist_col = "총 이동거리(km)"
        orig_d = fv(orig_row, dist_col)
        opt_d  = fv(opt_row,  dist_col)
        save_d = (orig_d - opt_d) if (orig_d and opt_d) else None
        rate   = round(save_d / orig_d * 100, 1) if (save_d and orig_d) else None

        # ── 이동시간 단축 — 절감효과 행에서 직접 읽기 ──────────
        time_col = "총 이동시간"
        save_t = None
        if time_col in summary_df.columns:
            # 절감효과 행의 이동시간 직접 사용
            sav_t_raw = sav_row[time_col].values[0] if not sav_row.empty else None
            save_t = hhmm_to_min(sav_t_raw)
            # 절감효과 행이 0분이면 원본-최적화 차이로 계산
            if save_t is None or save_t == 0:
                orig_t_raw = orig_row[time_col].values[0] if not orig_row.empty else None
                opt_t_raw  = opt_row[time_col].values[0]  if not opt_row.empty else None
                orig_t = hhmm_to_min(orig_t_raw)
                opt_t  = hhmm_to_min(opt_t_raw)
                if orig_t and opt_t and orig_t > opt_t:
                    save_t = orig_t - opt_t

        # ── 절감액 컬럼 (동적 컬럼명 대응) ───────────────
        day_col = find_col(summary_df, "일 총 절감액")
        mon_col = find_col(summary_df, "월 절감액")
        ann_col = find_col(summary_df, "연 절감액")

        day_save = fv(sav_row, day_col) if day_col else None
        mon_save = fv(sav_row, mon_col) if mon_col else None
        ann_save = fv(sav_row, ann_col) if ann_col else (mon_save * 12 if mon_save else None)

        # ── 카드 렌더링 ────────────────────────────────────
        cards = []

        if rate is not None:
            cards.append(("이동거리 절감률",  f"▼{rate}%",           "highlight"))
        if save_d is not None:
            cards.append(("이동거리 절감량",  f"▼{save_d:.2f}km",    "green"))
        if save_t is not None:
            h, m = divmod(int(save_t), 60)
            t_str = f"▼{h}시간 {m}분" if h > 0 else f"▼{m}분"
            cards.append(("이동시간 단축",    t_str,                  "green"))
        if day_save is not None:
            cards.append(("일 총 절감액",     fmt_money(day_save),    "green"))
        if mon_save is not None:
            cards.append((f"월 절감액({working_days}일)", fmt_money(mon_save), "green"))
        if ann_save is not None:
            cards.append(("연 절감액",        fmt_money(ann_save),    "green"))

        if not cards:
            return False

        html = '<div class="metric-grid">'
        for label, value, cls in cards[:6]:
            html += f"""
            <div class="metric-card">
                <div class="metric-label">{label}</div>
                <div class="metric-value {cls}">{value}</div>
            </div>"""
        html += '</div>'
        st.markdown(html, unsafe_allow_html=True)
        return True
    except Exception:
        return False


# =====================================================
# session_state 초기화
# =====================================================

if "ai_interpretation" not in st.session_state:
    st.session_state["ai_interpretation"] = None
if "last_result" not in st.session_state:
    st.session_state["last_result"] = None
if "last_final_params" not in st.session_state:
    st.session_state["last_final_params"] = None


# =====================================================
# 제목
# =====================================================

st.title("🚚 AI 집배순로 최적화")

# ── 네이버 API 키 사전 확인 ───────────────────────────────
_naver_id     = os.getenv("NAVER_CLIENT_ID", "").strip()
_naver_secret = os.getenv("NAVER_CLIENT_SECRET", "").strip()
if not _naver_id or not _naver_secret:
    st.error(
        "⛔ 네이버 지도 API 키가 설정되지 않았습니다.\n\n"
        "`.env` 파일에 아래 항목을 추가해주세요:\n"
        "```\nNAVER_CLIENT_ID=발급받은_ID\nNAVER_CLIENT_SECRET=발급받은_SECRET\n```"
    )
    st.stop()


# =====================================================
# 알고리즘 선택
# =====================================================

ui_algorithm = st.radio(
    "최적화 방식",
    ["corse7 - 통상순로 최적화", "corse8 - 통상코스 + 통상순로 최적화"],
    label_visibility="visible",
)

st.markdown("---")


# =====================================================
# 📂 파일 업로드  +  ⚙️ 기본값 설정 (나란히 배치)
# =====================================================

col_upload, col_params = st.columns([1, 1], gap="large")

with col_upload:
    st.markdown("##### 📂 엑셀 파일 업로드")
    uploaded_file = st.file_uploader(
        "집배 데이터 (.xlsx)",
        type=["xlsx"],
        label_visibility="collapsed",
    )
    st.caption("출발지 · 도착지 · 통상코스 · 통상순로 컬럼 포함")

    st.markdown("##### 📱 샘플 데이터 체험")

    # 툴팁용 샘플 정보 미리 파악
    try:
        _s = make_sample_excel()
        _tooltip = f"부산 동래구 {_s.n_courses}코스·{_s.n_rows}건으로 즉시 체험"
    except Exception:
        _tooltip = "샘플 데이터로 즉시 체험"

    use_sample = st.button(
        "🚀 샘플 데이터로 바로 실행",
        type="primary",
        use_container_width=True,
        help=_tooltip,
    )
    st.caption("파일 없이도 바로 체험 가능")

with col_params:
    st.markdown("##### ⚙️ 기본값 설정")
    st.caption("우리 기관 근무여건에 맞게 입력하세요")

    ui_cost_per_km = st.number_input(
        "운송비 (원 / km)",
        min_value=100,
        max_value=10000,
        value=DEFAULT_COST_PER_KM,
        step=25,
        help="km당 운송 단가",
    )
    ui_labor_cost = st.number_input(
        "인건비 (원 / 시간)",
        min_value=1000,
        max_value=100000,
        value=DEFAULT_LABOR_COST_PER_HOUR,
        step=500,
        help="시간당 인건비 단가",
    )
    ui_working_days = st.number_input(
        "근무일수 (일 / 월)",
        min_value=1,
        max_value=31,
        value=DEFAULT_WORKING_DAYS,
        step=1,
        help="월 평균 근무일수",
    )

st.markdown("---")

# ── 개인정보 안내 ──────────────────────────────────────────
st.markdown(
    """
    <div style="
        background: #f0f7ff;
        border-left: 3px solid #378ADD;
        border-radius: 0 6px 6px 0;
        padding: 8px 14px;
        font-size: 12px;
        color: #185FA5;
        margin-bottom: 12px;
    ">
    🔒 <b>개인정보 안내</b> &nbsp;
    입력된 주소 데이터는 경로 최적화 목적으로만 사용됩니다.
    도로명 주소는 공개 정보이며 개인 식별 정보(성명·연락처 등)를 포함하지 않습니다.
    주소 좌표 변환 시 네이버 지도 API를 경유하며, 그 외 외부 전송은 없습니다.
    개인 식별 정보 혼입 여부를 자동 검사하며, API 호출 내역은 로컬 로그(privacy_log.txt)에 기록됩니다.
    </div>
    """,
    unsafe_allow_html=True,
)


# =====================================================
# 실행 대상 결정
# =====================================================

target_file = None

if use_sample:
    target_file = make_sample_excel()
    st.info(f"📋 샘플 데이터 사용 중 — 부산 동래구 {target_file.n_courses}코스·{target_file.n_rows}건", icon="ℹ️")
elif uploaded_file is not None:
    target_file = uploaded_file
    st.success("✅ 업로드 완료")

    # ── ⑥ 데이터 미리보기 ─────────────────────────────────────
    try:
        uploaded_file.seek(0)
        df_preview = pd.read_excel(uploaded_file)
        uploaded_file.seek(0)   # 이후 process_excel이 다시 읽을 수 있도록 되감기

        course_cnt  = df_preview["통상코스"].nunique() if "통상코스" in df_preview.columns else "?"
        point_cnt   = len(df_preview)
        worker_list = sorted(df_preview["통상코스"].dropna().unique().tolist()) \
                      if "통상코스" in df_preview.columns else []

        st.markdown(
            f"""
            <div style="
                background: var(--color-background-secondary);
                border-radius: 8px;
                padding: 10px 16px;
                font-size: 13px;
                color: var(--color-text-secondary);
                margin: 6px 0 10px;
                display: flex;
                gap: 20px;
                flex-wrap: wrap;
            ">
            <span>🗂 <b style="color:var(--color-text-primary)">코스 수</b> {course_cnt}개</span>
            <span>📦 <b style="color:var(--color-text-primary)">총 배송지</b> {point_cnt:,}건</span>
            <span>🔢 <b style="color:var(--color-text-primary)">코스 목록</b> {worker_list}</span>
            </div>
            """,
            unsafe_allow_html=True,
        )
    except Exception:
        pass   # 미리보기 실패 시 조용히 스킵


# =====================================================
# 실행
# =====================================================

if target_file is not None:

    run_label = "샘플 데이터 최적화 실행" if use_sample else "⚡ 최적화 실행"

    if use_sample or st.button(run_label, type="primary", width='stretch'):

        algo_key = "corse7" if ui_algorithm.startswith("corse7") else "corse8"

        # optimizer 파라미터 주입
        import corse7_optimizer as c7
        import corse8_optimizer as c8
        for mod in (c7, c8):
            mod.COST_PER_KM         = ui_cost_per_km
            mod.LABOR_COST_PER_HOUR = ui_labor_cost

        moto_bar = st.empty()

        def update_progress(pct, msg):
            pct    = max(0, min(100, pct))
            total  = 20
            filled = int(pct / 100 * total)
            empty  = total - filled - 1
            bar    = "━" * filled + "🏍️" + "─" * empty
            moto_bar.markdown(f"**{bar}** `{pct}%`", unsafe_allow_html=True)

        update_progress(0, "")

        # ── ④ 예외처리 ────────────────────────────────────────
        try:
            with st.spinner("AI 최적화 진행중..."):
                process_fn = process7 if algo_key == "corse7" else process8
                result = process_fn(
                    target_file,
                    progress_callback=update_progress,
                    working_days=int(ui_working_days),
                )

            moto_bar.markdown("**━━━━━━━━━━━━━━━━━━━🏁** `완료!`")

            # 결과 session_state 저장
            st.session_state["last_result"] = result
            st.session_state["last_final_params"] = {
                "algorithm":           algo_key,
                "cost_per_km":         ui_cost_per_km,
                "labor_cost_per_hour": ui_labor_cost,
                "working_days":        ui_working_days,
            }
            st.session_state["ai_interpretation"] = None

        except ValueError as e:
            moto_bar.empty()
            st.error(f"⚠️ 데이터 오류: {e}\n\n엑셀 파일의 컬럼명(출발지·도착지·통상코스·통상순로)을 확인해주세요.")
        except Exception as e:
            moto_bar.empty()
            st.error(f"⚠️ 최적화 중 오류가 발생했습니다: {e}")


# =====================================================
# 결과 표시 — session_state 기반으로 항상 유지
# =====================================================

if st.session_state.get("last_result"):

    result       = st.session_state["last_result"]
    final_params = st.session_state["last_final_params"]

    algo_label = (
        "CORSE7 (통상순로 최적화)"
        if final_params["algorithm"] == "corse7"
        else "CORSE8 (통상코스 + 통상순로 최적화)"
    )

    st.success("🎉 최적화 완료!")

    # 실패 주소 — session_state 기반으로 항상 표시
    failed_addrs = result.get("failed_addresses", [])
    if failed_addrs:
        st.warning(
            f"⚠️ 좌표 변환 실패 주소 {len(failed_addrs)}건 — 최적화에서 제외됩니다.\n\n"
            + "\n".join(f"  • {a}" for a in failed_addrs)
        )

    # 실행 파라미터 배너
    st.markdown(
        f'<div class="param-box">'
        f'<b>실행 파라미터</b> &nbsp;'
        f'<span class="param-tag">{final_params["algorithm"].upper()}</span>'
        f'<span class="param-tag">운송비 {final_params["cost_per_km"]:,}원/km</span>'
        f'<span class="param-tag">인건비 {final_params["labor_cost_per_hour"]:,}원/h</span>'
        f'<span class="param-tag">월 {final_params["working_days"]}일</span>'
        f'</div>',
        unsafe_allow_html=True,
    )

    # ── 요약 결과 ─────────────────────────────────────────────
    st.markdown("#### 📊 요약 결과")
    summary = result["summary"]
    card_ok = render_summary_cards(summary, working_days=final_params.get("working_days", 21))

    with st.expander("📋 상세 요약 테이블 보기", expanded=not card_ok):
        st.dataframe(summary, width='stretch')

    # ── 엑셀 다운로드 ─────────────────────────────────────────
    excel_buffer = io.BytesIO()
    with pd.ExcelWriter(excel_buffer, engine="openpyxl") as writer:
        result["df_original"].to_excel(writer, sheet_name="원본", index=False)
        result["df_optimized"].to_excel(writer, sheet_name="최적화", index=False)
        result["summary"].to_excel(writer, sheet_name="요약", index=False)
        # 실패 주소 시트 — 변환 실패 건이 있을 때만 추가
        failed = result.get("failed_addresses", [])
        if failed:
            pd.DataFrame({"좌표변환 실패 주소": failed}).to_excel(
                writer, sheet_name="변환실패", index=False
            )

    st.download_button(
        label="📥 결과 엑셀 다운로드",
        data=excel_buffer.getvalue(),
        file_name="최적화_결과.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    # ── AI 결과 해석 (Claude API) ──────────────────────────────
    st.markdown("#### 🤖 AI 결과 해석")

    if not st.session_state.get("ai_interpretation"):
        col_btn, _ = st.columns([1, 2])
        with col_btn:
            gen_btn = st.button("✨ AI 해석 생성", type="primary", use_container_width=True)

        if gen_btn:
            if not os.getenv("ANTHROPIC_API_KEY", "").strip():
                st.warning("⚠️ ANTHROPIC_API_KEY가 .env에 설정되지 않았습니다.")
            else:
                with st.spinner("Claude AI가 결과를 분석 중입니다..."):
                    summary_text   = build_summary_text(
                        summary, algo_label,
                        final_params["cost_per_km"],
                        final_params["labor_cost_per_hour"],
                        final_params["working_days"],
                    )
                    interpretation = call_claude_interpretation(summary_text, algo_label)
                    st.session_state["ai_interpretation"] = interpretation
                    st.rerun()
    else:
        st.markdown(
            f'<div class="ai-box">'
            f'{st.session_state["ai_interpretation"].replace(chr(10), "<br>")}'
            f'</div>',
            unsafe_allow_html=True,
        )
        col_dl, col_re = st.columns([1, 1])
        with col_dl:
            st.download_button(
                label="📄 AI 해석 텍스트 저장",
                data=st.session_state["ai_interpretation"],
                file_name="AI_결과해석.txt",
                mime="text/plain",
                key="ai_text_dl",
            )
        with col_re:
            if st.button("↺ 다시 생성", use_container_width=True):
                st.session_state["ai_interpretation"] = None
                st.rerun()

    # ── 전국 절감 추정 계산기 ─────────────────────────────────
    st.markdown("#### 🇰🇷 전국 절감 추정 계산기")
    st.caption("집배원 수를 조정하면 전국 절감 추정액이 실시간으로 계산됩니다")

    col_slider, col_result = st.columns([2, 1])
    with col_slider:
        national_workers = st.slider(
            "집배원 수 (명)",
            min_value=1,
            max_value=20000,
            value=17000,
            step=1,
            help="우체국 단위(10명 내외) ~ 전국(약 17,000명) 규모로 조정 가능",
        )
    with col_result:
        try:
            saved_row = summary[summary["구분"] == "절감효과"]
            if not saved_row.empty:
                yearly_cols = [c for c in summary.columns if "연" in c and "절감" in c]
                if yearly_cols:
                    # 파일 1개 = 집배원 1명 기준 → 그대로 1인 절감액
                    yearly_per      = float(str(saved_row[yearly_cols[0]].values[0]).replace(",", ""))
                    national_saving = yearly_per * national_workers
                    st.metric(
                        label=f"집배원 {national_workers:,}명 기준",
                        value=f"{national_saving/100000000:,.1f}억원/년",
                        delta=f"1인 평균 {yearly_per/1000:,.0f}천원/년",
                    )
        except Exception:
            st.caption("계산 결과를 표시할 수 없습니다.")

    # ── 캐시 통계 ─────────────────────────────────────────────
    try:
        import cache_manager as cm
        stats = cm.get_stats()
        if stats["api_calls_made"] + stats["api_calls_saved"] > 0:
            with st.expander("💾 API 캐시 통계", expanded=False):
                c1, c2, c3, c4 = st.columns(4)
                c1.metric("Geocode 캐시", f"{stats['geocode_cache_size']:,}건")
                c2.metric("도로거리 캐시", f"{stats['road_cache_size']:,}건")
                c3.metric("API 절약", f"{stats['api_calls_saved']:,}회")
                c4.metric("캐시 히트율",
                          f"{max(stats['geocode_hit_rate'], stats['road_hit_rate']):.1f}%")
    except Exception:
        pass

    # ── 지도 비교 ─────────────────────────────────────────────
    st.markdown("#### 🗺️ 지도 비교")

    course_count = len(result.get("compare_maps", []))

    if course_count == 0:
        st.warning("생성된 지도가 없습니다.")
    else:
        MAP_HEIGHT = 450
        # 실제 코스 번호 추출
        try:
            course_nums = sorted(result["df_optimized"]["통상코스"].dropna().unique().astype(int).tolist())
        except Exception:
            course_nums = list(range(1, course_count + 1))
        course_tabs = st.tabs([f"코스 {c}" for c in course_nums])

        for i, tab in enumerate(course_tabs):
            with tab:
                map_tabs = st.tabs(["원본 🔴", "최적화 🔵", "비교"])

                with map_tabs[0]:
                    try:
                        with open(result["original_maps"][i], "r", encoding="utf-8") as f:
                            original_html = f.read()
                        st.iframe(original_html, height=MAP_HEIGHT)
                        st.download_button(
                            label="📥 원본 지도 다운로드",
                            data=original_html.encode("utf-8"),
                            file_name=os.path.basename(result["original_maps"][i]),
                            mime="text/html",
                            key=f"orig_{i}",
                        )
                    except Exception:
                        st.warning("원본 지도 파일을 불러올 수 없습니다.")

                with map_tabs[1]:
                    try:
                        with open(result["optimized_maps"][i], "r", encoding="utf-8") as f:
                            optimized_html = f.read()
                        st.iframe(optimized_html, height=MAP_HEIGHT)
                        st.download_button(
                            label="📥 최적화 지도 다운로드",
                            data=optimized_html.encode("utf-8"),
                            file_name=os.path.basename(result["optimized_maps"][i]),
                            mime="text/html",
                            key=f"opt_{i}",
                        )
                    except Exception:
                        st.warning("최적화 지도 파일을 불러올 수 없습니다.")

                with map_tabs[2]:
                    try:
                        with open(result["compare_maps"][i], "r", encoding="utf-8") as f:
                            compare_html = f.read()
                        st.iframe(compare_html, height=MAP_HEIGHT)
                        st.download_button(
                            label="📥 비교 지도 다운로드",
                            data=compare_html.encode("utf-8"),
                            file_name=os.path.basename(result["compare_maps"][i]),
                            mime="text/html",
                            key=f"cmp_{i}",
                        )
                    except Exception:
                        st.warning("비교 지도 파일을 불러올 수 없습니다.")

        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
            for path in result["original_maps"] + result["optimized_maps"] + result["compare_maps"]:
                try:
                    if os.path.exists(path):
                        zf.write(path, os.path.basename(path))
                except Exception:
                    pass

        st.download_button(
            label="🗺️ 지도 전체 ZIP 다운로드",
            data=zip_buffer.getvalue(),
            file_name="최적화_지도.zip",
            mime="application/zip",
        )

    # ── 결과 해석 한계 고지 ───────────────────────────────────
    st.markdown("---")
    st.markdown(
        """
        <div style="
            background: var(--color-background-secondary, #f8f9fa);
            border-left: 3px solid #888;
            border-radius: 0 6px 6px 0;
            padding: 8px 14px;
            font-size: 11px;
            color: #666;
            margin-top: 8px;
        ">
        ⚠️ <b>결과 해석 안내</b> &nbsp;
        본 최적화 결과는 입력 데이터·주소 정확도·외부 API 응답에 따라 달라질 수 있으며,
        현장 적용을 위한 <b>의사결정 참고자료</b>로 활용해야 합니다.
        실제 집배구역 특성·도로 상황·안전 문제·담당자 숙련도 등 현장 여건을 함께 고려하시기 바랍니다.
        </div>
        """,
        unsafe_allow_html=True,
    )
