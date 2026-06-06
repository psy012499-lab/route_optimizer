import streamlit as st
import io
import os
import zipfile
import pandas as pd
from corse7_optimizer import process_excel as process7
from corse8_optimizer import process_excel as process8


# =====================================================
# 설정 상수
# =====================================================

DEFAULT_COST_PER_KM         = 925
DEFAULT_LABOR_COST_PER_HOUR = 17000
DEFAULT_WORKING_DAYS        = 21


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
</style>
""", unsafe_allow_html=True)


# =====================================================
# 샘플 데이터 생성
# =====================================================

def make_sample_excel() -> io.BytesIO:
    sample_rows = [
        ("부산광역시 동구 중앙대로 270", "부산광역시 동구 초량중로 43-1",  1, 1),
        ("부산광역시 동구 중앙대로 270", "부산광역시 동구 범일로 100",     1, 2),
        ("부산광역시 동구 중앙대로 270", "부산광역시 동구 수정동 50-1",    1, 3),
        ("부산광역시 동구 중앙대로 270", "부산광역시 동구 좌천동 1가 10",  1, 4),
        ("부산광역시 동구 중앙대로 270", "부산광역시 동구 초량동 1183",    1, 5),
        ("부산광역시 동구 중앙대로 270", "부산광역시 동구 범일동 830-55",  1, 6),
        ("부산광역시 동구 중앙대로 270", "부산광역시 동구 수정4동 108-2",  1, 7),
        ("부산광역시 동구 중앙대로 270", "부산광역시 동구 좌천2동 22-5",   1, 8),
        ("부산광역시 동구 중앙대로 270", "부산광역시 동구 초량3동 55-1",   1, 9),
        ("부산광역시 동구 중앙대로 270", "부산광역시 동구 범일2동 44-3",   1, 10),
        ("부산광역시 동구 중앙대로 270", "부산광역시 동구 자성로 67",      2, 1),
        ("부산광역시 동구 중앙대로 270", "부산광역시 동구 중앙대로 999",   2, 2),
        ("부산광역시 동구 중앙대로 270", "부산광역시 동구 망양로 450",     2, 3),
        ("부산광역시 동구 중앙대로 270", "부산광역시 동구 초량동 792-4",   2, 4),
        ("부산광역시 동구 중앙대로 270", "부산광역시 동구 수정2동 65",     2, 5),
        ("부산광역시 동구 중앙대로 270", "부산광역시 동구 범일동 12-7",    2, 6),
        ("부산광역시 동구 중앙대로 270", "부산광역시 동구 좌천동 100",     2, 7),
        ("부산광역시 동구 중앙대로 270", "부산광역시 동구 초량2동 210",    2, 8),
        ("부산광역시 동구 중앙대로 270", "부산광역시 동구 수정5동 33-9",   2, 9),
        ("부산광역시 동구 중앙대로 270", "부산광역시 동구 범일3동 77",     2, 10),
    ]
    df = pd.DataFrame(sample_rows, columns=["출발지", "도착지", "통상코스", "통상순로"])
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Sheet1")
    buf.seek(0)
    buf.name = "sample_input.xlsx"
    return buf


# =====================================================
# 요약 카드 렌더링
# =====================================================

def render_summary_cards(summary_df: pd.DataFrame):
    try:
        mask = summary_df.astype(str).apply(
            lambda col: col.str.contains("합계|전체|합 계|total", case=False, na=False)
        ).any(axis=1)
        row = summary_df[mask].iloc[0] if mask.any() else summary_df.iloc[-1]

        num_cols = []
        for col in summary_df.columns:
            val = row[col]
            try:
                fval = float(str(val).replace(",", "").replace("%", ""))
                num_cols.append((str(col), fval, str(val)))
            except (ValueError, TypeError):
                pass

        if not num_cols:
            return False

        def card_class(col_name):
            c = col_name.lower()
            if any(k in c for k in ["절감률", "절감율", "단축률", "%", "율"]):
                return "highlight"
            if any(k in c for k in ["절감", "단축", "최적"]):
                return "green"
            return ""

        def fmt_val(col_name, raw_str, fval):
            c = col_name.lower()
            if any(k in c for k in ["절감액", "비용", "원"]):
                if fval >= 1_000_000:
                    return f"{fval/1000:,.0f}천원"
                elif fval >= 1_000:
                    return f"{fval/1000:.1f}천원"
                else:
                    return f"{fval:,.0f}원"
            return raw_str

        html = '<div class="metric-grid">'
        for col, fval, raw in num_cols[:6]:
            cls   = card_class(col)
            value = fmt_val(col, raw, fval)
            html += f"""
            <div class="metric-card">
                <div class="metric-label">{col}</div>
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

if "last_result" not in st.session_state:
    st.session_state["last_result"] = None
if "last_final_params" not in st.session_state:
    st.session_state["last_final_params"] = None


# =====================================================
# 제목
# =====================================================

st.title("🚚 AI 집배순로 최적화")


# =====================================================
# 알고리즘 선택
# =====================================================

ui_algorithm = st.radio(
    "최적화 방식",
    ["corse7 - 통상순로 최적화", "corse8 - 통상순로 + 통상코스 최적화"],
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
    use_sample = st.button(
        "🚀 샘플 데이터로 바로 실행",
        type="primary",
        use_container_width=True,
        help="부산 동구 2코스·20건으로 즉시 체험",
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


# =====================================================
# 실행 대상 결정
# =====================================================

target_file = None

if use_sample:
    target_file = make_sample_excel()
    st.info("📋 샘플 데이터 사용 중 — 부산 동구 2코스·20건", icon="ℹ️")
elif uploaded_file is not None:
    target_file = uploaded_file
    st.success("✅ 업로드 완료")


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


# =====================================================
# 결과 표시 — session_state 기반으로 항상 유지
# =====================================================

if st.session_state.get("last_result"):

    result       = st.session_state["last_result"]
    final_params = st.session_state["last_final_params"]

    algo_label = (
        "CORSE7 (통상순로 최적화)"
        if final_params["algorithm"] == "corse7"
        else "CORSE8 (통상순로+코스 최적화)"
    )

    st.success("🎉 최적화 완료!")

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
    card_ok = render_summary_cards(summary)

    with st.expander("📋 상세 요약 테이블 보기", expanded=not card_ok):
        st.dataframe(summary, width='stretch')

    # ── 엑셀 다운로드 ─────────────────────────────────────────
    excel_buffer = io.BytesIO()
    with pd.ExcelWriter(excel_buffer, engine="openpyxl") as writer:
        result["df_original"].to_excel(writer, sheet_name="원본", index=False)
        result["df_optimized"].to_excel(writer, sheet_name="최적화", index=False)
        result["summary"].to_excel(writer, sheet_name="요약", index=False)

    st.download_button(
        label="📥 결과 엑셀 다운로드",
        data=excel_buffer.getvalue(),
        file_name="최적화_결과.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    # ── 지도 비교 ─────────────────────────────────────────────
    st.markdown("#### 🗺️ 지도 비교")

    course_count = len(result.get("compare_maps", []))

    if course_count == 0:
        st.warning("생성된 지도가 없습니다.")
    else:
        MAP_HEIGHT  = 450
        course_tabs = st.tabs([f"코스 {i+1}" for i in range(course_count)])

        for i, tab in enumerate(course_tabs):
            with tab:
                map_tabs = st.tabs(["원본 🔴", "최적화 🔵", "비교"])

                with map_tabs[0]:
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

                with map_tabs[1]:
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

                with map_tabs[2]:
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

        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
            for path in result["original_maps"]:
                zf.write(path, os.path.basename(path))
            for path in result["optimized_maps"]:
                zf.write(path, os.path.basename(path))
            for path in result["compare_maps"]:
                zf.write(path, os.path.basename(path))

        st.download_button(
            label="🗺️ 지도 전체 ZIP 다운로드",
            data=zip_buffer.getvalue(),
            file_name="최적화_지도.zip",
            mime="application/zip",
        )
