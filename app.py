import streamlit as st
import streamlit.components.v1 as components
import io
import os
import zipfile
import pandas as pd

from corse7_optimizer import process_excel as process7
from corse8_optimizer import process_excel as process8


# =====================================================
# 페이지 설정 — 모바일 친화적으로 "centered" 사용
# =====================================================

st.set_page_config(
    page_title="AI 집배순로 최적화",
    page_icon="🚚",
    layout="centered",          # wide → centered: 모바일 여백 제거
    initial_sidebar_state="collapsed",
)

# ── 모바일 최적화 CSS ──────────────────────────────────────────────────
st.markdown("""
<style>
/* 전체 폰트 크기 조정 */
html, body, [class*="css"] { font-size: 15px; }

/* 제목 크기 축소 */
h1 { font-size: 1.5rem !important; }
h2 { font-size: 1.2rem !important; }
h3 { font-size: 1.05rem !important; }

/* 요약 카드 그리드 */
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
.metric-label {
    font-size: 11px;
    color: #666;
    margin-bottom: 4px;
}
.metric-value {
    font-size: 1.35rem;
    font-weight: 700;
    color: #1C2B4A;
}
.metric-sub {
    font-size: 11px;
    color: #888;
    margin-top: 2px;
}
.metric-value.highlight { color: #E67E22; }
.metric-value.green     { color: #2e7d32; }

/* 테이블 가로 스크롤 */
.stDataFrame { overflow-x: auto !important; }
.stDataFrame table { font-size: 12px !important; }
.stDataFrame th, .stDataFrame td { padding: 4px 8px !important; white-space: nowrap; }

/* 버튼 터치 영역 */
.stButton > button { min-height: 48px; font-size: 15px; }
.stDownloadButton > button { min-height: 44px; font-size: 14px; }

/* 탭 폰트 */
.stTabs [data-baseweb="tab"] { font-size: 13px; padding: 6px 12px; }

/* 진행바 폰트 축소 */
.progress-text { font-size: 13px; }
</style>
""", unsafe_allow_html=True)


# =====================================================
# 샘플 데이터 생성 함수
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


# ── 결과 요약을 모바일 친화적 카드로 렌더링 ───────────────────────────
def render_summary_cards(summary_df: pd.DataFrame):
    """summary DataFrame에서 핵심 수치를 카드 그리드로 표시"""

    # summary DataFrame 컬럼 구조에 맞게 값 추출
    # 예상 컬럼: 코스, 원본거리(km), 최적화거리(km), 절감거리(km), 절감률(%), 연간절감액 등
    try:
        total_row = summary_df[summary_df.iloc[:, 0].astype(str).str.contains("합계|전체|total", case=False, na=False)]
        if total_row.empty:
            total_row = summary_df.iloc[[-1]]  # 마지막 행을 합계로 간주

        row = total_row.iloc[0]
        cols = summary_df.columns.tolist()

        def get_col(keywords):
            for kw in keywords:
                for c in cols:
                    if kw in str(c):
                        return row[c]
            return None

        orig_km   = get_col(["원본거리", "원본_거리", "원본 거리", "기존거리"])
        opt_km    = get_col(["최적거리", "최적화거리", "최적화_거리", "최적 거리"])
        save_km   = get_col(["절감거리", "절감_거리", "단축"])
        save_pct  = get_col(["절감률", "절감율", "단축률", "%"])
        annual    = get_col(["연간절감", "연간_절감", "연간 절감", "절감액"])

        cards = []
        if save_pct  is not None: cards.append(("이동거리 단축", f"▼{float(save_pct):.1f}%",  "",          "highlight"))
        if save_km   is not None: cards.append(("절감 거리",     f"{float(save_km):.1f}km",    "하루 기준", "green"))
        if orig_km   is not None: cards.append(("원본 거리",     f"{float(orig_km):.1f}km",    "",          ""))
        if opt_km    is not None: cards.append(("최적화 거리",   f"{float(opt_km):.1f}km",     "",          ""))
        if annual    is not None:
            val = float(annual)
            label = f"{val/10000:.0f}만원" if val >= 10000 else f"{val:.0f}원"
            cards.append(("연간 절감액", label, "4명 기준", "green"))

        if not cards:
            raise ValueError("카드 데이터 없음")

        # 2열 그리드로 렌더링
        html = '<div class="metric-grid">'
        for label, value, sub, cls in cards:
            html += f"""
            <div class="metric-card">
                <div class="metric-label">{label}</div>
                <div class="metric-value {cls}">{value}</div>
                {'<div class="metric-sub">' + sub + '</div>' if sub else ''}
            </div>"""
        html += '</div>'
        st.markdown(html, unsafe_allow_html=True)
        return True

    except Exception:
        return False  # 카드 렌더링 실패 시 원본 dataframe으로 폴백


# =====================================================
# 제목
# =====================================================

st.title("🚚 AI 집배순로 최적화")


# =====================================================
# 알고리즘 선택
# =====================================================

algorithm = st.radio(
    "최적화 방식",
    ["corse7 - 통상순로 최적화", "corse8 - 통상순로 + 통상코스 최적화"],
    label_visibility="visible",
)


# =====================================================
# 샘플 데이터 체험
# =====================================================

st.markdown("---")
st.markdown("##### 📱 파일 없이 바로 체험")

use_sample = st.button(
    "🚀 샘플 데이터로 바로 실행",
    type="primary",
    use_container_width=True,
    help="부산 동구 주소 기반 샘플 2코스·20건으로 즉시 체험",
)
st.caption("엑셀 없이 체험 가능 · 부산 동구 샘플 2코스·20건")

st.markdown("---")


# =====================================================
# 파일 업로드
# =====================================================

st.markdown("##### 📂 직접 파일 업로드")

uploaded_file = st.file_uploader("엑셀 파일 (.xlsx)", type=["xlsx"])


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

    if use_sample or st.button(run_label, type="primary", use_container_width=True):

        moto_bar = st.empty()

        def update_progress(pct, msg):
            pct = max(0, min(100, pct))
            total = 20   # 모바일에서 바 길이 줄임
            filled = int(pct / 100 * total)
            empty = total - filled - 1
            bar = "━" * filled + "🏍️" + "─" * empty
            moto_bar.markdown(f"**{bar}** `{pct}%`", unsafe_allow_html=True)

        update_progress(0, "")

        with st.spinner("AI 최적화 진행중..."):
            if algorithm.startswith("corse7"):
                result = process7(target_file, progress_callback=update_progress)
            else:
                result = process8(target_file, progress_callback=update_progress)

        moto_bar.markdown("**━━━━━━━━━━━━━━━━━━━🏁** `완료!`")
        st.success("🎉 최적화 완료!")

        # ── 요약 결과 ─────────────────────────────────────────────
        st.markdown("#### 📊 요약 결과")

        # 카드 렌더링 시도 → 실패 시 compact dataframe
        card_ok = render_summary_cards(result["summary"])
        if not card_ok:
            # 컬럼 수가 많을 때 가로 스크롤 가능한 소형 테이블
            st.dataframe(
                result["summary"],
                use_container_width=True,
                height=min(200, 40 + 35 * len(result["summary"])),
            )

        # 전체 요약 테이블 (접을 수 있게)
        with st.expander("📋 상세 요약 테이블 보기"):
            st.dataframe(result["summary"], use_container_width=True)

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
            use_container_width=True,
        )

        # ── 지도 비교 ─────────────────────────────────────────────
        st.markdown("#### 🗺️ 지도 비교")

        course_count = len(result.get("compare_maps", []))

        if course_count == 0:
            st.warning("생성된 지도가 없습니다.")
        else:
            # 모바일 지도 높이: 화면 높이의 약 60% (450px)
            MAP_HEIGHT = 450

            course_tabs = st.tabs([f"코스 {i+1}" for i in range(course_count)])

            for i, tab in enumerate(course_tabs):
                with tab:
                    map_tabs = st.tabs(["원본 🔴", "최적화 🔵", "비교"])

                    with map_tabs[0]:
                        with open(result["original_maps"][i], "r", encoding="utf-8") as f:
                            original_html = f.read()
                        components.html(original_html, height=MAP_HEIGHT)
                        st.download_button(
                            label=f"📥 원본 지도 다운로드",
                            data=original_html.encode("utf-8"),
                            file_name=os.path.basename(result["original_maps"][i]),
                            mime="text/html",
                            key=f"orig_{i}",
                            use_container_width=True,
                        )

                    with map_tabs[1]:
                        with open(result["optimized_maps"][i], "r", encoding="utf-8") as f:
                            optimized_html = f.read()
                        components.html(optimized_html, height=MAP_HEIGHT)
                        st.download_button(
                            label=f"📥 최적화 지도 다운로드",
                            data=optimized_html.encode("utf-8"),
                            file_name=os.path.basename(result["optimized_maps"][i]),
                            mime="text/html",
                            key=f"opt_{i}",
                            use_container_width=True,
                        )

                    with map_tabs[2]:
                        with open(result["compare_maps"][i], "r", encoding="utf-8") as f:
                            compare_html = f.read()
                        components.html(compare_html, height=MAP_HEIGHT)
                        st.download_button(
                            label=f"📥 비교 지도 다운로드",
                            data=compare_html.encode("utf-8"),
                            file_name=os.path.basename(result["compare_maps"][i]),
                            mime="text/html",
                            key=f"cmp_{i}",
                            use_container_width=True,
                        )

            # ── ZIP 다운로드 ───────────────────────────────────────
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
                use_container_width=True,
            )
