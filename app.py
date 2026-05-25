import streamlit as st
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
    """summary DataFrame의 숫자 컬럼을 카드 그리드로 표시"""
    try:
        # 합계 행 찾기 (마지막 행 또는 '합계' 포함 행)
        mask = summary_df.astype(str).apply(
            lambda col: col.str.contains("합계|전체|합 계|total", case=False, na=False)
        ).any(axis=1)
        if mask.any():
            row = summary_df[mask].iloc[0]
        else:
            row = summary_df.iloc[-1]  # 마지막 행

        # 숫자형 컬럼만 추출
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

        # 컬럼명 기반으로 색상/강조 결정
        def card_class(col_name):
            c = col_name.lower()
            if any(k in c for k in ["절감률", "절감율", "단축률", "%", "율"]):
                return "highlight"
            if any(k in c for k in ["절감", "단축", "최적"]):
                return "green"
            return ""

        # 단위 포맷팅
        def fmt_val(col_name, raw_str, fval):
            c = col_name.lower()
            if any(k in c for k in ["절감액", "비용", "원"]):
                if fval >= 10_000_000:
                    return f"{fval/10_000_000:.1f}천만원"
                elif fval >= 10_000:
                    return f"{fval/10_000:.0f}만원"
                else:
                    return f"{fval:,.0f}원"
            return raw_str  # 원본 문자열 그대로 (이미 단위 포함된 경우)

        html = '<div class="metric-grid">'
        for col, fval, raw in num_cols[:6]:  # 최대 6개
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

    except Exception as e:
        return False


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
    width='stretch',
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

    if use_sample or st.button(run_label, type="primary", width='stretch'):

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

        # 숫자 컬럼만 골라서 모바일 친화적 소형 테이블로 표시
        summary = result["summary"]

        # 핵심 수치 카드 (파싱 성공 시만)
        card_ok = render_summary_cards(summary)

        # 항상 전체 테이블 표시 (가로 스크롤 가능)
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
            # 모바일 지도 높이: 화면 높이의 약 60% (450px)
            MAP_HEIGHT = 450

            course_tabs = st.tabs([f"코스 {i+1}" for i in range(course_count)])

            for i, tab in enumerate(course_tabs):
                with tab:
                    map_tabs = st.tabs(["원본 🔴", "최적화 🔵", "비교"])

                    with map_tabs[0]:
                        with open(result["original_maps"][i], "r", encoding="utf-8") as f:
                            original_html = f.read()
                        st.iframe(original_html, height=MAP_HEIGHT)
                        st.download_button(
                            label=f"📥 원본 지도 다운로드",
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
                            label=f"📥 최적화 지도 다운로드",
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
                            label=f"📥 비교 지도 다운로드",
                            data=compare_html.encode("utf-8"),
                            file_name=os.path.basename(result["compare_maps"][i]),
                            mime="text/html",
                            key=f"cmp_{i}",
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
            )
