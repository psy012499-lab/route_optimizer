import streamlit as st
import streamlit.components.v1 as components
import io
import os
import zipfile
import pandas as pd

from corse7_optimizer import process_excel as process7
from corse8_optimizer import process_excel as process8


# =====================================================
# 모바일 친화적 페이지 설정
# =====================================================

st.set_page_config(
    page_title="AI 집배순로 최적화",
    page_icon="🚚",
    layout="wide",
    initial_sidebar_state="collapsed",
)


# =====================================================
# 샘플 데이터 생성 함수
# =====================================================

def make_sample_excel() -> io.BytesIO:
    """
    실제 집배 데이터 구조와 동일한 샘플 엑셀을 메모리에 생성합니다.
    출발지(우체국), 도착지(배달지), 통상코스, 통상순로 4개 컬럼만 사용합니다.
    """
    sample_rows = [
        # 코스 1 — 순서가 비효율적으로 섞여 있어 최적화 효과가 잘 보임
        ("부산광역시 동구 중앙대로 270",  "부산광역시 동구 초량중로 43-1",   1, 1),
        ("부산광역시 동구 중앙대로 270",  "부산광역시 동구 범일로 100",      1, 2),
        ("부산광역시 동구 중앙대로 270",  "부산광역시 동구 수정동 50-1",     1, 3),
        ("부산광역시 동구 중앙대로 270",  "부산광역시 동구 좌천동 1가 10",   1, 4),
        ("부산광역시 동구 중앙대로 270",  "부산광역시 동구 초량동 1183",      1, 5),
        ("부산광역시 동구 중앙대로 270",  "부산광역시 동구 범일동 830-55",    1, 6),
        ("부산광역시 동구 중앙대로 270",  "부산광역시 동구 수정4동 108-2",   1, 7),
        ("부산광역시 동구 중앙대로 270",  "부산광역시 동구 좌천2동 22-5",    1, 8),
        ("부산광역시 동구 중앙대로 270",  "부산광역시 동구 초량3동 55-1",    1, 9),
        ("부산광역시 동구 중앙대로 270",  "부산광역시 동구 범일2동 44-3",    1, 10),
        # 코스 2
        ("부산광역시 동구 중앙대로 270",  "부산광역시 동구 자성로 67",        2, 1),
        ("부산광역시 동구 중앙대로 270",  "부산광역시 동구 중앙대로 999",     2, 2),
        ("부산광역시 동구 중앙대로 270",  "부산광역시 동구 망양로 450",       2, 3),
        ("부산광역시 동구 중앙대로 270",  "부산광역시 동구 초량동 792-4",     2, 4),
        ("부산광역시 동구 중앙대로 270",  "부산광역시 동구 수정2동 65",       2, 5),
        ("부산광역시 동구 중앙대로 270",  "부산광역시 동구 범일동 12-7",      2, 6),
        ("부산광역시 동구 중앙대로 270",  "부산광역시 동구 좌천동 100",       2, 7),
        ("부산광역시 동구 중앙대로 270",  "부산광역시 동구 초량2동 210",      2, 8),
        ("부산광역시 동구 중앙대로 270",  "부산광역시 동구 수정5동 33-9",    2, 9),
        ("부산광역시 동구 중앙대로 270",  "부산광역시 동구 범일3동 77",       2, 10),
    ]

    df = pd.DataFrame(sample_rows, columns=["출발지", "도착지", "통상코스", "통상순로"])

    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Sheet1")
    buf.seek(0)
    buf.name = "sample_input.xlsx"   # file_uploader 반환 객체처럼 .name 속성 부여
    return buf


# =====================================================
# 제목
# =====================================================

st.title("AI 집배순로 최적화 시스템")


# =====================================================
# 알고리즘 선택
# =====================================================

algorithm = st.radio(
    "최적화 방식 선택",
    [
        "corse7 - 통상순로 최적화",
        "corse8 - 통상순로 + 통상코스 최적화",
    ],
)


# =====================================================
# 샘플 데이터 체험 버튼
# =====================================================

st.markdown("---")
st.markdown("#### 📱 파일 없이 바로 체험")

col_btn, col_desc = st.columns([1, 2])
with col_btn:
    use_sample = st.button(
        "🚀 샘플 데이터로 바로 실행",
        type="primary",
        use_container_width=True,
        help="부산 동구 실제 주소 기반 샘플 20건으로 즉시 최적화를 체험합니다.",
    )
with col_desc:
    st.caption("엑셀 파일 없이 바로 체험할 수 있습니다.\n부산 동구 주소 기반 샘플 데이터 2코스·20건이 자동으로 입력됩니다.")

st.markdown("---")


# =====================================================
# 파일 업로드
# =====================================================

st.markdown("#### 📂 직접 파일 업로드")

uploaded_file = st.file_uploader(
    "엑셀 파일 업로드",
    type=["xlsx"],
)


# =====================================================
# 실행 대상 결정
# 우선순위: 샘플 버튼 > 업로드 파일
# =====================================================

target_file = None

if use_sample:
    target_file = make_sample_excel()
    st.info("📋 샘플 데이터를 사용합니다 — 부산 동구 2코스·20건", icon="ℹ️")
elif uploaded_file is not None:
    target_file = uploaded_file
    st.success("엑셀 업로드 완료")


# =====================================================
# 실행
# =====================================================

if target_file is not None:

    run_label = "최적화 실행" if not use_sample else "샘플 데이터 최적화 실행"

    if use_sample or st.button(run_label, type="primary", use_container_width=True):

        moto_bar = st.empty()

        def update_progress(pct, msg):
            pct = max(0, min(100, pct))
            total = 30
            filled = int(pct / 100 * total)
            empty = total - filled - 1
            bar = "━" * filled + "🏍️" + "─" * empty
            moto_bar.markdown(
                f"**{bar}** `{pct}%`",
                unsafe_allow_html=True,
            )

        update_progress(0, "")

        with st.spinner("AI 최적화 진행중..."):
            if algorithm.startswith("corse7"):
                result = process7(target_file, progress_callback=update_progress)
            else:
                result = process8(target_file, progress_callback=update_progress)

        moto_bar.markdown("**━━━━━━━━━━━━━━━━━━━━━━━━━━━━━🏁** `완료!`")

        st.success("최적화 완료 🎉")

        # =============================================
        # 결과 요약
        # =============================================

        st.subheader("요약 결과")
        st.dataframe(result["summary"], use_container_width=True)

        # =============================================
        # 엑셀 다운로드
        # =============================================

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

        # =============================================
        # 비교지도
        # =============================================

        st.subheader("지도 비교")

        course_count = len(result.get("compare_maps", []))

        if course_count == 0:
            st.warning("생성된 지도가 없습니다.")
        else:
            course_tabs = st.tabs([f"코스{i+1}" for i in range(course_count)])

            for i, tab in enumerate(course_tabs):
                with tab:
                    map_tabs = st.tabs(["원본", "최적화", "비교"])

                    with map_tabs[0]:
                        with open(result["original_maps"][i], "r", encoding="utf-8") as f:
                            original_html = f.read()
                        components.html(original_html, height=800)
                        st.download_button(
                            label=f"📥 코스{i+1} 원본 지도 다운로드",
                            data=original_html.encode("utf-8"),
                            file_name=os.path.basename(result["original_maps"][i]),
                            mime="text/html",
                            key=f"orig_{i}",
                        )

                    with map_tabs[1]:
                        with open(result["optimized_maps"][i], "r", encoding="utf-8") as f:
                            optimized_html = f.read()
                        components.html(optimized_html, height=800)
                        st.download_button(
                            label=f"📥 코스{i+1} 최적화 지도 다운로드",
                            data=optimized_html.encode("utf-8"),
                            file_name=os.path.basename(result["optimized_maps"][i]),
                            mime="text/html",
                            key=f"opt_{i}",
                        )

                    with map_tabs[2]:
                        with open(result["compare_maps"][i], "r", encoding="utf-8") as f:
                            compare_html = f.read()
                        components.html(compare_html, height=800)
                        st.download_button(
                            label=f"📥 코스{i+1} 비교 지도 다운로드",
                            data=compare_html.encode("utf-8"),
                            file_name=os.path.basename(result["compare_maps"][i]),
                            mime="text/html",
                            key=f"cmp_{i}",
                        )

            # =============================================
            # 지도 전체 ZIP 다운로드
            # =============================================

            zip_buffer = io.BytesIO()

            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
                for path in result["original_maps"]:
                    zf.write(path, os.path.basename(path))
                for path in result["optimized_maps"]:
                    zf.write(path, os.path.basename(path))
                for path in result["compare_maps"]:
                    zf.write(path, os.path.basename(path))

            st.download_button(
                label="🗺️ 지도 전체 다운로드 (ZIP)",
                data=zip_buffer.getvalue(),
                file_name="최적화_지도.zip",
                mime="application/zip",
                use_container_width=True,
            )
