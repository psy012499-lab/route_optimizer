import streamlit as st
import streamlit.components.v1 as components

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
# 파일 업로드
# =====================================================

uploaded_file = st.file_uploader(
    "엑셀 파일 업로드",
    type=["xlsx"],
)


# =====================================================
# 실행
# =====================================================

if uploaded_file is not None:

    st.success("엑셀 업로드 완료")

    if st.button("최적화 실행", type="primary", use_container_width=True):

        with st.spinner("AI 최적화 진행중..."):

            if algorithm.startswith("corse7"):
                result = process7(uploaded_file)
            else:
                result = process8(uploaded_file)

        st.success("최적화 완료")

        # =============================================
        # 결과 요약
        # =============================================

        st.subheader("요약 결과")
        st.dataframe(result["summary"], use_container_width=True)

        # =============================================
        # 엑셀 다운로드
        # =============================================

        import io
        import pandas as pd

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
                            components.html(f.read(), height=800)

                    with map_tabs[1]:
                        with open(result["optimized_maps"][i], "r", encoding="utf-8") as f:
                            components.html(f.read(), height=800)

                    with map_tabs[2]:
                        with open(result["compare_maps"][i], "r", encoding="utf-8") as f:
                            components.html(f.read(), height=800)