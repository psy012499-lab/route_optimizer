import streamlit as st
import io
import os
import zipfile
import json
import pandas as pd
import requests as _requests

from corse7_optimizer import process_excel as process7
from corse8_optimizer import process_excel as process8


# =====================================================
# Claude AI — 자연어 파라미터 파싱
# =====================================================

ANTHROPIC_API_URL = "https://api.anthropic.com/v1/messages"
HAIKU_MODEL       = "claude-haiku-4-5-20251001"

DEFAULT_PARAMS = {
    "algorithm":            None,   # None = UI 라디오 값 사용
    "cost_per_km":          925,
    "labor_cost_per_hour":  17000,
    "working_days":         21,
    "target_courses":       None,   # None = 전체 코스
}

def call_claude(system: str, user: str, max_tokens: int = 800) -> str:
    headers = {
        "Content-Type": "application/json",
        "anthropic-version": "2023-06-01",
    }
    payload = {
        "model": HAIKU_MODEL,
        "max_tokens": max_tokens,
        "system": system,
        "messages": [{"role": "user", "content": user}],
    }
    try:
        resp = _requests.post(ANTHROPIC_API_URL, headers=headers, json=payload, timeout=30)
        resp.raise_for_status()
        return resp.json()["content"][0]["text"]
    except Exception as e:
        return f"[API 오류] {e}"


def parse_params_from_nl(user_input: str, current_params: dict) -> tuple[dict, str]:
    """
    자연어 입력 → 파라미터 dict + 사람이 읽을 수 있는 변경 요약 반환
    """
    system = """당신은 우편 집배 경로 최적화 시스템의 파라미터 파서입니다.
사용자의 자연어 지시에서 변경할 파라미터를 추출하세요.

추출 가능한 파라미터:
- algorithm: "corse7" 또는 "corse8" (언급 없으면 null)
- cost_per_km: 정수 (원/km, 언급 없으면 null)
- labor_cost_per_hour: 정수 (원/시간, 언급 없으면 null)
- working_days: 정수 (월 근무일, 언급 없으면 null)
- target_courses: 정수 배열 (특정 코스 번호들, 전체이면 null)

반드시 아래 JSON 형식만 출력하세요. 설명 없이 JSON만:
{
  "params": {
    "algorithm": null,
    "cost_per_km": null,
    "labor_cost_per_hour": null,
    "working_days": null,
    "target_courses": null
  },
  "summary": "변경사항을 한 문장으로 설명"
}"""

    user = f"""현재 파라미터:
- 알고리즘: {current_params.get('algorithm', 'UI 선택값')}
- 운송비: {current_params['cost_per_km']}원/km
- 인건비: {current_params['labor_cost_per_hour']}원/시간
- 월 근무일: {current_params['working_days']}일
- 대상 코스: {current_params.get('target_courses') or '전체'}

사용자 지시: {user_input}"""

    raw = call_claude(system, user, max_tokens=400)

    try:
        # JSON 블록만 추출
        raw_clean = raw.strip()
        if "```" in raw_clean:
            raw_clean = raw_clean.split("```")[1]
            if raw_clean.startswith("json"):
                raw_clean = raw_clean[4:]
        parsed = json.loads(raw_clean)
        extracted = parsed.get("params", {})
        summary   = parsed.get("summary", "파라미터가 업데이트되었습니다.")

        # 기존 파라미터에 null 아닌 값만 덮어쓰기
        new_params = current_params.copy()
        for key, val in extracted.items():
            if val is not None:
                new_params[key] = val

        return new_params, summary

    except Exception:
        return current_params, f"파싱 실패 — 원본 응답: {raw[:200]}"


# =====================================================
# Claude AI — 결과 자연어 해석
# =====================================================

def build_summary_text(summary: pd.DataFrame, algorithm_name: str, params: dict) -> str:
    try:
        saving_row = summary[summary["구분"] == "절감효과"]
        if saving_row.empty:
            saving_row = summary.iloc[-1:]

        lines = [
            f"알고리즘: {algorithm_name}",
            f"운송비 단가: {params['cost_per_km']}원/km",
            f"인건비 단가: {params['labor_cost_per_hour']}원/시간",
            f"월 근무일: {params['working_days']}일",
        ]
        for col in summary.columns:
            if col == "구분":
                continue
            orig_val  = summary[summary["구분"] == "원본"][col].values
            opt_val   = summary[summary["구분"] == "최적화"][col].values
            saved_val = saving_row[col].values
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
    return call_claude(system, user, max_tokens=800)


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
.metric-sub   { font-size: 11px; color: #888; margin-top: 2px; }
.metric-value.highlight { color: #E67E22; }
.metric-value.green     { color: #2e7d32; }
.stDataFrame { overflow-x: auto !important; }
.stDataFrame table { font-size: 12px !important; }
.stDataFrame th, .stDataFrame td { padding: 4px 8px !important; white-space: nowrap; }
.stButton > button { min-height: 48px; font-size: 15px; }
.stDownloadButton > button { min-height: 44px; font-size: 14px; }
.stTabs [data-baseweb="tab"] { font-size: 13px; padding: 6px 12px; }
/* 파라미터 변경 박스 */
.param-box {
    background: #f8f4ff;
    border-left: 4px solid #7F77DD;
    border-radius: 0 8px 8px 0;
    padding: 12px 16px;
    margin: 8px 0 12px;
    font-size: 13px;
    line-height: 1.7;
    color: #26215C;
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
/* AI 해석 박스 */
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
                if "천원" in col_name:
                    return f"{fval:,.0f}천원"
                else:
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

if "run_params" not in st.session_state:
    st.session_state["run_params"] = DEFAULT_PARAMS.copy()
if "param_log" not in st.session_state:
    st.session_state["param_log"] = []
if "ai_interpretation" not in st.session_state:
    st.session_state["ai_interpretation"] = None


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
# 📂 파일 업로드  +  ⚙️ 기본값 설정  (나란히 배치)
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

    p = st.session_state["run_params"]

    ui_cost_per_km = st.number_input(
        "운송비 (원 / km)",
        min_value=100,
        max_value=10000,
        value=int(p["cost_per_km"]),
        step=25,
        help="km당 운송 단가",
    )
    ui_labor_cost = st.number_input(
        "인건비 (원 / 시간)",
        min_value=1000,
        max_value=100000,
        value=int(p["labor_cost_per_hour"]),
        step=500,
        help="시간당 인건비 단가",
    )
    ui_working_days = st.number_input(
        "근무일수 (일 / 월)",
        min_value=1,
        max_value=31,
        value=int(p["working_days"]),
        step=1,
        help="월 평균 근무일수",
    )

    # UI 입력값 즉시 반영
    st.session_state["run_params"]["cost_per_km"]         = ui_cost_per_km
    st.session_state["run_params"]["labor_cost_per_hour"] = ui_labor_cost
    st.session_state["run_params"]["working_days"]        = ui_working_days

st.markdown("---")

# =====================================================
# 🤖 AI 자연어 파라미터 조정 (Agent 입력창)
# =====================================================

st.markdown("##### 🤖 AI 자연어 파라미터 조정")
st.caption("위 기본값 외에 추가 조건을 자연어로 입력 · 예: \"CORSE8로 바꾸고 코스 1번만 최적화해줘\"")

nl_input = st.text_input(
    label="파라미터 지시",
    placeholder="예) CORSE8로 바꾸고 코스 1번만 최적화해줘",
    label_visibility="collapsed",
)

col_apply, col_reset = st.columns([1, 1])
with col_apply:
    apply_btn = st.button("✨ AI 파라미터 적용", use_container_width=True, type="primary")
with col_reset:
    reset_btn = st.button("↺ 기본값으로 초기화", use_container_width=True)

if reset_btn:
    st.session_state["run_params"] = DEFAULT_PARAMS.copy()
    st.session_state["param_log"]  = []
    st.session_state["ai_interpretation"] = None
    st.rerun()

if apply_btn and nl_input.strip():
    with st.spinner("Claude AI가 지시를 분석 중..."):
        new_params, summary_msg = parse_params_from_nl(
            nl_input,
            st.session_state["run_params"]
        )
    st.session_state["run_params"] = new_params
    st.session_state["param_log"].append(summary_msg)
    st.session_state["ai_interpretation"] = None
    st.rerun()

# 현재 적용 파라미터 태그 요약
p = st.session_state["run_params"]
algo_display    = p["algorithm"] or ("corse7" if ui_algorithm.startswith("corse7") else "corse8")
courses_display = f"코스 {p['target_courses']}" if p.get("target_courses") else "전체 코스"

tags_html = (
    f'<span class="param-tag">알고리즘: {algo_display.upper()}</span>'
    f'<span class="param-tag">운송비: {p["cost_per_km"]:,}원/km</span>'
    f'<span class="param-tag">인건비: {p["labor_cost_per_hour"]:,}원/h</span>'
    f'<span class="param-tag">월 근무일: {p["working_days"]}일</span>'
    f'<span class="param-tag">대상: {courses_display}</span>'
)
log_html = ""
if st.session_state["param_log"]:
    last = st.session_state["param_log"][-1]
    log_html = f'<div style="margin-top:8px; font-size:12px; color:#534AB7;">📝 {last}</div>'

st.markdown(
    f'<div class="param-box">{tags_html}{log_html}</div>',
    unsafe_allow_html=True,
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

        # 최종 파라미터 결정
        final_params = st.session_state["run_params"].copy()
        if final_params["algorithm"] is None:
            final_params["algorithm"] = "corse7" if ui_algorithm.startswith("corse7") else "corse8"

        algo_label = (
            "CORSE7 (통상순로 최적화)"
            if final_params["algorithm"] == "corse7"
            else "CORSE8 (통상순로+코스 최적화)"
        )

        # optimizer에 파라미터 주입
        import corse7_optimizer as c7
        import corse8_optimizer as c8
        for mod in (c7, c8):
            mod.COST_PER_KM          = final_params["cost_per_km"]
            mod.LABOR_COST_PER_HOUR  = final_params["labor_cost_per_hour"]

        moto_bar = st.empty()

        def update_progress(pct, msg):
            pct = max(0, min(100, pct))
            total  = 20
            filled = int(pct / 100 * total)
            empty  = total - filled - 1
            bar    = "━" * filled + "🏍️" + "─" * empty
            moto_bar.markdown(f"**{bar}** `{pct}%`", unsafe_allow_html=True)

        update_progress(0, "")

        with st.spinner("AI 최적화 진행중..."):
            process_fn = process7 if final_params["algorithm"] == "corse7" else process8
            result = process_fn(
                target_file,
                progress_callback=update_progress,
                target_courses=final_params.get("target_courses"),
                working_days=int(final_params["working_days"]),
            )

        moto_bar.markdown("**━━━━━━━━━━━━━━━━━━━🏁** `완료!`")
        st.success("🎉 최적화 완료!")

        # ── 적용된 파라미터 확인 배너 ─────────────────────────────
        courses_banner = "전체 코스" if not final_params.get("target_courses") else f"코스 {final_params['target_courses']}"
        st.markdown(
            f'<div class="param-box" style="margin-bottom:16px;">'
            f'<b>실행 파라미터</b> &nbsp;'
            f'<span class="param-tag">{final_params["algorithm"].upper()}</span>'
            f'<span class="param-tag">운송비 {final_params["cost_per_km"]:,}원/km</span>'
            f'<span class="param-tag">인건비 {final_params["labor_cost_per_hour"]:,}원/h</span>'
            f'<span class="param-tag">월 {final_params["working_days"]}일</span>'
            f'<span class="param-tag">대상: {courses_banner}</span>'
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

        # ── AI 결과 해석 ──────────────────────────────────────────
        st.markdown("#### 🤖 AI 결과 해석")

        col_btn, _ = st.columns([1, 2])
        with col_btn:
            gen_btn = st.button("✨ AI 해석 생성", type="primary", use_container_width=True)

        if gen_btn:
            with st.spinner("Claude AI가 결과를 분석 중입니다..."):
                summary_text   = build_summary_text(summary, algo_label, final_params)
                interpretation = call_claude_interpretation(summary_text, algo_label)
                st.session_state["ai_interpretation"] = interpretation

        if st.session_state.get("ai_interpretation"):
            st.markdown(
                f'<div class="ai-box">'
                f'{st.session_state["ai_interpretation"].replace(chr(10), "<br>")}'
                f'</div>',
                unsafe_allow_html=True,
            )
            st.download_button(
                label="📄 AI 해석 텍스트 저장",
                data=st.session_state["ai_interpretation"],
                file_name="AI_결과해석.txt",
                mime="text/plain",
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
