import base64
import copy
import io
import json
import math
import os
from datetime import datetime
from pathlib import Path

import openai
import pandas as pd
import plotly.graph_objects as go
import streamlit as st
from docx.oxml.ns import qn
from docxtpl import DocxTemplate
from dotenv import load_dotenv
from PIL import Image

BASE_DIR = Path(__file__).resolve().parent
TEMPLATE_PATH = BASE_DIR / "report_template.docx"
PASS_COLOR = "#003399"
FAIL_COLOR = "#E03A3A"
SAMPLE_FILES = ["img_1.png", "img_2.png", "img_3.png"]
STANDARDS = [
    "FCC Part 15 Subpart B Class B (3 m)",
    "CISPR 32 Class B (3 m)",
    "KN 32",
]

VISION_MODEL = "gpt-5.6-luna"
FREQ_TO_MHZ = {"GHz": 1000.0, "MHz": 1.0, "kHz": 1e-3, "Hz": 1e-6}
POWER_TO_UW = {"pW": 1e-6, "nW": 1e-3, "uW": 1.0, "mW": 1e3, "W": 1e6}
POWER_UNITS = [*POWER_TO_UW, "dBm"]

MARKER_PROMPT = """This is a screenshot of an Agilent/Keysight spectrum analyzer (Swept SA).
Read ONLY the green marker readout labelled "Mkr1" at the top-right of the trace area.
It has two lines: the marker frequency (e.g. "Mkr1 2.405 095 GHz") and the marker power (e.g. "641.83 µW").
Ignore the "Ref" level, the grid axis labels, and the Center/Start/Stop Freq values in the right-hand menu.

Return a JSON object with exactly these keys:
- "freq_val": number, with digit-grouping spaces removed ("2.405 095" -> 2.405095)
- "freq_unit": one of "GHz", "MHz", "kHz", "Hz"
- "power_val": number
- "power_unit": one of "pW", "nW", "uW", "mW", "W", "dBm" (write µW as "uW")

Example: {"freq_val": 2.405095, "freq_unit": "GHz", "power_val": 641.83, "power_unit": "uW"}"""

load_dotenv(BASE_DIR / ".env")


# ─────────────────────────────────────────────
# [STEP 3] Vision 마커 파싱 & RF 계산 엔진
# ─────────────────────────────────────────────
def _normalize_unit(unit: str, allowed: list[str]) -> str:
    text = str(unit).strip().replace("µ", "u").replace("μ", "u")
    for key in allowed:
        if text.lower() == key.lower():
            return key
    raise ValueError(f"지원하지 않는 단위입니다: {unit!r}")


@st.cache_data(show_spinner="OpenAI 5.6 Luna Vision이 Mkr1 마커를 판독하는 중...")
def extract_marker_data(image_bytes: bytes, api_key: str) -> dict:
    mime = Image.MIME[Image.open(io.BytesIO(image_bytes)).format]
    data_url = f"data:{mime};base64,{base64.b64encode(image_bytes).decode('ascii')}"

    client = openai.OpenAI(api_key=api_key)
    response = client.chat.completions.create(
        model=VISION_MODEL,
        response_format={"type": "json_object"},
        messages=[
            {
                "role": "user",
                "content": [
                    {"type": "text", "text": MARKER_PROMPT},
                    {"type": "image_url", "image_url": {"url": data_url}},
                ],
            }
        ],
    )
    raw = json.loads(response.choices[0].message.content)

    # 잘못된 응답은 예외로 올려서 캐시에 저장되지 않게 한다.
    return {
        "freq_val": float(raw["freq_val"]),
        "freq_unit": _normalize_unit(raw["freq_unit"], list(FREQ_TO_MHZ)),
        "power_val": float(raw["power_val"]),
        "power_unit": _normalize_unit(raw["power_unit"], POWER_UNITS),
    }


def calc_rf_metrics(marker: dict, cf_db: float, limit: float) -> dict:
    freq_mhz = marker["freq_val"] * FREQ_TO_MHZ[marker["freq_unit"]]
    if marker["power_unit"] == "dBm":
        power_uw = 10 ** (marker["power_val"] / 10) * 1000
    else:
        power_uw = marker["power_val"] * POWER_TO_UW[marker["power_unit"]]
    if power_uw <= 0:
        raise ValueError(f"전력 값이 0 이하입니다: {marker['power_val']} {marker['power_unit']}")

    # 성적서의 수치가 서로 맞아떨어지도록 반올림된 값으로 다음 단계를 계산한다.
    dbm = round(10 * math.log10(power_uw / 1000), 2)
    dbuv_m = round(dbm + 107.0 + cf_db, 2)
    margin = round(limit - dbuv_m, 2)
    return {
        "freq_mhz": round(freq_mhz, 2),
        "power_uw": round(power_uw, 2),
        "dbm": dbm,
        "dbuv_m": dbuv_m,
        "limit": round(limit, 2),
        "margin": margin,
        "verdict": "PASS" if margin >= 0 else "FAIL",
    }


# ─────────────────────────────────────────────
# [STEP 5] docxtpl 워드 성적서 렌더링 (인메모리)
# ─────────────────────────────────────────────
def _add_result_row_loop(doc: DocxTemplate) -> None:
    """결과 표의 {{r.*}} 행을 {%tr for r in rows %} 반복 구간으로 감싼다.

    제공된 템플릿에는 반복 태그가 없어 한 행만 채워지므로, 파일은 그대로 두고 메모리에서만 보강한다.
    """
    docx = doc.get_docx()
    if "for r in" in doc.get_xml():
        return
    for table in docx.tables:
        for row in table.rows:
            if "r.no" not in "".join(cell.text for cell in row.cells):
                continue
            start_row, end_row = copy.deepcopy(row._tr), copy.deepcopy(row._tr)
            for tag_row, tag in ((start_row, "{%tr for r in rows %}"), (end_row, "{%tr endfor %}")):
                texts = list(tag_row.iter(qn("w:t")))
                texts[0].text = tag
                for t in texts[1:]:
                    t.text = ""
            row._tr.addprevious(start_row)
            row._tr.addnext(end_row)
            return
    raise ValueError("템플릿에서 {{r.no}} 결과 행을 찾지 못했습니다.")


def build_report_docx(context: dict) -> io.BytesIO:
    doc = DocxTemplate(str(TEMPLATE_PATH))
    _add_result_row_loop(doc)
    doc.render(context)
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# ─────────────────────────────────────────────
# [STEP 2] 페이지 설정 & 사이드바 파라미터
# ─────────────────────────────────────────────
st.set_page_config(page_title="Eurofins KCTL RF Inspector", page_icon="📡", layout="wide")

st.session_state.setdefault("sample_images", [])
st.session_state.setdefault("analysis_results", [])

with st.sidebar:
    st.title("📡 KCTL RF 계측 Inspector")
    api_key = st.text_input("OpenAI API Key", type="password", value=os.getenv("OPENAI_API_KEY", ""))
    st.caption("AI Engine: OpenAI 5.6 Luna Vision")

    st.divider()
    tester = st.text_input("시험 담당자", value="홍길동 선임연구원")
    reviewer = st.text_input("기술 검토자", value="김선임 기술책임자")
    sample_name = st.text_input("시료명(EUT)", value="EUT-2026-BLE-MODULE")
    standard = st.selectbox("시험 규격", STANDARDS)

    st.divider()
    cf_db = st.slider("안테나 보정계수 (CF) [dB]", min_value=15.0, max_value=40.0, value=28.5, step=0.5)
    limit_dbuv = st.slider("규격 기준치 (Limit) [dBµV/m]", min_value=120.0, max_value=150.0, value=140.0, step=1.0)

# ─────────────────────────────────────────────
# [STEP 2] 메인 상단: 배너 & 이미지 입력
# ─────────────────────────────────────────────
st.title("📡 Eurofins KCTL RF 계측 자동 분석 대시보드")
st.markdown(
    """
    <div style="background:#003399;color:#ffffff;padding:1rem 1.25rem;border-radius:8px;margin-bottom:1rem;">
      <b>방사성 방출(RE) 스펙트럼 분석기 이미지 자동 판독 &amp; 공식 성적서 발행</b><br>
      Agilent Swept SA 캡처 이미지를 올리면 AI가 Mkr1 마커를 판독하고,
      보정계수(CF)와 규격 기준치(Limit)를 적용해 PASS/FAIL 판정과 워드 성적서를 만듭니다.
    </div>
    """,
    unsafe_allow_html=True,
)

col_upload, col_sample = st.columns([3, 1])
with col_upload:
    uploaded_files = st.file_uploader(
        "스펙트럼 분석기 캡처 이미지 업로드 (PNG/JPG, 여러 장 선택 가능)",
        accept_multiple_files=True,
        type=["png", "jpg"],
    )
with col_sample:
    if st.button("📂 기본 샘플 3종(img_1~3) 일괄 불러오기"):
        try:
            st.session_state["sample_images"] = [
                {"name": name, "bytes": (BASE_DIR / name).read_bytes()} for name in SAMPLE_FILES
            ]
        except Exception as e:
            st.error(f"오류 내용: {e}")

# 업로드 파일과 샘플을 합치되, 같은 파일명은 업로드 쪽을 우선한다.
images = [{"name": f.name, "bytes": f.getvalue()} for f in uploaded_files or []]
uploaded_names = {img["name"] for img in images}
images += [img for img in st.session_state["sample_images"] if img["name"] not in uploaded_names]

if not images:
    st.info("분석할 스펙트럼 이미지를 업로드하거나 기본 샘플을 불러오세요.")
    st.stop()

st.subheader(f"🖼️ 분석 대상 이미지 ({len(images)}건)")
preview_cols = st.columns(min(len(images), 4))
for i, img in enumerate(images):
    preview_cols[i % len(preview_cols)].image(img["bytes"], caption=img["name"])

if not api_key:
    st.warning("사이드바에 OpenAI API Key를 입력하세요.")
    st.stop()

# ─────────────────────────────────────────────
# [STEP 3] 마커 판독 실행 & RF 판정
# ─────────────────────────────────────────────
# Vision 호출은 캐시되므로 CF/Limit 슬라이더를 움직여도 계산만 다시 한다.
results = []
for img in images:
    try:
        marker = extract_marker_data(img["bytes"], api_key)
        results.append(
            {
                "no": len(results) + 1,
                "image_name": img["name"],
                **marker,
                **calc_rf_metrics(marker, cf_db, limit_dbuv),
            }
        )
    except Exception as e:
        st.error(f"오류 내용: [{img['name']}] {e}")

st.session_state["analysis_results"] = results
if not results:
    st.stop()

st.success(f"✅ {len(results)}건 마커 판독 완료 (CF {cf_db} dB, Limit {limit_dbuv} dBµV/m 적용)")

# ─────────────────────────────────────────────
# [STEP 4] KPI 요약 카드 & Plotly 규격 비교 차트
# ─────────────────────────────────────────────
total = len(results)
pass_cnt = sum(r["verdict"] == "PASS" for r in results)
fail_cnt = total - pass_cnt
overall_verdict = "PASS" if fail_cnt == 0 else "FAIL"

st.divider()
st.subheader("📊 판정 요약")
kpi_total, kpi_pass, kpi_fail, kpi_verdict = st.columns(4)
kpi_total.metric("총 측정 건수", f"{total}건")
kpi_pass.metric("PASS 건수", f"{pass_cnt}건", delta=f"{pass_cnt / total:.0%} 적합", delta_color="normal")
kpi_fail.metric("FAIL 건수", f"{fail_cnt}건", delta=f"{fail_cnt / total:.0%} 부적합", delta_color="inverse")
kpi_verdict.metric("종합 판정", ":green[PASS]" if overall_verdict == "PASS" else ":red[FAIL]")

try:
    df = pd.DataFrame(results)
    df["freq_label"] = df["freq_mhz"].map("{:.2f}".format)

    fig = go.Figure()
    for verdict, color in (("PASS", PASS_COLOR), ("FAIL", FAIL_COLOR)):
        part = df[df["verdict"] == verdict]
        if part.empty:
            continue
        fig.add_bar(
            x=part["freq_label"],
            y=part["dbuv_m"],
            name=verdict,
            marker_color=color,
            text=part["dbuv_m"].map("{:.2f}".format),
            textposition="inside",
            insidetextanchor="end",
            textfont_color="white",
            customdata=part[["image_name", "margin"]],
            hovertemplate=(
                "<b>%{customdata[0]}</b><br>"
                "주파수: %{x} MHz<br>"
                "측정값: %{y:.2f} dBµV/m<br>"
                "여유 마진: %{customdata[1]:.2f} dB"
                "<extra>%{fullData.name}</extra>"
            ),
        )
    fig.add_hline(
        y=limit_dbuv,
        line_dash="dash",
        line_color="red",
        annotation_text=f"FCC Limit 기준 ({limit_dbuv:.1f} dBµV/m)",
        annotation_position="top right",
        annotation_font_color="red",
    )

    # 막대가 0부터 그려지면 차이가 보이지 않으므로 측정값·Limit 주변으로 Y축을 확대한다.
    y_low = min(df["dbuv_m"].min(), limit_dbuv)
    y_high = max(df["dbuv_m"].max(), limit_dbuv)
    fig.update_layout(
        title="주파수별 측정 전계강도 vs 규격 한계",
        xaxis_title="주파수 (MHz)",
        yaxis_title="측정 전계강도 (dBµV/m)",
        barmode="overlay",
        bargap=0.5,
        height=450,
        legend={"orientation": "h", "yanchor": "bottom", "y": 1.02, "xanchor": "right", "x": 1},
    )
    fig.update_xaxes(
        type="category",
        categoryorder="array",
        categoryarray=df.sort_values("freq_mhz")["freq_label"].tolist(),
    )
    fig.update_yaxes(range=[math.floor((y_low - 15) / 10) * 10, math.ceil((y_high + 5) / 10) * 10])

    st.subheader("📈 규격 비교 차트")
    st.plotly_chart(fig)

    st.subheader("📋 판정 결과")
    table = df[["no", "image_name", "freq_mhz", "power_uw", "dbm", "dbuv_m", "limit", "margin", "verdict"]].rename(
        columns={
            "no": "No",
            "image_name": "이미지명",
            "freq_mhz": "주파수(MHz)",
            "power_uw": "측정전력(µW)",
            "dbm": "dBm",
            "dbuv_m": "측정값(dBµV/m)",
            "limit": "Limit",
            "margin": "마진(dB)",
            "verdict": "판정",
        }
    )
    number_format = st.column_config.NumberColumn(format="%.2f")
    st.dataframe(
        table,
        hide_index=True,
        column_config={
            col: number_format
            for col in ["주파수(MHz)", "측정전력(µW)", "dBm", "측정값(dBµV/m)", "Limit", "마진(dB)"]
        },
    )
except Exception as e:
    st.error(f"오류 내용: {e}")

# ─────────────────────────────────────────────
# [STEP 5] 워드 성적서 다운로드
# ─────────────────────────────────────────────
st.divider()
st.subheader("📄 공식 시험성적서 발행")

now = datetime.now()
fails = [r for r in results if r["verdict"] == "FAIL"]
remarks = f"Limit {limit_dbuv:.1f} dBµV/m 기준, 총 {total}건 중 PASS {pass_cnt}건 / FAIL {fail_cnt}건."
if fails:
    remarks += " 부적합 주파수: " + ", ".join(f"{r['freq_mhz']:.2f} MHz (마진 {r['margin']:.2f} dB)" for r in fails)
else:
    remarks += " 전 항목이 규격 기준치 이내입니다."

report_context = {
    "doc_no": f"KCTL-RE-{now:%Y%m%d-%H%M%S}",
    "test_date": f"{now:%Y-%m-%d}",
    "test_name": "방사성 방출 (RE) 측정 결과 보고서",
    "tester": tester,
    "standard": standard,
    "sample_name": sample_name,
    "cf_db": f"{cf_db:.1f}",
    "reviewer": reviewer,
    "remarks": remarks,
    "generated_at": f"{now:%Y-%m-%d %H:%M:%S}",
    "total": total,
    "pass_cnt": pass_cnt,
    "fail_cnt": fail_cnt,
    "verdict": overall_verdict,
    "rows": [
        {
            "no": r["no"],
            "freq_mhz": f"{r['freq_mhz']:.2f}",
            "power_uw": f"{r['power_uw']:.2f}",
            "dbm": f"{r['dbm']:.2f}",
            "dbuv_m": f"{r['dbuv_m']:.2f}",
            "limit": f"{r['limit']:.2f}",
            "margin": f"{r['margin']:.2f}",
            "verdict": r["verdict"],
        }
        for r in results
    ],
}

try:
    buffer = build_report_docx(report_context)
    st.download_button(
        label="📥 공식 시험성적서(.docx) 다운로드",
        data=buffer,
        file_name=f"KCTL_RE_Report_{now:%Y%m%d}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )
except Exception as e:
    st.error(f"오류 내용: {e}")
