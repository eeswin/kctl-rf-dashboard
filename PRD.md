# [PRD] Streamlit 기반 RF 계측 자동 분석 & 워드 성적서 발행 대시보드

## 1. 프로젝트 개요 및 자산
- **과제명**: 유로핀즈KCTL AI PBL - 스펙트럼 분석기 이미지 인식 및 방사성 방출(RE) 성적서 자동 발행 웹앱
- **실행 파일**: `app.py` (단일 메인 애플리케이션)
- **자산 파일**: 
  - 서식 템플릿: `report_template.docx`
  - 계측 이미지: `img_1.png` (2.405GHz / 641.83uW), `img_2.png` (2.440GHz / 477.12uW), `img_3.png` (2.480GHz / 378.31uW)

---

## 2. 단계별 세부 구현 명세

### [STEP 1] 프로젝트 초기화 & requirements.txt
- Python 가상환경 `.venv` 생성
- 필수 라이브러리 `requirements.txt` 작성:
  `streamlit>=1.35.0`, `openai>=1.30.0`, `plotly>=5.18.0`, `pandas>=2.0.0`, `docxtpl>=0.16.0`, `python-dotenv>=1.0.0`, `Pillow>=10.0.0`
- 패키지 일괄 설치 완료

### [STEP 2] Streamlit 기본 레이아웃 & 사이드바 파라미터
1. **페이지 설정**: `st.set_page_config(page_title="Eurofins KCTL RF Inspector", page_icon="📡", layout="wide")`
2. **사이드바 구성 (`st.sidebar`)**:
   - 제목: "📡 KCTL RF 계측 Inspector"
   - OpenAI API Key 입력: `st.text_input("OpenAI API Key", type="password", value=os.getenv("OPENAI_API_KEY", ""))`
   - 모델 안내: `st.caption("AI Engine: OpenAI 5.6 Luna Vision")`
   - 시험 담당자: 기본값 "홍길동 선임연구원"
   - 기술 검토자: 기본값 "김선임 기술책임자"
   - 시료명(EUT): 기본값 "EUT-2026-BLE-MODULE"
   - 시험 규격 선택: `st.selectbox` ("FCC Part 15 Subpart B Class B (3 m)", "CISPR 32 Class B (3 m)", "KN 32")
   - 안테나 보정계수 (CF): 슬라이더 범위 15.0 ~ 40.0 dB, 기본값 28.5 dB, step 0.5
   - 규격 기준치 (Limit): 슬라이더 범위 120.0 ~ 150.0 dBµV/m, 기본값 140.0 dBµV/m, step 1.0
3. **메인 상단 구성**:
   - 대제목 및 설명 배너
   - 다중 이미지 업로더: `st.file_uploader(accept_multiple_files=True, type=['png', 'jpg'])`
   - 원클릭 샘플 로드 버튼: "📂 기본 샘플 3종(img_1~3) 일괄 불러오기"

### [STEP 3] OpenAI 5.6 Luna Vision 마커 파싱 & RF 계산 엔진
1. **Vision 파싱 모듈**:
   - 함수: `extract_marker_data(image_bytes: bytes, api_key: str) -> dict`
   - 데코레이터: `@st.cache_data` (이미지 바이트 기반 캐싱)
   - 모델: `gpt-5.6-luna` (OpenAI 5.6 Luna)
   - 타겟 영역: 스펙트럼 분석기 화면 우측 상단 녹색 텍스트 `Mkr1` 영역
   - JSON 반환 스키마:
     `{"freq_val": 2.405095, "freq_unit": "GHz", "power_val": 641.83, "power_unit": "uW"}`
2. **RF 물리량 계산 수식**:
   - 주파수($f_{\text{MHz}}$): 단위가 `GHz`이면 $\times 1000$, `MHz`이면 그대로 유지
   - 전력($P_{\text{dBm}}$): $10 \times \log_{10}(P_{\mu\text{W}} / 1000)$
   - 측정 전계강도($E_{\text{dB}\mu\text{V/m}}$): $P_{\text{dBm}} + 107.0 + CF$
   - 마진($\text{Margin}$): $\text{Limit} - E_{\text{dB}\mu\text{V/m}}$
   - 개별 판정: $\text{Margin} \ge 0 \rightarrow \text{"PASS"}$, $\text{Margin} < 0 \rightarrow \text{"FAIL"}$
   - 모든 수치는 소수점 둘째 자리까지 반올림(`round(val, 2)`)

### [STEP 4] KPI 요약 카드 & Plotly 인터랙티브 규격 차트
1. **상단 KPI 메트릭 카드 (`st.columns(4)`)**:
   - ① 총 측정 건수 (`len(results)`)
   - ② PASS 건수 (초록색 델타)
   - ③ FAIL 건수 (빨간색 델타)
   - ④ 종합 판정: FAIL 건수가 0이면 초록색 `PASS`, 1건 이상이면 빨간색 `FAIL`
2. **Plotly 규격 비교 차트**:
   - X축: 주파수 (`freq_mhz`, 단위 MHz)
   - Y축: 측정 전계강도 (`dbuv_m`, 단위 dBµV/m)
   - 막대 색상: PASS 항목은 `#003399`(KCTL 블루), FAIL 항목은 `#E03A3A`(레드)
   - 기준선: 사이드바의 `Limit` 값을 Y축 수평 점선(빨간 대시선)으로 그리고 어노테이션(`FCC Limit 기준`) 추가
   - 툴팁: 마우스 오버 시 주파수, 측정값, 여유 마진 상세 표기
3. **판정 결과 데이터프레임 (`st.dataframe`)**:
   - 컬럼: No, 이미지명, 주파수(MHz), 측정전력(µW), dBm, 측정값(dBµV/m), Limit, 마진(dB), 판정

### [STEP 5] docxtpl 워드 성적서 다운로드 연동
1. **메모리 버퍼 기반 워드 렌더링**:
   - `docxtpl.DocxTemplate("report_template.docx")` 로드
   - `io.BytesIO` 메모리 스트림을 열어 `doc.save(buffer)` 실행 후 `buffer.seek(0)` 호출
2. **컨텍스트 변수 맵핑 규격**:
   - 메타: `doc_no`, `test_date`, `test_name`("방사성 방출 (RE) 측정 결과 보고서"), `tester`, `standard`, `sample_name`, `cf_db`, `reviewer`, `remarks`, `generated_at`
   - 요약: `total`, `pass_cnt`, `fail_cnt`, `verdict`
   - 반복 행 데이터 (`r`): `no`, `freq_mhz`, `power_uw`, `dbm`, `dbuv_m`, `limit`, `margin`, `verdict`
3. **다운로드 위젯**:
   - `st.download_button(label="📥 공식 시험성적서(.docx) 다운로드", data=buffer, file_name=f"KCTL_RE_Report_{date}.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")`