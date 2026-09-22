# Eurofins KCTL RF 계측 자동화 대시보드 - Claude Code 행동 헌법 (CLAUDE.md)

## 1. 프로젝트 목적 및 철학
- 본 시스템은 유로핀즈KCTL의 인증 시험 연구원이 브라우저에서 스펙트럼 분석기(Agilent Swept SA) 이미지를 분석하고 공식 워드 성적서를 발행하는 순수 Python 기반 Streamlit 웹 애플리케이션이다.
- 모든 비즈니스 로직, 수식 계산, UI 명세는 반드시 프로젝트 루트의 `PRD.md`를 최우선 기준으로 준수한다.
- 사용자가 "PRD.md의 [STEP N]을 구현해줘"와 같이 최소한의 원라인(One-line) 지시만 내리더라도, PRD.md의 상세 명세를 파악하여 누락 없이 완전하고 실행 가능한 코드를 작성한다.

## 2. 기술 스택 & 필수 종속성
- Frontend / Framework: `streamlit>=1.35.0`
- AI Vision Engine: `openai>=1.30.0` (모델: `5.6 luna` / 코드 표기: `gpt-5.6-luna`)
- Data & Visualization: `pandas>=2.0.0`, `plotly>=5.18.0`
- Document Rendering: `docxtpl>=0.16.0` (python-docx 인메모리 버퍼 렌더링)
- Environment: `python-dotenv>=1.0.0`, `Pillow>=10.0.0`

## 3. OpenAI 5.6 Luna Vision API 구현 규격
- **API 클라이언트 초기화**:
  - `openai.OpenAI(api_key=...)`를 사용한다.
  - API 키는 Streamlit 사이드바 인풋 필드(`st.sidebar.text_input(type="password")`)와 `.env` 파일의 `OPENAI_API_KEY` 환경 변수를 상호 연동한다.
- **모델 지정**:
  - 비전 추출 호출 시 `model="gpt-5.6-luna"`를 명시한다.
- **이미지 전달 포맷**:
  - 업로드된 이미지 파일(또는 로컬 파일)은 base64 문자열로 인코딩하여 Data URL(`data:image/png;base64,{base64_str}`) 형태로 `messages` 인자에 주입한다.
- **정형 JSON 응답 강제**:
  - 프롬프트에 마커 정보(Mkr1 주파수 값/단위, 전력 값/단위)를 반드시 JSON 오브젝트(`response_format={"type": "json_object"}`) 형태로 출력받아 파싱 실패율을 0%로 유지한다.

## 4. Streamlit 상태 관리 & 비용 최적화 원칙
- **API 호출 캐싱 (`@st.cache_data`)**:
  - 슬라이더(보정계수 CF, 규격 한계 Limit)를 조작할 때 Streamlit의 전체 재실행(Rerun) 특성으로 인해 OpenAI API가 중복 호출되는 것을 원천 차단한다.
  - 이미지 바이트(bytes)를 인자로 받는 Vision 추출 함수에 반드시 `@st.cache_data` 데코레이터를 적용한다.
- **세션 상태 관리 (`st.session_state`)**:
  - 분석 완료된 결과 데이터는 `st.session_state['analysis_results']`에 저장하여 화면 재렌더링 시에도 즉시 조회되도록 한다.

## 5. 인메모리 파일 I/O 규칙
- 워드 성적서(`report_template.docx`) 렌더링 시 서버 디스크에 임시 파일을 쓰지 않는다.
- `io.BytesIO()` 버퍼 객체를 생성하고 `doc.save(buffer)` 호출 후, **반드시 `buffer.seek(0)`**을 수행하여 파일 포인터를 원점으로 되감은 뒤 `st.download_button`의 `data`로 전달한다.

## 6. 에러 핸들링
- API 키 미입력 시 친절한 `st.warning("사이드바에 OpenAI API Key를 입력하세요.")` 표기 후 `st.stop()` 처리.
- 모든 예외는 `try-except`로 포착하여 `st.error(f"오류 내용: {e}")`로 처리하고 앱이 크래시되지 않게 보호한다.