# Progress

작성일: 2026-02-22
브랜치: `main`

## 업데이트 (2026-02-23)
- `docs` 디렉토리 기반 HWP↔PDF 배치 비교 루프 추가
  - `scripts/preview_pdf_diff.mjs`
    - `--docs-dir <dir>`: 같은 basename의 `.hwp/.pdf`(옵션으로 `.hwpx`) 자동 페어링
    - `--case` 반복 입력 지원(멀티 케이스)
    - `--case-limit <N>` 지원
    - 케이스 파일 존재 여부 검증, 케이스 디렉토리명 안전 처리
  - `README.md`에 배치 실행 옵션 반영
- 프리뷰 페이지네이션 튜닝(표 비중 문서 과분할 완화)
  - `computeTableDensityPageScale` 보정폭 확장
  - 대표 케이스 `20250828_수요품목서 양식_KAMPA` 페이지 수: `3 -> 2`(PDF=2) 정합
- 표 렌더 개선(가독/색상 일부 반영)
  - 셀 텍스트를 문단 블록 단위로 연결(`textBlocks`)하여 표 내부 컬러 텍스트 반영 경로 추가
  - `DocInfo`의 `HWPTAG_BORDER_FILL(20)` 파싱 추가 및 테이블 셀 배경색 반영 경로 추가(조건부)

### docs 전체 1p 스모크 비교 (16 케이스)
- baseline: `artifacts/diff_docs_baseline/summary.json`
- final: `artifacts/diff_docs_final/summary.json`
- 페이지 수 정합:
  - 절대 페이지 오차 합: `177 -> 173`
  - 정확 일치 케이스: `2/16 -> 6/16`
- 대표 개선:
  - `[별첨] 산업기술분류표...`: `12 -> 10` (PDF=10)
  - `[별첨] 기술수요조사서`: `6 -> 5` (PDF=5)
  - `20250828_수요품목서 양식_KAMPA`: `3 -> 2` (PDF=2)
  - `SBA 연구지표...`: `5 -> 4` (PDF=4)
- 품질 지표(1p 평균, 16 케이스): `avgRms 50.022 -> 50.785`, `avgMae 18.797 -> 18.980` (페이지 정합 개선 대비 픽셀 오차는 소폭 악화)

### 추가 업데이트 (2026-02-23, docs 16케이스 재검증)
- 페이지 시트 스타일 정합 개선(`src/styles.css`)
  - `.preview-sheet-doc`의 카드성 스타일(추가 패딩/라운드/그림자) 제거
  - `.doc-table-wrap` 라운드 제거(테이블 외곽 직각화)
- 단일 핵심 케이스(KAMPA) 재측정
  - `artifacts/diff_one_kampa_pagecss/summary.json`
  - `avgRms 34.8176 -> 34.6200`, `avgMae 7.7835 -> 7.5550`
- docs 전체 1p(16 케이스) 재측정
  - 기준: `artifacts/diff_docs_current/summary.json`
  - 최신: `artifacts/diff_docs_try11/summary.json`
  - 페이지 수 정합: `absDeltaSum 155` 유지, `exact 8/16` 유지
  - 픽셀 지표 개선: `avgRms 51.4555 -> 51.3378`, `avgMae 19.4411 -> 19.2277`
- 실험 후 폐기
  - 표 텍스트를 문단 스타일 기반으로 강하게 반영하는 패치(폰트 크기/가중치/라인렌더)는 전체 지표 악화로 원복
  - `line-seg page-first/column-first` 비트를 상단 시작 조건으로 제한 복원하는 실험은 대표 과대/과소 케이스에서 변화가 없어 원복

### 추가 업데이트 (2026-02-23, 스펙 재검증 + 페이지네이션 보정)
- 5.0 스펙 정합 수정 반영
  - `LIST_HEADER(72)` 파싱 오프셋을 스펙 기준(6B 헤더)으로 보정
  - 객체 공통 속성 `textFlow` 값 매핑을 표 70 기준으로 수정
  - 문단 헤더 `nchars` 상위 비트(`0x80000000`) 마스킹 반영
- legacy 3.x 경로 보완
  - `legacy-3x-binary`에서 휴리스틱 이전에 레코드 파싱 우선 시도
  - preview 스트림 선택에서 legacy fallback 경로 추가
- 페이지네이션 추가 보정
  - 문단 높이 추정에서 plain-text / control / table 문단을 분리 가중
  - 표 행 높이 추정 스케일 상향(행 fallback 포함)으로 table-heavy 문서 과소분할 완화
- 재측정 결과
  - 기준: `artifacts/diff_docs_after_spec/summary.json`
  - 최신: `artifacts/diff_docs_after_linecap/summary.json`
  - 페이지 수 정합: `absDeltaSum 153 -> 129`, `exact 8/16 -> 8/16`
  - 픽셀 지표(1p 평균): `avgRms 51.3378 -> 51.3191`, `avgMae 19.2277 -> 19.2016`
  - 주요 변화
    - `연구개발계획서_마크베이스_수정_v0.6_예산제외`: `118 -> 102` (PDF=66)
    - `1. 최종보고서_(주)마크베이스(날인본)`: `260 -> 249` (PDF=194)
    - `(최종 검토본) 수정사업계획서...`: `56 -> 53` (PDF=84, 과소분할은 잔존)
    - `(주)마크베이스_보고자료_KB_251031`: `9 -> 9` (PDF=7, 악화 없음)
  - 단일 케이스 유지 확인
    - `20250828_수요품목서 양식_KAMPA`: `renderedPages=2` (PDF=2), `avgRms=34.5350`, `avgMae=7.3781`

### 추가 업데이트 (2026-02-23, 3.0/5.0 스펙 재검증 2차)
- `5.0.pdf` 핵심 표(21, 26~29, 41~47p)와 코드 매핑 재검증
  - 일치 확인:
    - 레코드 헤더 구조(4.1): `parseRecords`
    - 문단 헤더(표 58): `parseParaHeader` (`nchars` 상위비트 마스킹 포함)
    - 문단 레이아웃(표 62): `parseParaLineSeg`
    - 컨트롤 헤더(표 64): `parseCtrlHeader`
    - 문단 리스트 헤더(표 65): `parseListHeader` (6B + 셀 26B 오프셋)
    - 개체 공통 속성 비트(표 70): `decodeObjectCommonProperty`/상수 매핑
    - 표 개체 기본 구조(표 75/79/80): `parseTablePropertiesAt`
  - 부분 구현/미구현 확인:
    - `HWPTAG_BORDER_FILL(20)`의 테두리선/대각선/그라데이션/이미지 채우기 전체 해석 미구현 (`parseDocBorderFill`은 일부 필드만 추출)
    - `HWPTAG_PARA_RANGE_TAG(70)` 전용 파서 미구현
    - `HWPTAG_CTRL_DATA(87)` Parameter Set 전용 해석 미구현
    - `HWPTAG_SHAPE_COMPONENT(76)`의 Rendering matrix(표 84/85) 상세 해석 미구현 (`parseShapeComponent`은 핵심 헤더만 파싱)
- `3.0.pdf` 재검증
  - 3.x 바이너리(그리기/OLE/특수문자 구조) 전체 파서는 아직 없음(legacy heuristic 중심).
  - 현재 샘플 `3.0.hwp`/`5.0.hwp`는 모두 `HWP 5.x distributable record`로 감지되어, 실제 렌더 경로는 5.x 레코드 파서가 사용됨.

### 남은 핵심 과제
- 대용량 문서 페이지 수 오차 잔존
  - `1. 최종보고서_(주)마크베이스(날인본)`: `249 vs PDF 194`
  - `연구개발계획서_마크베이스_수정_v0.6_예산제외`: `102 vs PDF 66`
  - `(최종 검토본) 수정사업계획서...`: `53 vs PDF 84`
- 표/객체 렌더 정밀도
  - border-fill/문단 스타일의 record-level 해석 정교화 필요(현재는 제한적 추출/적용)
  - 표 내부 폰트 메트릭/줄간격/행높이 추정식 추가 튜닝 필요

## 완료된 작업
- HWP 본문 문단 컨텍스트 수집 로직 보정
  - 최상위 문단 기준으로 문단 컨텍스트를 만들고, 직접 자식 레벨의 텍스트/스타일/라인세그먼트를 매핑하도록 개선
  - 컨트롤 내부 중첩 문단이 본문으로 중복 렌더되는 케이스 완화
- PDF 비교 자동화(`diff:preview`) 안정화
  - 캡처 대상 페이지 수가 안정될 때까지 대기
  - 페이지별 `scrollIntoView` 후 이미지 로드 완료 확인 뒤 캡처
  - 산출물 디렉토리 초기화 후 재생성
  - 기본 정책: 기존 서버 재사용 금지(필요시 `--reuse-server`)
  - Vite dev server `--strictPort` 적용
- TOC(차례) 탭 리더 렌더 개선
  - 탭 스톱에서 leader 타입(dot/dash/underline) 해석
  - `doc-tab-stop`/`doc-tab-leader` 기반 리더선 렌더링
  - 탭 포함 줄은 좌측 정렬 강제하여 justify로 인한 단어 분산 완화
- 표지(커버) 텍스트박스 줄바꿈 완화(조건부)
  - 큰 제목 텍스트박스에 대해 제한적으로 폭을 확장하여 줄바꿈 압박 완화
- `src` 기능 분리 리팩토링(1차)
  - 공용 유틸 분리: `src/utils/{bytes,format,numeric,html}.js`
  - HWP 상수/사전 분리: `src/constants/hwp.js`
  - UI 셸 분리: `src/ui/appShell.js` (초기 DOM 템플릿/바인딩)
  - `src/main.js`는 오케스트레이션 중심으로 정리
- `src` 기능 분리 리팩토링(2차)
  - 문서 로딩 파이프라인 분리: `src/parser/documentLoader.js`
    - `parseFileHeader`, `tryReadCfbEntries`, `detectDocumentFormat`, `buildDocumentFromBytes`
  - 스트림 경로 판별 분리: `src/parser/streamPath.js`
  - 요약 패널 렌더 분리: `src/render/summary.js`
  - 압축 복원 유틸 분리: `src/utils/compression.js` (`safeInflate`)
  - DocInfo 패널 렌더 분리: `src/render/docInfoPanel.js`
  - Stream/Record 패널 렌더 분리: `src/render/streamPanels.js`
  - Detail 패널 렌더 분리: `src/render/detailPanel.js`
  - 스트림 분석 코어 분리: `src/parser/streamAnalyzer.js`

## 커밋 이력(이번 라운드)
- `af74f7f` Stabilize diff capture and add tab leader rendering
- `115a78b` Refine TOC tabbed line rendering for leader alignment
- `7ceb6e4` Relax cover textbox wrapping for large slash titles
- `386d892` refactor: split shared helpers into src/utils modules
- `51efe24` refactor: extract hwp constants into dedicated module
- `9c55835` refactor: split document loading into parser modules

## 검증 결과
- 실행 명령
  - `npm run build`
  - `npm run diff:preview -- --url http://127.0.0.1:4230 --max-pages 2 --python python3 --artifacts artifacts/diff_kampa_after_spec --case kampa:docs/20250828_수요품목서\ 양식_KAMPA.hwp:docs/20250828_수요품목서\ 양식_KAMPA.pdf`
  - `npm run diff:preview -- --url http://127.0.0.1:4231 --max-pages 1 --python python3 --artifacts artifacts/diff_docs_after_spec --docs-dir docs`
  - `npm run diff:preview -- --url http://127.0.0.1:4234 --max-pages 1 --python python3 --artifacts artifacts/diff_docs_after_tune --docs-dir docs`
  - `npm run diff:preview -- --url http://127.0.0.1:4236 --max-pages 1 --python python3 --artifacts artifacts/diff_docs_after_linecap --docs-dir docs`
  - `npm run diff:preview -- --url http://127.0.0.1:4208 --max-pages 20 --python /home/sjkim/work/hwpv/.venv/bin/python`
  - `npm run diff:preview -- --url http://127.0.0.1:4210 --max-pages 1 --python /home/sjkim/work/hwpv/.venv/bin/python --case 3.0:3.0.hwp:3.0.pdf`
  - `npm run diff:preview -- --url http://127.0.0.1:4211 --max-pages 1 --python /home/sjkim/work/hwpv/.venv/bin/python --case 5.0:5.0.hwp:5.0.pdf`
  - `npm run diff:preview -- --url http://127.0.0.1:4214 --max-pages 1 --python /home/sjkim/work/hwpv/.venv/bin/python --case 3.0:3.0.hwp:3.0.pdf`
  - `npm run diff:preview -- --url http://127.0.0.1:4215 --max-pages 1 --python /home/sjkim/work/hwpv/.venv/bin/python --case 5.0:5.0.hwp:5.0.pdf`
- 결과(최근)
  - 3.0: `renderedPages=122`, `avgRms=33.726`, `avgMae=9.5495`
  - 5.0: `renderedPages=71`, `avgRms=32.6903`, `avgMae=8.8992`
  - 3.0(1p smoke): `renderedPages=122`, `avgRms=43.3063`, `avgMae=15.9137`
  - 5.0(1p smoke): `renderedPages=71`, `avgRms=50.5679`, `avgMae=14.5375`

## 현재 상태 요약
- 페이지 수는 3.0/5.0 모두 PDF와 일치
- TOC/표지의 시각 정합은 이전 대비 개선
- `main.js` 크기 축소 진행 중: 7182 -> 6752 lines (상수/유틸/UI 셸 분리 반영)
- `main.js` 크기 축소 진행 중: 7182 -> 6441 lines (상수/유틸/UI 셸/로더/요약렌더 분리 반영)
- `main.js` 크기 축소 진행 중: 7182 -> 5753 lines (스트림 분석 코어 분리 포함)
- 완전 일치까지는 추가 튜닝 필요(폰트 메트릭, 탭 리더 길이/페이지번호 필드, 오브젝트 텍스트박스 레이아웃)

## 향후 리팩토링 계획 (재개용)
목표:
- `src/main.js`를 오케스트레이션 전용으로 축소(중간 목표: 4000 lines 이하, 최종 목표: 3000 lines 이하)
- 파서/렌더/프리뷰 경계 고정 및 회귀 자동 검증 체계화

우선순위 P1:
1. `renderPreviewPanel` 분리
   - 대상: `buildPreviewModel`, `buildPreviewPagesWithOverflow`, `buildPreviewSegments` 호출부
   - 신규 모듈: `src/render/previewPanel.js`
   - 완료 기준: `main.js`에서 Preview 패널 HTML 생성 로직 제거, 동작/페이지 수 동일
2. 캐시/선택 상태 분리
   - 대상: `getSelectedStream`, `getStreamAnalysis`
   - 신규 모듈: `src/state/selection.js`, `src/state/analysisCache.js`
   - 완료 기준: 상태 전이(스트림 선택/레코드 선택) 사이드이펙트가 함수 단위로 고립

우선순위 P2:
1. Parser 세분화
   - 분리 대상:
     - 분산문서 복호화: `src/parser/distributableDecoder.js`
     - 레코드 파싱: `src/parser/recordParser.js`
     - DocInfo 파싱: `src/parser/docInfoParser.js`
   - 완료 기준: 파서 모듈 간 의존성이 단방향(`utils -> parser -> render`)으로 정리
2. Preview 세분화
   - 분리 대상:
     - 페이지네이션: `src/preview/pagination.js`
     - 표 분할: `src/preview/tablePagination.js`
     - 오브젝트 레이아웃: `src/preview/objectLayout.js`
   - 완료 기준: 표/오브젝트 overflow 분기 로직이 UI 코드와 분리

우선순위 P3:
1. 자동 회귀 검증 강화
   - `scripts/preview_pdf_diff.mjs`에 임계치 체크 추가
   - 케이스별 기준치 파일 예: `scripts/diff_thresholds.json`
   - 실패 조건 예: `avgRms`, `avgMae`, `renderedPages` mismatch
2. 아키텍처 문서화
   - `docs/architecture.md`에 계층/의존성/책임 범위 기록
   - 신규 모듈 추가 시 파일 책임 3~5줄 요약 유지

재개 체크리스트:
1. `git pull --rebase`
2. `npm ci`
3. `npm run build`
4. `npm run diff:preview -- --url http://127.0.0.1:4222 --max-pages 2 --python /home/sjkim/work/hwpv/.venv/bin/python --case 3.0:3.0.hwp:3.0.pdf`
5. `npm run diff:preview -- --url http://127.0.0.1:4223 --max-pages 2 --python /home/sjkim/work/hwpv/.venv/bin/python --case 5.0:5.0.hwp:5.0.pdf`

작업 규칙:
- `rnd.hwp`는 로컬 전용(커밋/푸시 금지) 유지
- 리팩토링 커밋은 기능 단위로 작게 분리
- 각 리팩토링 단계마다 `build + diff smoke` 결과를 `progress.md`에 누적

## 보안/운영 메모
- `rnd.hwp`는 외부 반출 금지 정책 유지
- 저장소 `.gitignore`에 `rnd.hwp`, `artifacts/` 반영됨
