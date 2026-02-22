# Progress

작성일: 2026-02-22
브랜치: `main`

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
- `main.js` 크기 축소 진행 중: 7182 -> 5995 lines (Stream/Record 패널 렌더 분리 포함)
- 완전 일치까지는 추가 튜닝 필요(폰트 메트릭, 탭 리더 길이/페이지번호 필드, 오브젝트 텍스트박스 레이아웃)

## 다음 리팩토링 단계(계획)
1. `src/parser` 분리
   - CFB 로드/엔트리 스캔/스트림 디코딩을 `cfbReader`, `recordStream` 모듈로 이동
2. `src/preview` 분리
   - 페이지 분할/문단 레이아웃/표 렌더를 `pagination`, `paragraph`, `table` 모듈로 분리
3. `src/render` 분리
   - Summary/DocInfo/Record/Detail 패널 렌더러를 각각 모듈화
4. 회귀 방지
   - `diff:preview` 스크립트에 케이스별 기준치 체크(예: avgRms/avgMae threshold) 추가

## 보안/운영 메모
- `rnd.hwp`는 외부 반출 금지 정책 유지
- 저장소 `.gitignore`에 `rnd.hwp`, `artifacts/` 반영됨
