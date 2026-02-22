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

## 커밋 이력(이번 라운드)
- `af74f7f` Stabilize diff capture and add tab leader rendering
- `115a78b` Refine TOC tabbed line rendering for leader alignment
- `7ceb6e4` Relax cover textbox wrapping for large slash titles

## 검증 결과
- 실행 명령
  - `npm run build`
  - `npm run diff:preview -- --url http://127.0.0.1:4208 --max-pages 20 --python /home/sjkim/work/hwpv/.venv/bin/python`
- 결과(최근)
  - 3.0: `renderedPages=122`, `avgRms=33.726`, `avgMae=9.5495`
  - 5.0: `renderedPages=71`, `avgRms=32.6903`, `avgMae=8.8992`

## 현재 상태 요약
- 페이지 수는 3.0/5.0 모두 PDF와 일치
- TOC/표지의 시각 정합은 이전 대비 개선
- 완전 일치까지는 추가 튜닝 필요(폰트 메트릭, 탭 리더 길이/페이지번호 필드, 오브젝트 텍스트박스 레이아웃)

## 보안/운영 메모
- `rnd.hwp`는 외부 반출 금지 정책 유지
- 저장소 `.gitignore`에 `rnd.hwp`, `artifacts/` 반영됨
