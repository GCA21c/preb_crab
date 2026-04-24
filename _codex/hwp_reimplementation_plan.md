# HWP Reimplementation Plan

## 목적

HWP/HWPX를 외부 한글 설치나 외부 bridge 없이, 우리 프로젝트 내부 코드로 읽기 전용 페이지 렌더 소스로 재구현한다.

최종 목표:

`HWP/HWPX -> 내부 문서 모델 -> 페이지 렌더 -> ORIGIN 표시 -> 기존 capture 재사용`

## 구현 철학

- 문자열 편집 기능은 우선 목표가 아님
- 핵심은 `사람이 보는 페이지를 그대로 그리는 것`
- 텍스트 추출기가 아니라 `페이지 렌더 엔진`을 만든다
- 외부 실행파일/sidecar를 최종 해법으로 채택하지 않는다
- 기술문서에 없는 휴리스틱 page split/merge는 기본 경로에서 빼고, record/page break 기반 구조를 우선한다

## 참고 전략

Codex가 매 단계 비교 분석 후 참고 대상을 고른다.

### 공식 기준 문서

- `_codex/resources/hwp file format_5.0_revision1.3.pdf`
  - 본문 레코드: p.38~41
  - 표 개체: p.44
  - 구역 정의/용지 설정: p.58~59
- Hancom Tech Python parsing series
  - `https://tech.hancom.com/python-hwp-parsing-1/`
  - OLE 접근, zlib 해제, record header 분해 예제
  - `https://tech.hancom.com/python-hwpx-parsing-1/`
  - HWPX zip/xml 구조, section xml 접근 예제

## 고정 우선순위

다음 순서는 절대 바꾸지 않는다.

1. `기술문서 PDF 확인`
2. `Hancom Tech Python 글 확인`
3. `openhwp / hwp-rs` 구조 확인`
4. `reader/model 수정`
5. 그 다음에만 `rendered vs reference` 비교

즉:

- 번호만 같은 페이지끼리 바로 diff 금지
- 휴리스틱 후처리부터 만지는 것 금지
- 기술문서/공식 글 확인 없이 page split/merge부터 만지는 것 금지
- rendered가 구조를 갖추기 전에는 비교를 품질 판단 근거로 쓰지 않는다

## 페이지 매칭 원칙

현재 compare 파이프라인은 `reference/page_###.png`와 `rendered/page_###.png`를 기계적으로 같은 번호로 비교한다.
이건 **rendered가 구조를 갖춘 뒤**에만 유효하다.

앞으로는 반드시 아래 순서를 지킨다.

1. rendered가 기술문서/한컴 글 기준의 페이지 구조를 먼저 갖춘다
2. 각 페이지의 `첫 제목/첫 문단/첫 표 헤더`를 추출한다
3. `reference`와 `rendered`의 의미상 같은 페이지를 먼저 매칭한다
4. 그 다음에만 review/diff 생성한다

금지:

- `11페이지 vs 11페이지`만 맞았다고 같은 페이지라고 간주
- page number만 맞춘 상태에서 diff 수치로 품질 판단
- page alignment가 불명확한 상태에서 이미지 오버랩 결과를 근거로 파서/렌더 수정

### 1. `hwp-rs`

- 용도: `.hwp` 저수준 파싱
- 참고 포인트: OLE stream, record, binary structure
- 현재 직접 반영 중:
  - `Body -> Section -> Paragraph`
  - `ParagraphHeader` break flags
  - `TableControl.row_count`
  - `PageDefinition`

### 2. `rhwp`

- 용도: 렌더/페이지네이션/조판 참고
- 참고 포인트: page-like render, table/object layout

### 3. `openhwp`

- 용도: 전체 아키텍처 참고
- 참고 포인트: parser/viewer/editor 계층 분리

### 4. `unhwp`

- 용도: 문서 모델(IR) 참고
- 참고 포인트: HWP/HWPX 공통 구조, 표/자산 표현

## 현재 구현 원칙

- rendered가 구조를 갖추기 전에는 reference와의 page diff를 품질 근거로 쓰지 않는다
- page_count_hint, page rebalance, attachment split 같은 휴리스틱 후처리는 기본 구조가 잡히기 전까지 사용하지 않는다
- 비교는 구조가 기술문서/한컴 글 기준으로 고정된 뒤에만 사용한다

## 현재 렌더 원칙

- 렌더러는 우선 generic paragraph/table flow만 보여준다
- title box / notice box / section badge 같은 시각 추정은 기본 렌더에서 빼둔다
- 실제 구조가 파싱되기 전에는 텍스트 내용으로 화면 스타일을 가공하지 않는다

## 현재 구현 상태

### 완료

- `core/hwp_types.py`
  - 내부 문서/페이지/문단/표 모델 기초
  - paragraph break flag / table control id / raw record bytes / raw attrs 보존 필드 추가
  - PARA_HEADER text/control/para shape fields와 TABLE row/col/border fields를 실제 필드로 승격
  - PARA_CHAR_SHAPE와 PARA_LINE_SEG를 문단 모델의 실제 필드로 승격
  - HWPX beginNum / refList / section attrs 보존용 source metadata 추가
- `core/hwp_probe.py`
  - `.hwp` OLE / `.hwpx` ZIP 구조 분석
  - Hancom Tech 글의 DocInfo 순서 기준으로 DocInfo 압축 해제 / record header 순회 추가
  - HWP DOCUMENT_PROPERTIES / ID_MAPPINGS / FACE_NAME 파싱
  - HWP BORDER_FILL / CHAR_SHAPE / PARA_SHAPE 상세 레코드 파싱
  - HWP 글꼴 목록과 ID mapping counts를 source metadata에 연결
  - HWPX beginNum / refList / head attrs metadata 추출
- `core/hwp_reader.py`
  - 최소 내부 reader 연결
  - 기술문서 기준 record/page break 구조 우선
  - HWP PARA_HEADER / TABLE header 기본 필드 해석
  - HWP PARA_CHAR_SHAPE / PARA_LINE_SEG 해석 후 문단에 연결
  - HWPX section attrs / page defs raw 보존
  - HWPX page size / margins 가능한 범위 해석
  - HWP / HWPX page source name 보존
  - page source index/name 보존
- `core/hwp_renderer.py`
  - 최소 page-like renderer 연결
  - generic paragraph/table flow 렌더
  - 문단 line segment의 줄 높이/텍스트 시작 위치를 기본 렌더에 반영
  - DocInfo FACE_NAME에서 추출한 글꼴을 시스템에 있으면 우선 사용
  - CHAR_SHAPE 기반 글자 크기/굵기/기울임/색상 최소 반영
  - PARA_SHAPE 기반 문단 좌우 여백/앞뒤 간격 최소 반영
  - BORDER_FILL 단색 채우기 필드 파싱 및 표 배경 최소 반영 경로 추가
  - TABLE cell 속성의 row/col/span/size/margin/border_fill_id를 셀 모델에 승격
  - 셀별 BORDER_FILL을 우선 적용하고 없으면 표 BORDER_FILL을 fallback으로 사용
  - BORDER_FILL의 4방향 선 종류/굵기/색상을 개별 렌더
  - PARA_SHAPE 정렬 비트를 일반 문단과 표 셀 텍스트에 반영
- `core/document_loader.py`
  - HWP/HWPX 내부 경로를 ORIGIN에 연결
- `_codex/render_reference_pdf.py`
  - 기준 PDF를 페이지 PNG로 렌더
- `_codex/render_hwp_sample.py`
  - 내부 HWP 렌더 결과를 페이지 PNG로 저장
- `_codex/compare_hwp_pdf.py`
  - 기준 PDF와 내부 HWP 렌더의 페이지별 diff/report 생성

### 현재 품질

- 문서는 빠르게 열린다
- 내용은 보인다
- sample HWP 기준 `BodyText/Section0`에서 3페이지가 생성된다
- sample HWP 기준 문단 55개 전체에 `PARA_CHAR_SHAPE`, `PARA_LINE_SEG`가 연결된다
- DocInfo에서 글꼴 14개, char_shape 160개, para_shape 112개, border_fill 69개가 추출된다
- 페이지 구조는 아직 기술문서/한컴 글 기준으로 더 맞춰야 한다
- 페이지 수를 억지로 맞추는 후처리는 기본 경로에서 배제한다
- 하지만 품질은 아직 낮다

현재 확인된 문제:

1. sample HWP는 기준 PDF 11페이지와 아직 페이지 수가 맞지 않는다
2. 기본 경로에서 `page_count_hint` 재분배와 attachment split은 제거되어 있다
3. 첫 페이지의 큰 구조는 보이나, 색상/테두리/굵기/정렬/문단 스타일은 아직 원본과 차이가 크다
4. 표는 셀 단위로 보이지만 열 너비/행 높이/정렬/강조가 아직 부정확하다
5. DocInfo의 char_shape / para_shape / border_fill 상세 레코드는 읽고 기본 렌더에 반영하지만, 그라데이션/이미지 채우기/복합 선종류는 아직 제한적이다
6. diff 수치만으로는 품질 판단이 왜곡될 수 있고, 사람 눈 비교가 계속 필요하다
7. 아직 page-like preview 수준이며 실사용 품질은 아니다

## 지금까지의 판단

- 속도는 충분히 가능성이 있다
- 병목은 파싱보다 `구조 해석`과 `렌더 품질`이다
- `PrvText`는 잘려 있어서 단독 기준이 될 수 없고, `BodyText`를 본문 소스로 유지해야 한다
- page_count_hint, heuristic rebalance, attachment split 같은 후처리는 기본 구조 파악이 끝난 뒤에만 검토한다
- 비교 자동화는 이미 구축됐지만, **페이지 구조가 고정된 뒤에만** `_codex/compare/` 산출물을 사용한다

## 현재 교훈

이번 세션에서 확인된 잘못된 접근:

1. `page-to-page semantic alignment`가 안 맞는데 번호 기준으로 diff를 본 것
2. 그 상태에서 diff 수치를 보고 렌더/후처리를 오래 수정한 것
3. rendered가 구조를 갖추기 전에 비교를 품질 판단 근거처럼 쓴 것

이건 다시 반복하면 안 된다.

앞으로 compare 산출물은 아래 상태에서만 신뢰한다.

- rendered가 기술문서/한컴 글 기준 구조를 먼저 갖춤
- `reference 첫 블럭`과 `rendered 첫 블럭`이 의미상 대응됨
- attachment/NCS/form 페이지가 실제로 같은 의미 페이지로 매칭됨
- 그 다음에만 overlay/diff 해석 가능

## 다음 우선순위

### Phase A

1. 기술문서 PDF와 Hancom Python 글에 있는 record/page break 구조를 먼저 정확히 따른다
2. `.hwp` `BodyText` record에서 `PARA_HEADER / CTRL_HEADER / TABLE / PAGE_DEF / PARA_CHAR_SHAPE / PARA_LINE_SEG` 해석을 더 늘린다
3. `PrvText`는 앞쪽 구조 힌트로만 참고하고, 본문 구조를 덮어쓰는 기준으로 쓰지 않는다
4. page_count_hint 기반 재분배와 휴리스틱 merge/split을 기본 경로에서 제거한다
5. DocInfo의 CHAR_SHAPE / PARA_SHAPE / BORDER_FILL 상세 레코드를 실제 모델 필드로 승격한다
6. 표 구조를 셀 단위로 더 안정화하고, 2열/5열 같은 흔한 패턴의 열 너비 규칙을 보강한다

### Phase B

7. `rhwp / hwp-rs / openhwp / unhwp`는 기술문서/한컴 글로 설명되지 않는 구조나 구현 세부를 보강할 때만 후순위로 본다
8. `.hwpx`와 `.hwp`의 공통 문서 모델을 더 정교하게 맞춘다
9. 제목 / 공지 박스 / 섹션 배너 / 표 헤더 같은 시각 패턴을 문서 모델 기반으로 일반화한다

### Phase C

10. 고해상도 렌더 강화
11. 벡터 기반 렌더 또는 고배율 내부 캔버스 검토
12. 확대 시 글자/선 깨짐 완화

### Phase D

13. header/footer
14. textbox/shape
15. equation/image
16. 실제 조판 품질 보정

## 작업 원칙

- HWP 재구현 중에도 ORIGIN/HERE/CAPTURE BLOCKS 기존 안정성을 깨면 안 된다
- bridge를 되살려 임시로 덮지 않는다
- 테스트 결과가 나쁘면 계획과 참고 우선순위를 바로 조정한다

## 현재 테스트 포인트

다음 확인이 필요:

1. `.hwpx`가 1페이지로만 뭉치는지
2. `.hwp`가 기준 PDF와 페이지 수를 계속 맞추는지
3. 첫 페이지에서 title / notice / section / table 구조가 더 원본에 가까워지는지
4. 표가 표처럼 보이기 시작하는지
5. 확대 시 가독성이 얼마나 유지되는지
