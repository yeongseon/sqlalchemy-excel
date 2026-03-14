# PRD.md — sqlalchemy-excel Product Requirements Document

> Version: 0.1.0 MVP Complete | Date: 2026-03-15 | Author: Yeongseon Choe

## 1. Vision & Positioning

### Vision Statement

> SQLAlchemy model을 단일 진실 소스(single source of truth)로 삼아, Excel 템플릿 생성 → 서버 검증(정확한 에러 리포트) → 안정적 DB 적재를 한 번에 제공하는 개발자용 툴킷.

### Problem Statement

많은 팀이 "Excel 업로드 → 검증 실패 → 원인 불명 → 재업로드 반복" 문제를 겪는다. 현재 Python 생태계에는:

- **openpyxl/pandas**: 읽기/쓰기는 강력하나, 검증 리포트/스키마 매핑/DB 적재 정책은 직접 구현 필요
- **django-import-export**: Django 전용. FastAPI/SQLAlchemy 프로젝트에서 재사용 불가
- **ExcelAlchemy**: Pydantic 중심이나 SQLAlchemy ORM 메타데이터와 1:1 통합 부족
- **pyexcel**: 포맷 통합은 강하나 스키마 계약/에러 리포트 품질이 핵심 가치가 아님

**sqlalchemy-excel**은 이 공백을 메운다.

### Package Identity

- **PyPI name**: `sqlalchemy-excel` (hyphen)
- **Import name**: `sqlalchemy_excel` (underscore)
- **License**: MIT

## 2. Target Users (Personas)

### P1: Backend Engineer (REST API 개발자)

- **상황**: FastAPI/Flask로 REST API 운영. 사용자가 Excel로 대량 등록/수정 요청
- **요구**: (a) 안전한 파일 수신, (b) 행 단위 에러 리포트, (c) 트랜잭션 적재, (d) 비동기 처리
- **성공 기준**: 업로드→검증→적재 파이프라인을 30분 내 구축

### P2: Data/Operations Staff (어드민 운영자)

- **상황**: DB 직접 접근 불가. Excel로 변경사항 정리 후 개발팀에 전달
- **요구**: 정확한 템플릿(필드명/필수 여부/허용값), 업로드 시 "어떤 셀이 왜 틀렸는지" 피드백
- **성공 기준**: 템플릿 다운로드 → 작성 → 업로드 → 명확한 에러 메시지 또는 성공 확인

### P3: Data Engineer / Analyst

- **상황**: Excel을 표준화하여 DB로 넣거나, DB 결과를 Excel로 추출
- **요구**: pandas와의 자연스러운 호환, 대용량 처리, 프로그래밍 API
- **성공 기준**: `pip install sqlalchemy-excel[pandas]` 후 5줄 코드로 import/export

### P4: Platform/DevOps Engineer

- **상황**: CLI로 배치 실행, CI/CD 파이프라인 연동
- **요구**: CLI 명령, 종료 코드, JSON/Excel 리포트 출력
- **성공 기준**: `sqlalchemy-excel validate --input data.xlsx` → exit 0 또는 exit 1 + report

## 3. MVP Feature Set (v0.1.0) — ✅ All Implemented

### F1: ORM → Excel Template Generation ✅

**설명**: SQLAlchemy ORM 모델에서 Excel 템플릿(.xlsx) 자동 생성

**요구사항**:
- ✅ ORM 모델의 컬럼명, 타입, nullable, default 값을 추출하여 헤더 생성
- ✅ 컬럼별 타입 힌트를 Excel 주석(comment)으로 표시
- ✅ Enum/선택지가 있는 컬럼은 Excel 드롭다운(DataValidation) 적용
- ✅ 샘플 데이터 행 생성 옵션
- ✅ BytesIO 출력 지원 (웹 응답용)

**수용 기준**:
- ✅ 10개 컬럼 모델에서 템플릿 생성 < 100ms
- ✅ 생성된 템플릿을 Excel에서 열면 드롭다운/주석이 정상 동작

**구현**: `template.py` — `ExcelWorkbookSession.open()` 경유하여 excel-dbapi를 통한 워크북 생성

### F2: Excel Parsing & Reading ✅

**설명**: 업로드된 Excel 파일을 파싱하여 구조화된 데이터로 변환

**요구사항**:
- ✅ 헤더 자동 감지 (첫 번째 비어있지 않은 행)
- ✅ 컬럼명 → ORM 필드 매핑 (대소문자 무시, 공백→underscore 정규화)
- ✅ 빈 행 스킵
- ✅ openpyxl 기반 읽기 모드 지원
- ✅ excel-dbapi SQL 기반 리더 (`ExcelDbapiReader`) 구현

**수용 기준**:
- ✅ 10,000행 파일 파싱 < 3초
- ✅ 파일 크기 검증 (기본 50MB 제한)

**구현**: `reader/excel_dbapi_reader.py` — `SELECT * FROM SheetName` via excel-dbapi 커서

### F3: Server-Side Validation Engine ✅

**설명**: 파싱된 데이터를 ORM 스키마/Pydantic 모델 기반으로 검증

**요구사항**:
- ✅ 타입 검증 (string→int 변환 시도 후 실패하면 에러)
- ✅ nullable 검증 (필수 필드 누락 감지)
- ✅ 길이/범위 검증 (String(50) → 50자 초과 에러)
- ✅ Enum 값 검증
- ✅ 행/열 단위 에러 수집 (ValidationReport)
- ✅ ValidationReport는 Excel 파일로 내보내기 가능 (에러 행만 하이라이트)

**수용 기준**:
- ✅ 1,000행 파일 검증 < 1초
- ✅ 에러 리포트에 행 번호, 컬럼명, 에러 메시지, 원본 값, 기대 타입 포함

**구현**: `validation/engine.py` — `ExcelDbapiReader`로 데이터 읽기, `PydanticBackend`로 동적 모델 검증

### F4: Database Import ✅

**설명**: 검증 통과 데이터를 SQLAlchemy Session으로 DB에 적재

**요구사항**:
- ✅ Insert 모드: 새 레코드 삽입
- ✅ Upsert 모드: 키 기반 업데이트 (있으면 UPDATE, 없으면 INSERT)
- ✅ Dry-run 모드: 실제 적재 없이 결과 미리보기
- ✅ 트랜잭션 안전: savepoint 기반 에러 복구
- ✅ 배치 크기 조절 (batch_size)
- ✅ ImportResult 반환 (inserted/updated/skipped/failed counts)

**수용 기준**:
- ✅ 10,000행 insert < 5초 (SQLite)
- ✅ dry-run 후 DB에 변경 없음 확인

**구현**: `load/importer.py` + `load/strategies.py` — `InsertStrategy`, `UpsertStrategy`, `DryRunStrategy` (savepoint 기반)

### F5: CLI Interface ✅

**설명**: 커맨드라인에서 템플릿/검증/임포트/엑스포트 실행

**명령어**:
```bash
sqlalchemy-excel template --model <dotpath> --output <path> [--sample-data] [--sheet-name NAME]
sqlalchemy-excel validate --model <dotpath> --input <path> [--format text|json|excel] [--output report.xlsx]
sqlalchemy-excel import --model <dotpath> --input <path> --db <url> [--mode insert|upsert] [--dry-run] [--batch-size 1000]
sqlalchemy-excel export --model <dotpath> --db <url> --output <path>
sqlalchemy-excel inspect --input <path>
```

**수용 기준**:
- ✅ `--help`로 모든 옵션 확인 가능
- ✅ 검증 실패 시 exit code 1 + stderr에 요약
- ✅ JSON 출력 옵션 (`--format json`)

**구현**: `cli.py` — Click 기반, 5개 커맨드 (template, validate, import, export, inspect)

### F6: FastAPI Integration ✅

**설명**: FastAPI 프로젝트에 즉시 통합 가능한 라우터 팩토리

**요구사항**:
- ✅ `create_import_router(model=User)` → GET /template, POST /validate, POST /import, GET /health 라우터 생성
- ✅ UploadFile → 검증 → 적재 파이프라인
- ✅ 의존성 주입(Depends)으로 Session 제공

**수용 기준**:
- ✅ 3줄 코드로 업로드 엔드포인트 추가 가능
- ✅ Swagger UI에서 파일 업로드 테스트 가능

**구현**: `integrations/fastapi.py` — `create_import_router()` 팩토리 (4 엔드포인트)

### F7: Export ✅

**설명**: SQLAlchemy 쿼리 결과를 서식이 적용된 Excel 파일로 내보내기

**요구사항**:
- ✅ Query/Select 결과 → xlsx 변환
- ✅ 컬럼 서식 (날짜, 숫자, 문자열)
- ✅ 헤더 스타일 (bold, 배경색)
- ✅ BytesIO 출력 지원

**수용 기준**:
- ✅ 10,000행 export < 3초
- ✅ 생성된 파일을 Excel에서 정상 열기 가능

**구현**: `export.py` — `ExcelWorkbookSession.open()` 경유하여 excel-dbapi를 통한 워크북 생성

### F8: excel-dbapi Integration ✅

**설명**: excel-dbapi를 핵심 Excel I/O 레이어로 사용하여 데이터 접근 통합

**요구사항**:
- ✅ `ExcelWorkbookSession`: 듀얼 채널 세션 (SQL cursor + openpyxl workbook)
- ✅ `ExcelDbapiReader`: SQL 기반 Excel 리더 (기존 openpyxl 직접 읽기 대체)
- ✅ `excel-dbapi>=1.0` 코어 의존성으로 등록
- ✅ BinaryIO → temp file 변환 (excel-dbapi는 파일 경로 필요)
- ✅ Template/Export는 `ExcelWorkbookSession.open(path, create=True)` 사용
- ✅ Validator/Importer는 `ExcelDbapiReader.read()` 사용

**수용 기준**:
- ✅ 모든 Excel I/O가 excel-dbapi를 경유하여 동작
- ✅ 기존 117개 테스트 전부 통과
- ✅ mypy --strict 통과

**구현**: `excelio/session.py` + `reader/excel_dbapi_reader.py`

## 4. Non-Goals (MVP 범위 밖)

- SQLAlchemy dialect (`create_engine("excel://...")`) — 장기 목표로 분리
- Excel-as-DB (Excel 파일을 직접 쿼리) — excel-dbapi 프로젝트에서 지원
- .xls (구 형식) 지원 — xlsx만 지원
- GUI/웹 UI — 레퍼런스 앱 수준만 제공
- 실시간 협업/동시 편집 — 범위 밖
- Excel Online/SharePoint 연동 — 범위 밖
- BackgroundTasks 비동기 처리 — 향후 개선 사항

## 5. Technical Constraints

1. **Python 3.10+**: `match` 문, `|` union 타입 등 현대 문법 사용
2. **SQLAlchemy 2.0+**: `Mapped[]`, `mapped_column()`, `DeclarativeBase` 전용
3. **Core 의존성**: `sqlalchemy`, `openpyxl`, `pydantic`, `defusedxml`, `click`, `excel-dbapi`
4. **Optional extras**: `pandas`, `pandera`, `fastapi`, `python-multipart`은 선택 설치
5. **Security**: defusedxml 필수 (XML bomb 방어), 파일 크기 제한, formula injection 방지
6. **excel-dbapi 제약**: file path 필수 (BinaryIO는 temp file로 변환), unquoted table names 사용

## 6. Quality Requirements

### Performance
- 10,000행 파일: 파싱+검증+적재 < 10초
- 100,000행 파일: 스트리밍 모드에서 메모리 < 200MB

### Reliability
- 적재 실패 시 savepoint 롤백 보장
- 손상된 Excel 파일에 대한 명확한 에러 메시지 (crash 방지)
- Batch 실패 시 per-row retry (UpsertStrategy)

### Security
- defusedxml으로 XML 공격 방어
- Formula injection 방지 (`sanitize_cell_value()` — 셀 값이 `=`, `+`, `-`, `@`, `\t`, `\r`로 시작하면 `'` prefix)
- 파일 크기 제한 기본값 제공 (50MB)

### Compatibility
- CI 매트릭스: Python 3.10–3.13
- SQLAlchemy 2.0.x 호환 검증
- openpyxl 3.1.x 호환 검증
- excel-dbapi 1.0.x 호환 검증

## 7. Success Metrics (v0.1.0)

| Metric | Target | Status |
|--------|--------|--------|
| PyPI 첫 릴리스 | 2026-05 이전 | 준비 완료 |
| Quickstart 완료 시간 | < 10분 | ✅ |
| CI 매트릭스 통과 | Python 3.10–3.13 green | ✅ |
| 테스트 커버리지 | > 80% | ✅ (117 tests) |
| 공개 API 타입 힌트 | 100% | ✅ (mypy --strict) |
| 문서화된 public 함수 | 100% | ✅ |
| excel-dbapi 통합 | 전면 의존성 | ✅ |

## 8. Future Roadmap (Post-MVP)

| Phase | Feature | Priority |
|-------|---------|----------|
| 0.2.0 | 대용량 스트리밍 최적화, Pandera 옵션 백엔드 | High |
| 0.3.0 | Alembic 연동 가이드, 스키마/매핑 파일 포맷 | Medium |
| 0.4.0 | BackgroundTasks 비동기 처리, 진행률 콜백 | Medium |
| 0.5.0 | 실험적 Excel dialect (read-only, excel-dbapi 연동) | Low |
| 1.0.0 | API 안정 선언, 기업 도입 사례 | — |
