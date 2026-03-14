# ARCH.md — sqlalchemy-excel Architecture Document

> **Version**: 0.1.0  
> **Last Updated**: 2026-03-15  
> **Status**: MVP Complete (feature-complete)

---

## 1. System Overview

sqlalchemy-excel는 SQLAlchemy ORM 모델을 **단일 진실 소스(single source of truth)**로 사용하여 Excel 워크플로(템플릿 생성, 업로드 검증, 데이터베이스 임포트, 쿼리 내보내기)를 자동화하는 툴킷입니다.

모든 Excel I/O는 **excel-dbapi**를 통해 수행됩니다. excel-dbapi는 PEP 249 DB-API 2.0 호환 드라이버로, SQL 기반 데이터 접근과 openpyxl 워크북 직접 접근을 모두 제공합니다.

```
┌──────────────────────────────────────────────────────────────────────┐
│                          sqlalchemy-excel                             │
│                                                                       │
│  ┌──────────┐   ┌──────────┐   ┌──────────┐   ┌───────────┐          │
│  │ mapping  │──▶│ template │   │  reader  │──▶│validation │          │
│  │          │   │          │   │          │   │           │          │
│  │ ORM →    │   │ Mapping→ │   │ .xlsx →  │   │ Rows →    │          │
│  │ Schema   │   │ .xlsx    │   │ Rows     │   │ Report    │          │
│  └──────────┘   └──────────┘   └──────────┘   └─────┬─────┘          │
│       │                                              │               │
│       │               ┌──────────┐                   │               │
│       └──────────────▶│   load   │◀──────────────────┘               │
│                       │          │                                    │
│                       │ Rows→DB  │                                    │
│                       └──────────┘                                    │
│                                                                       │
│  ┌──────────┐   ┌──────────────┐   ┌────────┐                        │
│  │  export  │   │ integrations │   │  cli   │                        │
│  │          │   │              │   │        │                        │
│  │ Query→   │   │ FastAPI      │   │ Click  │                        │
│  │ .xlsx    │   │ router       │   │ cmds   │                        │
│  └──────────┘   └──────────────┘   └────────┘                        │
│                                                                       │
│  ┌─────────────────────────────────────────┐                          │
│  │          excelio / reader               │                          │
│  │   excel-dbapi Integration Layer         │                          │
│  │                                         │                          │
│  │  ExcelWorkbookSession  ExcelDbapiReader │                          │
│  │  (쓰기: 듀얼 채널)     (읽기: SQL 기반) │                          │
│  └────────────────────┬────────────────────┘                          │
│                       │                                               │
└───────────────────────┼───────────────────────────────────────────────┘
                        │ excel_dbapi.connect()
                        ▼
┌──────────────────────────────────────────────────────────────────────┐
│                          excel-dbapi                                  │
│                                                                       │
│  ┌──────────────────────────────────────────────────────────────┐     │
│  │  PEP 249 DB-API 2.0 Interface                                │     │
│  │  ExcelConnection · ExcelCursor · PEP 249 Exceptions          │     │
│  ├──────────────────────────────────────────────────────────────┤     │
│  │  SQL Parser  →  Engine Abstraction  →  Executor              │     │
│  ├────────────────────┬─────────────────────────────────────────┤     │
│  │  OpenpyxlEngine    │  PandasEngine                           │     │
│  │  (workbook 속성)   │  (DataFrame 기반)                       │     │
│  ├────────────────────┴─────────────────────────────────────────┤     │
│  │  Storage: .xlsx File (Excel 2007+ Open XML)                   │     │
│  └──────────────────────────────────────────────────────────────┘     │
└──────────────────────────────────────────────────────────────────────┘
```

---

## 2. Module Architecture

### 2.1 Layer Diagram

```
┌──────────────────────────────────────────────────────┐
│                Public API Layer                       │
│  __init__.py: ExcelMapping, ExcelTemplate,            │
│  ExcelValidator, ExcelImporter, ExcelExporter,        │
│  ExcelWorkbookSession, ImportResult                   │
├──────────────────────────────────────────────────────┤
│             Interface Layer (CLI / Web)               │
│  cli.py (Click)  │  integrations/fastapi.py          │
├──────────────────────────────────────────────────────┤
│               Service Layer                           │
│  mapping.py  │  template.py  │  export.py            │
│  validation/engine.py  │  load/importer.py           │
├──────────────────────────────────────────────────────┤
│          excel-dbapi Integration Layer                │
│  excelio/session.py (ExcelWorkbookSession)           │
│  reader/excel_dbapi_reader.py (ExcelDbapiReader)     │
├──────────────────────────────────────────────────────┤
│               Backend Layer                           │
│  reader/openpyxl_reader.py (레거시, 호환용)          │
│  validation/pydantic_backend.py                      │
│  load/strategies.py                                  │
├──────────────────────────────────────────────────────┤
│            Infrastructure Layer                       │
│  _types.py  │  _compat.py  │  exceptions.py          │
├──────────────────────────────────────────────────────┤
│           External Dependencies                       │
│  excel-dbapi ≥ 1.0  │  SQLAlchemy ≥ 2.0             │
│  openpyxl ≥ 3.1  │  Pydantic ≥ 2.0                  │
│  defusedxml ≥ 0.7  │  Click ≥ 8.0                   │
└──────────────────────────────────────────────────────┘
```

### 2.2 Dependency Rules

1. **Public API** → Service Layer only (직접 Backend 접근 금지)
2. **Service Layer** → excel-dbapi Integration Layer + Backend Layer
3. **excel-dbapi Integration Layer** → excel-dbapi 외부 패키지
4. **Backend Layer** → External libraries (openpyxl, pydantic)
5. **Interface Layer** → Service Layer only
6. **Infrastructure** ← 모든 레이어에서 사용

### 2.3 Dependencies

```
sqlalchemy-excel (core dependencies)
├── sqlalchemy >= 2.0         # ORM 스키마 소스, DB 작업
├── excel-dbapi >= 1.0        # 전면 의존성: 모든 Excel I/O
├── openpyxl >= 3.1           # 셀 레벨 포맷팅 (workbook 채널 경유)
├── pydantic >= 2.0           # 행 레벨 유효성 검증
├── defusedxml >= 0.7         # 보안: XML 공격 방어
└── click >= 8.0              # CLI 인터페이스

sqlalchemy-excel[pandas]      (+ pandas >= 2.0)
sqlalchemy-excel[pandera]     (+ pandera >= 0.18, pandas >= 2.0)
sqlalchemy-excel[fastapi]     (+ fastapi >= 0.100, python-multipart >= 0.0.5)
sqlalchemy-excel[dev]         (+ pytest, hypothesis, ruff, mypy, coverage)
sqlalchemy-excel[all]         (위 모든 extras)
```

---

## 3. Module Specifications

### 3.1 `excelio/session.py` — ExcelWorkbookSession (Dual-Channel Wrapper)

**역할**: excel-dbapi 커넥션을 래핑하여 **데이터 채널**(SQL 커서)과 **포맷 채널**(openpyxl 워크북)을 동시에 제공합니다. 템플릿 생성과 내보내기에서 워크북 생성 및 저장에 사용됩니다.

**아키텍처 — 듀얼 채널**:

```
ExcelWorkbookSession
│
├── Data Channel (SQL)
│   └── self._cursor = self._conn.cursor()
│       ├── cursor.execute("SELECT * FROM Sheet1")
│       └── cursor.fetchall() → List[Tuple]
│
└── Format Channel (openpyxl)
    └── self._conn.workbook → openpyxl.Workbook
        ├── workbook[sheet_name] → Worksheet
        ├── cell.font, cell.fill, cell.alignment
        └── DataValidation, comments, freeze_panes
```

**Class API**:

```python
class ExcelWorkbookSession:
    def __init__(self, conn: Any) -> None:
        self._conn = conn
        self._cursor = conn.cursor()

    @classmethod
    def open(
        cls,
        path: str | Path,
        *,
        create: bool = False,
        data_only: bool = False,
    ) -> ExcelWorkbookSession:
        """excel-dbapi connect()를 호출하여 세션을 생성합니다.

        내부 구현:
            conn = excel_dbapi.connect(
                str(path),
                engine="openpyxl",
                autocommit=False,
                create=create,
                data_only=data_only,
            )
        """

    @property
    def conn(self) -> Any: ...      # excel-dbapi 커넥션 객체
    @property
    def cursor(self) -> Any: ...    # DB-API 커서
    @property
    def workbook(self) -> Any: ...  # openpyxl Workbook (conn.workbook 경유)

    def commit(self) -> None: ...   # 워크북 변경 디스크 저장
    def rollback(self) -> None: ... # 워크북 변경 롤백
    def close(self) -> None: ...    # 커넥션 종료

    def __enter__(self) -> ExcelWorkbookSession: ...
    def __exit__(self, exc_type, exc_val, exc_tb) -> None: ...
```

**사용 패턴 (template.py)**:

```python
with ExcelWorkbookSession.open(path, create=True) as session:
    workbook = session.workbook        # openpyxl Workbook 직접 접근
    ws = workbook.create_sheet("Users")
    ws.cell(row=1, column=1, value="name")
    ws.cell(row=1, column=1).font = Font(bold=True)
    session.commit()                    # 디스크에 저장
```

**설계 결정**:
- `engine="openpyxl"` 고정 — 포맷 채널(workbook 속성)은 openpyxl 엔진만 지원
- `autocommit=False` — 명시적 `commit()` 호출이 필요하며 트랜잭션 안전성 보장
- `create=True` — 새 파일 생성 시 사용 (빈 워크북 초기화)
- `data_only=True` — 수식이 아닌 캐시된 값 읽기 (읽기 전용 시)

### 3.2 `reader/excel_dbapi_reader.py` — ExcelDbapiReader (SQL-Based Reader)

**역할**: excel-dbapi의 SQL 인터페이스를 사용하여 Excel 파일을 읽습니다. `SELECT * FROM SheetName` 쿼리로 데이터를 가져오고, 정규화된 헤더와 행 딕셔너리로 반환합니다. `ExcelValidator`와 `ExcelImporter`에서 사용됩니다.

**Class API**:

```python
class ExcelDbapiReader:
    def __init__(
        self,
        *,
        read_only: bool = False,
        max_file_size: int = 50 * 1024 * 1024,  # 50MB
    ) -> None: ...

    def read(
        self,
        source: FileSource,          # str | Path | BinaryIO
        sheet_name: str | None = None,
        header_row: int | None = None,
    ) -> ReaderResult: ...

    def _resolve_source(self, source: FileSource) -> tuple[str, bool]: ...
    def _validate_file_size_path(self, path: str) -> None: ...
    @staticmethod
    def _normalize_headers(raw_headers: list[str]) -> list[str]: ...
    @staticmethod
    def _is_empty_cell(value: object) -> bool: ...
    @staticmethod
    def _remove_temp_file(path: str) -> None: ...
```

**읽기 파이프라인 (read 메서드 내부)**:

```
FileSource (str | Path | BinaryIO)
    │
    ▼
_resolve_source()
    ├── str / Path → (file_path, False)    # 파일 경로 그대로 사용
    └── BinaryIO → tempfile.NamedTemporaryFile(.xlsx) → (temp_path, True)
    │
    ▼
_validate_file_size_path()
    ├── os.path.getsize(path)
    └── size > max_file_size → ReaderError 발생
    │
    ▼
excel_dbapi.connect(
    str(file_path),
    engine="openpyxl",
    autocommit=True,
    data_only=True,
)
    │
    ▼
conn.workbook.sheetnames → 시트 이름 확인
    ├── sheet_name 지정 → 존재 확인 (SheetNotFoundError)
    └── sheet_name 미지정 → 첫 번째 시트 사용
    │
    ▼
cursor.execute(f"SELECT * FROM {resolved_sheet}")
    │ 주의: 테이블 이름은 따옴표 없이 사용 (excel-dbapi 규칙)
    │
    ▼
cursor.description → 헤더 추출
    │
    ▼
_normalize_headers() → normalize_header() 적용
    ├── strip, lowercase, spaces → underscores
    ├── 빈 헤더 검사 → ReaderError
    └── 중복 헤더 검사 → ReaderError
    │
    ▼
cursor.fetchall() → 행 반복
    ├── 빈 행 필터링 (_is_empty_cell)
    └── dict(zip(headers, values)) → RowDict
    │
    ▼
ReaderResult(headers, rows, total_rows)
```

**BinaryIO → 임시 파일 변환 (_resolve_source)**:

excel-dbapi는 파일 경로를 요구하므로, `BinaryIO` 입력은 임시 파일로 변환됩니다:

```python
def _resolve_source(self, source: FileSource) -> tuple[str, bool]:
    if isinstance(source, (str, Path, os.PathLike)):
        return str(source), False  # 경로 그대로, 삭제 불필요

    # BinaryIO: 내용을 읽어서 임시 파일에 저장
    content = binary_source.read()
    binary_source.seek(original_position)  # 커서 복원

    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
        tmp.write(content)
        temp_path = tmp.name

    return temp_path, True  # 임시 파일 경로, 읽기 후 삭제 필요
```

**에러 처리**:

```python
try:
    # 정상 읽기 로직
except (InvalidFileException, BadZipFile):
    raise FileFormatError("Input is not a valid .xlsx file")
except (SheetNotFoundError, ReaderError, FileFormatError):
    raise  # 재발생 (래핑 안 함)
except Exception as exc:
    raise ReaderError(f"Failed to read Excel data: {exc}") from exc
finally:
    if conn is not None:
        conn.close()  # 항상 커넥션 정리
    if remove_after_read:
        _remove_temp_file(file_path)  # 임시 파일 정리
```

**설계 결정**:
- `autocommit=True` — 읽기 전용이므로 트랜잭션 관리 불필요
- `data_only=True` — 수식이 아닌 캐시된 값 읽기
- 테이블 이름은 **따옴표 없이** 사용 — `SELECT * FROM Sheet1` (따옴표 사용 시 lookup 실패)
- `suppress(OSError)` 패턴으로 임시 파일 정리 실패를 안전하게 무시

### 3.3 `mapping.py` — ORM Schema Extraction

**역할**: SQLAlchemy ORM 모델을 인트로스펙션하여 `ExcelMapping` 데이터클래스를 생성합니다. 모든 후속 작업(템플릿, 검증, 임포트, 내보내기)의 기반이 됩니다.

**핵심 타입**:

```python
@dataclass(frozen=True)
class ColumnMapping:
    name: str                         # ORM 컬럼 이름
    excel_header: str                 # Excel 헤더 표시 텍스트
    python_type: type[object]         # Python 타입 (str, int, float 등)
    sqla_type: TypeEngine[object]     # SQLAlchemy 타입 객체
    nullable: bool                    # NULL 허용 여부
    primary_key: bool                 # 기본 키 여부
    has_default: bool                 # 기본값 존재 여부
    default_value: object | None      # 정적 기본값 (callable이면 None)
    enum_values: list[str] | None     # Enum 드롭다운 옵션
    max_length: int | None            # String(N) 최대 길이
    description: str | None           # 컬럼 doc/comment
    foreign_key: str | None           # 참조하는 table.column

@dataclass(frozen=True)
class ExcelMapping:
    model_class: type[DeclarativeBase]  # ORM 모델 클래스
    sheet_name: str                      # 워크시트 이름
    columns: list[ColumnMapping]         # 정렬된 컬럼 매핑
    key_columns: list[str]               # upsert 키 컬럼
```

**인트로스펙션 전략**:

```
1. sa_inspect(model) → mapper 획득
2. mapper.columns 순회 → 컬럼 메타데이터 추출
3. _TYPE_MAP 기반 타입 매핑:
   Integer → int, String → str, Text → str, Float → float,
   Boolean → bool, Date → date, DateTime → datetime,
   Numeric(asdecimal=True) → Decimal, Numeric(asdecimal=False) → float,
   Enum → str (enum_values 추출)
4. include/exclude 필터 적용
5. header_map 오버라이드 적용
6. key_columns 미지정 시 primary key 자동 사용
```

**설계 결정**:
- Frozen dataclass → 매핑 불변성 보장 (생성 후 수정 금지)
- `from_model()` 클래스 메서드가 유일한 팩토리
- `include`와 `exclude`는 상호 배타적 (MappingError 발생)

### 3.4 `template.py` — Excel Template Generation

**역할**: `ExcelMapping`을 서식이 적용된 `.xlsx` 템플릿 파일로 변환합니다. **ExcelWorkbookSession**을 사용하여 워크북을 생성하고 저장합니다.

**excel-dbapi 통합 방식**:

```python
class ExcelTemplate:
    def save(self, path: str | Path) -> None:
        with ExcelWorkbookSession.open(path, create=True) as session:
            self._populate_workbook(session.workbook)  # openpyxl Workbook 직접 조작
            session.commit()                            # excel-dbapi를 통해 디스크 저장

    def to_bytesio(self) -> BytesIO:
        # 1. 임시 파일 생성
        # 2. self.save(temp_path) → ExcelWorkbookSession 사용
        # 3. 임시 파일 읽어서 BytesIO 반환
        # 4. 임시 파일 정리
```

**스타일 상수**:

```python
HEADER_FILL = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
HEADER_FONT = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
HEADER_ALIGNMENT = Alignment(horizontal="center", vertical="center")
REQUIRED_FILL = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
```

**템플릿 생성 파이프라인**:

```
ExcelWorkbookSession.open(path, create=True)
    │
    ▼
session.workbook → openpyxl.Workbook
    │
    ▼
기존 시트 모두 삭제
    │
    ▼
각 ExcelMapping에 대해:
    ├── workbook.create_sheet(title=sheet_name)
    ├── 헤더 행 작성 (스타일: 폰트, 배경색, 테두리)
    │   ├── 필수 컬럼 → 노란색 배경 (FFF2CC)
    │   └── 선택 컬럼 → 파란색 배경 (4472C4)
    ├── 셀 코멘트 추가 (타입, nullable, PK, FK, 설명)
    ├── 컬럼 너비 자동 조정
    ├── Enum 컬럼 → DataValidation 드롭다운
    │   └── 255자 제한 초과 시 코멘트로 대체
    ├── Auto-filter 설정
    ├── Freeze panes (A2)
    └── 샘플 데이터 행 (선택)
    │
    ▼
session.commit() → 디스크 저장
```

**설계 결정**:
- `ExcelWorkbookSession.open(create=True)` — 새 파일 생성을 excel-dbapi에 위임
- `session.workbook` — openpyxl 워크북에 직접 접근하여 셀 레벨 포맷팅 수행
- `session.commit()` — excel-dbapi의 원자적 저장(tempfile + os.replace) 활용
- DataValidation은 클라이언트 힌트일 뿐 — 서버 검증이 항상 필요

### 3.5 `reader/` — Excel File Parsing

**아키텍처**: Protocol 기반 플러그인 구조. 주 리더는 `ExcelDbapiReader`, 레거시 `OpenpyxlReader` 유지.

```python
# reader/base.py
@dataclass(slots=True)
class ReaderResult:
    headers: list[str]                # 정규화된 헤더
    rows: Iterable[RowDict]          # 행 딕셔너리 반복자
    total_rows: int | None           # 행 수 (스트리밍 시 None)

class BaseReader(Protocol):
    def read(
        self,
        source: FileSource,
        sheet_name: str | None = None,
        header_row: int | None = None,
    ) -> ReaderResult: ...

def normalize_header(header: str) -> str:
    """strip → lowercase → spaces→underscores → 비알파벳 제거"""
```

**리더 선택 전략**:

| 리더 | 사용 위치 | 접근 방식 |
|------|----------|----------|
| `ExcelDbapiReader` | ExcelValidator, ExcelImporter | `SELECT * FROM SheetName` via excel-dbapi |
| `OpenpyxlReader` | 레거시 호환용 (직접 사용 없음) | openpyxl `iter_rows()` 직접 호출 |

### 3.6 `validation/` — Validation Engine

**역할**: 파싱된 행을 스키마에 대해 검증하고, 구조화된 에러 리포트를 생성합니다.

**아키텍처**: ExcelDbapiReader로 데이터를 읽고, PydanticBackend로 행별 검증 수행.

```python
class ExcelValidator:
    def __init__(
        self,
        mappings: list[ExcelMapping],
        *,
        backend: str = "pydantic",
    ) -> None:
        self._reader = ExcelDbapiReader(read_only=True)  # excel-dbapi 기반 리더
        self._backends = {
            mapping.sheet_name: PydanticBackend(mapping)
            for mapping in mappings
        }

    def validate(
        self,
        source: str | Path | BinaryIO,
        *,
        sheet_name: str | None = None,
        max_errors: int | None = None,
        stop_on_first_error: bool = False,
    ) -> ValidationReport: ...
```

**검증 파이프라인**:

```
.xlsx 파일 (업로드)
    │
    ▼
ExcelDbapiReader.read(source, sheet_name)
    │ → SELECT * FROM SheetName via excel-dbapi
    ▼
ReaderResult(headers, rows, total_rows)
    │
    ▼
_build_header_map() → 헤더 → 컬럼 이름 매핑
    ├── normalize_header(column.name)
    └── normalize_header(column.excel_header)
    │
    ▼
각 행에 대해:
    ├── _remap_row() → 헤더를 ORM 컬럼 이름으로 변환
    ├── PydanticBackend.validate_row(row_data, row_number)
    │   ├── _coerce_value() → 경량 타입 변환
    │   │   ├── "" → None
    │   │   ├── str → int/float/Decimal/date/datetime
    │   │   ├── Enum → value 문자열
    │   │   └── bool: "true"/"yes"/"1" → True
    │   ├── Pydantic model_validate() → 엄격한 타입 검증
    │   └── PydanticValidationError → CellError 변환
    ├── max_errors 체크 → 초과 시 중단
    └── stop_on_first_error 체크 → 첫 에러 시 중단
    │
    ▼
ValidationReport(errors, total_rows, valid_rows, invalid_rows)
```

**동적 Pydantic 모델 생성** (`pydantic_backend.py`):

```python
def _create_pydantic_model(mapping: ExcelMapping) -> type[BaseModel]:
    field_definitions = {}
    for column in mapping.columns:
        field_type = _field_type_for_column(column)
        # Enum → Literal[values], nullable → type | None
        field_info = Field(
            default=None if column.nullable else ...,
            max_length=column.max_length,
        )
        field_definitions[column.name] = (field_type, field_info)

    return create_model(f"{model_name}Validator", **field_definitions)
```

**CellError 및 ValidationReport**:

```python
@dataclass(frozen=True)
class CellError:
    row: int              # Excel 행 번호 (1-based)
    column: str           # 컬럼 이름
    value: Any            # 원본 셀 값
    expected_type: str    # 예상 타입 설명
    message: str          # 사람이 읽을 수 있는 에러 메시지
    error_code: str       # 머신 리더블 에러 코드
                          # null_error, type_error, length_error,
                          # enum_error, constraint_error

@dataclass
class ValidationReport:
    errors: list[CellError]
    total_rows: int
    valid_rows: int
    invalid_rows: int

    def summary(self) -> str: ...
    def to_dict(self) -> dict: ...           # JSON 직렬화
    def errors_by_row(self) -> dict: ...     # 행별 그룹핑
    def to_excel(self, path) -> None: ...    # 에러 리포트 Excel 내보내기
```

### 3.7 `load/` — Database Import

**역할**: 검증된 Excel 데이터를 SQLAlchemy Session을 통해 DB에 로드합니다.

**아키텍처**: Strategy 패턴으로 Insert/Upsert/DryRun 모드를 분리.

```python
class ExcelImporter:
    def __init__(self, mappings: list[ExcelMapping], session: Session) -> None:
        self._mappings = mappings
        self._session = session

    def insert(self, source, *, batch_size=1000, validate=True) -> ImportResult: ...
    def upsert(self, source, *, batch_size=1000, validate=True) -> ImportResult: ...
    def dry_run(self, source, *, validate=True) -> ImportResult: ...
```

**Import 파이프라인** (`_run` 메서드):

```
Excel 소스
    │
    ▼
[validate=True] → ExcelValidator.validate(source)
    ├── has_errors → ImportResult(failed=invalid_rows, errors=...) 조기 반환
    └── 검증 통과 → 계속
    │
    ▼
ExcelDbapiReader(read_only=True).read(source, sheet_name)
    │ → SELECT * FROM SheetName via excel-dbapi
    ▼
_extract_rows_for_mapping() → _align_row() 적용
    ├── normalize_header()로 헤더 정규화
    └── ORM 컬럼 이름에 맞게 행 재정렬
    │
    ▼
strategy.execute(session, model_class, rows, key_columns, batch_size)
    │
    ├── InsertStrategy
    │   ├── _chunk(rows, batch_size)로 배치 분할
    │   ├── session.begin_nested() → savepoint
    │   ├── model_class(**row)로 ORM 객체 생성
    │   ├── session.add_all(objects) → 배치 추가
    │   ├── session.flush() → DB 반영
    │   ├── IntegrityError → savepoint.rollback() → failed 카운트
    │   └── 성공 → savepoint.commit()
    │
    ├── UpsertStrategy
    │   ├── _chunk(rows, batch_size)로 배치 분할
    │   ├── session.begin_nested() → savepoint
    │   ├── 각 행: key_columns로 기존 레코드 조회
    │   │   ├── 없음 → session.add(model_class(**row)) → inserted++
    │   │   └── 있음 → setattr(existing, k, v) → updated++
    │   ├── session.flush()
    │   ├── IntegrityError → _recover_failed_batch()
    │   │   └── 행별 재시도 (개별 savepoint)
    │   └── 성공 → savepoint.commit()
    │
    └── DryRunStrategy
        ├── InsertStrategy와 동일한 로직
        └── finally: savepoint.rollback() (항상 롤백)
    │
    ▼
ImportResult(inserted, updated, skipped, failed, errors, duration_ms)
```

**ImportResult**:

```python
@dataclass(slots=True)
class ImportResult:
    inserted: int = 0
    updated: int = 0
    skipped: int = 0
    failed: int = 0
    errors: list[str] = field(default_factory=list)
    duration_ms: float = 0.0

    @property
    def total(self) -> int: ...     # inserted + updated + skipped + failed
    def summary(self) -> str: ...   # 요약 문자열
```

**설계 결정**:
- 트랜잭션 경계는 호출자가 관리 (`session.commit()`)
- 내부적으로 `session.flush()` + savepoint만 사용
- UpsertStrategy의 `_recover_failed_batch()`: 배치 실패 시 행별 개별 재시도
- DryRunStrategy: `finally` 블록에서 항상 savepoint 롤백 → 데이터 영구 저장 안 됨

### 3.8 `export.py` — Query Result Export

**역할**: SQLAlchemy 쿼리 결과를 서식 적용된 Excel 파일로 내보냅니다. **ExcelWorkbookSession**을 사용합니다.

**excel-dbapi 통합 방식**:

```python
class ExcelExporter:
    def export(self, rows, path=None, *, sheet_name=None) -> bytes | None:
        if path is not None:
            with ExcelWorkbookSession.open(path, create=True) as session:
                self._populate_workbook(session.workbook, rows, sheet_name)
                session.commit()
            return None

        # path=None → 임시 파일 → bytes 반환
        temp_path = tempfile.NamedTemporaryFile(suffix=".xlsx")
        self.export(rows, temp_path)
        return open(temp_path, "rb").read()
```

**내보내기 기능**:
- 헤더 스타일링 (파란색 배경, 흰색 볼드 폰트)
- 날짜/시간 포맷팅 (`YYYY-MM-DD`, `YYYY-MM-DD HH:MM:SS`)
- 컬럼 너비 자동 조정 (최대 50자)
- Auto-filter 및 Freeze panes
- ORM 인스턴스 또는 딕셔너리 행 지원
- `sanitize_cell_value()`로 수식 주입 방지

### 3.9 `integrations/fastapi.py` — FastAPI Router Factory

**역할**: 단일 모델로부터 템플릿 다운로드, 업로드 검증, 임포트 엔드포인트를 자동 생성합니다.

```python
def create_import_router(
    model: type,
    *,
    prefix: str = "",
    tags: list[str] | None = None,
    session_dependency: Any = None,
) -> APIRouter:
```

**생성되는 엔드포인트**:

| Method | Path | 설명 |
|--------|------|------|
| `GET` | `{prefix}/template` | Excel 템플릿 다운로드 (샘플 데이터 포함) |
| `POST` | `{prefix}/validate` | 업로드된 Excel 파일 검증 → JSON 리포트 |
| `POST` | `{prefix}/import` | 검증 후 DB 임포트 (session_dependency 필요) |
| `GET` | `{prefix}/health` | 헬스 체크 |

### 3.10 `cli.py` — CLI Interface

**역할**: Click 기반 5개 커맨드 제공.

| 커맨드 | 설명 |
|--------|------|
| `template` | ORM 모델에서 Excel 템플릿 생성 |
| `validate` | Excel 파일 검증 (text/json/excel 출력) |
| `import` | Excel → DB 임포트 (insert/upsert/dry-run) |
| `export` | DB → Excel 내보내기 |
| `inspect` | Excel 파일 구조 검사 (모델 불필요) |

---

## 4. Data Flow Diagrams

### 4.1 Template Generation Flow

```
ORM Model
    │
    ▼
ExcelMapping.from_model(Model)
    │ sa_inspect() → mapper → columns → ColumnMapping 리스트
    ▼
ExcelMapping (frozen dataclass)
    │
    ▼
ExcelTemplate([mapping], include_sample_data=True)
    │
    ▼
ExcelWorkbookSession.open(path, create=True)
    │ → excel_dbapi.connect(path, engine="openpyxl", autocommit=False, create=True)
    ▼
session.workbook → openpyxl.Workbook
    │
    ├── 기존 시트 삭제
    ├── create_sheet(title=sheet_name)
    ├── 헤더 행: 스타일 + 코멘트 + DataValidation
    ├── 컬럼 너비, auto-filter, freeze panes
    └── 샘플 데이터 행 (선택)
    │
    ▼
session.commit()
    │ → conn.commit() → engine.save() (atomic: tempfile + os.replace)
    ▼
.xlsx 파일 (또는 BytesIO)
```

### 4.2 Upload → Validate → Import Flow

```
.xlsx 파일 (업로드)
    │
    ▼
ExcelDbapiReader.read(source, sheet_name)
    │ → _resolve_source(): BinaryIO → 임시 파일 변환 (필요 시)
    │ → excel_dbapi.connect(path, engine="openpyxl", autocommit=True, data_only=True)
    │ → cursor.execute("SELECT * FROM SheetName")  # 따옴표 없이!
    │ → cursor.description → 헤더 추출 + 정규화
    │ → cursor.fetchall() → 행 딕셔너리 리스트
    ▼
ReaderResult(headers, rows, total_rows)
    │
    ▼
ExcelValidator.validate(source)
    │ → PydanticBackend.validate_row() per row
    │ → 타입 강제변환, nullable 체크, enum 체크, 길이 체크
    ▼
ValidationReport
    ├── [has_errors] → report.to_excel() → 에러 리포트
    │   또는 report.to_dict() → JSON 응답
    │
    └── [clean] →
        │
        ▼
    ExcelImporter.insert() 또는 .upsert()
        │ → ExcelDbapiReader.read() (재읽기)
        │ → _align_row() → ORM 컬럼에 정렬
        │ → strategy.execute()
        │   ├── InsertStrategy: session.add_all() per batch + savepoint
        │   ├── UpsertStrategy: select → exists? update : insert + savepoint
        │   └── DryRunStrategy: insert + always rollback
        ▼
    ImportResult(inserted, updated, skipped, failed, errors, duration_ms)
        │
        ▼
    session.commit()  ← 호출자 책임
```

### 4.3 Export Flow

```
SQLAlchemy Query
    │
    ▼
session.execute(select(Model)).scalars().all()
    │
    ▼
List[ORM 인스턴스]
    │
    ▼
ExcelExporter([mapping]).export(rows, path)
    │
    ▼
ExcelWorkbookSession.open(path, create=True)
    │ → excel_dbapi.connect(path, engine="openpyxl", autocommit=False, create=True)
    ▼
session.workbook → openpyxl.Workbook
    │
    ├── 기존 시트 삭제
    ├── create_sheet(title=sheet_name)
    ├── 헤더 행: 스타일 적용
    ├── 데이터 행: ORM 인스턴스/dict에서 값 추출
    │   ├── getattr(row, column_name) 또는 row.get(column_name)
    │   ├── sanitize_cell_value() → 수식 주입 방지
    │   └── datetime/date → 포맷 적용
    ├── 컬럼 너비 자동 조정 (최대 50자)
    ├── Auto-filter + Freeze panes
    │
    ▼
session.commit()
    │
    ▼
.xlsx 파일 (또는 bytes)
```

---

## 5. Error Handling Architecture

### 5.1 Exception Hierarchy

```
SqlalchemyExcelError (base)
├── MappingError              # ORM 인트로스펙션 실패
├── TemplateError             # 템플릿 생성 실패
├── ReaderError               # Excel 파싱 실패
│   ├── FileFormatError       # 유효하지 않은 .xlsx 파일
│   ├── SheetNotFoundError    # 시트 이름 없음
│   └── HeaderMismatchError   # 헤더/컬럼 불일치
├── ValidationError           # 데이터 검증 실패 (report 래핑)
├── ImportError_              # DB 임포트 실패 (builtin 충돌 방지 언더스코어)
│   ├── DuplicateKeyError     # 유니크 제약 조건 위반
│   └── ConstraintViolationError  # 기타 DB 제약 조건 위반
└── ExportError               # 내보내기 실패
```

### 5.2 Error Recovery Strategy

| 모듈 | 에러 처리 방식 |
|------|---------------|
| **mapping** | `MappingError` 즉시 발생 — 잘못된 모델 설정은 조기 실패 |
| **template** | 모든 예외를 `TemplateError`로 래핑하여 재발생 |
| **reader** | openpyxl/excel-dbapi 예외를 `FileFormatError`, `SheetNotFoundError`, `ReaderError`로 변환 |
| **validation** | 예외를 발생시키지 않음 — 모든 에러를 `ValidationReport`에 수집 |
| **import** | savepoint 기반 트랜잭션 롤백 + 부분 결과를 `ImportResult`에 기록 |
| **export** | 모든 예외를 `ExportError`로 래핑하여 재발생 |

### 5.3 excel-dbapi 예외 매핑

```
excel-dbapi Exception          → sqlalchemy-excel Exception
─────────────────────────────────────────────────────────
InvalidFileException, BadZipFile → FileFormatError
Sheet not in workbook.sheetnames → SheetNotFoundError
Connection failure               → ReaderError
Generic Exception                → ReaderError (래핑)
```

---

## 6. Security Architecture

### 6.1 Threats & Mitigations

| 위협 | 방어 수단 |
|------|----------|
| XML Bomb (billion laughs) | `defusedxml` 필수 의존성, import 시 확인 |
| 수식 주입 (`=`, `+`, `-`, `@`, `\t`, `\r`) | `sanitize_cell_value()`: 위험 접두사에 `'` 추가 |
| 파일 크기 DoS | `ExcelDbapiReader.max_file_size` (기본 50MB) |
| Zip bomb | `_validate_file_size_path()` + openpyxl 내부 보호 |
| 경로 순회 | 파일 경로 검증, 업로드 시 tempfile 사용 |
| SQL 주입 | 모든 DB 작업은 SQLAlchemy ORM 경유 (파라미터화됨) |

### 6.2 Security Initialization

```python
# __init__.py — 패키지 임포트 시 즉시 실행
from sqlalchemy_excel._compat import ensure_defusedxml
ensure_defusedxml()  # defusedxml 미설치 시 ImportError 발생
```

### 6.3 Cell Value Sanitization

```python
# _compat.py
_FORMULA_PREFIXES = ("=", "+", "-", "@", "\t", "\r")

def sanitize_cell_value(value: str) -> str:
    if value.startswith(_FORMULA_PREFIXES):
        return f"'{value}"  # 아포스트로피 접두사 → Excel 수식 해석 방지
    return value
```

사용 위치: `template.py` (샘플 데이터), `export.py` (내보내기 값), `report.py` (에러 리포트)

---

## 7. Performance Considerations

### 7.1 Memory Management

| 시나리오 | 접근 방식 |
|---------|----------|
| 소형 파일 (< 50MB) | 전체 로드: excel-dbapi가 openpyxl로 메모리에 적재 |
| 검증 | 행 단위 처리: Pydantic 모델로 개별 검증 |
| DB 임포트 | 배치 처리: `batch_size` 단위로 flush (기본 1000행) |

### 7.2 Optimization Targets

| 작업 | 목표 | 접근 방식 |
|------|------|----------|
| 템플릿 생성 | < 100ms | 스타일 사전 정의, 배치 셀 쓰기 |
| 10K행 파싱 | < 3s | excel-dbapi SQL 쿼리 기반 읽기 |
| 10K행 검증 | < 1s | Pydantic v2 컴파일된 검증기 |
| 10K행 임포트 | < 5s | 배치 add_all(), savepoint 기반 |
| 10K행 내보내기 | < 3s | ExcelWorkbookSession 직접 셀 쓰기 |

### 7.3 excel-dbapi Performance Characteristics

- **읽기**: `SELECT * FROM SheetName` → openpyxl `iter_rows()` 내부 사용
- **쓰기**: `session.commit()` → atomic save (tempfile + `os.replace`)
- **스냅샷**: 커밋마다 BytesIO 스냅샷 생성 (대형 파일에서 비용 발생)
- **BinaryIO 입력**: 임시 파일 생성 → 추가 디스크 I/O

---

## 8. Testing Architecture

### 8.1 Test Structure

```
tests/
├── conftest.py              # 공유 픽스처
│   ├── in_memory_engine     # SQLite :memory:
│   ├── sample_models        # User, Product, Order 모델
│   └── temp_workbook        # 임시 .xlsx 생성 헬퍼
├── unit/
│   ├── test_mapping.py      # ExcelMapping 추출
│   ├── test_template.py     # 템플릿 생성
│   ├── test_reader.py       # Excel 파일 파싱
│   ├── test_validation.py   # 검증 엔진
│   ├── test_importer.py     # DB 임포트
│   ├── test_report.py       # ValidationReport
│   └── test_export.py       # 쿼리 내보내기
├── integration/
│   └── test_end_to_end.py   # 전체 파이프라인
└── fixtures/                # 테스트용 .xlsx 파일
```

### 8.2 Test Count & Coverage

- **단위 테스트**: 7개 파일
- **통합 테스트**: 1개 파일 (전체 파이프라인)
- **총 테스트 수**: 117개, 모두 통과
- **CI 매트릭스**: Python 3.10–3.13

### 8.3 Test Invariants (Hypothesis)

1. **Round-trip**: `template → fill with valid data → validate → no errors`
2. **Idempotent validation**: `validate(data) == validate(data)` (동일 리포트)
3. **Export/Import**: `export(query) → import(file) → same data in DB`

---

## 9. Integration Architecture: excel-dbapi

### 9.1 excel-dbapi as Core Dependency

`excel-dbapi >= 1.0`은 `pyproject.toml`의 `dependencies`에 포함된 **전면 의존성**입니다. 선택적(optional)이 아닙니다. 모든 Excel I/O는 excel-dbapi를 통해 수행됩니다.

```toml
dependencies = [
    "sqlalchemy>=2.0",
    "openpyxl>=3.1",
    "pydantic>=2.0",
    "defusedxml>=0.7",
    "click>=8.0",
    "excel-dbapi>=1.0",    # 전면 의존성
]
```

### 9.2 Dual-Channel Architecture

sqlalchemy-excel은 excel-dbapi를 **두 가지 채널**로 활용합니다:

```
┌──────────────────────────────────────────────────────────────┐
│  excel-dbapi Usage in sqlalchemy-excel                        │
│                                                               │
│  ┌─────────────────────────────────────────────────────────┐ │
│  │  Format Channel (쓰기)                                   │ │
│  │                                                          │ │
│  │  ExcelWorkbookSession                                    │ │
│  │  ├── conn = excel_dbapi.connect(path, create=True)      │ │
│  │  ├── conn.workbook → openpyxl.Workbook                  │ │
│  │  │   ├── cell.font = Font(bold=True)                    │ │
│  │  │   ├── cell.fill = PatternFill(...)                   │ │
│  │  │   └── DataValidation, comments, freeze_panes         │ │
│  │  └── conn.commit() → atomic disk save                   │ │
│  │                                                          │ │
│  │  사용 위치: template.py, export.py                       │ │
│  └─────────────────────────────────────────────────────────┘ │
│                                                               │
│  ┌─────────────────────────────────────────────────────────┐ │
│  │  Data Channel (읽기)                                     │ │
│  │                                                          │ │
│  │  ExcelDbapiReader                                        │ │
│  │  ├── conn = excel_dbapi.connect(path, data_only=True)   │ │
│  │  ├── cursor = conn.cursor()                              │ │
│  │  ├── cursor.execute("SELECT * FROM SheetName")           │ │
│  │  │   └── 테이블 이름 따옴표 없이 사용!                   │ │
│  │  ├── cursor.description → 컬럼 메타데이터                │ │
│  │  └── cursor.fetchall() → List[Tuple]                    │ │
│  │                                                          │ │
│  │  사용 위치: validation/engine.py, load/importer.py       │ │
│  └─────────────────────────────────────────────────────────┘ │
└──────────────────────────────────────────────────────────────┘
```

### 9.3 Connection Configuration

| 용도 | engine | autocommit | create | data_only |
|------|--------|-----------|--------|-----------|
| 템플릿 생성 | `"openpyxl"` | `False` | `True` | `False` |
| 파일 내보내기 | `"openpyxl"` | `False` | `True` | `False` |
| 파일 읽기/검증 | `"openpyxl"` | `True` | `False` | `True` |
| 파일 읽기/임포트 | `"openpyxl"` | `True` | `False` | `True` |

### 9.4 Contract Requirements from excel-dbapi

sqlalchemy-excel이 excel-dbapi에 요구하는 계약:

1. `connect()` → `ExcelConnection` 반환 (표준 DB-API 메서드 포함)
2. `conn.workbook` → openpyxl `Workbook` 객체 반환 (openpyxl 엔진 전용)
3. `cursor.description` → PEP 249 7-tuple 컬럼 메타데이터
4. `cursor.fetchall()` → `List[Tuple]`
5. `create=True` → 빈 워크북이 포함된 유효한 .xlsx 파일 생성
6. 비인용 테이블 이름 → `SELECT * FROM Sheet1` (따옴표 → lookup 실패)
7. 원자적 저장 → tempfile + `os.replace`
8. 모든 예외는 PEP 249 계층을 따름
