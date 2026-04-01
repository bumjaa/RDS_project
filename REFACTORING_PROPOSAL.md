# RDS 리팩토링 제안서

## 현재 상태 요약

| 항목 | 현황 | 심각도 |
|------|------|--------|
| 배포 | U: 네트워크 공유 + GitHub 혼재, 이중 버전 체계 | 🔴 Critical |
| 자격증명 | GitHub 토큰·API키 코드에 평문 노출 | 🔴 Critical |
| 경로 | U:, s: 네트워크 드라이브 하드코딩 | 🔴 Critical |
| 표준 데이터 | 20+ 표준이 Case문으로 하드코딩 (standard_module) | 🟡 Major |
| 에러 처리 | On Error Resume Next 남발 (실패를 감춤) | 🟡 Major |
| 코드 중복 | URLEncodeUTF8 등 폼 간 복붙, Build*JSON 4개 거의 동일 | 🟡 Major |
| SQL 인젝션 | personalDB_module에서 문자열 결합으로 SQL 구성 | 🟠 Moderate |

---

## Phase 1: 설정 분리 + 배포 체계 통합 (우선)

### 1-1. 외부 설정 파일 도입 (`config.json`)

**현재:** 경로·API키·버전이 각 모듈에 흩어져 하드코딩
```vba
' version_module.bas:58
token = "<GITHUB_TOKEN_REMOVED>"
' update_module.bas:4
Const API_URL As String = "https://script.google.com/macros/s/..."
' personalDB_module.bas:26
"Data Source=" & "s:\rdsDB" & "\personalDB.accdb;"
```

**개선:** `config.json` 파일 하나로 통합
```json
{
  "version": "1.1.0",
  "github": {
    "repo": "bumjaa/RDS_project",
    "token_env": "RDS_GITHUB_TOKEN"
  },
  "api": {
    "base_url": "https://script.google.com/macros/s/.../exec",
    "key_env": "RDS_API_KEY"
  },
  "paths": {
    "personal_db": "s:\\rdsDB\\personalDB.accdb",
    "environment_data": "U:\\EMC센터\\01. 공통관리\\01. 시험환경 환경\\07. 환경일 기록\\",
    "original_copy": "U:\\EMC센터\\02. QC\\05. RDS 프로젝트\\RDSDB\\RDS_origianlCopy.xlsm"
  }
}
```

**새 모듈:** `Config_module.bas`
- `LoadConfig()` — config.json 읽기 (ThisWorkbook.Path 기준)
- `GetConfigValue(key)` — 점 표기법으로 설정값 접근
- 환경변수에서 자격증명 읽기 (`Environ()`)
- config.json은 `.gitignore`에 추가, `config.sample.json`만 Git 관리

### 1-2. 버전 체계 통합

**현재:** 3곳에서 서로 다른 버전 관리
- `version_module.bas:9` → `currentVersion = "1.0"` (GitHub raw)
- `version_module.bas:59` → `currentVer = "0.0.01"` (GitHub API + token)
- `update_module.bas:18` → `curVersion = "1.00.02"` (Google Sheets API)

**개선:** 단일 버전 소스 (`config.json`의 `version` 필드)
- `version_module.bas`와 `update_module.bas`를 **`Updater_module.bas`**로 통합
- 버전 체크: GitHub API 단일 경로
- 업데이트 흐름: GitHub → 다운로드 → 시트/모듈 교체
- 네트워크 공유(U:) 의존 제거

### 1-3. 배포 프로세스

```
[개발자 PC]
    ↓ git push (코드 + version.txt)
[GitHub: bumjaa/RDS_project]
    ↓ 사용자가 Excel 열 때 자동 체크
[사용자 PC]
    ↓ 버전 다르면 → 다운로드 + 자동 적용
[Excel xlsm 업데이트 완료]
```

- Release 태그 기반 배포 (GitHub Releases API)
- 시트 업데이트와 모듈 업데이트를 분리된 이벤트로 관리
- 롤백 지원: 이전 버전 xlsm 백업 후 업데이트

---

## Phase 2: 데이터-코드 분리

### 2-1. standard_module 리팩토링

**현재:** 20+ 표준 × 2개 함수(`standardFunction`, `DataSuffix`)가 거대 Case문
- 표준 추가 시 양쪽 모두 수정 필요
- 같은 표준의 쉼표/공백 변형이 Case에 나열됨 (예: `"KS C 9832 KS C 9835"`, `"KS C 9832, KS C 9835"`)

**개선:** Google Sheets의 `Standards` 시트로 데이터 이관
```
| StandardKey | Aliases              | Functions                    | DataSuffixes                    |
|-------------|----------------------|------------------------------|---------------------------------|
| KS_9832     | KS C 9832, KS C 9835| CE,ISN,CDVE,OUTCDVE,RE,...  | ESD_ENCLOSURE,RS_ENCLOSURE,...  |
```

- 앱 시작 시 한 번 로드 → Dictionary에 캐싱
- 표준명 정규화 함수 추가 (쉼표/공백 변형 자동 처리)
- **표준 추가/수정이 코드 배포 없이 가능**

### 2-2. 계측기/주변기기 데이터 캐싱 개선

**현재:** `Ins_module`이 매번 Google Sheets 전체 데이터를 가져옴 (541줄)

**개선:**
- 로컬 캐시 시트 도입 (숨김 시트)
- 마지막 갱신 시각 기록, TTL 기반 재조회
- 오프라인 모드 지원 (캐시가 있으면 네트워크 없이도 동작)

---

## Phase 3: 코드 품질 개선

### 3-1. 중복 코드 제거

| 중복 항목 | 위치 | 통합 방안 |
|-----------|------|-----------|
| `URLEncodeUTF8` | Instruments.frm, Peripherals.frm | → `Encoder_module.bas`로 이동 |
| `Build*ConfigJSON` | personalDB_module (4개 함수) | → 범용 `BuildRangeJSON(rangeName, colIndices, startRow)` |
| HTTP 요청 패턴 | 거의 모든 모듈 | → `HttpGet(url, headers)` 헬퍼 |

### 3-2. 에러 처리 정비

**현재:** `On Error Resume Next`가 정당한 용도(시트 존재 확인 등)와 에러 무시 용도가 혼재

**개선:**
- 정당한 용도: `SheetExists`, `RangeExists` 등 헬퍼 함수로 대체
- API 호출: 통합 에러 핸들러 + 사용자 메시지
- DB 작업: 트랜잭션 래퍼

### 3-3. SQL 인젝션 방지

**현재:** `personalDB_module`에서 문자열 결합으로 SQL 구성
```vba
strSQL = "SELECT ... WHERE Order_No = '" & orderNoParam & "'"
```

**개선:** ADODB Command + Parameter 사용
```vba
Set cmd = New ADODB.Command
cmd.CommandText = "SELECT ... WHERE Order_No = ?"
cmd.Parameters.Append cmd.CreateParameter(, adVarWChar, adParamInput, 50, orderNoParam)
```

---

## Phase 4: UX 개선

### 4-1. 업데이트 진행률 표시
- StatusBar 활용한 진행 상태 표시
- 업데이트 실패 시 명확한 에러 메시지 + 복구 안내

### 4-2. 오프라인 대응
- 네트워크 불가 시 graceful degradation
- 캐시된 데이터로 기본 기능 동작

---

## 실행 순서 (권장)

```
Phase 1-1  Config_module 생성 + config.json 도입
    ↓
Phase 1-2  Updater_module 통합 (version + update 합침)
    ↓
Phase 1-3  GitHub Releases 기반 배포 흐름 구현
    ↓
Phase 2-1  standard_module → 데이터 시트 이관
    ↓
Phase 3-1  중복 코드 제거 (URLEncode, Build*JSON, HTTP 헬퍼)
    ↓
Phase 3-2  에러 처리 정비
    ↓
Phase 3-3  SQL 파라미터화
    ↓
Phase 2-2  캐싱 개선
    ↓
Phase 4    UX 개선
```

각 Phase는 독립적으로 배포 가능하며, Phase 1이 완료되면 이후 Phase의 배포가 자동화됨.

---

## 파일 구조 변경 (예상)

```
project_rds/
├── config.sample.json          ← NEW: 설정 템플릿 (Git 관리)
├── config.json                 ← NEW: 실제 설정 (gitignore)
├── Config_module.bas           ← NEW: 설정 로더
├── Updater_module.bas          ← NEW: version_module + update_module 통합
├── Http_module.bas             ← NEW: HTTP 요청 헬퍼
├── standard_module.bas         ← REFACTOR: Case문 → 데이터 조회
├── personalDB_module.bas       ← REFACTOR: SQL 파라미터화 + 중복 제거
├── Encoder_module.bas          ← REFACTOR: URLEncode 통합
├── Ins_module.bas              ← REFACTOR: 캐싱 개선
├── ...기존 모듈들...
├── version_module.bas          ← DELETE: Updater_module로 통합
├── update_module.bas           ← DELETE: Updater_module로 통합
└── CLAUDE.md
```
