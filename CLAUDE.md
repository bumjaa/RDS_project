# RDS Project - EMC Test Record Form (시험기록지)

## Project Overview
EMC(전자파적합성) 시험기록지 관리 시스템. Excel VBA 기반.
- 각종 국제 표준(KS, KN, EN, FCC 등) 시험 기록 관리
- 시험 환경, 계측기, 주변기기 데이터 연동
- 시험 레이아웃 다이어그램 생성/내보내기
- Access DB(personalDB)로 개인 데이터 저장 (JSON 직렬화)

## Architecture
```
[Excel xlsm] ← 사용자 인터페이스 + 시험 시트
    ├── VBA Modules (.bas) ← 비즈니스 로직
    ├── UserForms (.frm) ← UI 폼
    └── Classes (.cls) ← 이벤트 핸들러
        │
        ├──→ config.json ← 모든 설정/경로/자격증명
        ├──→ standards_data.json ← 표준↔시험함수 매핑 데이터
        ├──→ Google Sheets API (preset, 계측기, 주변기기, 버전)
        ├──→ GitHub API (버전 체크, 파일 다운로드)
        ├──→ Access DB (personalDB.accdb)
        └──→ Network Share (원본 파일, .bas 배포)
```

## Module Responsibilities
| Module | Role |
|--------|------|
| **Config_module** | config.json 로드, GetCfg() 점 표기법 설정 접근 |
| **Http_module** | HTTP GET/다운로드/JSON 헬퍼 (MSXML2 래퍼) |
| **Encoder_module** | URLEncode + URLEncodeUTF8 통합 |
| standard_module | 표준↔시험함수 매핑 (standards_data.json에서 로드) |
| Layout_module | 시험 배치도 그리기/내보내기 |
| Ins_module | 계측기 데이터 (Google Sheets 연동, 캐싱) |
| personalDB_module | Access DB CRUD (JSON 직렬화, BuildRangeJSON 범용 빌더) |
| preset_module | 시험 프리셋 로드 (Google Sheets) |
| version_module | GitHub 기반 버전 체크 + 자동 업데이트 |
| update_module | 시트/모듈 업데이트 (이벤트 기반) |
| DataInsert_module | 시험 결과 자동 입력 |
| Environment_module | 환경 데이터 조회 (ADO) |
| Function_module | 유틸리티 (날짜, 오디오 계산) |
| LineModule | 동적 범위 확장/축소 |
| LayoutFunction | 레이아웃 헬퍼 |
| uiux_module | 체크박스 스타일링 |
| rename_module | PDF 파일 이름 변경 |
| makeSheets | 시트 생성 |
| JsonConverter | JSON 파서 (외부 라이브러리) |

## Configuration
- **config.json**: 모든 설정값 (Git 미추적, .gitignore)
- **config.sample.json**: 설정 템플릿 (Git 추적)
- **standards_data.json**: 표준 데이터 (Git 추적, 코드 배포 없이 표준 추가 가능)

설정값 접근: `GetCfg("api.key")`, `GetPersonalDBConn()`, `GetVersion()` 등

## Key External Dependencies
- **Google Sheets API**: preset, 계측기, 주변기기, 버전 데이터
- **GitHub**: bumjaa/RDS_project (버전 관리, 배포)
- **Access DB**: config.json → paths.personal_db
- **Network Share**: config.json → paths.original_copy, paths.bas_dir

## Development Notes
- VBA 코드는 .bas/.frm/.cls 파일로 export되어 Git 관리
- .frx 파일은 바이너리(폼 리소스)이므로 .gitignore 처리
- 원본 xlsm은 참조용으로만 보관 (코드 수정은 .bas 파일에서)
- 한글 인코딩: .bas 파일은 VBA 기본 인코딩 사용

## Refactoring Status (완료)
- [x] 하드코딩된 경로/자격증명 → config.json 분리
- [x] 이중 버전 체계 → Config_module의 단일 GetVersion() 사용
- [x] 표준 데이터 → standards_data.json으로 분리
- [x] HTTP 요청 패턴 → Http_module로 통합
- [x] URLEncodeUTF8 중복 → Encoder_module로 통합
- [x] Build*JSON 중복 → BuildRangeJSON 범용 함수
- [x] SQL 인젝션 기본 방어 (quote escaping)
