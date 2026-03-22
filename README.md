# Museum Item Card Assistant

박물관·미술관·기록물 관리용 카드 양식을 데이터 기반으로 렌더링하고, 개별 또는 일괄 PDF로 내보내는 정적 웹 애플리케이션입니다.

이 저장소는 기존 Excel/VBA 기반 `museum_card-v0.3.xlsm`의 작업 흐름을 GitHub Pages에서 운영 가능한 형태로 옮기는 것을 목표로 합니다.

## 목표

- 데이터 파일을 업로드하면 각 행을 카드/서식 단위로 렌더링합니다.
- 여러 개의 양식을 템플릿으로 관리합니다.
- 왼쪽에서 값 편집, 오른쪽에서 실시간 미리보기를 제공합니다.
- 현재 행 단건 PDF와 여러 행 일괄 PDF ZIP 다운로드를 지원합니다.
- 서버 없이 GitHub Pages에 배포 가능한 구조를 유지합니다.

## 원본 Excel/VBA 기준 동작

`data/museum_card-v0.3.xlsm` 내부 로직과 README 시트를 기준으로 다음 규칙을 유지합니다.

- 데이터 시트: `MuseumData`
- 양식 시트: `관리카드양식`
- 제어 컬럼
  - `<<개별저장>>`: 현재 행 단건 저장
  - `<<미리보기>>`: 현재 행 미리보기
  - `<<일괄여부>>`: 일괄 저장 대상 표시
  - `<<일괄저장>>`: 전체 또는 선택 행 일괄 저장
- 파일명 규칙: `NAMOK-자료번호-세부번호.pdf`
- 예외 처리: 자료번호/세부번호가 비면 `Unknown`
- 금지문자 치환: `/ \ : * ? " < > |`

확인한 VBA 핵심 함수는 아래와 같습니다.

- `PrintCardByRow`
- `BatchSaveCards`
- `RenderTemplateText`
- `ResolveToken`
- `ResolveChoiceToken`
- `ResolveChoiceMapToken`
- `GetCellValueByHeader`
- `GetHeaderColumn`

## 템플릿 문법

웹 버전은 Excel 양식의 토큰 문법을 그대로 계승합니다.

- 일반 치환: `[[명칭]]`
- 복합 치환: `[[수량]] [[수량단위]]`
- 선택형 출력: `[[CHOICE:저작권(유무)|유,무]]`
- 코드-라벨 매핑: `[[CHOICEMAP:손상등급|1=1등급,2-1=2-1등급,2-2=2-2등급,3=3등급]]`

웹 구현에서는 `CHOICE`와 `CHOICEMAP`을 실제 선택 상태가 드러나도록 렌더링합니다.

- `CHOICE`: 선택된 항목은 체크된 상태로 표시
- `CHOICEMAP`: 실제 값에 해당하는 라벨을 강조 표시

## 템플릿 구조

템플릿은 `/templates` 아래 폴더 단위로 관리합니다.

```text
templates/
├─ catalog.json
├─ museum-management-card/
│  ├─ meta.json
│  └─ template.html
├─ exhibition-label/
│  ├─ meta.json
│  └─ template.html
└─ condition-report/
   ├─ meta.json
   └─ template.html
```

각 템플릿은 다음 두 파일을 가집니다.

- `meta.json`: 이름, 설명, 용지 크기, 파일명 규칙, 대표 필드, 태그
- `template.html`: 실제 렌더링용 HTML 마크업

## 웹 앱 구조

```text
Museum-Item-Card-Assistant/
├─ index.html
├─ assets/
│  ├─ app.js
│  ├─ styles.css
├─ templates/
│  ├─ catalog.json
│  ├─ museum-management-card/
│  ├─ exhibition-label/
│  └─ condition-report/
└─ data/
   ├─ sample-data.csv
   ├─ sample-data.xlsx
   ├─ museum_card-v0.3.xlsm
   ├─ MICA 변수명.xlsx
   └─ 일괄 소장자료.xlsx
```

## 주요 기능

- `xlsx`, `xlsm`, `csv`, `json` 업로드
- `MuseumData` 시트 우선 인식
- 템플릿 선택 후 즉시 미리보기
- 현재 행 필드 편집
- 템플릿 필요 필드 중심의 입력 폼
- 단건 PDF 다운로드
- 여러 행 PDF 생성 후 ZIP 다운로드
- GitHub Pages 배포 대응 정적 파일 구조

## 사용 흐름

1. 데이터 업로드
   - `xlsx`, `xlsm`, `csv`, `json` 중 하나를 업로드합니다.
   - `MuseumData` 시트가 있으면 우선 사용합니다.

2. 템플릿 선택
   - 기본 제공 템플릿 중 하나를 고릅니다.
   - `museum-management-card`는 Excel `관리카드양식` 구조를 참고해 만든 템플릿입니다.

3. 레코드 선택
   - 목록에서 현재 편집할 행을 고릅니다.
   - 여러 행을 체크하면 일괄 ZIP 내보내기에 사용됩니다.

4. 값 편집
   - 왼쪽 입력 폼에서 현재 행 값을 수정합니다.
   - 오른쪽 미리보기에 즉시 반영됩니다.

5. 내보내기
   - 현재 레코드만 PDF로 저장하거나
   - 선택 행 또는 전체 행을 PDF ZIP으로 저장합니다.

## GitHub Pages 배포

이 프로젝트는 빌드 단계 없이 정적 호스팅이 가능하도록 구성되었습니다.

### 로컬 테스트

정적 파일을 직접 여는 대신 HTTP 서버로 확인하는 편이 안전합니다.

```powershell
./serve-local.ps1
```

기본 주소는 `http://127.0.0.1:8080` 입니다.

### 배포 방법

1. 저장소를 GitHub에 푸시합니다.
2. `Settings > Pages`에서 배포 브랜치를 `main` / root 로 설정합니다.
3. 배포 후 `index.html`이 시작 페이지로 제공됩니다.

### 런타임 외부 라이브러리

클라이언트에서만 사용하는 CDN 스크립트를 로드합니다.

- `SheetJS`: Excel/CSV 읽기
- `html2pdf.js`: PDF 생성
- `JSZip`: 다건 압축 다운로드

정적 호스팅에는 서버 설정이 필요하지 않습니다.

## 주의 사항

- 브라우저 기반 PDF 생성 품질은 사용 브라우저와 폰트 환경의 영향을 받습니다.
- 대량 일괄 출력 시 브라우저 메모리 사용량이 증가할 수 있습니다.
- GitHub Pages에서는 로컬 폰트가 아닌 웹 안전 폰트 또는 공개 웹폰트를 고려해야 합니다.
- 템플릿 내 필드명은 데이터 헤더와 정확히 일치해야 합니다.

## 향후 확장 제안

- 사용자 정의 템플릿 업로드
- 템플릿 시각 편집기
- 이미지 필드 및 바코드 지원
- 다국어 템플릿 세트
- 인쇄 프리셋 및 용지별 레이아웃 저장

## 라이선스

이 프로젝트는 [LICENSE](./LICENSE)를 따릅니다.
