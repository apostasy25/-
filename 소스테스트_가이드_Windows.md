# 빌드 전 Windows 소스 테스트 가이드

EXE로 빌드하기 전에 **Python 소스 상태**로 전체 플로우(런처 → 마스터 관리 → 정산 실행 → PDF → 메일 Draft)를 검증하는 방법입니다. 소스로 먼저 돌리면 문제를 빠르게 고치고, EXE는 마지막에 한 번만 빌드하면 됩니다.

> 핵심: 빌드된 EXE와 소스 실행은 **같은 진입점(`launcher.py`)·같은 로직**을 씁니다. 소스에서 통과하면 EXE도 동일하게 동작합니다(차이는 패키징뿐).

---

## 1. 사전 준비 (한 번만)

### 1-1. Python 설치
- Python 3.11 이상 (3.13 권장). 설치 시 **"tcl/tk and IDLE" 옵션 체크** — 이게 있어야 정산 실행 화면(tkinter)이 뜹니다.
- 설치 후 확인:
  ```
  python --version
  python -c "import tkinter; print('tkinter OK')"
  ```
  `tkinter OK`가 안 나오면 Python을 "Modify → tcl/tk 체크"로 다시 설치하세요.

### 1-2. LibreOffice 설치 (PDF 변환에 필요)
- https://www.libreoffice.org 에서 설치. 기본 경로면 자동 인식됩니다:
  `C:\Program Files\LibreOffice\program\soffice.exe`
- 다른 경로에 설치했다면 환경변수로 지정(아래 3-2).

### 1-3. 파이썬 패키지 설치
프로젝트 폴더에서:
```
pip install -r requirements.txt
```
(openpyxl · pypdf · PySide6 · pyinstaller 설치됨. pyinstaller는 테스트엔 안 쓰지만 미리 깔아둬도 무방)

---

## 2. 폴더 구성

받은 소스 폴더가 아래처럼 되어 있는지 확인합니다(핵심만):
```
프로젝트/
├─ launcher.py            ← 진입점
├─ pipeline_v2.py
├─ securepdf.py  maildraft.py  config.py  daterules.py
├─ engine\  (m3_run.py 등)
├─ gui\
│   ├─ app.py             ← 마스터 관리(PySide6)
│   ├─ pipeline_app.py    ← 정산 실행(tkinter)
│   └─ master\master_io.py
├─ 원작료정산_마스터.xlsx   ← 실제 마스터(개인정보 포함)
└─ inbox\
    └─ 매출리스트\…_테라핀_….xlsx,  기타수익_누적.xlsx
```
- **마스터 파일**과 **RAW(매출리스트·기타수익)**은 실제 파일을 직접 넣어야 합니다.
- 정산 실행 화면에서 "RAW 폴더"로 `inbox`(또는 매출리스트가 든 폴더)를 고르면 매출/기타수익/마스터를 자동 인식합니다.

---

## 3. 환경변수 설정 (PowerShell 기준, 그 세션에서만 유효)

### 3-1. (선택) 마스터 파일 위치 지정
마스터를 exe/소스 옆이 아닌 다른 곳에 두려면:
```powershell
$env:MASTER_PATH = "D:\정산\원작료정산_마스터.xlsx"
```

### 3-2. (LibreOffice가 기본 경로가 아닐 때만) soffice 경로 지정
```powershell
$env:SOFFICE_PATH = "D:\LibreOffice\program\soffice.exe"
```
기본 경로(`C:\Program Files\LibreOffice\…`)면 설정 불필요.

### 3-3. (선택) 수식 오류 점검(재계산)
소스 테스트에서는 재계산 스크립트가 없으면 **자동으로 건너뜁니다**(검수 리포트의 "오류 건수"가 0으로 표시). 플로우 검증에는 영향 없습니다. 정산 금액/수식 정확성은 개발 단계에서 이미 검증되었습니다.

---

## 4. 테스트 절차

작업 디렉터리를 프로젝트 폴더로 옮긴 뒤 진행합니다:
```powershell
cd D:\정산\프로젝트
```

### ① 런처 실행
```powershell
python launcher.py
```
→ "원작료 정산 자동화" 시작 화면에 **[정산 실행] / [마스터 관리]** 두 버튼이 보이면 정상.

### ② 마스터 관리 (PySide6)
- 시작 화면에서 **[마스터 관리]** 클릭 → 별도 창(PySide6)이 뜹니다.
- 업체/작품/환율 등 마스터를 열람·편집·저장해 봅니다.
- 단독으로 바로 띄워 디버깅하려면:
  ```powershell
  python launcher.py --mode=master
  ```

### ③ 정산 실행 (tkinter)
- 시작 화면에서 **[정산 실행]** 클릭 → 정산 파이프라인 창이 뜹니다.
- 단독 실행:
  ```powershell
  python launcher.py --mode=pipeline
  ```
- 창에서:
  1. **RAW 폴더** = `inbox` 선택 (매출/기타수익/마스터 자동 인식 로그 확인)
  2. **정산서월** = 예) `2026-05` 입력, **회사** = 테라핀
  3. **[▶ 정산 실행]** 클릭

### ④ 검수 리포트 (게이트 1) → PDF 생성
- 정산서 생성이 끝나면 **검수 리포트 팝업**이 뜹니다:
  업체 수 · 정산서 수 · MG 작품 수 · 해외 정산 수 · 분기/반기 수 · 이월 업체 수 · **오류 건수** · 경고 건수
- 숫자를 확인하고 **[확인 — PDF 생성]** → `output\2026-05\테라핀\PDF\`에 PDF가 생성됩니다.
  - (오류 건수가 0이 아니면 강행 여부를 한 번 더 묻습니다)
  - PDF가 안 만들어지면 LibreOffice 설치/`SOFFICE_PATH`를 확인하세요.

### ⑤ 메일 Draft (게이트 2)
- PDF 생성 후 **"메일 Draft 생성?"** 최종 확인이 뜹니다.
- **[예]** → PDF에 비밀번호(사업자/주민번호) 설정 + `output\…\메일Draft\`에 `.eml` 초안 생성.
- **메일은 자동 발송되지 않습니다**(.eml 초안만). 더블클릭하면 메일 프로그램에서 열려 첨부·수신자를 확인할 수 있습니다.
- 비밀번호가 안 걸린 업체는 로그에 "비밀번호 미설정"으로 표시됩니다(마스터에 증빙번호 미등록).

---

## 5. 결과 확인 위치
```
output\2026-05\테라핀\
├─ (업체별 정산서).xlsx
├─ PDF\(업체별).pdf          ← ④에서 생성, ⑤에서 비밀번호 적용
├─ 메일Draft\(업체별).eml     ← ⑤에서 생성(무발송)
└─ 실행로그_….txt             ← 검수 지표 로그
```

---

## 6. 빠른 문제 진단

| 증상 | 원인 / 조치 |
|---|---|
| 시작 화면은 뜨는데 [정산 실행]을 눌러도 창이 안 뜸 | tkinter 미설치 → Python을 tcl/tk 옵션으로 재설치. `python launcher.py --mode=pipeline` 직접 실행해 에러 로그 확인 |
| [마스터 관리] 창이 안 뜸 | PySide6 미설치 → `pip install PySide6`. `python launcher.py --mode=master`로 에러 확인 |
| PDF가 생성되지 않음 | LibreOffice 미설치/경로 → 설치 후 `$env:SOFFICE_PATH` 지정 |
| 검수 리포트 "오류 건수"가 항상 0 | 정상(소스 테스트는 재계산 건너뜀). 수식 정확성은 개발 단계서 검증됨 |
| 한글 폴더/파일 깨짐 | PowerShell에서 `chcp 65001`로 UTF-8 설정 후 실행 |

---

## 7. 각 모드 단독 실행 요약 (디버깅용)
```powershell
python launcher.py                  # 런처(선택 화면)
python launcher.py --mode=pipeline  # 정산 실행 화면만
python launcher.py --mode=master    # 마스터 관리 화면만
```
런처를 거치지 않고 각 도구를 직접 띄울 수 있어, 문제 구간을 빠르게 좁힐 수 있습니다.

소스에서 ①~⑤가 모두 통과하면, 그때 `pyinstaller build.spec --noconfirm`(또는 GitHub Actions)으로 `정산자동화.exe`를 빌드하면 됩니다.
