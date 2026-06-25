# GitHub로 exe 빌드하기 (비개발자용)

윈도우 PC나 개발도구 없이, 깃허브가 윈도우에서 자동으로 exe를 만들어 줍니다.
빌드 과정이 파일(`build.yml`)로 남아 **담당자가 바뀌어도 누구나 다시 빌드**할 수 있습니다.

## ⚠️ 가장 중요 — 저장소는 반드시 "Private"
마스터에는 사업자번호·주민번호·이메일 같은 개인정보가 있습니다.
- 저장소는 **반드시 Private(비공개)** 로 만드세요.
- **실제 마스터(원작료정산_마스터.xlsx)와 매출 파일은 깃허브에 올리지 마세요.**
  (`.gitignore`가 자동으로 막아두긴 했지만, 직접 올리지 않도록 주의)
- exe를 받은 뒤, 실제 마스터·데이터는 **내 PC의 프로그램 폴더에만** 둡니다.

## 처음 한 번 설정
1. github.com 가입 → **New repository** → 이름 입력 → **Private** 선택 → Create.
2. 아래 파일들을 저장소에 올립니다(웹에서 드래그앤드롭 가능):
   `*.py`(코드 전부), `engine/`, `gui/`, `requirements.txt`, `build.spec`,
   `.github/workflows/build.yml`, `.gitignore`
   → 실제 마스터·매출 xlsx는 올리지 않습니다.
3. 올리면 **Actions 탭**에서 `build-exe`가 자동으로 돌기 시작합니다.
   (수동으로 돌리려면 Actions → build-exe → **Run workflow**)

## exe 받기
1. Actions 탭 → 방금 끝난 실행 클릭.
2. 아래 **Artifacts** 의 `royalty-settlement-windows` 다운로드(zip).
3. 압축을 풀면 `royalty_settlement/` 폴더 안에 `royalty_settlement.exe`가 있습니다.

## 실행 준비 (PC에서 한 번)
1. **LibreOffice 설치** — 정산서를 PDF로 바꾸는 데 필요합니다(무료).
2. 압축 푼 폴더 안에 **실제 `원작료정산_마스터.xlsx`** 를 넣습니다.
3. `inbox/매출리스트`, `inbox/기타수익`, `inbox/직전정산서` 에 그 달 파일을 넣습니다.
4. `royalty_settlement.exe` 실행.

## 코드가 바뀌면
파일만 다시 올리면(같은 저장소에 덮어쓰기) Actions가 **자동으로 새 exe를 다시 빌드**합니다.
