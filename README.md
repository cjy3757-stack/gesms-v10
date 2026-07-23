# GunpoEunhye Corps Family Management System V8.5.1

## 출석 통계 자동화

V8.5.1부터 `attendance-index.json`을 직접 작성하거나 수정하지 않습니다.

1. 앱에서 날짜와 예배 종류를 선택해 출석을 체크합니다.
2. 날짜별 CSV를 저장합니다.
3. 해당 CSV만 GitHub 저장소에 업로드합니다.
4. 통계 화면에서 **전체 통계 새로고침** 또는 **GitHub CSV 자동검색**을 누릅니다.

파일명은 앱이 자동 생성하는 `attendance-YYYY-MM-DD-예배명.csv` 형식을 그대로 유지해야 합니다. 출석 CSV는 저장소 루트 또는 하위 폴더에 둘 수 있습니다.

GitHub API 자동검색이 일시적으로 실패할 경우 기존 `attendance-index.json`과 기기 저장 목록을 예비 수단으로 사용합니다.
