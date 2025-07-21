# 김준원표 주간통신문 아카이버

김준원표 만방학교 학부모페이지 아카이버

## ✨ 주요 기능

*   **자동화된 스크래핑:** Selenium WebDriver로 주간 통신문 페이지를를 자동으로 긁올수있음
*   **PDF 변환:** 긁어온 통신문이랑 이미지를 PDF로 변환해서 저장
*   **폴더 자동 정리:** 학기별, 주차별로 통신문이 자동으로 분류됨

## 🚀 시작하기

<p align="center">
  <a href="https://github.com/notj-code/WBHPA_archiever/archive/refs/heads/main.zip">
    <img src="https://img.shields.io/badge/Download%20for%20Windows-%20-blue?logo=windows11&logoColor=white&style=for-the-badge" alt="Windows (x86)">
  </a>
</p>



## 💡 사용법

1.  **앱 실행:** `gui_app.exe` (아니면 `python gui_app.py`) 실행하면 GUI 창이 뜰거임
2.  **로그인:**
    *   GUI에서 **"로그인하기"** 버튼 누르기
    *   새 크롬 창 뜨면 만방학교 학부모 페이지에 **직접 로그인**하기기
    *   로그인 완료되고 주간 통신문 페이지 뜨면 GUI에 "로그인 상태"가 "로그인 됨"으로 바뀜
3.  **아카이빙 시작:**
    *   로그인된 크롬 창에서 **아카이빙 원하는 년도/학기 직접 선택**하기
    *   GUI에서 **"시작"** 버튼 누르기
    *   앱이 선택한 학기 모든 주간 통신문 자동으로 스크래핑해서 PDF로 변환, `weekly_notices` 폴더에 저장함. 진행 상황은 GUI 하단 상태 바에
4.  **아카이빙 중지:**
    *   스크래핑 중에 멈추고 싶으면 GUI에서 **"정지"** 버튼 누르면 됨
5.  **저장된 폴더 열기:**
    *   아카이빙 끝나거나 중지한 뒤 **"저장된 폴더 열기"** 버튼 누르면 통신문 저장된 `weekly_notices` 폴더 열림

## ⚠️ 주의사항 및 문제 해결

*   **`chromedriver.exe` 오류:**
    *   `chromedriver.exe`는 사용 중인 Chrome 브라우저 버전과 일치해야 함
    *   Chrome 브라우저 업데이트했으면 `chromedriver.exe`도 최신으로 바꿔야 함
    *   혹시 오류가 나면 **본인 chrome 버전과 일치하는** `chromedriver.exe` 재설치 해볼 것
*   **웹사이트 변경:** 만방학교 학부모 페이지 HTML 구조 바뀌면 스크래핑 안될 수 있음. 이 경우 코드 수정 필요함.
