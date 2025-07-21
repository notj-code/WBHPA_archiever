from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
import os
import time
import requests
import base64 # 이미지 base64 인코딩을 위해 추가
import threading
import sys # sys 모듈 임포트
import tempfile # 임시 파일 생성을 위해 추가

# --- 설정 (필요에 따라 수정) ---
BASE_URL = "https://wb.manbangschool.org/_Parent/WeeklyReport/WeekReportView.asp"
SAVE_DIR = "weekly_notices"

# PyInstaller로 패키징될 때 chromedriver.exe의 경로를 동적으로 설정
def get_webdriver_path():
    if getattr(sys, 'frozen', False): # PyInstaller로 실행될 때
        # 실행 파일과 같은 디렉토리에 chromedriver.exe가 있다고 가정
        return os.path.join(os.path.dirname(sys.executable), "chromedriver.exe")
    else: # 개발 환경에서 실행될 때
        return "chromedriver.exe"

# --- 이미지 다운로드 함수 ---
def download_image(image_url, folder_path, image_name):
    try:
        if not image_url.startswith("http"):
            image_url = urljoin(BASE_URL, image_url)
            
        response = requests.get(image_url, stream=True)
        response.raise_for_status()
        file_path = os.path.join(folder_path, image_name)
        with open(file_path, 'wb') as f:
            for chunk in response.iter_content(chunk_size=8192):
                f.write(chunk)
        # print(f"  > 이미지 다운로드 성공: {image_name}") # GUI로 메시지 전달
        return file_path
    except requests.exceptions.RequestException as e:
        # print(f"  > 이미지 다운로드 실패: {image_url} - {e}") # GUI로 메시지 전달
        return None
    except Exception as e:
        # print(f"  > 이미지 처리 중 오류 발생: {e}") # GUI로 메시지 전달
        return None

# --- Selenium을 사용하여 HTML 내용을 PDF로 인쇄하는 함수 ---
def print_html_to_pdf_with_selenium(driver, title, page_text, page_images, pdf_file_path, progress_callback=None):
    temp_html_file = None
    try:
        # 모든 텍스트와 이미지를 포함하는 단일 HTML 파일 생성
        html_content = f"""
<!DOCTYPE html>
<html>
<head>
    <meta charset="utf-8">
    <title>{title}</title>
    <style>
        body {{ font-family: 'Noto Sans KR', 'Malgun Gothic', '맑은 고딕', sans-serif; line-height: 1.6; margin: 20mm; }}
        h1 {{ text-align: center; margin-bottom: 20px; }}
        img {{ max-width: 100%; height: auto; display: block; margin: 10px auto; border: 1px solid #eee; }}
        .page-break {{ page-break-before: always; }}
    </style>
</head>
<body>
    <h1>{title}</h1>
    <div class="content">
"""

        # 텍스트 내용 추가
        html_content += f"<pre style=\"white-space: pre-wrap; word-wrap: break-word;\">{page_text}</pre>"

        # 이미지 추가 (base64 인코딩 또는 로컬 경로 사용)
        for img_path in page_images:
            if img_path and os.path.exists(img_path):
                try:
                    if progress_callback: progress_callback(f"  > HTML에 이미지 삽입 시도: {img_path}")
                    # 이미지를 base64로 인코딩하여 HTML에 직접 삽입
                    with open(img_path, "rb") as image_file:
                        encoded_string = base64.b64encode(image_file.read()).decode('utf-8')
                    html_content += f'<img src="data:image/jpeg;base64,{encoded_string}" alt="Image"/>'
                    if progress_callback: progress_callback(f"  > HTML에 이미지 삽입 성공: {img_path}")
                except Exception as e:
                    if progress_callback: progress_callback(f"  > HTML에 이미지 삽입 실패: {img_path} - {e}")
            else:
                if progress_callback: progress_callback(f"  > 유효하지 않은 이미지 경로 또는 파일 없음: {img_path}")

        html_content += f"""
    </div>
</body>
</html>
"""

        # 임시 HTML 파일 생성
        fd, temp_html_file = tempfile.mkstemp(suffix=".html", prefix="weekly_notice_")
        with os.fdopen(fd, 'w', encoding='utf-8') as f:
            f.write(html_content)
        
        if progress_callback: progress_callback(f"  > 임시 HTML 파일 생성: {temp_html_file}")

        # Selenium 드라이버로 임시 HTML 파일 로드
        driver.get(f"file:///{os.path.abspath(temp_html_file)}")
        time.sleep(2) # 페이지 로드 대기

        # PDF 인쇄 옵션 (필요에 따라 조정)
        print_options = {
            'landscape': False,
            'displayHeaderFooter': False,
            'printBackground': True,
            'preferCSSPageSize': True,
        }

        # PDF로 인쇄
        pdf_base64 = driver.print_page(print_options)
        pdf_bytes = base64.b64decode(pdf_base64)

        with open(pdf_file_path, "wb") as f:
            f.write(pdf_bytes)
        
        if progress_callback: progress_callback(f"PDF 파일 생성 완료: {pdf_file_path}")

    except Exception as e:
        if progress_callback: progress_callback(f"PDF 변환 중 오류 발생: {e}")
    finally:
        if temp_html_file and os.path.exists(temp_html_file):
            os.remove(temp_html_file)
            if progress_callback: progress_callback(f"  > 임시 HTML 파일 삭제: {temp_html_file}")

class WeeklyNoticeScraper:
    def __init__(self, base_url=BASE_URL, save_dir=SAVE_DIR):
        self.webdriver_path = get_webdriver_path() # 동적으로 경로 가져오기
        self.base_url = base_url
        self.save_dir = save_dir
        self.driver = None
        self.wait = None
        self.stop_event = threading.Event()

        if not os.path.exists(self.save_dir):
            os.makedirs(self.save_dir)

    def init_driver(self, headless=False, progress_callback=None):
        if progress_callback: progress_callback(f"WebDriver 경로: {self.webdriver_path}")
        if not os.path.exists(self.webdriver_path):
            if progress_callback: progress_callback(f"오류: chromedriver.exe를 찾을 수 없습니다. 경로: {self.webdriver_path}")
            return False

        options = webdriver.ChromeOptions()
        if headless:
            options.add_argument("--headless")
        options.add_argument("--no-sandbox")
        options.add_argument("--disable-dev-shm-usage")
        options.add_experimental_option("prefs", {
            "download.default_directory": os.path.abspath(self.save_dir),
            "download.prompt_for_download": False,
            "download.directory_upgrade": True,
            "plugins.always_open_pdf_externally": True
        })
        try:
            service = webdriver.chrome.service.Service(self.webdriver_path)
            self.driver = webdriver.Chrome(service=service, options=options)
            self.wait = WebDriverWait(self.driver, 60)
            return True
        except Exception as e:
            if progress_callback: progress_callback(f"WebDriver 초기화 오류: {e}")
            return False

    def login(self, progress_callback=None):
        if not self.init_driver(headless=False, progress_callback=progress_callback): # 드라이버 초기화 실패 시
            return False

        try:
            self.driver.get(self.base_url)
            if progress_callback:
                progress_callback("브라우저가 열렸습니다. 로그인 페이지로 이동 중...")
                progress_callback("브라우저에서 직접 로그인하고, 아카이빙을 원하는 학기를 선택해주세요.")
            
            # 로그인 성공 여부를 판단하는 로직 (예: 특정 요소가 나타날 때까지 대기)
            # 여기서는 로그인 후 주간 통신문 페이지의 특정 요소가 나타날 때까지 기다립니다.
            self.wait.until(EC.presence_of_element_located((By.ID, "Left_weekYear")))
            if progress_callback:
                progress_callback("로그인 성공 및 주간 통신문 페이지 로드 확인.")
            return True
        except Exception as e:
            if progress_callback:
                progress_callback(f"로그인 실패 또는 페이지 로드 오류: {e}")
            return False

    def scrape_selected_semester_from_browser(self, progress_callback=None, stop_event=None):
        self.stop_event = stop_event if stop_event else threading.Event()

        if not self.driver:
            if progress_callback: progress_callback("드라이버가 초기화되지 않았습니다.")
            return False

        try:
            # 현재 브라우저에서 선택된 년도와 학기 값 가져오기
            year_select_element = self.wait.until(EC.presence_of_element_located((By.ID, "Left_weekYear")))
            year_select = Select(year_select_element)
            selected_year_str = year_select.first_selected_option.text # 표시용 텍스트
            selected_year_value = year_select.first_selected_option.get_attribute("value") # 실제 값

            semester_select_element = self.wait.until(EC.presence_of_element_located((By.ID, "Left_weekTerm")))
            semester_select = Select(semester_select_element)
            selected_semester_name = semester_select.first_selected_option.text # 표시용 텍스트
            selected_semester_value = semester_select.first_selected_option.get_attribute("value") # 실제 값

            if progress_callback:
                progress_callback(f"\n--- 현재 브라우저에서 선택된 {selected_year_str} {selected_semester_name} 학기 통신문 스크래핑 시작 ---")

            # 주간 통신문 목록이 로드될 때까지 명시적으로 대기
            self.wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, ".leftweeklytext01")))
            if progress_callback:
                progress_callback(f"{selected_year_str} {selected_semester_name} 학기의 주간 통신문 목록 로딩 완료.")
            time.sleep(1) # 추가적인 안정성 확보를 위한 짧은 대기

            # 주차별 링크 가져오기 (사이드바)
            weekly_notices_info = []
            weekly_notice_elements = self.driver.find_elements(By.CSS_SELECTOR, ".leftweeklytext01 a")
            
            for elem in weekly_notice_elements:
                text = elem.text
                href = elem.get_attribute("href")
                
                if text and href and ("주차" in text or "[읽음]" in text):
                    try:
                        week_num_str_parts = text.split("주차")[0].split(" ")
                        week_num = int(week_num_str_parts[-1].strip()) # 주차 번호 추출
                        weekly_notices_info.append((text, href))
                    except ValueError:
                        if progress_callback: progress_callback(f"DEBUG: 주차 번호로 변환할 수 없는 텍스트 감지 (스킵): '{text}'")
                        pass
                else:
                    if progress_callback: progress_callback(f"DEBUG: '주차' 또는 '[읽음]' 텍스트를 포함하지 않는 링크 스킵: '{text}'")

            if not weekly_notices_info:
                if progress_callback: progress_callback(f"WARNING: {selected_year_str} {selected_semester_name} 학기에 스크래핑할 주차 통신문이 현재 화면에 없습니다.")
                return False
            
            # 주차별로 정렬 (0주차, 1주차, ... 순서로)
            def sort_key_func(item):
                try:
                    return int(item[0].split("주차")[0].split(" ")[-1].strip())
                except ValueError:
                    return -1

            weekly_notices_info.sort(key=sort_key_func)
            if progress_callback: progress_callback(f"DEBUG: {selected_year_str} {selected_semester_name} 학기 주차 정보 정렬 완료. 스크래핑할 주차: {[item[0] for item in weekly_notices_info]}")

            # --- 주차별 통신문 스크래핑 ---
            for i, (notice_text, notice_href) in enumerate(weekly_notices_info):
                if self.stop_event.is_set():
                    if progress_callback: progress_callback("스크래핑 중지 요청 감지. 현재 주차 스크래핑을 중단합니다.")
                    break

                if progress_callback: progress_callback(f"\n--- ({i+1}/{len(weekly_notices_info)}) {notice_text} 스크래핑 시작 ---")
                
                self.driver.get(notice_href)
                time.sleep(2);

                # 파일명 및 폴더명에 불필요한 문자 제거
                clean_notice_text = notice_text.replace(' ', '_').replace('/', '_').replace(':', '').replace('[읽음]', '').strip()
                folder_name = f"{selected_year_str}_{selected_semester_name}_{clean_notice_text}"
                current_notice_folder = os.path.join(self.save_dir, folder_name)
                os.makedirs(current_notice_folder, exist_ok=True)

                page_num = 1

                while True:
                    if self.stop_event.is_set():
                        if progress_callback: progress_callback("스크래핑 중지 요청 감지. 현재 페이지 스크래핑을 중단합니다.")
                        break

                    if progress_callback: progress_callback(f"  > {notice_text} - 페이지 {page_num} 추출 중...")
                    
                    try:
                        title_element = self.wait.until(EC.presence_of_element_located((By.ID, "contenttitle")))
                        title = title_element.text.strip()
                    except:
                        title = f"{notice_text} - No Title"
                        if progress_callback: progress_callback("  > 제목 추출 실패. 기본 제목 사용.")

                    page_text = ""
                    try:
                        content_element = self.wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, ".weekly_contents")))
                        page_text = content_element.text.strip()
                    except Exception as e:
                        if progress_callback: progress_callback(f"  > 본문 텍스트 추출 실패: {e}")
                    
                    page_images = []
                    try:
                        content_element = self.wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, ".weekly_contents")))
                        img_elements = content_element.find_elements(By.TAG_NAME, "img")
                        if progress_callback: progress_callback(f"  > 페이지에서 {len(img_elements)}개의 이미지 요소 감지.")
                        for j, img in enumerate(img_elements):
                            img_url = img.get_attribute("src")
                            if img_url:
                                if progress_callback: progress_callback(f"  > 이미지 URL: {img_url}")
                                img_name = f"{clean_notice_text}_page{page_num}_img{j+1}.jpg"
                                downloaded_path = download_image(img_url, current_notice_folder, img_name)
                                if downloaded_path:
                                    page_images.append(downloaded_path)
                                    if progress_callback: progress_callback(f"  > 이미지 다운로드 성공: {downloaded_path}")
                                else:
                                    if progress_callback: progress_callback(f"  > 이미지 다운로드 실패: {img_url}")
                            else:
                                if progress_callback: progress_callback(f"  > 이미지 src 속성 없음: {img}")
                    except Exception as e:
                        if progress_callback: progress_callback(f"  > 이미지 추출/다운로드 실패: {e}")
                    
                    html_content = self.driver.page_source
                    html_file_path = os.path.join(current_notice_folder, f"{clean_notice_text}_page{page_num}.html")
                    with open(html_file_path, "w", encoding="utf-8") as f:
                        f.write(html_content)
                    if progress_callback: progress_callback(f"  > HTML 파일 저장 완료: {html_file_path}")

                    txt_file_path = os.path.join(current_notice_folder, f"{clean_notice_text}_page{page_num}.txt")
                    with open(txt_file_path, "w", encoding="utf-8") as f:
                        f.write(f"제목: {title}\n\n{page_text}")
                    if progress_callback: progress_callback(f"  > 텍스트 파일 저장 완료: {txt_file_path}")
                    
                    # PDF를 매 페이지마다 저장
                    pdf_output_path = os.path.join(current_notice_folder, f"{clean_notice_text}_page{page_num}.pdf")
                    print_html_to_pdf_with_selenium(self.driver, title, page_text, page_images, pdf_output_path, progress_callback)

                    # 주간 통신문 끝을 나타내는 요소 확인
                    end_of_notice_found = False
                    try:
                        # id가 P_LAST_WEEKREPORT_READED_YN인 input 태그가 존재하는지 확인
                        self.driver.find_element(By.ID, "P_LAST_WEEKREPORT_READED_YN")
                        end_of_notice_found = True
                        if progress_callback: progress_callback("  > 주간 통신문 끝 요소 감지. 현재 주차 스크래핑을 종료합니다.")
                    except:
                        pass # 요소가 없으면 계속 진행

                    next_button_found = False
                    if not end_of_notice_found: # 끝 요소가 발견되지 않았을 때만 다음 페이지 버튼 확인
                        try:
                            next_button = self.driver.find_element(By.CSS_SELECTOR, ".pagination .next")
                            if next_button.is_displayed() and next_button.is_enabled():
                                next_page_href = next_button.get_attribute("href")
                                if next_page_href:
                                    self.driver.get(next_page_href)
                                    time.sleep(2);
                                    page_num += 1
                                    next_button_found = True
                        except Exception as e:
                            pass

                    if not next_button_found or end_of_notice_found: # 다음 페이지 버튼이 없거나 끝 요소가 발견되면 페이지 루프 종료
                        break

                if progress_callback: progress_callback(f"--- {notice_text} 스크래핑 완료 ---")
                
                # 스크래핑 완료 후 다시 기본 페이지로 돌아가 다음 수동 선택을 기다림
                self.driver.get(self.base_url)
                # 드롭다운 엘리먼트가 다시 로드될 때까지 대기
                self.wait.until(EC.presence_of_element_located((By.ID, "Left_weekYear")))
                # 드라이버가 페이지를 다시 로드했으므로 Select 객체도 다시 초기화해야 함
                year_select_element = self.driver.find_element(By.ID, "Left_weekYear")
                year_select = Select(year_select_element)
                semester_select_element = self.driver.find_element(By.ID, "Left_weekTerm")
                semester_select = Select(semester_select_element)

            return True

        except Exception as e:
            if progress_callback: progress_callback(f"전체 스크래핑 중 예상치 못한 오류 발생: {e}")
            return False
        finally:
            # 드라이버 종료는 GUI에서 명시적으로 호출하도록 변경
            pass

    def quit_driver(self):
        if self.driver:
            self.driver.quit()
            self.driver = None
            self.wait = None
            # print("웹 드라이버 종료.") # GUI로 메시지 전달

# if __name__ == "__main__":
#     # 이 부분은 GUI에서 호출되므로 주석 처리하거나 삭제합니다.
#     # scraper = WeeklyNoticeScraper()
#     # scraper.init_driver()
#     # scraper.login()
#     # semesters = scraper.get_available_semesters()
#     # print(semesters)
#     # scraper.scrape_semester("2024", "1") # 예시
#     # scraper.quit_driver()