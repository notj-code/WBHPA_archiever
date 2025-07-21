import os
import time
import base64
import threading
import sys
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select

# --- 설정 (필요에 따라 수정) ---
BASE_URL = "https://wb.manbangschool.org/_Parent/WeeklyReport/WeekReportView.asp"
SAVE_DIR = "weekly_notices"

# PyInstaller로 패키징될 때 chromedriver.exe의 경로를 동적으로 설정
def get_webdriver_path():
    if getattr(sys, 'frozen', False): # PyInstaller로 실행될 때
        return os.path.join(os.path.dirname(sys.executable), "chromedriver.exe")
    else: # 개발 환경에서 실행될 때
        return "chromedriver.exe"

# --- Selenium을 사용하여 현재 페이지를 PDF로 저장하는 함수 ---
def save_current_page_as_pdf(driver, pdf_file_path, progress_callback=None):
    """Saves the current browser page to a PDF file with enhanced stability."""
    try:
        if progress_callback: progress_callback(f"  > PDF 저장을 위해 페이지 안정화 대기...")
        
        # 페이지의 핵심 콘텐츠가 완전히 로드될 때까지 명시적으로 대기
        WebDriverWait(driver, 10).until(
            EC.visibility_of_element_located((By.CSS_SELECTOR, ".weekly_contents"))
        )
        time.sleep(1) # 렌더링을 위한 추가적인 여유 시간

        if progress_callback: progress_callback(f"  > 현재 페이지를 PDF로 저장 시도: {pdf_file_path}")
        
        # A4 사이즈 및 여백을 명시한 PDF 인쇄 옵션
        print_options = {
            'landscape': False,
            'displayHeaderFooter': False,
            'printBackground': True,
            'preferCSSPageSize': True,
            'paperWidth': 8.27,  # A4 width in inches
            'paperHeight': 11.69, # A4 height in inches
            'marginTop': 0.4,      # Margin in inches
            'marginBottom': 0.4,
            'marginLeft': 0.4,
            'marginRight': 0.4
        }

        # CDP(Chrome DevTools Protocol)를 직접 사용하여 PDF로 인쇄
        result = driver.execute_cdp_cmd("Page.printToPDF", print_options)
        pdf_bytes = base64.b64decode(result['data'])

        with open(pdf_file_path, "wb") as f:
            f.write(pdf_bytes)
        
        if progress_callback: progress_callback(f"  > PDF 파일 저장 완료: {pdf_file_path}")
        return True
    except Exception as e:
        if progress_callback: progress_callback(f"  > PDF 저장 중 오류 발생: {e}")
        return False

class WeeklyNoticeScraper:
    def __init__(self, base_url=BASE_URL, save_dir=SAVE_DIR):
        self.webdriver_path = get_webdriver_path()
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
            self.wait = WebDriverWait(self.driver, 600) # 로그인 대기 시간을 10분(600초)으로 늘림
            return True
        except Exception as e:
            if progress_callback: progress_callback(f"WebDriver 초기화 오류: {e}")
            return False

    def login(self, progress_callback=None):
        if not self.init_driver(headless=False, progress_callback=progress_callback):
            return False

        try:
            self.driver.get(self.base_url)
            if progress_callback:
                progress_callback("브라우저가 열렸습니다. 로그인 페이지로 이동 중...")
                progress_callback("브라우저에서 직접 로그인하고, 아카이빙을 원하는 학기를 선택해주세요.")
            
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
            year_select_element = self.wait.until(EC.presence_of_element_located((By.ID, "Left_weekYear")))
            year_select = Select(year_select_element)
            selected_year_str = year_select.first_selected_option.text
            
            semester_select_element = self.wait.until(EC.presence_of_element_located((By.ID, "Left_weekTerm")))
            semester_select = Select(semester_select_element)
            selected_semester_name = semester_select.first_selected_option.text

            if progress_callback:
                progress_callback(f"\n--- 현재 브라우저에서 선택된 {selected_year_str} {selected_semester_name} 학기 통신문 스크래핑 시작 ---")

            self.wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, ".leftweeklytext01")))
            time.sleep(1)

            weekly_notice_elements = self.driver.find_elements(By.CSS_SELECTOR, ".leftweeklytext01 a")
            weekly_notices_info = []
            for elem in weekly_notice_elements:
                text = elem.text
                href = elem.get_attribute("href")
                if text and href and ("주차" in text or "[읽음]" in text):
                    weekly_notices_info.append((text, href))

            if not weekly_notices_info:
                if progress_callback: progress_callback(f"WARNING: {selected_year_str} {selected_semester_name} 학기에 스크래핑할 주차 통신문이 현재 화면에 없습니다.")
                return False
            
            def sort_key_func(item):
                try:
                    # "n주차" 형식의 텍스트에서 숫자 n을 추출합니다.
                    # "08주차" 와 같이 0으로 시작하는 경우도 처리합니다.
                    week_text = item[0].split("주차")[0].split(" ")[-1].strip()
                    return int(week_text)
                except (ValueError, IndexError):
                    # 변환에 실패하면 정렬 순서를 낮게 유지합니다.
                    return -1
            weekly_notices_info.sort(key=sort_key_func)

            # --- 주차별 통신문 스크래핑 ---
            for i, (notice_text, notice_href) in enumerate(weekly_notices_info):
                if self.stop_event.is_set():
                    if progress_callback: progress_callback("스크래핑 중지 요청 감지.")
                    break

                if progress_callback: progress_callback(f"\n--- ({i+1}/{len(weekly_notices_info)}) {notice_text} 스크래핑 시작 ---")
                
                self.driver.get(notice_href)
                
                clean_notice_text = notice_text.replace(' ', '_').replace('/', '_').replace(':', '').replace('[읽음]', '').strip()
                folder_name = f"{selected_year_str}_{selected_semester_name}_{clean_notice_text}"
                current_notice_folder = os.path.join(self.save_dir, folder_name)
                os.makedirs(current_notice_folder, exist_ok=True)

                page_num = 1
                while not self.stop_event.is_set():
                    self.wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, ".weekly_contents")))
                    time.sleep(1)

                    if progress_callback: progress_callback(f"  > {notice_text} - 페이지 {page_num} 저장 중...")
                    
                    pdf_output_path = os.path.join(current_notice_folder, f"{clean_notice_text}_page{page_num}.pdf")
                    save_current_page_as_pdf(self.driver, pdf_output_path, progress_callback)

                    # 현재 주차의 마지막 페이지인지 확인
                    try:
                        self.driver.find_element(By.ID, "P_LAST_WEEKREPORT_READED_YN")
                        if progress_callback: progress_callback(f"  > 마지막 페이지 확인됨. {notice_text} 스크래핑을 종료합니다.")
                        break 
                    except:
                        # 마지막 페이지가 아니면 다음 버튼 클릭
                        try:
                            next_button = self.driver.find_element(By.CSS_SELECTOR, "a.next")
                            if progress_callback: progress_callback("  > '다음' 페이지 버튼 발견. 클릭합니다.")
                            self.driver.execute_script("arguments[0].click();", next_button)
                            time.sleep(2)
                            page_num += 1
                        except:
                            if progress_callback: progress_callback(f"  > '다음' 페이지 버튼을 찾을 수 없어 {notice_text} 스크래핑을 종료합니다.")
                            break

                if progress_callback: progress_callback(f"--- {notice_text} 스크래핑 완료 ---")
            
            if progress_callback: progress_callback(f"\n--- {selected_year_str} {selected_semester_name} 학기 전체 스크래핑 완료 ---")
            
            # 스크래핑 완료 후 사용자가 다른 학기를 선택할 수 있도록 기본 페이지로 돌아갑니다.
            self.driver.get(self.base_url)
            self.wait.until(EC.presence_of_element_located((By.ID, "Left_weekYear")))
            return True

        except Exception as e:
            if progress_callback: progress_callback(f"전체 스크래핑 중 예상치 못한 오류 발생: {e}")
            return False
        finally:
            # GUI에서 명시적으로 드라이버를 종료하므로 여기서는 종료하지 않습니다.
            pass

    def quit_driver(self):
        if self.driver:
            self.driver.quit()
            self.driver = None
            self.wait = None

# GUI에서 직접 클래스를 사용하므로, 아래의 main 블록은 필요 없습니다.
# if __name__ == "__main__":
#     pass
