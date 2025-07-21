import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
import threading
import subprocess
import sys
from script import WeeklyNoticeScraper # script.py에서 WeeklyNoticeScraper 클래스 임포트

class Application(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("김준원표 주간 통신문 아카이버")
        self.geometry("500x400")
        self.resizable(False, False)

        self.scraper = WeeklyNoticeScraper() # Scraper 인스턴스 생성
        
        self.create_widgets()
        
        self.scraping_thread = None
        self.stop_scraping_flag = threading.Event()

        # GUI 종료 시 드라이버 종료
        self.protocol("WM_DELETE_WINDOW", self.on_closing)

    def create_widgets(self):
        # 로그인 섹션
        login_frame = ttk.LabelFrame(self, text="로그인", padding="10")
        login_frame.pack(pady=10, padx=10, fill="x")

        self.login_button = ttk.Button(login_frame, text="로그인하기", command=self.start_login_process)
        self.login_button.pack(pady=5)

        self.login_status_label = ttk.Label(login_frame, text="로그인 상태: 로그아웃", foreground="red")
        self.login_status_label.pack(pady=5)

        # 아카이빙 섹션
        scrape_frame = ttk.LabelFrame(self, text="아카이빙", padding="10")
        scrape_frame.pack(pady=10, padx=10, fill="x")

        ttk.Label(scrape_frame, text="브라우저에서 학기를 선택한 후 시작 버튼을 눌러주세요:").pack(pady=5)

        self.start_button = ttk.Button(scrape_frame, text="시작", command=self.start_scraping, state="disabled") # 초기에는 비활성화
        self.start_button.pack(side="left", padx=5, expand=True)

        self.stop_button = ttk.Button(scrape_frame, text="정지", command=self.stop_scraping, state="disabled")
        self.stop_button.pack(side="right", padx=5, expand=True)

        # 폴더 열기 섹션
        folder_frame = ttk.LabelFrame(self, text="저장된 폴더", padding="10")
        folder_frame.pack(pady=10, padx=10, fill="x")

        self.open_folder_button = ttk.Button(folder_frame, text="저장된 폴더 열기", command=self.open_save_folder)
        self.open_folder_button.pack(pady=5)

        # 상태 메시지 영역
        self.status_text = tk.Text(self, height=5, state="disabled", wrap="word")
        self.status_text.pack(pady=10, padx=10, fill="both", expand=True)

    def log_message(self, message):
        self.status_text.config(state="normal")
        self.status_text.insert(tk.END, message + "\n")
        self.status_text.see(tk.END)
        self.status_text.config(state="disabled")
        self.update_idletasks() # GUI 업데이트 강제

    def start_login_process(self):
        self.log_message("로그인 프로세스 시작...")
        self.login_button.config(state="disabled")
        self.login_status_label.config(text="로그인 상태: 브라우저 대기 중...", foreground="orange")
        
        # 별도 스레드에서 로그인 로직 실행
        threading.Thread(target=self._run_login).start()

    def _run_login(self):
        # headless=False로 설정하여 브라우저 창을 띄웁니다.
        login_success = self.scraper.login(progress_callback=self.log_message)
        
        if login_success:
            self.after(0, self.on_login_success) # 메인 스레드에서 GUI 업데이트
        else:
            self.after(0, self.on_login_failure) # 메인 스레드에서 GUI 업데이트

    def on_login_success(self):
        self.log_message("로그인 성공!")
        self.login_status_label.config(text="로그인 상태: 로그인 됨", foreground="green")
        self.login_button.config(text="로그인 됨", state="disabled")
        self.start_button.config(state="normal") # 로그인 성공 시 시작 버튼 활성화

    def on_login_failure(self):
        self.log_message("로그인 실패 또는 페이지 로드 오류 발생.")
        self.login_status_label.config(text="로그인 상태: 로그인 실패", foreground="red")
        self.login_button.config(state="normal") # 로그인 실패 시 버튼 다시 활성화
        self.scraper.quit_driver() # 실패 시 드라이버 종료

    def start_scraping(self):
        self.log_message(f"아카이빙 시작...")
        self.start_button.config(state="disabled")
        self.stop_button.config(state="normal")
        
        self.stop_scraping_flag.clear() # 중지 플래그 초기화
        self.scraping_thread = threading.Thread(target=self._run_scraping, 
                                                args=(self.stop_scraping_flag,))
        self.scraping_thread.start()

    def _run_scraping(self, stop_event):
        scraping_success = self.scraper.scrape_selected_semester_from_browser(
            progress_callback=self.log_message, 
            stop_event=stop_event
        )
        
        if scraping_success and not stop_event.is_set():
            self.after(0, self.on_scraping_complete)
        else:
            self.after(0, self.on_scraping_stopped)

    def stop_scraping(self):
        self.log_message("스크래핑 중지 요청...")
        self.stop_scraping_flag.set() # 중지 플래그 설정
        self.stop_button.config(state="disabled") # 중지 버튼 비활성화

    def on_scraping_complete(self):
        self.log_message("아카이빙 완료!")
        self.start_button.config(state="normal")
        self.stop_button.config(state="disabled")

    def on_scraping_stopped(self):
        self.log_message("아카이빙 중지됨.")
        self.start_button.config(state="normal")
        self.stop_button.config(state="disabled")

    def open_save_folder(self):
        save_dir = self.scraper.save_dir # Scraper 인스턴스에서 SAVE_DIR 가져오기
        if not os.path.exists(save_dir):
            os.makedirs(save_dir)
            self.log_message(f"'{save_dir}' 폴더가 존재하지 않아 생성했습니다.")

        try:
            if sys.platform == "win32":
                os.startfile(os.path.abspath(save_dir))
            elif sys.platform == "darwin": # macOS
                subprocess.Popen(["open", os.path.abspath(save_dir)])
            else: # linux variants
                subprocess.Popen(["xdg-open", os.path.abspath(save_dir)])
            self.log_message(f"'{save_dir}' 폴더를 열었습니다.")
        except Exception as e:
            self.log_message(f"폴더 열기 실패: {e}")
            messagebox.showerror("오류", f"폴더를 여는 데 실패했습니다: {e}")

    def on_closing(self):
        self.log_message("애플리케이션 종료 중...")
        self.scraper.quit_driver()
        self.destroy()

if __name__ == "__main__":
    app = Application()
    app.mainloop()
