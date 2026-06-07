import customtkinter as ctk
import threading
import time
import os
import logging
import openpyxl
import subprocess
import requests
from datetime import datetime
from tkinter import messagebox
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager
import json

# 텔레그램 설정
TOKEN = "8774609207:AAEGTPiEXkGHktvFoDXod0BSLpzGseRoo5E"
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
XL_PATH = os.path.join(BASE_DIR, "정리본.xlsx")
IDS_PATH = os.path.join(BASE_DIR, "chat_ids.txt")
CREDENTIALS_PATH = os.path.join(BASE_DIR, "spath_credentials.json")

class NoticeSearchApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("SSU 공지사항 자동 검색기")
        self.geometry("750x680")
        self.resizable(True, True)
        self.protocol("WM_DELETE_WINDOW", self.on_closing)
        
        # 최상단 고정 (초기 실행 시)
        self.attributes("-topmost", True)
        self.focus_force()
        self.after(500, lambda: self.attributes("-topmost", False))
        
        # 메인 프레임
        self.main_frame = ctk.CTkFrame(self)
        self.main_frame.pack(pady=20, padx=20, fill="both", expand=True)
        
        ctk.CTkLabel(self.main_frame, text="🔍 공지사항 실시간 모니터링", font=ctk.CTkFont(size=24, weight="bold")).pack(pady=(20, 5))
        ctk.CTkLabel(self.main_frame, text="입력한 키워드를 실시간으로 모니터링하여 '정리본.xlsx'에 저장하고 텔레그램으로 알려드립니다.", 
                     font=ctk.CTkFont(size=12), text_color="gray").pack(pady=(0, 20))
        
        # 입력 필드 프레임
        self.input_frame = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        self.input_frame.pack(pady=5)
        
        self.entry_kw = self._create_input("검색 키워드 (띄어쓰기 구분):", 0)
        self.entry_page = self._create_input("검색할 페이지 수 (예: 1 5, 기본 1 1):", 1, "1 1")
        
        ctk.CTkLabel(self.input_frame, text="검색 주기 설정:", font=ctk.CTkFont(size=14)).grid(row=2, column=0, padx=10, pady=10, sticky="e")
        self.combo_interval = ctk.CTkOptionMenu(self.input_frame, 
                                                values=["1분", "5분", "10분", "30분", "1시간", "2시간", "6시간", "12시간"],
                                                width=280, font=("Malgun Gothic", 13))
        self.combo_interval.set("1분")
        self.combo_interval.grid(row=2, column=1, padx=5, pady=10)
        
        # 제어 버튼 프레임
        self.btn_frame = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        self.btn_frame.pack(pady=10)
        
        self.btn_start = ctk.CTkButton(self.btn_frame, text="🔍 검색 시작", command=self.start_search, 
                                       width=140, height=45, font=("Malgun Gothic", 15, "bold"))
        self.btn_start.pack(side="left", padx=10)
        
        self.btn_stop = ctk.CTkButton(self.btn_frame, text="🛑 검색 중지", command=self.stop_search, 
                                      width=140, height=45, fg_color="#E74C3C", font=("Malgun Gothic", 15, "bold"), 
                                      state="disabled")
        self.btn_stop.pack(side="left", padx=10)
        
        # 유틸리티 버튼 프레임
        self.util_frame = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        self.util_frame.pack(pady=5)
        
        util_opts = [
            ("📂 결과 폴더 열기", "#34495E", lambda: os.startfile(BASE_DIR)),
            ("📊 엑셀 파일 보기", "#27AE60", self.open_excel),
            ("🗑️ 엑셀 데이터 초기화", "#D35400", self.reset_excel)
        ]
        
        for text, color, cmd in util_opts:
            ctk.CTkButton(self.util_frame, text=text, command=cmd, fg_color=color, width=140, height=35).pack(side="left", padx=10)
            
        # 로그 영역
        self.log_area = ctk.CTkTextbox(self.main_frame, width=650, height=220, font=("Malgun Gothic", 13))
        self.log_area.pack(pady=(10, 20), padx=20, fill="both", expand=True)
        self.log_area.configure(state="disabled")
        
        self.last_update_id = 0
        self.chat_ids = self.load_chat_ids()
        
        # 동시 검색을 위한 상태 관리
        self.excel_lock = threading.Lock()
        self.search_status = {"GUI": False}
        self.user_states = {} # 텔레그램 사용자별 상태 저장 (예: 검색어 입력 대기)
        
        # 텔레그램 대기 세션
        self.session = requests.Session()
        
        # 텔레그램 폴링 스레드 시작
        threading.Thread(target=self.start_telegram_polling, daemon=True).start()

    def load_chat_ids(self):
        """저장된 텔레그램 채팅 ID 목록을 불러옵니다."""
        ids = []
        # 기본 ID들 (코드에 직접 포함되어 있던 것)
        default_ids = ["8634359119", "8663774579"]
        
        if os.path.exists(IDS_PATH):
            try:
                with open(IDS_PATH, "r", encoding="utf-8") as f:
                    ids = [line.strip() for line in f if line.strip()]
            except Exception as e:
                self.log(f"[오류] 채팅 ID 로드 실패: {e}")
        
        # 기본 ID들이 파일에 없으면 추가
        updated = False
        for d_id in default_ids:
            if d_id not in ids:
                ids.append(d_id)
                updated = True
        
        if updated:
            self.save_chat_ids(ids)
            
        return ids

    def save_chat_ids(self, ids):
        """텔레그램 채팅 ID 목록을 파일에 저장합니다."""
        try:
            with open(IDS_PATH, "w", encoding="utf-8") as f:
                for chat_id in ids:
                    f.write(f"{chat_id}\n")
        except Exception as e:
            self.log(f"[오류] 채팅 ID 저장 실패: {e}")

    def add_new_chat_id(self, chat_id):
        """새로운 채팅 ID를 목록에 추가하고 저장합니다."""
        if chat_id not in self.chat_ids:
            self.chat_ids.append(chat_id)
            self.save_chat_ids(self.chat_ids)
            self.log(f"🆕 새로운 사용자 연동 완료: {chat_id}")
            return True
        return False

    def load_spath_credentials(self):
        """저장된 슈패스 계정 정보를 불러옵니다."""
        if os.path.exists(CREDENTIALS_PATH):
            try:
                with open(CREDENTIALS_PATH, "r", encoding="utf-8") as f:
                    return json.load(f)
            except Exception as e:
                self.log(f"[오류] 슈패스 계정 로드 실패: {e}")
        return {}

    def save_spath_credentials(self, credentials):
        """슈패스 계정 정보를 파일에 저장합니다."""
        try:
            with open(CREDENTIALS_PATH, "w", encoding="utf-8") as f:
                json.dump(credentials, f, ensure_ascii=False, indent=4)
        except Exception as e:
            self.log(f"[오류] 슈패스 계정 저장 실패: {e}")

    def login_to_spath(self, driver, username, password):
        """슈패스(SSU-PATH) 로그인 처리"""
        try:
            # 먼저 실제 공지사항 보드에 접속해 세션 확인
            board_url = "https://path.ssu.ac.kr/ptfol/comm/board/7d5781d33124f58f831fe856c5c86e53/index.do"
            driver.get(board_url)
            time.sleep(3)
            
            # 로그인 페이지로 리다이렉트되지 않았다면 로그인 성공 상태
            if "login.do" not in driver.current_url:
                return True, "이미 로그인되어 있습니다."
                
            # 테스트에서 100% 성공했던 로그인 URL로 강제 접속
            login_url = "https://path.ssu.ac.kr/comm/login/user/login.do?rtnUrl=4c75f592029de13d9ba58cf7ce7dea55af38b35135a7fc35f6f2d024ddbedc1e"
            driver.get(login_url)
            time.sleep(3)
            
            # 첫 번째 로그인 버튼 클릭
            login_btn = driver.find_element(By.CLASS_NAME, "log_btn2")
            driver.execute_script("arguments[0].click();", login_btn)
            time.sleep(3)
            
            # 통합로그인 창에 아이디/비번 입력
            id_input = driver.find_element(By.ID, "userid")
            pw_input = driver.find_element(By.ID, "pwd")
            
            id_input.clear()
            id_input.send_keys(username)
            pw_input.clear()
            pw_input.send_keys(password)
            
            # 최종 로그인 클릭
            submit_btn = driver.find_element(By.CLASS_NAME, "btn_login")
            driver.execute_script("arguments[0].click();", submit_btn)
            time.sleep(5)
            
            # 경고창(Alert) 발생 시 처리
            try:
                alert = driver.switch_to.alert
                alert_text = alert.text
                alert.accept()
                return False, f"로그인 오류: {alert_text}"
            except:
                pass
                
            if "smartid.ssu.ac.kr" in driver.current_url or "login.do" in driver.current_url:
                return False, "아이디 또는 비밀번호가 틀렸습니다."
                
            return True, "성공"
        except Exception as e:
            import traceback
            err_msg = traceback.format_exc()
            self.log(f"로그인 중 에러 상세:\n{err_msg}")
            try:
                with open("error_log.txt", "a", encoding="utf-8") as f:
                    f.write(err_msg + "\n" + "="*50 + "\n")
            except:
                pass
            return False, f"로그인 중 에러 발생: {e}"

    def scrape_spath_notices(self, driver, keywords, start_page, end_page):
        """슈패스 공지사항 게시판 크롤링"""
        import re
        results = []
        board_url = "https://path.ssu.ac.kr/ptfol/comm/board/7d5781d33124f58f831fe856c5c86e53/index.do"

        try:
            driver.get(board_url)
            time.sleep(2)

            for p in range(start_page, end_page + 1):
                if p > 1:
                    page_moved = False

                    # 방법 1: 페이지 번호 링크 클릭 시도
                    try:
                        page_links = driver.find_elements(By.XPATH,
                            f"//a[normalize-space(text())='{p}' or contains(@onclick, '{p}')]")
                        for pl in page_links:
                            if pl.is_displayed():
                                pl.click()
                                time.sleep(2)
                                page_moved = True
                                break
                    except:
                        pass

                    # 방법 2: pageIndex URL 파라미터로 직접 이동 (폴백)
                    if not page_moved:
                        driver.get(f"{board_url}?pageIndex={p}")
                        time.sleep(2)

                # 공지 제목 링크 수집
                notice_links = []

                # 방법 1: 테이블 tbody 행 기반 추출
                try:
                    rows = driver.find_elements(By.CSS_SELECTOR, "table tbody tr")
                    for row in rows:
                        cells = row.find_elements(By.TAG_NAME, "td")
                        for cell in cells:
                            for a in cell.find_elements(By.TAG_NAME, "a"):
                                title = a.text.strip()
                                if not title or len(title) <= 2:
                                    continue
                                    
                                href = a.get_attribute("href") or ""
                                onclick = a.get_attribute("onclick") or ""
                                
                                # 슈패스 게시판 링크 패턴: global.write('1002031', './view.do')
                                m = re.search(r"global\.write\(['\"](\d+)['\"]", onclick)
                                if m:
                                    seq = m.group(1)
                                    href = f"https://path.ssu.ac.kr/ptfol/comm/board/7d5781d33124f58f831fe856c5c86e53/view.do?dataSeq={seq}"
                                else:
                                    if not href or href.endswith("#") or href.endswith("#;"):
                                        safe_t = "".join(c for c in title if c.isalnum())
                                        href = f"{board_url}#notice_{safe_t}"
                                        
                                notice_links.append((title, href))
                except Exception as e:
                    self.log(f"[슈패스 공지 테이블 파싱 오류] {e}")

                # 방법 2: 전체 a 태그 폴백
                if not notice_links:
                    try:
                        for a in driver.find_elements(By.TAG_NAME, "a"):
                            title = a.text.strip()
                            href = a.get_attribute("href") or ""
                            if title and len(title) > 5 and "path.ssu.ac.kr" in href:
                                notice_links.append((title, href))
                    except:
                        pass

                # 키워드 매칭
                for title, href in notice_links:
                    for kw in keywords:
                        if kw in title:
                            results.append((f"[슈패스 공지] {title}", href, kw))
                            break

        except Exception as e:
            self.log(f"[슈패스 공지 크롤링 오류] {e}")

        return results

    def _create_input(self, label, row, default=""):
        ctk.CTkLabel(self.input_frame, text=label, font=ctk.CTkFont(size=14)).grid(row=row, column=0, padx=10, pady=10, sticky="e")
        entry = ctk.CTkEntry(self.input_frame, width=280, font=("Malgun Gothic", 14))
        if default:
            entry.insert(0, default)
        entry.grid(row=row, column=1, padx=5, pady=10)
        return entry

    def log(self, msg):
        self.log_area.configure(state="normal")
        self.log_area.insert("end", f"{msg}\n")
        self.log_area.see("end")
        self.log_area.configure(state="disabled")

    def open_excel(self):
        if os.path.exists(XL_PATH):
            os.startfile(XL_PATH)
        else:
            self.log("[안내] '정리본.xlsx' 파일이 아직 생성되지 않았습니다. 먼저 검색을 시작해 주세요.")

    def reset_excel(self):
        if not os.path.exists(XL_PATH):
            return messagebox.showinfo("알림", "초기화할 엑셀 파일이 존재하지 않습니다.")
        
        if messagebox.askyesno("초기화 확인", "모든 엑셀 데이터가 삭제됩니다. 정말 초기화하시겠습니까?"):
            try:
                os.remove(XL_PATH)
                self.log("\n[알림] 엑셀 파일이 성공적으로 초기화되었습니다.")
                self.send_telegram_basic_message("⚠️ [SSU 공지 알림] '정리본.xlsx' 파일이 사용자에 의해 초기화되었습니다.")
                messagebox.showinfo("완료", "엑셀 파일이 삭제되었습니다.")
            except Exception as e:
                self.log(f"[오류] 파일 초기화 실패: {e}")

    def start_search(self):
        kw = self.entry_kw.get().strip()
        page = self.entry_page.get().strip()
        
        page_parts = page.split()
        if len(page_parts) >= 2 and page_parts[0].isdigit() and page_parts[1].isdigit():
            start_page = int(page_parts[0])
            end_page = int(page_parts[1])
        elif page.isdigit():
            start_page = int(page)
            end_page = int(page)
        else:
            return self.log("[오류] 페이지 수를 정확히 입력해 주세요. (예: 1 5)")

        if not kw:
            return self.log("[오류] 키워드를 입력해 주세요.")
            
        self.keywords = kw.split()
        self.start_page = start_page
        self.end_page = end_page
        self.interval_str = self.combo_interval.get()
        self.running = True
        
        # UI 비활성화
        for w in [self.btn_start, self.entry_kw, self.entry_page, self.combo_interval]:
            w.configure(state="disabled")
            
        self.btn_start.configure(text="⏳ 검색 중...")
        self.btn_stop.configure(state="normal")
        
        # 로그 초기화
        self.log_area.configure(state="normal")
        self.log_area.delete("1.0", "end")
        self.log_area.configure(state="disabled")
        
        self.log("="*50)
        self.log(f"🚀 자동 공지 모니터링 시작 ({self.interval_str} 주기)")
        self.log("="*50)
        
        self.search_status["GUI"] = True
        threading.Thread(target=self.run_loop, args=(self.keywords, self.start_page, self.end_page, self.interval_str, "GUI"), daemon=True).start()

    def stop_search(self, search_id="GUI"):
        self.search_status[search_id] = False
        if search_id == "GUI":
            self.btn_stop.configure(state="disabled", text="🛑 중지 요청 중...")
            self.log("\n[!] 중지 요청을 받았습니다. 현재 작업이 마무리되면 종료됩니다.")

    def run_loop(self, keywords, start_page, end_page, interval_str, search_id="GUI", target_chat_id=None, search_mode="both"):
        opts = Options()
        for a in ['--window-size=1920,1080', '--disable-gpu', '--log-level=3', '--disable-blink-features=AutomationControlled']:
            opts.add_argument(a)
        opts.add_experimental_option('excludeSwitches', ['enable-automation', 'enable-logging'])
        opts.add_experimental_option('useAutomationExtension', False)
        
        os.environ['WDM_LOG'] = '0'
        logging.getLogger('WDM').setLevel(logging.WARNING)
        
        driver = None
        id_msg = f"[{search_id}] " if search_id != "GUI" else ""
        first_run = True
        
        try:
            self.log(f"\n{id_msg}브라우저 환경을 준비하는 중입니다...")
            driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=opts)
            
            while self.search_status.get(search_id, False):
                mode_label = {"both": "공지 + 슈패스 통합", "ssu": "숭실대 공지사항", "spath": "슈패스(SSU-PATH)"}.get(search_mode, "통합")
                self.log(f"\n{id_msg}[{datetime.now().strftime('%H:%M:%S')}] [{mode_label}] {start_page}페이지부터 {end_page}페이지까지 검색합니다...")
                
                results = []
                if search_mode in ["both", "ssu"]:
                    for p in range(start_page, end_page + 1):
                        if not self.search_status.get(search_id, False): break
                        
                        driver.get(f"https://scatch.ssu.ac.kr/%EA%B3%B5%EC%A7%80%EC%82%AC%ED%95%AD/page/{p}/")
                        time.sleep(2)
                        
                        posts = driver.find_elements(By.TAG_NAME, "a")
                        for post in posts:
                            title = post.text
                            link = post.get_attribute("href")
                            
                            if title and link and "scatch.ssu.ac.kr" in link:
                                for kw in keywords:
                                    if kw in title:
                                        results.append((title, link, kw))
                                        break
                    
                if not self.search_status.get(search_id, False): break
                
                # 2. 슈패스(SSU-PATH) 검색 추가
                spath_results = []
                if search_mode in ["both", "spath"]:
                    credentials = self.load_spath_credentials()
                    user_creds = None
                    if target_chat_id and str(target_chat_id) in credentials:
                        user_creds = credentials[str(target_chat_id)]
                    elif not target_chat_id and credentials:
                        user_creds = list(credentials.values())[0]

                    if user_creds:
                        self.log(f"{id_msg}슈패스(SSU-PATH) 로그인을 시도합니다...")
                        login_ok, msg = self.login_to_spath(driver, user_creds["id"], user_creds["pw"])
                        if login_ok:
                            self.log(f"{id_msg}슈패스 로그인 성공! 비교과 목록을 검색합니다...")
                            spath_items = self.scrape_spath_notices(driver, keywords, start_page, end_page)
                            spath_results.extend(spath_items)
                            self.log(f"{id_msg}-> 슈패스에서 {len(spath_items)}개의 키워드 매칭 항목을 발견했습니다.")
                        else:
                            self.log(f"{id_msg}❌ 슈패스 로그인 실패: {msg}")
                    else:
                        self.log(f"{id_msg}⚠️ 슈패스 계정 미연동. '로그인 아이디 비밀번호' 명령어로 먼저 연동해 주세요.")

                results.extend(spath_results)
                
                self.log(f"{id_msg}-> 🔍 총 {len(results)}개의 키워드 매칭 공지를 발견했습니다.")
                
                # 엑셀 저장 및 사용자별 발송 대상 추출
                target_ids = [target_chat_id] if target_chat_id else self.chat_ids
                self.log(f"{id_msg}분석 대상 유저 수: {len(target_ids)}명")
                notifications_to_send = self.save_to_excel_safe(results, target_ids)
                
                self.log(f"{id_msg}" + "-" * 50)
                
                sent_any = False
                for chat_id, items in notifications_to_send.items():
                    if items:
                        sent_any = True
                        self.log(f"{id_msg}🔔 {chat_id}님에게 보낼 새로운 공지 {len(items)}건 발견!")
                        for title, link, kw in items:
                            short_link = self.shorten_url(link)
                            message = f"📢 [새로운 공지 발견]\n✨ 키워드: {kw}\n📝 제목: {title}\n🔗 바로가기: {short_link}"
                            self._send_single_message(chat_id, message)
                
                if not sent_any:
                    self.log(f"{id_msg}😴 새로운 공지사항이 발견되지 않았습니다. (이미 알림을 받았을 수 있습니다.)")
                    
                    # 텔레그램 개별 검색 시 피드백 제공 (첫 실행 혹은 요청 시)
                    if search_id != "GUI" and first_run:
                        if not results:
                            self.send_telegram_basic_message("🔍 [안내] 입력하신 키워드와 일치하는 검색 결과가 없습니다.", target_id=search_id)
                        else:
                            self.send_telegram_basic_message("✅ [안내] 검색 결과는 있으나 모두 이미 보내드린 공지입니다 (새로운 공지 없음).", target_id=search_id)
                
                first_run = False
                self.log(f"{id_msg}🔎 검색 완료. {interval_str} 후 다시 검색을 시작합니다.")
                
                # 대기 로직 (중간 중지 확인 위해 1초씩 루프)
                wait_seconds = 60
                if "분" in interval_str:
                    wait_seconds = int(interval_str.replace("분", "")) * 60
                elif "시간" in interval_str:
                    wait_seconds = int(interval_str.replace("시간", "")) * 3600
                
                for _ in range(wait_seconds):
                    if not self.search_status.get(search_id, False): break
                    time.sleep(1)
                    
        except Exception as e:
            self.log(f"\n{id_msg}[오류 발생] {e}")
        finally:
            if driver:
                driver.quit()
            self.log(f"\n{id_msg}[알림] 모니터링 루프가 종료되었습니다.")
            if search_id == "GUI":
                self.after(0, self.reset_ui)
            else:
                self.search_status.pop(search_id, None)

    def save_to_excel(self, results, target_ids):
        """공지를 엑셀에 저장하고 사용자별로 발송할 대상을 반환합니다."""
        if not results: return {}
        
        notifications = {str(tid): [] for tid in target_ids}
        
        try:
            if os.path.exists(XL_PATH):
                wb = openpyxl.load_workbook(XL_PATH)
            else:
                wb = openpyxl.Workbook()
                
            # 시트 초기화 함수
            def init_sheet(sheet_name):
                if sheet_name not in wb.sheetnames:
                    if wb.active.title == "Sheet":
                        wb.active.title = sheet_name
                        ws = wb.active
                    else:
                        ws = wb.create_sheet(sheet_name)
                else:
                    ws = wb[sheet_name]
                    
                if ws.max_row == 1 and not ws.cell(1, 1).value:
                    headers = ["검색 일시", "매칭 키워드", "공지사항 제목", "원문 링크"]
                    ws.append(headers)
                    widths = [22, 18, 75, 85]
                    for i, width in enumerate(widths, 1):
                        ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = width
                return ws
            
            ws_main = init_sheet("공지목록")
            ws_spath = init_sheet("슈패스_공지목록")
            
            # 기존 링크 로드 (두 시트 모두)
            existing_links = set()
            for ws in [ws_main, ws_spath]:
                for row in ws.iter_rows(min_row=2, values_only=True):
                    if len(row) >= 4 and row[3]:
                        existing_links.add(str(row[3]).strip())
            
            # 결과 저장 시 링크 기준으로 분리
            for title, link, kw in results:
                link_str = str(link).strip()
                if link_str not in existing_links:
                    if "path.ssu.ac.kr" in link_str:
                        ws_spath.append([datetime.now().strftime("%Y-%m-%d %H:%M:%S"), kw, title, link_str])
                    else:
                        ws_main.append([datetime.now().strftime("%Y-%m-%d %H:%M:%S"), kw, title, link_str])
                    existing_links.add(link_str)
            
            # 2. 발송이력 시트 처리 (사용자별 알림 체크)
            if "발송이력" not in wb.sheetnames:
                ws_hist = wb.create_sheet("발송이력")
                ws_hist.append(["ID", "링크", "발송일시"])
                ws_hist.column_dimensions['A'].width = 15
                ws_hist.column_dimensions['B'].width = 85
                ws_hist.column_dimensions['C'].width = 22
            else:
                ws_hist = wb["발송이력"]
            
            # 발송 이력 로드 (효율을 위해 최근 것 위주나 전체 로드)
            sent_history = set()
            for row in ws_hist.iter_rows(min_row=2, values_only=True):
                if len(row) >= 2:
                    sent_history.add((str(row[0]).strip(), str(row[1]).strip()))
            
            # 각 사용자별로 새로 보낼 것이 있는지 확인
            for tid in target_ids:
                chat_id_str = str(tid).strip()
                for title, link, kw in results:
                    link_str = str(link).strip()
                    if (chat_id_str, link_str) not in sent_history:
                        notifications[chat_id_str].append((title, link, kw))
                        # 이력에 추가
                        ws_hist.append([chat_id_str, link_str, datetime.now().strftime("%Y-%m-%d %H:%M:%S")])
                        sent_history.add((chat_id_str, link_str))
            
            # 스타일링 (메인 시트 위주)
            header_fill = PatternFill("solid", start_color="EBF1DE")
            thin_side = Side(style='thin')
            common_border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
            
            for ws in [ws_main, ws_spath, ws_hist]:
                for row in ws.iter_rows(min_row=1, max_row=ws.max_row):
                    for cell in row:
                        cell.border = common_border
                        if cell.row == 1:
                            cell.font = Font(bold=True)
                            cell.fill = header_fill
                            cell.alignment = Alignment(horizontal="center", vertical="center")
            
            wb.save(XL_PATH)
            return notifications
        except PermissionError:
            self.log("[경고] '정리본.xlsx' 파일이 다른 프로그램(엑셀 등)에서 열려 있어 저장할 수 없습니다. 엑셀을 닫아주세요!")
            return {}
        except Exception as e:
            self.log(f"[엑셀 저장 오류] {e}")
            return {}

    def send_telegram_notification(self, title, short_link, kw):
        """새로운 공지를 발견했을 때 모든 등록된 사용자에게 알림 전송 (병렬 방식)"""
        message = f"📢 [새로운 공지 발견]\n✨ 키워드: {kw}\n📝 제목: {title}\n🔗 바로가기: {short_link}"
        
        threads = []
        for chat_id in self.chat_ids:
            t = threading.Thread(target=self._send_single_message, args=(chat_id, message))
            t.start()
            threads.append(t)

    def send_telegram_basic_message(self, text, target_id=None, reply_markup=None):
        """일반 메시지 전송 (대상자가 여러 명일 경우 병렬 전송)"""
        ids = [target_id] if target_id else self.chat_ids
        
        for chat_id in ids:
            threading.Thread(target=self._send_single_message, args=(chat_id, text, reply_markup)).start()

    def _send_single_message(self, chat_id, message, reply_markup=None):
        """실제 메시지 전송을 담당하는 내부 메서드 (재시도 로직 포함)"""
        url = f"https://api.telegram.org/bot{TOKEN}/sendMessage"
        payload = {"chat_id": chat_id, "text": message}
        if reply_markup:
            payload["reply_markup"] = reply_markup
        
        for i in range(3): # 최대 3회 시도
            try:
                # 1. 세션 사용 시도
                response = self.session.post(url, json=payload, timeout=15)
                if response.status_code == 200:
                    self.log(f"✅ 알림 발송 완료 ({chat_id})")
                    return True
                else:
                    self.log(f"[알림 오류] 상태 코드: {response.status_code} ({chat_id})")
                    break
            except (requests.exceptions.ConnectionError, requests.exceptions.Timeout) as e:
                self.log(f"[알림 알림] 연결 실패 ({i+1}/3): {e}")
                if i == 1: # 2차 실패 시 세션 초기화 시도
                    self.session = requests.Session()
                time.sleep(2)
            except Exception as e:
                # 세션 없이 다이렉트 시도 (최후의 수단)
                try:
                    requests.post(url, json=payload, timeout=10)
                    self.log(f"✅ 일반 요청으로 알림 발송 성공 ({chat_id})")
                    return True
                except:
                    self.log(f"[알림 오류] 최종 발송 실패 ({chat_id}): {e}")
                    break
        return False

    def shorten_url(self, long_url):
        try:
            api_url = f"http://tinyurl.com/api-create.php?url={long_url}"
            response = self.session.get(api_url, timeout=10)
            if response.status_code == 200:
                return response.text
        except Exception as e:
            # URL 단축 실패는 크리티컬하지 않으므로 무시하고 원본 링크 반환
            pass
        return long_url

    def save_to_excel_safe(self, results, target_ids):
        """스레드 세이프한 엑셀 저장"""
        with self.excel_lock:
            return self.save_to_excel(results, target_ids)

    def send_telegram_notification_to_id(self, title, short_link, kw, chat_id):
        """특정 사용자에게 알림 전송 (재시도 포함)"""
        message = f"📢 [새로운 공지 발견]\n✨ 키워드: {kw}\n📝 제목: {title}\n🔗 바로가기: {short_link}"
        self._send_single_message(chat_id, message)

    def on_closing(self):
        """프로그램 종료 시 사용자에게 알림 후 종료"""
        any_running = any(self.search_status.values())
        
        should_close = False
        if any_running:
            if messagebox.askokcancel("종료 확인", "현재 공지사항 검색이 진행 중입니다.\n정말로 프로그램을 종료하시겠습니까?"):
                should_close = True
        else:
            should_close = True
            
        if should_close:
            try:
                # 모든 검색 상태 중지
                for key in list(self.search_status.keys()):
                    self.search_status[key] = False
                
                # 텔레그램 종료 알림 발송
                self.log("\n[시스템] 텔레그램으로 종료 알림을 보냅니다...")
                msg = "🛑 [SSU 공지 알림] 모니터링 프로그램이 종료되었습니다.\n새로운 공지 감시가 중단됩니다."
                
                # 종료 시에는 스레드보다 직접 발송이 안전 (메시지 전송 대기)
                for chat_id in self.chat_ids:
                    try:
                        url = f"https://api.telegram.org/bot{TOKEN}/sendMessage"
                        payload = {"chat_id": chat_id, "text": msg}
                        # 세션을 사용하여 빠르게 발송
                        self.session.post(url, json=payload, timeout=5)
                    except:
                        pass
                
                self.log("[시스템] 종료 준비 완료.")
            except:
                pass
            finally:
                self.destroy()

    def reset_ui(self):
        self.btn_start.configure(state="normal", text="🔍 검색 시작")
        self.btn_stop.configure(state="disabled", text="🛑 검색 중지")
        self.entry_kw.configure(state="normal")
        self.entry_page.configure(state="normal")
        self.combo_interval.configure(state="normal")

    def start_telegram_polling(self):
        """텔레그램 메시지를 실시간으로 감시하는 스레드"""
        self.log("\n[시스템] 텔레그램 명령 대기 중...")
        
        # 텔레그램 봇 메뉴 설정 (좌측 하단 메뉴 버튼 생성)
        try:
            menu_url = f"https://api.telegram.org/bot{TOKEN}/setMyCommands"
            commands = {
                "commands": [
                    {"command": "help", "description": "도움말 및 기능 버튼 띄우기"},
                    {"command": "status", "description": "현재 검색 상태 확인"},
                    {"command": "stop", "description": "진행 중인 검색 중지"}
                ]
            }
            self.session.post(menu_url, json=commands, timeout=5)
        except:
            pass
            
        # 봇 시작 시 자동 알림은 사용자 요청에 의해 제거됨
        
        while True:
            try:
                # getUpdates를 이용한 롱 폴링 (Long Polling)
                url = f"https://api.telegram.org/bot{TOKEN}/getUpdates"
                params = {"offset": self.last_update_id + 1, "timeout": 30}
                
                # 타임아웃을 넉넉히 주어 연결 유지
                response = self.session.get(url, params=params, timeout=40)
                
                if response.status_code == 200:
                    data = response.json()
                    updates = data.get("result", [])
                    
                    for update in updates:
                        self.last_update_id = update["update_id"]
                        
                        if "callback_query" in update:
                            cq = update["callback_query"]
                            data = cq.get("data")
                            message = cq.get("message", {})
                            chat_id = str(message.get("chat", {}).get("id"))
                            
                            self.log(f"🔘 인라인 버튼 클릭: {data} (ID: {chat_id})")
                            
                            # 콜백 쿼리 응답 (버튼 로딩 표시 제거)
                            try:
                                cb_id = cq["id"]
                                url_cb = f"https://api.telegram.org/bot{TOKEN}/answerCallbackQuery"
                                self.session.post(url_cb, json={"callback_query_id": cb_id})
                            except:
                                pass
                                
                            self.handle_telegram_command(data, chat_id)
                            continue
                            
                        if "message" in update and "text" in update["message"]:
                            msg_text = update["message"]["text"]
                            chat_id = str(update["message"]["chat"]["id"])
                            
                            # 신규 사용자 자동 연동
                            is_new = self.add_new_chat_id(chat_id)
                            
                            self.log(f"📩 텔레그램 메시지 수신: {msg_text} (ID: {chat_id})")
                            
                            # 사용자 상태 확인 (검색어 입력 대기 중인지)
                            state = self.user_states.get(chat_id)
                            
                            # 입력받은 텍스트가 예약된 명령어(버튼 클릭)인지 확인
                            clean_msg = msg_text.strip()
                            for emoji in ["ℹ️", "📊", "🛑", "🔓", "▶️", "🍚"]:
                                clean_msg = clean_msg.replace(emoji, "").strip()
                            cmd_word = clean_msg.split()[0].lower() if clean_msg else ""
                            is_reserved = cmd_word in ["start", "시작", "도움말", "help", "중지", "stop", "상태", "status", "로그아웃", "공지검색", "슈패스검색", "학식", "학식메뉴", "메뉴", "오늘의"] or clean_msg.startswith("/")
                            
                            if state and state.startswith("awaiting_keyword_"):
                                if is_reserved:
                                    # 명령어 버튼을 누른 경우 대기 상태 취소 후 정상 명령어 처리
                                    self.user_states[chat_id] = None
                                else:
                                    mode = state.split("_")[-1]
                                    # 입력한 텍스트를 검색어로 취급하여 명령어 조합 (기본 1페이지)
                                    msg_text = f"{mode} {msg_text} 1"
                                    self.user_states[chat_id] = None # 상태 초기화
                                    self.handle_telegram_command(msg_text, chat_id)
                                    continue
                            
                            # /start 또는 시작 관련 명령어가 아니더라도 신규 사용자면 환영 메시지 전송
                            is_start_cmd = msg_text.strip().lower() in ["/start", "start", "/시작", "시작"]
                            if is_new and not is_start_cmd:
                                self.handle_telegram_command("/start", chat_id)
                            else:
                                self.handle_telegram_command(msg_text, chat_id)

                
            except requests.exceptions.RequestException:
                # 네트워크 연결 일시 오류 등은 무시하고 재시도
                pass
            except Exception as e:
                self.log(f"[폴링 오류] {e}")
                time.sleep(5) # 오류 발생 시 잠시 대기
            
            time.sleep(0.5)

    def handle_telegram_command(self, text, chat_id):
        """텔레그램 명령어를 해석하고 실행"""
        text = text.strip()
        # Reply Keyboard 버튼의 이모지 제거
        for emoji in ["ℹ️", "📊", "🛑", "🔓", "▶️", "🍚"]:
            text = text.replace(emoji, "").strip()
            
        if not text:
            return

        parts = text.split()
        cmd = parts[0].lower()
        if cmd.startswith("/"):
            cmd = cmd[1:]

        if cmd in ["start", "시작", "도움말", "help"]:
            msg = ("👋 안녕하세요! SSU 공지사항 검색 봇입니다.\n\n"
                   "하단에 고정된 메뉴 버튼을 눌러 편리하게 제어하거나,\n"
                   "직접 명령어를 채팅으로 입력할 수도 있습니다.\n\n"
                   "📌 검색 명령어 직접 입력 방법:\n"
                   "1️⃣ <검색어> <시작페이지> <끝페이지> : 숭실대 공지만 검색\n"
                   "   예: 장학금 2 5\n"
                   "2️⃣ 슈패스 <검색어> <시작페이지> <끝페이지> : 슈패스만 검색\n"
                   "   (슈패스 검색은 먼저 '로그인 <아이디> <비밀번호>' 입력 필요)\n"
                   "3️⃣ 로그인 <아이디> <비밀번호> : 슈패스 연동")
            
            reply_markup = {
                "keyboard": [
                    [
                        {"text": "ℹ️ 도움말"},
                        {"text": "📊 상태"}
                    ],
                    [
                        {"text": "🛑 중지"},
                        {"text": "🔓 로그아웃"}
                    ],
                    [
                        {"text": "▶️ 공지검색"},
                        {"text": "▶️ 슈패스검색"}
                    ],
                    [
                        {"text": "🍚 오늘의 학식 메뉴"}
                    ]
                ],
                "resize_keyboard": True,
                "persistent": True
            }
            self.send_telegram_basic_message(msg, target_id=chat_id, reply_markup=reply_markup)
            
        elif cmd in ["공지검색", "슈패스검색"]:
            # 공지검색, 슈패스검색 버튼 클릭 시
            mode = cmd.replace("검색", "")
            self.user_states[chat_id] = f"awaiting_keyword_{mode}"
            self.send_telegram_basic_message(f"🔍 **{mode} 검색**을 시작합니다.\n채팅창에 원하시는 **검색어**를 입력해주세요!\n(예: 장학금, 마일리지, 공모전)", target_id=chat_id)
            return

        elif cmd.startswith("ask_keyword_"):
            mode = cmd.split("_")[-1]
            self.user_states[chat_id] = f"awaiting_keyword_{mode}"
            self.send_telegram_basic_message(f"🔍 **{mode} 검색**을 시작합니다.\n채팅창에 원하시는 **검색어**를 입력해주세요!\n(예: 장학금, 마일리지, 공모전)", target_id=chat_id)
            return

        elif cmd in ["학식", "메뉴", "학식메뉴", "오늘의"]:
            self.send_telegram_basic_message("⏳ 학식 메뉴를 불러오는 중입니다...", target_id=chat_id)
            menu_text = self.get_cafeteria_menu()
            self.send_telegram_basic_message(menu_text, target_id=chat_id)
            return

        elif cmd == "로그인":
            if len(parts) < 3:
                self.send_telegram_basic_message("❌ 사용법: 로그인 <아이디> <비밀번호>\n예: 로그인 20201234 비밀번호123", target_id=chat_id)
                return
            username = parts[1]
            password = parts[2]
            
            credentials = self.load_spath_credentials()
            credentials[str(chat_id)] = {"id": username, "pw": password}
            self.save_spath_credentials(credentials)
            self.send_telegram_basic_message("🔑 슈패스 로그인 계정이 연동되었습니다. 이제 슈패스(SSU-PATH) 비교과 공지도 함께 검색됩니다.", target_id=chat_id)

        elif cmd == "로그아웃":
            credentials = self.load_spath_credentials()
            chat_id_str = str(chat_id)
            if chat_id_str in credentials:
                del credentials[chat_id_str]
                self.save_spath_credentials(credentials)
                self.send_telegram_basic_message("🔓 슈패스 로그인 계정이 연동 해제되었습니다.", target_id=chat_id)
            else:
                self.send_telegram_basic_message("😴 연동된 슈패스 계정이 없습니다.", target_id=chat_id)

        elif cmd in ["공지", "ssu", "notice", "숭실"]:
            # 숭실대 공지사항만 검색
            kw_parts = parts[1:]
            if not kw_parts:
                self.send_telegram_basic_message("❌ 사용법: 공지 <검색어> <시작페이지> <끝페이지>\n예: 공지 장학금 2 5", target_id=chat_id)
                return
            try:
                if len(kw_parts) >= 2 and kw_parts[-2].isdigit() and kw_parts[-1].isdigit():
                    start_page = int(kw_parts[-2])
                    end_page = int(kw_parts[-1])
                    keywords = kw_parts[:-2]
                elif kw_parts[-1].isdigit():
                    start_page = int(kw_parts[-1])
                    end_page = int(kw_parts[-1])
                    keywords = kw_parts[:-1]
                else:
                    start_page = 1
                    end_page = 1
                    keywords = kw_parts
                if not keywords:
                    self.send_telegram_basic_message("❌ 검색할 키워드를 입력해 주세요.", target_id=chat_id)
                    return
                self.search_status[chat_id] = True
                threading.Thread(target=self.run_loop,
                                 args=(keywords, start_page, end_page, "10분", chat_id, chat_id),
                                 kwargs={"search_mode": "ssu"},
                                 daemon=True).start()
                self.send_telegram_basic_message(
                    f"📢 숭실대 공지사항 검색을 시작합니다 (10분 주기).\n🔍 키워드: {' '.join(keywords)}\n📄 범위: {start_page} ~ {end_page}페이지",
                    target_id=chat_id)
                
                credentials = self.load_spath_credentials()
                if str(chat_id) not in credentials:
                    self.send_telegram_basic_message(
                        "💡 꿀팁: 슈패스 계정을 연동하시면 슈패스 비교과도 검색할 수 있어요!\n\n"
                        "📌 연동 방법:\n"
                        "로그인 <아이디> <비밀번호>",
                        target_id=chat_id)
            except Exception as e:
                self.send_telegram_basic_message(f"❌ 검색 설정 중 오류 발생: {e}", target_id=chat_id)

        elif cmd in ["슈패스", "spath", "path"]:
            # 슈패스(SSU-PATH)만 검색
            kw_parts = parts[1:]
            if not kw_parts:
                self.send_telegram_basic_message("❌ 사용법: 슈패스 <검색어> <시작페이지> <끝페이지>\n예: 슈패스 장학금 2 5", target_id=chat_id)
                return
            # 계정 연동 여부 확인
            credentials = self.load_spath_credentials()
            if str(chat_id) not in credentials:
                self.send_telegram_basic_message(
                    "🔐 슈패스(SSU-PATH) 로그인이 필요합니다!\n\n"
                    "슈패스 검색을 사용하려면 먼저 아이디와 비밀번호를 등록해 주세요.\n\n"
                    "📌 등록 방법:\n"
                    "로그인 <아이디> <비밀번호>\n\n"
                    "예시:\n"
                    "로그인 20231234 mypassword123\n\n"
                    "✅ 등록 후 다시 슈패스 검색을 시도해 주세요!",
                    target_id=chat_id)
                return
            try:
                if len(kw_parts) >= 2 and kw_parts[-2].isdigit() and kw_parts[-1].isdigit():
                    start_page = int(kw_parts[-2])
                    end_page = int(kw_parts[-1])
                    keywords = kw_parts[:-2]
                elif kw_parts[-1].isdigit():
                    start_page = int(kw_parts[-1])
                    end_page = int(kw_parts[-1])
                    keywords = kw_parts[:-1]
                else:
                    start_page = 1
                    end_page = 1
                    keywords = kw_parts
                if not keywords:
                    self.send_telegram_basic_message("❌ 검색할 키워드를 입력해 주세요.", target_id=chat_id)
                    return
                self.search_status[chat_id] = True
                threading.Thread(target=self.run_loop,
                                 args=(keywords, start_page, end_page, "10분", chat_id, chat_id),
                                 kwargs={"search_mode": "spath"},
                                 daemon=True).start()
                self.send_telegram_basic_message(
                    f"🎓 슈패스(SSU-PATH) 검색을 시작합니다 (10분 주기).\n🔍 키워드: {' '.join(keywords)}\n📄 범위: {start_page} ~ {end_page}페이지",
                    target_id=chat_id)
            except Exception as e:
                self.send_telegram_basic_message(f"❌ 검색 설정 중 오류 발생: {e}", target_id=chat_id)

        elif cmd in ["stop", "중지", "정지"]:
            if self.search_status.get(chat_id, False):
                self.stop_search(chat_id)
                self.send_telegram_basic_message("🛑 본인의 검색 중지 요청을 보냈습니다.", target_id=chat_id)
            else:
                self.send_telegram_basic_message("😴 현재 실행 중인 본인의 검색이 없습니다.", target_id=chat_id)

        elif cmd in ["status", "상태"]:
            gui_status = "✅ 대기 중" if not self.search_status.get("GUI", False) else "🔍 실행 중"
            my_status = "✅ 실행 중" if self.search_status.get(chat_id, False) else "💤 대기 중"
            
            msg = (f"📊 [현재 상태]\n"
                   f"🤖 전체 서버 상태: {gui_status}\n"
                   f"👤 내 검색 상태: {my_status}")
            self.send_telegram_basic_message(msg, target_id=chat_id)

        else:
            # 기본 명령어 이외의 입력은 모두 검색 키워드로 간주 (예: '장학금 5' 또는 '검색 장학금 5', '통합 장학금 1')
            if cmd in ["search", "검색", "통합"]:
                keywords_parts = parts[1:]
            else:
                keywords_parts = parts

            if not keywords_parts:
                self.send_telegram_basic_message("❌ 검색할 키워드를 입력해 주세요.", target_id=chat_id)
                return

            try:
                # 마지막 요소들이 숫자이면 페이지 수로 인식
                if len(keywords_parts) >= 2 and keywords_parts[-2].isdigit() and keywords_parts[-1].isdigit():
                    start_page = int(keywords_parts[-2])
                    end_page = int(keywords_parts[-1])
                    keywords = keywords_parts[:-2]
                elif keywords_parts[-1].isdigit():
                    start_page = int(keywords_parts[-1])
                    end_page = int(keywords_parts[-1])
                    keywords = keywords_parts[:-1]
                else:
                    start_page = 1
                    end_page = 1
                    keywords = keywords_parts

                if not keywords:
                    self.send_telegram_basic_message("❌ 검색할 키워드를 입력해 주세요.", target_id=chat_id)
                    return

                # 개별 스레드에서 검색 시작 (기본 공지 검색)
                self.search_status[chat_id] = True
                threading.Thread(target=self.run_loop, 
                                 args=(keywords, start_page, end_page, "10분", chat_id, chat_id),
                                 kwargs={"search_mode": "ssu"},
                                 daemon=True).start()
                
                self.send_telegram_basic_message(f"🚀 공지 검색을 시작합니다 (10분 주기).\n🔍 키워드: {' '.join(keywords)}\n📄 범위: {start_page} ~ {end_page}페이지", target_id=chat_id)

                credentials = self.load_spath_credentials()
                if str(chat_id) not in credentials:
                    self.send_telegram_basic_message(
                        "💡 꿀팁: 슈패스 계정을 연동하시면 슈패스 비교과도 검색할 수 있어요!\n\n"
                        "📌 연동 방법:\n"
                        "로그인 <아이디> <비밀번호>",
                        target_id=chat_id)

            except Exception as e:
                self.send_telegram_basic_message(f"❌ 검색 설정 중 오류 발생: {e}", target_id=chat_id)

    def remote_start_search(self, keywords, start_page, end_page):
        """텔레그램 등 외부에서 검색을 시작할 때 GUI 값을 설정하고 시작"""
        if self.running:
            self.log("[!] 이미 검색이 진행 중입니다. 먼저 중지해 주세요.")
            return

        # GUI 입력창 업데이트
        self.entry_kw.delete(0, "end")
        self.entry_kw.insert(0, keywords)
        
        self.entry_page.delete(0, "end")
        self.entry_page.insert(0, f"{start_page} {end_page}")
        
        # 검색 시작 버튼 클릭 시뮬레이션
        self.start_search()

    def get_cafeteria_menu(self):
        """soongguri.com 에서 당일 학식 메뉴 (중식 1,2,3 / 석식 1) 스크래핑"""
        import re
        import html
        try:
            url = "https://soongguri.com/main.php?mkey=2&w=3"
            headers = {"User-Agent": "Mozilla/5.0"}
            response = requests.get(url, headers=headers, timeout=10)
            response.encoding = 'utf-8'
            html_text = response.text
            
            # 셀 내용 정리 함수
            def clean_text(text):
                text = re.sub(r'<br\s*/?>', '\n', text, flags=re.IGNORECASE)
                text = re.sub(r'</(p|div|li|ul|td|tr)>', '\n', text, flags=re.IGNORECASE)
                text = re.sub(r'<[^>]+>', '', text)
                text = html.unescape(text)
                
                lines = [line.strip() for line in text.split('\n') if line.strip()]
                cleaned_lines = []
                
                skip_mode = False
                for line in lines:
                    if "알러지유발식품" in line or "원산지" in line or "*알러지" in line or "*원산지" in line or "* 알러지" in line or "* 원산지" in line:
                        skip_mode = True
                        continue
                    if skip_mode:
                        continue
                        
                    if line and len(line) < 60:
                        alpha_count = sum(1 for c in line if c.isalpha() and c.isascii())
                        if alpha_count > len(line) * 0.5:
                            continue
                        line = line.replace("★", "").replace("\xa0", " ").strip()
                        line = re.sub(r'\[.*?\]', '', line).strip()
                        if line:
                            cleaned_lines.append(line)
                return " / ".join(cleaned_lines)

            # 메뉴 칸(cell) 추출
            cells = re.findall(r'<td style="width:[^>]+;text-align:left;padding:3px;border:1px dotted #999999;vertical-align:top;">(.*?)</td>', html_text, flags=re.DOTALL|re.IGNORECASE)
            
            if len(cells) < 7:
                return "❌ 학식 메뉴 구조가 변경되어 불러올 수 없습니다."
                
            # 학생식당 중식 1,2,3은 cells[1], cells[2], cells[3]
            
            output = "🍚 **오늘의 학식 메뉴** 🍚\n\n"
            output += "🏢 **[학생식당]**\n"
            for i, corner_name in zip([1, 2, 3], ["중식 1코너", "중식 2코너", "중식 3코너"]):
                menu_items = clean_text(cells[i])
                if menu_items:
                    output += f"🔸 **{corner_name}**: {menu_items}\n"
                    
            return output.strip()
        except Exception as e:
            return f"❌ 학식 메뉴를 불러오는데 실패했습니다: {e}"

if __name__ == "__main__":
    app = NoticeSearchApp()
    app.mainloop()
