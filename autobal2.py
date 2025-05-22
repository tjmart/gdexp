import os
import threading
import time
import tkinter as tk
import json
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# 설정 로딩 함수
def load_config():
    default_config = {
        "downloads_path": os.path.expanduser("~\\Downloads"),
        "kd_id": "",
        "kd_pw": "",
        "kh_id": "",
        "kh_pw": ""
    }
    if os.path.exists("config.json"):
        with open("config.json", "r", encoding="utf-8") as f:
            return json.load(f)
    else:
        return default_config

config = load_config()

# 크롬 드라이버 설정
chrome_driver_path = "./chromedriver.exe"
service = Service(chrome_driver_path)
options = webdriver.ChromeOptions()
options.add_experimental_option("detach", True)

driver = webdriver.Chrome(service=service, options=options)
driver.set_window_size(1600, 900)

# 로그인 수행 함수
def login_kyungdong(driver, userid, password):
    try:
        print("[자동화] 경동 로그인 화면 진입")
        driver.get("https://ong.kdexp.com")

        # 아이디, 비밀번호, 로그인 버튼 대기 및 입력
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "mblid")))
        id_input = driver.find_element(By.ID, "mblid")
        pw_input = driver.find_element(By.ID, "mpw")
        login_btn = driver.find_element(By.ID, "login")

        id_input.clear()
        id_input.send_keys(userid)
        pw_input.clear()
        pw_input.send_keys(password)

        print("✅ 로그인 정보 입력 완료 (로그인 버튼 앞까지 대기 중)")
        # 여기서 로그인 클릭은 수동으로 기다리기
        WebDriverWait(driver, 300).until(
            EC.presence_of_element_located((By.XPATH, "//a[contains(text(), '발송')]"))
        )
        print("✅ 수동 로그인 성공, '발송' 메뉴 감지됨")

    except Exception as e:
        print(f"[오류] 로그인 처리 중 문제 발생: {e}")





# ============ 자동화 메인 루프 ============
def automation_loop():
    """반복 처리(연속 실행/1회 실행)까지 포함."""
    global stop_signal, driver

    driver = init_driver()

    # 탭1: 경동물류, 탭2: KH업무관리
    driver.get("https://ong.kdexp.com")
    time.sleep(1)
    driver.execute_script("window.open('https://kdexp.xyz', '_blank');")

    # 파일 이름 등 config
    downloads_path = config["downloads_path"]
    target_name = config["target_name"]
    file_path = os.path.join(downloads_path, target_name)

    # 아이디 / 비번
    kd_id = config["kd_id"]
    kd_pw = config["kd_pw"]
    kh_id = config["kh_id"]
    kh_pw = config["kh_pw"]
    print(f"[정보] 경동물류 ID/PW: {kd_id}/{kd_pw}, KH ID/PW: {kh_id}/{kh_pw}")

    first_run = True

    while not stop_signal:
        try:
            # ------------------- (A) 경동물류 자동화 -------------------
            driver.switch_to.window(driver.window_handles[0])
            driver.switch_to.default_content()

            # 1) '발송' 메뉴 (최초만)
            if first_run:
                menu_items = driver.find_elements(By.XPATH, "//*[contains(text(), '발송')]")
                found_menu = False
                for item in menu_items:
                    if item.text.strip() == '발송':
                        print("[자동화] '발송' 메뉴 마우스오버 or 클릭")
                        ActionChains(driver).move_to_element(item).perform()
                        time.sleep(1)
                        found_menu = True
                        break
                if not found_menu:
                    print("❌ '발송' 메뉴를 찾을 수 없습니다.")
                    raise Exception("발송 메뉴 클릭 실패")

                # 2) 하위 '발송자료조회' 클릭
                sub_item = WebDriverWait(driver, 5).until(
                    EC.element_to_be_clickable((By.XPATH, "//*[contains(text(), '발송자료조회')]"))
                )
                sub_item.click()
                time.sleep(2)

                # iframe 진입
                driver.switch_to.default_content()
                WebDriverWait(driver, 10).until(
                    EC.frame_to_be_available_and_switch_to_it("iframe-발송자료조회")
                )
                print("[자동화] iframe-발송자료조회 진입 완료")

                # '오늘' 버튼
                print("[자동화] '오늘' 버튼 클릭 (최초 한 번)")
                today_btn = driver.find_element(By.XPATH, "//button[contains(text(),'오늘')]")
                today_btn.click()
                time.sleep(1)
            else:
                # 이미 발송자료조회 열려 있다고 가정
                print("[자동화] 발송자료조회 이미 열림 (2회차 이상)")
                driver.switch_to.frame("iframe-발송자료조회")

            # 조회
            search_btn = driver.find_element(By.ID, "searchBtn")
            search_btn.click()
            time.sleep(2)

            # 목록수 1000
            driver.find_element(By.CSS_SELECTOR, "#sendDataPageSize .wj-input-group-btn").click()
            time.sleep(1)
            driver.find_element(By.XPATH, "//div[text()='1000']").click()
            time.sleep(2)

            driver.switch_to.default_content()
            driver.switch_to.frame("iframe-발송자료조회")
            time.sleep(1)

            # 전체선택
            checkboxes = driver.find_elements(By.CSS_SELECTOR, "input[type='checkbox']")
            for cb in checkboxes:
                if 'wj-column-selector-group' in (cb.get_attribute("class") or ""):
                    cb.click()
                    time.sleep(1.5)
                    break

            # 엑셀 다운로드
            excel_btns = driver.find_elements(By.TAG_NAME, "button")
            excel_buttons = [b for b in excel_btns if "excel" in (b.get_attribute("class") or "").lower()]
            if excel_buttons:
                excel_buttons[-1].click()
                time.sleep(3)

            # ------------------- (B) KH 자동화 -------------------
            driver.switch_to.default_content()
            driver.switch_to.window(driver.window_handles[1])
            time.sleep(1)

            # '발송관리' 클릭
            menu_items = driver.find_elements(By.XPATH, "//*[contains(text(), '발송관리')]")
            for item in menu_items:
                if item.is_displayed():
                    item.click()
                    time.sleep(2)
                    break

            # '발송UP(신)테스트'
            buttons = driver.find_elements(By.CSS_SELECTOR, "button.savebtn")
            for btn in buttons:
                txt = btn.text.replace("\n", "").replace(" ", "")
                if "발송UP" in txt and "테스트" in txt:
                    btn.click()
                    print("✅ '발송UP(신)테스트' 클릭 완료")
                    time.sleep(2)
                    break

            # 파일 업로드
            driver.execute_script("document.getElementById('upfile25').style.display='block';")
            file_input = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "#upfile25 input[type='file']"))
            )
            file_input.send_keys(file_path)
            time.sleep(2)

            upload_form = driver.find_element(By.CSS_SELECTOR, "form[name='depupload25']")
            upload_form.submit()
            time.sleep(3)

            # 알림 팝업 처리
            try:
                WebDriverWait(driver, 10).until(EC.alert_is_present())
                alert = driver.switch_to.alert
                alert.accept()
                print("[자동화] 확인 팝업 처리 완료")
                time.sleep(2)
            except:
                print("[자동화] 팝업 없음 or 오류")

            # 파일 삭제
            if os.path.exists(file_path):
                os.remove(file_path)
                print("[자동화] 엑셀 파일 삭제 완료")

            status_label.config(text="✅ 자동화 완료")

            # 루프 반복 or 종료
            if not loop_var.get():
                break
            first_run = False
            time.sleep(3)

        except Exception as e:
            print(f"[에러] 자동화 오류: {e}")
            stop_signal = True
            break

    driver.quit()
    print("[자동화] 드라이버 종료 및 루프 종료")

def start_automation():
    global stop_signal
    stop_signal = False
    status_label.config(text="⏳ 자동화 실행 중...")
    threading.Thread(target=automation_loop, daemon=True).start()

def stop_automation():
    global stop_signal
    stop_signal = True
    status_label.config(text="⏹️ 중지됨")

# ============ 설정 창 열기 ============
def open_settings():
    settings_win = tk.Toplevel(root)
    settings_win.title("설정")
    settings_win.geometry("400x420")

    tk.Label(settings_win, text="다운로드 폴더:").pack(pady=2)
    downloads_entry = tk.Entry(settings_win, width=40)
    downloads_entry.pack()
    downloads_entry.insert(0, config["downloads_path"])

    def browse_downloads():
        folder = filedialog.askdirectory()
        if folder:
            downloads_entry.delete(0, tk.END)
            downloads_entry.insert(0, folder)

    tk.Button(settings_win, text="폴더찾기", command=browse_downloads).pack(pady=2)

    tk.Label(settings_win, text="다운로드 파일명:").pack(pady=2)
    target_entry = tk.Entry(settings_win, width=40)
    target_entry.pack()
    target_entry.insert(0, config["target_name"])

    tk.Label(settings_win, text="ChromeDriver 경로:").pack(pady=2)
    driver_entry = tk.Entry(settings_win, width=40)
    driver_entry.pack()
    driver_entry.insert(0, config["chrome_driver_path"])

    tk.Label(settings_win, text="경동물류 ID:").pack(pady=2)
    kd_id_entry = tk.Entry(settings_win, width=40)
    kd_id_entry.pack()
    kd_id_entry.insert(0, config["kd_id"])

    tk.Label(settings_win, text="경동물류 PW:").pack(pady=2)
    kd_pw_entry = tk.Entry(settings_win, width=40, show='*')
    kd_pw_entry.pack()
    kd_pw_entry.insert(0, config["kd_pw"])

    tk.Label(settings_win, text="KH업무관리 ID:").pack(pady=2)
    kh_id_entry = tk.Entry(settings_win, width=40)
    kh_id_entry.pack()
    kh_id_entry.insert(0, config["kh_id"])

    tk.Label(settings_win, text="KH업무관리 PW:").pack(pady=2)
    kh_pw_entry = tk.Entry(settings_win, width=40, show='*')
    kh_pw_entry.pack()
    kh_pw_entry.insert(0, config["kh_pw"])

    def save_settings():
        config["downloads_path"] = downloads_entry.get()
        config["target_name"] = target_entry.get()
        config["chrome_driver_path"] = driver_entry.get()
        config["kd_id"] = kd_id_entry.get()
        config["kd_pw"] = kd_pw_entry.get()
        config["kh_id"] = kh_id_entry.get()
        config["kh_pw"] = kh_pw_entry.get()
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(config, f, ensure_ascii=False, indent=2)
        print("[설정] 저장 완료.")
        settings_win.destroy()

    tk.Button(settings_win, text="저장", command=save_settings).pack(pady=8)

# GUI 메인
root = tk.Tk()
root.title("경동-KH 자동화 with Settings")
root.geometry("420x230")

loop_var = tk.BooleanVar()

status_label = tk.Label(root, text="🟢 대기 중", fg="green")
status_label.pack(pady=5)

tk.Button(root, text="🛠 설정", command=open_settings, width=30).pack(pady=2)
tk.Button(root, text="▶️ Start 자동화", command=start_automation, width=30, bg="lightgreen").pack(pady=2)
tk.Checkbutton(root, text="🔁 연속 실행", variable=loop_var).pack(pady=2)
tk.Button(root, text="⏹️ Stop 자동화", command=stop_automation, width=30, bg="salmon").pack(pady=2)

root.mainloop()
