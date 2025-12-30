import json
import os
import time
import shutil
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

# --- 設定區 ---
# 程式會自動在桌面建立一個「輔具下載結果」資料夾
DOWNLOAD_DIR = os.path.join(os.path.expanduser("~"), "Desktop", "SFAA_Downloads")
if not os.path.exists(DOWNLOAD_DIR):
    os.makedirs(DOWNLOAD_DIR)

# --- Chrome 參數設定 ---
chrome_options = Options()
settings = {
    "recentDestinations": [{"id": "Save as PDF", "origin": "local", "account": ""}],
    "selectedDestinationId": "Save as PDF",
    "version": 2
}
prefs = {
    "printing.print_preview_sticky_settings.appState": json.dumps(settings),
    "savefile.default_directory": DOWNLOAD_DIR
}
chrome_options.add_experimental_option("prefs", prefs)
chrome_options.add_argument('--kiosk-printing') # 開啟自動列印

def get_latest_file(path):
    files = os.listdir(path)
    paths = [os.path.join(path, basename) for basename in files if basename.endswith('.pdf')]
    return max(paths, key=os.path.getctime) if paths else None

def start_process():
    # 讀取清單
    if not os.path.exists('items.txt'):
        print("錯誤：找不到 items.txt 檔案！")
        return

    with open('items.txt', 'r', encoding='utf-8') as f:
        items = [line.strip() for line in f if line.strip()]

    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
    wait = WebDriverWait(driver, 20)

    try:
        driver.get("https://ptps.sfaa.gov.tw/mall")
        
        for name in items:
            print(f"正在處理：{name}")
            try:
                # 找到搜尋框
                search_box = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "input[placeholder*='關鍵字']")))
                search_box.clear()
                search_box.send_keys(name)
                search_box.send_keys(Keys.ENTER)
                
                time.sleep(5) # 等待網頁產出結果
                
                # 紀錄下載前的檔案清單
                old_files = os.listdir(DOWNLOAD_DIR)
                
                # 執行列印
                driver.execute_script('window.print();')
                
                # 等待新檔案出現並改名
                timeout = 15
                while timeout > 0:
                    time.sleep(1)
                    new_files = os.listdir(DOWNLOAD_DIR)
                    if len(new_files) > len(old_files):
                        time.sleep(1) # 確保檔案寫入完成
                        latest_file = get_latest_file(DOWNLOAD_DIR)
                        if latest_file:
                            new_path = os.path.join(DOWNLOAD_DIR, f"{name.replace('/', '_')}.pdf")
                            os.rename(latest_file, new_path)
                            print(f"完成！存為：{new_path}")
                        break
                    timeout -= 1
                
            except Exception as e:
                print(f"跳過項目 {name}，原因：{e}")
                
    finally:
        print("任務結束。")
        driver.quit()

if __name__ == "__main__":
    start_process()
