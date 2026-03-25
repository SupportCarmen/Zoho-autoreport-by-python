import os
import json
from datetime import datetime
from zoneinfo import ZoneInfo
from playwright.sync_api import sync_playwright

from config import *
from discord_bot import send_to_discord
from excel_bot import update_master
from logger_util import log

SESSION_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'session.json')

def is_today_session():
    if not os.path.exists(SESSION_FILE):
        return False
    try:
        with open(SESSION_FILE, 'r', encoding='utf-8') as f:
            data = json.load(f)
            today = datetime.now(ZoneInfo('Asia/Bangkok')).strftime('%Y-%m-%d')
            return data.get('date') == today
    except Exception:
        return False

def main():
    for d in [FOLDER, REPORT_FOLDER]:
        os.makedirs(d, exist_ok=True)
        
    master_folder = os.path.join(os.path.expanduser("~"), 'Downloads', 'All')
    os.makedirs(master_folder, exist_ok=True)
    
    # %Y-%m-%dT%H:%M:%S in JS -> we want same format e.g. 2023-10-05T12-30-00 -> JS format string equivalent
    now_dt = datetime.now(ZoneInfo('Asia/Bangkok'))
    now_str = now_dt.strftime('%Y-%m-%d_%H-%M-%S')
    
    log.info("🔐 ตรวจสอบ session...")
    
    with sync_playwright() as p:
        browser = None
        context = None
        
        try:
            browser = p.chromium.launch(headless=False)
            
            if is_today_session():
                with open(SESSION_FILE, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    state = data.get('state')
                context = browser.new_context(storage_state=state)
                log.success("Session loaded, skip login")
            else:
                log.info("🔑 ไม่มี session, กำลัง login...")
                context = browser.new_context()
                login_page = context.new_page()
                
                login_page.goto("https://accounts.zoho.com/signin", timeout=60000)
                login_page.fill('#login_id', ZOHO_EMAIL)
                login_page.click('#nextbtn')
                login_page.wait_for_selector('#password', timeout=60000)
                login_page.fill('#password', ZOHO_PASSWORD)
                login_page.click('#nextbtn')
                login_page.wait_for_load_state('networkidle')
                login_page.wait_for_timeout(6000)
                log.success("Login success")
                
                state = context.storage_state()
                today = datetime.now(ZoneInfo('Asia/Bangkok')).strftime('%Y-%m-%d')
                with open(SESSION_FILE, 'w', encoding='utf-8') as f:
                    json.dump({"date": today, "state": state}, f)
                log.success("Session saved")
                login_page.close()
                
            log.step("📊 เปิด Dashboard...")
            page = context.new_page()
            page.goto(DASHBOARD_URL, timeout=60000)
            page.wait_for_timeout(6000)
            log.success("Dashboard โหลดสำเร็จ")
            
            try:
                btn = page.locator('text="Not Now"')
                if btn.count() > 0:
                    btn.click()
            except Exception:
                pass
                
            log.step("📸 เริ่ม capture dashboard...")
            selector = '.zd_v2-dashboarddetailcontainer-container'
            images = []
            scroll_steps = [0, 300, 400, 1200]
            
            for i in range(4):
                if scroll_steps[i] > 0:
                    page.mouse.wheel(0, scroll_steps[i])
                    page.wait_for_timeout(2000)
                
                file_path = os.path.join(FOLDER, f"{now_str}_dashboard_{i + 1}.png")
                page.locator(selector).screenshot(path=file_path)
                images.append(file_path)
                log.success(f"capture {i + 1}/4")
                
            log.step("📥 เริ่ม download reports...")
            reports = []
            report_entries = []
            
            for report in REPORTS:
                r_name = report['name']
                r_url = report['url']
                log.info(f"📥 กำลัง download: {r_name}...")
                response = page.request.get(r_url)
                buffer = response.body()
                
                file_path = os.path.join(REPORT_FOLDER, f"{r_name}_{now_str}.xls")
                with open(file_path, "wb") as f:
                    f.write(buffer)
                    
                reports.append(file_path)
                report_entries.append({"file": file_path, "name": r_name})
                log.success(f"{r_name} downloaded")
                
        except Exception as e:
            log.error(f"เกิดข้อผิดพลาดใน Playwright: {e}")
            raise e
        finally:
            if context:
                context.close()
            if browser:
                browser.close()
                log.info("Browser closed")
                
    # นอกเหนือจากการใช้ Playwright จะได้ประหยัด memory
    log.step("สร้าง Master Excel...")
    master_file = os.path.join(master_folder, f"Tickets_AllZoho_{now_str}.xlsx")
    update_master(report_entries, master_file)
    
    log.step("ส่งไฟล์ไป Discord...")
    # แพ็คไฟล์ทั้งหมดเตรียมส่ง (รูปถ่าย + Master Excel)
    if os.path.exists(master_file):
        files_to_send = images + [master_file]
    else:
        files_to_send = images # ถ้า Excel ช็อต ก็ส่งแค่รูป
        
    send_to_discord(files_to_send, now_str, WEBHOOK)

    log.step("เสร็จสิ้นทุกขั้นตอน 🎉")

if __name__ == "__main__":
    main()
