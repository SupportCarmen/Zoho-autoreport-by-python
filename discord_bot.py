import requests
from datetime import datetime
from logger_util import log

def send_to_discord(files, now_str, webhook_url):
    day = datetime.now().strftime("%A")
    
    message = f"📊Dashboards_Tickets\n🕒{day} {now_str}\n@everyone"
    data = {"content": message}
    
    multipart_files = []
    for i, file_path in enumerate(files):
        multipart_files.append((f'file{i}', open(file_path, 'rb')))
        
    response = requests.post(webhook_url, data=data, files=multipart_files)
    
    for _, f in multipart_files:
        f.close()
        
    log.success("ส่งข้อมูลไปที่ Discord สำเร็จแล้ว")
