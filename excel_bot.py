import os
import time
import pandas as pd
import win32com.client as win32
from logger_util import log

def clean_data(file_path):
    # อ่านไฟล์ Excel โดยใช้ engine xlrd สำหรับ .xls
    df = pd.read_excel(file_path, header=None, engine='xlrd')
    
    # 5 แถวแรก: โลโก้ และ header
    # header อยู่ใน row index 4 (บรรทัดที่ 5)
    header = df.iloc[4].tolist()
    
    # ข้อมูลเริ่มต้นที่ row index 5
    data = df.iloc[5:].copy()
    
    # กรองเอาแค่ แถวที่ column แรกไม่ว่างค่า และไม่เป็นคำว่า total records (การลบเซลล์รวมท้ายสุด)
    data = data.dropna(subset=[0])
    data = data[~data[0].astype(str).str.lower().str.contains('total records')]
    
    # กำหนด header ให้คอลัมน์
    data.columns = header
    return data

def update_master(report_files, master_file):
    open_all_file = next((r['file'] for r in report_files if r['name'] == 'OpenAll'), None)
    ticket_today_file = next((r['file'] for r in report_files if r['name'] == 'TicketToday'), None)
    
    if not open_all_file or not os.path.exists(open_all_file):
        log.warning("File not found: OpenAll")
        return None
    if not ticket_today_file or not os.path.exists(ticket_today_file):
        log.warning("File not found: TicketToday")
        return None
        
    log.info("📂 เปิดไฟล์: OpenAll.xls")
    df1 = clean_data(open_all_file)
    log.success(f"OpenAll → {len(df1)} rows")
    
    log.info("📂 เปิดไฟล์: TicketToday.xls")
    df2 = clean_data(ticket_today_file)
    log.success(f"TicketToday → {len(df2)} rows")
    
    log.info("💾 สร้าง Master Excel เบื้องต้น...")
    # บันทึกเป็น .xlsx แทน
    with pd.ExcelWriter(master_file, engine='openpyxl') as writer:
        df1.to_excel(writer, sheet_name='OpenAll_Data', index=False)
        df2.to_excel(writer, sheet_name='Today_Data', index=False)
        # สร้างแผ่นงานเปล่าๆ ให้ Pivot มาลง
        pd.DataFrame([['READY']]).to_excel(writer, sheet_name='Open_All', index=False, header=False)
        
    log.success("Master saved")
    
    # แทนที่จะเรียก VBScript เราสามารถสั่ง win32com ควบคุม Excel ได้ตรงๆ (เหมือนใช้งาน VBA ใน Python)
    abs_master = os.path.abspath(master_file)
    log.info("🔄 สร้าง PivotTable...")
    
    try:
        excel = win32.Dispatch("Excel.Application")
        excel.Visible = True
        excel.DisplayAlerts = False
        
        wb = excel.Workbooks.Open(abs_master)
        time.sleep(2)
        
        # ลบ sheet default ที่อาจติดมา
        for ws in list(wb.Sheets):
            if ws.Name == "Sheet1":
                ws.Delete()
                
        time.sleep(0.5)
        
        wsDst = wb.Sheets("Open_All")
        ws1src = wb.Sheets("OpenAll_Data")
        ws2src = wb.Sheets("Today_Data")
        
        wsDst.Cells.Clear()
        time.sleep(0.5)
        wsDst.Activate()
        time.sleep(0.5)
        
        # 1 = xlRowField, 2 = xlColumnField, -4112 = xlCount
        
        # ── PT 1: OpenAll / Support Member Assigned ──
        log.info("🔄 PT 1/4: OpenAll - Support Member Assigned...")
        pc1 = wb.PivotCaches().Create(1, ws1src.UsedRange)
        pt1 = pc1.CreatePivotTable(wsDst.Cells(1, 1), "PT_Open_Member")
        pt1.PivotFields("Support Member Assigned").Orientation = 1
        pt1.PivotFields("Status (Ticket)").Orientation = 2
        pt1.AddDataField(pt1.PivotFields("Ticket Id"), "Count of Ticket Id", -4112)
        time.sleep(2)
        log.success("PT 1/4 สำเร็จ")
        
        # ── PT 2: OpenAll / Product Category ──
        log.info("🔄 PT 2/4: OpenAll - Product Category...")
        lastRow = wsDst.UsedRange.Rows.Count + 3
        pt2 = pc1.CreatePivotTable(wsDst.Cells(lastRow, 1), "PT_Open_Product")
        pt2.PivotFields("Product Category").Orientation = 1
        pt2.PivotFields("Status (Ticket)").Orientation = 2
        pt2.AddDataField(pt2.PivotFields("Ticket Id"), "Count of Ticket Id", -4112)
        time.sleep(2)
        log.success("PT 2/4 สำเร็จ")
        
        # ── PT 3: Today / Support Member Assigned ──
        log.info("🔄 PT 3/4: Today - Support Member Assigned...")
        lastRow = wsDst.UsedRange.Rows.Count + 3
        pc2 = wb.PivotCaches().Create(1, ws2src.UsedRange)
        pt3 = pc2.CreatePivotTable(wsDst.Cells(lastRow, 1), "PT_Today_Member")
        pt3.PivotFields("Support Member Assigned").Orientation = 1
        pt3.PivotFields("Status (Ticket)").Orientation = 2
        pt3.AddDataField(pt3.PivotFields("Ticket Id"), "Count of Ticket Id", -4112)
        time.sleep(2)
        log.success("PT 3/4 สำเร็จ")
        
        # ── PT 4: Today / Product Category ──
        log.info("🔄 PT 4/4: Today - Product Category...")
        lastRow = wsDst.UsedRange.Rows.Count + 3
        pt4 = pc2.CreatePivotTable(wsDst.Cells(lastRow, 1), "PT_Today_Product")
        pt4.PivotFields("Product Category").Orientation = 1
        pt4.PivotFields("Status (Ticket)").Orientation = 2
        pt4.AddDataField(pt4.PivotFields("Ticket Id"), "Count of Ticket Id", -4112)
        time.sleep(2)
        log.success("PT 4/4 สำเร็จ")
        
        # ── สร้าง PivotChart สำหรับแต่ละ PivotTable ──
        log.info("📊 สร้าง PivotChart...")
        
        # สร้าง sheet ใหม่สำหรับ Charts
        wsChart = wb.Sheets.Add(After=wb.Sheets(wb.Sheets.Count))
        wsChart.Name = "Charts"
        wsChart.Activate()
        time.sleep(1)
        
        # xlColumnClustered = 51, xl3DColumnClustered = 54
        xlColumnClustered = 51
        
        chart_configs = [
            {"pt": pt1, "title": "OpenAll - Support Member Assigned", "idx": 1},
            {"pt": pt2, "title": "OpenAll - Product Category",        "idx": 2},
            {"pt": pt3, "title": "Today - Support Member Assigned",   "idx": 3},
            {"pt": pt4, "title": "Today - Product Category",          "idx": 4},
        ]
        
        for cfg in chart_configs:
            i = cfg["idx"]
            # คำนวณตำแหน่ง: 2 chart ต่อแถว, แต่ละ chart 480x300 px
            col_offset = 0 if (i % 2 == 1) else 500
            row_offset = ((i - 1) // 2) * 320
            
            log.info(f"📊 Chart {i}/4: {cfg['title']}...")
            
            chart_obj = wsChart.ChartObjects().Add(
                Left=10 + col_offset,
                Top=10 + row_offset,
                Width=480,
                Height=300
            )
            chart = chart_obj.Chart
            chart.SetSourceData(cfg["pt"].TableRange2)
            chart.ChartType = xlColumnClustered
            chart.HasTitle = True
            chart.ChartTitle.Text = cfg["title"]
            
            # ปรับ style ให้สวยงาม
            try:
                chart.Style = 201  # Dark Style
            except Exception:
                pass  # บาง version อาจไม่รองรับ style นี้
            
            time.sleep(1)
            log.success(f"Chart {i}/4 สำเร็จ")
        
        log.success("PivotChart ทั้ง 4 สร้างเสร็จ")
        
        wb.Save()
        time.sleep(1)
        wb.Close()
        excel.Quit()
        log.success("PivotTable & PivotChart created")
        
    except Exception as e:
        log.error(f"error: {e}")
        try:
            excel.Quit() # ให้ปิด excel แน่ๆ ในกรณี error
        except:
            pass

    return master_file
