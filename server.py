import http.server
import socketserver
import os
import time
import json
import sys
import threading
import win32com.client as win32

sys.stdout.reconfigure(encoding='utf-8')

PORT = 8080
DIRECTORY = r"D:\AI AGI_2"
EXCEL_PATH = r"D:\AI AGI_2\THEO DOI HOP DONG-2_Optimized.xlsx"
JS_DATA_PATH = r"D:\AI AGI_2\dashboard_data.js"
JSON_DATA_PATH = r"D:\AI AGI_2\dashboard_data.json"

last_mtime = 0
last_check_time = 0
extract_lock = threading.Lock()

def extract_excel_data():
    global last_mtime, last_check_time
    
    now = time.time()
    if now - last_check_time < 5:
        return False

    if not extract_lock.acquire(blocking=False):
        return False
        
    try:
        last_check_time = now
        if not os.path.exists(EXCEL_PATH):
            print(f"File Excel not found: {EXCEL_PATH}")
            return False
        
        current_mtime = os.path.getmtime(EXCEL_PATH)
        if current_mtime <= last_mtime:
            return False  # No change

        print(f"[{time.strftime('%H:%M:%S')}] Phát hiện file Excel thay đổi! Đang cập nhật dữ liệu Dashboard...")
        
        excel = None
        wb = None
        try:
            excel = win32.DispatchEx('Excel.Application')
            excel.Visible = False
            excel.DisplayAlerts = False
            excel.AutomationSecurity = 1

            abs_src = os.path.abspath(EXCEL_PATH)
            wb = excel.Workbooks.Open(abs_src, 0, True)
            time.sleep(1.0)
            
            ws = wb.Sheets('CHI')
            vals = ws.UsedRange.Value
            wb.Close(False)
            excel.Quit()
            excel = None
            wb = None
            
            rows = [list(r) for r in vals if any(x is not None for x in r)]
            data = rows[1:]
            
            processed_items = []
            for idx, r in enumerate(data):
                if len(r) < 17:
                    continue
                loai_cp = r[0]
                if loai_cp is None or str(loai_cp).strip() == '':
                    continue
                    
                tieu_muc = str(r[1]).strip() if r[1] is not None else ''
                so_hd = str(r[2]).strip() if r[2] is not None else ''
                ngay_hd = str(r[3])[:10] if r[3] is not None else ''
                thang = str(r[4]).strip() if r[4] is not None else ''
                ly_do = str(r[5]).strip() if r[5] is not None else ''
                chi_tiet = str(r[6]).strip() if r[6] is not None else ''
                chi_tiet_hd = str(r[7]).strip() if (len(r)>7 and r[7] is not None) else ''
                kho = str(r[8]).strip() if (len(r)>8 and r[8] is not None) else ''
                
                try: st_no_vat = float(r[9]) if (len(r)>9 and r[9] is not None) else 0.0
                except: st_no_vat = 0.0
                    
                try: vat = float(r[10]) if (len(r)>10 and r[10] is not None) else 0.0
                except: vat = 0.0

                try: st_vat = float(r[11]) if (len(r)>11 and r[11] is not None) else 0.0
                except: st_vat = 0.0
                    
                ngay_tt = str(r[12])[:10] if (len(r)>12 and r[12] is not None) else ''
                nguoi_thu_huong = str(r[13]).strip() if (len(r)>13 and r[13] is not None) else ''
                stk = str(r[14]).strip() if (len(r)>14 and r[14] is not None) else ''
                ngan_hang = str(r[15]).strip() if (len(r)>15 and r[15] is not None) else ''
                
                try: nam = int(float(r[16])) if (len(r)>16 and r[16] is not None) else 'N/A'
                except: nam = 'N/A'

                processed_items.append({
                    'id': idx + 1,
                    'loai_cp': str(loai_cp).strip(),
                    'tieu_muc': tieu_muc,
                    'so_hd': so_hd,
                    'ngay_hd': ngay_hd,
                    'thang': thang,
                    'ly_do': ly_do,
                    'chi_tiet': chi_tiet,
                    'chi_tiet_hd': chi_tiet_hd,
                    'kho': kho if kho != 'None' else 'Khác',
                    'st_no_vat': st_no_vat,
                    'vat': vat,
                    'st_vat': st_vat,
                    'ngay_tt': ngay_tt,
                    'nguoi_thu_huong': nguoi_thu_huong if nguoi_thu_huong != 'None' else '',
                    'stk': stk if stk != 'None' else '',
                    'ngan_hang': ngan_hang if ngan_hang != 'None' else '',
                    'nam': nam
                })
                
            sync_info = {
                'last_updated': time.strftime('%Y-%m-%d %H:%M:%S'),
                'total_rows': len(processed_items)
            }
            
            with open(JSON_DATA_PATH, 'w', encoding='utf-8') as f:
                json.dump(processed_items, f, ensure_ascii=False, indent=2)

            with open(JS_DATA_PATH, 'w', encoding='utf-8') as f:
                f.write(f'window.SYNC_INFO = {json.dumps(sync_info, ensure_ascii=False)};\n')
                f.write('window.RAW_DATA = ' + json.dumps(processed_items, ensure_ascii=False) + ';')

            last_mtime = current_mtime
            print(f"[{time.strftime('%H:%M:%S')}] Cập nhật thành công! Tổng số {len(processed_items)} dòng.")
            return True
        except Exception as e:
            print(f"Lỗi khi đọc file Excel: {e}")
            return False
        finally:
            if wb:
                try: wb.Close(False)
                except: pass
            if excel:
                try: excel.Quit()
                except: pass
    finally:
        extract_lock.release()

class AutoUpdateHandler(http.server.SimpleHTTPRequestHandler):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=DIRECTORY, **kwargs)

    def do_GET(self):
        try:
            if self.path in ('/', '/index.html'):
                extract_excel_data()
        except Exception as e:
            print(f"Lỗi trong do_GET: {e}")
        return super().do_GET()

if __name__ == '__main__':
    extract_excel_data()
    with socketserver.TCPServer(("", PORT), AutoUpdateHandler) as httpd:
        print(f"Server đang chạy tại http://localhost:{PORT}")
        print("Mỗi khi bạn mở hoặc F5 lại trang Dashboard, server sẽ tự kiểm tra và đọc file Excel mới nhất!")
        httpd.serve_forever()
