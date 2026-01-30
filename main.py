from nicegui import ui
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime
import os
import json
import openpyxl
import io

# ==========================================
# 1. CONFIGURATION
# ==========================================
SHEET_NAME = "DEPA_PO_SYSTEM"
CURRENT_YEAR_TAB = "PO_2569"
JSON_KEY_FILE = "service_account.json"
TEMPLATE_FILE = "template_po.xlsx"

# UI Styles
STYLE_INPUT = 'w-full bg-white' 
PROPS_INPUT = 'outlined dense color="teal"'
STYLE_CARD = 'w-full max-w-6xl bg-white shadow-lg rounded-lg border border-gray-200 p-0 mx-auto'
STYLE_LABEL = 'text-xs font-bold text-gray-500 uppercase mb-1'

# ==========================================
# 2. BACKEND LOGIC
# ==========================================

def get_client():
    """เชื่อมต่อ Google Sheets"""
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    json_env = os.getenv("GOOGLE_JSON_KEY")
    
    if json_env:
        creds = ServiceAccountCredentials.from_json_keyfile_dict(json.loads(json_env), scope)
    elif os.path.exists(JSON_KEY_FILE):
        creds = ServiceAccountCredentials.from_json_keyfile_name(JSON_KEY_FILE, scope)
    else:
        return None
    return gspread.authorize(creds)

def get_worksheet():
    """เปิด Tab และจัดการ Header ให้ตรงกับรูปภาพ"""
    client = get_client()
    if not client: return None
    try:
        sheet = client.open(SHEET_NAME)
        try:
            ws = sheet.worksheet(CURRENT_YEAR_TAB)
        except:
            # สร้างชีทใหม่ถ้ายังไม่มี
            ws = sheet.add_worksheet(title=CURRENT_YEAR_TAB, rows=100, cols=20)
            # Header ตามรูปภาพจริง (A-K)
            headers = ['วันที่', 'เลขที่ PO', 'รายการ', 'เลขที่ PR', 'รหัสงบ', 'ผู้ขาย', 'เลขภาษี', 'ยอดสุทธิ', 'กำหนดส่ง', 'ผู้จัดทำ', 'สถานะ']
            ws.append_row(headers)
        return ws
    except Exception as e:
        print(f"Sheet Error: {e}")
        return None

def get_next_po_number():
    """รันเลข PO อัตโนมัติ (อ่าน Column B)"""
    ws = get_worksheet()
    if not ws: return "ซจ.001"
    try:
        # อ่าน Column 2 (B) คือเลข PO (แก้จากเดิมที่อ่าน Col 3)
        po_col = ws.col_values(2) 
        if len(po_col) <= 1: return "ซจ.001"
        
        last_po = po_col[-1] # เอาตัวล่าสุด
        if "ซจ." in last_po:
            # รูปแบบ: ซจ.001/2569 -> ตัดเอาแค่ 001
            try:
                # แยก / ก่อน ถ้ามี
                parts = last_po.split('/')
                number_part = parts[0].replace("ซจ.", "").strip()
                new_num = int(number_part) + 1
                return f"ซจ.{new_num:03d}"
            except:
                return "ซจ.001"
        return "ซจ.001"
    except:
        return "ซจ.001"

def generate_excel_bytes(data):
    """สร้างไฟล์ Excel (เหมือนเดิม)"""
    if not os.path.exists(TEMPLATE_FILE):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws['A1'] = "TEMPLATE NOT FOUND"
    else:
        wb = openpyxl.load_workbook(TEMPLATE_FILE)
        ws = wb.active

    # Mapping Excel (ปรับพิกัดตาม Template จริงของคุณ)
    ws['H3'] = f"เลขที่ {data['po_no']}"
    ws['H4'] = data['date']
    ws['B6'] = data['vendor']
    ws['B7'] = data['address']
    ws['B8'] = data['tax_id']
    
    start_row = 17
    for i, item in enumerate(data['items']):
        r = start_row + i
        ws[f'B{r}'] = i + 1
        ws[f'C{r}'] = item['desc']
        ws[f'H{r}'] = float(item['qty'])
        ws[f'I{r}'] = item['unit']
        ws[f'K{r}'] = float(item['price'])
        ws[f'L{r}'] = float(item['qty']) * float(item['price'])

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output

# ==========================================
# 3. UI APPLICATION
# ==========================================
@ui.page('/')
def main_page():
    
    # Custom Font
    ui.add_head_html("""
        <style>
            @import url('https://fonts.googleapis.com/css2?family=Sarabun:wght@300;400;500;600;700&display=swap');
            body { font-family: 'Sarabun', sans-serif; background-color: #f0f2f5; }
        </style>
    """)

    # --- STATE MANAGEMENT (รวมตัวแปรทั้งหมดไว้ที่นี่) ---
    state = {
        'po_no_run': get_next_po_number(), # เลขรัน (ซจ.XXX)
        'date': datetime.now().strftime('%Y-%m-%d'),
        'pr_no': '',       # เลข PR (เดิมไม่มี)
        'budget_code': '', # รหัสงบ (เดิมไม่มี)
        'vendor': '',
        'address': '',
        'tax_id': '',
        'delivery_date': '', # กำหนดส่ง
        'items': [{'desc': '', 'qty': 1, 'unit': 'งาน', 'price': 0}],
        'grand_total': 0
    }

    # --- Functions ---
    def calculate():
        total = sum(float(x['qty']) * float(x['price']) for x in state['items'])
        vat = total * 0.07
        grand = total + vat
        label_total.text = f"{total:,.2f}"
        label_vat.text = f"{vat:,.2f}"
        label_grand.text = f"{grand:,.2f}"
        state['grand_total'] = grand

    def save_and_export():
        ws = get_worksheet()
        if not ws:
            ui.notify('เชื่อมต่อ Google Sheets ไม่ได้', type='negative')
            return

        ui.notify('กำลังบันทึก...', type='info', spinner=True)

        # 1. Format Data
        full_po_no = f"{state['po_no_run']}/{datetime.now().year + 543}"
        
        # รวมรายการสินค้าเป็น string เดียว (สำหรับโชว์ใน Sheet ช่องเดียว)
        items_summary = ", ".join([i['desc'] for i in state['items'] if i['desc']])

        # 2. Prepare Row Data (เรียงตามคอลัมน์ A-K ในรูปภาพเป๊ะๆ)
        row_data = [
            state['date'],              # A: วันที่
            full_po_no,                 # B: เลขที่ PO
            items_summary,              # C: รายการ
            state['pr_no'],             # D: เลขที่ PR
            state['budget_code'],       # E: รหัสงบ
            state['vendor'],            # F: ผู้ขาย
            state['tax_id'],            # G: เลขภาษี
            state['grand_total'],       # H: ยอดสุทธิ
            state['delivery_date'],     # I: กำหนดส่ง
            "เจ้าหน้าที่พัสดุ",           # J: ผู้จัดทำ
            "รออนุมัติ"                  # K: สถานะ
        ]

        # 3. Excel Data Object
        excel_data = {
            'po_no': full_po_no,
            'date': state['date'],
            'vendor': state['vendor'],
            'address': state['address'],
            'tax_id': state['tax_id'],
            'items': state['items']
        }

        try:
            # Save to Sheet
            ws.append_row(row_data)
            
            # Generate Excel
            excel_file = generate_excel_bytes(excel_data)
            filename = f"PO_{state['po_no_run']}.xlsx"
            ui.download(excel_file.read(), filename)

            ui.notify(f'✅ บันทึก {full_po_no} เรียบร้อย!', type='positive')
            
            # Reset & Refresh
            state['po_no_run'] = get_next_po_number()
            po_input.value = state['po_no_run']
            refresh_history_table()

        except Exception as e:
            ui.notify(f'Error: {e}', type='negative')

    # --- LAYOUT ---
    with ui.column().classes('w-full items-center py-6 px-4'):
        
        # HEADER
        with ui.row().classes('w-full max-w-6xl items-center justify-between mb-4'):
            with ui.row().classes('items-center'):
                ui.icon('assignment', size='lg').classes('text-teal-700')
                ui.label('DEPA PO SYSTEM').classes('text-2xl font-bold text-gray-800')
            ui.button('ประวัติ / ฐานข้อมูล', on_click=lambda: drawer.toggle(), icon='history').props('flat color=teal')

        # MAIN FORM CARD
        with ui.card().classes(STYLE_CARD):
            # Title Bar
            with ui.row().classes('bg-teal-700 w-full p-4 text-white'):
                ui.label('บันทึกขออนุมัติจัดซื้อ/จัดจ้าง').classes('text-lg font-bold')

            with ui.column().classes('p-6 w-full gap-6'):
                
                # --- SECTION 1: HEADER INFO ---
                with ui.grid(columns=4).classes('w-full gap-4'):
                    # PO Number
                    with ui.column():
                        ui.label('เลขที่ PO (Auto)').classes(STYLE_LABEL)
                        po_input = ui.input().bind_value(state, 'po_no_run').props(PROPS_INPUT + ' readonly').classes('font-bold')
                    
                    # Date
                    with ui.column():
                        ui.label('วันที่เอกสาร').classes(STYLE_LABEL)
                        ui.input().bind_value(state, 'date').props(PROPS_INPUT + ' type=date')

                    # PR No (Bind แล้ว!)
                    with ui.column():
                        ui.label('เลขที่ PR อ้างอิง').classes(STYLE_LABEL)
                        ui.input().bind_value(state, 'pr_no').props(PROPS_INPUT) # Bind

                    # Delivery Date (Bind แล้ว!)
                    with ui.column():
                        ui.label('กำหนดส่งของ (วว/ดด/ปป)').classes(STYLE_LABEL)
                        ui.input().bind_value(state, 'delivery_date').props(PROPS_INPUT) # Bind

                ui.separator()

                # --- SECTION 2: VENDOR & BUDGET ---
                ui.label('ข้อมูลผู้ขาย & งบประมาณ').classes('text-lg font-bold text-gray-700')
                with ui.grid(columns=3).classes('w-full gap-4'):
                    # Vendor
                    with ui.column().classes('col-span-2'):
                        ui.label('ชื่อบริษัท / ร้านค้า').classes(STYLE_LABEL)
                        ui.input(placeholder='ระบุชื่อผู้ขาย...').bind_value(state, 'vendor').props(PROPS_INPUT)
                    
                    # Budget Code (Bind แล้ว!)
                    with ui.column():
                        ui.label('รหัสงบประมาณ').classes(STYLE_LABEL)
                        ui.input().bind_value(state, 'budget_code').props(PROPS_INPUT) # Bind

                    # Address
                    with ui.column().classes('col-span-2'):
                        ui.label('ที่อยู่').classes(STYLE_LABEL)
                        ui.textarea(placeholder='ที่อยู่...').bind_value(state, 'address').props(PROPS_INPUT + ' rows=2')

                    # Tax ID (Bind แล้ว!)
                    with ui.column():
                        ui.label('เลขประจำตัวผู้เสียภาษี').classes(STYLE_LABEL)
                        ui.input().bind_value(state, 'tax_id').props(PROPS_INPUT) # Bind

                ui.separator()

                # --- SECTION 3: ITEMS ---
                ui.label('รายการสินค้า').classes('text-lg font-bold text-gray-700')
                
                # Header Row
                with ui.row().classes('w-full bg-gray-100 py-2 px-2 rounded border border-gray-200 text-xs font-bold text-gray-600'):
                    ui.label('รายการ').classes('flex-grow pl-2')
                    ui.label('จำนวน').classes('w-24 text-center')
                    ui.label('หน่วย').classes('w-24 text-center')
                    ui.label('ราคา/หน่วย').classes('w-32 text-right pr-2')

                # Dynamic List
                @ui.refreshable
                def items_list():
                    for i, item in enumerate(state['items']):
                        with ui.row().classes('w-full gap-2 mb-1 items-start'):
                            ui.input().bind_value(item, 'desc').props('dense outlined bg-color=white').classes('flex-grow')
                            ui.number(min=1, on_change=calculate).bind_value(item, 'qty').props('dense outlined bg-color=white input-class=text-center').classes('w-24')
                            ui.input().bind_value(item, 'unit').props('dense outlined bg-color=white input-class=text-center').classes('w-24')
                            ui.number(min=0, on_change=calculate).bind_value(item, 'price').props('dense outlined bg-color=white input-class=text-right').classes('w-32')
                            
                            # ปุ่มลบแถว (ถ้ามีมากกว่า 1 แถว)
                            if len(state['items']) > 1:
                                ui.button(icon='delete', on_click=lambda idx=i: delete_item(idx)).props('flat dense color=red round')
                    
                    with ui.row().classes('mt-2'):
                        ui.button('เพิ่มรายการ', icon='add', on_click=add_item).props('outline dense color=teal')

                def add_item():
                    state['items'].append({'desc':'', 'qty':1, 'unit':'', 'price':0})
                    items_list.refresh()

                def delete_item(index):
                    state['items'].pop(index)
                    calculate()
                    items_list.refresh()

                items_list()

                # --- SECTION 4: SUMMARY & ACTIONS ---
                with ui.row().classes('w-full justify-end mt-4 gap-6'):
                    with ui.column().classes('items-end gap-1'):
                        with ui.row().classes('gap-4'):
                            ui.label('รวมเป็นเงิน:').classes('text-gray-600')
                            label_total = ui.label('0.00').classes('font-bold w-24 text-right')
                        with ui.row().classes('gap-4'):
                            ui.label('ภาษีมูลค่าเพิ่ม 7%:').classes('text-gray-600')
                            label_vat = ui.label('0.00').classes('font-bold w-24 text-right')
                        ui.separator().classes('my-1')
                        with ui.row().classes('gap-4 items-center'):
                            ui.label('ยอดสุทธิ:').classes('text-lg font-bold text-teal-800')
                            label_grand = ui.label('0.00').classes('text-xl font-bold text-teal-700 w-32 text-right')

                with ui.row().classes('w-full border-t pt-4 justify-end gap-4'):
                    ui.button('บันทึกและออกใบ PO', on_click=save_and_export, icon='save').props('color=teal size=lg')

    # --- DRAWER: HISTORY ---
    with ui.right_drawer(fixed=False).classes('bg-white w-2/3 border-l') as drawer:
        with ui.row().classes('p-4 border-b justify-between items-center'):
            ui.label('ประวัติการสั่งซื้อ (เชื่อมต่อ Google Sheets)').classes('text-xl font-bold')
            ui.button(icon='close', on_click=lambda: drawer.toggle()).props('flat round')
        
        grid = ui.aggrid({
            'columnDefs': [
                {'headerName': 'วันที่', 'field': 'วันที่', 'width': 100},
                {'headerName': 'PO', 'field': 'เลขที่ PO', 'width': 120},
                {'headerName': 'ผู้ขาย', 'field': 'ผู้ขาย', 'width': 150},
                {'headerName': 'ยอดสุทธิ', 'field': 'ยอดสุทธิ', 'width': 100},
                {'headerName': 'สถานะ', 'field': 'สถานะ', 'width': 100},
            ],
            'rowData': []
        }).classes('h-full w-full')

        def refresh_history_table():
            ws = get_worksheet()
            if ws:
                data = ws.get_all_records()
                grid.options['rowData'] = data
                grid.update()

    # Initial Calculation
    calculate()
    drawer.hide() # ซ่อน Drawer ก่อน

ui.run(title='DEPA PO System Fixed', port=8080, language='th')
