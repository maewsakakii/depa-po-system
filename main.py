from nicegui import ui
import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime
import os
import json
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side
import io

# ==========================================
# 1. CONFIGURATION & STYLES
# ==========================================
SHEET_NAME = "DEPA_PO_SYSTEM"
CURRENT_YEAR_TAB = "PO_2569"
JSON_KEY_FILE = "service_account.json"
TEMPLATE_FILE = "template_po.xlsx"

# --- UI STYLES (Tailwind mimicking) ---
# บังคับ Style ให้เหมือน HTML ที่คุณชอบ
STYLE_INPUT = 'w-full bg-white' 
PROPS_INPUT = 'outlined dense color="teal"' # ใช้ Quasar style แบบเส้นบาง
STYLE_CARD = 'w-full max-w-5xl bg-white shadow-xl rounded-lg overflow-hidden border border-gray-200 p-0 mx-auto'
STYLE_LABEL = 'text-xs font-bold text-gray-500 uppercase mb-1'

# ==========================================
# 2. BACKEND LOGIC (Google Sheets + Excel)
# ==========================================

def get_client():
    """เชื่อมต่อ Google Sheets"""
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    # 1. Environment Variable (Cloud)
    json_env = os.getenv("GOOGLE_JSON_KEY")
    if json_env:
        creds = ServiceAccountCredentials.from_json_keyfile_dict(json.loads(json_env), scope)
    # 2. Local File
    elif os.path.exists(JSON_KEY_FILE):
        creds = ServiceAccountCredentials.from_json_keyfile_name(JSON_KEY_FILE, scope)
    else:
        return None
    return gspread.authorize(creds)

def get_worksheet():
    """เปิด Tab ปัจจุบัน"""
    client = get_client()
    if not client: return None
    try:
        sheet = client.open(SHEET_NAME)
        try:
            ws = sheet.worksheet(CURRENT_YEAR_TAB)
        except:
            ws = sheet.add_worksheet(title=CURRENT_YEAR_TAB, rows=100, cols=20)
            ws.append_row(['Timestamp', 'วันที่', 'เลขที่ PO', 'รายการ', 'ผู้ขาย', 'ยอดสุทธิ', 'ผู้จัดทำ', 'สถานะ'])
        return ws
    except Exception as e:
        print(f"Sheet Error: {e}")
        return None

def get_next_po_number():
    """รันเลข PO อัตโนมัติ"""
    ws = get_worksheet()
    if not ws: return "ซจ.001"
    try:
        po_col = ws.col_values(3) # Column C = PO No.
        if len(po_col) <= 1: return "ซจ.001"
        last_po = po_col[-1]
        if "ซจ." in last_po:
            # ตัด ซจ.001/2569 เอาแค่เลข
            num_part = last_po.split('/')[0].replace("ซจ.", "")
            return f"ซจ.{int(num_part) + 1:03d}"
        return "ซจ.001"
    except:
        return "ซจ.001"

def generate_excel_bytes(data):
    """สร้างไฟล์ Excel PO (Map ข้อมูลลง Template)"""
    # ถ้าไม่มี Template ให้สร้างไฟล์เปล่าๆ ขึ้นมา (กัน Error)
    if not os.path.exists(TEMPLATE_FILE):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws['A1'] = "TEMPLATE NOT FOUND (Please upload template_po.xlsx)"
    else:
        wb = openpyxl.load_workbook(TEMPLATE_FILE)
        ws = wb.active

    # --- MAPPING FIELD (แก้พิกัด Cell ตรงนี้ตามไฟล์จริง) ---
    ws['H3'] = f"เลขที่ {data['po_no']}"  # PO Number
    ws['H4'] = data['date']             # Date
    ws['B6'] = data['vendor']           # Vendor Name
    ws['B7'] = data['address']          # Address
    ws['B8'] = data['tax_id']           # Tax ID
    
    # Mapping Items
    start_row = 17 # สมมติเริ่มบรรทัด 17
    for i, item in enumerate(data['items']):
        r = start_row + i
        ws[f'B{r}'] = i + 1
        ws[f'C{r}'] = item['desc']
        ws[f'H{r}'] = float(item['qty'])
        ws[f'I{r}'] = item['unit']
        ws[f'K{r}'] = float(item['price'])
        ws[f'L{r}'] = float(item['qty']) * float(item['price'])

    # Save to Memory (RAM) ไม่ต้องเซฟลงเครื่อง Server
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
            body { font-family: 'Sarabun', sans-serif; background-color: #f3f4f6; }
            .nice-input { background: white !important; }
        </style>
    """)

    # State
    state = {
        'po_no': get_next_po_number(),
        'date': datetime.now().strftime('%Y-%m-%d'),
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
            ui.notify('เชื่อมต่อ Google Sheets ไม่ได้ (ตรวจสอบไฟล์ JSON Key)', type='negative')
            return

        ui.notify('กำลังบันทึกและสร้างเอกสาร...', type='info', spinner=True)

        # 1. Prepare Data
        full_po_no = f"{po_input.value}/{datetime.now().year + 543}"
        row_data = [
            datetime.now().strftime("%Y-%m-%d %H:%M:%S"), # Timestamp
            date_input.value,
            full_po_no,
            state['items'][0]['desc'],
            vendor_input.value,
            state['grand_total'],
            "เจ้าหน้าที่พัสดุ",
            "อนุมัติแล้ว"
        ]

        # 2. Data Object for Excel
        excel_data = {
            'po_no': full_po_no,
            'date': date_input.value,
            'vendor': vendor_input.value,
            'address': address_input.value,
            'tax_id': tax_input.value,
            'items': state['items']
        }

        try:
            # Save DB
            ws.append_row(row_data)
            
            # Generate & Download Excel
            excel_file = generate_excel_bytes(excel_data)
            filename = f"PO_{po_input.value}.xlsx"
            ui.download(excel_file.read(), filename)

            ui.notify('✅ บันทึกและดาวน์โหลดเรียบร้อย!', type='positive')
            
            # Reset
            state['po_no'] = get_next_po_number()
            po_input.value = state['po_no']
            refresh_history_table() # Refresh Table

        except Exception as e:
            ui.notify(f'เกิดข้อผิดพลาด: {e}', type='negative')

    # --- UI Layout ---
    
    # 1. HEADER AREA
    with ui.header().classes('bg-teal-700 text-white shadow-md h-16 flex items-center px-4'):
        ui.icon('description', size='md').classes('mr-2')
        with ui.column().classes('gap-0'):
            ui.label('ระบบออกใบสั่งซื้อ/จ้าง (PO System)').classes('text-lg font-bold leading-tight')
            ui.label('สำนักงานส่งเสริมเศรษฐกิจดิจิทัล').classes('text-xs text-teal-100 opacity-80')
        ui.space()
        with ui.tabs().classes('bg-transparent') as tabs:
            tab_form = ui.tab('ออกใบ PO', icon='edit')
            tab_history = ui.tab('ประวัติ/ฐานข้อมูล', icon='history')

    # 2. CONTENT AREA
    with ui.tab_panels(tabs, value=tab_form).classes('w-full bg-transparent p-0'):
        
        # === TAB 1: FORM (UX ปรับปรุงใหม่) ===
        with ui.tab_panel(tab_form).classes('p-4'):
            with ui.card().classes(STYLE_CARD):
                
                # Header Card
                with ui.row().classes('bg-teal-50 p-6 border-b border-gray-100 w-full justify-between items-center'):
                    ui.label('แบบฟอร์มขออนุมัติจัดซื้อ').classes('text-xl font-bold text-teal-800')
                    ui.label('สถานะ: รอการบันทึก').classes('text-xs bg-teal-100 text-teal-800 px-2 py-1 rounded-full')

                with ui.column().classes('p-8 w-full gap-6'):
                    
                    # Section 1: Doc Info
                    with ui.grid(columns=3).classes('w-full gap-6'):
                        with ui.column():
                            ui.label('เลขที่ PO (Auto)').classes(STYLE_LABEL)
                            po_input = ui.input().bind_value(state, 'po_no').props(PROPS_INPUT + ' readonly input-class="font-bold text-teal-700"').classes(STYLE_INPUT)
                        with ui.column():
                            ui.label('วันที่เอกสาร').classes(STYLE_LABEL)
                            date_input = ui.input().bind_value(state, 'date').props(PROPS_INPUT + ' type=date').classes(STYLE_INPUT)
                        with ui.column():
                            ui.label('เลขที่ PR อ้างอิง').classes(STYLE_LABEL)
                            ui.input().props(PROPS_INPUT).classes(STYLE_INPUT)

                    ui.separator()

                    # Section 2: Vendor
                    ui.label('1. ข้อมูลผู้ขาย').classes('text-lg font-bold text-gray-700')
                    with ui.grid(columns=2).classes('w-full gap-4'):
                        with ui.column().classes('col-span-2'):
                            ui.label('ชื่อบริษัท / ร้านค้า').classes(STYLE_LABEL)
                            vendor_input = ui.input(placeholder='ระบุชื่อผู้ขาย...').props(PROPS_INPUT).classes(STYLE_INPUT)
                        with ui.column().classes('col-span-2'):
                            ui.label('ที่อยู่').classes(STYLE_LABEL)
                            address_input = ui.textarea(placeholder='ระบุที่อยู่...').props(PROPS_INPUT + ' rows=2').classes(STYLE_INPUT)
                        with ui.column():
                            ui.label('เลขผู้เสียภาษี').classes(STYLE_LABEL)
                            tax_input = ui.input().props(PROPS_INPUT).classes(STYLE_INPUT)
                        with ui.column():
                            ui.label('รหัสงบประมาณ').classes(STYLE_LABEL)
                            ui.input().props(PROPS_INPUT).classes(STYLE_INPUT)

                    ui.separator()

                    # Section 3: Items Table (Custom UI)
                    ui.label('2. รายการสินค้า').classes('text-lg font-bold text-gray-700')
                    
                    # Table Header
                    with ui.row().classes('w-full bg-gray-100 py-2 px-4 rounded text-xs font-bold text-gray-500'):
                        ui.label('รายการ').classes('flex-grow')
                        ui.label('จำนวน').classes('w-20 text-right')
                        ui.label('หน่วย').classes('w-20 text-center')
                        ui.label('ราคา/หน่วย').classes('w-28 text-right')

                    # Dynamic Items
                    @ui.refreshable
                    def items_list():
                        for item in state['items']:
                            with ui.row().classes('w-full gap-2 mb-2 items-center'):
                                ui.input().bind_value(item, 'desc').props(PROPS_INPUT).classes('flex-grow bg-white')
                                ui.number(min=1, on_change=calculate).bind_value(item, 'qty').props(PROPS_INPUT).classes('w-20 bg-white')
                                ui.input().bind_value(item, 'unit').props(PROPS_INPUT).classes('w-20 bg-white')
                                ui.number(min=0, on_change=calculate).bind_value(item, 'price').props(PROPS_INPUT).classes('w-28 bg-white')
                        
                        ui.button('+ เพิ่มรายการ', on_click=lambda: (state['items'].append({'desc':'', 'qty':1, 'price':0, 'unit':''}), items_list.refresh())).props('flat dense color=teal icon=add').classes('mt-2')
                    
                    items_list()

                    # Section 4: Summary
                    with ui.row().classes('w-full justify-end mt-4'):
                        with ui.column().classes('w-64 bg-gray-50 p-4 rounded border border-gray-200'):
                            with ui.row().classes('w-full justify-between'):
                                ui.label('รวมเงิน').classes('text-sm')
                                label_total = ui.label('0.00').classes('font-bold')
                            with ui.row().classes('w-full justify-between text-red-500'):
                                ui.label('VAT 7%').classes('text-sm')
                                label_vat = ui.label('0.00').classes('font-bold')
                            ui.separator().classes('my-2')
                            with ui.row().classes('w-full justify-between items-center'):
                                ui.label('ยอดสุทธิ').classes('text-lg font-bold text-gray-800')
                                label_grand = ui.label('0.00').classes('text-xl font-bold text-teal-700')

                # Footer Actions
                with ui.row().classes('w-full bg-gray-50 p-4 border-t border-gray-200 justify-end gap-3'):
                    ui.button('ล้างข้อมูล', on_click=lambda: ui.notify('Clear Form')).props('flat color=grey')
                    ui.button('บันทึกและออกใบ PO', on_click=save_and_export, icon='print').props('color=teal')

        # === TAB 2: HISTORY (Database View) ===
        with ui.tab_panel(tab_history).classes('p-4'):
            with ui.card().classes(STYLE_CARD + ' h-[80vh]'):
                with ui.row().classes('p-4 border-b justify-between items-center'):
                    ui.label('ฐานข้อมูลใบสั่งซื้อ (Google Sheets)').classes('text-lg font-bold text-gray-700')
                    ui.button('รีโหลดข้อมูล', icon='refresh', on_click=lambda: refresh_history_table()).props('outline dense color=teal')

                # AG Grid Table
                grid = ui.aggrid({
                    'columnDefs': [
                        {'headerName': 'วันที่', 'field': 'วันที่', 'width': 120},
                        {'headerName': 'เลขที่ PO', 'field': 'เลขที่ PO', 'width': 150, 'checkboxSelection': True},
                        {'headerName': 'ผู้ขาย', 'field': 'ผู้ขาย', 'width': 200},
                        {'headerName': 'รายการ', 'field': 'รายการ', 'width': 250},
                        {'headerName': 'ยอดสุทธิ', 'field': 'ยอดสุทธิ', 'width': 120},
                        {'headerName': 'สถานะ', 'field': 'สถานะ', 'width': 100, 'cellStyle': {'color': 'green', 'fontWeight': 'bold'}}
                    ],
                    'rowData': [],
                    'pagination': True,
                    'paginationPageSize': 15
                }).classes('w-full h-full')

                def refresh_history_table():
                    ws = get_worksheet()
                    if ws:
                        records = ws.get_all_records()
                        grid.options['rowData'] = records
                        grid.update()
                        ui.notify(f'โหลดข้อมูลสำเร็จ {len(records)} รายการ')
                    else:
                        grid.options['rowData'] = []
                        grid.update()
                        ui.notify('ไม่พบข้อมูล หรือเชื่อมต่อไม่ได้', type='warning')

    # Initial Load
    refresh_history_table()

ui.run(title='DEPA PO System Pro', port=8080, language='th')
