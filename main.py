from nicegui import ui
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime
import os
import json

# --- CONFIGURATION ---
SHEET_NAME = "DEPA_PO_SYSTEM"  # ชื่อไฟล์ Google Sheets ต้องตรงเป๊ะ
CURRENT_YEAR_TAB = "PO_2569"  # ชื่อ Tab ที่จะทำงานด้วย

# โหลด Key จาก Environment Variable (สำหรับ Cloud) หรือไฟล์ local
# วิธีใช้บนเครื่อง: วางไฟล์ json ไว้ที่เดียวกับโค้ดแล้วแก้ชื่อไฟล์ตรงนี้
JSON_KEY_FILE = "depa-po-bot-4a8b48470390.json"


# --- GOOGLE SHEETS CONNECTION ---
def get_client():
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]

    # 1. ลองอ่านจาก Environment Variable (สำหรับ Render/Railway)
    json_env = os.getenv("GOOGLE_JSON_KEY")
    if json_env:
        creds_dict = json.loads(json_env)
        creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
    # 2. ถ้าไม่มี ให้ลองอ่านจากไฟล์ (สำหรับรันในเครื่อง)
    elif os.path.exists(JSON_KEY_FILE):
        creds = ServiceAccountCredentials.from_json_keyfile_name(JSON_KEY_FILE, scope)
    else:
        return None

    return gspread.authorize(creds)


def get_worksheet():
    client = get_client()
    if not client: return None
    try:
        sheet = client.open(SHEET_NAME)
        # ลองเปิด Tab ปีปัจจุบัน ถ้าไม่มีให้สร้าง
        try:
            ws = sheet.worksheet(CURRENT_YEAR_TAB)
        except:
            ws = sheet.add_worksheet(title=CURRENT_YEAR_TAB, rows=100, cols=20)
            # สร้างหัวตารางให้อัตโนมัติ
            ws.append_row(
                ['วันที่', 'เลขที่ PO', 'รายการ', 'เลขที่ PR', 'รหัสงบ', 'ผู้ขาย', 'เลขภาษี', 'ยอดสุทธิ', 'กำหนดส่ง',
                 'ผู้จัดทำ', 'สถานะ'])
        return ws
    except Exception as e:
        print(f"Sheet Error: {e}")
        return None


def get_next_po_number():
    ws = get_worksheet()
    if not ws: return "ซจ.001"

    try:
        # ดึงข้อมูลคอลัมน์ B (เลขที่ PO)
        po_col = ws.col_values(2)
        if len(po_col) <= 1: return "ซจ.001"  # มีแค่หัวตาราง

        last_po = po_col[-1]  # ตัวสุดท้าย
        if "ซจ." in last_po:
            num = int(last_po.replace("ซจ.", "").split("/")[0])  # เผื่อมี /2569 ต่อท้าย
            return f"ซจ.{num + 1:03d}"
        return "ซจ.001"
    except:
        return "ซจ.001"


# --- UI Application ---
@ui.page('/')
def main_page():
    # --- STYLES (Tailwind classes) ---
    input_style = "w-full border border-gray-300 rounded px-3 py-2 focus:outline-none focus:border-teal-500 bg-white"
    label_style = "block text-xs font-bold text-gray-500 uppercase mb-1"

    # --- STATE ---
    state = {
        'po_no': get_next_po_number(),
        'date': datetime.now().strftime('%Y-%m-%d'),
        'items': [{'desc': '', 'qty': 1, 'unit': 'งาน', 'price': 0}]
    }

    def calculate_total():
        total = sum(float(x['qty']) * float(x['price']) for x in state['items'])
        vat = total * 0.07
        grand = total + vat
        total_lbl.text = f"{grand:,.2f}"
        vat_lbl.text = f"{vat:,.2f}"
        state['grand_total'] = grand

    def save_data():
        ws = get_worksheet()
        if not ws:
            ui.notify('ไม่สามารถเชื่อมต่อ Google Sheets ได้', type='negative')
            return

        ui.notify('กำลังบันทึก...', type='info', spinner=True)

        # เตรียมข้อมูล (Format ให้ตรงกับหัวตาราง)
        row = [
            date_input.value,
            f"{po_input.value}/{datetime.now().year + 543}",  # ใส่ปีงบประมาณต่อท้าย
            state['items'][0]['desc'],  # เอาแค่รายการแรกเป็นหัวเรื่อง
            pr_input.value,
            budget_input.value,
            vendor_input.value,
            tax_input.value,
            state.get('grand_total', 0),
            due_input.value,
            "เจ้าหน้าที่พัสดุ",  # Hardcode หรือทำ login เพิ่ม
            "รออนุมัติ"
        ]

        try:
            ws.append_row(row)
            ui.notify('บันทึกข้อมูลเรียบร้อย!', type='positive')

            # Reset
            state['po_no'] = get_next_po_number()
            po_input.value = state['po_no']
            vendor_input.value = ""
            state['items'] = [{'desc': '', 'qty': 1, 'unit': 'งาน', 'price': 0}]
            items_container.refresh()

        except Exception as e:
            ui.notify(f'เกิดข้อผิดพลาด: {e}', type='negative')

    # --- LAYOUT (เลียนแบบ HTML Mockup) ---
    with ui.column().classes('w-full min-h-screen bg-gray-100 items-center py-10 px-4'):

        # Main Card
        with ui.card().classes(
            'w-full max-w-5xl bg-white shadow-xl rounded-lg overflow-hidden border border-gray-200 p-0'):
            # Header
            with ui.row().classes('w-full bg-teal-700 text-white p-6 justify-between items-center'):
                with ui.column().classes('gap-0'):
                    ui.label('ระบบทะเบียนคุมและออกใบ PO').classes('text-2xl font-bold')
                    ui.label('สำนักงานส่งเสริมเศรษฐกิจดิจิทัล (depa)').classes('text-teal-100 text-sm opacity-80')
                with ui.column().classes('items-end gap-1'):
                    ui.label('สถานะ: พร้อมใช้งาน').classes('text-xs bg-teal-800 py-1 px-3 rounded-full')
                    ui.label(f'Connect to: {CURRENT_YEAR_TAB}').classes('text-sm')

            # Form Content
            with ui.column().classes('w-full p-8 gap-8'):
                # Section 1: Top Info
                with ui.grid(columns=3).classes('w-full gap-6 bg-gray-50 p-4 rounded-lg border border-gray-100'):
                    with ui.column().classes('w-full'):
                        ui.label('เลขที่ PO (AUTO)').classes(label_style)
                        po_input = ui.input(value=state['po_no']).props('readonly').classes(
                            input_style + " text-teal-700 font-bold")
                    with ui.column().classes('w-full'):
                        ui.label('วันที่เอกสาร').classes(label_style)
                        date_input = ui.input(value=state['date']).props('type=date').classes(input_style)
                    with ui.column().classes('w-full'):
                        ui.label('เลขที่ PR อ้างอิง').classes(label_style)
                        pr_input = ui.input(placeholder='เช่น PR-69-00xxx').classes(input_style)

                # Section 2: Vendor
                with ui.column().classes('w-full gap-4'):
                    ui.label('1. ข้อมูลผู้ขาย / คู่สัญญา').classes(
                        'text-lg font-bold text-gray-700 border-b pb-2 w-full')
                    with ui.grid(columns=2).classes('w-full gap-6'):
                        with ui.column().classes('col-span-2'):
                            ui.label('ชื่อบริษัท / ร้านค้า').classes('font-medium text-sm text-gray-700')
                            vendor_input = ui.input().classes(input_style)
                        with ui.column():
                            ui.label('เลขประจำตัวผู้เสียภาษี').classes('font-medium text-sm text-gray-700')
                            tax_input = ui.input().classes(input_style)
                        with ui.column():
                            ui.label('รหัสงบประมาณ').classes('font-medium text-sm text-gray-700')
                            budget_input = ui.input().classes(input_style)
                        with ui.column():
                            ui.label('กำหนดส่งมอบ').classes('font-medium text-sm text-gray-700')
                            due_input = ui.input().props('type=date').classes(input_style)

                # Section 3: Items Table
                with ui.column().classes('w-full gap-4'):
                    ui.label('2. รายการสินค้า').classes('text-lg font-bold text-gray-700 border-b pb-2 w-full')

                    # Table Header
                    with ui.row().classes('w-full bg-gray-100 text-gray-600 text-sm py-2 px-4 rounded-t'):
                        ui.label('รายการ').classes('flex-grow')
                        ui.label('จำนวน').classes('w-20 text-right')
                        ui.label('หน่วย').classes('w-20 text-center')
                        ui.label('ราคา/หน่วย').classes('w-24 text-right')

                    # Items Rows
                    @ui.refreshable
                    def items_container():
                        for item in state['items']:
                            with ui.row().classes('w-full items-center gap-2 border-b border-gray-100 py-2'):
                                ui.input().bind_value(item, 'desc').classes(
                                    'flex-grow bg-transparent focus:bg-white border-b border-transparent focus:border-teal-500')
                                ui.number(min=1, on_change=calculate_total).bind_value(item, 'qty').classes('w-20')
                                ui.input().bind_value(item, 'unit').classes('w-20 text-center')
                                ui.number(min=0, on_change=calculate_total).bind_value(item, 'price').classes(
                                    'w-24 text-right')

                        ui.button('+ เพิ่มรายการ', on_click=lambda: (
                            state['items'].append({'desc': '', 'qty': 1, 'price': 0, 'unit': ''}),
                            items_container.refresh())).classes('text-teal-600 font-bold w-full mt-2')

                    items_container()

                    # Summary
                    with ui.row().classes('w-full justify-end mt-4'):
                        with ui.column().classes('w-1/3 bg-gray-50 p-4 rounded-lg border border-gray-200'):
                            with ui.row().classes('w-full justify-between mb-2'):
                                ui.label('ภาษี (7%)').classes('text-gray-600')
                                vat_lbl = ui.label('0.00').classes('font-medium text-red-500')
                            ui.separator()
                            with ui.row().classes('w-full justify-between items-center mt-2'):
                                ui.label('ยอดสุทธิ').classes('text-lg font-bold text-gray-800')
                                total_lbl = ui.label('0.00').classes('text-2xl font-bold text-teal-700')

            # Footer
            with ui.row().classes('w-full bg-gray-50 p-6 border-t border-gray-200 justify-end gap-3'):
                ui.button('บันทึกข้อมูล', on_click=save_data).classes(
                    'bg-teal-600 text-white px-8 py-2 rounded shadow-md hover:bg-teal-700')


ui.run(title='DEPA PO System', port=8080)