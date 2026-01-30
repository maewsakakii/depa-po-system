from nicegui import ui
from datetime import datetime
from bahttext import bahttext
import io
import os
import openpyxl
from openpyxl.styles import Alignment, Border, Side
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# --- CONFIG ---
SHEET_NAME = "DEPA_PO_SYSTEM"
CURRENT_YEAR_TAB = "PO_2569"
TEMPLATE_FILE = "template_po.xlsx"  # เปลี่ยนเป็น Excel
JSON_KEY_FILE = "service_account.json"

# --- UI STYLES ---
STYLE_INPUT = 'w-full bg-white'
PROPS_INPUT = 'outlined dense color="teal"'
STYLE_CARD = 'w-full max-w-6xl bg-white shadow-lg rounded-lg border border-gray-200 p-0 mx-auto'
STYLE_LABEL = 'text-xs font-bold text-gray-500 uppercase mb-1'

# --- BACKEND LOGIC ---
def get_worksheet():
    """เชื่อมต่อ Google Sheet (Log Database)"""
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    if os.path.exists(JSON_KEY_FILE):
        try:
            creds = ServiceAccountCredentials.from_json_keyfile_name(JSON_KEY_FILE, scope)
            client = gspread.authorize(creds)
            sheet = client.open(SHEET_NAME)
            try:
                ws = sheet.worksheet(CURRENT_YEAR_TAB)
            except:
                ws = sheet.add_worksheet(title=CURRENT_YEAR_TAB, rows=100, cols=20)
                headers = [
                    'Date', 'PO No', 'Project', 'PR No', 'Quote No', 
                    'Vendor', 'Tax ID', 'Grand Total', 'Preparer'
                ]
                ws.append_row(headers)
            return ws
        except Exception as e:
            print(f"GSheet Error: {e}")
            return None
    return None

def replace_text_in_cell(cell, context):
    """ฟังก์ชันช่วยแทนที่ {{ key }} ใน Cell"""
    if cell.value and isinstance(cell.value, str):
        for key, value in context.items():
            if f'{{{{ {key} }}}}' in cell.value or f'{{{{{key}}}}}' in cell.value:
                # ถ้า Cell มีแค่ Tag ให้แทนค่าลงไปเลย (เพื่อรักษา Type เช่น ตัวเลข)
                if cell.value.strip() == f'{{{{ {key} }}}}' or cell.value.strip() == f'{{{{{key}}}}}':
                    cell.value = value
                else:
                    # ถ้ามีข้อความอื่นปน ให้ replace string
                    cell.value = cell.value.replace(f'{{{{ {key} }}}}', str(value))
                    cell.value = cell.value.replace(f'{{{{{key}}}}}', str(value))

def generate_excel_bytes(data, total_vars):
    if not os.path.exists(TEMPLATE_FILE):
        ui.notify(f'ไม่พบไฟล์ Template: {TEMPLATE_FILE}', type='negative')
        return None

    wb = openpyxl.load_workbook(TEMPLATE_FILE)
    ws = wb.active

    # 1. เตรียม Context สำหรับข้อมูลทั่วไป (Header/Footer)
    # Mapping ตาม PDF Source [cite: 9, 23, 1-8]
    context = {
        'po_no': data['po_no'],
        'date': data['date'],
        'project_name': data['project_name'],
        'pr_no': data['pr_no'],
        'pr_date': data['pr_date'],
        'budget_code': data['budget_code'],
        'quote_no': data['quote_no'],       # 
        'quote_date': data['quote_date'],   # 
        'vendor_name': data['vendor_name'],
        'vendor_address': data['vendor_address'],
        'vendor_contact': data['vendor_contact'],
        'tax_id': data['tax_id'],
        'contact_person': data['contact_person'], # 
        'contact_ext': data['contact_ext'],       # 
        'contact_email': data['contact_email'],   # 
        'preparer': data['preparer'],
        'subtotal': total_vars['subtotal'],
        'vat_amount': total_vars['vat'],
        'grand_total': total_vars['grand_total'],
        'baht_text': bahttext(total_vars['grand_total'])
    }

    # 2. Loop แทนที่ค่าทั่วไปใน Sheet (ที่ไม่ใช่ตารางสินค้า)
    # ค้นหาบรรทัดที่มี {{ item.desc }} เพื่อระบุจุดเริ่มตาราง
    item_start_row = -1
    for row in ws.iter_rows():
        for cell in row:
            if cell.value and isinstance(cell.value, str) and 'item.desc' in cell.value:
                item_start_row = cell.row
                break
        if item_start_row != -1:
            break
            
    # แทนค่า General Tags ทั้งหมด
    for row in ws.iter_rows():
        for cell in row:
            replace_text_in_cell(cell, context)

    # 3. จัดการตารางสินค้า (Items)
    if item_start_row != -1:
        # ลบ Placeholder แถวแรกของ Item ออกก่อนแล้วค่อยเขียนทับ
        # (ในทางปฏิบัติ เราจะเขียนทับแถว item_start_row ไปเรื่อยๆ)
        
        current_row = item_start_row
        for i, item in enumerate(data['items']):
            total_line = float(item['qty']) * float(item['price'])
            
            # Map ข้อมูลลงแต่ละคอลัมน์ (ต้องแก้ Column Index A,B,C ให้ตรงกับ Template จริง)
            # สมมติ: A=ลำดับ, B=รายการ, H=จำนวน, I=หน่วย, J=ราคา, K=รวม
            # (คุณต้องปรับ Column index ให้ตรงกับไฟล์ Excel ของคุณ)
            
            # วิธีที่ยืดหยุ่นกว่า: ค้นหาว่า column ไหนมี tag อะไรในแถว item_start_row เดิม
            # แต่เพื่อความง่าย เราจะใช้การ hardcode column ตาม Template มาตรฐาน หรือเขียนทับ
            
            ws[f'A{current_row}'] = i + 1
            ws[f'B{current_row}'] = item['desc']      # 
            ws[f'H{current_row}'] = float(item['qty']) # 
            ws[f'I{current_row}'] = item['unit']
            ws[f'J{current_row}'] = float(item['price']) # 
            ws[f'K{current_row}'] = total_line

            # จัด Format ตัวเลข
            ws[f'J{current_row}'].number_format = '#,##0.00'
            ws[f'K{current_row}'].number_format = '#,##0.00'
            
            current_row += 1

    # Save
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output

# --- UI PAGE ---
@ui.page('/')
def main_page():
    ui.add_head_html("""
        <style>
            @import url('https://fonts.googleapis.com/css2?family=Sarabun:wght@300;400;600&display=swap');
            body { font-family: 'Sarabun', sans-serif; background-color: #f3f4f6; }
        </style>
    """)

    # State Init
    state = {
        'po_no': 'ซจ.001',
        'date': datetime.now().strftime('%d/%m/%Y'),
        'project_name': '',
        'pr_no': '',
        'pr_date': '',
        'budget_code': '',
        'quote_no': '',    # 
        'quote_date': '',  # 
        'vendor_name': '',
        'vendor_address': '',
        'vendor_contact': '',
        'tax_id': '',
        'contact_person': 'พบธรรม',             # 
        'contact_ext': '1131',                  # 
        'contact_email': 'pobthum.sa@depa.or.th', # 
        'preparer': 'เจ้าหน้าที่พัสดุ',
        'items': [{'desc': '', 'qty': 1, 'unit': 'งาน', 'price': 0}],
        'grand_total': 0
    }

    def calculate():
        total = sum(float(x['qty']) * float(x['price']) for x in state['items'])
        state['grand_total'] = total * 1.07
        label_total.text = f"{total:,.2f}"
        label_grand.text = f"{state['grand_total']:,.2f}"

    async def save_and_export():
        # คำนวณ
        sub = sum(float(x['qty']) * float(x['price']) for x in state['items'])
        vat = sub * 0.07
        grand = sub + vat
        total_vars = {'subtotal': sub, 'vat': vat, 'grand_total': grand}
        
        # Save Log to Sheet
        ws = get_worksheet()
        if ws:
            ws.append_row([
                state['date'], state['po_no'], state['project_name'], 
                state['pr_no'], state['quote_no'], state['vendor_name'], 
                state['tax_id'], grand, state['contact_person']
            ])
            ui.notify('บันทึก Log แล้ว', type='positive')

        # Generate Excel
        file_bytes = generate_excel_bytes(state, total_vars)
        if file_bytes:
            ui.download(file_bytes.read(), f"PO_{state['po_no']}.xlsx")

    # --- LAYOUT ---
    with ui.column().classes('w-full py-8 px-4 items-center'):
        with ui.card().classes(STYLE_CARD):
            # Header Title
            with ui.row().classes('w-full bg-teal-800 p-4 rounded-t-lg items-center'):
                ui.label('ระบบออกใบสั่งซื้อ (Excel Template)').classes('text-white text-xl font-bold')

            with ui.column().classes('p-6 w-full gap-4'):
                
                # 1. Document Info
                ui.label('ข้อมูลเอกสาร & ใบเสนอราคา').classes(STYLE_LABEL)
                with ui.grid(columns=4).classes('w-full gap-4'):
                    ui.input('เลขที่ PO').bind_value(state, 'po_no').props(PROPS_INPUT)
                    ui.input('วันที่เอกสาร').bind_value(state, 'date').props(PROPS_INPUT)
                    ui.input('เลขที่ใบเสนอราคา').bind_value(state, 'quote_no').props(PROPS_INPUT) # 
                    ui.input('ลงวันที่ใบเสนอราคา').bind_value(state, 'quote_date').props(PROPS_INPUT)

                with ui.grid(columns=3).classes('w-full gap-4'):
                    ui.input('เลขที่ PR').bind_value(state, 'pr_no').props(PROPS_INPUT)
                    ui.input('รหัสงบประมาณ').bind_value(state, 'budget_code').props(PROPS_INPUT) # 
                    ui.input('ชื่องาน/โครงการ').bind_value(state, 'project_name').props(PROPS_INPUT)

                ui.separator()

                # 2. Vendor Info
                ui.label('ข้อมูลผู้ขาย').classes(STYLE_LABEL) # 
                with ui.grid(columns=2).classes('w-full gap-4'):
                    ui.input('ชื่อบริษัท/ร้านค้า').bind_value(state, 'vendor_name').props(PROPS_INPUT)
                    ui.input('เลขผู้เสียภาษี').bind_value(state, 'tax_id').props(PROPS_INPUT) # 
                    ui.textarea('ที่อยู่').bind_value(state, 'vendor_address').props(PROPS_INPUT).classes('col-span-2') # 
                    ui.input('ผู้ติดต่อ (Vendor)').bind_value(state, 'vendor_contact').props(PROPS_INPUT).classes('col-span-2') # 

                ui.separator()

                # 3. Internal Contact
                ui.label('ผู้ประสานงาน (DEPA)').classes(STYLE_LABEL) # 
                with ui.grid(columns=3).classes('w-full gap-4'):
                    ui.input('ชื่อเจ้าหน้าที่').bind_value(state, 'contact_person').props(PROPS_INPUT)
                    ui.input('เบอร์ต่อ').bind_value(state, 'contact_ext').props(PROPS_INPUT)
                    ui.input('อีเมล').bind_value(state, 'contact_email').props(PROPS_INPUT)

                ui.separator()

                # 4. Items
                ui.label('รายการสินค้า').classes(STYLE_LABEL) # [cite: 13]
                @ui.refreshable
                def items_list():
                    # Header row
                    with ui.row().classes('w-full gap-2 px-2'):
                        ui.label('รายการ').classes('flex-grow text-xs text-gray-500')
                        ui.label('จำนวน').classes('w-20 text-xs text-gray-500')
                        ui.label('หน่วย').classes('w-20 text-xs text-gray-500')
                        ui.label('ราคา/หน่วย').classes('w-28 text-xs text-gray-500')

                    for i, item in enumerate(state['items']):
                        with ui.row().classes('w-full gap-2 mb-1 items-start'):
                            ui.textarea().bind_value(item, 'desc').props('outlined dense rows=1').classes('flex-grow')
                            ui.number(on_change=calculate).bind_value(item, 'qty').props(PROPS_INPUT).classes('w-20')
                            ui.input().bind_value(item, 'unit').props(PROPS_INPUT).classes('w-20')
                            ui.number(on_change=calculate).bind_value(item, 'price').props(PROPS_INPUT).classes('w-28')
                            
                    ui.button('เพิ่มรายการ', on_click=lambda: (state['items'].append({'desc':'', 'qty':1, 'price':0, 'unit':''}), items_list.refresh())).props('flat dense icon=add color=teal')
                items_list()

                # 5. Footer Summary
                with ui.column().classes('w-full items-end mt-4 p-4 bg-gray-50 rounded border'):
                    with ui.row().classes('gap-4 items-center'):
                        ui.label('รวมเงิน:').classes('font-bold')
                        label_total = ui.label('0.00')
                    with ui.row().classes('gap-4 items-center'):
                        ui.label('VAT 7%:').classes('font-bold')
                        ui.label('(Auto Calc)')
                    with ui.row().classes('gap-4 items-center'):
                        ui.label('ยอดสุทธิ:').classes('text-xl font-bold text-teal-800')
                        label_grand = ui.label('0.00').classes('text-xl font-bold text-teal-800')

                ui.button('สร้างไฟล์ Excel PO', on_click=save_and_export).props('unelevated color=teal icon=file_download w-full size=lg')

ui.run(title='DEPA PO System (Excel)', port=8080)
