from nicegui import ui
from datetime import datetime
from docxtpl import DocxTemplate
from bahttext import bahttext
import io
import os
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# --- CONFIG ---
SHEET_NAME = "DEPA_PO_SYSTEM"
CURRENT_YEAR_TAB = "PO_2569"
TEMPLATE_FILE = "template_po.docx"
JSON_KEY_FILE = "service_account.json"

# UI Styles
STYLE_INPUT = 'w-full bg-white' 
PROPS_INPUT = 'outlined dense color="teal"'
STYLE_CARD = 'w-full max-w-6xl bg-white shadow-lg rounded-lg border border-gray-200 p-0 mx-auto'
STYLE_LABEL = 'text-xs font-bold text-gray-500 uppercase mb-1'

# --- BACKEND LOGIC ---
def get_worksheet():
    # (ใช้โค้ดเดิมส่วน connect sheet)
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    if os.path.exists(JSON_KEY_FILE):
        creds = ServiceAccountCredentials.from_json_keyfile_name(JSON_KEY_FILE, scope)
        client = gspread.authorize(creds)
        try:
            sheet = client.open(SHEET_NAME)
            try:
                ws = sheet.worksheet(CURRENT_YEAR_TAB)
            except:
                ws = sheet.add_worksheet(title=CURRENT_YEAR_TAB, rows=100, cols=20)
                # Header ตาม Database ใหม่
                headers = [
                    'วันที่ออกเลข', 'เลขที่ PO/2569', 'รายการ', 'เลขที่ PR', 'ใบ PR ลงวันที่',
                    'รหัสงบประมาณ', 'งบประมาณ', 'ผู้รับจ้าง/ผู้ขาย', 'เลขประจำตัวผู้เสียภาษีอากร',
                    'จำนวนเงิน', 'วันที่อนุมัติ', 'วันที่ครบกำหนด', 'ผู้จัดทำ'
                ]
                ws.append_row(headers)
            return ws
        except Exception as e:
            print(f"Error: {e}")
            return None
    return None

def get_next_po_number():
    # (ใช้ฟังก์ชันรันเลข PO เดิมของคุณตรงนี้)
    return "ซจ.001" 

def generate_document_bytes(data, total_vars):
    if not os.path.exists(TEMPLATE_FILE):
        return None

    doc = DocxTemplate(TEMPLATE_FILE)
    
    # เตรียมตารางสินค้า
    items_context = []
    for i, item in enumerate(data['items']):
        total_line = float(item['qty']) * float(item['price'])
        items_context.append({
            'index': i + 1,
            'desc': item['desc'],           # {{ item.desc }}
            'qty': f"{float(item['qty']):,.0f}", # {{ item.qty }}
            'unit': item['unit'],
            'price': f"{float(item['price']):,.2f}", # {{ item.price }}
            'total': f"{total_line:,.2f}"
        })

    # รวม Data ทั้งหมดส่งไป Word
    context = {
        # Header
        'po_no': data['po_no'],
        'date': data['date'],
        'project_name': data['project_name'],
        'pr_no': data['pr_no'],
        'pr_date': data['pr_date'],
        'budget_code': data['budget_code'],
        
        # Quotation Info (เพิ่มใหม่)
        'quote_no': data['quote_no'],     # {{ quote_no }}
        'quote_date': data['quote_date'], # {{ quote_date }}

        # Vendor Info
        'vendor_name': data['vendor_name'],
        'vendor_address': data['vendor_address'], # {{ vendor_address }}
        'vendor_contact': data['vendor_contact'], # {{ vendor_contact }}
        'tax_id': data['tax_id'],

        # Internal Contact (เพิ่มใหม่)
        'contact_person': data['contact_person'], # {{ contact_person }}
        'contact_ext': data['contact_ext'],       # {{ contact_ext }}
        'contact_email': data['contact_email'],   # {{ contact_email }}

        # Items & Footer
        'items': items_context,
        'subtotal': f"{total_vars['subtotal']:,.2f}",
        'vat_amount': f"{total_vars['vat']:,.2f}",
        'grand_total': f"{total_vars['grand_total']:,.2f}",
        'baht_text': bahttext(total_vars['grand_total']),
        'due_date': data['due_date'],
        'approve_date': data['approve_date'],
        'preparer': data['preparer']
    }

    doc.render(context)
    output = io.BytesIO()
    doc.save(output)
    output.seek(0)
    return output

# --- UI PAGE ---
@ui.page('/')
def main_page():
    # Style
    ui.add_head_html("""
        <style>
            @import url('https://fonts.googleapis.com/css2?family=Sarabun:wght@300;400;500;600;700&display=swap');
            body { font-family: 'Sarabun', sans-serif; background-color: #f0f2f5; }
        </style>
    """)

    # State รวมตัวแปรทั้งหมด
    state = {
        'po_no': get_next_po_number(),
        'date': datetime.now().strftime('%d/%m/%Y'), # แปลงเป็น format ไทยง่ายๆ
        'project_name': '',
        'pr_no': '',
        'pr_date': '',
        'budget_code': '',
        'budget_amt': 0,
        
        # เพิ่ม: ใบเสนอราคา
        'quote_no': '',
        'quote_date': '',

        # ผู้ขาย
        'vendor_name': '',
        'vendor_address': '',
        'vendor_contact': '',
        'tax_id': '',

        # เพิ่ม: ผู้ติดต่อภายใน (DEPA)
        'contact_person': 'พบธรรม',
        'contact_ext': '1131',
        'contact_email': 'pobthum.sa@depa.or.th',

        # อื่นๆ
        'approve_date': '',
        'due_date': '',
        'preparer': 'เจ้าหน้าที่พัสดุ',
        'items': [{'desc': '', 'qty': 1, 'unit': 'งาน', 'price': 0}],
        'grand_total': 0
    }

    def calculate():
        total = sum(float(x['qty']) * float(x['price']) for x in state['items'])
        state['grand_total'] = total * 1.07 # รวม VAT
        label_total.text = f"{total:,.2f}"
        label_grand.text = f"{state['grand_total']:,.2f}"

    def save_and_export():
        # คำนวณยอด
        sub = sum(float(x['qty']) * float(x['price']) for x in state['items'])
        vat = sub * 0.07
        grand = sub + vat
        total_vars = {'subtotal': sub, 'vat': vat, 'grand_total': grand}
        
        # บันทึกลง Sheet (เฉพาะ 13 คอลัมน์หลัก)
        ws = get_worksheet()
        if ws:
            row = [
                state['date'], state['po_no'], state['project_name'],
                state['pr_no'], state['pr_date'], state['budget_code'],
                state['budget_amt'], state['vendor_name'], state['tax_id'],
                grand, state['approve_date'], state['due_date'], state['preparer']
            ]
            ws.append_row(row)
            ui.notify('บันทึกข้อมูลแล้ว')

        # สร้าง Word
        file_bytes = generate_document_bytes(state, total_vars)
        if file_bytes:
            ui.download(file_bytes.read(), f"PO_{state['po_no']}.docx")

    # --- LAYOUT UI ---
    with ui.column().classes('w-full py-6 px-4 items-center'):
        with ui.card().classes(STYLE_CARD):
            with ui.column().classes('p-6 w-full gap-4'):
                ui.label('แบบฟอร์มออกใบ PO').classes('text-xl font-bold text-teal-800')
                
                # 1. ข้อมูลหลัก + ใบเสนอราคา (เพิ่มใหม่)
                ui.label('ข้อมูลเอกสาร & ใบเสนอราคาอ้างอิง').classes(STYLE_LABEL)
                with ui.grid(columns=4).classes('w-full gap-4'):
                    ui.input('เลขที่ PO').bind_value(state, 'po_no').props(PROPS_INPUT)
                    ui.input('วันที่เอกสาร').bind_value(state, 'date').props(PROPS_INPUT)
                    # เพิ่มช่องใบเสนอราคา
                    ui.input('เลขที่ใบเสนอราคา').bind_value(state, 'quote_no').props(PROPS_INPUT) 
                    ui.input('ลงวันที่ใบเสนอราคา').bind_value(state, 'quote_date').props(PROPS_INPUT)

                with ui.grid(columns=3).classes('w-full gap-4'):
                    ui.input('เลขที่ PR').bind_value(state, 'pr_no').props(PROPS_INPUT)
                    ui.input('วันที่ PR').bind_value(state, 'pr_date').props(PROPS_INPUT)
                    ui.input('ชื่องาน/โครงการ').bind_value(state, 'project_name').props(PROPS_INPUT).classes('col-span-1')

                ui.separator()

                # 2. ผู้ขาย (เพิ่มที่อยู่/ผู้ติดต่อ)
                ui.label('ข้อมูลผู้ขาย').classes(STYLE_LABEL)
                with ui.grid(columns=2).classes('w-full gap-4'):
                    ui.input('ชื่อบริษัท/ร้านค้า').bind_value(state, 'vendor_name').props(PROPS_INPUT)
                    ui.input('เลขผู้เสียภาษี').bind_value(state, 'tax_id').props(PROPS_INPUT)
                    # เพิ่มช่องที่อยู่และผู้ติดต่อฝั่งผู้ขาย
                    ui.textarea('ที่อยู่ผู้ขาย').bind_value(state, 'vendor_address').props(PROPS_INPUT).classes('col-span-2')
                    ui.input('ชื่อผู้ติดต่อ (ผู้ขาย)').bind_value(state, 'vendor_contact').props(PROPS_INPUT)

                ui.separator()

                # 3. ข้อมูลผู้ติดต่อภายใน (เพิ่มใหม่ตามที่ขอ)
                ui.label('ผู้ประสานงาน (DEPA)').classes(STYLE_LABEL)
                with ui.grid(columns=3).classes('w-full gap-4'):
                    ui.input('ชื่อเจ้าหน้าที่').bind_value(state, 'contact_person').props(PROPS_INPUT)
                    ui.input('เบอร์ต่อ').bind_value(state, 'contact_ext').props(PROPS_INPUT)
                    ui.input('อีเมล').bind_value(state, 'contact_email').props(PROPS_INPUT)

                ui.separator()

                # 4. รายการสินค้า
                ui.label('รายการสินค้า').classes(STYLE_LABEL)
                @ui.refreshable
                def items_list():
                    for i, item in enumerate(state['items']):
                        with ui.row().classes('w-full gap-2 mb-1'):
                            ui.input('รายการ').bind_value(item, 'desc').props(PROPS_INPUT).classes('flex-grow')
                            ui.number('จำนวน', on_change=calculate).bind_value(item, 'qty').props(PROPS_INPUT).classes('w-20')
                            ui.input('หน่วย').bind_value(item, 'unit').props(PROPS_INPUT).classes('w-20')
                            ui.number('ราคา/หน่วย', on_change=calculate).bind_value(item, 'price').props(PROPS_INPUT).classes('w-28')
                    ui.button('เพิ่มรายการ', on_click=lambda: (state['items'].append({'desc':'', 'qty':1, 'price':0, 'unit':''}), items_list.refresh())).props('flat dense icon=add')
                items_list()

                # 5. สรุปยอด
                with ui.row().classes('w-full justify-end gap-4 mt-2'):
                    label_total = ui.label('0.00')
                    ui.label('+ VAT 7%')
                    label_grand = ui.label('0.00').classes('font-bold text-xl text-teal-700')

                ui.button('บันทึก & ออกใบ PO', on_click=save_and_export).props('color=teal w-full')

ui.run(title='DEPA PO System v2', port=8080)
