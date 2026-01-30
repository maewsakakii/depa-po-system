from nicegui import ui
from datetime import datetime
import os
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from docxtpl import DocxTemplate # สำหรับ Word
# from xlxtpl.writerx import BookWriter # ถ้าใช้ Excel ให้เปิดตัวนี้แทน
import io

# --- CONFIG ---
SHEET_NAME = "DEPA_PO_SYSTEM"
CURRENT_YEAR_TAB = "PO_2569" # ตั้งชื่อ Tab ตามปีงบประมาณ
TEMPLATE_FILE = "template_po.docx" # หรือ .xlsx
JSON_KEY_FILE = "service_account.json"

# --- HELPER: แปลงเงินบาทเป็นตัวอักษร ---
from bahttext import bahttext # ต้อง pip install bahttext ก่อน (หรือใช้ฟังก์ชันเขียนเอง)

def get_worksheet():
    """เชื่อมต่อ Sheet และสร้าง Header ตามภาพใหม่"""
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
                # สร้าง Header ตามภาพ Database ใหม่ (A-M)
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

def generate_document_bytes(data, total_vars):
    """สร้างไฟล์เอกสาร (Word) โดยแทนที่ Tag {{...}}"""
    if not os.path.exists(TEMPLATE_FILE):
        return None

    doc = DocxTemplate(TEMPLATE_FILE)
    
    # 1. เตรียมข้อมูลสินค้า
    items_context = []
    for i, item in enumerate(data['items']):
        total_line = item['qty'] * item['price']
        items_context.append({
            'index': i + 1,
            'desc': item['desc'],
            'qty': f"{item['qty']:,.0f}",
            'unit': item['unit'],
            'price': f"{item['price']:,.2f}",
            'total': f"{total_line:,.2f}"
        })

    # 2. เตรียม Context (Tags)
    context = {
        # Header
        'po_no': data['po_no'],
        'date': data['date'],         # วันที่ออกเลข
        'pr_no': data['pr_no'],
        'pr_date': data['pr_date'],
        'budget_code': data['budget_code'],
        'project_name': data['project_name'], # รายการหลัก (Col C)

        # Vendor
        'vendor_name': data['vendor_name'],
        'tax_id': data['tax_id'],
        'vendor_address': data['vendor_address'],   # ไม่มีใน DB แต่ใส่ในเอกสาร
        'vendor_contact': data['vendor_contact'],   # ไม่มีใน DB แต่ใส่ในเอกสาร

        # Items
        'items': items_context,

        # Footer / Totals
        'subtotal': f"{total_vars['subtotal']:,.2f}",
        'vat_amount': f"{total_vars['vat']:,.2f}",
        'grand_total': f"{total_vars['grand_total']:,.2f}",
        'baht_text': bahttext(total_vars['grand_total']), # แปลงตัวเลขเป็นบาทถ้วน
        'due_date': data['due_date'],
        'approve_date': data['approve_date'],
        'preparer': data['preparer']
    }

    doc.render(context)
    
    output = io.BytesIO()
    doc.save(output)
    output.seek(0)
    return output

# --- UI LOGIC (เฉพาะส่วน Save) ---
def save_data(state):
    ws = get_worksheet()
    
    # คำนวณยอดเงิน
    subtotal = sum(item['qty'] * item['price'] for item in state['items'])
    vat = subtotal * 0.07
    grand_total = subtotal + vat
    total_vars = {'subtotal': subtotal, 'vat': vat, 'grand_total': grand_total}

    # ข้อมูลที่จะใช้ (State จากหน้าจอ UI)
    # สมมติ state มี field ครบตามฟอร์ม
    data = {
        'date': state['date'],                 # A
        'po_no': state['po_no'],               # B
        'project_name': state['project_name'], # C (ชื่อโครงการ/งานจ้าง)
        'pr_no': state['pr_no'],               # D
        'pr_date': state['pr_date'],           # E
        'budget_code': state['budget_code'],   # F
        'budget_amount': state['budget_amt'],  # G (งบที่มี)
        'vendor_name': state['vendor_name'],   # H
        'tax_id': state['tax_id'],             # I
        # J คือ grand_total
        'approve_date': state['approve_date'], # K
        'due_date': state['due_date'],         # L
        'preparer': state['preparer'],         # M
        
        # ส่วนเสริม (ไม่ลง DB แต่ลง Template)
        'vendor_address': state['vendor_address'],
        'vendor_contact': state['vendor_contact'],
        'items': state['items']
    }

    # 1. เรียงข้อมูลลง Database (A - M)
    row_to_append = [
        data['date'],           # A: วันที่ออกเลข
        data['po_no'],          # B: เลขที่ PO
        data['project_name'],   # C: รายการ (ชื่องาน)
        data['pr_no'],          # D: เลขที่ PR
        data['pr_date'],        # E: ใบ PR ลงวันที่
        data['budget_code'],    # F: รหัสงบ
        data['budget_amount'],  # G: งบประมาณ
        data['vendor_name'],    # H: ผู้ขาย
        data['tax_id'],         # I: เลขภาษี
        total_vars['grand_total'], # J: จำนวนเงิน (ยอดสุทธิ)
        data['approve_date'],   # K: วันที่อนุมัติ
        data['due_date'],       # L: วันที่ครบกำหนด
        data['preparer']        # M: ผู้จัดทำ
    ]

    # บันทึกลง Sheet
    ws.append_row(row_to_append)
    ui.notify('บันทึกข้อมูลเรียบร้อย!')

    # สร้างและโหลดไฟล์เอกสาร
    file_bytes = generate_document_bytes(data, total_vars)
    if file_bytes:
        ui.download(file_bytes.read(), f"{data['po_no']}.docx")
