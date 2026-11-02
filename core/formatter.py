"""
Officia Formatter Pro v3 - نظام التنسيق المحاسبي الاحترافي
متخصص في التقارير المالية والأصول الثابتة بمستوى Excel احترافي
"""

import os
from datetime import datetime
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn, nsdecls
from docx.oxml import parse_xml
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, numbers
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, PieChart, LineChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.formatting.rule import CellIsRule
from openpyxl.worksheet.table import Table, TableStyleInfo

# ==========================================
# نظام الألوان المحاسبي الاحترافي
# ==========================================
ACCOUNTING_PALETTE = {
    'header_dark': '1F4E79',
    'header_mid': '2E75B6',
    'header_light': '9DC3E6',
    'section_bg': 'D6E4F0',
    'total_bg': 'E2EFDA',
    'subtotal_bg': 'FCE4D6',
    'negative': 'FF0000',
    'positive': '00B050',
    'warning': 'FFC000',
    'border': 'B4C7E7',
    'text_dark': '1F1F1F',
    'text_mid': '404040',
    'text_light': '808080',
    'white': 'FFFFFF',
    'row_even': 'F2F2F2',
    'row_odd': 'FFFFFF',
}

ACCOUNTING_FONTS = {
    'title': 'Calibri',
    'heading': 'Calibri',
    'body': 'Calibri',
    'numbers': 'Calibri',
    'arabic': 'Traditional Arabic',
}

FONT_SIZES = {
    'mega_title': 18,
    'title': 16,
    'heading': 14,
    'subheading': 12,
    'body': 11,
    'small': 10,
    'caption': 9,
}

# ==========================================
# دوال مساعدة عامة
# ==========================================
def set_cell_shading(cell, color_hex):
    """تعيين لون خلفية خلية"""
    from openpyxl.styles import PatternFill
    fill = PatternFill(start_color=color_hex, end_color=color_hex, fill_type='solid')
    cell.fill = fill

def get_unique_filename(directory, base_name, extension):
    """إنشاء اسم ملف فريد"""
    counter = 1
    filename = f"{base_name}.{extension}"
    while os.path.exists(os.path.join(directory, filename)):
        filename = f"{base_name}_{counter}.{extension}"
        counter += 1
    return filename

# ==========================================
# دوال Word الاحترافية
# ==========================================
def create_professional_letter(
    recipient, subject, content_paragraphs,
    table_data=None, table_headers=None, table_caption=None,
    closing_note=None, closing_phrase=None,
    sender_name=None, sender_title=None, sender_organization=None,
    reference_number=None, attachments=None,
    output_dir=None, filename=None
):
    """إنشاء خطاب رسمي بمستوى احترافي"""
    from core.config import WORKSPACE_DIR
    
    doc = Document()
    section = doc.sections[0]
    section.top_margin = Cm(2.5)
    section.bottom_margin = Cm(2.5)
    section.left_margin = Cm(2.5)
    section.right_margin = Cm(2.5)
    
    style = doc.styles['Normal']
    style.font.name = ACCOUNTING_FONTS['body']
    style.font.size = Pt(FONT_SIZES['body'])
    
    if sender_organization:
        org_para = doc.add_paragraph()
        org_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        org_run = org_para.add_run(sender_organization)
        org_run.font.size = Pt(FONT_SIZES['title'])
        org_run.font.bold = True
        org_run.font.color.rgb = RGBColor.from_string(ACCOUNTING_PALETTE['header_dark'])
    
    para = doc.add_paragraph()
    para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    pPr = para._p.get_or_add_pPr()
    pBdr = parse_xml(f'<w:pBdr {nsdecls("w")}><w:bottom w:val="single" w:sz="12" w:space="1" w:color="{ACCOUNTING_PALETTE["header_dark"]}"/></w:pBdr>')
    pPr.append(pBdr)
    
    title = doc.add_heading(subject, level=0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for run in title.runs:
        run.font.size = Pt(FONT_SIZES['mega_title'])
        run.font.color.rgb = RGBColor.from_string(ACCOUNTING_PALETTE['header_dark'])
    
    info_table = doc.add_table(rows=1, cols=2)
    date_cell = info_table.cell(0, 0)
    date_para = date_cell.paragraphs[0]
    date_para.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    date_run = date_para.add_run(f"التاريخ: {datetime.now().strftime('%Y/%m/%d')}")
    date_run.font.bold = True
    
    ref_cell = info_table.cell(0, 1)
    ref_para = ref_cell.paragraphs[0]
    ref_para.alignment = WD_ALIGN_PARAGRAPH.LEFT
    ref_run = ref_para.add_run(f"رقم: {reference_number or datetime.now().strftime('%Y%m%d')}")
    ref_run.font.bold = True
    
    doc.add_paragraph()
    
    to_para = doc.add_paragraph()
    to_para.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    to_para.add_run("السادة/ ").font.bold = True
    to_para.add_run(recipient).font.bold = True
    
    subj_para = doc.add_paragraph()
    subj_para.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    subj_para.add_run("الموضوع: ").font.bold = True
    subj_value = subj_para.add_run(subject)
    subj_value.font.color.rgb = RGBColor.from_string(ACCOUNTING_PALETTE['header_mid'])
    
    doc.add_paragraph()
    
    greeting = doc.add_paragraph()
    greeting.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    greeting.add_run("تحية طيبة وبعد،،،").font.bold = True
    
    doc.add_paragraph()
    
    intro = doc.add_paragraph()
    intro.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    intro.paragraph_format.first_line_indent = Cm(1.5)
    intro.add_run("بالإشارة إلى الموضوع أعلاه، نتقدم إليكم بهذا الخطاب.")
    
    for i, para_text in enumerate(content_paragraphs):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        p.paragraph_format.first_line_indent = Cm(1.5)
        p.paragraph_format.line_spacing = 1.5
        text = f"{i+1}. {para_text}" if len(content_paragraphs) > 1 else para_text
        p.add_run(text)
    
    doc.add_paragraph()
    
    if table_data and table_headers:
        table = doc.add_table(rows=1, cols=len(table_headers))
        table.style = 'Table Grid'
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        
        for i, header in enumerate(table_headers):
            cell = table.rows[0].cells[i]
            cell.text = ''
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = p.add_run(header)
            run.font.bold = True
            run.font.color.rgb = RGBColor.from_string(ACCOUNTING_PALETTE['white'])
            run.font.size = Pt(FONT_SIZES['body'])
            set_cell_shading(cell, ACCOUNTING_PALETTE['header_dark'])
        
        for row_idx, row_data in enumerate(table_data):
            row_cells = table.add_row().cells
            bg_color = ACCOUNTING_PALETTE['row_even'] if row_idx % 2 == 0 else ACCOUNTING_PALETTE['row_odd']
            for col_idx, cell_data in enumerate(row_data):
                cell = row_cells[col_idx]
                cell.text = ''
                p = cell.paragraphs[0]
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                run = p.add_run(str(cell_data))
                run.font.size = Pt(FONT_SIZES['body'])
                set_cell_shading(cell, bg_color)
        
        doc.add_paragraph()
    
    closing = closing_phrase or "وتفضلوا بقبول فائق الاحترام والتقدير،،"
    closing_para = doc.add_paragraph()
    closing_para.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    closing_para.add_run(closing).font.bold = True
    
    if closing_note:
        doc.add_paragraph()
        note_para = doc.add_paragraph()
        note_para.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        note_icon = note_para.add_run("⚠ تنبيه: ")
        note_icon.font.bold = True
        note_icon.font.color.rgb = RGBColor.from_string(ACCOUNTING_PALETTE['negative'])
        note_text = note_para.add_run(closing_note)
        note_text.font.color.rgb = RGBColor.from_string(ACCOUNTING_PALETTE['negative'])
    
    doc.add_paragraph()
    doc.add_paragraph()
    if sender_name:
        sig_para = doc.add_paragraph()
        sig_para.alignment = WD_ALIGN_PARAGRAPH.LEFT
        sig_para.add_run(sender_name).font.bold = True
        if sender_title:
            sig_para.add_run('\n' + sender_title).font.size = Pt(FONT_SIZES['small'])
    
    save_dir = output_dir or WORKSPACE_DIR
    os.makedirs(save_dir, exist_ok=True)
    
    if filename is None:
        safe_subject = subject[:30].replace(' ', '_').replace('/', '-')
        filename = f"خطاب_{safe_subject}"
    
    filename = get_unique_filename(save_dir, filename, 'docx')
    output_path = os.path.join(save_dir, filename)
    
    doc.save(output_path)
    return output_path

# --- نهاية الجزء 1 من 2 ---
for col in [5, 6, 7, 8, 10, 11] + list(range(12, 29)):
                cell = ws.cell(row=current_row, column=col)
                cell.number_format = '#,##0.00'
                cell.alignment = Alignment(horizontal='center')
            
            # خلفية بديلة
            bg_color = ACCOUNTING_PALETTE['row_even'] if current_row % 2 == 0 else ACCOUNTING_PALETTE['row_odd']
            for col in range(1, 29):
                set_cell_shading(ws.cell(row=current_row, column=col), bg_color)
            
            current_row += 1
        
        # صف الإجمالي
        ws.cell(row=current_row, column=2, value="Total")
        ws.cell(row=current_row, column=2).font = Font(bold=True, color=ACCOUNTING_PALETTE['header_dark'])
        ws.cell(row=current_row, column=5, value=section_total_value)
        ws.cell(row=current_row, column=11, value=section_total_depr)
        
        for col in [5, 11] + list(range(12, 29)):
            cell = ws.cell(row=current_row, column=col)
            cell.number_format = '#,##0.00'
            cell.font = Font(bold=True)
            set_cell_shading(cell, ACCOUNTING_PALETTE['total_bg'])
        
        # ضبط عرض الأعمدة
        ws.column_dimensions['A'].width = 6
        ws.column_dimensions['B'].width = 25
        ws.column_dimensions['C'].width = 35
        ws.column_dimensions['D'].width = 10
        for col in range(5, 29):
            ws.column_dimensions[get_column_letter(col)].width = 12
        
        # تجميد الصفوف
        ws.freeze_panes = f'A{header_row + 1}'
        
        # فلتر تلقائي
        ws.auto_filter.ref = f'A{header_row}:{get_column_letter(len(detail_headers))}{current_row - 1}'
    
    # ==========================================
    # الحفظ
    # ==========================================
    save_dir = output_dir or WORKSPACE_DIR
    os.makedirs(save_dir, exist_ok=True)
    
    if filename is None:
        filename = f"Fixed_Asset_{company_name}_{year}"
    
    filename = get_unique_filename(save_dir, filename, 'xlsx')
    output_path = os.path.join(save_dir, filename)
    
    wb.save(output_path)
    return output_path


# ==========================================
# دالة Excel متعددة الأغراض
# ==========================================
def create_multi_sheet_excel(
    sheets_config,
    output_dir=None,
    filename=None
):
    """
    إنشاء ملف Excel متعدد الأوراق بتنسيق احترافي
    
    sheets_config: قائمة من القواميس، كل ورقة تحتوي على:
    - name: اسم الورقة
    - title: عنوان الورقة (اختياري)
    - headers: قائمة العناوين
    - data: قائمة البيانات
    - has_formulas: هل تحتوي على معادلات (اختياري)
    """
    from core.config import WORKSPACE_DIR
    
    wb = openpyxl.Workbook()
    
    for idx, sheet_cfg in enumerate(sheets_config):
        if idx == 0:
            ws = wb.active
            ws.title = sheet_cfg['name'][:31]
        else:
            ws = wb.create_sheet(title=sheet_cfg['name'][:31])
        
        ws.sheet_view.rightToLeft = True
        
        # العنوان
        if sheet_cfg.get('title'):
            ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(sheet_cfg['headers']))
            title_cell = ws.cell(row=1, column=1, value=sheet_cfg['title'])
            title_cell.font = Font(name=ACCOUNTING_FONTS['title'], size=FONT_SIZES['title'], bold=True, color=ACCOUNTING_PALETTE['header_dark'])
            title_cell.alignment = Alignment(horizontal='center', vertical='center')
            ws.row_dimensions[1].height = 30
            current_row = 3
        else:
            current_row = 1
        
        # الرؤوس
        header_fill = PatternFill(start_color=ACCOUNTING_PALETTE['header_dark'], end_color=ACCOUNTING_PALETTE['header_dark'], fill_type='solid')
        header_font = Font(name=ACCOUNTING_FONTS['heading'], size=FONT_SIZES['body'], bold=True, color=ACCOUNTING_PALETTE['white'])
        
        for col_num, header in enumerate(sheet_cfg['headers'], 1):
            cell = ws.cell(row=current_row, column=col_num, value=header)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
            cell.border = Border(
                left=Side(style='thin', color=ACCOUNTING_PALETTE['border']),
                right=Side(style='thin', color=ACCOUNTING_PALETTE['border']),
                top=Side(style='thin', color=ACCOUNTING_PALETTE['border']),
                bottom=Side(style='medium', color=ACCOUNTING_PALETTE['header_dark'])
            )
        
        ws.row_dimensions[current_row].height = 28
        data_start_row = current_row + 1
        
        # البيانات
        for row_idx, row_data in enumerate(sheet_cfg['data']):
            for col_idx, cell_data in enumerate(row_data, 1):
                cell = ws.cell(row=data_start_row + row_idx, column=col_idx, value=cell_data)
                cell.font = Font(name=ACCOUNTING_FONTS['body'], size=FONT_SIZES['body'])
                cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
                cell.border = Border(
                    left=Side(style='thin', color=ACCOUNTING_PALETTE['border']),
                    right=Side(style='thin', color=ACCOUNTING_PALETTE['border']),
                    top=Side(style='thin', color=ACCOUNTING_PALETTE['border']),
                    bottom=Side(style='thin', color=ACCOUNTING_PALETTE['border'])
                )
                
                # تنسيق الأرقام
                if isinstance(cell_data, (int, float)):
                    cell.number_format = '#,##0.00'
                
                # خلفية بديلة
                bg_color = ACCOUNTING_PALETTE['row_even'] if row_idx % 2 == 0 else ACCOUNTING_PALETTE['row_odd']
                cell.fill = PatternFill(start_color=bg_color, end_color=bg_color, fill_type='solid')
        
        # تجميد
        ws.freeze_panes = f'A{data_start_row}'
        
        # فلتر
        if len(sheet_cfg['data']) > 0:
            last_row = data_start_row + len(sheet_cfg['data']) - 1
            ws.auto_filter.ref = f'A{current_row}:{get_column_letter(len(sheet_cfg["headers"]))}{last_row}'
        
        # ضبط العرض
        for col_num in range(1, len(sheet_cfg['headers']) + 1):
            max_length = len(str(sheet_cfg['headers'][col_num - 1])) * 1.5
            for row in ws.iter_rows(min_row=data_start_row, max_row=data_start_row + len(sheet_cfg['data']) - 1, min_col=col_num, max_col=col_num):
                for cell in row:
                    if cell.value:
                        max_length = max(max_length, len(str(cell.value)) * 1.2)
            ws.column_dimensions[get_column_letter(col_num)].width = min(max(max_length + 4, 12), 45)
    
    # الحفظ
    save_dir = output_dir or WORKSPACE_DIR
    os.makedirs(save_dir, exist_ok=True)
    
    if filename is None:
        filename = f"Report_{datetime.now().strftime('%Y%m%d')}"
    
    filename = get_unique_filename(save_dir, filename, 'xlsx')
    output_path = os.path.join(save_dir, filename)
    
    wb.save(output_path)
    return output_path


# ==========================================
# دالة إنشاء فاتورة احترافية
# ==========================================
def create_professional_invoice(
    company_name,
    client_name,
    invoice_number,
    invoice_date,
    items,
    tax_rate=0.05,
    notes=None,
    output_dir=None,
    filename=None
):
    """
    إنشاء فاتورة احترافية
    
    items: قائمة من القواميس، كل عنصر يحتوي على:
    - description: الوصف
    - quantity: الكمية
    - unit_price: سعر الوحدة
    """
    from core.config import WORKSPACE_DIR
    
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "الفاتورة"
    ws.sheet_view.rightToLeft = True
    
    # عنوان الفاتورة
    ws.merge_cells('A1:F1')
    title_cell = ws['A1']
    title_cell.value = f"فاتورة - {company_name}"
    title_cell.font = Font(name=ACCOUNTING_FONTS['title'], size=FONT_SIZES['mega_title'], bold=True, color=ACCOUNTING_PALETTE['header_dark'])
    title_cell.alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[1].height = 40
    
    # معلومات الفاتورة
    ws['A2'] = f"رقم الفاتورة: {invoice_number}"
    ws['A2'].font = Font(bold=True, size=FONT_SIZES['body'])
    ws['A3'] = f"التاريخ: {invoice_date}"
    ws['A3'].font = Font(bold=True, size=FONT_SIZES['body'])
    ws['A4'] = f"العميل: {client_name}"
    ws['A4'].font = Font(bold=True, size=FONT_SIZES['body'])
    
    # رؤوس الجدول
    headers = ['#', 'الوصف', 'الكمية', 'سعر الوحدة', 'الإجمالي']
    header_row = 6
    
    for col_num, header in enumerate(headers, 1):
        cell = ws.cell(row=header_row, column=col_num, value=header)
        cell.fill = PatternFill(start_color=ACCOUNTING_PALETTE['header_dark'], end_color=ACCOUNTING_PALETTE['header_dark'], fill_type='solid')
        cell.font = Font(bold=True, color=ACCOUNTING_PALETTE['white'], size=FONT_SIZES['body'])
        cell.alignment = Alignment(horizontal='center', vertical='center')
        cell.border = Border(
            left=Side(style='thin', color=ACCOUNTING_PALETTE['border']),
            right=Side(style='thin', color=ACCOUNTING_PALETTE['border']),
            top=Side(style='thin', color=ACCOUNTING_PALETTE['border']),
            bottom=Side(style='medium', color=ACCOUNTING_PALETTE['header_dark'])
        )
    
    ws.row_dimensions[header_row].height = 28
    
    # البيانات
    data_start_row = header_row + 1
    total = 0
    
    for idx, item in enumerate(items, 1):
        row = data_start_row + idx - 1
        quantity = item['quantity']
        unit_price = item['unit_price']
        item_total = quantity * unit_price
        total += item_total
        
        ws.cell(row=row, column=1, value=idx)
        ws.cell(row=row, column=2, value=item['description'])
        ws.cell(row=row, column=3, value=quantity)
        ws.cell(row=row, column=4, value=unit_price)
        ws.cell(row=row, column=5, value=item_total)
        
        for col in range(1, 6):
            cell = ws.cell(row=row, column=col)
            cell.alignment = Alignment(horizontal='center', vertical='center')
            cell.border = Border(
                left=Side(style='thin', color=ACCOUNTING_PALETTE['border']),
                right=Side(style='thin', color=ACCOUNTING_PALETTE['border']),
                top=Side(style='thin', color=ACCOUNTING_PALETTE['border']),
                bottom=Side(style='thin', color=ACCOUNTING_PALETTE['border'])
            )
            if col >= 3:
                cell.number_format = '#,##0.00'
        
        bg_color = ACCOUNTING_PALETTE['row_even'] if idx % 2 == 0 else ACCOUNTING_PALETTE['row_odd']
        for col in range(1, 6):
            ws.cell(row=row, column=col).fill = PatternFill(start_color=bg_color, end_color=bg_color, fill_type='solid')
    
    # الإجماليات
    total_row = data_start_row + len(items)
    ws.cell(row=total_row, column=4, value="الإجمالي الفرعي:")
    ws.cell(row=total_row, column=4).font = Font(bold=True, size=FONT_SIZES['body'])
    ws.cell(row=total_row, column=5, value=total)
    ws.cell(row=total_row, column=5).font = Font(bold=True, size=FONT_SIZES['body'])
    ws.cell(row=total_row, column=5).number_format = '#,##0.00'
    
    tax_row = total_row + 1
    tax_amount = total * tax_rate
    ws.cell(row=tax_row, column=4, value=f"الضريبة ({tax_rate*100}%):")
    ws.cell(row=tax_row, column=4).font = Font(bold=True, size=FONT_SIZES['body'])
    ws.cell(row=tax_row, column=5, value=tax_amount)
    ws.cell(row=tax_row, column=5).font = Font(bold=True, size=FONT_SIZES['body'])
    ws.cell(row=tax_row, column=5).number_format = '#,##0.00'
    
    final_row = tax_row + 1
    final_total = total + tax_amount
    ws.cell(row=final_row, column=4, value="الإجمالي النهائي:")
    ws.cell(row=final_row, column=4).font = Font(bold=True, size=FONT_SIZES['heading'], color=ACCOUNTING_PALETTE['header_dark'])
    ws.cell(row=final_row, column=5, value=final_total)
    ws.cell(row=final_row, column=5).font = Font(bold=True, size=FONT_SIZES['heading'], color=ACCOUNTING_PALETTE['header_dark'])
    ws.cell(row=final_row, column=5).number_format = '#,##0.00'
    
    for col in range(4, 6):
        for row in [total_row, tax_row, final_row]:
            ws.cell(row=row, column=col).fill = PatternFill(start_color=ACCOUNTING_PALETTE['total_bg'], end_color=ACCOUNTING_PALETTE['total_bg'], fill_type='solid')
    
    # الملاحظات
    if notes:
        notes_row = final_row + 2
        ws.cell(row=notes_row, column=1, value="ملاحظات:")
        ws.cell(row=notes_row, column=1).font = Font(bold=True, size=FONT_SIZES['body'])
        ws.cell(row=notes_row + 1, column=1, value=notes)
        ws.cell(row=notes_row + 1, column=1).font = Font(size=FONT_SIZES['body'])
    
    # ضبط العرض
    ws.column_dimensions['A'].width = 5
    ws.column_dimensions['B'].width = 40
    ws.column_dimensions['C'].width = 12
    ws.column_dimensions['D'].width = 15
    ws.column_dimensions['E'].width = 15
    
    # تجميد
    ws.freeze_panes = f'A{data_start_row}'
    
    # الحفظ
    save_dir = output_dir or WORKSPACE_DIR
    os.makedirs(save_dir, exist_ok=True)
    
    if filename is None:
        filename = f"فاتورة_{invoice_number}"
    
    filename = get_unique_filename(save_dir, filename, 'xlsx')
    output_path = os.path.join(save_dir, filename)
    
    wb.save(output_path)
    return output_path
