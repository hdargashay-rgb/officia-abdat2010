"""
Officia Formatter Pro v4 - نظام التنسيق الاحترافي المتقدم
دوال تنسيق ذكية متعددة الاستخدام
"""

import os
from datetime import datetime
from docx import Document
from docx.shared import Pt, Cm, RGBColor, Inches, Emu
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.enum.section import WD_ORIENT
from docx.oxml.ns import qn, nsdecls
from docx.oxml import parse_xml
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, numbers, GradientFill
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, PieChart, LineChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.worksheet.table import Table, TableStyleInfo
from core.styles import (
    apply_header_style, apply_content_style, apply_table_style,
    apply_footer_style, apply_signature_style,
    classify_document_type, get_recommended_styles
)
import re

# ==========================================
# نظام الألوان الاحترافي المتقدم
# ==========================================
PALETTE = {
    # الألوان الأساسية - تدرج أزرق ملكي
    'royal_dark': '1B3A5C',
    'royal_mid': '2E5C8A',
    'royal_light': '4A7FB5',
    'royal_pale': 'D6E4F0',
    
    # ألوان التمييز
    'gold': 'C9A961',
    'gold_light': 'F5E6C8',
    'silver': 'BDC3C7',
    
    # ألوان التنبيه
    'accent_red': 'A93226',
    'accent_green': '1E8449',
    'accent_orange': 'D68910',
    'accent_yellow': 'F4D03F',
    
    # ألوان محايدة
    'charcoal': '2C3E50',
    'slate': '5D6D7E',
    'cream': 'FDFEFE',
    'white': 'FFFFFF',
    'black': '000000',
    
    # ألوان الصفوف
    'row_even': 'D6E4F0',
    'row_odd': 'FFFFFF',
    'row_highlight': 'FFF9C4',
}

# الخطوط الاحترافية
FONTS = {
    'title': 'Traditional Arabic',
    'heading': 'Simplified Arabic',
    'body': 'Calibri',
    'accent': 'Amiri',
    'numbers': 'Arial',
}

# أحجام الخطوط
SIZES = {
    'mega_title': 26,
    'title': 20,
    'heading': 16,
    'subheading': 14,
    'body': 13,
    'small': 11,
    'caption': 10,
    'footer': 9,
    'tiny': 8,
}

# ==========================================
# دوال مساعدة عامة
# ==========================================
def set_cell_shading(cell, color_hex):
    """تعيين لون خلفية خلية"""
    shading = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{color_hex}"/>')
    cell._tc.get_or_add_tcPr().append(shading)

def set_cell_gradient(cell, color1, color2):
    """تعيين تدرج لوني للخلية"""
    shading = parse_xml(
        f'<w:shd {nsdecls("w")} w:fill="{color1}" w:val="clear"/>'
    )
    cell._tc.get_or_add_tcPr().append(shading)

def set_paragraph_rtl(paragraph):
    """تعيين اتجاه الفقرة RTL"""
    pPr = paragraph._p.get_or_add_pPr()
    pPr.set(qn('w:bidi'), '1')
    pPr.set(qn('w:jc'), 'right')

def set_cell_rtl(cell):
    """تعيين اتجاه الخلية RTL"""
    for para in cell.paragraphs:
        set_paragraph_rtl(para)

def get_unique_filename(directory, base_name, extension):
    """إنشاء اسم ملف فريد"""
    counter = 1
    filename = f"{base_name}.{extension}"
    while os.path.exists(os.path.join(directory, filename)):
        filename = f"{base_name}_{counter}.{extension}"
        counter += 1
    return filename

# ==========================================
# دوال تنسيق Word الاحترافية
# ==========================================
def add_elegant_header(doc, organization_name, title, subtitle=None):
    """
    إضافة رأس خطاب أنيق
    - اسم المؤسسة
    - خط فاصل مزدوج
    - العنوان الرئيسي
    - خط ذهبي
    - العنوان الفرعي (اختياري)
    """
    # اسم المؤسسة
    if organization_name:
        org_para = doc.add_paragraph()
        org_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        org_para.paragraph_format.space_after = Pt(6)
        org_run = org_para.add_run(organization_name)
        org_run.font.name = FONTS['title']
        org_run.font.size = Pt(SIZES['title'])
        org_run.font.bold = True
        org_run.font.color.rgb = RGBColor.from_string(PALETTE['royal_dark'])
        set_paragraph_rtl(org_para)
    
    # خط فاصل مزدوج
    para = doc.add_paragraph()
    para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    para.paragraph_format.space_before = Pt(6)
    para.paragraph_format.space_after = Pt(6)
    pPr = para._p.get_or_add_pPr()
    pBdr = parse_xml(
        f'<w:pBdr {nsdecls("w")}>'
        f'  <w:top w:val="double" w:sz="12" w:space="4" w:color="{PALETTE["royal_dark"]}"/>'
        f'</w:pBdr>'
    )
    pPr.append(pBdr)
    
    # العنوان الرئيسي
    title_para = doc.add_heading(title, level=0)
    title_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title_para.paragraph_format.space_before = Pt(12)
    title_para.paragraph_format.space_after = Pt(12)
    for run in title_para.runs:
        run.font.name = FONTS['title']
        run.font.size = Pt(SIZES['mega_title'])
        run.font.bold = True
        run.font.color.rgb = RGBColor.from_string(PALETTE['royal_dark'])
    
    # خط ذهبي
    para = doc.add_paragraph()
    para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    para.paragraph_format.space_before = Pt(6)
    para.paragraph_format.space_after = Pt(6)
    pPr = para._p.get_or_add_pPr()
    pBdr = parse_xml(
        f'<w:pBdr {nsdecls("w")}>'
        f'  <w:bottom w:val="single" w:sz="8" w:space="1" w:color="{PALETTE["gold"]}"/>'
        f'</w:pBdr>'
    )
    pPr.append(pBdr)
    
    # العنوان الفرعي
    if subtitle:
        sub_para = doc.add_paragraph()
        sub_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        sub_para.paragraph_format.space_after = Pt(12)
        sub_run = sub_para.add_run(subtitle)
        sub_run.font.name = FONTS['heading']
        sub_run.font.size = Pt(SIZES['subheading'])
        sub_run.font.color.rgb = RGBColor.from_string(PALETTE['royal_mid'])
        set_paragraph_rtl(sub_para)

def add_document_info(doc, date=None, reference_number=None, recipient=None, subject=None):
    """
    إضافة معلومات الوثيقة
    - التاريخ والرقم
    - المرسل إليه
    - الموضوع
    """
    # التاريخ والرقم
    info_table = doc.add_table(rows=1, cols=2)
    info_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    
    date_cell = info_table.cell(0, 0)
    date_para = date_cell.paragraphs[0]
    date_para.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    date_label = date_para.add_run("التاريخ: ")
    date_label.font.bold = True
    date_label.font.size = Pt(SIZES['body'])
    date_label.font.color.rgb = RGBColor.from_string(PALETTE['royal_mid'])
    date_value = date_para.add_run(date or datetime.now().strftime('%Y/%m/%d'))
    date_value.font.size = Pt(SIZES['body'])
    set_paragraph_rtl(date_para)
    
    ref_cell = info_table.cell(0, 1)
    ref_para = ref_cell.paragraphs[0]
    ref_para.alignment = WD_ALIGN_PARAGRAPH.LEFT
    ref_label = ref_para.add_run("رقم: ")
    ref_label.font.bold = True
    ref_label.font.size = Pt(SIZES['body'])
    ref_label.font.color.rgb = RGBColor.from_string(PALETTE['royal_mid'])
    ref_value = ref_para.add_run(reference_number or datetime.now().strftime('%Y%m%d%H%M'))
    ref_value.font.size = Pt(SIZES['body'])
    
    doc.add_paragraph()
    
    # المرسل إليه
    if recipient:
        to_para = doc.add_paragraph()
        to_para.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        to_para.paragraph_format.space_after = Pt(6)
        to_label = to_para.add_run("السادة/ ")
        to_label.font.name = FONTS['accent']
        to_label.font.bold = True
        to_label.font.size = Pt(SIZES['body'])
        to_label.font.color.rgb = RGBColor.from_string(PALETTE['royal_dark'])
        to_value = to_para.add_run(recipient)
        to_value.font.name = FONTS['accent']
        to_value.font.bold = True
        to_value.font.size = Pt(SIZES['body'])
        set_paragraph_rtl(to_para)
    
    # الموضوع
    if subject:
        subj_para = doc.add_paragraph()
        subj_para.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        subj_para.paragraph_format.space_after = Pt(12)
        subj_label = subj_para.add_run("الموضوع: ")
        subj_label.font.bold = True
        subj_label.font.size = Pt(SIZES['body'])
        subj_label.font.color.rgb = RGBColor.from_string(PALETTE['royal_dark'])
        subj_value = subj_para.add_run(subject)
        subj_value.font.bold = True
        subj_value.font.size = Pt(SIZES['body'])
        subj_value.font.color.rgb = RGBColor.from_string(PALETTE['royal_mid'])
        set_paragraph_rtl(subj_para)

def add_greeting(doc, greeting_text="تحية طيبة وبعد،،،"):
    """إضافة تحية رسمية"""
    greeting = doc.add_paragraph()
    greeting.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    greeting.paragraph_format.space_before = Pt(12)
    greeting.paragraph_format.space_after = Pt(12)
    greet_run = greeting.add_run(greeting_text)
    greet_run.font.name = FONTS['accent']
    greet_run.font.bold = True
    greet_run.font.size = Pt(SIZES['body'])
    greet_run.font.color.rgb = RGBColor.from_string(PALETTE['royal_dark'])
    set_paragraph_rtl(greeting)

def add_content_paragraphs(doc, paragraphs, numbered=False, first_line_indent=1.5):
    """
    إضافة فقرات المحتوى
    - numbered: ترقيم الفقرات
    - first_line_indent: إزاحة السطر الأول
    """
    for i, para_text in enumerate(paragraphs):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        p.paragraph_format.first_line_indent = Cm(first_line_indent)
        p.paragraph_format.line_spacing = 1.6
        p.paragraph_format.space_after = Pt(10)
        
        text = f"{i+1}. {para_text}" if numbered and len(paragraphs) > 1 else para_text
        run = p.add_run(text)
        run.font.name = FONTS['body']
        run.font.size = Pt(SIZES['body'])
        run.font.color.rgb = RGBColor.from_string(PALETTE['charcoal'])
        set_paragraph_rtl(p)

def create_professional_table(doc, headers, data, caption=None):
    """
    إنشاء جدول احترافي منسق
    - رؤوس بأزرق داكن
    - صفوف بديلة بألوان متناسقة
    - محاذاة مركزية
    - عنوان للجدول (اختياري)
    """
    if caption:
        cap_para = doc.add_paragraph()
        cap_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        cap_para.paragraph_format.space_before = Pt(12)
        cap_para.paragraph_format.space_after = Pt(8)
        cap_run = cap_para.add_run(caption)
        cap_run.font.name = FONTS['heading']
        cap_run.font.bold = True
        cap_run.font.size = Pt(SIZES['subheading'])
        cap_run.font.color.rgb = RGBColor.from_string(PALETTE['royal_dark'])
        set_paragraph_rtl(cap_para)
    
    # إنشاء الجدول
    table = doc.add_table(rows=1, cols=len(headers))
    table.style = 'Table Grid'
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = True
    
    # تنسيق رأس الجدول
    header_cells = table.rows[0].cells
    for i, header in enumerate(headers):
        cell = header_cells[i]
        cell.text = ''
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(4)
        p.paragraph_format.space_after = Pt(4)
        
        run = p.add_run(header)
        run.font.name = FONTS['heading']
        run.font.bold = True
        run.font.color.rgb = RGBColor.from_string(PALETTE['white'])
        run.font.size = Pt(SIZES['body'])
        
        set_cell_shading(cell, PALETTE['royal_dark'])
        set_cell_rtl(cell)
    
    # إضافة البيانات
    for row_idx, row_data in enumerate(data):
        row_cells = table.add_row().cells
        bg_color = PALETTE['row_even'] if row_idx % 2 == 0 else PALETTE['cream']
        
        for col_idx, cell_data in enumerate(row_data):
            cell = row_cells[col_idx]
            cell.text = ''
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_before = Pt(3)
            p.paragraph_format.space_after = Pt(3)
            
            run = p.add_run(str(cell_data))
            run.font.name = FONTS['body']
            run.font.size = Pt(SIZES['body'])
            run.font.color.rgb = RGBColor.from_string(PALETTE['charcoal'])
            
            set_cell_shading(cell, bg_color)
            set_cell_rtl(cell)
    
    doc.add_paragraph()
    return table

def add_closing(doc, closing_text="وتفضلوا بقبول فائق الاحترام والتقدير،،", signature_name=None, signature_title=None):
    """
    إضافة خاتمة وتوقيع
    - عبارة الخاتمة
    - اسم الموقع
    - صفة الموقع
    """
    # الخاتمة
    closing = doc.add_paragraph()
    closing.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    closing.paragraph_format.space_before = Pt(12)
    closing.paragraph_format.space_after = Pt(12)
    closing_run = closing.add_run(closing_text)
    closing_run.font.name = FONTS['accent']
    closing_run.font.bold = True
    closing_run.font.size = Pt(SIZES['body'])
    closing_run.font.color.rgb = RGBColor.from_string(PALETTE['royal_dark'])
    set_paragraph_rtl(closing)
    
    # التوقيع
    if signature_name:
        doc.add_paragraph()
        doc.add_paragraph()
        sig_para = doc.add_paragraph()
        sig_para.alignment = WD_ALIGN_PARAGRAPH.LEFT
        
        name_run = sig_para.add_run(signature_name)
        name_run.font.name = FONTS['accent']
        name_run.font.bold = True
        name_run.font.size = Pt(SIZES['body'])
        name_run.font.color.rgb = RGBColor.from_string(PALETTE['royal_dark'])
        
        if signature_title:
            sig_para.add_run('\n')
            title_run = sig_para.add_run(signature_title)
            title_run.font.size = Pt(SIZES['small'])
            title_run.font.color.rgb = RGBColor.from_string(PALETTE['royal_mid'])

def add_note(doc, note_text, icon="⚠"):
    """إضافة ملاحظة/تنبيه"""
    note_para = doc.add_paragraph()
    note_para.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    note_para.paragraph_format.space_before = Pt(6)
    
    note_icon = note_para.add_run(f"{icon} تنبيه: ")
    note_icon.font.bold = True
    note_icon.font.size = Pt(SIZES['body'])
    note_icon.font.color.rgb = RGBColor.from_string(PALETTE['accent_red'])
    
    note_text_run = note_para.add_run(note_text)
    note_text_run.font.size = Pt(SIZES['body'])
    note_text_run.font.color.rgb = RGBColor.from_string(PALETTE['accent_red'])
    set_paragraph_rtl(note_para)

def add_attachments(doc, attachments_list):
    """إضافة قائمة المرفقات"""
    att_label_para = doc.add_paragraph()
    att_label_para.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    att_label = att_label_para.add_run("المرفقات: ")
    att_label.font.bold = True
    att_label.font.size = Pt(SIZES['body'])
    att_label.font.color.rgb = RGBColor.from_string(PALETTE['royal_dark'])
    set_paragraph_rtl(att_label_para)
    
    for att in attachments_list:
        att_item = doc.add_paragraph()
        att_item.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        att_item.paragraph_format.left_indent = Cm(1)
        att_run = att_item.add_run(f"• {att}")
        att_run.font.size = Pt(SIZES['body'])
        set_paragraph_rtl(att_item)

def setup_document(doc, margins_cm=1.8):
    """إعداد المستند بهوامش أصغر لتوفير المساحة"""
    section = doc.sections[0]
    section.top_margin = Cm(margins_cm)
    section.bottom_margin = Cm(margins_cm)
    section.left_margin = Cm(margins_cm)
    section.right_margin = Cm(margins_cm)
    
    style = doc.styles['Normal']
    style.font.name = FONTS['body']
    style.font.size = Pt(SIZES['body'])
    style.font.color.rgb = RGBColor.from_string(PALETTE['charcoal'])
    style.paragraph_format.line_spacing = 1.3
    style.paragraph_format.space_after = Pt(6)

def add_footer(doc, footer_text="Officia - الوكيل الذكي للمهام المكتبية"):
    """إضافة تذييل احترافي"""
    section = doc.sections[0]
    footer = section.footer
    footer_para = footer.paragraphs[0]
    footer_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    footer_line = footer_para.add_run("━" * 40 + "\n")
    footer_line.font.color.rgb = RGBColor.from_string(PALETTE['silver'])
    footer_line.font.size = Pt(SIZES['caption'])
    
    footer_text_run = footer_para.add_run(footer_text)
    footer_text_run.font.size = Pt(SIZES['footer'])
    footer_text_run.font.color.rgb = RGBColor.from_string(PALETTE['slate'])
    footer_text_run.font.italic = True
    # ==========================================
# دوال Excel الاحترافية
# ==========================================
def create_excel_sheet(ws, headers, data, title=None, 
                       start_row=1, freeze_row=True,
                       auto_filter=True, table_style=True):
    """
    إنشاء ورقة Excel منسقة باحترافية كاملة
    
    تطبق: RTL، عناوين ملونة، حدود، auto-fit، تجميد، فلاتر
    """
    ws.sheet_view.rightToLeft = True
    ws.sheet_view.showGridLines = True
    
    current_row = start_row
    
    # العنوان
    if title:
        ws.merge_cells(
            start_row=current_row, start_column=1,
            end_row=current_row, end_column=len(headers)
        )
        title_cell = ws.cell(row=current_row, column=1, value=title)
        title_cell.font = Font(
            name=FONTS['title'],
            size=SIZES['title'],
            bold=True,
            color=PALETTE['royal_dark']
        )
        title_cell.alignment = Alignment(horizontal='center', vertical='center')
        title_cell.fill = PatternFill(
            start_color=PALETTE['royal_pale'],
            end_color=PALETTE['royal_pale'],
            fill_type='solid'
        )
        ws.row_dimensions[current_row].height = 35
        current_row += 2
    
    # رؤوس الأعمدة
    header_fill = PatternFill(
        start_color=PALETTE['royal_dark'],
        end_color=PALETTE['royal_mid'],
        fill_type='solid'
    )
    header_font = Font(
        name=FONTS['heading'],
        size=SIZES['body'],
        bold=True,
        color=PALETTE['white']
    )
    header_alignment = Alignment(
        horizontal='center',
        vertical='center',
        wrap_text=True
    )
    
    header_border = Border(
        left=Side(style='thin', color=PALETTE['silver']),
        right=Side(style='thin', color=PALETTE['silver']),
        top=Side(style='thin', color=PALETTE['silver']),
        bottom=Side(style='medium', color=PALETTE['gold'])
    )
    
    header_row = current_row
    for col_num, header in enumerate(headers, 1):
        cell = ws.cell(row=header_row, column=col_num, value=header)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = header_alignment
        cell.border = header_border
    
    ws.row_dimensions[header_row].height = 28
    data_start_row = header_row + 1
    
    # البيانات
    data_font = Font(name=FONTS['body'], size=SIZES['body'])
    data_alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    
    even_fill = PatternFill(start_color=PALETTE['row_even'], end_color=PALETTE['row_even'], fill_type='solid')
    odd_fill = PatternFill(start_color=PALETTE['cream'], end_color=PALETTE['cream'], fill_type='solid')
    
    data_border = Border(
        left=Side(style='thin', color=PALETTE['silver']),
        right=Side(style='thin', color=PALETTE['silver']),
        top=Side(style='thin', color=PALETTE['silver']),
        bottom=Side(style='thin', color=PALETTE['silver'])
    )
    
    for row_idx, row_data in enumerate(data):
        for col_idx, cell_data in enumerate(row_data, 1):
            cell = ws.cell(row=data_start_row + row_idx, column=col_idx, value=cell_data)
            cell.font = data_font
            cell.alignment = data_alignment
            cell.border = data_border
            
            cell.fill = even_fill if row_idx % 2 == 0 else odd_fill
            
            if isinstance(cell_data, (int, float)):
                cell.number_format = '#,##0.00'
    
    # auto-fit ذكي
    for col_num in range(1, len(headers) + 1):
        header_width = len(str(headers[col_num - 1])) * 2.2
        max_data_width = 0
        
        for row in ws.iter_rows(
            min_row=data_start_row, 
            max_row=data_start_row + len(data) - 1,
            min_col=col_num, 
            max_col=col_num
        ):
            for cell in row:
                if cell.value:
                    cell_width = len(str(cell.value)) * 1.5
                    max_data_width = max(max_data_width, cell_width)
        
        final_width = max(header_width, max_data_width) + 4
        final_width = min(final_width, 50)
        final_width = max(final_width, 12)
        
        ws.column_dimensions[get_column_letter(col_num)].width = final_width
    
    # تجميد
    if freeze_row:
        ws.freeze_panes = f'A{data_start_row}'
    
    # فلتر
    if auto_filter and len(data) > 0:
        last_row = data_start_row + len(data) - 1
        last_col = get_column_letter(len(headers))
        ws.auto_filter.ref = f'A{header_row}:{last_col}{last_row}'
    
    return data_start_row


def create_multi_sheet_excel(
    sheets_config,
    output_dir=None,
    filename=None
):
    """
    إنشاء ملف Excel متعدد الأوراق
    
    sheets_config: قائمة من القواميس:
    - name: اسم الورقة
    - title: عنوان الورقة (اختياري)
    - headers: قائمة العناوين
    - data: قائمة البيانات
    """
    from core.config import WORKSPACE_DIR
    
    wb = openpyxl.Workbook()
    
    for idx, sheet_cfg in enumerate(sheets_config):
        if idx == 0:
            ws = wb.active
            ws.title = sheet_cfg['name'][:31]
        else:
            ws = wb.create_sheet(title=sheet_cfg['name'][:31])
        
        create_excel_sheet(
            ws=ws,
            headers=sheet_cfg['headers'],
            data=sheet_cfg['data'],
            title=sheet_cfg.get('title'),
            freeze_row=True,
            auto_filter=True
        )
    
    save_dir = output_dir or WORKSPACE_DIR
    os.makedirs(save_dir, exist_ok=True)
    
    if filename is None:
        filename = f"Report_{datetime.now().strftime('%Y%m%d')}"
    
    filename = get_unique_filename(save_dir, filename, 'xlsx')
    output_path = os.path.join(save_dir, filename)
    
    wb.save(output_path)
    return output_path


def create_fixed_asset_report(
    company_name,
    year,
    sections,
    output_dir=None,
    filename=None
):
    """
    إنشاء تقرير أصول ثابتة احترافي كامل
    (مطابق لملف FIXED ASSET LISTING المرجعي)
    
    sections: قائمة من القواميس:
    - name: اسم القسم
    - depreciation_rate: نسبة الإهلاك
    - assets: قائمة الأصول
    """
    from core.config import WORKSPACE_DIR
    
    wb = openpyxl.Workbook()
    
    # ورقة الملخص
    ws_summary = wb.active
    ws_summary.title = "Summary"
    ws_summary.sheet_view.rightToLeft = True
    
    # عنوان التقرير
    ws_summary.merge_cells('A1:Y1')
    title_cell = ws_summary['A1']
    title_cell.value = f"Fixed Asset Summary - {year} {company_name}"
    title_cell.font = Font(name=FONTS['title'], size=SIZES['mega_title'], bold=True, color=PALETTE['royal_dark'])
    title_cell.alignment = Alignment(horizontal='center', vertical='center')
    ws_summary.row_dimensions[1].height = 35
    
    # رؤوس الأعمدة
    headers = [
        'Fixed Assets', 'Opening Balance', 'Additions', 'Disposal', 'Total',
        'Op. Accu. Dep', 'Depr. 2025',
        'Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec',
        'Monthly Total', 'Dep. During the year', 'Write of AD', 'Total Accu. Dep', 'Net Book Value', 'For check'
    ]
    
    header_row = 3
    for col_num, header in enumerate(headers, 1):
        cell = ws_summary.cell(row=header_row, column=col_num, value=header)
        cell.font = Font(name=FONTS['heading'], size=SIZES['body'], bold=True, color=PALETTE['white'])
        cell.fill = PatternFill(start_color=PALETTE['royal_dark'], end_color=PALETTE['royal_dark'], fill_type='solid')
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        cell.border = Border(
            left=Side(style='thin', color=PALETTE['silver']),
            right=Side(style='thin', color=PALETTE['silver']),
            top=Side(style='thin', color=PALETTE['silver']),
            bottom=Side(style='medium', color=PALETTE['royal_dark'])
        )
    
    ws_summary.row_dimensions[header_row].height = 30
    
    # بيانات الأقسام
    current_row = header_row + 1
    total_opening = 0
    total_depr = 0
    total_net = 0
    
    for section in sections:
        section_opening = sum(asset['book_value'] for asset in section['assets'])
        section_depr_rate = section['depreciation_rate']
        section_annual_depr = section_opening * section_depr_rate
        section_monthly_depr = section_annual_depr / 12
        section_accumulated_depr = section_opening * section_depr_rate * 2
        section_net_value = section_opening - section_accumulated_depr
        
        total_opening += section_opening
        total_depr += section_annual_depr
        total_net += section_net_value
        
        ws_summary.cell(row=current_row, column=1, value=f"RADIANT - {section['name']}")
        ws_summary.cell(row=current_row, column=1).font = Font(bold=True, color=PALETTE['royal_dark'])
        ws_summary.cell(row=current_row, column=2, value=section_opening)
        ws_summary.cell(row=current_row, column=3, value=0)
        ws_summary.cell(row=current_row, column=4, value=0)
        ws_summary.cell(row=current_row, column=5, value=section_opening)
        ws_summary.cell(row=current_row, column=6, value=section_accumulated_depr)
        ws_summary.cell(row=current_row, column=7, value=section_annual_depr)
        
        for month in range(8, 20):
            ws_summary.cell(row=current_row, column=month, value=section_monthly_depr)
        
        ws_summary.cell(row=current_row, column=20, value=section_monthly_depr)
        ws_summary.cell(row=current_row, column=21, value=section_annual_depr)
        ws_summary.cell(row=current_row, column=22, value=0)
        ws_summary.cell(row=current_row, column=23, value=section_accumulated_depr + section_annual_depr)
        ws_summary.cell(row=current_row, column=24, value=section_net_value)
        
        for col in range(2, 25):
            cell = ws_summary.cell(row=current_row, column=col)
            cell.number_format = '#,##0.00'
            cell.alignment = Alignment(horizontal='center')
        
        for col in range(1, 25):
            set_cell_shading(ws_summary.cell(row=current_row, column=col), PALETTE['section_bg'])
        
        current_row += 1
    
    # صف الإجمالي
    ws_summary.cell(row=current_row, column=1, value="Total")
    ws_summary.cell(row=current_row, column=1).font = Font(bold=True, size=SIZES['body'], color=PALETTE['royal_dark'])
    ws_summary.cell(row=current_row, column=2, value=total_opening)
    ws_summary.cell(row=current_row, column=7, value=total_depr)
    ws_summary.cell(row=current_row, column=24, value=total_net)
    
    for col in range(2, 25):
        cell = ws_summary.cell(row=current_row, column=col)
        cell.number_format = '#,##0.00'
        cell.font = Font(bold=True)
        set_cell_shading(cell, PALETTE['total_bg'])
    
    ws_summary.column_dimensions['A'].width = 30
    for col in range(2, 25):
        ws_summary.column_dimensions[get_column_letter(col)].width = 12
    
    # أوراق الأقسام التفصيلية
    for section in sections:
        ws = wb.create_sheet(title=section['name'][:31])
        ws.sheet_view.rightToLeft = True
        
        ws.merge_cells('A1:Q1')
        title_cell = ws['A1']
        title_cell.value = f"RADIANT - {section['name']}"
        title_cell.font = Font(name=FONTS['title'], size=SIZES['title'], bold=True, color=PALETTE['white'])
        title_cell.fill = PatternFill(start_color=PALETTE['royal_dark'], end_color=PALETTE['royal_dark'], fill_type='solid')
        title_cell.alignment = Alignment(horizontal='center', vertical='center')
        ws.row_dimensions[1].height = 30
        
        ws.merge_cells('A2:Q2')
        rate_cell = ws['A2']
        rate_cell.value = f"Depreciation Rate: {section['depreciation_rate']*100:.2f}%"
        rate_cell.font = Font(name=FONTS['body'], size=SIZES['body'], bold=True, color=PALETTE['royal_mid'])
        rate_cell.alignment = Alignment(horizontal='center')
        
        detail_headers = [
            'S.No.', 'Supplier Name', 'Description', 'FA No.', 'Book Value',
            'Addition', 'Sold/Dispose', 'Total', 'Date Of Purchase',
            'AD Dep. 12/31/24', 'Dep. for Full 2025',
            'Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec',
            'Total', 'Diff', 'W/Off AD', 'AD Dep. 12/31/25', 'Net Value 12/31/25'
        ]
        
        header_row = 4
        for col_num, header in enumerate(detail_headers, 1):
            cell = ws.cell(row=header_row, column=col_num, value=header)
            cell.font = Font(name=FONTS['heading'], size=SIZES['small'], bold=True, color=PALETTE['white'])
            cell.fill = PatternFill(start_color=PALETTE['royal_mid'], end_color=PALETTE['royal_mid'], fill_type='solid')
            cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
            cell.border = Border(
                left=Side(style='thin', color=PALETTE['silver']),
                right=Side(style='thin', color=PALETTE['silver']),
                top=Side(style='thin', color=PALETTE['silver']),
                bottom=Side(style='medium', color=PALETTE['royal_dark'])
            )
        
        ws.row_dimensions[header_row].height = 35
        
        current_row = header_row + 1
        section_total_value = 0
        section_total_depr = 0
        
        for asset in section['assets']:
            book_value = asset['book_value']
            annual_depr = book_value * section['depreciation_rate']
            monthly_depr = annual_depr / 12
            accumulated_depr = book_value * section['depreciation_rate'] * 2
            net_value = book_value - accumulated_depr
            
            section_total_value += book_value
            section_total_depr += annual_depr
            
            ws.cell(row=current_row, column=1, value=asset.get('sno', ''))
            ws.cell(row=current_row, column=2, value=asset.get('supplier', ''))
            ws.cell(row=current_row, column=3, value=asset.get('description', ''))
            ws.cell(row=current_row, column=4, value=asset.get('fa_no', ''))
            ws.cell(row=current_row, column=5, value=book_value)
            ws.cell(row=current_row, column=6, value=asset.get('addition', 0))
            ws.cell(row=current_row, column=7, value=asset.get('disposal', 0))
            ws.cell(row=current_row, column=8, value=book_value)
            ws.cell(row=current_row, column=9, value=asset.get('purchase_date', ''))
            ws.cell(row=current_row, column=10, value=accumulated_depr)
            ws.cell(row=current_row, column=11, value=annual_depr)
            
            for month_idx, month in enumerate(range(12, 24)):
                ws.cell(row=current_row, column=month, value=monthly_depr)
            
            ws.cell(row=current_row, column=24, value=annual_depr)
            ws.cell(row=current_row, column=25, value=0)
            ws.cell(row=current_row, column=26, value=0)
            ws.cell(row=current_row, column=27, value=accumulated_depr + annual_depr)
            ws.cell(row=current_row, column=28, value=net_value)
            
            for col in [5, 6, 7, 8, 10, 11] + list(range(12, 29)):
                cell = ws.cell(row=current_row, column=col)
                cell.number_format = '#,##0.00'
                cell.alignment = Alignment(horizontal='center')
            
            bg_color = PALETTE['row_even'] if current_row % 2 == 0 else PALETTE['row_odd']
            for col in range(1, 29):
                set_cell_shading(ws.cell(row=current_row, column=col), bg_color)
            
            current_row += 1
        
        ws.cell(row=current_row, column=2, value="Total")
        ws.cell(row=current_row, column=2).font = Font(bold=True, color=PALETTE['royal_dark'])
        ws.cell(row=current_row, column=5, value=section_total_value)
        ws.cell(row=current_row, column=11, value=section_total_depr)
        
        for col in [5, 11] + list(range(12, 29)):
            cell = ws.cell(row=current_row, column=col)
            cell.number_format = '#,##0.00'
            cell.font = Font(bold=True)
            set_cell_shading(cell, PALETTE['total_bg'])
        
        ws.column_dimensions['A'].width = 6
        ws.column_dimensions['B'].width = 25
        ws.column_dimensions['C'].width = 35
        ws.column_dimensions['D'].width = 10
        for col in range(5, 29):
            ws.column_dimensions[get_column_letter(col)].width = 12
        
        ws.freeze_panes = f'A{header_row + 1}'
        ws.auto_filter.ref = f'A{header_row}:{get_column_letter(len(detail_headers))}{current_row - 1}'
    
    save_dir = output_dir or WORKSPACE_DIR
    os.makedirs(save_dir, exist_ok=True)
    
    if filename is None:
        filename = f"Fixed_Asset_{company_name}_{year}"
    
    filename = get_unique_filename(save_dir, filename, 'xlsx')
    output_path = os.path.join(save_dir, filename)
    
    wb.save(output_path)
    return output_path
    # ==========================================
# دوال Excel متعددة الأغراض
# ==========================================
def create_multi_sheet_excel(
    sheets_config,
    output_dir=None,
    filename=None
):
    """
    إنشاء ملف Excel متعدد الأوراق
    
    sheets_config: قائمة من القواميس:
    - name: اسم الورقة
    - title: عنوان الورقة (اختياري)
    - headers: قائمة العناوين
    - data: قائمة البيانات
    """
    from core.config import WORKSPACE_DIR
    
    wb = openpyxl.Workbook()
    
    for idx, sheet_cfg in enumerate(sheets_config):
        if idx == 0:
            ws = wb.active
            ws.title = sheet_cfg['name'][:31]
        else:
            ws = wb.create_sheet(title=sheet_cfg['name'][:31])
        
        create_excel_sheet(
            ws=ws,
            headers=sheet_cfg['headers'],
            data=sheet_cfg['data'],
            title=sheet_cfg.get('title'),
            freeze_row=True,
            auto_filter=True
        )
    
    save_dir = output_dir or WORKSPACE_DIR
    os.makedirs(save_dir, exist_ok=True)

    # ==========================================
# دالة شاملة لإنشاء خطاب احترافي
# ==========================================
def generate_professional_letter(
    organization_name,
    title,
    recipient,
    subject,
    paragraphs,
    table_headers=None,
    table_data=None,
    table_caption=None,
    closing_text="وتفضلوا بقبول فائق الاحترام والتقدير،،",
    signature_name=None,
    signature_title=None,
    note=None,
    output_dir=None,
    filename=None
):
    """دالة شاملة لإنشاء خطاب احترافي كامل - بدون ترقيم"""
    from core.config import WORKSPACE_DIR
    
    doc = Document()
    setup_document(doc)
    add_elegant_header(doc, organization_name, title)
    add_document_info(doc, recipient=recipient, subject=subject)
    add_greeting(doc)
    
    # إضافة الفقرات بدون ترقيم نهائياً
    for para_text in paragraphs:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        p.paragraph_format.first_line_indent = Cm(1.5)
        p.paragraph_format.line_spacing = 1.6
        p.paragraph_format.space_after = Pt(10)
        run = p.add_run(para_text)
        run.font.name = FONTS['body']
        run.font.size = Pt(SIZES['body'])
        run.font.color.rgb = RGBColor.from_string(PALETTE['charcoal'])
        set_paragraph_rtl(p)
    
    if table_headers and table_data:
        create_professional_table(doc, table_headers, table_data, table_caption)
    
    add_closing(doc, closing_text, signature_name, signature_title)
    
    if note:
        add_note(doc, note)
    
    add_footer(doc)
    
    save_dir = output_dir or WORKSPACE_DIR
    os.makedirs(save_dir, exist_ok=True)
    
    if filename is None:
        safe_title = title[:30].replace(' ', '_').replace('/', '-')
        filename = f"خطاب_{safe_title}"
    
    filename = get_unique_filename(save_dir, filename, 'docx')
    output_path = os.path.join(save_dir, filename)
    
    doc.save(output_path)
    return output_path
    
    if filename is None:
        filename = f"Report_{datetime.now().strftime('%Y%m%d')}"
    
    filename = get_unique_filename(save_dir, filename, 'xlsx')
    output_path = os.path.join(save_dir, filename)
    
    wb.save(output_path)
    return output_path


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
    
    items: قائمة من القواميس:
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
    title_cell.font = Font(name=FONTS['title'], size=SIZES['mega_title'], bold=True, color=PALETTE['royal_dark'])
    title_cell.alignment = Alignment(horizontal='center', vertical='center')
    title_cell.fill = PatternFill(start_color=PALETTE['royal_pale'], end_color=PALETTE['royal_pale'], fill_type='solid')
    ws.row_dimensions[1].height = 40
    
    # معلومات الفاتورة
    ws['A2'] = f"رقم الفاتورة: {invoice_number}"
    ws['A2'].font = Font(bold=True, size=SIZES['body'])
    ws['A3'] = f"التاريخ: {invoice_date}"
    ws['A3'].font = Font(bold=True, size=SIZES['body'])
    ws['A4'] = f"العميل: {client_name}"
    ws['A4'].font = Font(bold=True, size=SIZES['body'])
    
    # رؤوس الجدول
    headers = ['#', 'الوصف', 'الكمية', 'سعر الوحدة', 'الإجمالي']
    header_row = 6
    
    for col_num, header in enumerate(headers, 1):
        cell = ws.cell(row=header_row, column=col_num, value=header)
        cell.fill = PatternFill(start_color=PALETTE['royal_dark'], end_color=PALETTE['royal_dark'], fill_type='solid')
        cell.font = Font(bold=True, color=PALETTE['white'], size=SIZES['body'])
        cell.alignment = Alignment(horizontal='center', vertical='center')
        cell.border = Border(
            left=Side(style='thin', color=PALETTE['silver']),
            right=Side(style='thin', color=PALETTE['silver']),
            top=Side(style='thin', color=PALETTE['silver']),
            bottom=Side(style='medium', color=PALETTE['gold'])
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
                left=Side(style='thin', color=PALETTE['silver']),
                right=Side(style='thin', color=PALETTE['silver']),
                top=Side(style='thin', color=PALETTE['silver']),
                bottom=Side(style='thin', color=PALETTE['silver'])
            )
            if col >= 3:
                cell.number_format = '#,##0.00'
        
        bg_color = PALETTE['row_even'] if idx % 2 == 0 else PALETTE['cream']
        for col in range(1, 6):
            ws.cell(row=row, column=col).fill = PatternFill(start_color=bg_color, end_color=bg_color, fill_type='solid')
    
    # الإجماليات
    total_row = data_start_row + len(items)
    ws.cell(row=total_row, column=4, value="الإجمالي الفرعي:")
    ws.cell(row=total_row, column=4).font = Font(bold=True, size=SIZES['body'])
    ws.cell(row=total_row, column=5, value=total)
    ws.cell(row=total_row, column=5).font = Font(bold=True, size=SIZES['body'])
    ws.cell(row=total_row, column=5).number_format = '#,##0.00'
    
    tax_row = total_row + 1
    tax_amount = total * tax_rate
    ws.cell(row=tax_row, column=4, value=f"الضريبة ({tax_rate*100}%):")
    ws.cell(row=tax_row, column=4).font = Font(bold=True, size=SIZES['body'])
    ws.cell(row=tax_row, column=5, value=tax_amount)
    ws.cell(row=tax_row, column=5).font = Font(bold=True, size=SIZES['body'])
    ws.cell(row=tax_row, column=5).number_format = '#,##0.00'
    
    final_row = tax_row + 1
    final_total = total + tax_amount
    ws.cell(row=final_row, column=4, value="الإجمالي النهائي:")
    ws.cell(row=final_row, column=4).font = Font(bold=True, size=SIZES['heading'], color=PALETTE['royal_dark'])
    ws.cell(row=final_row, column=5, value=final_total)
    ws.cell(row=final_row, column=5).font = Font(bold=True, size=SIZES['heading'], color=PALETTE['royal_dark'])
    ws.cell(row=final_row, column=5).number_format = '#,##0.00'
    
    for col in range(4, 6):
        for row in [total_row, tax_row, final_row]:
            ws.cell(row=row, column=col).fill = PatternFill(start_color=PALETTE['total_bg'], end_color=PALETTE['total_bg'], fill_type='solid')
    
    # الملاحظات
    if notes:
        notes_row = final_row + 2
        ws.cell(row=notes_row, column=1, value="ملاحظات:")
        ws.cell(row=notes_row, column=1).font = Font(bold=True, size=SIZES['body'])
        ws.cell(row=notes_row + 1, column=1, value=notes)
        ws.cell(row=notes_row + 1, column=1).font = Font(size=SIZES['body'])
    
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

# ==========================================
# دالة شاملة لإنشاء خطاب احترافي
# ==========================================
def generate_professional_letter(
    organization_name,
    title,
    recipient,
    subject,
    paragraphs,
    table_headers=None,
    table_data=None,
    table_caption=None,
    closing_text="وتفضلوا بقبول فائق الاحترام والتقدير،،",
    signature_name=None,
    signature_title=None,
    note=None,
    output_dir=None,
    filename=None
):
    """دالة شاملة لإنشاء خطاب احترافي كامل - مع تنظيف تلقائي للترقيم"""
    from core.config import WORKSPACE_DIR
    import re
    
    doc = Document()
    setup_document(doc)
    add_elegant_header(doc, organization_name, title)
    add_document_info(doc, recipient=recipient, subject=subject)
    add_greeting(doc)
    
    # تنظيف الفقرات من الترقيم التلقائي
    cleaned_paragraphs = []
    for para_text in paragraphs:
        # إزالة أي رقم في بداية النص (عربي أو إنجليزي)
        cleaned = re.sub(r'^[\d\u0660-\u0669]+[\.\-\)]\s*', '', para_text.strip())
        cleaned_paragraphs.append(cleaned)
    
    # إضافة الفقرات المنظفة
    for para_text in cleaned_paragraphs:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        p.paragraph_format.first_line_indent = Cm(1.5)
        p.paragraph_format.line_spacing = 1.6
        p.paragraph_format.space_after = Pt(10)
        run = p.add_run(para_text)
        run.font.name = FONTS['body']
        run.font.size = Pt(SIZES['body'])
        run.font.color.rgb = RGBColor.from_string(PALETTE['charcoal'])
        set_paragraph_rtl(p)
    
    if table_headers and table_data:
        create_professional_table(doc, table_headers, table_data, table_caption)
    
    add_closing(doc, closing_text, signature_name, signature_title)
    
    if note:
        add_note(doc, note)
    
    add_footer(doc)
    
    save_dir = output_dir or WORKSPACE_DIR
    os.makedirs(save_dir, exist_ok=True)
    
    if filename is None:
        safe_title = title[:30].replace(' ', '_').replace('/', '-')
        filename = f"خطاب_{safe_title}"
    
    filename = get_unique_filename(save_dir, filename, 'docx')

    # ==========================================
# 📊 دالة شاملة لإنشاء أي تقرير مالي احترافي
# ==========================================
def generate_financial_report(
    report_type: str,
    company_name: str,
    period: str,
    sections: list,
    currency: str = "ريال",
    include_totals: bool = True,
    include_formulas: bool = True,
    style: str = "accounting",
    output_dir: str = None,
    filename: str = None
) -> str:
    """
    إنشاء أي تقرير مالي احترافي بمعايير محاسبية دولية
    
    report_type: نوع التقرير
        - 'income_statement': قائمة الدخل
        - 'cash_flow_direct': التدفقات النقدية (مباشرة)
        - 'cash_flow_indirect': التدفقات النقدية (غير مباشرة)
        - 'equity_statement': قائمة حقوق الملكية
        - 'balance_sheet': قائمة المركز المالي (الميزانية)
        - 'trial_balance': ميزان المراجعة
        - 'fixed_assets': قائمة الأصول الثابتة والإهلاكات
        - 'sales_report': تقرير المبيعات
        - 'expenses_report': تقرير المصروفات
        - 'account_statement': كشف حساب
        - 'bank_reconciliation': مطابقة بنكية
    
    sections: قائمة الأقسام، كل قسم:
    {
        'title': 'عنوان القسم',
        'headers': ['العمود1', 'العمود2', ...],
        'data': [['بيانات', 1000, 2000], ...],
        'is_total': False,  # هل هذا صف إجمالي؟
        'is_subtotal': False  # هل هذا صف إجمالي فرعي؟
    }
    
    style: نمط التنسيق
        - 'accounting': محاسبي كلاسيكي (أزرق داكن)
        - 'modern': حديث (أخضر/رمادي)
        - 'classic': تقليدي (أسود/أبيض)
    """
    from core.config import WORKSPACE_DIR
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, numbers
    from openpyxl.utils import get_column_letter
    from openpyxl.formatting.rule import CellIsRule
    
    wb = openpyxl.Workbook()
    if 'Sheet' in wb.sheetnames:
        del wb['Sheet']
    
    # ==========================================
    # أنماط التنسيق حسب النوع
    # ==========================================
    styles_config = {
        'accounting': {
            'header_fill': PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid"),
            'header_font': Font(name="Arial", size=12, bold=True, color="FFFFFF"),
            'subtotal_fill': PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid"),
            'subtotal_font': Font(name="Arial", size=11, bold=True, color="1F4E79"),
            'total_fill': PatternFill(start_color="2E5C8A", end_color="2E5C8A", fill_type="solid"),
            'total_font': Font(name="Arial", size=12, bold=True, color="FFFFFF"),
            'data_font': Font(name="Arial", size=10),
            'title_color': "1F4E79",
            'border_color': "B4C7E7"
        },
        'modern': {
            'header_fill': PatternFill(start_color="1E8449", end_color="1E8449", fill_type="solid"),
            'header_font': Font(name="Calibri", size=12, bold=True, color="FFFFFF"),
            'subtotal_fill': PatternFill(start_color="D5F5E3", end_color="D5F5E3", fill_type="solid"),
            'subtotal_font': Font(name="Calibri", size=11, bold=True, color="1E8449"),
            'total_fill': PatternFill(start_color="145A32", end_color="145A32", fill_type="solid"),
            'total_font': Font(name="Calibri", size=12, bold=True, color="FFFFFF"),
            'data_font': Font(name="Calibri", size=10),
            'title_color': "1E8449",
            'border_color': "A9DFBF"
        },
        'classic': {
            'header_fill': PatternFill(start_color="2C3E50", end_color="2C3E50", fill_type="solid"),
            'header_font': Font(name="Times New Roman", size=12, bold=True, color="FFFFFF"),
            'subtotal_fill': PatternFill(start_color="F2F3F4", end_color="F2F3F4", fill_type="solid"),
            'subtotal_font': Font(name="Times New Roman", size=11, bold=True, color="2C3E50"),
            'total_fill': PatternFill(start_color="1C2833", end_color="1C2833", fill_type="solid"),
            'total_font': Font(name="Times New Roman", size=12, bold=True, color="FFFFFF"),
            'data_font': Font(name="Times New Roman", size=10),
            'title_color': "2C3E50",
            'border_color': "5D6D7E"
        }
    }
    
    current_style = styles_config.get(style, styles_config['accounting'])
    
    # حدود الخلايا
    thin_border = Border(
        left=Side(style="thin", color=current_style['border_color']),
        right=Side(style="thin", color=current_style['border_color']),
        top=Side(style="thin", color=current_style['border_color']),
        bottom=Side(style="thin", color=current_style['border_color'])
    )
    
    # محاذاة
    center_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    right_align = Alignment(horizontal="right", vertical="center", wrap_text=True)
    left_align = Alignment(horizontal="left", vertical="center", wrap_text=True)
    
    # ==========================================
    # إنشاء ورقة التقرير الرئيسية
    # ==========================================
    ws = wb.active
    ws.title = report_type
    ws.sheet_view.rightToLeft = True
    
    # عنوان التقرير (مدمج)
    num_cols = max(len(sec['headers']) for sec in sections) if sections else 6
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=num_cols)
    title_cell = ws.cell(row=1, column=1)
    title_cell.value = f"{company_name}\n{report_type.replace('_', ' ').title()}\nعن الفترة: {period}"
    title_cell.font = Font(name="Arial", size=14, bold=True, color=current_style['title_color'])
    title_cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[1].height = 50
    
    # ==========================================
    # إضافة الأقسام والبيانات
    # ==========================================
    current_row = 3
    
    for sec_idx, section in enumerate(sections):
        # عنوان القسم
        ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=num_cols)
        section_title = ws.cell(row=current_row, column=1)
        section_title.value = section.get('title', '')
        section_title.font = Font(name="Arial", size=11, bold=True, italic=True, color=current_style['title_color'])
        section_title.alignment = left_align
        current_row += 1
        
        # رؤوس الأعمدة
        for col_idx, header in enumerate(section['headers'], 1):
            cell = ws.cell(row=current_row, column=col_idx, value=header)
            cell.fill = current_style['header_fill']
            cell.font = current_style['header_font']
            cell.alignment = center_align
            cell.border = thin_border
        ws.row_dimensions[current_row].height = 25
        header_row = current_row
        current_row += 1
        
        # بيانات القسم
        for row_data in section['data']:
            is_total_row = row_data[-1] if isinstance(row_data[-1], dict) and row_data[-1].get('is_total') else False
            is_subtotal_row = row_data[-1] if isinstance(row_data[-1], dict) and row_data[-1].get('is_subtotal') else False
            
            # استخراج البيانات الفعلية (بدون القاموس إن وجد)
            clean_data = [item for item in row_data if not isinstance(item, dict)]
            
            for col_idx, cell_data in enumerate(clean_data, 1):
                cell = ws.cell(row=current_row, column=col_idx, value=cell_data)
                cell.font = current_style['data_font']
                cell.alignment = right_align if col_idx > 1 else left_align
                cell.border = thin_border
                
                # تنسيق الأرقام المحاسبي
                if isinstance(cell_data, (int, float)) and col_idx > 1:
                    cell.number_format = '#,##0.00'
                
                # تنسيق صف الإجمالي الفرعي
                if is_subtotal_row:
                    cell.fill = current_style['subtotal_fill']
                    cell.font = current_style['subtotal_font']
                
                # تنسيق صف الإجمالي الكلي
                if is_total_row:
                    cell.fill = current_style['total_fill']
                    cell.font = current_style['total_font']
            
            current_row += 1
        
        # مسافة بين الأقسام
        current_row += 1
    
    # ==========================================
    # إضافة معادلات Excel التلقائية (إن طُلب)
    # ==========================================
    if include_formulas and include_totals:
        # يمكن إضافة معادلات SUM للأعمدة المالية هنا
        pass
    
    # ==========================================
    # تجميد الصفوف وتفعيل الفلاتر
    # ==========================================
    ws.freeze_panes = 'A4'
    ws.auto_filter.ref = f"A3:{get_column_letter(num_cols)}{current_row}"
    
    # ==========================================
    # ضبط عرض الأعمدة
    # ==========================================
    for col in range(1, num_cols + 1):
        ws.column_dimensions[get_column_letter(col)].width = 18
    
    # ==========================================
    # حفظ الملف
    # ==========================================
    save_dir = output_dir or WORKSPACE_DIR
    os.makedirs(save_dir, exist_ok=True)
    
    if filename is None:
        safe_company = company_name.replace(' ', '_')[:20]
        safe_type = report_type.replace('_', '_')
        filename = f"{safe_type}_{safe_company}_{period}"
    
    if not filename.endswith('.xlsx'):
        filename += '.xlsx'
    
    output_path = os.path.join(save_dir, filename)
    wb.save(output_path)
    
    return output_path
    output_path = os.path.join(save_dir, filename)
    
    doc.save(output_path)
    return output_path

# ==========================================
# 📝 دالة إنشاء خطاب بأنماط ديناميكية
# ==========================================
def generate_styled_letter(
    organization_name,
    title,
    recipient,
    subject,
    paragraphs,
    table_headers=None,
    table_data=None,
    table_caption=None,
    header_style='official',
    content_style='paragraphs',
    table_style='royal_blue',
    footer_style='professional',
    signature_style='single',
    signatures=None,
    output_dir=None,
    filename=None
):
    """إنشاء خطاب باستخدام نظام الأنماط الديناميكي"""
    from core.config import WORKSPACE_DIR
    
    doc = Document()
    setup_document(doc)
    
    apply_header_style(doc, header_style, organization_name, title)
    add_document_info(doc, recipient=recipient, subject=subject)
    add_greeting(doc)
    apply_content_style(doc, content_style, paragraphs)
    
    if table_headers and table_data:
        apply_table_style(doc, table_style, table_headers, table_data, table_caption)
    
    apply_signature_style(doc, signature_style, signatures)
    apply_footer_style(doc, footer_style)
    
    save_dir = output_dir or WORKSPACE_DIR
    os.makedirs(save_dir, exist_ok=True)
    
    if filename is None:
        safe_title = title[:30].replace(' ', '_').replace('/', '-')
        filename = f"خطاب_{safe_title}"
    
    filename = get_unique_filename(save_dir, filename, 'docx')
    output_path = os.path.join(save_dir, filename)
    
    doc.save(output_path)
    return output_path


# ==========================================
# 📊 دالة إنشاء تقرير مالي احترافي (مثل Fixed Asset Listing)
# ==========================================
def create_fixed_asset_report(
    company_name: str,
    year: str,
    sections: list,
    output_dir: str = None,
    filename: str = None
) -> str:
    """
    إنشاء تقرير أصول ثابتة / تقرير مالي متعدد الأوراق بمعايير محاسبية احترافية
    
    sections: قائمة تحتوي على أقسام التقرير، كل قسم عبارة عن قاموس:
    {
        'sheet_name': 'اسم_الورقة',
        'headers': ['رقم الأصل', 'الوصف', 'تاريخ الشراء', 'التكلفة', 'الإهلاك', 'الصافي'],
        'data': [['A001', 'سيارة', '2023-01-01', 150000, 30000, 120000], ...]
    }
    """
    from core.config import WORKSPACE_DIR
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, numbers
    from openpyxl.utils import get_column_letter
    
    wb = openpyxl.Workbook()
    
    # إزالة الورقة الافتراضية
    if 'Sheet' in wb.sheetnames:
        del wb['Sheet']
    
    # ألوان وتنسيقات احترافية
    header_fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
    header_font = Font(name="Arial", size=12, bold=True, color="FFFFFF")
    total_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    total_font = Font(name="Arial", size=11, bold=True, color="1F4E79")
    data_font = Font(name="Arial", size=10)
    center_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    right_align = Alignment(horizontal="right", vertical="center", wrap_text=True)
    
    thin_border = Border(
        left=Side(style="thin", color="B4C7E7"),
        right=Side(style="thin", color="B4C7E7"),
        top=Side(style="thin", color="B4C7E7"),
        bottom=Side(style="thin", color="B4C7E7")
    )

    # 1️⃣ إنشاء ورقة الملخص (Summary)
    summary_ws = wb.create_sheet(title="ملخص التقرير")
    summary_ws.sheet_view.rightToLeft = True
    
    # عنوان التقرير (خلايا مدمجة)
    summary_ws.merge_cells('A1:F1')
    title_cell = summary_ws['A1']
    title_cell.value = f"تقرير الأصول الثابتة - {company_name} ({year})"
    title_cell.font = Font(name="Arial", size=16, bold=True, color="1F4E79")
    title_cell.alignment = Alignment(horizontal="center", vertical="center")
    
    summary_ws.row_dimensions[1].height = 40
    
    # جدول الملخص
    summary_headers = ["القسم", "عدد الأصول", "إجمالي التكلفة", "إجمالي الإهلاك", "صافي القيمة"]
    summary_ws.append(summary_headers)
    
    # تنسيق رأس جدول الملخص
    for col_num, header in enumerate(summary_headers, 1):
        cell = summary_ws.cell(row=2, column=col_num)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = center_align
        cell.border = thin_border
    
    summary_ws.row_dimensions[2].height = 25
    
    # إضافة بيانات الملخص (حسابات تلقائية بسيطة)
    total_cost = 0
    total_depreciation = 0
    total_net = 0
    
    row_idx = 3
    for sec in sections:
        data = sec.get('data', [])
        cost = sum(float(row[3]) if len(row) > 3 and str(row[3]).replace('.','').isdigit() else 0 for row in data)
        dep = sum(float(row[4]) if len(row) > 4 and str(row[4]).replace('.','').isdigit() else 0 for row in data)
        net = cost - dep
        
        summary_ws.append([sec['sheet_name'], len(data), cost, dep, net])
        
        # تنسيق صف البيانات
        for col_num in range(1, 6):
            cell = summary_ws.cell(row=row_idx, column=col_num)
            cell.font = data_font
            cell.alignment = center_align if col_num == 2 else right_align
            cell.border = thin_border
            if col_num in [3, 4, 5]:
                cell.number_format = '#,##0.00'
        
        total_cost += cost
        total_depreciation += dep
        total_net += net
        row_idx += 1
    
    # صف الإجمالي النهائي
    summary_ws.append(["الإجمالي الكلي", "", total_cost, total_depreciation, total_net])
    for col_num in range(1, 6):
        cell = summary_ws.cell(row=row_idx, column=col_num)
        cell.fill = total_fill
        cell.font = total_font
        cell.alignment = center_align if col_num in [1, 2] else right_align
        cell.border = thin_border
        if col_num in [3, 4, 5]:
            cell.number_format = '#,##0.00'
            
    # ضبط عرض الأعمدة في الملخص
    for col in range(1, 6):
        summary_ws.column_dimensions[get_column_letter(col)].width = 18

    # 2️⃣ إنشاء الأوراق التفصيلية لكل قسم
    for sec in sections:
        ws = wb.create_sheet(title=sec['sheet_name'])
        ws.sheet_view.rightToLeft = True
        
        # عنوان الورقة
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(sec['headers']))
        title_cell = ws.cell(row=1, column=1)
        title_cell.value = f"تفاصيل: {sec['sheet_name']}"
        title_cell.font = Font(name="Arial", size=14, bold=True, color="1F4E79")
        title_cell.alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[1].height = 30
        
        # إضافة الرؤوس
        ws.append(sec['headers'])
        header_row = 2
        for col_num, header in enumerate(sec['headers'], 1):
            cell = ws.cell(row=header_row, column=col_num)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = center_align
            cell.border = thin_border
        ws.row_dimensions[header_row].height = 25
        
        # إضافة البيانات
        for row_data in sec['data']:
            ws.append(row_data)
            current_row = ws.max_row
            
            for col_num, cell_data in enumerate(row_data, 1):
                cell = ws.cell(row=current_row, column=col_num)
                cell.font = data_font
                cell.alignment = center_align if col_num == 1 else right_align
                cell.border = thin_border
                
                # تنسيق محاسبي للأعمدة المالية (افترض أن الأعمدة من 3 فما بعد هي أرقام)
                if col_num >= 3:
                    cell.number_format = '#,##0.00'
        
        # تجميد الصف العلوي وتفعيل الفلتر
        ws.freeze_panes = 'A3'
        ws.auto_filter.ref = f"A2:{get_column_letter(len(sec['headers']))}{ws.max_row}"
        
        # ضبط عرض الأعمدة
        for col in range(1, len(sec['headers']) + 1):
            ws.column_dimensions[get_column_letter(col)].width = 18

    # 3️⃣ حفظ الملف
    save_dir = output_dir or WORKSPACE_DIR
    os.makedirs(save_dir, exist_ok=True)
    
    if filename is None:
        safe_company = company_name.replace(' ', '_')[:20]
        filename = f"تقرير_الأصول_{safe_company}_{year}"
    
    # إضافة امتداد إذا لم يكن موجوداً
    if not filename.endswith('.xlsx'):
        filename += '.xlsx'
        
    output_path = os.path.join(save_dir, filename)
    wb.save(output_path)
    
    return output_path

# ==========================================
# 📄 دالة إنشاء سيرة ذاتية متوافقة مع ATS (موسعة)
# ==========================================
def generate_ats_cv(
    full_name,
    phone,
    email,
    location,
    linkedin,
    summary,
    experience,
    education,
    skills,
    languages=None,
    certifications=None,
    projects=None,
    interests=None,
    objective=None,
    output_dir=None,
    filename=None
):
    """
    إنشاء سيرة ذاتية متوافقة 100% مع أنظمة ATS
    - تنسيق بسيط ونظيف
    - خطوط قياسية
    - عناوين واضحة
    - بدون جداول معقدة
    """
    from core.config import WORKSPACE_DIR
    
    doc = Document()
    
    # إعدادات أساسية
    section = doc.sections[0]
    section.top_margin = Cm(1.5)
    section.bottom_margin = Cm(1.5)
    section.left_margin = Cm(2)
    section.right_margin = Cm(2)
    
    style = doc.styles['Normal']
    style.font.name = 'Calibri'
    style.font.size = Pt(11)
    style.font.color.rgb = RGBColor.from_string('000000')
    style.paragraph_format.line_spacing = 1.15
    style.paragraph_format.space_after = Pt(4)

    # 1. الاسم الكامل
    name_para = doc.add_paragraph()
    name_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    name_run = name_para.add_run(full_name.upper())
    name_run.font.name = 'Calibri'
    name_run.font.size = Pt(20)
    name_run.font.bold = True
    name_run.font.color.rgb = RGBColor.from_string('1B3A5C')

    # 2. معلومات الاتصال
    contact_para = doc.add_paragraph()
    contact_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    contact_text = f"{phone} | {email} | {location}"
    if linkedin:
        contact_text += f" | {linkedin}"
    contact_run = contact_para.add_run(contact_text)
    contact_run.font.name = 'Calibri'
    contact_run.font.size = Pt(10)
    contact_run.font.color.rgb = RGBColor.from_string('5D6D7E')
    
    doc.add_paragraph()

    # دالة مساعدة لعناوين الأقسام
    def add_ats_section_title(title_text):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(10)
        p.paragraph_format.space_after = Pt(4)
        p.paragraph_format.bottom_border = parse_xml(
            f'<w:pBdr {nsdecls("w")}>'
            f'  <w:bottom w:val="single" w:sz="4" w:space="1" w:color="1B3A5C"/>'
            f'</w:pBdr>'
        )
        run = p.add_run(title_text.upper())
        run.font.name = 'Calibri'
        run.font.size = Pt(12)
        run.font.bold = True
        run.font.color.rgb = RGBColor.from_string('1B3A5C')

    # دالة مساعدة للتحقق من اللغة العربية
    def is_arabic(text):
        return any('\u0600' <= c <= '\u06FF' for c in str(text))

    # 3. الهدف المهني (اختياري)
    if objective:
        add_ats_section_title("الهدف المهني" if is_arabic(objective) else "Career Objective")
        obj_para = doc.add_paragraph()
        obj_para.alignment = WD_ALIGN_PARAGRAPH.RIGHT if is_arabic(objective) else WD_ALIGN_PARAGRAPH.LEFT
        obj_run = obj_para.add_run(objective)
        obj_run.font.name = 'Calibri'
        obj_run.font.size = Pt(11)

    # 4. الملخص المهني
    add_ats_section_title("ملخص مهني" if is_arabic(summary) else "Professional Summary")
    summary_para = doc.add_paragraph()
    summary_para.alignment = WD_ALIGN_PARAGRAPH.RIGHT if is_arabic(summary) else WD_ALIGN_PARAGRAPH.LEFT
    summary_run = summary_para.add_run(summary)
    summary_run.font.name = 'Calibri'
    summary_run.font.size = Pt(11)

    # 5. الخبرة العملية
    if experience:
        add_ats_section_title("الخبرة العملية" if is_arabic(experience) else "Work Experience")
        for exp in experience:
            header_p = doc.add_paragraph()
            header_p.alignment = WD_ALIGN_PARAGRAPH.RIGHT if is_arabic(exp) else WD_ALIGN_PARAGRAPH.LEFT
            
            role_run = header_p.add_run(exp.get('role', ''))
            role_run.font.bold = True
            role_run.font.size = Pt(11)
            
            header_p.add_run(" | ")
            
            company_run = header_p.add_run(exp.get('company', ''))
            company_run.font.italic = True
            
            header_p.add_run(" | ")
            
            date_run = header_p.add_run(exp.get('date', ''))
            date_run.font.size = Pt(10)
            date_run.font.color.rgb = RGBColor.from_string('5D6D7E')
            
            for responsibility in exp.get('responsibilities', []):
                bullet_p = doc.add_paragraph()
                bullet_p.alignment = WD_ALIGN_PARAGRAPH.RIGHT if is_arabic(responsibility) else WD_ALIGN_PARAGRAPH.LEFT
                bullet_p.paragraph_format.left_indent = Cm(0.5)
                bullet_p.paragraph_format.space_after = Pt(2)
                bullet_run = bullet_p.add_run(f"• {responsibility}")
                bullet_run.font.name = 'Calibri'
                bullet_run.font.size = Pt(11)

    # 6. التعليم
    if education:
        add_ats_section_title("التعليم" if is_arabic(education) else "Education")
        for edu in education:
            edu_p = doc.add_paragraph()
            edu_p.alignment = WD_ALIGN_PARAGRAPH.RIGHT if is_arabic(edu) else WD_ALIGN_PARAGRAPH.LEFT
            
            degree_run = edu_p.add_run(edu.get('degree', ''))
            degree_run.font.bold = True
            
            edu_p.add_run(" - ")
            
            uni_run = edu_p.add_run(edu.get('university', ''))
            uni_run.font.italic = True
            
            edu_p.add_run(" | ")
            
            year_run = edu_p.add_run(edu.get('year', ''))
            year_run.font.size = Pt(10)
            year_run.font.color.rgb = RGBColor.from_string('5D6D7E')

    # 7. المهارات (كنقاط)
    if skills:
        add_ats_section_title("المهارات" if is_arabic(skills) else "Skills")
        for skill in skills:
            skill_p = doc.add_paragraph()
            skill_p.alignment = WD_ALIGN_PARAGRAPH.RIGHT if is_arabic(skill) else WD_ALIGN_PARAGRAPH.LEFT
            skill_p.paragraph_format.space_after = Pt(2)
            skill_run = skill_p.add_run(f"• {skill}")
            skill_run.font.name = 'Calibri'
            skill_run.font.size = Pt(11)

    # 8. اللغات (جديد)
    if languages:
        add_ats_section_title("اللغات" if is_arabic(languages) else "Languages")
        for lang in languages:
            lang_p = doc.add_paragraph()
            lang_p.alignment = WD_ALIGN_PARAGRAPH.RIGHT if is_arabic(lang) else WD_ALIGN_PARAGRAPH.LEFT
            lang_p.paragraph_format.space_after = Pt(2)
            
            name_run = lang_p.add_run(lang.get('name', ''))
            name_run.font.bold = True
            lang_p.add_run(" - ")
            level_run = lang_p.add_run(lang.get('level', ''))
            level_run.font.size = Pt(10)

    # 9. الشهادات (جديد)
    if certifications:
        add_ats_section_title("الشهادات" if is_arabic(certifications) else "Certifications")
        for cert in certifications:
            cert_p = doc.add_paragraph()
            cert_p.alignment = WD_ALIGN_PARAGRAPH.RIGHT if is_arabic(cert) else WD_ALIGN_PARAGRAPH.LEFT
            cert_p.paragraph_format.space_after = Pt(2)
            
            name_run = cert_p.add_run(cert.get('name', ''))
            name_run.font.bold = True
            cert_p.add_run(" | ")
            issuer_run = cert_p.add_run(cert.get('issuer', ''))
            cert_p.add_run(" | ")
            year_run = cert_p.add_run(cert.get('year', ''))
            year_run.font.size = Pt(10)
            year_run.font.color.rgb = RGBColor.from_string('5D6D7E')

    # 10. المشاريع (جديد)
    if projects:
        add_ats_section_title("المشاريع" if is_arabic(projects) else "Projects")
        for proj in projects:
            proj_p = doc.add_paragraph()
            proj_p.alignment = WD_ALIGN_PARAGRAPH.RIGHT if is_arabic(proj) else WD_ALIGN_PARAGRAPH.LEFT
            
            name_run = proj_p.add_run(proj.get('name', ''))
            name_run.font.bold = True
            proj_p.add_run(" | ")
            date_run = proj_p.add_run(proj.get('date', ''))
            date_run.font.size = Pt(10)
            date_run.font.color.rgb = RGBColor.from_string('5D6D7E')
            
            if proj.get('description'):
                desc_p = doc.add_paragraph()
                desc_p.alignment = WD_ALIGN_PARAGRAPH.RIGHT if is_arabic(proj['description']) else WD_ALIGN_PARAGRAPH.LEFT
                desc_p.paragraph_format.left_indent = Cm(0.5)
                desc_run = desc_p.add_run(proj['description'])
                desc_run.font.size = Pt(11)

    # 11. الاهتمامات (جديد)
    if interests:
        add_ats_section_title("الاهتمامات" if is_arabic(interests) else "Interests")
        interests_p = doc.add_paragraph()
        interests_p.alignment = WD_ALIGN_PARAGRAPH.RIGHT if is_arabic(interests) else WD_ALIGN_PARAGRAPH.LEFT
        interests_text = " • ".join(interests)
        interests_run = interests_p.add_run(interests_text)
        interests_run.font.name = 'Calibri'
        interests_run.font.size = Pt(11)

    # الحفظ
    save_dir = output_dir or WORKSPACE_DIR
    os.makedirs(save_dir, exist_ok=True)
    
    if filename is None:
        safe_name = full_name.replace(' ', '_')
        filename = f"CV_{safe_name}"
    
    filename = get_unique_filename(save_dir, filename, 'docx')
    output_path = os.path.join(save_dir, filename)
    
    doc.save(output_path)
    return output_path
