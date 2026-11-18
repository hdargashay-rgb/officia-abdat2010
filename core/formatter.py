"""
Officia Formatter Pro v5 - نظام التنسيق الاحترافي المتقدم
دوال تنسيق ذكية متعددة الاستخدام

الإصدار 5: كل دوال Excel المحاسبية تعتمد على table_engine (معادلات حية + أعمدة مرنة)،
مع حذف الدوال المكررة الميتة وإصلاح الدوال المبتورة.
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
try:  # الملفان داخل حزمة core (الموصى به)
    from core.table_engine import (_num, col, resolve_columns, write_table, apply_widths, warn_ignored, print_warnings,
                          sheet_ref, apply_header_formulas, add_totals_by_header)
    from core.formula_helpers import finalize_workbook
except ImportError:  # أو بجانب formatter.py مباشرة
    from table_engine import (_num, col, resolve_columns, write_table, apply_widths, warn_ignored, print_warnings,
                          sheet_ref, apply_header_formulas, add_totals_by_header)
    from formula_helpers import finalize_workbook

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
    'total_bg': 'D9E1F2',     # خلفية صفوف الإجمالي (كان مفقوداً فيسبب KeyError في الفاتورة)
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
                       auto_filter=True, table_style=True,
                       formulas=None, totals=None, total_label="الإجمالي"):
    """
    إنشاء ورقة Excel منسقة باحترافية كاملة
    
    تطبق: RTL، عناوين ملونة، حدود، auto-fit، تجميد، فلاتر

    formulas: معادلات أعمدة بعناوينها، مثل {'الصافي': '={التكلفة}-{الخصم}'}
              (عمود غير موجود يُنشأ في آخر الجدول)
    totals  : صف إجمالي بمعادلات SUM: True أو 'auto' لكل الأعمدة الرقمية، أو قائمة عناوين
    """
    ws.sheet_view.rightToLeft = True
    ws.sheet_view.showGridLines = True

    # أعمدة المعادلات غير الموجودة تُضاف إلى العناوين
    headers = list(headers)
    for _target in (formulas or {}):
        if _target not in headers:
            headers.append(_target)
    
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
    
    # معادلات الأعمدة وصف الإجمالي (حيّة، تتحدث مع أي تعديل)
    last_data_row = data_start_row + len(data) - 1
    if formulas and len(data) > 0:
        apply_header_formulas(ws, header_row, data_start_row, last_data_row, formulas)
        for _r in range(data_start_row, last_data_row + 1):
            for _target in formulas:
                _c = ws.cell(row=_r, column=headers.index(_target) + 1)
                _c.font, _c.alignment, _c.border = data_font, data_alignment, data_border
                _c.fill = even_fill if (_r - data_start_row) % 2 == 0 else odd_fill
    if totals and len(data) > 0:
        add_totals_by_header(ws, header_row, data_start_row, last_data_row,
                             "auto" if totals is True else totals, total_label)

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
                if cell.value and not str(cell.value).startswith('='):
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
    - formulas: معادلات أعمدة بعناوينها (اختياري)  مثل {'الصافي': '={التكلفة}-{الخصم}'}
    - totals: صف إجمالي (اختياري): True أو قائمة عناوين
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
            auto_filter=True,
            formulas=sheet_cfg.get('formulas'),
            totals=sheet_cfg.get('totals'),
            total_label=sheet_cfg.get('total_label', "الإجمالي")
        )

    save_dir = output_dir or WORKSPACE_DIR
    os.makedirs(save_dir, exist_ok=True)

    if filename is None:
        filename = f"Report_{datetime.now().strftime('%Y%m%d')}"

    filename = get_unique_filename(save_dir, filename, 'xlsx')
    output_path = os.path.join(save_dir, filename)

    finalize_workbook(wb)
    wb.save(output_path)

    # ✅ طبقة الحماية النهائية: ضمان طباعة مسار الملف بنجاح
    print("SUCCESS: " + output_path)

    return output_path

    # ==========================================
# دالة شاملة لإنشاء خطاب احترافي
# ==========================================


INVOICE_COLUMNS = [
    col("idx", "#", "index", width=5),
    col("description", "الوصف", "text", width=40),
    col("quantity", "الكمية", width=12),
    col("unit_price", "سعر الوحدة", width=15),
    col("line_total", "الإجمالي", formula="={quantity}*{unit_price}", width=15),
]


def create_professional_invoice(
    company_name,
    client_name,
    invoice_number,
    invoice_date,
    items,
    tax_rate=0.05,
    notes=None,
    output_dir=None,
    filename=None,
    columns=None,
    add_columns=None,
    remove_columns=None,
    rename_columns=None,
    **kwargs
):
    """
    إنشاء فاتورة احترافية (إجمالي البند والمجموع والضريبة والنهائي كلها معادلات حية)

    items: قائمة من القواميس: description, quantity, unit_price (+ أي حقول لأعمدة مضافة)

    الضريبة: نسبتها خلية قابلة للتعديل (B5) وتتحدث بها الضريبة والإجمالي النهائي.
    مرونة الأعمدة: add_columns=[{'key': 'discount', 'title': 'الخصم'}] يُطرح تلقائياً من إجمالي البند،
    ويمكن إضافة أي عمود آخر نصي أو رقمي، وحذف أي عمود بـ remove_columns.
    """
    from core.config import WORKSPACE_DIR

    warn_ignored("create_professional_invoice", kwargs)

    cols, col_warns = resolve_columns(columns or INVOICE_COLUMNS, add_columns, remove_columns, rename_columns)
    keys = [c['key'] for c in cols]
    # الخصم عمود قيمة يُطرح من إجمالي البند
    for c in cols:
        if c['key'] == 'line_total' and 'discount' in keys and 'discount' not in c['formula']:
            c['formula'] = c['formula'] + "-{discount}"
    n = len(cols)
    if n < 2:
        raise ValueError("❌ الفاتورة تحتاج عمودين على الأقل.")

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "الفاتورة"
    ws.sheet_view.rightToLeft = True

    # عنوان الفاتورة
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=n)
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

    # نسبة الضريبة كخلية مدخلات (تغييرها يعيد حساب الفاتورة)
    ws['A5'] = "نسبة الضريبة:"
    ws['A5'].font = Font(bold=True, size=SIZES['body'])
    rate_cell = ws['B5']
    rate_cell.value = float(tax_rate or 0)
    rate_cell.number_format = '0.0%'
    rate_cell.font = Font(bold=True, size=SIZES['body'], color="0000FF")
    rate_cell.alignment = Alignment(horizontal='right')
    rate_cell.fill = PatternFill(start_color="FFF9C4", end_color="FFF9C4", fill_type='solid')

    # جدول البنود
    header_row = 6
    inv_styles = {
        "header_fill": PatternFill(start_color=PALETTE['royal_dark'], end_color=PALETTE['royal_dark'], fill_type='solid'),
        "header_font": Font(bold=True, color=PALETTE['white'], size=SIZES['body']),
        "header_border": Border(
            left=Side(style='thin', color=PALETTE['silver']), right=Side(style='thin', color=PALETTE['silver']),
            top=Side(style='thin', color=PALETTE['silver']), bottom=Side(style='medium', color=PALETTE['gold'])),
        "data_font": Font(size=SIZES['body']),
        "data_border": Border(
            left=Side(style='thin', color=PALETTE['silver']), right=Side(style='thin', color=PALETTE['silver']),
            top=Side(style='thin', color=PALETTE['silver']), bottom=Side(style='thin', color=PALETTE['silver'])),
    }
    info = write_table(ws, cols, items or [], header_row=header_row, styles=inv_styles,
                       total=False, header_height=28)
    apply_widths(ws, cols)
    first, last = info['first_row'], info['last_row']
    L = info['letters']

    # أعمدة الألوان المتناوبة كما في التصميم الأصلي
    for r in range(first, last + 1):
        bg = PALETTE['row_even'] if (r - first + 1) % 2 == 0 else PALETTE['cream']
        for c in range(1, n + 1):
            ws.cell(row=r, column=c).fill = PatternFill(start_color=bg, end_color=bg, fill_type='solid')

    # الإجماليات: كلها معادلات
    amount_col = n
    label_col = n - 1
    amt = get_column_letter(amount_col)
    sum_src = L.get('line_total', amt)
    total_row = last + 1
    tax_row = total_row + 1
    final_row = tax_row + 1

    ws.cell(row=total_row, column=label_col, value="الإجمالي الفرعي:").font = Font(bold=True, size=SIZES['body'])
    ws.cell(row=total_row, column=amount_col,
            value=f"=SUM({sum_src}{first}:{sum_src}{last})" if last >= first else 0)

    ws.cell(row=tax_row, column=label_col,
            value='="الضريبة ("&ROUND($B$5*100,2)&"%):"').font = Font(bold=True, size=SIZES['body'])
    ws.cell(row=tax_row, column=amount_col, value=f"={amt}{total_row}*$B$5")

    ws.cell(row=final_row, column=label_col, value="الإجمالي النهائي:").font = Font(
        bold=True, size=SIZES['heading'], color=PALETTE['royal_dark'])
    ws.cell(row=final_row, column=amount_col, value=f"={amt}{total_row}+{amt}{tax_row}")

    for r in (total_row, tax_row, final_row):
        big = r == final_row
        f = Font(bold=True, size=SIZES['heading'] if big else SIZES['body'],
                 color=PALETTE['royal_dark'] if big else None)
        ws.cell(row=r, column=amount_col).font = f
        ws.cell(row=r, column=amount_col).number_format = '#,##0.00'
        ws.cell(row=r, column=amount_col).alignment = Alignment(horizontal='right')
        for c in (label_col, amount_col):
            ws.cell(row=r, column=c).fill = PatternFill(
                start_color=PALETTE['total_bg'], end_color=PALETTE['total_bg'], fill_type='solid')

    # الملاحظات
    if notes:
        notes_row = final_row + 2
        ws.cell(row=notes_row, column=1, value="ملاحظات:")
        ws.cell(row=notes_row, column=1).font = Font(bold=True, size=SIZES['body'])
        ws.cell(row=notes_row + 1, column=1, value=notes)
        ws.cell(row=notes_row + 1, column=1).font = Font(size=SIZES['body'])

    ws.freeze_panes = f'A{first}'
    print_warnings("create_professional_invoice", col_warns + info['warnings'])

    # الحفظ
    save_dir = output_dir or WORKSPACE_DIR
    os.makedirs(save_dir, exist_ok=True)
    if filename is None:
        filename = f"فاتورة_{invoice_number}"
    filename = get_unique_filename(save_dir, filename, 'xlsx')
    output_path = os.path.join(save_dir, filename)
    finalize_workbook(wb)
    wb.save(output_path)

    # ✅ طبقة الحماية النهائية: ضمان طباعة مسار الملف بنجاح
    print("SUCCESS: " + output_path)
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
    filename=None,
    **kwargs
):
    """دالة شاملة لإنشاء خطاب احترافي كامل - مع تنظيف تلقائي للترقيم"""
    from core.config import WORKSPACE_DIR

    if isinstance(paragraphs, str):
        paragraphs = [paragraphs]

    doc = Document()
    setup_document(doc)
    add_elegant_header(doc, organization_name, title)
    add_document_info(doc, recipient=recipient, subject=subject)
    add_greeting(doc)

    # تنظيف الفقرات من الترقيم التلقائي (رقم عربي أو إنجليزي في بداية النص)
    cleaned_paragraphs = [
        re.sub(r'^[\d\u0660-\u0669]+[\.\-\)]\s*', '', str(para_text).strip())
        for para_text in paragraphs
    ]

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
    output_path = os.path.join(save_dir, filename)

    doc.save(output_path)

    # ✅ طبقة الحماية النهائية: ضمان طباعة مسار الملف بنجاح
    print("SUCCESS: " + output_path)

    return output_path

    # ==========================================
# 📊 دالة شاملة لإنشاء أي تقرير مالي احترافي
# ==========================================
# ==========================================
# 📊 دالة التقارير المالية - الإصدار النهائي (v5.1)
# بنية: ورقة ملخص في البداية + أوراق تفاصيل + معادلات ربط حية
# الأنواع غير المعرّفة بقالب (أو التي لا تطابق أقسامها القالب) تُنشأ بالمسار العام _generate_generic_financial_report
# ==========================================

# قوالب التقارير المحاسبية
# aliases: أسماء بديلة تُطابَق بها عناوين الأقسام. key في القسم نفسه يربطه بالقالب مباشرة.
ACCOUNTING_TEMPLATES = {
    "income_statement": {
        "name_ar": "قائمة الدخل",
        "first_sheet": "قائمة الدخل",
        "sections": [
            {"key": "revenue", "title": "الإيرادات", "sheet": "الإيرادات",
             "aliases": ["المبيعات", "إيرادات المبيعات", "الإيرادات التشغيلية"]},
            {"key": "sales_returns", "title": "مردودات المبيعات", "sheet": "مردودات المبيعات", "optional": True,
             "aliases": ["مردودات ومسموحات المبيعات"]},
            {"key": "net_revenue", "title": "صافي الإيرادات", "is_formula": True, "formula": "{revenue} - {sales_returns}"},
            {"key": "cogs", "title": "تكلفة المبيعات", "sheet": "تكلفة المبيعات", "optional": True,
             "aliases": ["تكلفة البضاعة المباعة", "تكلفة الإيرادات"]},
            {"key": "gross_profit", "title": "مجمل الربح", "is_formula": True, "formula": "{net_revenue} - {cogs}"},
            {"key": "operating_expenses", "title": "المصروفات التشغيلية", "sheet": "المصروفات التشغيلية",
             "aliases": ["المصروفات", "المصاريف", "المصاريف التشغيلية", "مصروفات تشغيلية", "المصروفات الإدارية والعمومية"]},
            {"key": "operating_profit", "title": "الربح التشغيلي", "is_formula": True, "formula": "{gross_profit} - {operating_expenses}"},
            {"key": "other_income", "title": "الإيرادات الأخرى", "sheet": "الإيرادات الأخرى", "optional": True},
            {"key": "other_expenses", "title": "المصروفات الأخرى", "sheet": "المصروفات الأخرى", "optional": True},
            {"key": "profit_before_tax", "title": "الربح قبل الضريبة", "is_formula": True,
             "formula": "{operating_profit} + {other_income} - {other_expenses}"},
            {"key": "tax", "title": "الضريبة", "sheet": "الضريبة", "optional": True,
             "aliases": ["ضريبة الدخل", "مصروف الضريبة", "الزكاة والضريبة"]},
            {"key": "net_profit", "title": "صافي الربح", "is_formula": True, "formula": "{profit_before_tax} - {tax}"},
        ],
    },
    "balance_sheet": {
        "name_ar": "قائمة المركز المالي",
        "first_sheet": "الميزانية",
        "sections": [
            {"key": "current_assets", "title": "الأصول المتداولة", "sheet": "الأصول المتداولة"},
            {"key": "non_current_assets", "title": "الأصول غير المتداولة", "sheet": "الأصول غير المتداولة",
             "aliases": ["الأصول الثابتة"]},
            {"key": "total_assets", "title": "إجمالي الأصول", "is_formula": True, "formula": "{current_assets} + {non_current_assets}"},
            {"key": "current_liabilities", "title": "الخصوم المتداولة", "sheet": "الخصوم المتداولة",
             "aliases": ["الخصوم", "الالتزامات", "الالتزامات المتداولة"]},
            {"key": "non_current_liabilities", "title": "الخصوم غير المتداولة", "sheet": "الخصوم غير المتداولة", "optional": True,
             "aliases": ["الالتزامات غير المتداولة"]},
            {"key": "total_liabilities", "title": "إجمالي الخصوم", "is_formula": True, "formula": "{current_liabilities} + {non_current_liabilities}"},
            {"key": "equity", "title": "حقوق الملكية", "sheet": "حقوق الملكية"},
            {"key": "total_equity_liabilities", "title": "إجمالي الخصوم وحقوق الملكية", "is_formula": True, "formula": "{total_liabilities} + {equity}"},
            {"key": "balance_check", "title": "فحص التوازن", "is_check": True, "formula": "{total_assets} - {total_equity_liabilities}"},
        ],
    },
    "cash_flow": {
        "name_ar": "قائمة التدفقات النقدية",
        "first_sheet": "التدفقات النقدية",
        "sections": [
            {"key": "operating", "title": "الأنشطة التشغيلية", "sheet": "التدفقات التشغيلية",
             "aliases": ["التدفقات التشغيلية", "التدفقات النقدية من الأنشطة التشغيلية"]},
            {"key": "investing", "title": "الأنشطة الاستثمارية", "sheet": "التدفقات الاستثمارية", "optional": True,
             "aliases": ["التدفقات الاستثمارية", "التدفقات النقدية من الأنشطة الاستثمارية"]},
            {"key": "financing", "title": "الأنشطة التمويلية", "sheet": "التدفقات التمويلية", "optional": True,
             "aliases": ["التدفقات التمويلية", "التدفقات النقدية من الأنشطة التمويلية"]},
            {"key": "net_change", "title": "صافي التغير في النقدية", "is_formula": True, "formula": "{operating} + {investing} + {financing}"},
        ],
    },
    "generic": {"name_ar": "تقرير مالي", "first_sheet": "التقرير", "sections": []},
}


def _detect_template(report_type: str) -> str:
    """يكتشف نوع التقرير من الاسم (الأنواع الأخرى مثل ميزان المراجعة والمطابقة البنكية = generic)."""
    rt = str(report_type).lower().replace('-', '_').replace(' ', '_')
    if any(w in rt for w in ('trial', 'reconcil', 'تجريبي', 'مراجعة', 'مطابقة', 'fixed_asset', 'sales', 'expenses', 'account_statement')):
        return 'generic'
    if 'cash' in rt or 'التدفقات' in rt or 'تدفق' in rt:
        return 'cash_flow'
    if 'income' in rt or 'pnl' in rt or 'profit' in rt or 'الدخل' in rt:
        return 'income_statement'
    if 'balance' in rt or 'financial_position' in rt or 'الميزانية' in rt or 'المركز_المالي' in rt:
        return 'balance_sheet'
    return 'generic'


def _safe_sheet_name(name: str) -> str:
    """ينظف اسم الورقة ليكون صالحاً في Excel."""
    name = str(name).strip().replace('/', '-').replace('\\', '-')
    for ch in "[]*?:'":
        name = name.replace(ch, '')
    name = name.strip()
    return name[:31] if name else "ورقة"


def _norm_title(s) -> str:
    """توحيد العنوان للمطابقة (مسافات + همزات + تاء مربوطة)."""
    s = re.sub(r'\s+', ' ', str(s)).strip()
    for a, b in (('أ', 'ا'), ('إ', 'ا'), ('آ', 'ا'), ('ة', 'ه'), ('ى', 'ي')):
        s = s.replace(a, b)
    return s


def _build_excel_formula(formula_template: str, cell_refs: dict) -> str:
    """
    يحول قالب معادلة إلى معادلة Excel فعلية (المراجع بلا علامة = في أولها).
    مثال: "{revenue} - {cogs}" → "=B7-B9"
    """
    def replace_ref(match):
        return cell_refs.get(match.group(1), "0")
    return "=" + re.sub(r'\{(\w+)\}', replace_ref, formula_template)


def _match_sections_to_template(sections, template_sections):
    """
    يربط أقسام المستخدم بمفاتيح القالب. الأولوية: key صريح ← تطابق تام (عنوان/اسم ورقة/بديل)
    ← احتواء (أطول اسم قالب موجود داخل عنوان القسم). كل مفتاح يُربط بقسم واحد.
    يرجع (قاموس مفتاح→فهرس القسم، فهارس الأقسام غير المرتبطة).
    """
    candidates = [t for t in template_sections if not t.get('is_formula') and not t.get('is_check')]
    result, used = {}, set()

    for i, sec in enumerate(sections):                      # 1) key صريح
        k = sec.get('key')
        if k and k not in result and any(t['key'] == k for t in candidates):
            result[k] = i
            used.add(i)

    for i, sec in enumerate(sections):                      # 2) تطابق تام
        if i in used:
            continue
        nt = _norm_title(sec.get('title', ''))
        for t in candidates:
            if t['key'] in result:
                continue
            names = {_norm_title(t['title']), _norm_title(t.get('sheet', ''))}
            names |= {_norm_title(a) for a in t.get('aliases', [])}
            if nt in names:
                result[t['key']] = i
                used.add(i)
                break

    for i, sec in enumerate(sections):                      # 3) احتواء (الأطول أولاً)
        if i in used:
            continue
        nt = _norm_title(sec.get('title', ''))
        best = None
        for t in candidates:
            if t['key'] in result:
                continue
            for name in [t['title']] + t.get('aliases', []):
                n = _norm_title(name)
                if n and n in nt and (best is None or len(n) > best[0]):
                    best = (len(n), t['key'])
        if best:
            result[best[1]] = i
            used.add(i)

    return result, [i for i in range(len(sections)) if i not in used]


def _fin_styles(style: str) -> dict:
    """أنماط التقرير: accounting (أزرق) / modern (أخضر) / classic (رمادي داكن)."""
    main, pale, border = {
        'accounting': ("1F4E79", "D6E4F0", "B4C7E7"),
        'modern': ("1E8449", "D5F5E3", "A9DFBF"),
        'classic': ("1C2833", "F2F3F4", "5D6D7E"),
    }.get(style, ("1F4E79", "D6E4F0", "B4C7E7"))
    solid = lambda c: PatternFill(start_color=c, end_color=c, fill_type="solid")
    side = lambda s: Side(style=s, color=border)
    return {
        'title_font': Font(name="Arial", size=18, bold=True, color=main),
        'subtitle_font': Font(name="Arial", size=14, bold=True, color=main),
        'period_font': Font(name="Arial", size=12, italic=True, color="5D6D7E"),
        'header_fill': solid(main),
        'header_font': Font(name="Arial", size=12, bold=True, color="FFFFFF"),
        'formula_row_fill': solid(pale),
        'formula_row_font': Font(name="Arial", size=11, bold=True, color=main),
        'total_fill': solid(main),
        'total_font': Font(name="Arial", size=13, bold=True, color="FFFFFF"),
        'data_font': Font(name="Arial", size=11),
        'thin': Border(left=side("thin"), right=side("thin"), top=side("thin"), bottom=side("thin")),
        'thick': Border(left=side("medium"), right=side("medium"), top=side("medium"), bottom=side("double")),
    }


def _write_fin_detail_sheet(wb, section, idx, S):
    """ورقة تفاصيل قسم واحد: بنود + صف إجمالي بمعادلات SUM. يرجع (الورقة، صف الإجمالي أو None، عدد الأعمدة)."""
    title = section.get('title', f'قسم {idx + 1}')
    ws = wb.create_sheet(title=_safe_sheet_name(title))     # openpyxl يضمن فرادة الاسم، فنقرأ ws.title دائماً
    ws.sheet_view.rightToLeft = True

    headers = list(section.get('headers') or ['البيان', 'المبلغ'])
    rows = []
    for row_data in section.get('data', []):
        if not isinstance(row_data, (list, tuple)):
            continue
        flags = row_data[-1] if row_data and isinstance(row_data[-1], dict) else {}
        if flags.get('is_total') or flags.get('is_subtotal'):
            continue            # الإجمالي تحسبه الورقة نفسها بمعادلة؛ إبقاؤه يضاعف الجمع
        rows.append([v for v in row_data if not isinstance(v, dict)])
    ncols = max([len(headers)] + [len(r) for r in rows])
    headers += ['المبلغ'] * (ncols - len(headers))

    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=max(ncols, 2))
    c = ws.cell(row=1, column=1, value=f"تفاصيل: {title}")
    c.font, c.alignment = S['title_font'], Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 35

    for i, h in enumerate(headers, 1):
        c = ws.cell(row=2, column=i, value=h)
        c.fill, c.font, c.border = S['header_fill'], S['header_font'], S['thick']
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[2].height = 30

    first = 3
    for n, row in enumerate(rows):
        r = first + n
        for ci, val in enumerate(row, 1):
            val = _num(val) if ci > 1 else val
            c = ws.cell(row=r, column=ci, value=val)
            c.font, c.border = S['data_font'], S['thin']
            c.alignment = Alignment(horizontal="right" if ci == 1 else "center", vertical="center", wrap_text=True)
            if ci > 1 and isinstance(val, (int, float)) and not isinstance(val, bool):
                c.number_format = '#,##0.00'
        ws.row_dimensions[r].height = 24
    last = first + len(rows) - 1

    total_row = None
    if rows:
        total_row = last + 1
        lc = ws.cell(row=total_row, column=1, value="الإجمالي")
        lc.font, lc.fill, lc.border = S['total_font'], S['total_fill'], S['thick']
        lc.alignment = Alignment(horizontal="right", vertical="center")
        for ci in range(2, ncols + 1):
            L = get_column_letter(ci)
            c = ws.cell(row=total_row, column=ci, value=f"=SUM({L}{first}:{L}{last})")
            c.font, c.fill, c.border = S['total_font'], S['total_fill'], S['thick']
            c.alignment = Alignment(horizontal="center", vertical="center")
            c.number_format = '#,##0.00'
        ws.row_dimensions[total_row].height = 28

    ws.column_dimensions['A'].width = 45
    for ci in range(2, ncols + 1):
        ws.column_dimensions[get_column_letter(ci)].width = 22
    ws.freeze_panes = 'A3'
    return ws, total_row, ncols


# ==========================================
# 📊 الدالة الرئيسية
# ==========================================
def generate_financial_report(
    report_type: str = "تقرير_مالي",
    company_name: str = "شركة",
    period: str = "2026",
    sections: list = None,
    currency: str = "ر.س",
    include_totals: bool = True,
    include_formulas: bool = True,
    style: str = "accounting",
    output_dir: str = None,
    filename: str = None,
    **kwargs
) -> str:
    """
    إنشاء تقرير مالي احترافي.

    قائمة الدخل / المركز المالي / التدفقات النقدية: ورقة ملخص أولى + ورقة تفاصيل لكل قسم،
    والملخص كله معادلات تربط بإجماليات التفاصيل (تتحدث مع أي تعديل). أعمدة المقارنة
    (السنة الحالية / السابقة...) مدعومة: كل عمود قيمة في التفاصيل يظهر في الملخص.

    sections: [{'title': 'الإيرادات', 'headers': ['البند', 'السنة الحالية', 'السنة السابقة'],
                'data': [['مبيعات', 1000, 800], ...], 'key': 'revenue' (اختياري: ربط صريح بالقالب)}]
      - يُربط القسم بالقالب بالعنوان (أو بدائله أو key). أي قسم غير مربوط يبقى ورقة تفصيلية فقط.
      - صفوف الإجمالي المرفقة بالبيانات (is_total/is_subtotal) تُهمل لأن الإجمالي معادلة.
      - فرضية: المصروفات والتكاليف تُدخل بقيم موجبة، والقالب يطرحها.

    الأنواع الأخرى (ميزان مراجعة، أصول ثابتة، مبيعات...) أو إن لم يطابق أي قسم القالب:
    تُنشأ بالمسار العام (ورقة واحدة، صفوف إجمالي بمعادلات، ومعاملات id/formula).
    """
    # 🛡️ طبقة الحماية
    if sections is None:
        sections = []
    if isinstance(sections, dict):
        sections = [sections]
    if not isinstance(sections, list):
        sections = []
    cleaned = []
    for sec in sections:
        if not isinstance(sec, dict):
            continue
        sec = dict(sec)
        sec.setdefault('title', 'قسم بدون عنوان')
        sec.setdefault('headers', ['البيان', 'المبلغ'])
        if not isinstance(sec.get('data'), list):
            sec['data'] = []
        cleaned.append(sec)
    sections = cleaned

    generic_args = dict(report_type=report_type, company_name=company_name, period=period,
                        sections=sections, currency=currency, include_totals=include_totals,
                        include_formulas=include_formulas, style=style, output_dir=output_dir,
                        filename=filename, **kwargs)

    template_key = _detect_template(report_type)
    template = ACCOUNTING_TEMPLATES.get(template_key, ACCOUNTING_TEMPLATES['generic'])
    template_sections = template.get('sections', [])
    if template_key == 'generic' or not sections:
        return _generate_generic_financial_report(**generic_args)

    matched, unmatched = _match_sections_to_template(sections, template_sections)
    if not matched:
        print(f"⚠️ generate_financial_report: لم يطابق أي قسم قالب «{template['name_ar']}»؛ "
              f"استُخدم المسار العام. (عناوين القالب: "
              f"{', '.join(t['title'] for t in template_sections if not t.get('is_formula') and not t.get('is_check'))})")
        return _generate_generic_financial_report(**generic_args)

    try:
        from core.config import WORKSPACE_DIR
    except ImportError:
        WORKSPACE_DIR = "/tmp"

    S = _fin_styles(style)
    wb = openpyxl.Workbook()
    if 'Sheet' in wb.sheetnames:
        del wb['Sheet']

    # 🗂️ أوراق التفاصيل أولاً
    detail = {}   # مفتاح القالب -> (اسم الورقة، صف الإجمالي، عدد الأعمدة)
    key_of = {i: k for k, i in matched.items()}
    for idx, section in enumerate(sections):
        ws_d, total_row, ncols = _write_fin_detail_sheet(wb, section, idx, S)
        if idx in key_of and total_row:
            detail[key_of[idx]] = (ws_d.title, total_row, ncols)
        elif idx in unmatched:
            print(f"⚠️ generate_financial_report: القسم «{section['title']}» لم يُربط بالقالب (ورقة تفاصيل فقط). "
                  f"استخدم key أو عنواناً من القالب لربطه.")

    # 📊 ورقة الملخص (الأولى)
    ref_section = sections[min(matched.values())]
    nv = max([n - 1 for (_, _, n) in detail.values()] + [1])           # عدد أعمدة القيم
    value_headers = list(ref_section.get('headers', ['البيان', 'المبلغ']))[1:nv + 1]
    value_headers += ['المبلغ'] * (nv - len(value_headers))
    last_col = nv + 1

    first_name = _safe_sheet_name(template.get('first_sheet', 'الملخص'))
    ws = wb.create_sheet(first_name, 0)
    ws.sheet_view.rightToLeft = True
    right = Alignment(horizontal="right", vertical="center", wrap_text=True)
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)

    for r, (text, font, h) in enumerate([
        (company_name, S['title_font'], 40),
        (template.get('name_ar', report_type), S['subtitle_font'], 30),
        (f"عن الفترة: {period}", S['period_font'], 25),
    ], 1):
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=last_col)
        c = ws.cell(row=r, column=1, value=text)
        c.font, c.alignment = font, Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[r].height = h
    ws.merge_cells(start_row=4, start_column=1, end_row=4, end_column=last_col)
    ws.cell(row=4, column=1).fill = S['header_fill']
    ws.row_dimensions[4].height = 3

    for i, h in enumerate(['البيان'] + value_headers, 1):
        c = ws.cell(row=5, column=i, value=h)
        c.fill, c.font, c.alignment, c.border = S['header_fill'], S['header_font'], center, S['thick']
    ws.row_dimensions[5].height = 30

    row_of = {}          # مفتاح الصف -> رقم الصف في الملخص
    current_row = 6
    missing_required = []
    for t in template_sections:
        key, title = t['key'], t['title']
        is_formula, is_check = t.get('is_formula', False), t.get('is_check', False)
        lc = ws.cell(row=current_row, column=1, value=title)
        lc.alignment = right

        for j in range(nv):
            L = get_column_letter(j + 2)
            cell = ws.cell(row=current_row, column=j + 2)
            if is_formula or is_check:
                refs = {k: f"{L}{r}" for k, r in row_of.items()}
                f = _build_excel_formula(t['formula'], refs)
                if is_check:
                    f = f'=IF(ROUND({f[1:]},2)=0,"✅ متوازنة","⚠️ فرق: "&TEXT({f[1:]},"#,##0.00"))'
                cell.value = f
            elif key in detail and j + 2 <= detail[key][2]:
                sheet_title, trow, _ = detail[key]
                cell.value = f"={sheet_ref(sheet_title)}!{L}{trow}"
            else:
                cell.value = 0
                if j == 0 and key not in detail and not t.get('optional'):
                    missing_required.append(title)
            cell.alignment = center
            cell.number_format = '#,##0.00'

        styled = is_formula or is_check
        for ci in range(1, last_col + 1):
            c = ws.cell(row=current_row, column=ci)
            c.font = S['formula_row_font'] if styled else S['data_font']
            c.border = S['thick'] if styled else S['thin']
            if styled:
                c.fill = S['formula_row_fill']
        ws.row_dimensions[current_row].height = 28 if styled else 24
        row_of[key] = current_row
        current_row += 1

    if missing_required:
        print(f"⚠️ generate_financial_report: أقسام مطلوبة غير موجودة (عُرضت أصفاراً): {', '.join(missing_required)}")
    if not (include_totals and include_formulas):
        print("ℹ️ generate_financial_report: قوالب الدخل/المركز المالي/التدفقات تحسب دائماً بمعادلات حية؛ "
              "include_totals/include_formulas تخص المسار العام فقط.")

    ws.column_dimensions['A'].width = 50
    for j in range(nv):
        ws.column_dimensions[get_column_letter(j + 2)].width = 25
    ws.freeze_panes = 'A6'

    # 💾 الحفظ
    save_dir = output_dir or WORKSPACE_DIR
    os.makedirs(save_dir, exist_ok=True)
    if filename is None:
        safe_company = company_name.replace(' ', '_').replace('/', '_')[:20]
        safe_type = template.get('name_ar', report_type).replace(' ', '_').replace('/', '_')[:20]
        filename = f"{safe_type}_{safe_company}_{period}.xlsx"
    if not filename.endswith('.xlsx'):
        filename += '.xlsx'
    output_path = os.path.join(save_dir, filename)
    finalize_workbook(wb)
    wb.save(output_path)
    print("SUCCESS: " + output_path)
    return output_path


# ==========================================
# المسار العام للتقارير المالية (صفوف إجمالي بمعادلات + id/formula)
# ==========================================
def _generate_generic_financial_report(
    report_type: str,
    company_name: str,
    period: str,
    sections: list,
    currency: str = "ريال",
    include_totals: bool = True,
    include_formulas: bool = True,
    style: str = "accounting",
    output_dir: str = None,
    filename: str = None,
    **kwargs
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
    
    معادلات حية (include_formulas=True): كل صف قاموس أخير يمكن أن يحمل:
        'id': 'rev'                      # اسم يُشار به من معادلات أخرى
        'formula': '={rev}-{exp}'        # معادلة تُطبَّق على كل الأعمدة الرقمية في الصف ({id} = خلية الصف ذاك في نفس العمود)
        'is_subtotal': True              # يُجمع تلقائياً ما فوقه منذ آخر إجمالي فرعي
        'is_total': True                 # الإجمالي الكلي: مجموع الإجماليات الفرعية في القسم، وإلا مجموع بنود القسم
    الإجمالي الكلي في تقارير غير جمعية (قائمة الدخل مثلاً: الإيرادات − المصروفات) يحتاج 'formula'
    وإلا يبقى رقماً ثابتاً مع تنبيه.

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
    
    # إنشاء الورقة الرئيسية (بدلاً من wb.active الذي يصبح None)
    ws = wb.create_sheet(title=report_type)
    ws.sheet_view.rightToLeft = True
    
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
    row_ids = {}                # id الصف -> رقم الصف
    pending_formulas = []       # معادلات مخصصة تُحَلّ بعد معرفة كل المعرّفات
    fin_warns = []
    ADDITIVE_REPORT_TYPES = {'balance_sheet', 'trial_balance', 'fixed_assets', 'sales_report', 'expenses_report'}
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
        data_first_row = current_row
        since_sub = current_row      # بداية نطاق الإجمالي الفرعي التالي
        sub_rows = []                # صفوف الإجماليات الفرعية في هذا القسم
        for row_data in section['data']:
            flags = row_data[-1] if row_data and isinstance(row_data[-1], dict) else {}
            is_total_row = bool(flags.get('is_total'))
            is_subtotal_row = bool(flags.get('is_subtotal'))
            row_formula = flags.get('formula') if include_formulas else None
            if flags.get('id'):
                row_ids[str(flags['id'])] = current_row

            # استخراج البيانات الفعلية (بدون القاموس إن وجد)
            clean_data = [item for item in row_data if not isinstance(item, dict)]

            for col_idx, cell_data in enumerate(clean_data, 1):
                value = cell_data
                L_ = get_column_letter(col_idx)
                numeric = isinstance(cell_data, (int, float)) and not isinstance(cell_data, bool)
                if include_formulas and col_idx > 1:
                    if row_formula and (numeric or cell_data in (None, '')):
                        pending_formulas.append((current_row, col_idx, row_formula))
                        value = 0
                    elif numeric and include_totals and is_subtotal_row and current_row > since_sub:
                        value = f"=SUM({L_}{since_sub}:{L_}{current_row - 1})"
                    elif numeric and include_totals and is_total_row:
                        if report_type in ADDITIVE_REPORT_TYPES and sub_rows:
                            value = "=SUM(" + ",".join(f"{L_}{r_}" for r_ in sub_rows) + ")"
                        elif report_type in ADDITIVE_REPORT_TYPES and current_row > data_first_row:
                            value = f"=SUM({L_}{data_first_row}:{L_}{current_row - 1})"
                        elif current_row not in [w[0] for w in fin_warns if isinstance(w, tuple)]:
                            fin_warns.append((current_row,
                                f"صف الإجمالي في القسم '{section.get('title', '')}' بقي رقماً ثابتاً "
                                f"(نوع التقرير {report_type} غير جمعي): مرّر 'formula' و'id' لربطه بمعادلة"))

                cell = ws.cell(row=current_row, column=col_idx, value=value)
                cell.font = current_style['data_font']
                cell.alignment = right_align if col_idx > 1 else left_align
                cell.border = thin_border

                # تنسيق الأرقام المحاسبي
                is_num_cell = (isinstance(value, (int, float)) and not isinstance(value, bool)) or \
                              (isinstance(value, str) and value.startswith('='))
                if is_num_cell and col_idx > 1:
                    cell.number_format = '#,##0.00'

                # تنسيق صف الإجمالي الفرعي
                if is_subtotal_row:
                    cell.fill = current_style['subtotal_fill']
                    cell.font = current_style['subtotal_font']

                # تنسيق صف الإجمالي الكلي
                if is_total_row:
                    cell.fill = current_style['total_fill']
                    cell.font = current_style['total_font']

            if is_subtotal_row or is_total_row:
                since_sub = current_row + 1
                if is_subtotal_row:
                    sub_rows.append(current_row)
                else:
                    sub_rows = []

            current_row += 1

        # مسافة بين الأقسام
        current_row += 1
    
    # ==========================================
    # إضافة معادلات Excel التلقائية (إن طُلب)
    # ==========================================
    if include_formulas:
        for (r_, c_, tpl_) in pending_formulas:
            L_ = get_column_letter(c_)

            def _sub(m_, L_=L_):
                key_ = m_.group(1).strip()
                if key_ in row_ids:
                    return f"{L_}{row_ids[key_]}"
                fin_warns.append((0, f"المعرّف '{key_}' غير موجود في صفوف التقرير فاستُبدل بصفر"))
                return "0"
            f_ = re.sub(r"\{([^{}]+)\}", _sub, tpl_)
            ws.cell(row=r_, column=c_).value = f_ if f_.startswith('=') else '=' + f_
        for _w in fin_warns:
            print("⚠️ generate_financial_report: " + _w[1])
    
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
    finalize_workbook(wb)
    wb.save(output_path)
    
    # ✅ طبقة الحماية النهائية: ضمان طباعة مسار الملف بنجاح
    print("SUCCESS: " + output_path)
    
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
    filename=None,
    closing_text=None,
    note=None,
    **kwargs
):
    """
    إنشاء خطاب باستخدام نظام الأنماط الديناميكي.
    closing_text: عبارة ختام تُكتب قبل التوقيع (اختياري). note: ملاحظة بعد التوقيع (اختياري).
    يُنظَّف ترقيم بداية الفقرات (1- أو أ) تلقائياً.
    """
    from core.config import WORKSPACE_DIR

    if isinstance(paragraphs, str):
        paragraphs = [paragraphs]
    paragraphs = [re.sub(r'^[\d\u0660-\u0669]+[\.\-\)]\s*', '', str(p).strip()) for p in (paragraphs or [])]

    doc = Document()
    setup_document(doc)
    
    apply_header_style(doc, header_style, organization_name, title)
    add_document_info(doc, recipient=recipient, subject=subject)
    add_greeting(doc)
    apply_content_style(doc, content_style, paragraphs)
    
    if table_headers and table_data:
        apply_table_style(doc, table_style, table_headers, table_data, table_caption)

    if closing_text:
        _cp = doc.add_paragraph()
        _cp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        _cp.paragraph_format.space_before = Pt(10)
        _run = _cp.add_run(closing_text)
        _run.font.name = FONTS['body']
        _run.font.size = Pt(SIZES['body'])
        set_paragraph_rtl(_cp)

    apply_signature_style(doc, signature_style, signatures)
    if note:
        add_note(doc, note)
    apply_footer_style(doc, footer_style)
    
    save_dir = output_dir or WORKSPACE_DIR
    os.makedirs(save_dir, exist_ok=True)
    
    if filename is None:
        safe_title = title[:30].replace(' ', '_').replace('/', '-')
        filename = f"خطاب_{safe_title}"
    
    filename = get_unique_filename(save_dir, filename, 'docx')
    output_path = os.path.join(save_dir, filename)
    
    doc.save(output_path)

    # ✅ طبقة الحماية النهائية: ضمان طباعة مسار الملف بنجاح
    print("SUCCESS: " + output_path)
    
    return output_path


# ==========================================
# 📊 دالة إنشاء تقرير مالي احترافي (مثل Fixed Asset Listing)
# ==========================================

# قاموس نسب الإهلاك الافتراضية (طريقة الرصيد المتناقص)
DEFAULT_DEPRECIATION_RATES = {
    'المركبات': 33.33,
    'الأجهزة': 25.00,
    'الأجهزة والمعدات': 25.00,
    'المعدات': 25.00,
    'الأثاث': 20.00,
    'الأثاث والمفروشات': 20.00,
    'المفروشات': 20.00,
    'المباني': 5.00,
    'البرمجيات': 33.33,
}

def _find_header(headers, include, exclude=()):
    """يرجع فهرس أول عنوان يحتوي إحدى الكلمات (ولا يحتوي المستبعدة)، أو None."""
    for i, h in enumerate(headers):
        hs = str(h)
        if any(k in hs for k in include) and not any(k in hs for k in exclude):
            return i
    return None


def create_fixed_asset_report(
    company_name: str,
    year: str,
    sections: list,
    output_dir: str = None,
    filename: str = None,
    assets: list = None,  # ✅ معلمة جديدة: بيانات الأصول التفصيلية
    include_depreciation_schedule: bool = True,  # ✅ معلمة جديدة: هل نضيف ورقة الإهلاك؟
    add_dep_columns: list = None,       # أعمدة إضافية لورقة الإهلاك
    remove_dep_columns: list = None,
    **kwargs
) -> str:
    """
    إنشاء تقرير أصول ثابتة / تقرير مالي متعدد الأوراق (ملخص + أوراق تفصيلية + جدول إهلاك)

    sections: قائمة الأقسام، كل قسم:
    {
        'sheet_name': 'اسم_الورقة',
        'headers': ['رقم الأصل', 'الوصف', 'تاريخ الشراء', 'التكلفة', 'الإهلاك', 'الصافي'],
        'data': [['A001', 'سيارة', '2023-01-01', 150000, 30000, 120000], ...],
        'totals': True | False | ['التكلفة', 'الإهلاك'],   # اختياري: صف الإجمالي (الافتراضي: كل الأعمدة الرقمية)
        'formulas': {'الصافي': '={التكلفة}-{الإهلاك}'}       # اختياري: معادلات بعناوين الأعمدة
    }

    ما يُحسب تلقائياً كمعادلات حية:
      - عمود الصافي = التكلفة − الإهلاك (يُكتشف من العناوين إن لم تُحدَّد معادلة)
      - صف الإجمالي لكل ورقة، وكل ورقة الملخص (مراجع إلى الأوراق التفصيلية)
      - ورقة الإهلاك الشهري (التكلفة ومجمع الإهلاك مربوطان بالأوراق التفصيلية)
    """
    from core.config import WORKSPACE_DIR
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

    warn_ignored("create_fixed_asset_report", kwargs)

    wb = openpyxl.Workbook()
    if 'Sheet' in wb.sheetnames:
        del wb['Sheet']

    # ألوان وتنسيقات احترافية
    header_fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
    header_font = Font(name="Arial", size=12, bold=True, color="FFFFFF")
    data_font = Font(name="Arial", size=10)
    center_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    right_align = Alignment(horizontal="right", vertical="center", wrap_text=True)
    thin_border = Border(
        left=Side(style="thin", color="B4C7E7"), right=Side(style="thin", color="B4C7E7"),
        top=Side(style="thin", color="B4C7E7"), bottom=Side(style="thin", color="B4C7E7")
    )

    # 1️⃣ ورقة الملخص (تُملأ بعد الأوراق التفصيلية لأنها تشير إليها)
    summary_ws = wb.create_sheet(title="ملخص التقرير")
    summary_ws.sheet_view.rightToLeft = True

    # 2️⃣ الأوراق التفصيلية
    meta = []   # معلومات كل قسم لاستخدامها في الملخص وورقة الإهلاك
    for sec in sections:
        sheet_name = str(sec['sheet_name'])[:31]
        headers = list(sec['headers'])
        data = sec.get('data', [])
        ws = wb.create_sheet(title=sheet_name)
        ws.sheet_view.rightToLeft = True

        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(headers))
        title_cell = ws.cell(row=1, column=1)
        title_cell.value = f"تفاصيل: {sec['sheet_name']}"
        title_cell.font = Font(name="Arial", size=14, bold=True, color="1F4E79")
        title_cell.alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[1].height = 30

        ws.append(headers)
        header_row = 2
        for col_num, header in enumerate(headers, 1):
            cell = ws.cell(row=header_row, column=col_num)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = center_align
            cell.border = thin_border
        ws.row_dimensions[header_row].height = 25

        for row_data in data:
            ws.append(row_data)
            current_row = ws.max_row
            for col_num, _ in enumerate(row_data, 1):
                cell = ws.cell(row=current_row, column=col_num)
                cell.font = data_font
                cell.alignment = center_align if col_num == 1 else right_align
                cell.border = thin_border
                if col_num >= 3:
                    cell.number_format = '#,##0.00'

        first, last = 3, 2 + len(data)

        # أعمدة التكلفة والإهلاك والصافي من العناوين (مع الرجوع للمواضع القديمة 4 و5)
        cost_i = _find_header(headers, ("التكلفة", "تكلفة", "cost"), ("نسبة",))
        dep_i = _find_header(headers, ("الإهلاك", "الاهلاك", "إهلاك", "مجمع", "dep"), ("نسبة", "معدل"))
        net_i = _find_header(headers, ("الصافي", "صافي", "net"))
        if cost_i is None and len(headers) > 3:
            cost_i = 3
        if dep_i is None and len(headers) > 4:
            dep_i = 4

        # معادلات الأعمدة (المحددة صراحة، وإلا الصافي = التكلفة − الإهلاك)
        formulas = dict(sec.get('formulas') or {})
        if not formulas and net_i is not None and cost_i is not None and dep_i is not None \
                and len({net_i, cost_i, dep_i}) == 3 and data \
                and not any(isinstance(r[net_i], str) and str(r[net_i]).startswith('=')
                            for r in data if len(r) > net_i):
            formulas = {headers[net_i]: "={%s}-{%s}" % (headers[cost_i], headers[dep_i])}
        if formulas and data:
            apply_header_formulas(ws, header_row, first, last, formulas)
            for r in range(first, last + 1):
                for h in formulas:
                    c = headers.index(h) + 1 if h in headers else None
                    if c:
                        cell = ws.cell(row=r, column=c)
                        cell.font, cell.alignment, cell.border = data_font, right_align, thin_border

        # صف الإجمالي (معادلات SUM)
        totals_opt = sec.get('totals', 'auto')
        total_row = None
        if totals_opt is not False and data:
            total_row = add_totals_by_header(ws, header_row, first, last,
                                             "auto" if totals_opt in (True, None) else totals_opt)

        ws.freeze_panes = 'A3'
        ws.auto_filter.ref = f"A2:{get_column_letter(len(headers))}{max(last, 2)}"
        for ci in range(1, len(headers) + 1):
            ws.column_dimensions[get_column_letter(ci)].width = 18

        meta.append({'name': sheet_name, 'first': first, 'last': last, 'total_row': total_row,
                     'cost_i': cost_i, 'dep_i': dep_i, 'data': data, 'count': len(data)})

    # الملخص: كل رقم مرجع إلى الورقة التفصيلية
    summary_cols = [
        col("section", "القسم", "text", width=18),
        col("count", "عدد الأصول", total="sum", width=18),
        col("cost", "إجمالي التكلفة", total="sum", width=18),
        col("dep", "إجمالي الإهلاك", total="sum", width=18),
        col("net", "صافي القيمة", formula="={cost}-{dep}", total="sum", width=18),
    ]
    summary_rows = []
    for m in meta:
        ref = sheet_ref(m['name'])

        def colref(idx):
            if idx is None or m['count'] == 0:
                return 0
            L = get_column_letter(idx + 1)
            if m['total_row']:
                return f"={ref}!{L}{m['total_row']}"
            return f"=SUM({ref}!{L}{m['first']}:{L}{m['last']})"

        summary_rows.append({
            'section': m['name'],
            'count': f"=COUNTA({ref}!A{m['first']}:A{m['last']})" if m['count'] else 0,
            'cost': colref(m['cost_i']),
            'dep': colref(m['dep_i']),
        })

    summary_ws.merge_cells('A1:E1')
    title_cell = summary_ws['A1']
    title_cell.value = f"تقرير الأصول الثابتة - {company_name} ({year})"
    title_cell.font = Font(name="Arial", size=16, bold=True, color="1F4E79")
    title_cell.alignment = Alignment(horizontal="center", vertical="center")
    summary_ws.row_dimensions[1].height = 40
    write_table(summary_ws, summary_cols, summary_rows, header_row=2, total=True,
                total_label="الإجمالي الكلي", header_height=25)
    apply_widths(summary_ws, summary_cols, default=18)

    # 3️⃣ ورقة الإهلاك الشهري (مرتبطة بالأوراق التفصيلية)
    if include_depreciation_schedule:
        extracted_assets = []
        for m, sec in zip(meta, sections):
            ref = sheet_ref(m['name'])
            for k, row in enumerate(m['data']):
                if m['dep_i'] is None or m['cost_i'] is None or len(row) <= max(m['dep_i'], m['cost_i']):
                    continue
                r = m['first'] + k

                def link(idx, r=r, row=row):
                    v = _num_cell(row[idx])
                    return f"={ref}!{get_column_letter(idx + 1)}{r}" if isinstance(v, (int, float)) else v
                extracted_assets.append({
                    'asset_id': row[0] if len(row) > 0 else '',
                    'description': row[1] if len(row) > 1 else '',
                    'category': m['name'],
                    'cost': link(m['cost_i']),
                    'accumulated_dep_start': link(m['dep_i']),
                    'depreciation_rate': DEFAULT_DEPRECIATION_RATES.get(m['name'], 20.00),
                    'salvage_value': 0,
                })
        if extracted_assets:
            _add_depreciation_sheet(wb=wb, company_name=company_name, year=year, assets=extracted_assets,
                                    add_columns=add_dep_columns, remove_columns=remove_dep_columns)

    # 4️⃣ حفظ الملف
    save_dir = output_dir or WORKSPACE_DIR
    os.makedirs(save_dir, exist_ok=True)
    if filename is None:
        safe_company = company_name.replace(' ', '_')[:20]
        filename = f"تقرير_الأصول_{safe_company}_{year}"
    if not filename.endswith('.xlsx'):
        filename += '.xlsx'

    output_path = os.path.join(save_dir, filename)
    finalize_workbook(wb)
    wb.save(output_path)

    # ✅ طبقة الحماية النهائية: ضمان طباعة مسار الملف بنجاح
    print("SUCCESS: " + output_path)
    return output_path


def _num_cell(v):
    """قيمة رقمية من خلية قد تكون نصاً مثل '1,200.5'، وإلا ترجع كما هي."""
    if isinstance(v, str):
        try:
            return float(v.replace(',', '').strip())
        except ValueError:
            return v
    return v


# ==========================================
# 📅 دالة داخلية: إضافة ورقة الإهلاك الشهري
# ==========================================
# ==========================================
# 📅 جدول الإهلاك الشهري - معادلات حية (رصيد متناقص)
# ==========================================
_DEP_MONTH_NAMES = ['يناير', 'فبراير', 'مارس', 'أبريل', 'مايو', 'يونيو',
                    'يوليو', 'أغسطس', 'سبتمبر', 'أكتوبر', 'نوفمبر', 'ديسمبر']


def _depreciation_columns():
    from openpyxl.styles import Font
    cols = [
        col("asset_id", "رقم الأصل", "text", width=12),
        col("description", "الوصف", "text", width=25),
        col("cost", "التكلفة", total="sum", width=13),
        col("acc_dep", "مجمع الإهلاك أول المدة", total="sum", width=14),
        col("salvage", "قيمة الخردة", total="sum", width=12),
        col("beg_bv", "رصيد أول المدة", formula="={cost}-{acc_dep}", total="sum", width=13),
        col("rate", "نسبة الإهلاك %", "points", width=11,
            font=Font(name="Arial", size=10, bold=True, color="C9A961")),
    ]
    for i, name in enumerate(_DEP_MONTH_NAMES, 1):
        cols.append(col(f"m{i}", name, group="months", total="sum", width=11,
                        formula="=MAX(0,MIN({beg_bv}*{rate}/100,{beg_bv}-{salvage}))/12"))
    cols.append(col("annual", "إجمالي السنوي", formula="=SUM({group:months})", total="sum", width=13,
                    font=Font(name="Arial", size=9, bold=True, color="1F4E79")))
    cols.append(col("end_bv", "رصيد آخر المدة", formula="={beg_bv}-{annual}", total="sum", width=13,
                    font=Font(name="Arial", size=9, bold=True, color="1E8449")))
    return cols


def _fill_depreciation_sheet(ws, company_name, year, assets, add_columns=None,
                             remove_columns=None, rename_columns=None):
    """يملأ ورقة الإهلاك: المدخلات قيم، وكل ما عداها (الرصيد، الأشهر، السنوي، الختامي) معادلات."""
    from openpyxl.styles import Font, Alignment

    cols, warns = resolve_columns(_depreciation_columns(), add_columns, remove_columns, rename_columns)
    n = len(cols)
    ws.sheet_view.rightToLeft = True

    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=n)
    title_cell = ws['A1']
    title_cell.value = f"جدول الإهلاك الشهري (طريقة الرصيد المتناقص) - {company_name} ({year})"
    title_cell.font = Font(name="Arial", size=14, bold=True, color="1F4E79")
    title_cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[1].height = 30

    rows = []
    for a in assets:
        row = dict(a)
        category = a.get('category', '')
        rate = a.get('depreciation_rate')
        if rate is None:
            rate = DEFAULT_DEPRECIATION_RATES.get(category, 20.00)
        row.update({
            'asset_id': a.get('asset_id', ''),
            'description': a.get('description', ''),
            'cost': a.get('cost', 0),
            'acc_dep': a.get('accumulated_dep_start', 0),
            'salvage': a.get('salvage_value', 0),
            'rate': rate,
        })
        rows.append(row)

    info = write_table(ws, cols, rows, header_row=2, total=True, total_label="الإجمالي",
                       header_height=40, row_height=18)
    apply_widths(ws, cols, default=11)
    ws.freeze_panes = 'C3'
    print_warnings("جدول الإهلاك", warns + info['warnings'])
    return info


def _add_depreciation_sheet(
    wb,
    company_name: str,
    year: str,
    assets: list,
    header_fill=None,
    header_font=None,
    data_font=None,
    thin_border=None,
    center_align=None,
    add_columns=None,
    remove_columns=None,
    rename_columns=None
):
    """
    إضافة ورقة الإهلاك الشهري إلى ملف الأصول الثابتة (معاملات الأنماط القديمة تُقبل وتُهمل).
    """
    ws = wb.create_sheet(title="جدول_الإهلاك_الشهري")
    return _fill_depreciation_sheet(ws, company_name, year, assets, add_columns, remove_columns, rename_columns)

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
    filename=None,
    **kwargs
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

    # ✅ طبقة الحماية النهائية: ضمان طباعة مسار الملف بنجاح
    print("SUCCESS: " + output_path)
    
    return output_path

# ==========================================
# 🔄 دالة محاكاة الملفات (Replication)
# تحلل ملف مرجعي وتولد كود Python لإعادة بنائه
# ==========================================
def replicate_document(
    source_file: str,
    output_dir: str = None,
    filename: str = None,
    **kwargs
) -> dict:
    """
    تحليل ملف مرجعي (Word/Excel) وتوليد كود Python لإعادة بنائه
    
    تُستخدم عندما يرفع المستخدم ملفاً ويطلب إنشاء نموذج مماثل.
    
    ترجع:
    {
        'success': True/False,
        'file_type': 'excel'/'word',
        'analysis': {...},  # تحليل الهيكل والتنسيق
        'generated_code': '...',  # كود Python قابل للتنفيذ
        'output_path': '...'  # مسار الملف المُنشأ
    }
    """
    from core.config import WORKSPACE_DIR
    import os
    
    save_dir = output_dir or WORKSPACE_DIR
    os.makedirs(save_dir, exist_ok=True)
    
    file_ext = source_file.lower().split('.')[-1]
    
    if file_ext == 'xlsx':
        return _replicate_excel(source_file, save_dir, filename)
    elif file_ext == 'docx':
        return _replicate_word(source_file, save_dir, filename)
    else:
        return {
            'success': False,
            'error': f'نوع الملف غير مدعوم: {file_ext}. المدعوم: xlsx, docx'
        }


def _replicate_excel(source_file: str, save_dir: str, filename: str = None) -> dict:
    """محاكاة ملف Excel"""
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    
    wb = openpyxl.load_workbook(source_file)
    
    analysis = {
        'type': 'excel',
        'sheets': {},
        'total_sheets': len(wb.sheetnames)
    }
    
    generated_code_lines = [
        'import openpyxl',
        'from openpyxl.styles import Font, PatternFill, Alignment, Border, Side',
        'from openpyxl.utils import get_column_letter',
        '',
        'wb = openpyxl.Workbook()',
        "if 'Sheet' in wb.sheetnames: del wb['Sheet']",
        ''
    ]
    
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        sheet_info = {
            'title': sheet_name,
            'dimensions': ws.dimensions,
            'merged_cells': [str(m) for m in ws.merged_cells.ranges],
            'rows': ws.max_row,
            'cols': ws.max_column,
            'right_to_left': ws.sheet_view.rightToLeft if ws.sheet_view else False
        }
        
        # استخراج البيانات والتنسيقات (أول 50 صف و20 عمود لتجنب الحجم الكبير)
        rows_data = []
        for row_idx in range(1, min(ws.max_row + 1, 50)):
            row_data = []
            for col_idx in range(1, min(ws.max_column + 1, 20)):
                cell = ws.cell(row=row_idx, column=col_idx)
                cell_info = {
                    'value': cell.value,
                    'font': {
                        'name': cell.font.name,
                        'size': cell.font.size,
                        'bold': cell.font.bold,
                        'italic': cell.font.italic,
                        'color': str(cell.font.color.rgb) if cell.font.color and cell.font.color.rgb else None
                    },
                    'fill': str(cell.fill.start_color.rgb) if cell.fill and cell.fill.start_color and cell.fill.start_color.rgb else None,
                    'alignment': str(cell.alignment.horizontal) if cell.alignment else None,
                    'number_format': cell.number_format,
                    'is_formula': str(cell.value).startswith('=') if cell.value else False
                }
                row_data.append(cell_info)
            rows_data.append(row_data)
        
        sheet_info['data'] = rows_data
        analysis['sheets'][sheet_name] = sheet_info
        
        # توليد كود Python لهذه الورقة
        safe_name = sheet_name.replace("'", "\\'")
        generated_code_lines.append(f"# ورقة: {sheet_name}")
        generated_code_lines.append(f"ws = wb.create_sheet(title='{safe_name}')")
        generated_code_lines.append(f"ws.sheet_view.rightToLeft = True")
        generated_code_lines.append('')
        
        # الخلايا المدمجة
        for merge_range in sheet_info['merged_cells']:
            generated_code_lines.append(f"ws.merge_cells('{merge_range}')")
        
        # البيانات
        for row_idx, row_data in enumerate(rows_data, 1):
            for col_idx, cell_info in enumerate(row_data, 1):
                col_letter = get_column_letter(col_idx)
                value = cell_info['value']
                
                if value is None:
                    continue
                
                # تحويل القيمة إلى تمثيل Python
                value_repr = repr(value)   # تمثيل آمن للنصوص (يتحمل الفواصل العليا والأسطر)
                
                generated_code_lines.append(f"cell = ws.cell(row={row_idx}, column={col_idx}, value={value_repr})")
                
                # التنسيق
                if cell_info['font']['bold']:
                    generated_code_lines.append(f"cell.font = Font(bold=True)")
                if cell_info['fill']:
                    generated_code_lines.append(f"cell.fill = PatternFill(start_color='{cell_info['fill']}', fill_type='solid')")
                if cell_info['number_format'] and cell_info['number_format'] != 'General':
                    generated_code_lines.append(f"cell.number_format = '{cell_info['number_format']}'")
                
                generated_code_lines.append('')
        
        # ضبط الأعمدة
        for col in range(1, min(sheet_info['cols'] + 1, 20)):
            generated_code_lines.append(f"ws.column_dimensions['{get_column_letter(col)}'].width = 18")
        
        generated_code_lines.append('')
    
    # كود الحفظ
    if filename is None:
        filename = 'replicated_document.xlsx'
    if not filename.endswith('.xlsx'):
        filename += '.xlsx'
    
    output_path = os.path.join(save_dir, filename)
    safe_path = output_path.replace('\\', '\\\\')
    generated_code_lines.append(f"output_path = r'{safe_path}'")
    generated_code_lines.append("wb.save(output_path)")
    generated_code_lines.append("print(f'SUCCESS: {output_path}')")
    
    generated_code = '\n'.join(generated_code_lines)
    
    return {
        'success': True,
        'file_type': 'excel',
        'analysis': analysis,
        'generated_code': generated_code,
        'output_path': output_path
    }


def _replicate_word(source_file: str, save_dir: str, filename: str = None) -> dict:
    """محاكاة ملف Word"""
    from docx import Document
    from docx.shared import Pt, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    
    doc = Document(source_file)
    
    analysis = {
        'type': 'word',
        'paragraphs': [],
        'tables': [],
        'styles': set()
    }
    
    generated_code_lines = [
        'from docx import Document',
        'from docx.shared import Pt, RGBColor',
        'from docx.enum.text import WD_ALIGN_PARAGRAPH',
        '',
        'doc = Document()',
        ''
    ]
    
    # تحليل الفقرات
    for para in doc.paragraphs:
        para_info = {
            'text': para.text,
            'style': para.style.name if para.style else None,
            'alignment': str(para.alignment) if para.alignment else None,
            'runs': []
        }
        
        for run in para.runs:
            run_info = {
                'text': run.text,
                'bold': run.font.bold,
                'italic': run.font.italic,
                'size': str(run.font.size) if run.font.size else None,
                'color': str(run.font.color.rgb) if run.font.color and run.font.color.rgb else None
            }
            para_info['runs'].append(run_info)
            if para.style and para.style.name:
                analysis['styles'].add(para.style.name)
        
        analysis['paragraphs'].append(para_info)
        
        # توليد كود
        if para.text.strip():
            generated_code_lines.append(f"p = doc.add_paragraph({para.text!r})")
            if para.alignment:
                generated_code_lines.append(f"p.alignment = WD_ALIGN_PARAGRAPH.{para.alignment.name}")
            for run_info in para_info['runs']:
                if run_info['bold']:
                    generated_code_lines.append(f"for run in p.runs: run.bold = True")
            generated_code_lines.append('')
    
    # تحليل الجداول
    for table_idx, table in enumerate(doc.tables):
        table_data = []
        for row in table.rows:
            row_data = [cell.text for cell in row.cells]
            table_data.append(row_data)
        analysis['tables'].append(table_data)
        
        generated_code_lines.append(f"# جدول {table_idx + 1}")
        generated_code_lines.append(f"table = doc.add_table(rows={len(table.rows)}, cols={len(table.columns)})")
        generated_code_lines.append(f"table.style = 'Table Grid'")
        for row_idx, row_data in enumerate(table_data):
            for col_idx, cell_text in enumerate(row_data):
                generated_code_lines.append(f"table.cell({row_idx}, {col_idx}).text = {cell_text!r}")
        generated_code_lines.append('')
    
    # كود الحفظ
    if filename is None:
        filename = 'replicated_document.docx'
    if not filename.endswith('.docx'):
        filename += '.docx'
    
    output_path = os.path.join(save_dir, filename)
    safe_path = output_path.replace('\\', '\\\\')
    generated_code_lines.append(f"output_path = r'{safe_path}'")
    generated_code_lines.append("doc.save(output_path)")
    generated_code_lines.append("print(f'SUCCESS: {output_path}')")
    
    generated_code = '\n'.join(generated_code_lines)
    
    # تحويل set إلى list للـ JSON
    analysis['styles'] = list(analysis['styles'])
    
    return {
        'success': True,
        'file_type': 'word',
        'analysis': analysis,
        'generated_code': generated_code,
        'output_path': output_path
    }
    # ==========================================
# 🎨 دالة التنسيق المرن لتحليلات النموذج
# تأخذ تحليل النموذج وتنسقه باحترافية
# ==========================================
def format_analysis_report(
    analysis_data: dict,
    report_title: str,
    company_name: str = None,
    period: str = None,
    output_format: str = 'excel',
    style: str = 'accounting',
    output_dir: str = None,
    filename: str = None,
    **kwargs
) -> str:
    """
    تنسيق تحليل النموذج في ملف احترافي
    
    analysis_data: قاموس يحتوي على تحليل النموذج، مثل:
    {
        'summary': 'نص الملخص التنفيذي',
        'sections': [
            {
                'title': 'عنوان القسم',
                'metrics': [
                    {'label': 'إجمالي الإيرادات', 'value': 500000, 'change': '+11%'},
                    ...
                ],
                'insights': ['ملاحظة 1', 'ملاحظة 2'],
                'recommendations': ['توصية 1']
            }
        ],
        'conclusion': 'الخلاصة النهائية',
        'charts_data': [...]  # اختياري
    }
    
    output_format: 'excel' أو 'word'
    """
    from core.config import WORKSPACE_DIR
    import os
    
    if output_format == 'excel':
        return _format_as_excel(analysis_data, report_title, company_name, period, style, output_dir, filename)
    elif output_format == 'word':
        return _format_as_word(analysis_data, report_title, company_name, period, style, output_dir, filename)
    else:
        raise ValueError(f"تنسيق غير مدعوم: {output_format}. استخدم 'excel' أو 'word'")


def _format_as_excel(analysis_data, report_title, company_name, period, style, output_dir, filename):
    """تنسيق التحليل كملف Excel احترافي"""
    from core.config import WORKSPACE_DIR
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    
    wb = openpyxl.Workbook()
    if 'Sheet' in wb.sheetnames:
        del wb['Sheet']
    
    # الأنماط
    styles_config = {
        'accounting': {
            'header_fill': PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid"),
            'header_font': Font(name="Arial", size=12, bold=True, color="FFFFFF"),
            'title_color': "1F4E79",
            'border_color': "B4C7E7",
            'section_fill': PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid"),
            'positive_color': "1E8449",
            'negative_color': "A93226"
        },
        'modern': {
            'header_fill': PatternFill(start_color="1E8449", end_color="1E8449", fill_type="solid"),
            'header_font': Font(name="Calibri", size=12, bold=True, color="FFFFFF"),
            'title_color': "1E8449",
            'border_color': "A9DFBF",
            'section_fill': PatternFill(start_color="D5F5E3", end_color="D5F5E3", fill_type="solid"),
            'positive_color': "1E8449",
            'negative_color': "A93226"
        }
    }
    
    current_style = styles_config.get(style, styles_config['accounting'])
    thin_border = Border(
        left=Side(style="thin", color=current_style['border_color']),
        right=Side(style="thin", color=current_style['border_color']),
        top=Side(style="thin", color=current_style['border_color']),
        bottom=Side(style="thin", color=current_style['border_color'])
    )
    
    # ==========================================
    # ورقة الملخص التنفيذي
    # ==========================================
    ws_summary = wb.create_sheet(title="الملخص التنفيذي")
    ws_summary.sheet_view.rightToLeft = True
    
    # العنوان
    ws_summary.merge_cells('A1:D1')
    title_cell = ws_summary['A1']
    title_text = report_title
    if company_name:
        title_text += f"\n{company_name}"
    if period:
        title_text += f"\n{period}"
    title_cell.value = title_text
    title_cell.font = Font(name="Arial", size=14, bold=True, color=current_style['title_color'])
    title_cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws_summary.row_dimensions[1].height = 60
    
    # الملخص النصي
    if analysis_data.get('summary'):
        ws_summary.merge_cells('A3:D3')
        summary_cell = ws_summary['A3']
        summary_cell.value = "الملخص التنفيذي"
        summary_cell.font = Font(name="Arial", size=12, bold=True, color=current_style['title_color'])
        
        ws_summary.merge_cells('A4:D6')
        summary_text = ws_summary['A4']
        summary_text.value = analysis_data['summary']
        summary_text.font = Font(name="Arial", size=11)
        summary_text.alignment = Alignment(wrap_text=True, vertical="top")
        ws_summary.row_dimensions[4].height = 80
    
    # الأقسام
    current_row = 8
    for section in analysis_data.get('sections', []):
        # عنوان القسم
        ws_summary.merge_cells(f'A{current_row}:D{current_row}')
        section_title = ws_summary.cell(row=current_row, column=1, value=section.get('title', ''))
        section_title.font = Font(name="Arial", size=12, bold=True, color=current_style['title_color'])
        section_title.fill = current_style['section_fill']
        current_row += 1
        
        # المقاييس
        if section.get('metrics'):
            headers = ['البند', 'القيمة', 'التغيير', 'ملاحظات']
            for col_idx, header in enumerate(headers, 1):
                cell = ws_summary.cell(row=current_row, column=col_idx, value=header)
                cell.fill = current_style['header_fill']
                cell.font = current_style['header_font']
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.border = thin_border
            current_row += 1
            
            for metric in section['metrics']:
                ws_summary.cell(row=current_row, column=1, value=metric.get('label', '')).border = thin_border
                value_cell = ws_summary.cell(row=current_row, column=2, value=metric.get('value', ''))
                value_cell.border = thin_border
                if isinstance(metric.get('value'), (int, float)):
                    value_cell.number_format = '#,##0.00'
                
                change_cell = ws_summary.cell(row=current_row, column=3, value=metric.get('change', ''))
                change_cell.border = thin_border
                # تلوين التغيير
                change_str = str(metric.get('change', ''))
                if '+' in change_str:
                    change_cell.font = Font(color=current_style['positive_color'], bold=True)
                elif '-' in change_str:
                    change_cell.font = Font(color=current_style['negative_color'], bold=True)
                
                ws_summary.cell(row=current_row, column=4, value=metric.get('notes', '')).border = thin_border
                current_row += 1
        
        # الرؤى
        if section.get('insights'):
            current_row += 1
            ws_summary.merge_cells(f'A{current_row}:D{current_row}')
            ws_summary.cell(row=current_row, column=1, value="الرؤى والتحليلات").font = Font(bold=True, italic=True)
            current_row += 1
            for insight in section['insights']:
                ws_summary.merge_cells(f'A{current_row}:D{current_row}')
                ws_summary.cell(row=current_row, column=1, value=f"• {insight}").font = Font(size=10)
                current_row += 1
        
        # التوصيات
        if section.get('recommendations'):
            current_row += 1
            ws_summary.merge_cells(f'A{current_row}:D{current_row}')
            ws_summary.cell(row=current_row, column=1, value="التوصيات").font = Font(bold=True, italic=True, color="C9A961")
            current_row += 1
            for rec in section['recommendations']:
                ws_summary.merge_cells(f'A{current_row}:D{current_row}')
                ws_summary.cell(row=current_row, column=1, value=f"✓ {rec}").font = Font(size=10)
                current_row += 1
        
        current_row += 2
    
    # الخلاصة
    if analysis_data.get('conclusion'):
        ws_summary.merge_cells(f'A{current_row}:D{current_row}')
        ws_summary.cell(row=current_row, column=1, value="الخلاصة النهائية").font = Font(bold=True, size=12, color=current_style['title_color'])
        current_row += 1
        ws_summary.merge_cells(f'A{current_row}:D{current_row+2}')
        conclusion_cell = ws_summary.cell(row=current_row, column=1, value=analysis_data['conclusion'])
        conclusion_cell.font = Font(size=11)
        conclusion_cell.alignment = Alignment(wrap_text=True, vertical="top")
        ws_summary.row_dimensions[current_row].height = 60
    
    # ضبط الأعمدة
    ws_summary.column_dimensions['A'].width = 30
    ws_summary.column_dimensions['B'].width = 18
    ws_summary.column_dimensions['C'].width = 15
    ws_summary.column_dimensions['D'].width = 40
    
    # ==========================================
    # حفظ الملف
    # ==========================================
    save_dir = output_dir or WORKSPACE_DIR
    os.makedirs(save_dir, exist_ok=True)
    
    if filename is None:
        safe_title = report_title.replace(' ', '_')[:30]
        filename = f"تحليل_{safe_title}"
    if not filename.endswith('.xlsx'):
        filename += '.xlsx'
    
    output_path = os.path.join(save_dir, filename)
    wb.save(output_path)

    # ✅ طبقة الحماية النهائية: ضمان طباعة مسار الملف بنجاح
    print("SUCCESS: " + output_path)
    
    return output_path


def _format_as_word(analysis_data, report_title, company_name, period, style, output_dir, filename):
    """تنسيق التحليل كملف Word احترافي"""
    from core.config import WORKSPACE_DIR
    from docx import Document
    from docx.shared import Pt, Cm, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml.ns import qn, nsdecls
    from docx.oxml import parse_xml
    
    doc = Document()
    
    # الهوامش
    section = doc.sections[0]
    section.top_margin = Cm(2)
    section.bottom_margin = Cm(2)
    section.left_margin = Cm(2)
    section.right_margin = Cm(2)
    
    # العنوان الرئيسي
    title_para = doc.add_paragraph()
    title_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title_run = title_para.add_run(report_title)
    title_run.font.size = Pt(20)
    title_run.font.bold = True
    title_run.font.color.rgb = RGBColor.from_string('1F4E79')
    
    if company_name:
        org_para = doc.add_paragraph()
        org_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        org_run = org_para.add_run(company_name)
        org_run.font.size = Pt(14)
        org_run.font.color.rgb = RGBColor.from_string('2E5C8A')
    
    if period:
        period_para = doc.add_paragraph()
        period_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        period_run = period_para.add_run(f"الفترة: {period}")
        period_run.font.size = Pt(12)
        period_run.font.italic = True
    
    # خط فاصل
    sep_para = doc.add_paragraph()
    sep_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    pPr = sep_para._p.get_or_add_pPr()
    pBdr = parse_xml(
        f'<w:pBdr {nsdecls("w")}>'
        f'  <w:bottom w:val="single" w:sz="8" w:space="1" w:color="C9A961"/>'
        f'</w:pBdr>'
    )
    pPr.append(pBdr)
    
    # الملخص
    if analysis_data.get('summary'):
        doc.add_heading('الملخص التنفيذي', level=1)
        summary_para = doc.add_paragraph()
        summary_run = summary_para.add_run(analysis_data['summary'])
        summary_run.font.size = Pt(12)
    
    # الأقسام
    for section_data in analysis_data.get('sections', []):
        doc.add_heading(section_data.get('title', ''), level=1)
        
        # المقاييس كجدول
        if section_data.get('metrics'):
            table = doc.add_table(rows=1, cols=3)
            table.style = 'Table Grid'
            headers = ['البند', 'القيمة', 'التغيير']
            for i, header in enumerate(headers):
                cell = table.rows[0].cells[i]
                cell.text = header
                for paragraph in cell.paragraphs:
                    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    for run in paragraph.runs:
                        run.bold = True
            
            for metric in section_data['metrics']:
                row_cells = table.add_row().cells
                row_cells[0].text = metric.get('label', '')
                row_cells[1].text = str(metric.get('value', ''))
                row_cells[2].text = str(metric.get('change', ''))
        
        # الرؤى
        if section_data.get('insights'):
            doc.add_heading('الرؤى والتحليلات', level=2)
            for insight in section_data['insights']:
                p = doc.add_paragraph()
                p.add_run(f"• {insight}")
        
        # التوصيات
        if section_data.get('recommendations'):
            doc.add_heading('التوصيات', level=2)
            for rec in section_data['recommendations']:
                p = doc.add_paragraph()
                run = p.add_run(f"✓ {rec}")
                run.font.color.rgb = RGBColor.from_string('1E8449')
    
    # الخلاصة
    if analysis_data.get('conclusion'):
        doc.add_heading('الخلاصة النهائية', level=1)
        conclusion_para = doc.add_paragraph()
        conclusion_run = conclusion_para.add_run(analysis_data['conclusion'])
        conclusion_run.font.size = Pt(12)
        conclusion_run.font.bold = True
    
    # الحفظ
    save_dir = output_dir or WORKSPACE_DIR
    os.makedirs(save_dir, exist_ok=True)
    
    if filename is None:
        safe_title = report_title.replace(' ', '_')[:30]
        filename = f"تحليل_{safe_title}"
    if not filename.endswith('.docx'):
        filename += '.docx'
    
    output_path = os.path.join(save_dir, filename)
    doc.save(output_path)

    # ✅ طبقة الحماية النهائية: ضمان طباعة مسار الملف بنجاح
    print("SUCCESS: " + output_path)
    
    return output_path

# ==========================================
# 📅 جدول الإهلاك الشهري المضغوط (طريقة الرصيد المتناقص)
# ==========================================
def generate_depreciation_schedule(
    company_name: str,
    year: str,
    assets: list,
    output_dir: str = None,
    filename: str = None,
    add_columns: list = None,
    remove_columns: list = None,
    rename_columns: dict = None,
    **kwargs
) -> str:
    """
    إنشاء جدول إهلاك شهري مضغوط واحترافي (كل الحسابات معادلات حية)

    assets: قائمة الأصول، كل أصل:
    {'asset_id': 'V-001', 'description': 'تويوتا هايلكس 2024', 'cost': 185000,
     'accumulated_dep_start': 37000, 'depreciation_rate': 33.33, 'salvage_value': 0}

    مرونة الأعمدة: add_columns / remove_columns / rename_columns كما في ميزان المراجعة
    (مثلاً add_columns=[{'key': 'location', 'title': 'الموقع', 'kind': 'text', 'after': 'description'}]).
    تغيير التكلفة أو النسبة أو الخردة في الملف يعيد حساب كل الأشهر تلقائياً.
    """
    from core.config import WORKSPACE_DIR
    import openpyxl

    warn_ignored("generate_depreciation_schedule", kwargs)

    wb = openpyxl.Workbook()
    if 'Sheet' in wb.sheetnames:
        del wb['Sheet']
    ws = wb.create_sheet(title="جدول_الإهلاك_الشهري")
    _fill_depreciation_sheet(ws, company_name, year, assets, add_columns, remove_columns, rename_columns)

    save_dir = output_dir or WORKSPACE_DIR
    os.makedirs(save_dir, exist_ok=True)
    if filename is None:
        safe_company = company_name.replace(' ', '_')[:20]
        filename = f"جدول_الاهلاك_الشهري_{safe_company}_{year}"
    if not filename.endswith('.xlsx'):
        filename += '.xlsx'

    output_path = os.path.join(save_dir, filename)
    finalize_workbook(wb)
    wb.save(output_path)

    # ✅ طبقة الحماية النهائية
    print("SUCCESS: " + output_path)
    return output_path

# ==========================================
# ⚖️ ميزان المراجعة (Trial Balance)
# أرصدة افتتاحية + حركة + أرصدة ختامية
# ==========================================

def _parse_accounts_from_string(text: str) -> list:
    """
    دالة مساعدة: تحويل نص الحسابات إلى قائمة قواميس
    تتعامل مع الصيغ المختلفة التي قد يولدها النموذج
    """
    import re
    accounts = []
    
    # تنظيف النص
    text = text.strip()
    
    # محاولة 1: البحث عن قواميس JSON
    if '{' in text and '}' in text:
        try:
            import json
            # استخراج كل قاموس من النص
            pattern = r'\{[^}]+\}'
            matches = re.findall(pattern, text)
            for match in matches:
                try:
                    # تحويل النص إلى قاموس
                    match = match.replace("'", '"')
                    acc = json.loads(match)
                    accounts.append(acc)
                except:
                    continue
        except:
            pass
    
    # محاولة 2: البحث عن أنماط نصية عربية
    if not accounts:
        lines = text.split('\n')
        for line in lines:
            line = line.strip()
            if not line or line.startswith('#'):
                continue
            
            # البحث عن: رقم الحساب + اسم الحساب + القيم
            match = re.match(r'(\d{4})\s+([^\s:]+)', line)
            if match:
                acc_num = match.group(1)
                acc_name = match.group(2)
                
                # استخراج القيم
                opening_debit = 0
                opening_credit = 0
                debit_movement = 0
                credit_movement = 0
                
                if 'افتتاحي مدين' in line:
                    m = re.search(r'افتتاحي مدين\s+([\d,]+)', line)
                    if m:
                        opening_debit = float(m.group(1).replace(',', ''))
                
                if 'افتتاحي دائن' in line:
                    m = re.search(r'افتتاحي دائن\s+([\d,]+)', line)
                    if m:
                        opening_credit = float(m.group(1).replace(',', ''))
                
                if 'حركة مدينة' in line:
                    m = re.search(r'حركة مدينة\s+([\d,]+)', line)
                    if m:
                        debit_movement = float(m.group(1).replace(',', ''))
                
                if 'حركة دائنة' in line:
                    m = re.search(r'حركة دائنة\s+([\d,]+)', line)
                    if m:
                        credit_movement = float(m.group(1).replace(',', ''))
                
                accounts.append({
                    'account_number': acc_num,
                    'account_name': acc_name,
                    'opening_debit': opening_debit,
                    'opening_credit': opening_credit,
                    'debit_movement': debit_movement,
                    'credit_movement': credit_movement
                })
    
    return accounts


# ==========================================
# ⚖️ ميزان المراجعة - أعمدة بمواصفات + معادلات حية
# ==========================================
TRIAL_BALANCE_COLUMNS = [
    col("account_number", "رقم الحساب", "text", width=14),
    col("account_name", "اسم الحساب", "text", width=30, align="right"),
    col("opening_debit", "رصيد افتتاحي مدين", total="sum", width=18),
    col("opening_credit", "رصيد افتتاحي دائن", total="sum", width=18),
    col("debit_movement", "حركة مدينة", total="sum", width=18),
    col("credit_movement", "حركة دائنة", total="sum", width=18),
    col("closing_debit", "رصيد ختامي مدين", width=18, total="sum",
        formula="=MAX(({opening_debit}-{opening_credit}+{debit_movement}-{credit_movement}),0)"),
    col("closing_credit", "رصيد ختامي دائن", width=18, total="sum",
        formula="=MAX(-({opening_debit}-{opening_credit}+{debit_movement}-{credit_movement}),0)"),
]

_TB_PAIRS = [("opening_debit", "opening_credit"),
             ("debit_movement", "credit_movement"),
             ("closing_debit", "closing_credit")]


def generate_trial_balance(
    company_name: str,
    period: str,
    accounts: list,
    output_dir: str = None,
    filename: str = None,
    columns: list = None,
    add_columns: list = None,
    remove_columns: list = None,
    rename_columns: dict = None,
    **kwargs  # ✅ بروتوكول 1: قبول أي معاملات زائدة (مع تنبيه بما لم يُنفَّذ)
) -> str:
    """
    إنشاء ميزان مراجعة احترافي (كل الأرصدة الختامية والإجماليات وفحص التوازن معادلات حية)

    accounts: قائمة الحسابات، كل حساب:
    {'account_number': '1101', 'account_name': 'الصندوق', 'opening_debit': 50000,
     'opening_credit': 0, 'debit_movement': 20000, 'credit_movement': 15000}

    مرونة الأعمدة:
      add_columns    = [{'key': 'notes', 'title': 'ملاحظات', 'kind': 'text'}]   # القيم من accounts[i]['notes']
      remove_columns = ['opening_debit', 'opening_credit']                      # ميزان بلا أرصدة افتتاحية
      rename_columns = {'account_name': 'البيان'}
      columns        = قائمة أعمدة كاملة تحل محل الافتراضية
    """
    from core.config import WORKSPACE_DIR
    import openpyxl
    from openpyxl.styles import Font, Alignment
    from openpyxl.formatting.rule import FormulaRule

    warn_ignored("generate_trial_balance", kwargs)

    # ✅ بروتوكول 3: معالجة ذكية للأنواع
    if isinstance(accounts, str):
        print("⚠️ تم تحويل accounts من string إلى list")
        accounts = _parse_accounts_from_string(accounts)
    elif isinstance(accounts, dict):
        print("⚠️ تم تحويل accounts من dict إلى list")
        accounts = [accounts]

    # ✅ بروتوكول 4: التحقق من البيانات الفارغة
    if not accounts:
        raise ValueError("❌ قائمة الحسابات فارغة. يرجى تمرير بيانات الحسابات.")
    if not company_name:
        raise ValueError("❌ اسم الشركة مطلوب.")
    if not period:
        raise ValueError("❌ الفترة الزمنية مطلوبة.")

    cols, col_warns = resolve_columns(columns or TRIAL_BALANCE_COLUMNS,
                                      add_columns, remove_columns, rename_columns)
    n = len(cols)

    wb = openpyxl.Workbook()
    if 'Sheet' in wb.sheetnames:
        del wb['Sheet']
    ws = wb.create_sheet(title="ميزان_المراجعة")
    ws.sheet_view.rightToLeft = True

    # العنوان
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=n)
    title_cell = ws['A1']
    title_cell.value = f"ميزان المراجعة - {company_name} ({period})"
    title_cell.font = Font(name="Arial", size=16, bold=True, color="1F4E79")
    title_cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[1].height = 35

    # الجدول (الأرصدة الختامية والإجماليات معادلات)
    info = write_table(ws, cols, accounts, header_row=2, total=True,
                       total_label="الإجمالي", header_height=28)
    apply_widths(ws, cols)

    # فحص التوازن (معادلة تتحدث مع أي تعديل)
    L, tr = info['letters'], info['total_row']
    conds = [f"ROUND({L[a]}{tr}-{L[b]}{tr},2)=0" for a, b in _TB_PAIRS if a in L and b in L]
    if conds:
        check_row = tr + 2
        ws.merge_cells(start_row=check_row, start_column=1, end_row=check_row, end_column=n)
        check_cell = ws.cell(row=check_row, column=1)
        check_cell.value = (f'=IF(AND({",".join(conds)}),"✅ الميزان متوازن",'
                            f'"⚠️ تنبيه: الميزان غير متوازن - يرجى مراجعة البيانات")')
        check_cell.font = Font(name="Arial", size=11, bold=True, color="1E8449")
        ws.conditional_formatting.add(
            f"A{check_row}",
            FormulaRule(formula=[f'ISNUMBER(SEARCH("غير متوازن",$A${check_row}))'],
                        font=Font(name="Arial", size=11, bold=True, color="A93226")))

    ws.freeze_panes = 'C3' if n > 2 else 'A3'
    ws.auto_filter.ref = f"A2:{get_column_letter(n)}{max(info['last_row'], 2)}"
    print_warnings("generate_trial_balance", col_warns + info['warnings'])

    save_dir = output_dir or WORKSPACE_DIR
    os.makedirs(save_dir, exist_ok=True)
    if filename is None:
        safe_company = company_name.replace(' ', '_')[:20]
        filename = f"ميزان_المراجعة_{safe_company}_{period}"
    if not filename.endswith('.xlsx'):
        filename += '.xlsx'
    output_path = os.path.join(save_dir, filename)
    finalize_workbook(wb)
    wb.save(output_path)

    # ✅ بروتوكول 2: طبقة الحماية النهائية
    print("SUCCESS: " + output_path)
    return output_path

# ==========================================
# 💰 كشف الرواتب المتكامل - نسخة نهائية ومجربة
# ==========================================

# ==========================================
# 💰 كشف الرواتب - أعمدة بمواصفات + معادلات حية
# ==========================================
PAYROLL_COLUMNS = [
    col("idx", "م", "index", width=6),
    col("employee_id", "رقم الموظف", "text", width=14),
    col("employee_name", "اسم الموظف", "text", width=28, bold=True, align="left"),
    col("department", "القسم", "text", width=16),
    col("basic_salary", "الراتب الأساسي", group="income", total="sum", width=16),
    col("housing_allowance", "بدل سكن", group="income", total="sum", width=13),
    col("transport_allowance", "بدل مواصلات", group="income", total="sum", width=15),
    col("other_allowances", "بدل أخرى", group="income", total="sum", width=13, slip_label="بدلات أخرى"),
    col("gross_pay", "إجمالي الدخل", formula="=SUM({group:income})", total="sum", width=16, role="gross"),
    col("insurance_deductions", "تأمينات", group="deduction", total="sum", width=13,
        slip_label="خصومات تأمينات"),
    col("other_deductions", "خصومات أخرى", group="deduction", total="sum", width=15),
    col("net_pay", "صافي الراتب", formula="=SUM({group:income})-SUM({group:deduction})",
        total="sum", width=16, role="net"),
]


def _payroll_styles():
    from openpyxl.styles import Font, PatternFill, Border, Side
    thick = Side(style="thick", color="1F4E79")
    thick_border = Border(left=thick, right=thick, top=thick, bottom=thick)
    return {
        "header_fill": PatternFill(start_color="1F4E79", end_color="2E75B6", fill_type="solid"),
        "header_font": Font(name="Arial", size=11, bold=True, color="FFFFFF"),
        "header_border": thick_border,
        "total_fill": PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid"),
        "total_font": Font(name="Arial", size=11, bold=True, color="C00000"),
        "total_border": thick_border,
    }


def _payroll_sheet(wb, sheet_name, heading, company_name, period, cols, emps):
    """ورقة رواتب (رئيسية أو إدارية أو تشغيلية) بنفس الأعمدة والمعادلات."""
    from openpyxl.styles import Font, Alignment
    ws = wb.create_sheet(title=sheet_name)
    ws.sheet_view.rightToLeft = True
    n = len(cols)
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)

    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=n)
    ws['A1'].value = heading
    ws['A1'].font = Font(name="Arial", size=18, bold=True, color="1F4E79")
    ws['A1'].alignment = center
    ws.row_dimensions[1].height = 45

    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=n)
    ws['A2'].value = f"{company_name} - {period}"
    ws['A2'].font = Font(name="Arial", size=12, italic=True, color="404040")
    ws['A2'].alignment = center
    ws.row_dimensions[2].height = 30

    info = write_table(ws, cols, emps, header_row=3, styles=_payroll_styles(), total=True,
                       total_label="الإجمالي العام", label_span=min(3, n),
                       header_height=35, row_height=25)
    apply_widths(ws, cols)
    ws.freeze_panes = 'A4'
    ws.auto_filter.ref = f"A3:{get_column_letter(n)}{max(info['last_row'], 3)}"
    return ws, info


def _create_payroll_signature_slip(ws, row, col, employee, period, items=None):
    """
    إنشاء قسيمة راتب احترافية في مربع واضح.
    items (اختياري): قائمة (التسمية، القيمة أو المعادلة، النوع) والنوع: income | deduction | total
    وإن لم تُمرَّر تُستخدم القيم الثابتة من بيانات الموظف (السلوك القديم).
    """
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

    title_fill = PatternFill(start_color="1F4E79", end_color="2E75B6", fill_type="solid")
    title_font = Font(name="Arial", size=11, bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="D9E2F3", end_color="D9E2F3", fill_type="solid")
    header_font = Font(name="Arial", size=9, bold=True, color="1F4E79")
    label_font = Font(name="Arial", size=8, bold=True, color="404040")
    data_font = Font(name="Arial", size=8, color="000000")
    total_fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
    total_font = Font(name="Arial", size=9, bold=True, color="C00000")

    outer_border = Border(left=Side(style="thick", color="1F4E79"), right=Side(style="thick", color="1F4E79"), top=Side(style="thick", color="1F4E79"), bottom=Side(style="thick", color="1F4E79"))
    inner_border = Border(left=Side(style="thin", color="B4C7E7"), right=Side(style="thin", color="B4C7E7"), top=Side(style="thin", color="B4C7E7"), bottom=Side(style="thin", color="B4C7E7"))

    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    right = Alignment(horizontal="right", vertical="center")
    left = Alignment(horizontal="left", vertical="center")

    start_row = 2 + (row - 1) * 24
    start_col = 1 + (col - 1) * 14

    for r in range(start_row, start_row + 22):
        for c in range(start_col, start_col + 13):
            ws.cell(row=r, column=c).border = inner_border

    ws.merge_cells(start_row=start_row, start_column=start_col, end_row=start_row, end_column=start_col + 12)
    title_cell = ws.cell(row=start_row, column=start_col, value=f"قسيمة راتب - {period}")
    title_cell.font = title_font
    title_cell.fill = title_fill
    title_cell.alignment = center
    title_cell.border = outer_border
    ws.row_dimensions[start_row].height = 25

    emp_data = [
        ("اسم الموظف:", employee.get('employee_name', ''), "رقم الموظف:", employee.get('employee_id', '')),
        ("القسم:", employee.get('department', ''), "تاريخ الصرف:", employee.get('payment_date', ''))
    ]

    for i, (label1, val1, label2, val2) in enumerate(emp_data):
        r = start_row + 2 + i
        ws.row_dimensions[r].height = 18

        ws.cell(row=r, column=start_col, value=label1).font = label_font
        ws.cell(row=r, column=start_col).alignment = right
        ws.cell(row=r, column=start_col).border = inner_border

        ws.merge_cells(start_row=r, start_column=start_col+1, end_row=r, end_column=start_col+4)
        ws.cell(row=r, column=start_col+1, value=val1).font = data_font
        ws.cell(row=r, column=start_col+1).alignment = left
        ws.cell(row=r, column=start_col+1).border = inner_border

        ws.cell(row=r, column=start_col+7, value=label2).font = label_font
        ws.cell(row=r, column=start_col+7).alignment = right
        ws.cell(row=r, column=start_col+7).border = inner_border

        ws.merge_cells(start_row=r, start_column=start_col+8, end_row=r, end_column=start_col+12)
        ws.cell(row=r, column=start_col+8, value=val2).font = data_font
        ws.cell(row=r, column=start_col+8).alignment = left
        ws.cell(row=r, column=start_col+8).border = inner_border

    header_r = start_row + 5
    ws.row_dimensions[header_r].height = 20

    ws.merge_cells(start_row=header_r, start_column=start_col, end_row=header_r, end_column=start_col+6)
    ws.cell(row=header_r, column=start_col, value="البيان").font = header_font
    ws.cell(row=header_r, column=start_col).fill = header_fill
    ws.cell(row=header_r, column=start_col).alignment = center
    ws.cell(row=header_r, column=start_col).border = outer_border

    ws.merge_cells(start_row=header_r, start_column=start_col+7, end_row=header_r, end_column=start_col+12)
    ws.cell(row=header_r, column=start_col+7, value="المبلغ (ر.س)").font = header_font
    ws.cell(row=header_r, column=start_col+7).fill = header_fill
    ws.cell(row=header_r, column=start_col+7).alignment = center
    ws.cell(row=header_r, column=start_col+7).border = outer_border

    if items is None:
        items = [
            ('الراتب الأساسي', employee.get('basic_salary', 0), 'income'),
            ('بدل سكن', employee.get('housing_allowance', 0), 'income'),
            ('بدل مواصلات', employee.get('transport_allowance', 0), 'income'),
            ('بدلات أخرى', employee.get('other_allowances', 0), 'income'),
            ('إجمالي الدخل', employee.get('gross_pay', 0), 'total'),
            ('خصومات تأمينات', employee.get('insurance_deductions', 0), 'deduction'),
            ('خصومات أخرى', employee.get('other_deductions', 0), 'deduction'),
            ('صافي الراتب', employee.get('net_pay', 0), 'total'),
        ]
    if len(items) > 12:   # المساحة المتاحة في القسيمة 12 بنداً
        print(f"⚠️ القسيمة تتسع لـ 12 بنداً فقط، تم تجاهل {len(items) - 12} بنداً")
        items = items[:11] + items[-1:]

    for i, (label, amount, kind) in enumerate(items):
        r = header_r + 1 + i
        ws.row_dimensions[r].height = 18
        is_total = kind == 'total'
        is_deduction = kind == 'deduction'

        ws.merge_cells(start_row=r, start_column=start_col, end_row=r, end_column=start_col+6)
        label_cell = ws.cell(row=r, column=start_col, value=label)
        label_cell.font = total_font if is_total else data_font
        label_cell.alignment = right
        label_cell.border = inner_border
        if is_total:
            label_cell.fill = total_fill
        elif is_deduction:
            label_cell.font = Font(name="Arial", size=8, color="C00000")

        ws.merge_cells(start_row=r, start_column=start_col+7, end_row=r, end_column=start_col+12)
        amount_cell = ws.cell(row=r, column=start_col+7, value=amount)
        amount_cell.font = total_font if is_total else data_font
        amount_cell.alignment = right
        amount_cell.number_format = '#,##0.00'
        amount_cell.border = inner_border
        if is_total:
            amount_cell.fill = total_fill
        elif is_deduction:
            amount_cell.font = Font(name="Arial", size=8, color="C00000")

    sign_row = start_row + 18
    ws.row_dimensions[sign_row].height = 30
    ws.merge_cells(start_row=sign_row, start_column=start_col, end_row=sign_row, end_column=start_col+12)
    sign_cell = ws.cell(row=sign_row, column=start_col, value="توقيع الموظف:")
    sign_cell.font = Font(name="Arial", size=9, bold=True, color="1F4E79")
    sign_cell.alignment = Alignment(horizontal="right", vertical="center")
    sign_cell.border = outer_border

    sign_row2 = sign_row + 1
    ws.merge_cells(start_row=sign_row2, start_column=start_col+3, end_row=sign_row2, end_column=start_col+9)
    ws.cell(row=sign_row2, column=start_col+3).border = Border(bottom=Side(style="medium", color="000000"))
    ws.row_dimensions[sign_row2].height = 25


def _payroll_slip_items(cols, main_sheet, info, emp_index):
    """بنود القسيمة كمراجع حية إلى صف الموظف في الورقة الرئيسية (تتحدث مع أي تعديل)."""
    r = info['first_row'] + emp_index
    L = info['letters']

    def ref(c):
        return f"={sheet_ref(main_sheet)}!{L[c['key']]}{r}"

    items = [(c.get('slip_label', c['title']), ref(c), 'income') for c in cols if c.get('group') == 'income']
    items += [(c['title'], ref(c), 'total') for c in cols if c.get('role') == 'gross']
    items += [(c.get('slip_label', c['title']), ref(c), 'deduction') for c in cols if c.get('group') == 'deduction']
    items += [(c['title'], ref(c), 'total') for c in cols if c.get('role') == 'net']
    return items


def generate_payroll_register(company_name: str, period: str, employees: list, output_dir: str = None,
                              filename: str = None, include_signature_slips: bool = True,
                              columns: list = None, add_columns: list = None,
                              remove_columns: list = None, rename_columns: dict = None,
                              **kwargs) -> str:
    """
    إنشاء كشف رواتب احترافي متكامل (إجمالي الدخل والصافي والإجماليات معادلات حية).

    مرونة الأعمدة:
      add_columns = [{'key': 'overtime', 'title': 'بدل إضافي', 'group': 'income'}]      # يدخل في إجمالي الدخل تلقائياً
      add_columns = [{'key': 'loan', 'title': 'قسط سلفة', 'group': 'deduction'}]        # يُخصم من الصافي تلقائياً
      remove_columns = ['housing_allowance']
      rename_columns = {'department': 'الإدارة'}
    القيم الإضافية تؤخذ من employees[i]['overtime'] وهكذا. القسائم تتسع لـ 12 بنداً.
    """
    from core.config import WORKSPACE_DIR
    import openpyxl
    from openpyxl.styles import Font, Alignment
    from openpyxl.worksheet.pagebreak import Break
    from openpyxl.worksheet.properties import PageSetupProperties

    warn_ignored("generate_payroll_register", kwargs)

    if isinstance(employees, str):
        employees = []
    elif isinstance(employees, dict):
        employees = [employees]
    if not employees:
        raise ValueError("❌ قائمة الموظفين فارغة.")

    cols, col_warns = resolve_columns(columns or PAYROLL_COLUMNS, add_columns, remove_columns, rename_columns)

    wb = openpyxl.Workbook()
    if 'Sheet' in wb.sheetnames:
        del wb['Sheet']

    main_name = "مسير_الرواتب_الرئيسي"
    main_ws, main_info = _payroll_sheet(wb, main_name, "مسير الرواتب الشهرية", company_name, period, cols, employees)

    # تقسيم الموظفين (إدارية / تشغيلية)
    admin_employees, operational_employees = [], []
    for emp in employees:
        dept = str(emp.get('department', '') or '').lower()
        if 'إدار' in dept or 'admin' in dept:
            admin_employees.append(emp)
        else:
            operational_employees.append(emp)

    all_warns = list(col_warns) + main_info['warnings']
    if admin_employees:
        _, i2 = _payroll_sheet(wb, "الرواتب_الإدارية", "الرواتب الإدارية والعمومية",
                               company_name, period, cols, admin_employees)
        all_warns += i2['warnings']
    if operational_employees:
        _, i3 = _payroll_sheet(wb, "الرواتب_التشغيلية", "الرواتب التشغيلية",
                               company_name, period, cols, operational_employees)
        all_warns += i3['warnings']

    if include_signature_slips:
        slip_ws = wb.create_sheet(title="قسائم_التوقيع")
        slip_ws.sheet_view.rightToLeft = True
        slip_ws.page_setup.orientation = slip_ws.ORIENTATION_PORTRAIT
        slip_ws.page_setup.paper_size = slip_ws.PAPERSIZE_A4
        slip_ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
        slip_ws.page_setup.fitToWidth = 1
        slip_ws.page_setup.fitToHeight = 0
        slip_ws.merge_cells('A1:N1')
        slip_ws['A1'].value = f"قسائم استلام الرواتب - {period}"
        slip_ws['A1'].font = Font(name="Arial", size=18, bold=True, color="1F4E79")
        slip_ws['A1'].alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        slip_ws.row_dimensions[1].height = 35
        # قسيمة لكل موظف (4 في الصفحة) بدل أول 4 فقط
        for n, emp in enumerate(employees):
            row_idx, col_idx = n // 2 + 1, n % 2 + 1
            items = _payroll_slip_items(cols, main_name, main_info, n)
            _create_payroll_signature_slip(slip_ws, row_idx, col_idx, emp, period, items=items)
            if col_idx == 2 and row_idx % 2 == 0 and n != len(employees) - 1:
                slip_ws.row_breaks.append(Break(id=1 + row_idx * 24))   # صفحة جديدة كل 4 قسائم
        for c in range(1, 15):
            slip_ws.column_dimensions[get_column_letter(c)].width = 11

    print_warnings("generate_payroll_register", all_warns)

    save_dir = output_dir or WORKSPACE_DIR
    os.makedirs(save_dir, exist_ok=True)
    if filename is None:
        safe_company = company_name.replace(' ', '_')[:20]
        filename = f"كشف_الرواتب_{safe_company}_{period}"
    if not filename.endswith('.xlsx'):
        filename += '.xlsx'

    output_path = os.path.join(save_dir, filename)
    finalize_workbook(wb)
    wb.save(output_path)

    print("SUCCESS: " + output_path)
    return output_path
