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

def setup_document(doc, margins_cm=2.5):
    """إعداد المستند بهوامش موحدة"""
    section = doc.sections[0]
    section.top_margin = Cm(margins_cm)
    section.bottom_margin = Cm(margins_cm)
    section.left_margin = Cm(margins_cm)
    section.right_margin = Cm(margins_cm)
    
    style = doc.styles['Normal']
    style.font.name = FONTS['body']
    style.font.size = Pt(SIZES['body'])
    style.font.color.rgb = RGBColor.from_string(PALETTE['charcoal'])
    style.paragraph_format.line_spacing = 1.5

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
