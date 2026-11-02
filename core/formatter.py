"""
Officia Formatter - مكتبة التنسيق الاحترافي للمستندات
توفر دوال جاهزة لإنشاء خطابات وتقارير منسقة باحترافية عالية
متوافقة تماماً مع بنية Officia v2.1
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
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, Reference

# استيراد WORKSPACE_DIR كقيمة افتراضية للتوافق مع app.py
try:
    from core.config import WORKSPACE_DIR
except ImportError:
    WORKSPACE_DIR = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'workspace')

# ==========================================
# الثوابت والألوان الرسمية
# ==========================================
COLORS = {
    'primary': '003366',       # أزرق داكن (رسمي)
    'secondary': '006699',     # أزرق متوسط
    'accent': 'C00000',        # أحمر (للتنبيهات)
    'gold': 'B8860B',          # ذهبي (للعناوين الفرعية)
    'light_gray': 'F2F2F2',    # رمادي فاتح
    'white': 'FFFFFF',
    'row_alt': 'D9E2F3',       # أزرق فاتح للصفوف البديلة
}

FONTS = {'default': 'Calibri', 'heading': 'Calibri'}
SIZES = {'title': 22, 'heading': 16, 'subheading': 14, 'body': 12, 'small': 10, 'footer': 9}

# ==========================================
# دوال مساعدة عامة
# ==========================================
def set_cell_shading(cell, color_hex):
    """تعيين لون خلفية خلية في جدول Word"""
    shading_elm = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{color_hex}"/>')
    cell._tc.get_or_add_tcPr().append(shading_elm)

def set_paragraph_rtl(paragraph):
    """تعيين اتجاه الفقرة من اليمين لليسار"""
    pPr = paragraph._p.get_or_add_pPr()
    pPr.set(qn('w:bidi'), '1')

def set_cell_rtl(cell):
    """تعيين اتجاه الخلية من اليمين لليسار"""
    for paragraph in cell.paragraphs:
        set_paragraph_rtl(paragraph)

def add_styled_paragraph(doc, text, bold=False, italic=False, size=None, color=None, 
                         alignment=None, font_name=None, space_after=6, line_spacing=1.15):
    """إضافة فقرة منسقة بشكل احترافي"""
    para = doc.add_paragraph()
    para.alignment = alignment or WD_ALIGN_PARAGRAPH.RIGHT
    para.paragraph_format.space_after = Pt(space_after)
    para.paragraph_format.line_spacing = line_spacing
    
    run = para.add_run(text)
    run.font.name = font_name or FONTS['default']
    run.font.size = Pt(size or SIZES['body'])
    run.font.bold = bold
    run.font.italic = italic
    if color:
        run.font.color.rgb = RGBColor.from_string(color) if isinstance(color, str) else color
    
    set_paragraph_rtl(para)
    return para

def add_horizontal_line(doc, color='003366', thickness=1):
    """إضافة خط فاصل أنيق"""
    para = doc.add_paragraph()
    para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    pPr = para._p.get_or_add_pPr()
    pBdr = parse_xml(f'<w:pBdr {nsdecls("w")}><w:bottom w:val="single" w:sz="{thickness*8}" w:space="1" w:color="{color}"/></w:pBdr>')
    pPr.append(pBdr)

def get_unique_filename(directory, base_name, extension):
    """إنشاء اسم ملف فريد لتجنب التكرار"""
    counter = 1
    filename = f"{base_name}.{extension}"
    while os.path.exists(os.path.join(directory, filename)):
        filename = f"{base_name}_{counter}.{extension}"
        counter += 1
    return filename

# ==========================================
# 1. إنشاء خطاب رسمي احترافي
# ==========================================
def create_professional_letter(
    recipient, subject, content_paragraphs,
    table_data=None, table_headers=None, table_caption=None,
    closing_note=None, closing_phrase=None,
    sender_name=None, sender_title=None, sender_organization=None,
    reference_number=None, attachments=None,
    output_dir=None, filename=None
):
    """إنشاء خطاب رسمي منسق باحترافية عالية"""
    doc = Document()
    section = doc.sections[0]
    section.top_margin = section.bottom_margin = section.left_margin = section.right_margin = Cm(2.5)
    
    style = doc.styles['Normal']
    style.font.name = FONTS['default']
    style.font.size = Pt(SIZES['body'])

    # 1. رأس الخطاب
    if sender_organization:
        org_para = doc.add_paragraph()
        org_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        org_run = org_para.add_run(sender_organization)
        org_run.font.size = Pt(SIZES['heading'])
        org_run.font.bold = True
        org_run.font.color.rgb = RGBColor.from_string(COLORS['primary'])
        set_paragraph_rtl(org_para)

    add_horizontal_line(doc)

    # 2. العنوان الرئيسي
    title = doc.add_heading(subject, level=0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for run in title.runs:
        run.font.color.rgb = RGBColor.from_string(COLORS['primary'])
        run.font.size = Pt(SIZES['title'])
        run.font.bold = True
        run.font.name = FONTS['heading']
    
    add_horizontal_line(doc, color=COLORS['gold'])

    # 3. معلومات الخطاب
    info_table = doc.add_table(rows=1, cols=2)
    info_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    
    date_cell = info_table.cell(0, 0)
    date_para = date_cell.paragraphs[0]
    date_para.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    date_run = date_para.add_run(f"التاريخ: {datetime.now().strftime('%Y/%m/%d')}")
    date_run.font.size = Pt(SIZES['body'])
    date_run.font.bold = True
    set_paragraph_rtl(date_para)
    
    ref_cell = info_table.cell(0, 1)
    ref_para = ref_cell.paragraphs[0]
    ref_para.alignment = WD_ALIGN_PARAGRAPH.LEFT
    ref_run = ref_para.add_run(f"رقم: {reference_number or datetime.now().strftime('%Y%m%d%H%M')}")
    ref_run.font.size = Pt(SIZES['body'])
    ref_run.font.bold = True

    doc.add_paragraph()

    # 4. المرسل إليه والموضوع
    to_para = doc.add_paragraph()
    to_para.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    to_para.add_run("السادة/ ").font.bold = True
    to_para.add_run(recipient).font.bold = True
    set_paragraph_rtl(to_para)
    
    subj_para = doc.add_paragraph()
    subj_para.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    subj_label = subj_para.add_run("الموضوع: ")
    subj_label.font.bold = True
    subj_value = subj_para.add_run(subject)
    subj_value.font.bold = True
    subj_value.font.color.rgb = RGBColor.from_string(COLORS['primary'])
    set_paragraph_rtl(subj_para)
    
    doc.add_paragraph()

    # 5. التحية والمقدمة
    add_styled_paragraph(doc, "تحية طيبة وبعد،،،", bold=True, space_after=12)
    add_styled_paragraph(doc, "بالإشارة إلى الموضوع أعلاه، نتقدم إليكم بهذا الخطاب.", space_after=10)

    # 6. فقرات المحتوى
    for i, para_text in enumerate(content_paragraphs):
        numbered_text = f"{i+1}. {para_text}" if len(content_paragraphs) > 1 else para_text
        add_styled_paragraph(doc, numbered_text, space_after=10, line_spacing=1.5)
    
    doc.add_paragraph()

    # 7. الجدول (إن وجد)
    if table_data and table_headers:
        if table_caption:
            cap_para = doc.add_paragraph()
            cap_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            cap_run = cap_para.add_run(table_caption)
            cap_run.font.bold = True
            cap_run.font.size = Pt(SIZES['subheading'])
            cap_run.font.color.rgb = RGBColor.from_string(COLORS['primary'])
            set_paragraph_rtl(cap_para)
            doc.add_paragraph()
        
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
            run.font.color.rgb = RGBColor.from_string(COLORS['white'])
            run.font.size = Pt(SIZES['body'])
            set_cell_shading(cell, COLORS['primary'])
            set_cell_rtl(cell)
        
        for row_idx, row_data in enumerate(table_data):
            row_cells = table.add_row().cells
            bg_color = COLORS['row_alt'] if row_idx % 2 == 0 else COLORS['white']
            for col_idx, cell_data in enumerate(row_data):
                cell = row_cells[col_idx]
                cell.text = ''
                p = cell.paragraphs[0]
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                run = p.add_run(str(cell_data))
                run.font.size = Pt(SIZES['body'])
                set_cell_shading(cell, bg_color)
                set_cell_rtl(cell)
        doc.add_paragraph()

    # 8. الخاتمة والملاحظات
    closing = closing_phrase or "وتفضلوا بقبول فائق الاحترام والتقدير،،"
    add_styled_paragraph(doc, closing, bold=True, space_after=12)
    
    if closing_note:
        note_para = doc.add_paragraph()
        note_para.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        note_para.add_run("⚠️ تنبيه: ").font.bold = True
        note_para.add_run("⚠️ تنبيه: ").font.color.rgb = RGBColor.from_string(COLORS['accent'])
        note_para.add_run(closing_note).font.color.rgb = RGBColor.from_string(COLORS['accent'])
        set_paragraph_rtl(note_para)

    # 9. التوقيع
    doc.add_paragraph()
    doc.add_paragraph()
    if sender_name or sender_organization:
        sig_para = doc.add_paragraph()
        sig_para.alignment = WD_ALIGN_PARAGRAPH.LEFT
        if sender_name:
            sig_para.add_run(sender_name).font.bold = True
        if sender_title:
            sig_para.add_run('\n' + sender_title).font.size = Pt(SIZES['small'])
            sig_para.add_run('\n' + sender_title).font.color.rgb = RGBColor.from_string(COLORS['secondary'])

    # 10. الحفظ
    save_dir = output_dir or WORKSPACE_DIR
    os.makedirs(save_dir, exist_ok=True)
    
    if filename is None:
        safe_subject = subject[:30].replace(' ', '_').replace('/', '-')
        filename = f"خطاب_{safe_subject}"
    
    filename = get_unique_filename(save_dir, filename, 'docx')
    output_path = os.path.join(save_dir, filename)
    doc.save(output_path)
    return output_path


# ==========================================
# 2. إنشاء تقرير Excel احترافي
# ==========================================
def create_excel_report(
    headers, data, title=None, sheet_name="التقرير",
    add_chart=False, freeze_panes=True, auto_filter=True,
    output_dir=None, filename=None
):
    """إنشاء تقرير Excel منسق باحترافية"""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = sheet_name
    ws.sheet_view.rightToLeft = True
    
    current_row = 1
    if title:
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(headers))
        title_cell = ws.cell(row=1, column=1, value=title)
        title_cell.font = Font(name=FONTS['heading'], size=SIZES['heading'], bold=True, color=COLORS['primary'])
        title_cell.alignment = Alignment(horizontal='center', vertical='center')
        ws.row_dimensions[1].height = 30
        current_row = 3

    header_fill = PatternFill(start_color=COLORS['primary'], end_color=COLORS['primary'], fill_type='solid')
    header_font = Font(name=FONTS['default'], size=SIZES['body'], bold=True, color='FFFFFF')
    header_alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

    for col_num, header in enumerate(headers, 1):
        cell = ws.cell(row=current_row, column=col_num, value=header)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = header_alignment
        cell.border = thin_border
    
    ws.row_dimensions[current_row].height = 25
    data_start_row = current_row + 1

    data_font = Font(name=FONTS['default'], size=SIZES['body'])
    data_alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    alt_fill = PatternFill(start_color=COLORS['row_alt'], end_color=COLORS['row_alt'], fill_type='solid')

    for row_idx, row_data in enumerate(data):
        for col_idx, cell_data in enumerate(row_data, 1):
            cell = ws.cell(row=data_start_row + row_idx, column=col_idx, value=cell_data)
            cell.font = data_font
            cell.alignment = data_alignment
            cell.border = thin_border
            if row_idx % 2 == 0:
                cell.fill = alt_fill
            if isinstance(cell_data, (int, float)):
                cell.number_format = '#,##0.00'

    if freeze_panes:
        ws.freeze_panes = f'A{data_start_row}'
    if auto_filter:
        ws.auto_filter.ref = f'A{current_row}:{get_column_letter(len(headers))}{data_start_row + len(data) - 1}'

    for col_num in range(1, len(headers) + 1):
        max_length = len(str(headers[col_num - 1]))
        for row in ws.iter_rows(min_row=data_start_row, max_row=data_start_row + len(data) - 1, min_col=col_num, max_col=col_num):
            for cell in row:
                if cell.value:
                    max_length = max(max_length, len(str(cell.value)))
        ws.column_dimensions[get_column_letter(col_num)].width = min(max_length + 4, 40)

    if add_chart and len(data) > 1:
        chart = BarChart()
        chart.title = title or 'رسم بياني'
        data_ref = Reference(ws, min_col=2, min_row=current_row, max_row=data_start_row + len(data) - 1)
        cats_ref = Reference(ws, min_col=1, min_row=data_start_row, max_row=data_start_row + len(data) - 1)
        chart.add_data(data_ref, titles_from_data=True)
        chart.set_categories(cats_ref)
        ws.add_chart(chart, f'A{data_start_row + len(data) + 2}')

    save_dir = output_dir or WORKSPACE_DIR
    os.makedirs(save_dir, exist_ok=True)
    
    if filename is None:
        safe_title = (title or 'تقرير')[:30].replace(' ', '_').replace('/', '-')
        filename = f"{safe_title}"
    
    filename = get_unique_filename(save_dir, filename, 'xlsx')
    output_path = os.path.join(save_dir, filename)
    wb.save(output_path)
    return output_path
