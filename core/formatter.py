"""
Officia Formatter Pro - مكتبة التنسيق الاحترافي المتقدم
مستوى: مستندات رسمية راقية بخطوط عربية وألوان متناسقة وجماليات عالية
"""

import os
from datetime import datetime
from docx import Document
from docx.shared import Pt, Cm, RGBColor, Emu, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn, nsdecls
from docx.oxml import parse_xml
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, GradientFill, numbers
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, PieChart, Reference
from openpyxl.chart.label import DataLabelList

# ==========================================
# نظام الألوان الاحترافي المتدرج
# ==========================================
PALETTE = {
    # الألوان الأساسية - تدرج أزرق ملكي
    'royal_dark': '1B3A5C',      # أزرق ملكي داكن
    'royal_mid': '2E5C8A',        # أزرق ملكي متوسط
    'royal_light': '4A7FB5',      # أزرق ملكي فاتح
    'royal_pale': 'D6E4F0',       # أزرق باهت للخلفيات
    
    # ألوان التمييز
    'gold': 'C9A961',             # ذهبي راقي
    'gold_light': 'F5E6C8',       # ذهبي فاتح
    'accent_red': 'A93226',       # أحمر أنيق للتنبيهات
    'accent_green': '1E8449',     # أخضر للنجاح
    
    # ألوان محايدة
    'charcoal': '2C3E50',         # فحمي للنصوص
    'slate': '5D6D7E',            # رمادي أزرق
    'silver': 'BDC3C7',           # فضي للحدود
    'cream': 'FDFEFE',            # كريمي للخلفيات
    'white': 'FFFFFF',
}

# الخطوط العربية الاحترافية
ARABIC_FONTS = {
    'title': 'Traditional Arabic',    # للعناوين الرئيسية
    'heading': 'Simplified Arabic',   # للعناوين الفرعية
    'body': 'Calibri',                # للنصوص العادية
    'accent': 'Amiri',                # للتحية والخاتمة
}

# أحجام الخطوط المدروسة
FONT_SIZES = {
    'mega_title': 26,
    'title': 20,
    'heading': 16,
    'subheading': 14,
    'body': 13,
    'small': 11,
    'caption': 10,
    'footer': 9,
}

# ==========================================
# دوال مساعدة متقدمة
# ==========================================
def set_cell_shading(cell, color_hex):
    """تعيين لون خلفية خلية"""
    shading = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{color_hex}"/>')
    cell._tc.get_or_add_tcPr().append(shading)

def set_cell_gradient(cell, color1, color2):
    """تعيين تدرج لوني للخلية (يحاكى بلون صلب متقدم)"""
    set_cell_shading(cell, color1)

def set_paragraph_rtl(paragraph):
    """تعيين اتجاه الفقرة RTL مع محاذاة صحيحة"""
    pPr = paragraph._p.get_or_add_pPr()
    pPr.set(qn('w:bidi'), '1')
    pPr.set(qn('w:jc'), 'right')

def set_cell_rtl(cell):
    """تعيين اتجاه الخلية RTL"""
    for para in cell.paragraphs:
        set_paragraph_rtl(para)

def add_elegant_paragraph(doc, text, font_name=None, size=None, 
                          bold=False, italic=False, color=None,
                          alignment=None, space_before=0, space_after=8,
                          line_spacing=1.5, first_line_indent=None):
    """إضافة فقرة أنيقة بتنسيق احترافي"""
    para = doc.add_paragraph()
    
    # المحاذاة
    if alignment:
        para.alignment = alignment
    else:
        para.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    
    # المسافات والتباعد
    para.paragraph_format.space_before = Pt(space_before)
    para.paragraph_format.space_after = Pt(space_after)
    para.paragraph_format.line_spacing = line_spacing
    
    if first_line_indent:
        para.paragraph_format.first_line_indent = Cm(first_line_indent)
    
    # النص
    run = para.add_run(text)
    run.font.name = font_name or ARABIC_FONTS['body']
    run.font.size = Pt(size or FONT_SIZES['body'])
    run.font.bold = bold
    run.font.italic = italic
    
    if color:
        run.font.color.rgb = RGBColor.from_string(color)
    else:
        run.font.color.rgb = RGBColor.from_string(PALETTE['charcoal'])
    
    # تعيين RTL
    set_paragraph_rtl(para)
    
    return para

def add_decorative_line(doc, style='royal'):
    """إضافة خط فاصل مزخرف"""
    para = doc.add_paragraph()
    para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    para.paragraph_format.space_before = Pt(6)
    para.paragraph_format.space_after = Pt(6)
    
    pPr = para._p.get_or_add_pPr()
    
    if style == 'royal':
        # خط مزدوج أنيق
        pBdr = parse_xml(
            f'<w:pBdr {nsdecls("w")}>'
            f'  <w:top w:val="double" w:sz="12" w:space="4" w:color="{PALETTE["royal_dark"]}"/>'
            f'</w:pBdr>'
        )
    elif style == 'gold':
        # خط ذهبي رفيع
        pBdr = parse_xml(
            f'<w:pBdr {nsdecls("w")}>'
            f'  <w:bottom w:val="single" w:sz="8" w:space="1" w:color="{PALETTE["gold"]}"/>'
            f'</w:pBdr>'
        )
    else:
        pBdr = parse_xml(
            f'<w:pBdr {nsdecls("w")}>'
            f'  <w:bottom w:val="single" w:sz="4" w:space="1" w:color="{PALETTE["silver"]}"/>'
            f'</w:pBdr>'
        )
    
    pPr.append(pBdr)

def get_unique_filename(directory, base_name, extension):
    """إنشاء اسم ملف فريد"""
    counter = 1
    filename = f"{base_name}.{extension}"
    while os.path.exists(os.path.join(directory, filename)):
        filename = f"{base_name}_{counter}.{extension}"
        counter += 1
    return filename

# ==========================================
# 1. خطاب رسمي فاخر
# ==========================================
def create_professional_letter(
    recipient, subject, content_paragraphs,
    table_data=None, table_headers=None, table_caption=None,
    closing_note=None, closing_phrase=None,
    sender_name=None, sender_title=None, sender_organization=None,
    reference_number=None, attachments=None,
    output_dir=None, filename=None
):
    """إنشاء خطاب رسمي بمستوى احترافي راقٍ"""
    
    doc = Document()
    
    # ==========================================
    # إعدادات الصفحة والخطوط الافتراضية
    # ==========================================
    section = doc.sections[0]
    section.top_margin = Cm(3)
    section.bottom_margin = Cm(2.5)
    section.left_margin = Cm(2.5)
    section.right_margin = Cm(2.5)
    
    style = doc.styles['Normal']
    style.font.name = ARABIC_FONTS['body']
    style.font.size = Pt(FONT_SIZES['body'])
    style.font.color.rgb = RGBColor.from_string(PALETTE['charcoal'])
    style.paragraph_format.line_spacing = 1.5
    
    # ==========================================
    # 1. رأس الخطاب (الجهة المرسلة)
    # ==========================================
    if sender_organization:
        # اسم المؤسسة بخط كبير وأنيق
        org_para = doc.add_paragraph()
        org_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        org_para.paragraph_format.space_after = Pt(4)
        org_run = org_para.add_run(sender_organization)
        org_run.font.name = ARABIC_FONTS['title']
        org_run.font.size = Pt(FONT_SIZES['title'])
        org_run.font.bold = True
        org_run.font.color.rgb = RGBColor.from_string(PALETTE['royal_dark'])
        set_paragraph_rtl(org_para)
    
    if sender_title:
        title_para = doc.add_paragraph()
        title_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        title_para.paragraph_format.space_after = Pt(12)
        title_run = title_para.add_run(sender_title)
        title_run.font.name = ARABIC_FONTS['heading']
        title_run.font.size = Pt(FONT_SIZES['subheading'])
        title_run.font.color.rgb = RGBColor.from_string(PALETTE['royal_mid'])
        set_paragraph_rtl(title_para)
    
    add_decorative_line(doc, 'royal')
    
    # ==========================================
    # 2. العنوان الرئيسي (موضوع الخطاب)
    # ==========================================
    title = doc.add_heading(subject, level=0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title.paragraph_format.space_before = Pt(12)
    title.paragraph_format.space_after = Pt(12)
    
    for run in title.runs:
        run.font.name = ARABIC_FONTS['title']
        run.font.size = Pt(FONT_SIZES['mega_title'])
        run.font.bold = True
        run.font.color.rgb = RGBColor.from_string(PALETTE['royal_dark'])
    
    add_decorative_line(doc, 'gold')
    
    # ==========================================
    # 3. معلومات الخطاب (التاريخ والرقم)
    # ==========================================
    info_table = doc.add_table(rows=1, cols=2)
    info_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    info_table.autofit = True
    
    # التاريخ
    date_cell = info_table.cell(0, 0)
    date_para = date_cell.paragraphs[0]
    date_para.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    date_label = date_para.add_run("التاريخ: ")
    date_label.font.bold = True
    date_label.font.size = Pt(FONT_SIZES['body'])
    date_label.font.color.rgb = RGBColor.from_string(PALETTE['royal_mid'])
    date_value = date_para.add_run(datetime.now().strftime('%Y/%m/%d'))
    date_value.font.size = Pt(FONT_SIZES['body'])
    set_paragraph_rtl(date_para)
    
    # رقم المرجع
    ref_cell = info_table.cell(0, 1)
    ref_para = ref_cell.paragraphs[0]
    ref_para.alignment = WD_ALIGN_PARAGRAPH.LEFT
    ref_label = ref_para.add_run("رقم: ")
    ref_label.font.bold = True
    ref_label.font.size = Pt(FONT_SIZES['body'])
    ref_label.font.color.rgb = RGBColor.from_string(PALETTE['royal_mid'])
    ref_value = ref_para.add_run(reference_number or datetime.now().strftime('%Y/%m/%d/%H%M'))
    ref_value.font.size = Pt(FONT_SIZES['body'])
    
    doc.add_paragraph()
    
    # ==========================================
    # 4. المرسل إليه والموضوع
    # ==========================================
    to_para = doc.add_paragraph()
    to_para.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    to_para.paragraph_format.space_after = Pt(6)
    to_label = to_para.add_run("السادة/ ")
    to_label.font.name = ARABIC_FONTS['accent']
    to_label.font.bold = True
    to_label.font.size = Pt(FONT_SIZES['body'])
    to_label.font.color.rgb = RGBColor.from_string(PALETTE['royal_dark'])
    to_value = to_para.add_run(recipient)
    to_value.font.name = ARABIC_FONTS['accent']
    to_value.font.bold = True
    to_value.font.size = Pt(FONT_SIZES['body'])
    set_paragraph_rtl(to_para)
    
    subj_para = doc.add_paragraph()
    subj_para.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    subj_para.paragraph_format.space_after = Pt(12)
    subj_label = subj_para.add_run("الموضوع: ")
    subj_label.font.bold = True
    subj_label.font.size = Pt(FONT_SIZES['body'])
    subj_label.font.color.rgb = RGBColor.from_string(PALETTE['royal_dark'])
    subj_value = subj_para.add_run(subject)
    subj_value.font.bold = True
    subj_value.font.size = Pt(FONT_SIZES['body'])
    subj_value.font.color.rgb = RGBColor.from_string(PALETTE['royal_mid'])
    set_paragraph_rtl(subj_para)
    
    # ==========================================
    # 5. التحية
    # ==========================================
    add_elegant_paragraph(
        doc, "تحية طيبة وبعد،،،",
        font_name=ARABIC_FONTS['accent'],
        size=FONT_SIZES['body'],
        bold=True,
        color=PALETTE['royal_dark'],
        space_before=12, space_after=12
    )
    
    # ==========================================
    # 6. المقدمة الذكية
    # ==========================================
    add_elegant_paragraph(
        doc, "بالإشارة إلى الموضوع أعلاه، نتقدم إليكم بهذا الخطاب.",
        font_name=ARABIC_FONTS['body'],
        size=FONT_SIZES['body'],
        space_after=10,
        first_line_indent=1.5
    )
    
    # ==========================================
    # 7. فقرات المحتوى
    # ==========================================
    for i, para_text in enumerate(content_paragraphs):
        if len(content_paragraphs) > 1:
            numbered_text = f"{i+1}. {para_text}"
        else:
            numbered_text = para_text
        
        add_elegant_paragraph(
            doc, numbered_text,
            size=FONT_SIZES['body'],
            space_after=10,
            line_spacing=1.6,
            first_line_indent=1.5
        )
    
    doc.add_paragraph()
    
    # ==========================================
    # 8. الجدول الاحترافي
    # ==========================================
    if table_data and table_headers:
        if table_caption:
            cap_para = doc.add_paragraph()
            cap_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            cap_para.paragraph_format.space_before = Pt(12)
            cap_para.paragraph_format.space_after = Pt(8)
            cap_run = cap_para.add_run(table_caption)
            cap_run.font.name = ARABIC_FONTS['heading']
            cap_run.font.bold = True
            cap_run.font.size = Pt(FONT_SIZES['subheading'])
            cap_run.font.color.rgb = RGBColor.from_string(PALETTE['royal_dark'])
            set_paragraph_rtl(cap_para)
        
        # إنشاء الجدول
        table = doc.add_table(rows=1, cols=len(table_headers))
        table.style = 'Table Grid'
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        table.autofit = True
        
        # تنسيق رأس الجدول بتدرج لوني
        header_cells = table.rows[0].cells
        for i, header in enumerate(table_headers):
            cell = header_cells[i]
            cell.text = ''
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_before = Pt(4)
            p.paragraph_format.space_after = Pt(4)
            
            run = p.add_run(header)
            run.font.name = ARABIC_FONTS['heading']
            run.font.bold = True
            run.font.color.rgb = RGBColor.from_string(PALETTE['white'])
            run.font.size = Pt(FONT_SIZES['body'])
            
            # خلفية بتدرج أزرق ملكي
            set_cell_shading(cell, PALETTE['royal_dark'])
            set_cell_rtl(cell)
        
        # إضافة البيانات مع تلوين ذكي
        for row_idx, row_data in enumerate(table_data):
            row_cells = table.add_row().cells
            # تلوين الصفوف البديلة بألوان متناسقة
            bg_color = PALETTE['royal_pale'] if row_idx % 2 == 0 else PALETTE['cream']
            
            for col_idx, cell_data in enumerate(row_data):
                cell = row_cells[col_idx]
                cell.text = ''
                p = cell.paragraphs[0]
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p.paragraph_format.space_before = Pt(3)
                p.paragraph_format.space_after = Pt(3)
                
                run = p.add_run(str(cell_data))
                run.font.name = ARABIC_FONTS['body']
                run.font.size = Pt(FONT_SIZES['body'])
                run.font.color.rgb = RGBColor.from_string(PALETTE['charcoal'])
                
                set_cell_shading(cell, bg_color)
                set_cell_rtl(cell)
        
        doc.add_paragraph()
    
    # ==========================================
    # 9. الخاتمة الأنيقة
    # ==========================================
    closing = closing_phrase or "وتفضلوا بقبول فائق الاحترام والتقدير،،"
    add_elegant_paragraph(
        doc, closing,
        font_name=ARABIC_FONTS['accent'],
        size=FONT_SIZES['body'],
        bold=True,
        color=PALETTE['royal_dark'],
        space_before=12, space_after=12
    )
    
    # ==========================================
    # 10. الملاحظة الختامية
    # ==========================================
    if closing_note:
        doc.add_paragraph()
        note_para = doc.add_paragraph()
        note_para.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        note_para.paragraph_format.space_before = Pt(6)
        
        note_icon = note_para.add_run("⚠ تنبيه: ")
        note_icon.font.bold = True
        note_icon.font.size = Pt(FONT_SIZES['body'])
        note_icon.font.color.rgb = RGBColor.from_string(PALETTE['accent_red'])
        
        note_text = note_para.add_run(closing_note)
        note_text.font.size = Pt(FONT_SIZES['body'])
        note_text.font.color.rgb = RGBColor.from_string(PALETTE['accent_red'])
        set_paragraph_rtl(note_para)
    
    # ==========================================
    # 11. المرفقات
    # ==========================================
    if attachments:
        doc.add_paragraph()
        att_label_para = doc.add_paragraph()
        att_label_para.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        att_label = att_label_para.add_run("المرفقات: ")
        att_label.font.bold = True
        att_label.font.size = Pt(FONT_SIZES['body'])
        att_label.font.color.rgb = RGBColor.from_string(PALETTE['royal_dark'])
        set_paragraph_rtl(att_label_para)
        
        for att in attachments:
            att_item = doc.add_paragraph()
            att_item.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            att_item.paragraph_format.left_indent = Cm(1)
            att_run = att_item.add_run(f"• {att}")
            att_run.font.size = Pt(FONT_SIZES['body'])
            set_paragraph_rtl(att_item)
    
    # ==========================================
    # 12. التوقيع
    # ==========================================
    doc.add_paragraph()
    doc.add_paragraph()
    
    if sender_name or sender_organization:
        sig_para = doc.add_paragraph()
        sig_para.alignment = WD_ALIGN_PARAGRAPH.LEFT
        
        if sender_name:
            name_run = sig_para.add_run(sender_name)
            name_run.font.name = ARABIC_FONTS['accent']
            name_run.font.bold = True
            name_run.font.size = Pt(FONT_SIZES['body'])
            name_run.font.color.rgb = RGBColor.from_string(PALETTE['royal_dark'])
        
        if sender_title:
            sig_para.add_run('\n')
            title_run = sig_para.add_run(sender_title)
            title_run.font.size = Pt(FONT_SIZES['small'])
            title_run.font.color.rgb = RGBColor.from_string(PALETTE['royal_mid'])
        
        if sender_organization:
            sig_para.add_run('\n')
            org_run = sig_para.add_run(sender_organization)
            org_run.font.size = Pt(FONT_SIZES['small'])
            org_run.font.italic = True
            org_run.font.color.rgb = RGBColor.from_string(PALETTE['slate'])
    
    # ==========================================
    # تذييل الصفحة الاحترافي
    # ==========================================
    section = doc.sections[0]
    footer = section.footer
    footer_para = footer.paragraphs[0]
    footer_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    footer_line = footer_para.add_run("━" * 40 + "\n")
    footer_line.font.color.rgb = RGBColor.from_string(PALETTE['silver'])
    footer_line.font.size = Pt(FONT_SIZES['caption'])
    
    footer_text = footer_para.add_run("Officia - الوكيل الذكي للمهام المكتبية")
    footer_text.font.size = Pt(FONT_SIZES['footer'])
    footer_text.font.color.rgb = RGBColor.from_string(PALETTE['slate'])
    footer_text.font.italic = True
    
    # ==========================================
    # الحفظ
    # ==========================================
    from core.config import WORKSPACE_DIR
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
# 2. تقرير Excel احترافي متقدم
# ==========================================
def create_excel_report(
    headers, data, title=None, sheet_name="التقرير",
    add_chart=False, chart_type="bar",
    freeze_panes=True, auto_filter=True,
    output_dir=None, filename=None
):
    """إنشاء تقرير Excel بمستوى احترافي راقٍ"""
    
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = sheet_name
    ws.sheet_view.rightToLeft = True
    
    current_row = 1
    
    # ==========================================
    # العنوان الرئيسي
    # ==========================================
    if title:
        ws.merge_cells(start_row=1, start_column=1, 
                      end_row=1, end_column=len(headers))
        title_cell = ws.cell(row=1, column=1, value=title)
        title_cell.font = Font(
            name=ARABIC_FONTS['title'],
            size=FONT_SIZES['title'],
            bold=True,
            color=PALETTE['royal_dark']
        )
        title_cell.alignment = Alignment(horizontal='center', vertical='center')
        title_cell.fill = PatternFill(
            start_color=PALETTE['royal_pale'],
            end_color=PALETTE['royal_pale'],
            fill_type='solid'
        )
        ws.row_dimensions[1].height = 40
        current_row = 3
    
    # ==========================================
    # رؤوس الأعمدة بتدرج لوني
    # ==========================================
    header_fill = PatternFill(
        start_color=PALETTE['royal_dark'],
        end_color=PALETTE['royal_mid'],
        fill_type='linearGradient'
    )
    header_font = Font(
        name=ARABIC_FONTS['heading'],
        size=FONT_SIZES['body'],
        bold=True,
        color=PALETTE['white']
    )
    header_alignment = Alignment(
        horizontal='center',
        vertical='center',
        wrap_text=True
    )
    
    # حدود أنيقة
    thin_border = Border(
        left=Side(style='thin', color=PALETTE['silver']),
        right=Side(style='thin', color=PALETTE['silver']),
        top=Side(style='thin', color=PALETTE['silver']),
        bottom=Side(style='medium', color=PALETTE['royal_dark'])
    )
    
    for col_num, header in enumerate(headers, 1):
        cell = ws.cell(row=current_row, column=col_num, value=header)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = header_alignment
        cell.border = thin_border
    
    ws.row_dimensions[current_row].height = 30
    data_start_row = current_row + 1
    
    # ==========================================
    # البيانات بتنسيق متقدم
    # ==========================================
    data_font = Font(name=ARABIC_FONTS['body'], size=FONT_SIZES['body'])
    data_alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    
    # تعبئات الصفوف
    even_fill = PatternFill(start_color=PALETTE['royal_pale'], end_color=PALETTE['royal_pale'], fill_type='solid')
    odd_fill = PatternFill(start_color=PALETTE['cream'], end_color=PALETTE['cream'], fill_type='solid')
    
    for row_idx, row_data in enumerate(data):
        for col_idx, cell_data in enumerate(row_data, 1):
            cell = ws.cell(row=data_start_row + row_idx, column=col_idx, value=cell_data)
            cell.font = data_font
            cell.alignment = data_alignment
            cell.border = thin_border
            
            # تلوين الصفوف البديلة
            cell.fill = even_fill if row_idx % 2 == 0 else odd_fill
            
            # تنسيق الأرقام
            if isinstance(cell_data, (int, float)):
                cell.number_format = '#,##0.00'
    
    # ==========================================
    # تجميد الصف الأول
    # ==========================================
    if freeze_panes:
        ws.freeze_panes = f'A{data_start_row}'
    
    # ==========================================
    # الفلتر التلقائي
    # ==========================================
    if auto_filter:
        ws.auto_filter.ref = f'A{current_row}:{get_column_letter(len(headers))}{data_start_row + len(data) - 1}'
    
    # ==========================================
    # ضبط عرض الأعمدة بدقة
    # ==========================================
    for col_num in range(1, len(headers) + 1):
        max_length = len(str(headers[col_num - 1])) * 1.5
        for row in ws.iter_rows(min_row=data_start_row, max_row=data_start_row + len(data) - 1, 
                               min_col=col_num, max_col=col_num):
            for cell in row:
                if cell.value:
                    cell_length = len(str(cell.value)) * 1.2
                    max_length = max(max_length, cell_length)
        adjusted_width = min(max_length + 4, 45)
        ws.column_dimensions[get_column_letter(col_num)].width = adjusted_width
    
    # ==========================================
    # الرسم البياني الاحترافي
    # ==========================================
    if add_chart and len(data) > 1:
        if chart_type == 'bar':
            chart = BarChart()
            chart.type = 'col'
            chart.title = title or 'رسم بياني'
            chart.y_axis.title = 'القيمة'
            chart.x_axis.title = headers[0] if headers else ''
            chart.style = 10
            
            data_ref = Reference(ws, min_col=2, min_row=current_row, 
                                max_row=data_start_row + len(data) - 1)
            cats_ref = Reference(ws, min_col=1, min_row=data_start_row, 
                                max_row=data_start_row + len(data) - 1)
            chart.add_data(data_ref, titles_from_data=True)
            chart.set_categories(cats_ref)
            
            ws.add_chart(chart, f'A{data_start_row + len(data) + 3}')
    
    # ==========================================
    # الحفظ
    # ==========================================
    from core.config import WORKSPACE_DIR
    save_dir = output_dir or WORKSPACE_DIR
    os.makedirs(save_dir, exist_ok=True)
    
    if filename is None:
        safe_title = (title or 'تقرير')[:30].replace(' ', '_').replace('/', '-')
        filename = f"{safe_title}"
    
    filename = get_unique_filename(save_dir, filename, 'xlsx')
    output_path = os.path.join(save_dir, filename)
    
    wb.save(output_path)
    return output_path
