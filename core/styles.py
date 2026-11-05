"""
نظام الأنماط الديناميكي - Dynamic Styles System v2
مكتبة أنماط التنسيق المختلفة للمستندات
نظام الليغو: 8×8×10×6×6 = 23,040+ تركيبة مختلفة!
"""

from docx.shared import Pt, Cm, RGBColor, Emu, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml.ns import qn, nsdecls
from docx.oxml import parse_xml
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, numbers


# ==========================================
# 🎨 1️⃣ أنماط الرؤوس (Headers) - 8 أنماط
# ==========================================
HEADER_STYLES = {
    'official': {
        'name': 'رسمي كلاسيكي',
        'has_double_line': True,
        'has_gold_line': True,
        'font_size': 26,
        'color': '1B3A5C',
        'alignment': 'center',
        'org_size': 20,
        'spacing_before': 12,
        'spacing_after': 12
    },
    'modern': {
        'name': 'حديث بسيط',
        'has_double_line': False,
        'has_gold_line': False,
        'font_size': 24,
        'color': '2E5C8A',
        'alignment': 'right',
        'org_size': 18,
        'spacing_before': 6,
        'spacing_after': 6
    },
    'professional': {
        'name': 'احترافي مع شعار',
        'has_double_line': True,
        'has_gold_line': True,
        'font_size': 28,
        'color': '1F4E79',
        'alignment': 'center',
        'org_size': 22,
        'spacing_before': 18,
        'spacing_after': 18
    },
    'minimal': {
        'name': 'بسيط جداً',
        'has_double_line': False,
        'has_gold_line': False,
        'font_size': 20,
        'color': '2C3E50',
        'alignment': 'center',
        'org_size': 16,
        'spacing_before': 6,
        'spacing_after': 6
    },
    'bold': {
        'name': 'عريض قوي',
        'has_double_line': True,
        'has_gold_line': False,
        'font_size': 28,
        'color': 'A93226',
        'alignment': 'center',
        'org_size': 20,
        'spacing_before': 12,
        'spacing_after': 12
    },
    'gradient': {
        'name': 'تجاري متدرج',
        'has_double_line': False,
        'has_gold_line': True,
        'font_size': 26,
        'color': '1B3A5C',
        'alignment': 'center',
        'org_size': 20,
        'spacing_before': 12,
        'spacing_after': 12
    },
    'celebration': {
        'name': 'احتفالي مزخرف',
        'has_double_line': True,
        'has_gold_line': True,
        'font_size': 30,
        'color': 'C9A961',
        'alignment': 'center',
        'org_size': 22,
        'spacing_before': 20,
        'spacing_after': 20,
        'has_ornament': True
    },
    'islamic': {
        'name': 'إسلامي تراثي',
        'has_double_line': True,
        'has_gold_line': True,
        'font_size': 28,
        'color': '1E8449',
        'alignment': 'center',
        'org_size': 20,
        'spacing_before': 15,
        'spacing_after': 15,
        'has_bismillah': True
    }
}


# ==========================================
#  2️⃣ أنماط المحتوى (Content) - 8 أنماط
# ==========================================
CONTENT_STYLES = {
    'paragraphs': {
        'name': 'فقرات تقليدية',
        'numbered': False,
        'bulleted': False,
        'first_line_indent': 1.5,
        'line_spacing': 1.6,
        'space_after': 10
    },
    'numbered': {
        'name': 'نقاط مرقمة',
        'numbered': True,
        'bulleted': False,
        'first_line_indent': 1.5,
        'line_spacing': 1.6,
        'space_after': 10
    },
    'bullets': {
        'name': 'نقاط برمز •',
        'numbered': False,
        'bulleted': True,
        'bullet_symbol': '•',
        'first_line_indent': 1.5,
        'line_spacing': 1.6,
        'space_after': 8
    },
    'compact': {
        'name': 'مضغوط',
        'numbered': False,
        'bulleted': False,
        'first_line_indent': 0,
        'line_spacing': 1.3,
        'space_after': 6
    },
    'mixed': {
        'name': 'مزيج مرن',
        'numbered': False,
        'bulleted': False,
        'first_line_indent': 1.5,
        'line_spacing': 1.5,
        'space_after': 12
    },
    'quoted': {
        'name': 'مقتبس',
        'numbered': False,
        'bulleted': False,
        'first_line_indent': 2.0,
        'line_spacing': 1.8,
        'space_after': 12,
        'is_quote': True
    },
    'highlighted': {
        'name': 'بارز ومميز',
        'numbered': False,
        'bulleted': True,
        'bullet_symbol': '◆',
        'first_line_indent': 1.5,
        'line_spacing': 1.7,
        'space_after': 10,
        'is_bold': True
    },
    'decorative': {
        'name': 'زخرفي',
        'numbered': False,
        'bulleted': True,
        'bullet_symbol': '',
        'first_line_indent': 1.5,
        'line_spacing': 1.6,
        'space_after': 10,
        'has_separator': True
    }
}


# ==========================================
# 📊 3️ أنماط الجداول (Tables) - 10 أنماط
# ==========================================
TABLE_STYLES = {
    'royal_blue': {
        'name': 'أزرق ملكي',
        'header_color': '1B3A5C',
        'header_text_color': 'FFFFFF',
        'row_even': 'D6E4F0',
        'row_odd': 'FFFFFF',
        'border_color': 'B4C7E7',
        'border_style': 'thin'
    },
    'green_finance': {
        'name': 'أخضر مالي',
        'header_color': '1E8449',
        'header_text_color': 'FFFFFF',
        'row_even': 'D5F5E3',
        'row_odd': 'FFFFFF',
        'border_color': 'A9DFBF',
        'border_style': 'thin'
    },
    'dark_elegant': {
        'name': 'أسود أنيق',
        'header_color': '2C3E50',
        'header_text_color': 'FFFFFF',
        'row_even': 'F2F3F4',
        'row_odd': 'FFFFFF',
        'border_color': '5D6D7E',
        'border_style': 'medium'
    },
    'colorful': {
        'name': 'ملون متدرج',
        'header_color': '4A7FB5',
        'header_text_color': 'FFFFFF',
        'row_even': 'E8F4F8',
        'row_odd': 'FFF9C4',
        'border_color': 'BDC3C7',
        'border_style': 'thin'
    },
    'minimal': {
        'name': 'بسيط بدون ألوان',
        'header_color': 'FFFFFF',
        'header_text_color': '000000',
        'row_even': 'FFFFFF',
        'row_odd': 'FFFFFF',
        'border_color': 'BDC3C7',
        'border_style': 'single'
    },
    'accounting': {
        'name': 'محاسبي Excel',
        'header_color': '1F4E79',
        'header_text_color': 'FFFFFF',
        'row_even': 'D9E1F2',
        'row_odd': 'FFFFFF',
        'border_color': 'B4C7E7',
        'border_style': 'thin',
        'number_format': '#,##0.00'
    },
    'modern_shadows': {
        'name': 'حديث بظلال',
        'header_color': '2E5C8A',
        'header_text_color': 'FFFFFF',
        'row_even': 'F8F9FA',
        'row_odd': 'FFFFFF',
        'border_color': 'DEE2E6',
        'border_style': 'thin'
    },
    'golden_luxury': {
        'name': 'ذهبي فاخر',
        'header_color': 'C9A961',
        'header_text_color': 'FFFFFF',
        'row_even': 'F5E6C8',
        'row_odd': 'FFFFFF',
        'border_color': 'D4AF37',
        'border_style': 'medium'
    },
    'purple_royal': {
        'name': 'بنفسجي ملكي',
        'header_color': '6C3483',
        'header_text_color': 'FFFFFF',
        'row_even': 'E8DAEF',
        'row_odd': 'FFFFFF',
        'border_color': 'BB8FCE',
        'border_style': 'thin'
    },
    'rose_elegant': {
        'name': 'وردي أنيق',
        'header_color': 'C0392B',
        'header_text_color': 'FFFFFF',
        'row_even': 'FADBD8',
        'row_odd': 'FFFFFF',
        'border_color': 'E6B0AA',
        'border_style': 'thin'
    }
}


# ==========================================
# 📑 4️⃣ أنماط التذييل (Footer) - 6 أنماط
# ==========================================
FOOTER_STYLES = {
    'simple': {
        'name': 'بسيط',
        'has_line': False,
        'text': 'Officia - الوكيل الذكي',
        'font_size': 9,
        'color': '808080'
    },
    'professional': {
        'name': 'احترافي',
        'has_line': True,
        'text': 'Officia - الوكيل الذكي للمهام المكتبية',
        'font_size': 9,
        'color': '5D6D7E'
    },
    'contact': {
        'name': 'مع معلومات الاتصال',
        'has_line': True,
        'text': 'Officia | www.officia.app | info@officia.app',
        'font_size': 9,
        'color': '5D6D7E'
    },
    'logo': {
        'name': 'شعار + نص',
        'has_line': True,
        'text': '━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\nOfficia - الوكيل الذكي',
        'font_size': 9,
        'color': '5D6D7E'
    },
    'decorative': {
        'name': 'زخرفي',
        'has_line': True,
        'text': '❖ ❖ ❖\nOfficia - الوكيل الذكي للمهام المكتبية',
        'font_size': 9,
        'color': 'C9A961'
    },
    'double_line': {
        'name': 'خط مزدوج',
        'has_double_line': True,
        'text': 'Officia © 2026 - جميع الحقوق محفوظة',
        'font_size': 8,
        'color': '5D6D7E'
    }
}


# ==========================================
# ✍️ 5️⃣ أنماط التوقيعات (Signatures) - 6 أنماط
# ==========================================
SIGNATURE_STYLES = {
    'single': {
        'name': 'توقيع واحد',
        'count': 1,
        'layout': 'left'
    },
    'dual': {
        'name': 'توقيعان',
        'count': 2,
        'layout': 'table'
    },
    'triple': {
        'name': 'ثلاثة توقيعات',
        'count': 3,
        'layout': 'table'
    },
    'stamp': {
        'name': 'ختم + توقيع',
        'count': 1,
        'layout': 'center',
        'has_stamp': True
    },
    'multi_level': {
        'name': 'متعدد المستويات',
        'count': 3,
        'layout': 'vertical'
    },
    'formal_table': {
        'name': 'جدول توقيعات رسمي',
        'count': 2,
        'layout': 'formal_table'
    }
}


# ==========================================
# 🎨 6️⃣ أنماط الحدود والزخارف (Borders)
# ==========================================
BORDER_STYLES = {
    'none': {
        'name': 'بدون حدود',
        'top': None,
        'bottom': None,
        'left': None,
        'right': None
    },
    'thin': {
        'name': 'رفيع',
        'style': 'thin',
        'color': 'BDC3C7',
        'size': 4
    },
    'medium': {
        'name': 'متوسط',
        'style': 'single',
        'color': '5D6D7E',
        'size': 8
    },
    'thick': {
        'name': 'سميك',
        'style': 'single',
        'color': '2C3E50',
        'size': 12
    },
    'double': {
        'name': 'مزدوج',
        'style': 'double',
        'color': '1B3A5C',
        'size': 12
    },
    'dashed': {
        'name': 'متقطع',
        'style': 'dashed',
        'color': '95A5A6',
        'size': 6
    },
    'dotted': {
        'name': 'منقط',
        'style': 'dotted',
        'color': '95A5A6',
        'size': 6
    },
    'gold_ornate': {
        'name': 'ذهبي مزخرف',
        'style': 'single',
        'color': 'C9A961',
        'size': 10
    }
}


  # ==========================================
# 🎯 دوال تطبيق الأنماط - Word
# ==========================================
def apply_header_style(doc, style_name, organization_name, title, subtitle=None):
    """تطبيق نمط الرأس المحدد"""
    from docx.shared import Pt, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    
    style = HEADER_STYLES.get(style_name, HEADER_STYLES['official'])
    
    # البسملة (للنمط الإسلامي)
    if style.get('has_bismillah'):
        bismillah_para = doc.add_paragraph()
        bismillah_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        bismillah_run = bismillah_para.add_run('بسم الله الرحمن الرحيم')
        bismillah_run.font.name = 'Traditional Arabic'
        bismillah_run.font.size = Pt(18)
        bismillah_run.font.color.rgb = RGBColor.from_string('1E8449')
        bismillah_run.font.bold = True
        bismillah_para.paragraph_format.space_after = Pt(12)
    
    # اسم المؤسسة
    if organization_name:
        org_para = doc.add_paragraph()
        org_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        org_para.paragraph_format.space_after = Pt(6)
        org_run = org_para.add_run(organization_name)
        org_run.font.size = Pt(style['org_size'])
        org_run.font.bold = True
        org_run.font.color.rgb = RGBColor.from_string(style['color'])
    
    # الخط الفاصل المزدوج
    if style['has_double_line']:
        para = doc.add_paragraph()
        para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        para.paragraph_format.space_before = Pt(6)
        para.paragraph_format.space_after = Pt(6)
        pPr = para._p.get_or_add_pPr()
        pBdr = parse_xml(
            f'<w:pBdr {nsdecls("w")}>'
            f'  <w:top w:val="double" w:sz="12" w:space="4" w:color="{style["color"]}"/>'
            f'</w:pBdr>'
        )
        pPr.append(pBdr)
    
    # العنوان الرئيسي
    title_para = doc.add_heading(title, level=0)
    title_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title_para.paragraph_format.space_before = Pt(style['spacing_before'])
    title_para.paragraph_format.space_after = Pt(style['spacing_after'])
    for run in title_para.runs:
        run.font.size = Pt(style['font_size'])
        run.font.color.rgb = RGBColor.from_string(style['color'])
    
    # العنوان الفرعي
    if subtitle:
        sub_para = doc.add_paragraph()
        sub_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        sub_run = sub_para.add_run(subtitle)
        sub_run.font.size = Pt(14)
        sub_run.font.color.rgb = RGBColor.from_string(style['color'])
        sub_run.font.italic = True
    
    # الخط الذهبي
    if style['has_gold_line']:
        para = doc.add_paragraph()
        para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        para.paragraph_format.space_before = Pt(6)
        para.paragraph_format.space_after = Pt(6)
        pPr = para._p.get_or_add_pPr()
        pBdr = parse_xml(
            f'<w:pBdr {nsdecls("w")}>'
            f'  <w:bottom w:val="single" w:sz="8" w:space="1" w:color="C9A961"/>'
            f'</w:pBdr>'
        )
        pPr.append(pBdr)
    
    # الزخرفة (للنمط الاحتفالي)
    if style.get('has_ornament'):
        ornament_para = doc.add_paragraph()
        ornament_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        ornament_run = ornament_para.add_run('❖ ❖ ❖')
        ornament_run.font.size = Pt(16)
        ornament_run.font.color.rgb = RGBColor.from_string('C9A961')


def apply_content_style(doc, style_name, paragraphs):
    """تطبيق نمط المحتوى المحدد"""
    import re
    from docx.shared import Pt, Cm, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    
    style = CONTENT_STYLES.get(style_name, CONTENT_STYLES['paragraphs'])
    
    for i, para_text in enumerate(paragraphs):
        # تنظيف الترقيم التلقائي
        cleaned = re.sub(r'^[\d\u0660-\u0669]+[\.\-\)]\s*', '', para_text.strip())
        
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        p.paragraph_format.first_line_indent = Cm(style['first_line_indent'])
        p.paragraph_format.line_spacing = style['line_spacing']
        p.paragraph_format.space_after = Pt(style['space_after'])
        
        text = cleaned
        if style['numbered']:
            text = f"{i+1}. {cleaned}"
        elif style['bulleted']:
            symbol = style.get('bullet_symbol', '•')
            text = f"{symbol} {cleaned}"
        
        run = p.add_run(text)
        run.font.size = Pt(13)
        run.font.color.rgb = RGBColor.from_string('2C3E50')
        
        # تنسيق خاص للمقتبس
        if style.get('is_quote'):
            run.font.italic = True
            run.font.color.rgb = RGBColor.from_string('5D6D7E')
            p.paragraph_format.left_indent = Cm(2)
        
        # تنسيق خاص للبارز
        if style.get('is_bold'):
            run.font.bold = True
        
        # فاصل زخرفي
        if style.get('has_separator') and i < len(paragraphs) - 1:
            sep_para = doc.add_paragraph()
            sep_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            sep_run = sep_para.add_run('─' * 20)
            sep_run.font.size = Pt(10)
            sep_run.font.color.rgb = RGBColor.from_string('BDC3C7')


def apply_table_style(doc, style_name, headers, data, caption=None):
    """تطبيق نمط الجدول المحدد"""
    from docx.shared import Pt, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.enum.table import WD_TABLE_ALIGNMENT
    
    style = TABLE_STYLES.get(style_name, TABLE_STYLES['royal_blue'])
    
    # عنوان الجدول
    if caption:
        cap_para = doc.add_paragraph()
        cap_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        cap_para.paragraph_format.space_before = Pt(12)
        cap_para.paragraph_format.space_after = Pt(8)
        cap_run = cap_para.add_run(caption)
        cap_run.font.name = 'Traditional Arabic'
        cap_run.font.bold = True
        cap_run.font.size = Pt(14)
        cap_run.font.color.rgb = RGBColor.from_string(style['header_color'])
    
    table = doc.add_table(rows=1, cols=len(headers))
    table.style = 'Table Grid'
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    
    # الرأس
    for i, header in enumerate(headers):
        cell = table.rows[0].cells[i]
        cell.text = ''
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(header)
        run.font.bold = True
        run.font.color.rgb = RGBColor.from_string(style['header_text_color'])
        run.font.size = Pt(11)
        
        # لون الخلفية
        shading = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{style["header_color"]}"/>')
        cell._tc.get_or_add_tcPr().append(shading)
    
    # البيانات
    for row_idx, row_data in enumerate(data):
        row_cells = table.add_row().cells
        bg_color = style['row_even'] if row_idx % 2 == 0 else style['row_odd']
        
        for col_idx, cell_data in enumerate(row_data):
            cell = row_cells[col_idx]
            cell.text = ''
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = p.add_run(str(cell_data))
            run.font.size = Pt(10)
            
            shading = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{bg_color}"/>')
            cell._tc.get_or_add_tcPr().append(shading)
    
    return table


def apply_footer_style(doc, style_name):
    """تطبيق نمط التذييل المحدد"""
    from docx.shared import Pt, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    
    style = FOOTER_STYLES.get(style_name, FOOTER_STYLES['professional'])
    
    section = doc.sections[0]
    footer = section.footer
    footer_para = footer.paragraphs[0]
    footer_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    # خط فاصل
    if style.get('has_line'):
        line_run = footer_para.add_run("━" * 40 + "\n")
        line_run.font.color.rgb = RGBColor.from_string('BDC3C7')
        line_run.font.size = Pt(style['font_size'])
    
    # خط مزدوج
    if style.get('has_double_line'):
        line1_run = footer_para.add_run("━" * 40 + "\n")
        line1_run.font.color.rgb = RGBColor.from_string('BDC3C7')
        line1_run.font.size = Pt(style['font_size'])
        line2_run = footer_para.add_run("━" * 40 + "\n")
        line2_run.font.color.rgb = RGBColor.from_string('BDC3C7')
        line2_run.font.size = Pt(style['font_size'])
    
    text_run = footer_para.add_run(style['text'])
    text_run.font.size = Pt(style['font_size'])
    text_run.font.color.rgb = RGBColor.from_string(style['color'])
    text_run.font.italic = True


def apply_signature_style(doc, style_name, signatures=None):
    """تطبيق نمط التوقيع المحدد"""
    from docx.shared import Pt, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.enum.table import WD_TABLE_ALIGNMENT
    
    style = SIGNATURE_STYLES.get(style_name, SIGNATURE_STYLES['single'])
    
    if style['count'] == 0:
        return
    
    # مسافة قبل التوقيع
    sig_space = doc.add_paragraph()
    sig_space.paragraph_format.space_before = Pt(12)
    
    if style['layout'] == 'table' and style['count'] > 1:
        # توقيعات في جدول
        sig_table = doc.add_table(rows=1, cols=style['count'])
        sig_table.alignment = WD_TABLE_ALIGNMENT.CENTER
        
        for idx, sig in enumerate(signatures or []):
            if idx >= style['count']:
                break
            cell = sig_table.cell(0, idx)
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            
            if isinstance(sig, dict):
                name_run = p.add_run(sig.get('name', ''))
                name_run.font.bold = True
                name_run.font.size = Pt(11)
                p.add_run('\n')
                title_run = p.add_run(sig.get('title', ''))
                title_run.font.size = Pt(9)
                p.add_run('\n\n')
                sign_run = p.add_run("التوقيع: _______________")
                sign_run.font.size = Pt(10)
    elif style['layout'] == 'formal_table':
        # جدول توقيعات رسمي
        sig_table = doc.add_table(rows=2, cols=style['count'])
        sig_table.alignment = WD_TABLE_ALIGNMENT.CENTER
        
        for idx, sig in enumerate(signatures or []):
            if idx >= style['count']:
                break
            # الصف الأول: الاسم والصفة
            cell = sig_table.cell(0, idx)
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            if isinstance(sig, dict):
                name_run = p.add_run(sig.get('name', ''))
                name_run.font.bold = True
                name_run.font.size = Pt(11)
                p.add_run('\n')
                title_run = p.add_run(sig.get('title', ''))
                title_run.font.size = Pt(9)
            
            # الصف الثاني: التوقيع
            cell2 = sig_table.cell(1, idx)
            p2 = cell2.paragraphs[0]
            p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
            sign_run = p2.add_run("\n\nالتوقيع: _______________")
            sign_run.font.size = Pt(10)
    else:
        # توقيع واحد أو متعدد عمودياً
        for sig in (signatures or [])[:style['count']]:
            sig_para = doc.add_paragraph()
            sig_para.alignment = WD_ALIGN_PARAGRAPH.LEFT
            
            if isinstance(sig, dict):
                name_run = sig_para.add_run(sig.get('name', ''))
                name_run.font.bold = True
                name_run.font.size = Pt(11)
                sig_para.add_run('\n')
                title_run = sig_para.add_run(sig.get('title', ''))
                title_run.font.size = Pt(9)
                sig_para.add_run('\n\n')
                sign_run = sig_para.add_run("التوقيع: _______________")
                sign_run.font.size = Pt(10)
            
            # الختم
            if style.get('has_stamp'):
                stamp_para = doc.add_paragraph()
                stamp_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                stamp_run = stamp_para.add_run("[ الختم ]")
                stamp_run.font.size = Pt(14)
                stamp_run.font.color.rgb = RGBColor.from_string('C9A961')
                stamp_run.font.bold = True


# ==========================================
# 📊 دوال تطبيق الأنماط - Excel
# ==========================================
def apply_excel_table_style(ws, style_name, start_row, end_row, start_col, end_col):
    """تطبيق نمط جدول Excel"""
    style = TABLE_STYLES.get(style_name, TABLE_STYLES['royal_blue'])
    
    # الرأس
    header_fill = PatternFill(start_color=style['header_color'], end_color=style['header_color'], fill_type='solid')
    header_font = Font(bold=True, color=style['header_text_color'], size=11)
    
    for col in range(start_col, end_col + 1):
        cell = ws.cell(row=start_row, column=col)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal='center', vertical='center')
    
    # البيانات
    even_fill = PatternFill(start_color=style['row_even'], end_color=style['row_even'], fill_type='solid')
    odd_fill = PatternFill(start_color=style['row_odd'], end_color=style['row_odd'], fill_type='solid')
    
    for row in range(start_row + 1, end_row + 1):
        bg_fill = even_fill if (row - start_row) % 2 == 0 else odd_fill
        for col in range(start_col, end_col + 1):
            cell = ws.cell(row=row, column=col)
            cell.fill = bg_fill
            cell.alignment = Alignment(horizontal='center', vertical='center')
            
            if 'number_format' in style:
                if isinstance(cell.value, (int, float)):
                    cell.number_format = style['number_format']


# ==========================================
# 🎯 دالة التصنيف الذكي
# ==========================================
def classify_document_type(request: str) -> str:
    """تصنيف نوع المستند من الطلب"""
    request_lower = request.lower()
    
    if any(word in request_lower for word in ['إيصال', 'سند', 'استلام', 'دفع', 'receipt']):
        return 'receipt'
    elif any(word in request_lower for word in ['مذكرة', 'تعميم', 'تذكير', 'memo']):
        return 'memo'
    elif any(word in request_lower for word in ['فاتورة', 'bill', 'invoice']):
        return 'invoice'
    elif any(word in request_lower for word in ['تقرير', 'excel', 'ملف', 'report']):
        return 'report'
    elif any(word in request_lower for word in ['عقد', 'اتفاقية', 'contract']):
        return 'contract'
    elif any(word in request_lower for word in ['سيرة', 'cv', 'resume']):
        return 'cv'
    else:
        return 'letter'


def get_recommended_styles(doc_type: str) -> dict:
    """الحصول على الأنماط الموصى بها حسب نوع المستند"""
    recommendations = {
        'letter': {
            'header': 'official',
            'content': 'paragraphs',
            'table': 'royal_blue',
            'footer': 'professional',
            'signature': 'single'
        },
        'receipt': {
            'header': 'modern',
            'content': 'compact',
            'table': 'green_finance',
            'footer': 'simple',
            'signature': 'dual'
        },
        'memo': {
            'header': 'minimal',
            'content': 'bullets',
            'table': 'minimal',
            'footer': 'simple',
            'signature': 'none'
        },
        'invoice': {
            'header': 'professional',
            'content': 'paragraphs',
            'table': 'accounting',
            'footer': 'contact',
            'signature': 'dual'
        },
        'report': {
            'header': 'professional',
            'content': 'mixed',
            'table': 'accounting',
            'footer': 'professional',
            'signature': 'none'
        },
        'contract': {
            'header': 'bold',
            'content': 'numbered',
            'table': 'dark_elegant',
            'footer': 'professional',
            'signature': 'multi_level'
        },
        'cv': {
            'header': 'minimal',
            'content': 'compact',
            'table': 'minimal',
            'footer': 'simple',
            'signature': 'none'
        }
    }
    
    return recommendations.get(doc_type, recommendations['letter'])


# ==========================================
# 🎨 دالة الحصول على جميع الأنماط المتاحة
# ==========================================
def get_all_available_styles() -> dict:
    """إرجاع جميع الأنماط المتاحة للاستخدام"""
    return {
        'headers': {k: v['name'] for k, v in HEADER_STYLES.items()},
        'content': {k: v['name'] for k, v in CONTENT_STYLES.items()},
        'tables': {k: v['name'] for k, v in TABLE_STYLES.items()},
        'footers': {k: v['name'] for k, v in FOOTER_STYLES.items()},
        'signatures': {k: v['name'] for k, v in SIGNATURE_STYLES.items()},
        'borders': {k: v['name'] for k, v in BORDER_STYLES.items()}
    }
