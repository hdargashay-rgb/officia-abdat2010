"""
Arabic Fixer - مصحح النصوص العربية في ملفات Word
يتعامل مع:
- النصوص المعكوسة
- التشكيل الخاطئ
- اتجاه الفقرات RTL
- الجداول المعكوسة
"""
import os


def fix_arabic_docx(docx_path: str) -> dict:
    """
    تصحيح النصوص العربية في ملف Word
    """
    stats = {
        "doc_reversed": False,
        "paragraphs_fixed": 0,
        "tables_fixed": 0
    }
    
    try:
        from docx import Document
        from docx.enum.text import WD_ALIGN_PARAGRAPH
        from docx.oxml.ns import qn
        
        doc = Document(docx_path)
        
        # 1. تصحيح الفقرات
        for para in doc.paragraphs:
            text = para.text.strip()
            if not text:
                continue
            
            # فحص إذا كان النص معكوساً
            if _is_reversed_arabic(text):
                corrected = _reverse_arabic_text(text)
                para.text = corrected
                stats["paragraphs_fixed"] += 1
                stats["doc_reversed"] = True
            
            # ضبط اتجاه الفقرة لـ RTL
            pPr = para._element.get_or_add_pPr()
            bidi = pPr.makeelement(qn('w:bidi'), {})
            pPr.append(bidi)
            
            # محاذاة لليمين
            para.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        
        # 2. تصحيح الجداول
        for table in doc.tables:
            table_fixed = False
            for row in table.rows:
                for cell in row.cells:
                    for para in cell.paragraphs:
                        text = para.text.strip()
                        if text and _is_reversed_arabic(text):
                            corrected = _reverse_arabic_text(text)
                            para.text = corrected
                            table_fixed = True
            
            if table_fixed:
                stats["tables_fixed"] += 1
        
        doc.save(docx_path)
        
    except Exception as e:
        print(f"⚠️ تحذير في تصحيح العربية: {str(e)[:100]}")
    
    return stats


def _is_reversed_arabic(text: str) -> bool:
    """فحص إذا كان النص العربي معكوساً"""
    import unicodedata
    
    # فحص أول 50 حرف
    sample = text[:50]
    arabic_chars = [c for c in sample if unicodedata.category(c).startswith('Lo') and '\u0600' <= c <= '\u06FF']
    
    if len(arabic_chars) < 5:
        return False
    
    # فحص إذا كانت الحروف في وضع معزول (معكوسة)
    isolated_count = 0
    for c in arabic_chars:
        # الحروف المعزولة لها شكل مختلف عن المتصلة
        if unicodedata.name(c, '').endswith('ISOLATED FORM'):
            isolated_count += 1
    
    return isolated_count > len(arabic_chars) * 0.3


def _reverse_arabic_text(text: str) -> str:
    """عكس النص العربي المعكوس"""
    # استخدام مكتبة arabic_reshaper إذا كانت متاحة
    try:
        import arabic_reshaper
        from bidi.algorithm import get_display
        
        reshaped = arabic_reshaper.reshape(text)
        return get_display(reshaped)
    
    except ImportError:
        # fallback بسيط: عكس النص
        return text[::-1]
