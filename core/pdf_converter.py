"""
PDF Converter - محول PDF إلى Word و Excel
يدعم اللغة العربية مع تصحيح تلقائي
"""
import os
import warnings
from core.config import WORKSPACE_DIR


# ==========================================
# 🔍 دوال الفحص المساعدة
# ==========================================

def _inspect_pdf(pdf_path: str) -> dict:
    """فحص ملف PDF وإرجاع معلوماته"""
    try:
        import fitz  # PyMuPDF
        
        if not os.path.exists(pdf_path):
            return {"ok": False, "error": "الملف غير موجود"}
        
        if not pdf_path.lower().endswith('.pdf'):
            return {"ok": False, "error": "الملف ليس بصيغة PDF"}
        
        file_size_mb = os.path.getsize(pdf_path) / (1024 * 1024)
        if file_size_mb > 50:
            return {"ok": False, "error": "حجم الملف يتجاوز 50 ميجابايت"}
        
        doc = fitz.open(pdf_path)
        pages = len(doc)
        
        # فحص الصفحات الفارغة
        empty_pages = []
        for i in range(min(pages, 10)):  # فحص أول 10 صفحات فقط
            page = doc[i]
            text = page.get_text().strip()
            if not text:
                empty_pages.append(i + 1)
        
        doc.close()
        
        return {
            "ok": True,
            "pages": pages,
            "size_mb": round(file_size_mb, 2),
            "empty_pages": empty_pages
        }
    
    except Exception as e:
        return {"ok": False, "error": str(e)}


def _docx_text_length(docx_path: str) -> int:
    """قياس طول النص في ملف Word"""
    try:
        from docx import Document
        doc = Document(docx_path)
        text = " ".join(p.text for p in doc.paragraphs)
        return len(text)
    except:
        return 0


# ==========================================
# 🔧 دوال التحويل
# ==========================================

def convert_pdf_to_word(
    pdf_path: str,
    output_path: str = None,
    output_dir: str = None,
    max_pages: int = 50,
    **kwargs
) -> str:
    """
    تحويل PDF إلى Word مع الحفاظ على التنسيق
    يدعم اللغة العربية مع تصحيح تلقائي
    """
    # فحص الملف
    info = _inspect_pdf(pdf_path)
    if not info["ok"]:
        raise ValueError(f"❌ {info['error']}")
    
    # تحديد مسار الإخراج
    save_dir = output_dir or WORKSPACE_DIR
    os.makedirs(save_dir, exist_ok=True)
    
    if output_path is None:
        base_name = os.path.splitext(os.path.basename(pdf_path))[0]
        output_path = os.path.join(save_dir, f"{base_name}_converted.docx")
    
    warnings_list = []
    pages = info["pages"]
    end_page = min(pages, max_pages)
    
    if pages > max_pages:
        warnings_list.append(f"تم تحويل أول {max_pages} صفحة فقط من أصل {pages}")
    
    # المحاولة الأولى: pdf2docx (يحافظ على التنسيق)
    method = "pdf2docx"
    success = False
    
    try:
        from pdf2docx import Converter
        
        cv = Converter(pdf_path)
        cv.convert(output_path, start=0, end=end_page)
        cv.close()
        
        # التحقق من نجاح التحويل
        if os.path.exists(output_path) and _docx_text_length(output_path) >= 20:
            success = True
        else:
            warnings_list.append("التحويل الأولي أنتج نصاً قليلاً، جاري المحاولة البديلة")
    
    except Exception as e:
        warnings_list.append(f"فشل التحويل الأولي: {str(e)[:100]}")
    
    # المحاولة الثانية: استخراج نص بسيط
    if not success:
        method = "text_fallback"
        try:
            import fitz
            from docx import Document
            from docx.shared import Pt
            from docx.enum.text import WD_ALIGN_PARAGRAPH
            
            doc = fitz.open(pdf_path)
            word_doc = Document()
            
            for page_num in range(min(doc.page_count, end_page)):
                page = doc[page_num]
                text = page.get_text()
                
                if text.strip():
                    para = word_doc.add_paragraph(text)
                    para.alignment = WD_ALIGN_PARAGRAPH.RIGHT  # RTL للعربية
                    para.paragraph_format.space_after = Pt(12)
            
            doc.close()
            word_doc.save(output_path)
            
            if _docx_text_length(output_path) < 20:
                raise ValueError("تعذّر استخراج نص كافٍ من الملف")
            
            warnings_list.append("استُخدم تحويل مبسّط: النص محفوظ لكن التخطيط قد يتغير")
            success = True
        
        except Exception as e:
            raise ValueError(f"تعذّر تحويل الملف: {str(e)}")
    
    # تصحيح العربية
    try:
        from core.arabic_fixer import fix_arabic_docx
        stats = fix_arabic_docx(output_path)
        if stats.get("doc_reversed", False):
            warnings_list.append("تم تصحيح نص معكوس، راجع الفقرات المهمة")
    except Exception as e:
        warnings_list.append("تعذّر تصحيح اتجاه النص العربي تلقائياً")
    
    # فحوص الجودة
    try:
        from docx import Document
        doc = Document(output_path)
        all_text = " ".join(p.text for p in doc.paragraphs)
        
        # فحص الحروف المشوهة
        junk_chars = sum(1 for c in all_text if '\ue000' <= c <= '\uf8ff' or c == '\ufffd')
        if all_text and junk_chars / max(len(all_text), 1) > 0.03:
            warnings_list.append("بعض الحروف لم تُقرأ بسبب خطوط غير قياسية")
        
        doc.close()
    except:
        pass
    
    # طباعة التحذيرات
    if warnings_list:
        print("⚠️ تحذيرات:")
        for w in warnings_list:
            print(f"  - {w}")
    
    print("SUCCESS: " + output_path)
    return output_path


def convert_pdf_to_excel(
    pdf_path: str,
    output_path: str = None,
    output_dir: str = None,
    max_pages: int = 50,
    **kwargs
) -> str:
    """
    استخراج الجداول من PDF إلى Excel
    مثالي للتقارير المالية والكشوفات
    """
    # فحص الملف
    info = _inspect_pdf(pdf_path)
    if not info["ok"]:
        raise ValueError(f"❌ {info['error']}")
    
    # تحديد مسار الإخراج
    save_dir = output_dir or WORKSPACE_DIR
    os.makedirs(save_dir, exist_ok=True)
    
    if output_path is None:
        base_name = os.path.splitext(os.path.basename(pdf_path))[0]
        output_path = os.path.join(save_dir, f"{base_name}_tables.xlsx")
    
    warnings_list = []
    
    try:
        import pdfplumber
        import openpyxl
        from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
        from openpyxl.utils import get_column_letter
        
        wb = openpyxl.Workbook()
        wb.remove(wb.active)  # حذف الورقة الافتراضية
        
        pdf = pdfplumber.open(pdf_path)
        tables_found = 0
        
        for page_num, page in enumerate(pdf.pages[:max_pages]):
            tables = page.extract_tables()
            
            if not tables:
                continue
            
            for table_idx, table in enumerate(tables):
                if not table or len(table) < 2:
                    continue
                
                tables_found += 1
                sheet_name = f"صفحة_{page_num+1}_جدول_{table_idx+1}"[:31]  # حد 31 حرف
                ws = wb.create_sheet(title=sheet_name)
                ws.sheet_view.rightToLeft = True
                
                # الأنماط
                header_fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
                header_font = Font(name="Arial", size=11, bold=True, color="FFFFFF")
                data_font = Font(name="Arial", size=10)
                center = Alignment(horizontal="center", vertical="center", wrap_text=True)
                thin_border = Border(
                    left=Side(style="thin"),
                    right=Side(style="thin"),
                    top=Side(style="thin"),
                    bottom=Side(style="thin")
                )
                
                # كتابة البيانات
                for row_idx, row in enumerate(table):
                    for col_idx, cell_value in enumerate(row):
                        cell = ws.cell(row=row_idx+1, column=col_idx+1, value=cell_value)
                        
                        if row_idx == 0:  # صف الرؤوس
                            cell.fill = header_fill
                            cell.font = header_font
                            cell.alignment = center
                            cell.border = thin_border
                        else:
                            cell.font = data_font
                            cell.alignment = Alignment(horizontal="right", vertical="center")
                            cell.border = thin_border
                
                # ضبط عرض الأعمدة
                for col in range(1, len(table[0]) + 1):
                    ws.column_dimensions[get_column_letter(col)].width = 15
                
                # تجميد الصف الأول
                ws.freeze_panes = 'A2'
        
        pdf.close()
        
        if tables_found == 0:
            raise ValueError("لم يتم العثور على جداول في الملف")
        
        wb.save(output_path)
        
        print(f"✅ تم استخراج {tables_found} جدول من {min(len(pdf.pages), max_pages)} صفحة")
        print("SUCCESS: " + output_path)
        return output_path
    
    except ImportError:
        raise ValueError("مكتبة pdfplumber غير مثبتة. قم بتثبيتها: pip install pdfplumber")
    except Exception as e:
        raise ValueError(f"فشل استخراج الجداول: {str(e)}")
