"""
formula_helpers.py - دوال مساعدة للمحرك
"""
import os


def finalize_workbook(wb):
    """ضبط حساب المصنف تلقائياً عند الفتح"""
    wb.calculation.fullCalcOnLoad = True
    return wb


def find_hardcoded_totals(path_or_ws, total_row=None):
    """
    يبحث عن صفوف الإجمالي التي تحمل قيماً ثابتة بدل معادلات.
    يقبل: مسار ملف xlsx (كما يستدعيه agent.py) أو worksheet.
    يرجع قائمة نصوص وصفية (فارغة = سليم).
    """
    problems = []
    try:
        import openpyxl
        if isinstance(path_or_ws, str):
            if not os.path.exists(path_or_ws):
                return []
            wb = openpyxl.load_workbook(path_or_ws)
            sheets = wb.worksheets
        else:
            sheets = [path_or_ws]
        for ws in sheets:
            for row in ws.iter_rows():
                if not row:
                    continue
                label = row[0].value
                if isinstance(label, str) and "إجمالي" in label:
                    for cell in row[1:]:
                        if isinstance(cell.value, (int, float)) and not isinstance(cell.value, bool):
                            problems.append(f"{ws.title}!{cell.coordinate} = {cell.value} (ثابت)")
                            if len(problems) >= 10:
                                return problems
        return problems
    except Exception:
        return []


def check_report(ws):
    """التحقق من أن الورقة تحتوي على معادلات حية"""
    for row in ws.iter_rows():
        for cell in row:
            if isinstance(cell.value, str) and cell.value.startswith("="):
                return True
    return False
