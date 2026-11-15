"""
formula_helpers.py - دوال مساعدة للمحرك
"""

def finalize_workbook(wb):
    """ضبط حساب المصنف تلقائياً عند الفتح"""
    wb.calculation.fullCalcOnLoad = True
    return wb

def find_hardcoded_totals(ws, total_row):
    """دالة للتحقق (مستقبلية)"""
    return []

def check_report(ws):
    """دالة للتحقق (مستقبلية)"""
    return True
