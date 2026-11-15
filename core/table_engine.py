"""
table_engine.py  -  Officia
محرك الجداول: الأعمدة تُعرَّف بمواصفات (Spec) فتصبح إضافتها وحذفها وتغيير معادلاتها
مسألة بيانات لا مسألة كود، وكل عمود مشتق وكل إجمالي يُكتب معادلة Excel حيّة.

الفكرة:
    columns = [col("basic", "الراتب الأساسي", group="income", total="sum"),
               col("gross", "الإجمالي", kind="formula", formula="=SUM({group:income})", total="sum")]
    cols, warns = resolve_columns(columns, add_columns=[...], remove_columns=[...])
    info = write_table(ws, cols, rows, header_row=3)

رموز المعادلات داخل القوالب:
    {key}          خلية العمود key في نفس الصف          ->  E5
    {group:NAME}   كل أعمدة المجموعة في نفس الصف         ->  E5,F5,G5   (أو  0 إن كانت فارغة)
    {key:col}      نطاق العمود كاملاً (مطلق)             ->  $E$4:$E$20
    {key:tot}      خلية الإجمالي لذلك العمود (مطلقة)     ->  $E$21
عمود محذوف مذكور في معادلة يُستبدل بـ 0 مع تحذير، فلا ينكسر النموذج.
"""

import re
import copy
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

from formula_helpers import finalize_workbook, find_hardcoded_totals, check_report  # noqa: F401

_TOKEN = re.compile(r"\{([^{}]+)\}")

NUM_FMT = "#,##0.00"
_FMT_BY_KIND = {"number": NUM_FMT, "formula": NUM_FMT, "percent": "0.0%", "points": '0.00"%"'}

_THIN = Side(style="thin", color="B4C7E7")
DEFAULT_STYLES = {
    "header_fill": PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid"),
    "header_font": Font(name="Arial", size=11, bold=True, color="FFFFFF"),
    "header_align": Alignment(horizontal="center", vertical="center", wrap_text=True),
    "header_border": Border(left=_THIN, right=_THIN, top=_THIN, bottom=_THIN),
    "data_font": Font(name="Arial", size=10),
    "data_border": Border(left=_THIN, right=_THIN, top=_THIN, bottom=_THIN),
    "total_fill": PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid"),
    "total_font": Font(name="Arial", size=11, bold=True, color="1F4E79"),
    "total_border": Border(left=_THIN, right=_THIN, top=_THIN,
                           bottom=Side(style="double", color="1F4E79")),
}


# ---------------------------------------------------------------------------
# مواصفات الأعمدة
# ---------------------------------------------------------------------------
def col(key, title, kind="number", **kw):
    """
    kind : text | number | percent (كسر 0.15) | points (نقاط مئوية 15) | index (ترقيم) | formula
    kw   : formula, total ('sum'|'avg'|'count'|'min'|'max'|'same'|None), group, width,
           fmt, align ('center'|'left'|'right'), bold (True/False)
    """
    d = {"key": key, "title": title, "kind": kind}
    d.update(kw)
    if "formula" in d:
        d["kind"] = "formula"
    return d


def _as_list(x):
    if x is None:
        return []
    if isinstance(x, (str, dict)):
        return [x]
    return list(x)


def resolve_columns(base, add_columns=None, remove_columns=None, rename_columns=None):
    """
    يطبّق الإضافة والحذف وإعادة التسمية على مواصفات الأعمدة الافتراضية.
    يرجع (قائمة الأعمدة، قائمة التحذيرات).

    add_columns : قائمة dict (أو نصوص = أعمدة أرقام). مفاتيح اختيارية: after / before (مفتاح أو عنوان)،
                  وإن حُدِّد group دخل العمود بعد آخر عمود من المجموعة نفسها.
    remove_columns : مفاتيح أو عناوين.
    rename_columns : {مفتاح_أو_عنوان: عنوان_جديد}
    """
    cols = [dict(c) for c in base]
    warns = []

    def find(ref):
        for i, c in enumerate(cols):
            if c["key"] == ref or c["title"] == ref:
                return i
        return -1

    for ref in _as_list(remove_columns):
        i = find(ref)
        if i < 0:
            warns.append(f"عمود الحذف غير موجود: {ref}")
        else:
            cols.pop(i)

    for ref, new_title in (rename_columns or {}).items():
        i = find(ref)
        if i < 0:
            warns.append(f"عمود إعادة التسمية غير موجود: {ref}")
        else:
            cols[i]["title"] = new_title

    for spec in _as_list(add_columns):
        if isinstance(spec, str):
            spec = {"key": spec, "title": spec}
        spec = dict(spec)
        spec.setdefault("title", spec.get("key", "عمود"))
        spec.setdefault("key", spec["title"])
        if "formula" in spec:
            spec["kind"] = "formula"
        spec.setdefault("kind", "number")
        if "total" not in spec:   # الافتراضي: الأرقام تُجمع، والمعادلات تُحسب على صف الإجمالي بنفس القالب
            spec["total"] = {"number": "sum", "formula": "same"}.get(spec["kind"])
        if find(spec["key"]) >= 0:
            warns.append(f"العمود موجود مسبقاً ولم يُضَف: {spec['key']}")
            continue
        pos = len(cols)
        after, before = spec.pop("after", None), spec.pop("before", None)
        if after is not None and find(after) >= 0:
            pos = find(after) + 1
        elif before is not None and find(before) >= 0:
            pos = find(before)
        elif spec.get("group"):
            idx = [i for i, c in enumerate(cols) if c.get("group") == spec["group"]]
            if idx:
                pos = max(idx) + 1
        cols.insert(pos, spec)

    return cols, warns


# ---------------------------------------------------------------------------
# المعادلات
# ---------------------------------------------------------------------------
def compile_formula(template, letters, groups, row, first, last, total_row=None, warns=None):
    """يحوّل قالب المعادلة إلى معادلة Excel فعلية للصف المعطى."""

    def sub(m):
        tok = m.group(1).strip()
        if tok.startswith("group:"):
            refs = [f"{letters[k]}{row}" for k in groups.get(tok[6:], [])]
            return ",".join(refs) if refs else "0"
        key, mode = (tok.rsplit(":", 1) + [None])[:2] if ":" in tok else (tok, None)
        L = letters.get(key)
        if L is None:
            if warns is not None:
                warns.append(f"المعادلة تشير إلى عمود غير موجود ({key}) فاستُبدل بصفر")
            return "0"
        if mode == "col":
            return f"${L}${first}:${L}${last}"
        if mode == "tot":
            return f"${L}${total_row}" if total_row else "0"
        return f"{L}{row}"

    return _TOKEN.sub(sub, template)


def _num(v):
    """يحوّل '1,200' أو '1200.5' إلى رقم، ويترك غير الأرقام كما هي."""
    if isinstance(v, bool) or v is None:
        return 0 if v is None else v
    if isinstance(v, (int, float)):
        return v
    if isinstance(v, str):
        s = v.replace(",", "").replace("٬", "").strip()
        try:
            return float(s) if s not in ("", "-") else 0
        except ValueError:
            return v
    return v


def sheet_ref(name):
    """اسم ورقة مقتبس للاستخدام داخل المعادلات."""
    return "'" + str(name).replace("'", "''") + "'"


# ---------------------------------------------------------------------------
# كتابة الجدول
# ---------------------------------------------------------------------------
def _rows_to_dicts(rows, columns):
    data_cols = [c["key"] for c in columns if c["kind"] not in ("formula", "index")]
    out = []
    for r in rows:
        if isinstance(r, dict):
            out.append(r)
        elif isinstance(r, (list, tuple)):
            out.append(dict(zip(data_cols, r)))
        else:
            out.append({})
    return out


def write_table(ws, columns, rows, header_row, first_col=1, styles=None, total=True,
                total_label="الإجمالي", label_span=1, header_height=None, row_height=None):
    """
    يكتب رأس الجدول وصفوفه وصف الإجمالي (بمعادلات حية) ويرجع معلومات الموضع:
        {'first_row','last_row','total_row','letters','warnings','last_col'}
    """
    st = {**DEFAULT_STYLES, **(styles or {})}
    letters = {c["key"]: get_column_letter(first_col + i) for i, c in enumerate(columns)}
    groups = {}
    for c in columns:
        if c.get("group"):
            groups.setdefault(c["group"], []).append(c["key"])

    rows = _rows_to_dicts(rows, columns)
    first = header_row + 1
    last = header_row + len(rows)
    total_row = last + 1 if total else None
    warns = []

    # --- الرأس
    for i, c in enumerate(columns):
        cell = ws.cell(row=header_row, column=first_col + i, value=c["title"])
        cell.fill, cell.font = st["header_fill"], st["header_font"]
        cell.alignment, cell.border = st["header_align"], st["header_border"]
    if header_height:
        ws.row_dimensions[header_row].height = header_height

    # --- الصفوف
    for n, row in enumerate(rows):
        r = first + n
        for i, c in enumerate(columns):
            cell = ws.cell(row=r, column=first_col + i)
            kind = c["kind"]
            if kind == "index":
                cell.value = n + 1
            elif kind == "formula":
                cell.value = compile_formula(c["formula"], letters, groups, r, first, last,
                                             total_row, warns)
            elif kind == "text":
                v = row.get(c["key"])
                cell.value = "" if v is None else v
            else:
                cell.value = _num(row.get(c["key"]))
            cell.font = c.get("font") or (Font(name="Arial", size=10, bold=True) if c.get("bold")
                                          else st["data_font"])
            default_align = "right" if kind in ("number", "formula", "percent", "points") else "center"
            cell.alignment = Alignment(horizontal=c.get("align", default_align),
                                       vertical="center", wrap_text=True)
            cell.border = st["data_border"]
            if kind in _FMT_BY_KIND:
                cell.number_format = c.get("fmt", _FMT_BY_KIND[kind])
        if row_height:
            ws.row_dimensions[r].height = row_height

    # --- صف الإجمالي
    if total:
        for i, c in enumerate(columns):
            cell = ws.cell(row=total_row, column=first_col + i)
            L = letters[c["key"]]
            t = c.get("total")
            if t in ("sum", "avg", "count", "min", "max"):
                fn = {"sum": "SUM", "avg": "AVERAGE", "count": "COUNT", "min": "MIN", "max": "MAX"}[t]
                cell.value = f"={fn}({L}{first}:{L}{last})" if last >= first else 0
                cell.number_format = c.get("fmt", _FMT_BY_KIND.get(c["kind"], NUM_FMT))
            elif t == "same" and c["kind"] == "formula":
                cell.value = compile_formula(c["formula"], letters, groups, total_row, first, last,
                                             total_row, warns)
                cell.number_format = c.get("fmt", NUM_FMT)
            cell.fill, cell.font = st["total_fill"], st["total_font"]
            cell.border = st["total_border"]
            cell.alignment = Alignment(horizontal="right" if cell.value is not None else "center",
                                       vertical="center")
        label_cell = ws.cell(row=total_row, column=first_col, value=total_label)
        label_cell.alignment = Alignment(horizontal="right", vertical="center")
        if label_span > 1:
            ws.merge_cells(start_row=total_row, start_column=first_col,
                           end_row=total_row, end_column=first_col + label_span - 1)
        if row_height:
            ws.row_dimensions[total_row].height = row_height + 5

    return {"first_row": first, "last_row": last, "total_row": total_row, "letters": letters,
            "warnings": list(dict.fromkeys(warns)), "last_col": first_col + len(columns) - 1}


def apply_widths(ws, columns, first_col=1, default=16):
    for i, c in enumerate(columns):
        ws.column_dimensions[get_column_letter(first_col + i)].width = c.get("width", default)


def warn_ignored(func_name, kwargs, handled=()):
    """يطبع تنبيهاً بالمعاملات التي لم تفهمها الدالة، بدل تجاهلها بصمت."""
    ignored = [k for k in kwargs if k not in handled]
    if ignored:
        print(f"⚠️ {func_name}: معاملات لم تُنفَّذ (غير مدعومة): {', '.join(ignored)}")
    return ignored


def print_warnings(func_name, warns):
    for w in warns:
        print(f"⚠️ {func_name}: {w}")


# ---------------------------------------------------------------------------
# أدوات تعمل بعناوين الأعمدة (للجداول العامة غير المعرّفة بمواصفات)
# ---------------------------------------------------------------------------
_SKIP_TOTAL_WORDS = ("رقم", "م", "#", "نسبة", "%", "تاريخ", "الكود", "code", "id", "no")


def _headers_map(ws, header_row):
    return {str(c.value).strip(): c.column for c in ws[header_row] if c.value is not None}


def apply_header_formulas(ws, header_row, first_row, last_row, formulas, number_format=NUM_FMT):
    """
    formulas = {"الصافي": "={التكلفة}-{الإهلاك}"}  (الرموز بعناوين الأعمدة)
    يكتب المعادلة في كل صف، وإن لم يوجد عمود بهذا العنوان يُنشئه في آخر الجدول.
    يرجع ورقة العناوين المحدّثة (dict).
    """
    headers = _headers_map(ws, header_row)
    for target in formulas:
        if target not in headers:
            new_col = max(headers.values()) + 1 if headers else 1
            src = ws.cell(row=header_row, column=new_col - 1)
            h = ws.cell(row=header_row, column=new_col, value=target)
            h.fill, h.font = copy.copy(src.fill), copy.copy(src.font)
            h.alignment, h.border = copy.copy(src.alignment), copy.copy(src.border)
            headers[target] = new_col

    for target, template in formulas.items():
        tcol = headers[target]
        for r in range(first_row, last_row + 1):
            def sub(m, r=r):
                name = m.group(1).strip()
                if name not in headers:
                    raise ValueError(
                        f"العمود '{name}' غير موجود. الأعمدة المتاحة: {', '.join(headers)}")
                return f"{get_column_letter(headers[name])}{r}"
            cell = ws.cell(row=r, column=tcol, value=_TOKEN.sub(sub, template))
            cell.number_format = number_format
    return headers


def add_totals_by_header(ws, header_row, first_row, last_row, totals="auto",
                         label="الإجمالي", fill=None, font=None, border=None):
    """
    totals: "auto" (كل عمود رقمي) أو قائمة عناوين.
    يضيف صف إجمالي بمعادلات SUM تحت البيانات ويرجع رقمه (أو None إن لم يجد ما يُجمع).
    """
    headers = _headers_map(ws, header_row)
    if totals in (True, "auto"):
        targets = []
        for h, c in headers.items():
            low = h.lower()
            if any(low == w or (len(w) > 2 and w in low) for w in _SKIP_TOTAL_WORDS):
                continue
            vals = [ws.cell(row=r, column=c).value for r in range(first_row, last_row + 1)]
            numeric = [v for v in vals if isinstance(v, (int, float)) and not isinstance(v, bool)
                       or (isinstance(v, str) and v.startswith("="))]
            if numeric and len(numeric) >= max(1, len([v for v in vals if v not in (None, "")]) // 2):
                targets.append(h)
    else:
        targets = [t for t in _as_list(totals)]
        missing = [t for t in targets if t not in headers]
        if missing:
            raise ValueError(f"أعمدة الإجمالي غير موجودة: {', '.join(missing)}. المتاحة: {', '.join(headers)}")
    if not targets or last_row < first_row:
        return None

    total_row = last_row + 1
    fill = fill or DEFAULT_STYLES["total_fill"]
    font = font or DEFAULT_STYLES["total_font"]
    border = border or DEFAULT_STYLES["total_border"]
    first_c, last_c = min(headers.values()), max(headers.values())
    for c in range(first_c, last_c + 1):
        cell = ws.cell(row=total_row, column=c)
        cell.fill, cell.font, cell.border = fill, font, border
    ws.cell(row=total_row, column=first_c, value=label).alignment = Alignment(horizontal="right", vertical="center")
    for h in targets:
        L = get_column_letter(headers[h])
        cell = ws.cell(row=total_row, column=headers[h], value=f"=SUM({L}{first_row}:{L}{last_row})")
        cell.number_format = NUM_FMT
        cell.alignment = Alignment(horizontal="right", vertical="center")
    return total_row
