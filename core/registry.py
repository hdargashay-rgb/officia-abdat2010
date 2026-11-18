"""
core/registry.py  -  Officia
مصدر الحقيقة الوحيد لدوال النظام + طبقة حماية الـ Router من اختراع الدوال.
"""
import ast
import difflib
import importlib
import inspect
import json
import re

# ===========================================================================
# سجل الدوال: أضف دالة جديدة هنا مرة واحدة فتظهر في التوجيه والموجهات والحارس
# ===========================================================================
REGISTRY = {
    "generate_styled_letter": {
        "module": "core.formatter", "type": "word", "planner": "letter",
        "desc": "خطابات ومذكرات ورسائل رسمية (Word)",
        "keywords": ["خطاب", "مذكرة", "رسالة رسمية", "مخاطبة", "letter"],
    },
    "generate_ats_cv": {
        "module": "core.formatter", "type": "word",
        "desc": "سيرة ذاتية احترافية متوافقة مع أنظمة التوظيف (Word)",
        "keywords": ["سيرة ذاتية", "السيرة الذاتية", "cv", "resume"],
    },
    "generate_financial_report": {
        "module": "core.formatter", "type": "excel",
        "desc": "قائمة الدخل / المركز المالي (الميزانية) / التدفقات النقدية (Excel)",
        "keywords": ["تقرير مالي", "قائمة الدخل", "قائمة دخل", "ميزان عمومية", "ميزانية",
                     "المركز المالي", "التدفقات النقدية", "قائمة التدفقات",
                     "income statement", "balance sheet", "cash flow"],
    },
    "create_fixed_asset_report": {
        "module": "core.formatter", "type": "excel",
        "desc": "تقرير الأصول الثابتة (ملخص + أوراق تفصيلية + جدول إهلاك)",
        "keywords": ["أصول ثابتة", "الأصول الثابتة", "تقرير الأصول", "fixed asset"],
    },
    "generate_depreciation_schedule": {
        "module": "core.formatter", "type": "excel",
        "desc": "جدول إهلاك شهري (رصيد متناقص)",
        "keywords": ["جدول إهلاك", "جدول الإهلاك", "إهلاك شهري", "الإهلاك الشهري", "depreciation"],
    },
    "generate_trial_balance": {
        "module": "core.formatter", "type": "excel",
        "desc": "ميزان مراجعة (أرصدة افتتاحية وحركة وختامية)",
        "keywords": ["ميزان مراجعة", "ميزان المراجعة", "trial balance"],
    },
    "generate_payroll_register": {
        "module": "core.formatter", "type": "excel",
        "desc": "كشف رواتب ومسير موظفين مع قسائم",
        "keywords": ["كشف رواتب", "كشف الرواتب", "مسير رواتب", "مسير الرواتب",
                     "رواتب الموظفين", "payroll"],
    },
    "create_professional_invoice": {
        "module": "core.formatter", "type": "excel",
        "desc": "فاتورة احترافية بضريبة (Excel)",
        "keywords": ["فاتورة", "فاتوره", "invoice"],
    },
    "create_multi_sheet_excel": {
        "module": "core.formatter", "type": "excel", "route": False,
        "desc": "ملف Excel متعدد الأوراق (جداول عامة بمعادلات وإجماليات)",
        "keywords": [],
    },
    "convert_pdf_to_word": {
        "module": "core.pdf_converter", "type": "file", "needs_file": True, "file_ext": ".pdf",
        "desc": "تحويل PDF إلى Word",
        "keywords": ["تحويل pdf", "pdf to word", "pdf الى word", "pdf إلى word",
                     "pdf الى ورد", "pdf إلى ورد"],
    },
    "convert_pdf_to_excel": {
        "module": "core.pdf_converter", "type": "file", "needs_file": True, "file_ext": ".pdf",
        "desc": "استخراج جداول PDF إلى Excel",
        "keywords": ["pdf to excel", "استخراج جداول", "pdf الى excel", "pdf إلى excel",
                     "pdf الى اكسل", "pdf إلى إكسل"],
    },
    "replicate_document": {
        "module": "core.formatter", "type": "file", "needs_file": True,
        "desc": "محاكاة/إعادة بناء ملف Word أو Excel مرفوع",
        "keywords": ["محاكاة ملف", "محاكاة", "إعادة بناء", "replicate"],
    },
    "format_analysis_report": {
        "module": "core.formatter", "type": "excel",
        "desc": "تنسيق تقرير تحليلي (Excel/Word)",
        "keywords": ["تحليل نموذج", "تحليل ملف", "تقرير تحليلي"],
    },
}

# أسماء يخترعها النموذج عادةً → (الدالة الحقيقية، وسائط مسبقة)
ALIASES = {
    "generate_balance_sheet": ("generate_financial_report", {"report_type": "balance_sheet"}),
    "create_balance_sheet": ("generate_financial_report", {"report_type": "balance_sheet"}),
    "generate_income_statement": ("generate_financial_report", {"report_type": "income_statement"}),
    "create_income_statement": ("generate_financial_report", {"report_type": "income_statement"}),
    "generate_cash_flow": ("generate_financial_report", {"report_type": "cash_flow"}),
    "generate_cash_flow_statement": ("generate_financial_report", {"report_type": "cash_flow"}),
    "create_cash_flow_statement": ("generate_financial_report", {"report_type": "cash_flow"}),
    "create_financial_report": ("generate_financial_report", {}),
    "generate_financial_statement": ("generate_financial_report", {}),
    "generate_letter": ("generate_styled_letter", {}),
    "create_letter": ("generate_styled_letter", {}),
    "generate_official_letter": ("generate_styled_letter", {}),
    "create_styled_letter": ("generate_styled_letter", {}),
    "generate_cv": ("generate_ats_cv", {}),
    "create_cv": ("generate_ats_cv", {}),
    "generate_resume": ("generate_ats_cv", {}),
    "generate_invoice": ("create_professional_invoice", {}),
    "create_invoice": ("create_professional_invoice", {}),
    "generate_professional_invoice": ("create_professional_invoice", {}),
    "generate_depreciation": ("generate_depreciation_schedule", {}),
    "create_depreciation_schedule": ("generate_depreciation_schedule", {}),
    "create_trial_balance": ("generate_trial_balance", {}),
    "generate_payroll": ("generate_payroll_register", {}),
    "create_payroll_register": ("generate_payroll_register", {}),
    "create_payroll": ("generate_payroll_register", {}),
    "generate_fixed_asset_report": ("create_fixed_asset_report", {}),
    "generate_fixed_assets_report": ("create_fixed_asset_report", {}),
    "create_fixed_assets_report": ("create_fixed_asset_report", {}),
}

# وحدات الأدوات المسموح استيراد أسمائها من core
ENGINE_MODULES = ["core.table_engine", "core.formula_helpers"]

# محظور في الكود المولَّد
BLOCKED_MODULES = {"subprocess", "socket", "requests", "urllib", "http", "ftplib",
                   "smtplib", "telnetlib", "shutil", "ctypes", "importlib", "pickle",
                   "marshal", "multiprocessing", "webbrowser"}
BLOCKED_CALLS = {"eval", "exec", "compile", "__import__", "os.system", "os.popen",
                 "os.remove", "os.unlink", "os.rmdir", "os.removedirs", "os.rename",
                 "os.kill", "os.execv", "os.spawnl"}


# ===========================================================================
# أدوات عامة
# ===========================================================================
def norm(text) -> str:
    """توحيد النص للمطابقة: تشكيل، همزات، ياء/ألف مقصورة، تاء مربوطة، حروف صغيرة."""
    t = re.sub(r"[\u064B-\u0652\u0640]", "", str(text).lower())
    for a, b in (("أ", "ا"), ("إ", "ا"), ("آ", "ا"), ("ة", "ه"), ("ى", "ي")):
        t = t.replace(a, b)
    return re.sub(r"\s+", " ", t).strip()


def extract_json(text):
    """أول كائن JSON في رد النموذج (يتحمل `json والنص المحيط)."""
    if isinstance(text, dict):
        return text
    s = re.sub(r"^`(?:json)?\s*|\s*```$", "", str(text).strip(), flags=re.S)
    start, depth = s.find("{"), 0
    if start < 0:
        raise ValueError("لم يرجع النموذج JSON")
    in_str, esc = False, False
    for i in range(start, len(s)):
        ch = s[i]
        if in_str:
            if esc:
                esc = False
            elif ch == "\\":
                esc = True
            elif ch == '"':
                in_str = False
            continue
        if ch == '"':
            in_str = True
        elif ch == "{":
            depth += 1
        elif ch == "}":
            depth -= 1
            if depth == 0:
                return json.loads(s[start:i + 1])
    raise ValueError("JSON غير مكتمل في رد النموذج")


def _module(name):
    try:
        return importlib.import_module(name)
    except Exception:
        return None


def get_function(name):
    """الدالة الحقيقية من وحدتها، أو None."""
    meta = REGISTRY.get(name)
    mod = _module(meta["module"]) if meta else None
    return getattr(mod, name, None) if mod else None


def get_signature(name):
    fn = get_function(name)
    try:
        return inspect.signature(fn) if fn else None
    except (TypeError, ValueError):
        return None


def routable():
    return [n for n, m in REGISTRY.items() if m.get("route", True)]


def known_names():
    """كل الأسماء المسموح استيرادها: الدوال المسجلة + أدوات المحرك."""
    names = set(REGISTRY)
    for modname in ENGINE_MODULES:
        mod = _module(modname)
        if mod:
            names |= {n for n, o in vars(mod).items() if callable(o) and not n.startswith("_")}
    return names


# ===========================================================================
# 1) التوجيه
# ===========================================================================
def _kw_hits(text_norm, kw_norm):
    if re.fullmatch(r"[a-z0-9 ]+", kw_norm):
        return re.search(r"(?<![a-z0-9])" + re.escape(kw_norm) + r"(?![a-z0-9])", text_norm) is not None
    return kw_norm in text_norm


def route_keywords(user_request, uploaded_file_path=None):
    """أطول كلمة مفتاحية مطابقة تحدد الدالة (التعادل: الأسبق تسجيلاً)."""
    text = norm(user_request)
    ext = ("." + uploaded_file_path.rsplit(".", 1)[-1].lower()) if uploaded_file_path and "." in uploaded_file_path else None
    best = (0, None)
    for name in routable():
        meta = REGISTRY[name]
        if meta.get("needs_file") and not uploaded_file_path:
            continue
        if meta.get("file_ext") and ext and ext != meta["file_ext"]:
            continue
        for kw in meta["keywords"]:
            k = norm(kw)
            if k and _kw_hits(text, k) and len(k) > best[0]:
                best = (len(k), name)
    return best[1]


def match_function_in_text(text):
    """يبحث عن اسم دالة مسجلة أو رقمها داخل رد المصنّف الذكي."""
    names = routable()
    for n in sorted(names, key=len, reverse=True):
        if n in text:
            return n
    m = re.search(r"\b(\d{1,2})\b", text)
    if m and 1 <= int(m.group(1)) <= len(names):
        return names[int(m.group(1)) - 1]
    return None


def classify_prompt(user_request, uploaded_file_path=None):
    names = routable()
    lines = "\n".join(f"{i}. {n}: {REGISTRY[n]['desc']}" for i, n in enumerate(names, 1))
    file_note = f"\n(يوجد ملف مرفوع: {uploaded_file_path})" if uploaded_file_path else ""
    return (f"أنت مصنف لنظام Officia. اختر الدالة الأنسب لطلب المستخدم من القائمة فقط.\n"
            f"طلب المستخدم: \"{user_request}\"{file_note}\n\nالدوال:\n{lines}\n\n"
            f"أجب باسم الدالة فقط (أو رقمها). إن لم تنطبق أي دالة (يحتاج برمجة مخصصة) اكتب 0. لا شرح.")


# ===========================================================================
# 2) وسائط الدوال (النموذج يملأ JSON، النظام ينفّذ)
# ===========================================================================
def _financial_hints():
    mod = _module("core.formatter")
    tmpl = getattr(mod, "ACCOUNTING_TEMPLATES", None) if mod else None
    if not tmpl:
        return ""
    out = ["قيم report_type: income_statement | balance_sheet | cash_flow (وغيرها تُنتج تقريراً عاماً بورقة واحدة)."]
    for key in ("income_statement", "balance_sheet", "cash_flow"):
        t = tmpl.get(key)
        if t:
            titles = [s["title"] for s in t["sections"] if not s.get("is_formula") and not s.get("is_check")]
            opt = {s["title"] for s in t["sections"] if s.get("optional")}
            out.append(f"عناوين أقسام {key} المقبولة (استخدمها حرفياً في 'title'؛ * اختياري): " +
                       "، ".join(f"{x}{'*' if x in opt else ''}" for x in titles))
    out.append("كل قسم: {'title','headers':['البيان','المبلغ'],'data':[['بند',قيمة],...]}. "
               "المصروفات والتكاليف تُدخل بقيم موجبة (القالب يطرحها). لا تضف صفوف إجمالي: تُحسب بمعادلات.")
    return "\n".join(out)


def function_args_prompt(name, user_request, uploaded_file_path=None):
    fn = get_function(name)
    sig = get_signature(name)
    doc = (inspect.getdoc(fn) or "")[:1800] if fn else REGISTRY[name]["desc"]
    params = []
    if sig:
        for p in sig.parameters.values():
            if p.kind in (p.VAR_KEYWORD, p.VAR_POSITIONAL):
                continue
            req = "إلزامي" if p.default is inspect._empty else f"اختياري (افتراضي {p.default!r})"
            params.append(f"  - {p.name}: {req}")
    extra = _financial_hints() if name == "generate_financial_report" else ""
    file_line = f"\nالملف المرفوع: {uploaded_file_path}" if uploaded_file_path else ""
    return f"""مهمتك تعبئة وسائط الدالة {name} من طلب المستخدم. أعد JSON فقط بالشكل {{ "args": {{...}}}}.
طلب المستخدم: {user_request}{file_line}
وصف الدالة:
{doc}
الوسائط المسموحة فقط (لا تخترع وسائط غير مذكورة):
{chr(10).join(params)}
{extra}
قواعد:
لا تخترع بيانات لم يذكرها المستخدم؛ ما ينقص اترك له قيمته الافتراضية أو ضع نصاً محايداً.
الأرقام أرقام (بلا فواصل أو رموز عملة)، والقوائم قوائم JSON.
لا تضع output_dir (يحدده النظام).
لا تكتب أي نص خارج JSON."""


def validate_call_args(name, args):
    """يرجع قائمة مشكلات (فارغة = صالح): وسائط غير مدعومة، ناقصة، أو بنوع خاطئ."""
    sig = get_signature(name)
    if sig is None:
        return [f"تعذر قراءة توقيع {name}"]
    issues = []
    params = {p.name: p for p in sig.parameters.values() if p.kind not in (p.VAR_KEYWORD, p.VAR_POSITIONAL)}
    for k in args:
        if k not in params:
            close = difflib.get_close_matches(k, params, n=1)
            issues.append(f"الوسيط '{k}' غير مدعوم في {name}" + (f" (هل تقصد '{close[0]}'؟)" if close else ""))
    for pn, p in params.items():
        if p.default is inspect._empty and pn not in args:
            issues.append(f"الوسيط الإلزامي '{pn}' ناقص")
    types = {"list": list, "dict": dict, "str": str, "int": int, "float": (int, float), "bool": bool}
    for k, v in args.items():
        ann = getattr(params.get(k), "annotation", inspect._empty)
        tname = getattr(ann, "__name__", str(ann))
        if v is not None and tname in types and not isinstance(v, types[tname]):
            issues.append(f"الوسيط '{k}' يجب أن يكون {tname} وليس {type(v).__name__}")
    return issues


# ===========================================================================
# قائمة مسموحة مضغوطة للموجهات (مولّدة من السجل والتوقيعات)
# ===========================================================================
def allowed_functions_prompt():
    lines = ["⛔ الدوال المسموحة حصراً من core (ممنوع استيراد أو استدعاء أي دالة أخرى منها):"]
    for name, meta in REGISTRY.items():
        sig = get_signature(name)
        params = ", ".join(p.name for p in sig.parameters.values()
                          if p.kind not in (p.VAR_KEYWORD, p.VAR_POSITIONAL)) if sig else "..."
        lines.append(f"  • {name}({params}) — {meta['desc']}")
    lines.append("أدوات الجداول (core.table_engine / core.formula_helpers): col, resolve_columns, write_table, "
                 "apply_widths, write_total_row, write_derived_column, write_balance_check, finalize_workbook.")
    lines.append("أي دالة ليست هنا غير موجودة (مثل generate_balance_sheet) وستُرفض قبل التنفيذ.")
    return "\n".join(lines)


# ===========================================================================
# 3) حارس الكود المولَّد (AST)
# ===========================================================================
class _AliasHealer(ast.NodeTransformer):
    """يستبدل الأسماء المخترعة المعروفة بالدالة الحقيقية مع وسائطها المسبقة."""
    def __init__(self):
        self.fixes = []

    def visit_ImportFrom(self, node):
        if node.module and node.module.startswith("core"):
            new_names, seen = [], set()
            for a in node.names:
                real = ALIASES.get(a.name, (a.name, {}))[0]
                if real != a.name:
                    self.fixes.append(f"استيراد {a.name} → {real}")
                if real not in seen:
                    seen.add(real)
                    new_names.append(ast.alias(name=real, asname=a.asname if real == a.name else None))
            node.names = new_names
        return node

    def visit_Call(self, node):
        self.generic_visit(node)
        name = node.func.id if isinstance(node.func, ast.Name) else (
            node.func.attr if isinstance(node.func, ast.Attribute) else None)
        if name in ALIASES:
            real, preset = ALIASES[name]
            if isinstance(node.func, ast.Name):
                node.func = ast.Name(id=real, ctx=ast.Load())
            else:
                node.func.attr = real
            have = {k.arg for k in node.keywords}
            for k, v in preset.items():
                if k not in have:
                    node.keywords.append(ast.keyword(arg=k, value=ast.Constant(value=v)))
            self.fixes.append(f"استدعاء {name}(...) → {real}(...)" + (f" مع {preset}" if preset else ""))
        return node


def _dotted(node):
    parts = []
    while isinstance(node, ast.Attribute):
        parts.append(node.attr)
        node = node.value
    if isinstance(node, ast.Name):
        parts.append(node.id)
    return ".".join(reversed(parts))


def check_and_fix_code(code: str) -> dict:
    """
    يرجع {'code': الكود بعد الإصلاح التلقائي، 'fixes': [...], 'issues': [...], 'ok': bool}.
    issues = مشكلات تمنع التنفيذ ويجب إعادتها للنموذج.
    """
    result = {"code": code, "fixes": [], "issues": [], "ok": False}
    try:
        tree = ast.parse(code)
    except SyntaxError as e:
        result["issues"].append(f"خطأ نحوي في السطر {e.lineno}: {e.msg}")
        return result

    healer = _AliasHealer()
    tree = healer.visit(tree)
    ast.fix_missing_locations(tree)
    if healer.fixes:
        try:
            result["code"] = ast.unparse(tree)
            result["fixes"] = healer.fixes
        except AttributeError:
            result["issues"].append("أسماء دوال مخترعة: " + "؛ ".join(healer.fixes))

    issues = result["issues"]
    allowed = known_names()
    all_core_public = set()
    for modname in ("core.formatter", "core.pdf_converter"):
        mod = _module(modname)
        if mod:
            all_core_public |= {n for n, o in vars(mod).items() if callable(o) and not n.startswith("_")}

    for node in ast.walk(tree):
        if isinstance(node, ast.Import):
            for a in node.names:
                if a.name.split(".")[0] in BLOCKED_MODULES:
                    issues.append(f"استيراد محظور: {a.name}")
        elif isinstance(node, ast.ImportFrom):
            top = (node.module or "").split(".")[0]
            if top in BLOCKED_MODULES:
                issues.append(f"استيراد محظور: {node.module}")
            if node.module in ("core.formatter", "core.pdf_converter") or (node.module or "").startswith("core."):
                for a in node.names:
                    if a.name == "*":
                        issues.append(f"ممنوع 'import *' من {node.module}")
                    elif a.name not in allowed and a.name not in all_core_public:
                        close = difflib.get_close_matches(a.name, sorted(allowed), n=2)
                        issues.append(f"الدالة '{a.name}' غير موجودة في {node.module}"
                                      + (f". أقرب الدوال: {', '.join(close)}" if close else ""))
        elif isinstance(node, ast.Call):
            dotted = _dotted(node.func)
            if dotted in BLOCKED_CALLS or dotted.split(".")[0] == "subprocess":
                issues.append(f"استدعاء محظور: {dotted}()")
            fname = node.func.id if isinstance(node.func, ast.Name) else None
            if fname in REGISTRY:
                sig = get_signature(fname)
                if sig is None or any(isinstance(a, ast.Starred) for a in node.args) \
                        or any(k.arg is None for k in node.keywords):
                    continue
                params = {p.name: p for p in sig.parameters.values()
                          if p.kind not in (p.VAR_KEYWORD, p.VAR_POSITIONAL)}
                pos_names = list(params)
                given = set(pos_names[:len(node.args)])
                for k in node.keywords:
                    if k.arg not in params:
                        close = difflib.get_close_matches(k.arg, params, n=1)
                        issues.append(f"{fname}(): الوسيط '{k.arg}' غير مدعوم"
                                      + (f" (هل تقصد '{close[0]}'؟)" if close else ""))
                    given.add(k.arg)
                missing = [n for n, p in params.items() if p.default is inspect._empty and n not in given]
                if missing:
                    issues.append(f"{fname}(): وسائط إلزامية ناقصة: {', '.join(missing)}")

    result["issues"] = list(dict.fromkeys(issues))
    result["ok"] = not result["issues"]
    return result
