"""
Agent Loop - الحلقة الذكية للوكيل (الإصدار 3)

الفكرة: لا نترك النموذج يكتب كوداً إلا عند الضرورة.
  • طلب يطابق دالة مسجلة  → وضع الدالة: النموذج يملأ وسائط JSON فقط، والنظام يتحقق منها (توقيع الدالة) ثم ينفذ.
                              (الخطابات تمر أولاً بـ letter_planner لصياغة خطاب ضافٍ كامل.)
  • طلب مخصص              → وضع الكود: النموذج يكتب كوداً، وحارس AST (core.registry) يصلح الأسماء المخترعة
                              ويرفض الدوال/الوسائط غير الموجودة والاستيرادات الخطرة قبل أي تنفيذ.
  • بعد كل نجاح           → تحقق: هل أُنشئ ملف؟ هل ظهرت تحذيرات ⚠️؟ هل في Excel إجماليات ثابتة؟ جولة إصلاح واحدة.
مصدر الحقيقة للدوال: core/registry.py (أضف دالة جديدة هناك فتظهر في التوجيه والموجهات والحارس).
"""
import ast
import json
import os
import re

from core.llm_client import LLMClient
from core.executor import Executor
from core.config import MAX_AUTO_FIX_ATTEMPTS, WORKSPACE_DIR
from core import registry as R

# ==========================================================
# 🎨 بروتوكول التنسيق (يُستخدم في وضع الكود المخصص فقط، فالدوال الجاهزة تنسّق نفسها)
# ==========================================================
STYLE_PROTOCOL = """
🎨🎨🎨 بروتوكول التنسيق الاحترافي الإلزامي - يجب تطبيقه في كل ملف 🎨🎨🎨

📊 للملفات Excel (openpyxl):
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

1️⃣ الألوان الاحترافية:
   • الأزرق الغامق 1F4E79: للعناوين الرئيسية ورؤوس الأعمدة وصف الإجمالي النهائي
   • الأزرق المتوسط 2E5C8A: للعناوين الفرعية
   • الأزرق الفاتح D6E4F0: للمجاميع الفرعية
   • الأزرق الفاتح جداً E8F0F7: للتلوين المتبادل للبيانات
   • الأبيض FFFFFF: لصفوف البيانات العادية
   • الذهبي C9A961: للخطوط الفاصلة والتمييز

2️⃣ الخطوط:
   • العناوين: Arial, 16-18pt, Bold, لون 1F4E79
   • رؤوس الأعمدة: Arial, 12pt, Bold, أبيض على خلفية 1F4E79
   • البيانات: Arial, 10-11pt, عادي
   • الإجماليات: Arial, 12pt, Bold, أبيض على خلفية 1F4E79

3️⃣ أبعاد الأعمدة:
   • عمود النصوص (البيان/الوصف): عرض 45
   • أعمدة الأرقام: عرض 20-22
   • أعمدة التواريخ: عرض 15
   • أعمدة النسب المئوية: عرض 12

4️⃣ الحدود (Borders):
   • حدود رفيعة (thin) لجميع الخلايا: لون B4C7E7
   • حدود متوسطة (medium) تحت رؤوس الأعمدة
   • حدود مزدوجة (double) تحت صف الإجمالي النهائي

5️⃣ المعادلات الحية:
   • كل صف إجمالي يجب أن يحتوي على معادلة SUM حقيقية
   • مثال: =SUM(B3:B10) وليس قيمة ثابتة
   • معادلات الربط بين الأوراق: ='اسم_الورقة'!B7

6️⃣ التنسيق المحاسبي:
   • تنسيق الأرقام: '#,##0.00' (فواصل الآلاف ومنزلتان عشريتان)
   • النسب المئوية: '0.00%'
   • التواريخ: 'YYYY-MM-DD'

7️⃣ اتجاه الورقة:
   • RTL (من اليمين لليسار) للمستندات العربية
   • ws.sheet_view.rightToLeft = True

8️⃣ التجميد (Freeze Panes):
   • تجميد رؤوس الأعمدة: ws.freeze_panes = 'A3' (إذا كانت الرؤوس في الصف 2)
   • تجميد الأعمدة الأولى إذا لزم الأمر

9️⃣ الفلاتر التلقائية:
   • تفعيل AutoFilter على رؤوس الأعمدة
   • ws.auto_filter.ref = 'A2:D20'

🔟 التلوين المتبادل (Zebra Striping):
   • الصفوف الزوجية: لون D6E4F0 (أزرق فاتح)
   • الصفوف الفردية: لون FFFFFF (أبيض)
   • يبدأ التلوين من أول صف بيانات

📄 للمستندات Word (python-docx):
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

1️⃣ الهوامش:
   • جميع الهوامش: 1.8-2.0 سم

2️⃣ الخطوط:
   • العناوين: Traditional Arabic أو Simplified Arabic
   • المحتوى: Calibri أو Arial
   • الأرقام: Arial

3️⃣ الألوان:
   • نفس نظام الألوان المستخدم في Excel

4️⃣ الجداول:
   • رؤوس الأعمدة: خلفية 1F4E79، نص أبيض
   • صفوف بديلة: D6E4F0 و FFFFFF
   • حدود: Table Grid style

5️⃣ الفقرات:
   • محاذاة: RIGHT للعربية
   • تباعد الأسطر: 1.5
   • مسافة بعد الفقرة: 6-10pt

6️⃣ الترويسة والتذييل:
   • ترويسة: اسم المؤسسة (اختياري)
   • تذييل: "Officia - الوكيل الذكي" + رقم الصفحة

⚠️⚠️⚠️ تطبيق هذا البروتوكول إلزامي في كل ملف يتم إنشاؤه ⚠️⚠️⚠️
"""

ENGINE_CHEATSHEET = """
🧱 أدوات جاهزة للجداول (تُنتج تنسيقاً موحداً ومعادلات حية) — فضّلها على openpyxl الخام:
```python
from core.table_engine import col, write_table, apply_widths
from core.formula_helpers import finalize_workbook
cols = [col('name', 'البيان', 'text', width=30),
        col('qty', 'الكمية'), col('price', 'السعر'),
        col('total', 'الإجمالي', formula='={qty}*{price}', total='sum')]   # عمود مشتق = معادلة
info = write_table(ws, cols, rows, header_row=2)   # rows: قائمة dict بمفاتيح الأعمدة. يكتب صف إجمالي بمعادلات SUM
apply_widths(ws, cols); finalize_workbook(wb); wb.save(path)
```
رموز المعادلات: {key} خلية نفس الصف | {group:اسم} أعمدة المجموعة | {key:col} نطاق العمود | {key:tot} إجمالي العمود.
قاعدة: كل إجمالي أو عمود مشتق يُكتب معادلة Excel، لا قيمة محسوبة في Python.
"""

# دوال لا يملأ النموذج وسائطها: تُبنى من الملف المرفوع
_FILE_ARGS = {
    "convert_pdf_to_word": lambda p: {"pdf_path": p, "output_dir": WORKSPACE_DIR},
    "convert_pdf_to_excel": lambda p: {"pdf_path": p, "output_dir": WORKSPACE_DIR},
    "replicate_document": lambda p: {"source_file": p, "output_dir": WORKSPACE_DIR},
}


class Agent:
    def __init__(self):
        self.llm = LLMClient()
        self.executor = Executor()
        self.max_attempts = MAX_AUTO_FIX_ATTEMPTS
        self.max_repairs = 1          # جولات إصلاح بعد نجاح ظهرت فيه تحذيرات
        print("تم تهيئة Agent - الوكيل الذكي")

    # ======================================================
    # 🧭 التوجيه
    # ======================================================
    def _route_request(self, user_request: str, uploaded_file_path: str = None) -> dict:
        """كلمات مفتاحية (الأطول يفوز) ثم تصنيف ذكي من قائمة السجل."""
        name = R.route_keywords(user_request, uploaded_file_path)
        if name:
            return {"found": True, "function": name, "type": R.REGISTRY[name]["type"], "method": "keyword",
                    "instruction": f"⚠️ توجيه تلقائي: {name}"}
        return self._llm_classify(user_request, uploaded_file_path)

    def _llm_classify(self, user_request: str, uploaded_file_path: str = None) -> dict:
        try:
            response = str(self.llm.ask(R.classify_prompt(user_request, uploaded_file_path))).strip()
            name = R.match_function_in_text(response) if not re.fullmatch(r"\W*0\W*", response) else None
            meta = R.REGISTRY.get(name) if name else None
            if meta and not (meta.get("needs_file") and not uploaded_file_path):
                return {"found": True, "function": name, "type": meta["type"], "method": "llm",
                        "instruction": f"🧠 توجيه ذكي: {name}"}
        except Exception as e:
            print(f"⚠️ فشل التصنيف الذكي: {e}")
        return {"found": False, "function": None, "type": "custom", "method": "fallback", "instruction": ""}

    # ======================================================
    # 🚀 نقطة الدخول
    # ======================================================
    def execute_task(self, user_request: str, uploaded_file_path: str = None) -> dict:
        print("\n" + "=" * 60)
        print("طلب المستخدم: " + user_request)
        if uploaded_file_path:
            print("ملف مرفوع للمعالجة: " + uploaded_file_path)
        print("=" * 60)

        routing = self._route_request(user_request, uploaded_file_path)
        print(f"  -> التوجيه: {routing['function'] or 'كود مخصص'} ({routing['method']})")
        try:
            if routing["found"]:
                result = self._run_function_mode(user_request, uploaded_file_path, routing["function"])
            else:
                result = self._run_custom_mode(user_request, uploaded_file_path)
        except Exception as e:      # شبكة أمان أخيرة: لا نرجع None أبداً
            result = self._failure(0, "", f"خطأ غير متوقع في الوكيل: {type(e).__name__}: {e}", [], "")
        result.setdefault("mode", "function" if routing["found"] else "custom")
        result.setdefault("function", routing["function"])
        return result

    # ======================================================
    # 🧩 وضع الدالة: النموذج يملأ JSON، النظام يتحقق وينفذ
    # ======================================================
    def _run_function_mode(self, user_request, uploaded_file_path, fn_name):
        meta = R.REGISTRY[fn_name]
        if meta.get("needs_file") and not uploaded_file_path:
            return self._failure(0, "", f"الدالة {fn_name} تتطلب ملفاً مرفوعاً.", [], "")

        extras = {}
        args = None
        if fn_name in _FILE_ARGS:
            args = _FILE_ARGS[fn_name](uploaded_file_path)
        elif meta.get("planner") == "letter":
            args = self._letter_args(user_request, extras)

        best, last_error, last_args, repairs = None, "", None, 0
        for attempt in range(1, self.max_attempts + 1):
            print(f"\n[محاولة {attempt}/{self.max_attempts}] وضع الدالة: {fn_name}")
            if args is None:
                args = self._ask_args(fn_name, user_request, uploaded_file_path, last_args, last_error)
                if args is None:
                    last_error = "رد النموذج ليس JSON صالحاً"
                    continue
            args = self._finalize_args(fn_name, args)
            issues = R.validate_call_args(fn_name, args)
            if issues:
                last_error, last_args, args = "؛ ".join(issues), args, None
                print(f"  -> وسائط غير صالحة: {last_error[:160]}")
                continue

            code = self._call_code(fn_name, args)
            print("  -> جاري التنفيذ...")
            result = self.executor.run_python(code)
            if not result["success"]:
                last_error, last_args, args = result["error"], args, None
                print(f"  -> فشل التنفيذ: {str(result['error'])[:100]}")
                continue

            problems = self._verify(result)
            cand = self._success(attempt, result, code, problems, extras)
            if best is None or len(problems) < len(best["warnings"]):
                best = cand
            if not problems or repairs >= self.max_repairs or attempt >= self.max_attempts:
                break
            repairs += 1
            print(f"  -> تم التنفيذ مع {len(problems)} ملاحظة، جولة إصلاح...")
            last_error = "نجح التنفيذ لكن ظهرت ملاحظات يجب معالجتها:\n- " + "\n- ".join(problems)
            last_args, args = args, None

        if best:
            print(f"\n✅ نجحت المهمة" + (" (مع ملاحظات)" if best["warnings"] else ""))
            return best
        print(f"\n❌ فشلت المهمة بعد {self.max_attempts} محاولات")
        return self._failure(self.max_attempts, "", last_error, [], "")

    def _letter_args(self, user_request, extras):
        """الخطابات: صياغة ضافية بثلاث مراحل (ملخص → مسودة → فحص تغطية) ثم الإنشاء."""
        try:
            try:
                from core.letter_planner import plan_letter
            except ImportError:
                from letter_planner import plan_letter
            plan = plan_letter(user_request, self.llm.ask)
        except Exception as e:
            print(f"⚠️ تعذر تخطيط الخطاب ({e}); سيُملأ الخطاب بالمسار العادي")
            return None
        brief, draft = plan["brief"], plan["draft"]
        extras["placeholders"] = plan["placeholders"]
        extras["uncovered"] = plan["uncovered"]
        if plan["placeholders"]:
            print("  -> حقول ناقصة يجب أن يملأها المستخدم: " + "، ".join(plan["placeholders"]))
        return {"organization_name": brief.get("sender") or "المؤسسة",
                "title": brief.get("letter_type") or "خطاب رسمي",
                "recipient": brief.get("recipient") or "السادة / [المرسل إليه]",
                "subject": draft.get("subject") or brief.get("subject") or "",
                "paragraphs": draft["paragraphs"],
                "closing_text": draft.get("closing_text") or "وتفضلوا بقبول فائق الاحترام والتقدير،،",
                "note": draft.get("note")}

    def _ask_args(self, fn_name, user_request, uploaded_file_path, last_args, last_error):
        prompt = R.function_args_prompt(fn_name, user_request, uploaded_file_path)
        if last_error:
            prompt += ("\n\n⚠️ المحاولة السابقة لم تنجح.\nالوسائط السابقة:\n"
                       + json.dumps(last_args, ensure_ascii=False, default=str)[:3500]
                       + f"\nالمشكلة: {str(last_error)[:1500]}\nأعد JSON كاملاً بعد تصحيح المشكلة.")
        try:
            data = R.extract_json(self.llm.ask(prompt))
        except Exception as e:
            print(f"  -> {e}")
            return None
        inner = data.get("args")
        return inner if isinstance(inner, dict) else data

    def _finalize_args(self, fn_name, args):
        """النظام يحدد مسار الإخراج، ويحذف ما يضعه النموذج هناك."""
        args = dict(args)
        sig = R.get_signature(fn_name)
        if sig and "output_dir" in sig.parameters:
            args["output_dir"] = WORKSPACE_DIR
        return args

    @staticmethod
    def _call_code(fn_name, args):
        module = R.REGISTRY[fn_name]["module"]
        return (f"from {module} import {fn_name}\n"
                f"_args = {args!r}\n"
                f"_result = {fn_name}(**_args)\n"
                f"_path = _result.get('output_path') if isinstance(_result, dict) else _result\n"
                f"print('SUCCESS: ' + str(_path))\n")

    # ======================================================
    # 🛠️ وضع الكود المخصص (مع حارس AST)
    # ======================================================
    def _run_custom_mode(self, user_request, uploaded_file_path):
        prompt = self._build_custom_prompt(user_request, uploaded_file_path)
        best, last_error, last_code, repairs = None, "", "", 0
        for attempt in range(1, self.max_attempts + 1):
            print(f"\n[محاولة {attempt}/{self.max_attempts}] وضع الكود المخصص")
            try:
                generated = self.llm.ask(prompt)
            except Exception as e:      # انقطاع/حد معدل: نحتفظ بأفضل نتيجة نجحت سابقاً
                last_error = f"تعذر الاتصال بالنموذج: {e}"
                print("  -> " + last_error)
                continue
            code = self._extract_code(generated)
            if not code:
                last_error = "لم يتم استخراج كود Python صالح (ضع الكود في كتلة ```python)"
                print("  -> " + last_error)
                prompt = self._build_fix_prompt(user_request, str(generated)[:3000], last_error)
                continue

            guard = R.check_and_fix_code(code)
            if guard["fixes"]:
                print("  -> إصلاح تلقائي للأسماء المخترعة: " + "؛ ".join(guard["fixes"]))
                code = guard["code"]
            if guard["issues"]:
                last_error, last_code = "؛ ".join(guard["issues"]), code
                print(f"  -> رفض الحارس: {last_error[:160]}")
                prompt = self._build_fix_prompt(user_request, code, "رفضه الحارس قبل التنفيذ: " + last_error)
                continue

            print("  -> جاري التنفيذ...")
            result = self.executor.run_python(code)
            if not result["success"]:
                last_error, last_code = result["error"], code
                print(f"  -> فشل التنفيذ: {str(result['error'])[:100]}")
                if attempt < self.max_attempts:
                    prompt = self._build_fix_prompt(user_request, code, result["error"])
                continue

            problems = self._verify(result)
            cand = self._success(attempt, result, code, problems, {})
            if best is None or len(problems) < len(best["warnings"]):
                best = cand
            if not problems or repairs >= self.max_repairs or attempt >= self.max_attempts:
                break
            repairs += 1
            print(f"  -> تم التنفيذ مع {len(problems)} ملاحظة، جولة إصلاح...")
            prompt = self._build_fix_prompt(
                user_request, code, "نجح التنفيذ لكن ظهرت ملاحظات يجب معالجتها:\n- " + "\n- ".join(problems))

        if best:
            print(f"\n✅ نجحت المهمة" + (" (مع ملاحظات)" if best["warnings"] else ""))
            return best
        print(f"\n❌ فشلت المهمة بعد {self.max_attempts} محاولات")
        return self._failure(self.max_attempts, "", last_error, [], last_code)

    # ======================================================
    # ✅ التحقق من المخرجات
    # ======================================================
    def _verify(self, result):
        """مشكلات تستدعي إصلاحاً: لا ملف، تحذيرات ⚠️ من الدوال، إجماليات ثابتة في Excel."""
        output = str(result.get("output", ""))
        files = [f if os.path.isabs(str(f)) else os.path.join(WORKSPACE_DIR, str(f))
                 for f in (result.get("files_created") or [])]
        m = re.search(r"SUCCESS:\s*(.+)", output)
        if m and m.group(1).strip() not in files and os.path.exists(m.group(1).strip()):
            files.append(m.group(1).strip())
        problems = []
        if not files and "SUCCESS:" not in output:
            problems.append("لم يُنشأ أي ملف ولم تُطبع رسالة SUCCESS")
        problems += [l.strip() for l in output.splitlines() if l.strip().startswith("⚠️")]
        try:
            try:
                from core.formula_helpers import find_hardcoded_totals
            except ImportError:
                from formula_helpers import find_hardcoded_totals
            for f in files:
                if str(f).lower().endswith(".xlsx") and os.path.exists(f):
                    problems += [f"إجمالي ثابت بدل معادلة: {p}" for p in find_hardcoded_totals(f)[:5]]
        except Exception:
            pass
        return list(dict.fromkeys(problems))

    @staticmethod
    def _success(attempt, result, code, problems, extras):
        return {"success": True, "attempts": attempt, "final_output": result["output"],
                "files_created": result.get("files_created", []), "code": code,
                "warnings": problems, **extras}

    @staticmethod
    def _failure(attempts, output, error, files, code):
        return {"success": False, "attempts": attempts, "final_output": output, "error": error,
                "files_created": files, "code": code, "warnings": []}

    # ======================================================
    # 📝 الموجهات
    # ======================================================
    def _build_custom_prompt(self, user_request, uploaded_file_path=None):
        file_line = f"\n📎 ملف مرفوع: {uploaded_file_path}" if uploaded_file_path else ""
        return f"""أنت خبير في إنشاء مستندات احترافية بمعايير محاسبية وإدارية عالية.

🎯 طلب المستخدم: {user_request}{file_line}

لا توجد دالة جاهزة تطابق هذا الطلب، فاكتب كود Python كاملاً من الصفر (openpyxl / python-docx)
واحفظ الناتج في المجلد: {WORKSPACE_DIR}  ثم اطبع: print("SUCCESS: " + output_path)

{R.allowed_functions_prompt()}
{ENGINE_CHEATSHEET}
{STYLE_PROTOCOL}

اكتب الكود الكامل الآن داخل كتلة ```python واحدة."""

    def _build_fix_prompt(self, user_request, previous_code, error):
        """موجه التصحيح (كل استدعاء للنموذج بلا ذاكرة، فنعيد القائمة المسموحة هنا)."""
        return f"""الكود السابق لم ينجح.

طلب المستخدم: {user_request}

الكود السابق:
```python
{previous_code}
```

المشكلة: {str(error)[:2500]}

{R.allowed_functions_prompt()}

صحّح الكود وأعد كتابته كاملاً داخل كتلة ```python واحدة، واحفظ الناتج في {WORKSPACE_DIR} ثم اطبع SUCCESS: والمسار.
لا تخترع دوالاً ولا وسائط غير موجودة."""

    def _extract_code(self, response: str) -> str:
        """استخراج كود Python: يفضّل الكتلة التي فيها SUCCESS ثم الأطول، ويقبل كوداً بلا أسوار إن كان صالحاً."""
        if not response:
            return None
        blocks = [b.strip() for b in re.findall(r"```(?:python|py)?[ \t]*\n(.*?)```", response, re.S) if b.strip()]
        if blocks:
            good = [b for b in blocks if "SUCCESS" in b] or blocks
            return max(good, key=len)
        m = re.search(r"```(?:python|py)?[ \t]*\n(.*)$", response, re.S)     # كتلة غير مغلقة (رد مقطوع)
        if m and m.group(1).strip():
            return m.group(1).strip()
        s = response.strip()
        if re.match(r"^(from|import)\s", s):
            try:
                ast.parse(s)
                return s
            except SyntaxError:
                return None
        return None
