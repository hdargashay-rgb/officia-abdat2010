"""
letter_planner.py  -  Officia
خط إنتاج الخطابات بثلاث مراحل بدل مرحلة واحدة:

    طلب المستخدم ──► (1) ملخص منظّم (Brief): الوقائع والآثار والمطالب، لكل نقطة كلمات مفتاحية
                 ──► (2) صياغة خطاب ضافٍ كامل (Draft) وفق هيكل ثابت
                 ──► (3) فحص تغطية برمجي: هل ورد كل ما ذكره المستخدم؟ إن نقص شيء → جولة مراجعة
                 ──► إنشاء ملف Word عبر generate_professional_letter

الدالة لا تعرف أي نموذج تستخدم: تأخذ llm(prompt) -> str، فاربطها بنموذجك (Groq / Gemini / ...).

الاستخدام:
    from letter_planner import plan_letter, create_letter_from_request
    path = create_letter_from_request(user_text, llm, organization_name="شركة ...",
                                      signature_name="...", signature_title="المدير العام")
"""

import json
import re

# ---------------------------------------------------------------------------
# المرحلة 1: استخراج الملخص المنظّم
# ---------------------------------------------------------------------------
BRIEF_PROMPT = """أنت مساعد متخصص في تحليل طلبات المراسلات الرسمية.
اقرأ طلب المستخدم أدناه واستخرج منه ملخصاً منظماً بصيغة JSON فقط (بلا أي نص خارجه).

قواعد صارمة:
- لا تخترع أي واقعة أو تاريخ أو رقم أو اسم لم يذكره المستخدم. ما ينقص ضعه في missing_info.
- كل معلومة ذكرها المستخدم يجب أن تظهر في facts أو impacts أو demands. لا تُسقط شيئاً، حتى التفاصيل الصغيرة (المدد، أنواع المعاملات، العواقب).
- لكل عنصر في facts/impacts/demands أعطِ keywords: كلمتان إلى أربع كلمات مفتاحية عربية قصيرة (جذور أو عبارات قصيرة) تدل عليه، لنتحقق لاحقاً أن الخطاب غطّاه.

الشكل المطلوب:
{{
  "letter_type": "نوع الخطاب (مطالبة / إنذار / شكوى / طلب / اعتذار / ...)",
  "sender": "الجهة المرسلة كما يفهم من الطلب",
  "recipient": "الجهة أو الشخص المرسل إليه",
  "relationship": "طبيعة العلاقة بين الطرفين",
  "subject": "موضوع الخطاب في سطر واحد",
  "facts": [{{"text": "...", "keywords": ["...", "..."]}}],
  "impacts": [{{"text": "...", "keywords": ["...", "..."]}}],
  "demands": [{{"text": "...", "keywords": ["...", "..."]}}],
  "deadline": "المهلة إن ذكرها المستخدم، وإلا null",
  "consequence_if_ignored": "ما سيحدث عند عدم الاستجابة كما ذكره المستخدم، وإلا null",
  "tone": "النبرة المناسبة (رسمي ودي / رسمي حازم / ...)",
  "missing_info": ["معلومات لم يذكرها المستخدم وتلزم الخطاب: أسماء، تواريخ، أرقام عقود، مهلة..."]
}}

طلب المستخدم:
\"\"\"{request}\"\"\"
"""

# ---------------------------------------------------------------------------
# المرحلة 2: صياغة الخطاب الضافي
# ---------------------------------------------------------------------------
DRAFT_PROMPT = """اكتب خطاباً رسمياً ضافياً بالعربية الفصحى اعتماداً على الملخص التالي، وأعد JSON فقط.

الهيكل الإلزامي للفقرات (كل بند فقرة أو أكثر، حسب الحاجة):
1. فقرة افتتاحية تحدد العلاقة وسبب الخطاب.
2. الخلفية والالتزام الذي لم يُنفَّذ.
3. الواقعة وما ترتب عليها مباشرة.
4. الآثار على المرسل بالتفصيل: اذكر كل أثر وارد في الملخص صراحة وبالاسم (لا تجمعها في عبارة عامة).
5. المطالب المحددة (ما المطلوب فعله بالضبط).
6. المهلة الزمنية والمتابعة (إن لم تُذكر مهلة فضع [المهلة المحددة]).
7. التحفظ على الحقوق وتحميل المسؤولية عند عدم الاستجابة (بلغة نظامية رصينة، بلا تهديد غير لائق).
8. ختام يؤكد الرغبة في الحل الودي العاجل.

قواعد:
- غطِّ كل عنصر في facts وimpacts وdemands. يجب أن تظهر كلماته المفتاحية في النص.
- لا تخترع وقائع. ما نقص من معلومات اكتبه بين أقواس معقوفة هكذا: [اسم المالك]، [رقم العقد]، [تاريخ رفض التجديد].
- لا تذكر مواد قانونية أو أرقام أنظمة من عندك.
- لا ترقّم الفقرات (لا تبدأ بـ 1- أو أ-)، ولا تستخدم رموزاً أو قوائم نقطية. فقرات نثرية متصلة.
- الطول: ما بين 6 و10 فقرات، كل فقرة 2 إلى 5 جمل. نبرة الخطاب: {tone}.

الشكل المطلوب:
{{
  "subject": "موضوع الخطاب في سطر واحد",
  "paragraphs": ["...", "..."],
  "closing_text": "عبارة الختام (مثل: وتفضلوا بقبول فائق الاحترام والتقدير،،)",
  "note": null
}}

الملخص:
{brief}
"""

# ---------------------------------------------------------------------------
# المرحلة 3: مراجعة إن نقصت تغطية
# ---------------------------------------------------------------------------
REVISE_PROMPT = """الخطاب التالي لم يغطِّ العناصر الآتية من طلب المستخدم:
{missing}

أعد كتابة الخطاب كاملاً (لا الأجزاء الناقصة فقط) مع إدخال هذه العناصر في الفقرات المناسبة بوضوح وبالاسم،
مع الحفاظ على الهيكل والنبرة وعدم اختراع وقائع. أعد JSON بالشكل نفسه فقط.

الخطاب الحالي:
{draft}
"""

# أقل أطوال مقبولة للتحقق من أن الخطاب لم يُجتزأ
MIN_PARAGRAPHS = 5
MIN_WORDS = 180


# ---------------------------------------------------------------------------
# أدوات مساعدة
# ---------------------------------------------------------------------------
def _norm(text) -> str:
    """توحيد النص العربي للمطابقة (تشكيل، همزات، تاء مربوطة، ألف مقصورة)."""
    t = re.sub(r"[\u064B-\u0652\u0640]", "", str(text))
    for a, b in (("أ", "ا"), ("إ", "ا"), ("آ", "ا"), ("ة", "ه"), ("ى", "ي")):
        t = t.replace(a, b)
    t = t.translate(str.maketrans("٠١٢٣٤٥٦٧٨٩", "0123456789"))   # الأرقام الهندية → غربية
    return re.sub(r"\s+", " ", t).strip().lower()


def parse_json(text):
    """يستخلص أول كائن JSON من رد النموذج (حتى لو أحاطه بنص أو ```json)."""
    if isinstance(text, dict):
        return text
    s = str(text).strip()
    s = re.sub(r"^```(?:json)?\s*|\s*```$", "", s, flags=re.S)
    start, depth = s.find("{"), 0
    if start < 0:
        raise ValueError("لم يرجع النموذج JSON")
    for i in range(start, len(s)):
        depth += (s[i] == "{") - (s[i] == "}")
        if depth == 0:
            return json.loads(s[start:i + 1])
    raise ValueError("JSON غير مكتمل في رد النموذج")


def check_coverage(brief: dict, draft: dict) -> list:
    """
    يرجع قائمة نصوص العناصر التي لم يغطِّها الخطاب (فارغة = تغطية كاملة).
    العنصر مغطّى إن ورد نصف كلماته المفتاحية على الأقل (بعد التوحيد).
    """
    body = _norm(" ".join(draft.get("paragraphs", [])) + " " + str(draft.get("subject", "")))
    missing = []
    for group in ("facts", "impacts", "demands"):
        for item in brief.get(group, []) or []:
            kws = [_norm(k) for k in item.get("keywords", []) if str(k).strip()]
            if not kws:
                continue
            hits = sum(1 for k in kws if k in body)
            if hits < max(1, (len(kws) + 1) // 2):
                missing.append(item.get("text", ""))
    return missing


def check_length(draft: dict) -> list:
    """أسباب قصر الخطاب (اجتزاء) إن وجدت."""
    paras = draft.get("paragraphs", []) or []
    words = sum(len(str(p).split()) for p in paras)
    problems = []
    if len(paras) < MIN_PARAGRAPHS:
        problems.append(f"عدد الفقرات {len(paras)} أقل من {MIN_PARAGRAPHS}")
    if words < MIN_WORDS:
        problems.append(f"عدد الكلمات {words} أقل من {MIN_WORDS}")
    return problems


def find_placeholders(draft: dict) -> list:
    """الحقول الناقصة بين أقواس معقوفة التي يجب أن يملأها المستخدم قبل الإرسال."""
    text = " ".join(draft.get("paragraphs", []))
    return list(dict.fromkeys(re.findall(r"\[[^\]\n]{2,60}\]", text)))


# ---------------------------------------------------------------------------
# الخط الكامل
# ---------------------------------------------------------------------------
def plan_letter(user_request: str, llm, max_revisions: int = 2) -> dict:
    """
    يرجع {'brief', 'draft', 'uncovered', 'length_problems', 'placeholders', 'revisions'}.
    llm: دالة تأخذ نص التعليمة وترجع رد النموذج نصاً.
    """
    brief = parse_json(llm(BRIEF_PROMPT.format(request=user_request)))
    draft = parse_json(llm(DRAFT_PROMPT.format(
        brief=json.dumps(brief, ensure_ascii=False, indent=2),
        tone=brief.get("tone") or "رسمي")))

    revisions = 0
    while revisions < max_revisions:
        uncovered = check_coverage(brief, draft)
        length_problems = check_length(draft)
        if not uncovered and not length_problems:
            break
        needs = list(uncovered) + [f"(الطول) {p}: وسّع الفقرات بتفصيل الآثار والمطالب" for p in length_problems]
        draft = parse_json(llm(REVISE_PROMPT.format(
            missing="\n".join(f"- {m}" for m in needs),
            draft=json.dumps(draft, ensure_ascii=False, indent=2))))
        revisions += 1

    return {
        "brief": brief,
        "draft": draft,
        "uncovered": check_coverage(brief, draft),
        "length_problems": check_length(draft),
        "placeholders": find_placeholders(draft),
        "revisions": revisions,
    }


def create_letter_from_request(user_request: str, llm, organization_name: str,
                               title: str = "خطاب رسمي", recipient: str = None,
                               signature_name: str = None, signature_title: str = None,
                               output_dir: str = None, filename: str = None,
                               max_revisions: int = 2, **letter_kwargs) -> dict:
    """
    من طلب المستخدم إلى ملف Word. يرجع {'path', 'placeholders', 'uncovered', 'plan'}.
    placeholders: حقول يجب إخبار المستخدم بها ليملأها (الخطاب يحمل [ ] حيث ينقص شيء).
    """
    try:
        from core.formatter import generate_professional_letter
    except ImportError:
        from formatter import generate_professional_letter

    plan = plan_letter(user_request, llm, max_revisions=max_revisions)
    brief, draft = plan["brief"], plan["draft"]
    path = generate_professional_letter(
        organization_name=organization_name,
        title=title,
        recipient=recipient or brief.get("recipient") or "السادة / [المرسل إليه]",
        subject=draft.get("subject") or brief.get("subject") or "",
        paragraphs=draft["paragraphs"],
        closing_text=draft.get("closing_text") or "وتفضلوا بقبول فائق الاحترام والتقدير،،",
        signature_name=signature_name,
        signature_title=signature_title,
        note=draft.get("note"),
        output_dir=output_dir,
        filename=filename,
        **letter_kwargs,
    )
    return {"path": path, "placeholders": plan["placeholders"],
            "uncovered": plan["uncovered"], "plan": plan}
