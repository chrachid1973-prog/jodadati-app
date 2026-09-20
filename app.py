# -*- coding: utf-8 -*-
import sys
import os

# فرض ترميز UTF-8 لمنع أخطاء الترميز نهائياً
os.environ["PYTHONIOENCODING"] = "utf-8"
if sys.platform.startswith('win'):
    try:
        import io
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
        sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8')
    except Exception:
        pass

import streamlit as st
from google import genai
import docx
from docx.shared import Cm, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import json
import io

st.set_page_config(page_title="منصة الجذاذات التربوية الشاملة", layout="wide")

# محاذاة الجدول ليلتصق بالحافة اليمنى تماماً مع اتجاه RTL
def set_table_rtl_and_right(table):
    tblPr = table._tbl.tblPr
    bidiVisual = OxmlElement('w:bidiVisual')
    tblPr.append(bidiVisual)
    jc = OxmlElement('w:jc')
    jc.set(qn('w:val'), 'right')
    tblPr.append(jc)

# ضبط خصائص الخلية وتظليلها
def set_cell_properties(cell, fill_hex=None):
    tcPr = cell._element.get_or_add_tcPr()
    bidi = OxmlElement('w:bidi')
    bidi.set(qn('w:val'), '1')
    tcPr.append(bidi)
    if fill_hex:
        shd = OxmlElement('w:shd')
        shd.set(qn('w:val'), 'clear')
        shd.set(qn('w:color'), 'auto')
        shd.set(qn('w:fill'), fill_hex)
        tcPr.append(shd)

# ضبط الفقرة لتبدأ من اليمين مع خاصية RTL
def set_paragraph_rtl_right(p):
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    pPr = p._p.get_or_add_pPr()
    bidi = OxmlElement('w:bidi')
    bidi.set(qn('w:val'), '1')
    pPr.append(bidi)

# بناء مستند Word الرسمي
def create_word_jodada(data):
    doc = docx.Document()

    # الهوامش الرسمية: 0.5 سم علوي/سفلي، 1.0 سم يمين/يسار
    for section in doc.sections:
        section.top_margin = Cm(0.5)
        section.bottom_margin = Cm(0.5)
        section.right_margin = Cm(1.0)
        section.left_margin = Cm(1.0)
        section.header_distance = Cm(0.2)
        section.footer_distance = Cm(0.2)

    # البسملة
    p_bism = doc.add_paragraph()
    set_paragraph_rtl_right(p_bism)
    p_bism.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_bism.paragraph_format.space_before = Pt(0)
    p_bism.paragraph_format.space_after = Pt(2)
    run_bism = p_bism.add_run("بـــســـم الله الـرحـمـن الـرحــيــــم")
    run_bism.bold = True
    run_bism.font.name = 'Traditional Arabic'
    run_bism.font.size = Pt(14)

    # جدول الترويسة (3 أعمدة × 5 صفوف)
    t_head = doc.add_table(rows=5, cols=3)
    set_table_rtl_and_right(t_head)

    # الصف 1
    t_head.cell(0, 0).text = f"المستوى : {data['level']}"
    t_head.cell(0, 1).text = f"المادة : {data['subject']}"
    t_head.cell(0, 2).text = f"الأستاذ(ة) : {data['teacher']}"

    # الصف 2
    t_head.cell(1, 0).text = f"المؤسسة : {data['school']}"
    t_head.cell(1, 1).text = f"المرجع : {data['reference']}"
    t_head.cell(1, 2).text = f"المكون : {data['component']}"

    # الصف 3
    c_lesson = t_head.cell(2, 0).merge(t_head.cell(2, 1))
    c_lesson.text = f"الدرس : {data['lesson_title']}"
    t_head.cell(2, 2).text = f"رقم الجذاذة : {data['jodada_num']}"

    # الصف 4
    t_head.cell(3, 0).text = f"أسبوع السنة : {data['week']}"
    t_head.cell(3, 1).text = f"الحصة : {data['session']}"
    t_head.cell(3, 2).text = f"مدة الإنجاز : {data['duration']}"

    # الصف 5 (الأهداف كاملة)
    c_goals = t_head.cell(4, 0).merge(t_head.cell(4, 2))
    goals_text = "أهداف التعلم المسطرة في الدليل الرسمي :\n" + "\n".join([f"• {g}" for g in data['learning_goals']])
    c_goals.text = goals_text

    # تنسيق خلايا الترويسة
    for row in t_head.rows:
        for cell in row.cells:
            set_cell_properties(cell)
            for p in cell.paragraphs:
                set_paragraph_rtl_right(p)
                p.paragraph_format.space_before = Pt(1)
                p.paragraph_format.space_after = Pt(1)
                for r in p.runs:
                    r.font.name = 'Traditional Arabic'
                    r.font.size = Pt(12)

    doc.add_paragraph().paragraph_format.space_after = Pt(2)

    # بناء جداول الحصص التفصيلية
    def build_session_table(title, steps):
        t = doc.add_table(rows=1, cols=2)
        set_table_rtl_and_right(t)

        hdr = t.cell(0, 0).merge(t.cell(0, 1))
        hdr.text = title
        set_cell_properties(hdr, "E6E6E6")
        set_paragraph_rtl_right(hdr.paragraphs[0])
        hdr.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        hdr.paragraphs[0].runs[0].bold = True
        hdr.paragraphs[0].runs[0].font.size = Pt(13)
        hdr.paragraphs[0].runs[0].font.name = 'Traditional Arabic'

        cols_row = t.add_row()
        cols_row.cells[0].text = "المراحل / خطوات الدرس"
        cols_row.cells[1].text = "الأنشطة والسيناريو الديداكتيكي الحرفي كما ورد في دليل الأستاذ"
        for c in cols_row.cells:
            set_cell_properties(c, "F2F2F2")
            set_paragraph_rtl_right(c.paragraphs[0])
            c.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            c.paragraphs[0].runs[0].bold = True
            c.paragraphs[0].runs[0].font.size = Pt(12)
            c.paragraphs[0].runs[0].font.name = 'Traditional Arabic'

        for item in steps:
            r = t.add_row()
            c0 = r.cells[0]
            c0.text = item.get("step", "")
            set_cell_properties(c0)
            set_paragraph_rtl_right(c0.paragraphs[0])
            c0.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            if c0.paragraphs[0].runs:
                c0.paragraphs[0].runs[0].bold = True
                c0.paragraphs[0].runs[0].font.name = 'Traditional Arabic'
                c0.paragraphs[0].runs[0].font.size = Pt(11.5)

            c1 = r.cells[1]
            c1.text = item.get("activities", "")
            set_cell_properties(c1)
            for p in c1.paragraphs:
                set_paragraph_rtl_right(p)
                p.paragraph_format.space_before = Pt(2)
                p.paragraph_format.space_after = Pt(2)
                for run in p.runs:
                    run.font.name = 'Traditional Arabic'
                    run.font.size = Pt(11.5)

        for r in t.rows:
            r.cells[0].width = Cm(4.5)
            r.cells[1].width = Cm(14.5)

        doc.add_paragraph().paragraph_format.space_after = Pt(3)

    if data.get("session_1"):
        build_session_table("الحصة الأولى : بناء التعلمات (المدة الزمنية : 45 دقيقة)", data["session_1"])
    if data.get("session_2"):
        build_session_table("الحصة الثانية : تقويم التعلمات ودعمها وتثبيتها (المدة الزمنية : 45 دقيقة)", data["session_2"])

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# واجهة المنصة
st.markdown("<h2 style='text-align: right; direction: rtl; color: #1E3A8A;'>منصة الجذاذات التربوية الشاملة (المنهاج المغربي)</h2>", unsafe_allow_html=True)
st.markdown("<p style='text-align: right; direction: rtl;'>توليد جذاذات رسمية دقيقة لجميع المستويات والمواد في جداول Word مطابقة للدليل وملتصقة باليمين</p>", unsafe_allow_html=True)
st.write("---")

with st.sidebar:
    st.header("إعدادات الاتصال")
    api_key = st.text_input("مفتاح Gemini API:", type="password")
    st.markdown("[احصل على مفتاح مجاني من Google AI Studio](https://aistudio.google.com/app/apikey)")

# حفظ واسترجاع بيانات الأستاذ عبر Session State
if "teacher_name" not in st.session_state:
    st.session_state.teacher_name = "مولاي ارشيد شكيري"
if "school_name" not in st.session_state:
    st.session_state.school_name = "مدرسة المصلى"
if "directorate" not in st.session_state:
    st.session_state.directorate = "مكناس"

with st.expander("بيانات الأستاذ(ة) والمؤسسة (تُحفظ تلقائياً لكل الوثائق)", expanded=True):
    col_adm1, col_adm2, col_adm3 = st.columns(3)
    with col_adm1:
        st.session_state.teacher_name = st.text_input("اسم الأستاذ(ة):", value=st.session_state.teacher_name)
    with col_adm2:
        st.session_state.school_name = st.text_input("المؤسسة التعليمية:", value=st.session_state.school_name)
    with col_adm3:
        st.session_state.directorate = st.text_input("المديرية الإقليمية:", value=st.session_state.directorate)

# قاعدة بيانات المناهج الشاملة (جميع المستويات والمواد والمراجع)
CURRICULUM_DB = {
    "التربية الإسلامية": {
        "components": ["الحكمة", "القسط", "الاستجابة", "الاقتداء", "التزكية (العقيدة)", "التزكية (القرآن الكريم)"],
        "references": {
            "الخامس": ["في رحاب التربية الإسلامية", "واحة التربية الإسلامية", "الممتاز في التربية الإسلامية", "مرجع آخر / كتابة يدوية"],
            "السادس": ["في رحاب التربية الإسلامية", "واحة التربية الإسلامية", "الممتاز في التربية الإسلامية", "مرجع آخر / كتابة يدوية"],
            "الرابع": ["الممتاز في التربية الإسلامية", "واحة التربية الإسلامية", "في رحاب التربية الإسلامية", "الواضح في التربية الإسلامية", "مرجع آخر / كتابة يدوية"],
            "الثالث": ["الممتاز في التربية الإسلامية", "واحة التربية الإسلامية", "مرجع آخر / كتابة يدوية"],
            "الثاني": ["في رحاب التربية الإسلامية", "المفيد في التربية الإسلامية", "الواضح في التربية الإسلامية", "مرجع آخر / كتابة يدوية"],
            "الأول": ["في رحاب التربية الإسلامية", "المفيد في التربية الإسلامية", "الواضح في التربية الإسلامية", "مرجع آخر / كتابة يدوية"]
        }
    },
    "الاجتماعيات": {
        "components": ["التاريخ", "الجغرافيا", "التربية المدنية"],
        "references": {
            "السادس": ["الجديد في الاجتماعيات", "المسار في الاجتماعيات", "في رحاب الاجتماعيات", "مرجع آخر / كتابة يدوية"],
            "الخامس": ["المفيد في الاجتماعيات", "النجاح في الاجتماعيات", "مسار الاجتماعيات", "مرجع آخر / كتابة يدوية"],
            "الرابع": ["الجديد في الاجتماعيات", "المفيد في الاجتماعيات", "الواضح في الاجتماعيات", "مسار الاجتماعيات", "مرجع آخر / كتابة يدوية"]
        }
    },
    "اللغة العربية": {
        "components": ["القراءة", "الظواهر اللغوية (تراكيب / صرف وتحويل / إملاء)", "التعبير الكتابي (الإنشاء)", "التواصل الشفهي", "مشروع الوحدة"],
        "references": {
            "السادس": ["كتابي في اللغة العربية", "في رحاب اللغة العربية", "المنير في اللغة العربية", "مرجع آخر / كتابة يدوية"],
            "الخامس": ["مرشدي في اللغة العربية", "المنير في اللغة العربية", "كتابي في اللغة العربية", "مرجع آخر / كتابة يدوية"],
            "الرابع": ["الواحة في اللغة العربية", "المفيد في اللغة العربية", "مرشدي في اللغة العربية", "مرجع آخر / كتابة يدوية"],
            "الثالث": ["المفيد في اللغة العربية", "مرشدي في اللغة العربية", "واحة الكلمات العربية", "مرجع آخر / كتابة يدوية"],
            "الثاني": ["المفيد في اللغة العربية", "كتابي في اللغة العربية", "مرجع آخر / كتابة يدوية"],
            "الأول": ["المفيد في اللغة العربية", "كتابي في اللغة العربية", "مرجع آخر / كتابة يدوية"]
        }
    },
    "النشاط العلمي": {
        "components": ["علوم الحياة والأرض", "العلوم الفيزيائية والتكنولوجية", "الفلك والفضاء"],
        "references": {
            "السادس": ["المنير في النشاط العلمي", "فضاء النشاط العلمي", "الواضح في النشاط العلمي", "مرجع آخر / كتابة يدوية"],
            "الخامس": ["المنير في النشاط العلمي", "الواضح في النشاط العلمي", "المفيد في النشاط العلمي", "فضاء النشاط العلمي", "مرجع آخر / كتابة يدوية"],
            "الرابع": ["المرشد في النشاط العلمي", "المنهل في النشاط العلمي", "فضاء النشاط العلمي", "مرجع آخر / كتابة يدوية"],
            "الثالث": ["المنهل في النشاط العلمي", "الواضح في النشاط العلمي", "فضاء النشاط العلمي", "مرجع آخر / كتابة يدوية"],
            "الثاني": ["الواضح في النشاط العلمي", "فضاء النشاط العلمي", "الأساس في النشاط العلمي", "مرجع آخر / كتابة يدوية"],
            "الأول": ["الواضح في النشاط العلمي", "فضاء النشاط العلمي", "الأساس في النشاط العلمي", "مرجع آخر / كتابة يدوية"]
        }
    },
    "الرياضيات": {
        "components": ["الأعداد والحساب", "الهندسة والفضاء", "القياس", "تنظيم ومعالجة البيانات", "حل المسائل"],
        "references": {
            "السادس": ["النجاح في الرياضيات", "الجيد في الرياضيات", "المفيد في الرياضيات", "مرجع آخر / كتابة يدوية"],
            "الخامس": ["المفيد في الرياضيات", "النجاح في الرياضيات", "المرجع في الرياضيات", "الجيد في الرياضيات", "مرجع آخر / كتابة يدوية"],
            "الرابع": ["المفيد في الرياضيات", "الجيد في الرياضيات", "المرجع في الرياضيات", "مرجع آخر / كتابة يدوية"],
            "الثالث": ["المفيد في الرياضيات", "المرجع في الرياضيات", "فضاء الرياضيات", "مرجع آخر / كتابة يدوية"],
            "الثاني": ["المفيد في الرياضيات", "فضاء الرياضيات", "مرجع آخر / كتابة يدوية"],
            "الأول": ["المفيد في الرياضيات", "فضاء الرياضيات", "مرجع آخر / كتابة يدوية"]
        }
    },
    "اللغة الفرنسية": {
        "components": ["Communication et actes de langage", "Lecture", "Grammaire / Conjugaison / Orthographe", "Production de l'écrit", "Projet de classe"],
        "references": {
            "السادس": ["Mes apprentissages en français", "Pour communiquer en français", "L'oasis des mots", "Le chemin des lettres", "مرجع آخر / كتابة يدوية"],
            "الخامس": ["Mes apprentissages en français", "Pour communiquer en français", "L'oasis des mots", "Le chemin des lettres", "مرجع آخر / كتابة يدوية"],
            "الرابع": ["Mes apprentissages en français", "Pour communiquer en français", "L'oasis des mots", "مرجع آخر / كتابة يدوية"],
            "الثالث": ["Mes apprentissages en français", "Pour communiquer en français", "L'oasis des mots", "مرجع آخر / كتابة يدوية"],
            "الثاني": ["Mes apprentissages en français", "Pour communiquer en français", "مرجع آخر / كتابة يدوية"],
            "الأول": ["Mes apprentissages en français", "Dire, faire et agir", "مرجع آخر / كتابة يدوية"]
        }
    },
    "التربية الفنية": {
        "components": ["الرسم والتشكيل", "الموسيقى والأناشيد", "المسرح"],
        "references": {
            "السادس": ["عالم التربية الفنية", "الممتاز في التربية الفنية", "المفيد في التربية الفنية", "مرجع آخر / كتابة يدوية"],
            "الخامس": ["عالم التربية الفنية", "الممتاز في التربية الفنية", "المفيد في التربية الفنية", "مرجع آخر / كتابة يدوية"],
            "الرابع": ["عالم التربية الفنية", "الممتاز في التربية الفنية", "المفيد في التربية الفنية", "مرجع آخر / كتابة يدوية"],
            "الثالث": ["عالم التربية الفنية", "الممتاز في التربية الفنية", "المفيد في التربية الفنية", "مرجع آخر / كتابة يدوية"],
            "الثاني": ["عالم التربية الفنية", "الممتاز في التربية الفنية", "مرجع آخر / كتابة يدوية"],
            "الأول": ["عالم التربية الفنية", "الممتاز في التربية الفنية", "مرجع آخر / كتابة يدوية"]
        }
    }
}

st.markdown("#### تحديد محددات الدرس")
col1, col2 = st.columns(2)
with col1:
    level = st.selectbox("المستوى الدراسي:", ["الخامس", "السادس", "الرابع", "الثالث", "الثاني", "الأول"])

valid_subjects = [s for s, d in CURRICULUM_DB.items() if level in d["references"]]

with col2:
    subject = st.selectbox("المادة الدراسية:", valid_subjects)

col3, col4 = st.columns(2)
with col3:
    component = st.selectbox("المكون / المدخل:", CURRICULUM_DB[subject]["components"])

with col4:
    ref_options = CURRICULUM_DB[subject]["references"].get(level, ["مرجع رسمي معتمد"])
    selected_ref = st.selectbox("المرجع المعتمد في المادة:", ref_options)
    if selected_ref == "مرجع آخر / كتابة يدوية":
        reference = st.text_input("اكتب اسم المرجع هنا:")
    else:
        reference = selected_ref

col5, col6, col7 = st.columns(3)
with col5:
    lesson_order = st.number_input("ترتيب الدرس في المقرر (فقط الرقم):", min_value=1, max_value=40, value=1)
with col6:
    session_choice = st.selectbox("الحصة المطلوبة:", ["الحصتان معاً (1 و 2)", "الحصة الأولى فقط (بناء التعلمات)", "الحصة الثانية فقط (تقويم ودعم)"])
with col7:
    week_num = st.number_input("أسبوع السنة التقديري:", min_value=1, max_value=34, value=2)

st.write("---")

if st.button("توليد الجذاذة المطابقة للدليل بصيغة Word", type="primary", use_container_width=True):
    if not api_key:
        st.error("يرجى إدخال مفتاح Gemini API في الشريط الجانبي.")
    elif not st.session_state.teacher_name.strip() or not st.session_state.school_name.strip():
        st.warning("يرجى التأكد من ملء بيانات الأستاذ والمؤسسة في الصندوق أعلاه.")
    else:
        with st.spinner(f"جاري استخراج الجذاذة الرسمية للدرس {lesson_order} في {subject} ({reference}) عبر محرك Gemini 3.8/3.6..."):
            try:
                client = genai.Client(api_key=api_key)

                prompt = f"""
                أنت مفتش تربوي معتمد بالمملكة المغربية.
                المهمة: استخراج جذاذة تربوية رسمية مطابقة تماماً وبدون أي تلخيص أو نقص لدليل الأستاذ وكتاب التلميذ المعتمدين في المغرب.

                المحددات الإلزامية:
                - المستوى الدراسي: {level} ابتدائي حصراً.
                - المادة: {subject}.
                - المرجع الدراسي: {reference}.
                - المكون / المدخل: {component}.
                - ترتيب الدرس: الدرس رقم {lesson_order}.
                - أسبوع الإنجاز: {week_num}.

                قواعد التطابق البيداغوجي الصارمة:
                1. استخرج العنوان الرسمي الحرفي للدرس رقم {lesson_order} من فهرس كتاب التلميذ الخاص بالمرجع {reference} للمستوى {level}.
                2. انقل نص أهداف التعلم كاملة من ترويسة الدرس في دليل الأستاذ.
                3. انقل مراحل الحصة الأولى (بناء التعلمات 45د) خطوة بخطوة: النصوص الشرعية أو الوثائق أو الصور أو الأسناد بنصوصها وأرقامها، أسئلة الفهم والتحليل نصاً، خطاطة الاستخلاص والترسيخ، والتشبع بالقيم والبحث المنزلي.
                4. انقل مراحل الحصة الثانية (تقويم ودعم 45د): أسئلة التقويم الشفهي، نصوص الأنشطة والتمارين الكتابية المنجزة على الدفاتر، وإجراءات المعالجة والدعم الفوري.

                أخرج الناتج بصيغة JSON صارمة حصراً:
                {{
                    "lesson_title": "الدرس {lesson_order} : [العنوان الحرفي المستخرج من الدليل]",
                    "learning_goals": ["الهدف الأول كاملاً", "الهدف الثاني كاملاً..."],
                    "session_1": [
                        {{"step": "اسم المرحلة الأولى كما وردت في الدليل", "activities": "النصوص والأسئلة والأنشطة كاملة دون أي تلخيص"}},
                        {{"step": "اسم المرحلة التالية", "activities": "التفاصيل الدقيقة والأسئلة والخلاصات"}}
                    ],
                    "session_2": [
                        {{"step": "تقويم شفهي", "activities": "الأسئلة التقويمية كاملة"}},
                        {{"step": "تقويم كتابي (دفاتر)", "activities": "نص التمارين والوضعيات التقويمية"}},
                        {{"step": "أدعم وأعالج", "activities": "خطة المعالجة والدعم المعتمدة"}}
                    ]
                }}
                """

                # استخدام أحدث نماذج الجيل 3.8 فما فوق و 3.6 بالترتيب
                candidate_models = [
                    "gemini-3.8-flash",
                    "gemini-3.7-flash",
                    "gemini-3.6-flash"
                ]

                response = None
                last_err = None

                for m in candidate_models:
                    try:
                        response = client.models.generate_content(
                            model=m,
                            contents=prompt
                        )
                        if response and response.text:
                            break
                    except Exception as e:
                        last_err = e
                        continue

                if response is None:
                    raise Exception(f"فشل الاتصال بالنماذج (3.8 و 3.6): {last_err}")

                raw_text = response.text.strip()
                if raw_text.startswith("```json"):
                    raw_text = raw_text[7:]
                if raw_text.startswith("```"):
                    raw_text = raw_text[3:]
                if raw_text.endswith("```"):
                    raw_text = raw_text[:-3]

                res_data = json.loads(raw_text.strip())

                s1_data = res_data.get("session_1", []) if "الأولى" in session_choice or "معاً" in session_choice else []
                s2_data = res_data.get("session_2", []) if "الثانية" in session_choice or "معاً" in session_choice else []

                jodada_full = {
                    "level": level,
                    "subject": subject,
                    "teacher": st.session_state.teacher_name,
                    "school": st.session_state.school_name,
                    "unit": (lesson_order - 1) // 4 + 1,
                    "reference": reference,
                    "component": component,
                    "lesson_title": res_data.get("lesson_title", f"الدرس {lesson_order}"),
                    "jodada_num": lesson_order,
                    "week": week_num,
                    "session": session_choice,
                    "duration": "45 دقيقة" if "فقط" in session_choice else "45 دقيقة × 2",
                    "learning_goals": res_data.get("learning_goals", []),
                    "session_1": s1_data,
                    "session_2": s2_data
                }

                word_buffer = create_word_jodada(jodada_full)

                st.success(f"تم بنجاح استخراج وتنسيق جذاذة: {jodada_full['lesson_title']}")

                st.download_button(
                    label="تحميل الجذاذة بصيغة Word (.docx) منسقة في جداول",
                    data=word_buffer,
                    file_name=f"جذاذة_{subject}_{component}_{level}_الدرس_{lesson_order}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True
                )

            except Exception as e:
                st.error(f"حدث خطأ أثناء التوليد: {str(e)}")