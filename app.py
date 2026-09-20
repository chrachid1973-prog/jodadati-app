# -*- coding: utf-8 -*-
import sys
import os

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
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
import json
import io

st.set_page_config(page_title="منصة الجذاذات التربوية الرسمية", layout="wide")

# رسم حدود الجدول بأسلوب مسطر ورسمي
def set_table_borders(table):
    tblPr = table._tbl.tblPr
    borders = parse_xml(r'''
        <w:tblBorders {} >
            <w:top w:val="single" w:sz="6" w:space="0" w:color="000000"/>
            <w:left w:val="single" w:sz="6" w:space="0" w:color="000000"/>
            <w:bottom w:val="single" w:sz="6" w:space="0" w:color="000000"/>
            <w:right w:val="single" w:sz="6" w:space="0" w:color="000000"/>
            <w:insideH w:val="single" w:sz="4" w:space="0" w:color="888888"/>
            <w:insideV w:val="single" w:sz="4" w:space="0" w:color="888888"/>
        </w:tblBorders>
    '''.format(nsdecls('w')))
    tblPr.append(borders)

# محاذاة الجدول ليلتصق بالحافة اليمنى تماماً مع خاصية RTL
def set_table_rtl_and_right(table):
    tblPr = table._tbl.tblPr
    bidiVisual = parse_xml(r'<w:bidiVisual {}/>'.format(nsdecls('w')))
    tblPr.append(bidiVisual)
    jc = parse_xml(r'<w:jc {} w:val="right"/>'.format(nsdecls('w')))
    tblPr.append(jc)
    set_table_borders(table)

# ضبط خصائص الخلية وتظليلها
def format_cell(cell, fill_hex=None):
    tcPr = cell._element.get_or_add_tcPr()
    bidi = parse_xml(r'<w:bidi {} w:val="1"/>'.format(nsdecls('w')))
    tcPr.append(bidi)
    if fill_hex:
        shd = parse_xml(r'<w:shd {} w:val="clear" w:color="auto" w:fill="{}"/>'.format(nsdecls('w'), fill_hex))
        tcPr.append(shd)

# ضبط الفقرة باتجاه RTL نقي
def format_paragraph_rtl(p, align=WD_ALIGN_PARAGRAPH.RIGHT, space_before=2, space_after=2):
    p.alignment = align
    pPr = p._p.get_or_add_pPr()
    pPr.append(parse_xml(r'<w:bidi {} w:val="1"/>'.format(nsdecls('w'))))
    p.paragraph_format.space_before = Pt(space_before)
    p.paragraph_format.space_after = Pt(space_after)

# إضافة النص مع خاصية RTL العربية
def add_arabic_run(paragraph, text, font_size=11.5, bold=False):
    run = paragraph.add_run(text)
    run.bold = bold
    run.font.name = 'Traditional Arabic'
    run.font.size = Pt(font_size)
    rPr = run._r.get_or_add_rPr()
    rPr.append(parse_xml(r'<w:rtl {} w:val="1"/>'.format(nsdecls('w'))))
    return run

# بناء المستند بجداول مسطرة ورأسية مضبوطة
def create_word_jodada(data):
    doc = docx.Document()

    for section in doc.sections:
        section.top_margin = Cm(0.5)
        section.bottom_margin = Cm(0.5)
        section.right_margin = Cm(1.0)
        section.left_margin = Cm(1.0)
        section.header_distance = Cm(0.2)
        section.footer_distance = Cm(0.2)

    # البسملة
    p_bism = doc.add_paragraph()
    format_paragraph_rtl(p_bism, align=WD_ALIGN_PARAGRAPH.CENTER, space_before=0, space_after=3)
    add_arabic_run(p_bism, "بـــســـم الله الـرحـمـن الـرحــيــــم", font_size=13, bold=True)

    # جدول الرأسية المسطر (3 أعمدة × 5 صفوف)
    t_head = doc.add_table(rows=5, cols=3)
    set_table_rtl_and_right(t_head)

    # الصف 1
    p = t_head.cell(0, 0).paragraphs[0]
    format_paragraph_rtl(p)
    add_arabic_run(p, f"المستوى : {data['level']}", bold=True)

    p = t_head.cell(0, 1).paragraphs[0]
    format_paragraph_rtl(p)
    add_arabic_run(p, f"المادة : {data['subject']}", bold=True)

    p = t_head.cell(0, 2).paragraphs[0]
    format_paragraph_rtl(p)
    add_arabic_run(p, f"الأستاذ(ة) : {data['teacher']}", bold=True)

    # الصف 2 (خانة المؤسسة الرسمية)
    p = t_head.cell(1, 0).paragraphs[0]
    format_paragraph_rtl(p)
    add_arabic_run(p, f"المؤسسة : {data['school']}", bold=True)

    p = t_head.cell(1, 1).paragraphs[0]
    format_paragraph_rtl(p)
    add_arabic_run(p, f"المرجع : {data['reference']}", bold=True)

    p = t_head.cell(1, 2).paragraphs[0]
    format_paragraph_rtl(p)
    add_arabic_run(p, f"المكون : {data['component']}", bold=True)

    # الصف 3
    c_lesson = t_head.cell(2, 0).merge(t_head.cell(2, 1))
    p = c_lesson.paragraphs[0]
    format_paragraph_rtl(p)
    add_arabic_run(p, f"الدرس : {data['lesson_title']}", bold=True)

    p = t_head.cell(2, 2).paragraphs[0]
    format_paragraph_rtl(p)
    add_arabic_run(p, f"رقم الجذاذة : {data['jodada_num']}", bold=True)

    # الصف 4
    p = t_head.cell(3, 0).paragraphs[0]
    format_paragraph_rtl(p)
    add_arabic_run(p, f"أسبوع السنة : {data['week']}", bold=True)

    p = t_head.cell(3, 1).paragraphs[0]
    format_paragraph_rtl(p)
    add_arabic_run(p, f"الحصة : {data['session']}", bold=True)

    p = t_head.cell(3, 2).paragraphs[0]
    format_paragraph_rtl(p)
    add_arabic_run(p, f"مدة الإنجاز : {data['duration']}", bold=True)

    # الصف 5 (أهداف التعلم)
    c_goals = t_head.cell(4, 0).merge(t_head.cell(4, 2))
    p_title = c_goals.paragraphs[0]
    format_paragraph_rtl(p_title, space_before=1, space_after=1)
    add_arabic_run(p_title, "أهداف التعلم :", bold=True, font_size=12)

    for goal in data['learning_goals']:
        p_g = c_goals.add_paragraph()
        format_paragraph_rtl(p_g, space_before=1, space_after=1)
        add_arabic_run(p_g, f"• {goal}", font_size=11.5)

    for row in t_head.rows:
        for cell in row.cells:
            format_cell(cell)

    doc.add_paragraph().paragraph_format.space_after = Pt(2)

    # جدول خطوات الحصة
    def build_session_table(session_title, steps):
        t = doc.add_table(rows=1, cols=2)
        set_table_rtl_and_right(t)

        hdr = t.cell(0, 0).merge(t.cell(0, 1))
        format_cell(hdr, "E8E8E8")
        p_hdr = hdr.paragraphs[0]
        format_paragraph_rtl(p_hdr, align=WD_ALIGN_PARAGRAPH.CENTER)
        add_arabic_run(p_hdr, session_title, font_size=12.5, bold=True)

        cols_row = t.add_row()
        cols_row.cells[0].text = ""
        cols_row.cells[1].text = ""

        p0 = cols_row.cells[0].paragraphs[0]
        format_paragraph_rtl(p0, align=WD_ALIGN_PARAGRAPH.CENTER)
        add_arabic_run(p0, "المراحل / خطوات الدرس", font_size=12, bold=True)

        p1 = cols_row.cells[1].paragraphs[0]
        format_paragraph_rtl(p1, align=WD_ALIGN_PARAGRAPH.CENTER)
        add_arabic_run(p1, "الأنشطة التعليمية التعلمية", font_size=12, bold=True)

        format_cell(cols_row.cells[0], "F2F2F2")
        format_cell(cols_row.cells[1], "F2F2F2")

        for item in steps:
            r = t.add_row()
            c0 = r.cells[0]
            c1 = r.cells[1]
            format_cell(c0)
            format_cell(c1)

            p_step = c0.paragraphs[0]
            format_paragraph_rtl(p_step, align=WD_ALIGN_PARAGRAPH.CENTER)
            add_arabic_run(p_step, item.get("step", ""), font_size=11.5, bold=True)

            p_act = c1.paragraphs[0]
            format_paragraph_rtl(p_act, align=WD_ALIGN_PARAGRAPH.RIGHT, space_before=2, space_after=2)
            act_text = item.get("activities", "")
            
            lines = act_text.split("\n")
            for i, line in enumerate(lines):
                if i > 0:
                    p_act = c1.add_paragraph()
                    format_paragraph_rtl(p_act, align=WD_ALIGN_PARAGRAPH.RIGHT, space_before=1, space_after=1)
                add_arabic_run(p_act, line.strip(), font_size=11.5)

        for r in t.rows:
            r.cells[0].width = Cm(4.5)
            r.cells[1].width = Cm(14.5)

        doc.add_paragraph().paragraph_format.space_after = Pt(3)

    if data.get("session_1"):
        build_session_table("الحصة الأولى (45 دقيقة)", data["session_1"])
    if data.get("session_2"):
        build_session_table("الحصة الثانية (45 دقيقة)", data["session_2"])

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

st.markdown("<h2 style='text-align: right; direction: rtl; color: #1E3A8A;'>منصة الجذاذات التربوية الرسمية</h2>", unsafe_allow_html=True)
st.markdown("<p style='text-align: right; direction: rtl;'>توليد جذاذات رسمية مطابقة للكتاب المدرسي في جداول مسطرة ومحاذاة لليمين 100%</p>", unsafe_allow_html=True)
st.write("---")

with st.sidebar:
    st.header("إعدادات الاتصال")
    api_key = st.text_input("مفتاح Gemini API:", type="password")
    st.markdown("[احصل على مفتاح مجاني من Google AI Studio](https://aistudio.google.com/app/apikey)")

if "teacher_name" not in st.session_state:
    st.session_state.teacher_name = "مولاي ارشيد شكيري"
if "school_name" not in st.session_state:
    st.session_state.school_name = "مدرسة المصلى"
if "directorate" not in st.session_state:
    st.session_state.directorate = "مكناس"

with st.expander("البيانات الإدارية (تُحفظ تلقائياً في الترويسة المسطرة)", expanded=True):
    col_adm1, col_adm2, col_adm3 = st.columns(3)
    with col_adm1:
        st.session_state.teacher_name = st.text_input("اسم الأستاذ(ة):", value=st.session_state.teacher_name)
    with col_adm2:
        st.session_state.school_name = st.text_input("المؤسسة التعليمية:", value=st.session_state.school_name)
    with col_adm3:
        st.session_state.directorate = st.text_input("المديرية الإقليمية:", value=st.session_state.directorate)

CURRICULUM_DB = {
    "التربية الإسلامية": {
        "components": ["الحكمة", "القسط", "الاستجابة", "الاقتداء", "التزكية (العقيدة)", "التزكية (القرآن الكريم)"],
        "references": {
            "الخامس": ["في رحاب التربية الإسلامية", "واحة التربية الإسلامية", "الممتاز في التربية الإسلامية"],
            "السادس": ["في رحاب التربية الإسلامية", "واحة التربية الإسلامية", "الممتاز في التربية الإسلامية"],
            "الرابع": ["الممتاز في التربية الإسلامية", "واحة التربية الإسلامية", "في رحاب التربية الإسلامية"],
            "الثالث": ["الممتاز في التربية الإسلامية", "واحة التربية الإسلامية"],
            "الثاني": ["في رحاب التربية الإسلامية", "المفيد في التربية الإسلامية"],
            "الأول": ["في رحاب التربية الإسلامية", "المفيد في التربية الإسلامية"]
        }
    },
    "الاجتماعيات": {
        "components": ["التاريخ", "الجغرافيا", "التربية المدنية"],
        "references": {
            "السادس": ["الجديد في الاجتماعيات", "المسار في الاجتماعيات", "في رحاب الاجتماعيات"],
            "الخامس": ["المفيد في الاجتماعيات", "النجاح في الاجتماعيات", "مسار الاجتماعيات"],
            "الرابع": ["الجديد في الاجتماعيات", "المفيد في الاجتماعيات", "الواضح في الاجتماعيات"]
        }
    },
    "اللغة العربية": {
        "components": ["القراءة", "الظواهر اللغوية (تراكيب / صرف وتحويل / إملاء)", "التعبير الكتابي (الإنشاء)", "التواصل الشفهي"],
        "references": {
            "السادس": ["كتابي في اللغة العربية", "في رحاب اللغة العربية", "المنير في اللغة العربية"],
            "الخامس": ["مرشدي في اللغة العربية", "المنير في اللغة العربية", "كتابي في اللغة العربية"],
            "الرابع": ["الواحة في اللغة العربية", "المفيد في اللغة العربية"],
            "الثالث": ["المفيد في اللغة العربية", "مرشدي في اللغة العربية"],
            "الثاني": ["المفيد في اللغة العربية", "كتابي في اللغة العربية"],
            "الأول": ["المفيد في اللغة العربية", "كتابي في اللغة العربية"]
        }
    },
    "النشاط العلمي": {
        "components": ["علوم الحياة والأرض", "العلوم الفيزيائية والتكنولوجية", "الفلك والفضاء"],
        "references": {
            "السادس": ["المنير في النشاط العلمي", "فضاء النشاط العلمي"],
            "الخامس": ["المنير في النشاط العلمي", "الواضح في النشاط العلمي", "المفيد في النشاط العلمي"],
            "الرابع": ["المرشد في النشاط العلمي", "المنهل في النشاط العلمي"],
            "الثالث": ["المنهل في النشاط العلمي", "الواضح في النشاط العلمي"],
            "الثاني": ["الواضح في النشاط العلمي", "فضاء النشاط العلمي"],
            "الأول": ["الواضح في النشاط العلمي", "فضاء النشاط العلمي"]
        }
    },
    "الرياضيات": {
        "components": ["الأعداد والحساب", "الهندسة والفضاء", "القياس", "تنظيم ومعالجة البيانات"],
        "references": {
            "السادس": ["النجاح في الرياضيات", "الجيد في الرياضيات"],
            "الخامس": ["المفيد في الرياضيات", "النجاح في الرياضيات"],
            "الرابع": ["المفيد في الرياضيات", "الجيد في الرياضيات"],
            "الثالث": ["المفيد في الرياضيات", "المرجع في الرياضيات"],
            "الثاني": ["المفيد في الرياضيات", "فضاء الرياضيات"],
            "الأول": ["المفيد في الرياضيات", "فضاء الرياضيات"]
        }
    }
}

st.markdown("#### اختيار محددات الدرس")
col1, col2 = st.columns(2)
with col1:
    level = st.selectbox("المستوى الدراسي:", ["السادس", "الخامس", "الرابع", "الثالث", "الثاني", "الأول"])

valid_subjects = [s for s, d in CURRICULUM_DB.items() if level in d["references"]]

with col2:
    subject = st.selectbox("المادة الدراسية:", valid_subjects)

col3, col4 = st.columns(2)
with col3:
    component = st.selectbox("المكون / المدخل:", CURRICULUM_DB[subject]["components"])

with col4:
    ref_list = CURRICULUM_DB[subject]["references"].get(level, ["مرجع رسمي معتمد"])
    reference = st.selectbox("المرجع المعتمد:", ref_list)

col5, col6, col7 = st.columns(3)
with col5:
    lesson_order = st.number_input("ترتيب الدرس في المقرر (رقم فقط):", min_value=1, max_value=40, value=1)
with col6:
    session_choice = st.selectbox("الحصة:", ["الحصتان معاً (1 و 2)", "الحصة الأولى", "الحصة الثانية"])
with col7:
    week_num = st.number_input("أسبوع السنة المعتمد:", min_value=1, max_value=34, value=2)

st.write("---")

if st.button("توليد الجذاذة المسطرة المطابقة للدليل", type="primary", use_container_width=True):
    if not api_key:
        st.error("يرجى إدخال مفتاح Gemini API في الشريط الجانبي.")
    else:
        with st.spinner(f"جاري نسخ الجذاذة الحرفية للدرس {lesson_order} في {subject} ({reference})..."):
            try:
                client = genai.Client(api_key=api_key)

                prompt = f"""
                أنت خبير تربوي ومفتش بيداغوجي بالتعليم الابتدائي بالمملكة المغربية.
                المهمة: كتابة الجذاذة التربوية الرسمية للدرس رقم {lesson_order} مطابقة تماماً كنسخة طبق الأصل لما هو وارد في دليل الأستاذ وكتاب التلميذ للمرجع {reference} الخاص بـ {level} ابتدائي.

                المعطيات:
                - المستوى: {level} ابتدائي.
                - المادة: {subject}.
                - المكون: {component}.
                - المرجع: {reference}.
                - ترتيب الدرس: الدرس {lesson_order}.
                - الأسبوع: {week_num}.

                الشروط البيداغوجية واللغوية:
                1. استخرج العنوان الرسمي الحرفي للدرس {lesson_order}.
                2. اذكر أهداف التعلم الرسمية بنصها ودون اختصار.
                3. اكتب مراحل وسيرورة الحصة الأولى بدقة متناهية:
                   - التمهيد واستحضار المكتسبات.
                   - الأنشطة: قراءة الوثائق، أرقام الوثائق ونصوصها، الأسئلة الدقيقة الموجهة للتلاميذ، خطوات التفسير والتركيب والاستخلاص.
                4. اكتب مراحل الحصة الثانية: أنشطة تقويم التعلمات، التمارين والأنشطة التطبيقية، وإجراءات الدعم والتثبيت.
                5. ممنوع استخدام عبارات استدراكية مثل: "كما ورد في الكتاب" أو "بحسب الدليل". اكتب المحتوى مباشرة كجذاذة رسمية.

                أخرج الناتج بصيغة JSON صارمة فقط:
                {{
                    "lesson_title": "الدرس {lesson_order} : [عنوان الدرس الدقيق]",
                    "learning_goals": [
                        "الهدف الأول كاملاً",
                        "الهدف الثاني كاملاً"
                    ],
                    "session_1": [
                        {{"step": "تمهيد واستحضار المكتسبات", "activities": "تفاصيل الأسئلة والأنشطة"}},
                        {{"step": "النشاط 1 : [عنوان النشاط]", "activities": "دراسة الوثيقة 1 والأسئلة كاملة بالتفصيل"}},
                        {{"step": "النشاط 2 : [عنوان النشاط]", "activities": "دراسة الوثيقة 2 واستخلاص النتائج"}}
                    ],
                    "session_2": [
                        {{"step": "تقويم المكتسبات", "activities": "الأسئلة التقويمية كاملة"}},
                        {{"step": "أنشطة التثبيت والدعم", "activities": "التمارين والأنشطة التطبيقية المنجزة"}}
                    ]
                }}
                """

                candidate_models = ["gemini-2.5-flash", "gemini-3.6-flash", "gemini-2.0-flash"]
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
                    raise Exception(f"فشل الاتصال: {last_err}")

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
                    "duration": "45 دقيقة" if "فقط" in session_choice or "الأولى" in session_choice or "الثانية" in session_choice else "45 دقيقة × 2",
                    "learning_goals": res_data.get("learning_goals", []),
                    "session_1": s1_data,
                    "session_2": s2_data
                }

                word_buffer = create_word_jodada(jodada_full)

                st.success(f"تم إعداد الجذاذة المسطرة بنجاح: {jodada_full['lesson_title']}")

                st.download_button(
                    label="تحميل الجذاذة الرسمية المسطرة بصيغة Word (.docx)",
                    data=word_buffer,
                    file_name=f"جذاذة_{subject}_{component}_{level}_الدرس_{lesson_order}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True
                )

            except Exception as e:
                st.error(f"حدث خطأ أثناء التوليد: {str(e)}")
