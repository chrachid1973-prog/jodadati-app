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
from google.genai import types
import docx
from docx.shared import Cm, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
import json
import io
import re

st.set_page_config(page_title="منصة الجذاذات التربوية الرسمية المفصلة", layout="wide")

# رسم حدود الجدول بأسلوب مسطر كامل ورسمي
def set_table_borders(table):
    tblPr = table._tbl.tblPr
    borders = parse_xml(r'''
        <w:tblBorders {} >
            <w:top w:val="single" w:sz="6" w:space="0" w:color="000000"/>
            <w:left w:val="single" w:sz="6" w:space="0" w:color="000000"/>
            <w:bottom w:val="single" w:sz="6" w:space="0" w:color="000000"/>
            <w:right w:val="single" w:sz="6" w:space="0" w:color="000000"/>
            <w:insideH w:val="single" w:sz="4" w:space="0" w:color="000000"/>
            <w:insideV w:val="single" w:sz="4" w:space="0" w:color="000000"/>
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

# ضبط الفقرة باتجاه يمين صارم مع RTL
def format_paragraph_rtl(p, align=WD_ALIGN_PARAGRAPH.RIGHT, space_before=2, space_after=2):
    p.alignment = align
    pPr = p._p.get_or_add_pPr()
    pPr.append(parse_xml(r'<w:bidi {} w:val="1"/>'.format(nsdecls('w'))))
    pPr.append(parse_xml(r'<w:jc {} w:val="right"/>'.format(nsdecls('w'))))
    p.paragraph_format.space_before = Pt(space_before)
    p.paragraph_format.space_after = Pt(space_after)

# إضافة النص مع تفعيل الخط العربي ووسوم RTL والمحاذاة
def add_arabic_run(paragraph, text, font_size=11, bold=False):
    run = paragraph.add_run(text)
    run.bold = bold
    run.font.name = 'Traditional Arabic'
    run.font.size = Pt(font_size)
    rPr = run._r.get_or_add_rPr()
    rPr.append(parse_xml(r'<w:rtl {} w:val="1"/>'.format(nsdecls('w'))))
    rPr.append(parse_xml(r'<w:lang {} w:bidi="ar-MA" w:val="ar-MA"/>'.format(nsdecls('w'))))
    return run

# بناء المستند بجداول مسطرة ورأسية كاملة
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

    # جدول الترويسة المسطر بالكامل
    t_head = doc.add_table(rows=5, cols=3)
    set_table_rtl_and_right(t_head)

    # الصف 1: المؤسسة في الأعلى جهة اليمين
    p = t_head.cell(0, 0).paragraphs[0]
    format_paragraph_rtl(p)
    add_arabic_run(p, f"المؤسسة : {data['school']}", bold=True)

    p = t_head.cell(0, 1).paragraphs[0]
    format_paragraph_rtl(p)
    add_arabic_run(p, f"المستوى : {data['level']}", bold=True)

    p = t_head.cell(0, 2).paragraphs[0]
    format_paragraph_rtl(p)
    add_arabic_run(p, f"الأستاذ(ة) : {data['teacher']}", bold=True)

    # الصف 2: المادة والمرجع والمكون
    p = t_head.cell(1, 0).paragraphs[0]
    format_paragraph_rtl(p)
    add_arabic_run(p, f"المادة : {data['subject']}", bold=True)

    p = t_head.cell(1, 1).paragraphs[0]
    format_paragraph_rtl(p)
    add_arabic_run(p, f"المرجع : {data['reference']}", bold=True)

    p = t_head.cell(1, 2).paragraphs[0]
    format_paragraph_rtl(p)
    add_arabic_run(p, f"المكون : {data['component']}", bold=True)

    # الصف 3: عنوان الدرس المباشر
    c_lesson = t_head.cell(2, 0).merge(t_head.cell(2, 1))
    p = c_lesson.paragraphs[0]
    format_paragraph_rtl(p)
    add_arabic_run(p, f"الدرس {data['jodada_num']} : {data['lesson_title']}", bold=True)

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

    # الصف 5: الأهداف دون اختصار
    c_goals = t_head.cell(4, 0).merge(t_head.cell(4, 2))
    p_title = c_goals.paragraphs[0]
    format_paragraph_rtl(p_title, space_before=1, space_after=1)
    add_arabic_run(p_title, "أهداف التعلم المسطرة في الدليل :", bold=True, font_size=11.5)

    for goal in data['learning_goals']:
        p_g = c_goals.add_paragraph()
        format_paragraph_rtl(p_g, space_before=1, space_after=1)
        add_arabic_run(p_g, f"• {goal}", font_size=11)

    for row in t_head.rows:
        for cell in row.cells:
            format_cell(cell)

    doc.add_paragraph().paragraph_format.space_after = Pt(2)

    # بناء جداول الحصص التفصيلية مع دعم الجداول الفرعية
    def build_session_table(session_title, steps):
        t = doc.add_table(rows=1, cols=2)
        set_table_rtl_and_right(t)

        hdr = t.cell(0, 0).merge(t.cell(0, 1))
        format_cell(hdr, "E8E8E8")
        p_hdr = hdr.paragraphs[0]
        format_paragraph_rtl(p_hdr, align=WD_ALIGN_PARAGRAPH.CENTER)
        add_arabic_run(p_hdr, session_title, font_size=12.5, bold=True)

        cols_row = t.add_row()
        p0 = cols_row.cells[0].paragraphs[0]
        format_paragraph_rtl(p0, align=WD_ALIGN_PARAGRAPH.CENTER)
        add_arabic_run(p0, "المراحل / خطوات الدرس", font_size=11.5, bold=True)

        p1 = cols_row.cells[1].paragraphs[0]
        format_paragraph_rtl(p1, align=WD_ALIGN_PARAGRAPH.CENTER)
        add_arabic_run(p1, "الأنشطة والسيناريو الديداكتيكي الكامل", font_size=11.5, bold=True)

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
            add_arabic_run(p_step, item.get("step", ""), font_size=11, bold=True)

            p_act = c1.paragraphs[0]
            act_text = item.get("activities", "")
            lines = [l.strip() for l in act_text.split("\n") if l.strip()]

            for i, line in enumerate(lines):
                if i == 0:
                    format_paragraph_rtl(p_act, align=WD_ALIGN_PARAGRAPH.RIGHT, space_before=1, space_after=1)
                    add_arabic_run(p_act, line, font_size=11)
                else:
                    p_new = c1.add_paragraph()
                    format_paragraph_rtl(p_new, align=WD_ALIGN_PARAGRAPH.RIGHT, space_before=1, space_after=1)
                    add_arabic_run(p_new, line, font_size=11)

            sub_table_data = item.get("table_data", None)
            if sub_table_data and isinstance(sub_table_data, list) and len(sub_table_data) > 0:
                sub_rows = len(sub_table_data)
                sub_cols = len(sub_table_data[0]) if sub_rows > 0 else 0
                if sub_cols > 0:
                    st_p = c1.add_paragraph()
                    format_paragraph_rtl(st_p, space_before=2, space_after=2)
                    
                    sub_t = c1.add_table(rows=sub_rows, cols=sub_cols)
                    set_table_rtl_and_right(sub_t)
                    
                    for r_idx, row_content in enumerate(sub_table_data):
                        for c_idx, val in enumerate(row_content):
                            sc = sub_t.cell(r_idx, c_idx)
                            fill_color = "EEEEEE" if r_idx == 0 else None
                            format_cell(sc, fill_color)
                            p_sc = sc.paragraphs[0]
                            format_paragraph_rtl(p_sc, align=WD_ALIGN_PARAGRAPH.CENTER if r_idx == 0 else WD_ALIGN_PARAGRAPH.RIGHT)
                            add_arabic_run(p_sc, str(val), font_size=10, bold=(r_idx == 0))

        for r in t.rows:
            r.cells[0].width = Cm(4.5)
            r.cells[1].width = Cm(14.5)

        doc.add_paragraph().paragraph_format.space_after = Pt(4)

    if data.get("session_1"):
        build_session_table("الحصة الأولى : بناء التعلمات (45 دقيقة)", data["session_1"])
    if data.get("session_2"):
        build_session_table("الحصة الثانية : تقويم ودعم وتثبيت (45 دقيقة)", data["session_2"])

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

st.markdown("<h2 style='text-align: right; direction: rtl; color: #1E3A8A;'>منصة الجذاذات التربوية الرسمية الشاملة</h2>", unsafe_allow_html=True)
st.markdown("<p style='text-align: right; direction: rtl;'>توليد جذاذات رسمية دقيقة مطابقة للدليل باستخدام نماذج Gemini 3.8 / 3.6</p>", unsafe_allow_html=True)
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

with st.expander("البيانات الإدارية (تُحفظ تلقائياً في أعلى الترويسة)", expanded=True):
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
            "الخامس": ["المفيد في الاجتماعيات", "النجاح في الاجتماعيات", "مسار الاجتماعيات"],
            "السادس": ["الجديد في الاجتماعيات", "المسار في الاجتماعيات", "في رحاب الاجتماعيات"],
            "الرابع": ["الجديد في الاجتماعيات", "المفيد في الاجتماعيات", "الواضح في الاجتماعيات"]
        }
    },
    "اللغة العربية": {
        "components": ["القراءة", "الظواهر اللغوية (تراكيب / صرف وتحويل / إملاء)", "التعبير الكتابي (الإنشاء)", "التواصل الشفهي"],
        "references": {
            "الخامس": ["مرشدي في اللغة العربية", "المنير في اللغة العربية", "كتابي في اللغة العربية"],
            "السادس": ["كتابي في اللغة العربية", "في رحاب اللغة العربية", "المنير في اللغة العربية"],
            "الرابع": ["الواحة في اللغة العربية", "المفيد في اللغة العربية"],
            "الثالث": ["المفيد في اللغة العربية", "مرشدي في اللغة العربية"],
            "الثاني": ["المفيد في اللغة العربية", "كتابي في اللغة العربية"],
            "الأول": ["المفيد في اللغة العربية", "كتابي في اللغة العربية"]
        }
    },
    "النشاط العلمي": {
        "components": ["علوم الحياة والأرض", "العلوم الفيزيائية والتكنولوجية", "الفلك والفضاء"],
        "references": {
            "الخامس": ["المنير في النشاط العلمي", "الواضح في النشاط العلمي", "المفيد في النشاط العلمي"],
            "السادس": ["المنير في النشاط العلمي", "فضاء النشاط العلمي"],
            "الرابع": ["المرشد في النشاط العلمي", "المنهل في النشاط العلمي"],
            "الثالث": ["المنهل في النشاط العلمي", "الواضح في النشاط العلمي"],
            "الثاني": ["الواضح في النشاط العلمي", "فضاء النشاط العلمي"],
            "الأول": ["الواضح في النشاط العلمي", "فضاء النشاط العلمي"]
        }
    },
    "الرياضيات": {
        "components": ["الأعداد والحساب", "الهندسة والفضاء", "القياس", "تنظيم ومعالجة البيانات"],
        "references": {
            "الخامس": ["المفيد في الرياضيات", "النجاح في الرياضيات"],
            "السادس": ["النجاح في الرياضيات", "الجيد في الرياضيات"],
            "الرابع": ["المفيد في الرياضيات", "الجيد في الرياضيات"],
            "الثالث": ["المفيد في الرياضيات", "المرجع في الرياضيات"],
            "الثاني": ["المفيد في الرياضيات", "فضاء الرياضيات"],
            "الأول": ["المفيد في الرياضيات", "فضاء الرياضيات"]
        }
    }
}

OFFICIAL_SYLLABUS = {
    "الخامس": {
        "الاجتماعيات": {
            "التاريخ": {
                1: "التاريخ والمؤرخ",
                2: "عصور ما قبل التاريخ والعصور التاريخية",
                3: "المغرب ما قبل التاريخ: الإنسان ونمط عيشه",
                4: "المغرب القديم: التأثير الحضاري المتبادل بين الفينيقيين والأمازيغ",
                5: "المغرب القديم: الاحتلال الروماني والمقاومة الأمازيغية",
                6: "المغرب القديم: مظاهر من الحضارة الأمازيغية",
                7: "قيام الدولة الإدريسية وانتشار الإسلام في المغرب",
                8: "المرابطون: توحيد البلاد وامتداد الدولة",
                9: "الدولة الموحدية: توحيد الغرب الإسلامي",
                10: "الدولة المرينية: إبداع حضاري",
                11: "الدولة السعدية: ازدهار اقتصادي",
                12: "الدولة العلوية: توحيد البلاد وبناء الدولة"
            },
            "الجغرافيا": {
                1: "مميزات وطني: الخريطة والحدود والمدن",
                2: "مميزات وطني: الجبال والهضاب والسهول",
                3: "مميزات وطني: رسم خريطة تضاريس وطني",
                4: "مميزات مناخ وطني: الحرارة والتساقطات وتمثيلها",
                5: "مميزات وطني: الأنهار والسدود",
                6: "مميزات وطني: الفلاحة والثروة النباتية والحيوانية",
                7: "مميزات وطني: الواجهتان البحريتان",
                8: "السكان في وطني: التوزع والبنية",
                9: "السياحة: رافعة لتنمية وطني",
                10: "مميزات وطني: المعادن والصناعة المعدنية (الفوسفاط نموذجاً)",
                11: "مميزات وطني: المواصلات",
                12: "مشاكل بيئية بوطني: التصحر والتلوث ونقص الماء"
            },
            "التربية المدنية": {
                1: "أنظم عملي وأقيم أدائي",
                2: "أشارك في وضع ميثاق القسم",
                3: "أشارك في تكوين تعاونية القسم وأنشط فيها",
                4: "أنمي قدراتي على التعلم باستقلالية",
                5: "أستفيد من الخدمات العمومية في محيطي وأحافظ عليها",
                6: "أحافظ على صحتي وسلامتي",
                7: "أحمي نفسي من أخطار التدخين",
                8: "أعي مزايا وأخطار الإنترنت",
                9: "حقي في عدم التعرض لاعتداء ودوري في حماية نفسي",
                10: "أحترم القانون في استعمال الطريق",
                11: "الإحساس بالآخر ضمانة للاحترام المتبادل",
                12: "التسامح سلوك يعزز العيش المشترك"
            }
        }
    }
}

st.markdown("#### اختيار محددات الدرس")
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
    ref_list = CURRICULUM_DB[subject]["references"].get(level, ["مرجع رسمي معتمد"])
    reference = st.selectbox("المرجع المعتمد:", ref_list)

col5, col6, col7 = st.columns(3)
with col5:
    lesson_order = st.number_input("ترتيب الدرس في المقرر (رقم فقط):", min_value=1, max_value=40, value=12)
with col6:
    session_choice = st.selectbox("الحصة المطلوبة:", ["الحصتان معاً (1 و 2)", "الحصة الأولى", "الحصة الثانية"])
with col7:
    week_num = st.number_input("أسبوع السنة المعتمد:", min_value=1, max_value=34, value=28)

st.write("---")

if st.button("توليد الجذاذة الكاملة غير المختصرة المطابقة للدليل", type="primary", use_container_width=True):
    if not api_key:
        st.error("يرجى إدخال مفتاح Gemini API في الشريط الجانبي.")
    else:
        exact_title = OFFICIAL_SYLLABUS.get(level, {}).get(subject, {}).get(component, {}).get(lesson_order, None)
        title_hint = f"العنوان الرسمي الحقيقي للدرس {lesson_order} في فهرس هذا المقرر هو: '{exact_title}'." if exact_title else ""

        with st.spinner(f"جاري استخراج الجذاذة الحرفية بدقة متناهية عبر محرك Gemini 3.8 / 3.6..."):
            try:
                client = genai.Client(api_key=api_key)

                prompt = f"""
                أنت مفتش بيداغوجي مقتدر بالتعليم الابتدائي بالمملكة المغربية.
                المهمة الصارمة: استخراج ونقل الجذاذة التربوية الرسمية للدرس رقم {lesson_order} كنسخة طبق الأصل بنسبة 100% كما هي مدونة في "دليل الأستاذ" و"كتاب التلميذ" للمرجع {reference} (الطبعة المنقحة الحديثة).

                المعطيات المحددة:
                - المستوى الدراسي: {level} ابتدائي.
                - المادة: {subject}.
                - المرجع الدراسي: {reference}.
                - المكون: {component}.
                - ترتيب الدرس: الدرس {lesson_order}.
                - الأسبوع: {week_num}.
                {title_hint}

                شروط وقواعد النقل البيداغوجي الصارمة:
                1. عنوان الدرس: انقل العنوان الرسمي الصحيح للدرس رقم {lesson_order}.
                2. أهداف التعلم: انقل جميع الأهداف المسطرة في ترويسة الدرس بالدليل كاملة كلمة بكلمة.
                3. الحصة الأولى (45 دقيقة):
                   - التمهيد واستحضار المكتسبات: اكتب الوضعية والأسئلة كاملة.
                   - الأنشطة التعلمية: اكتب كل نشاط باسمه الرسمي، مع نقل الوثائق كاملة (أرقامها، أنواعها، ونصوصها)، وكل الأسئلة الموجهة للتلاميذ والمهام المطروحة، والأجوبة أو الاستنتاجات المرحلية المقابلة لها دون أي تلخيص.
                   - إذا ورد في الدليل جدول لتحليل وثائق أو مقارنة أو خلاصة، انقله بالكامل في حقل table_data.
                4. الحصة الثانية (45 دقيقة):
                   - أنشطة التقويم والدعم: انقل الأسئلة الشفهية، نصوص التوضيحات والوضعيات التقويمية المكتوبة على الدفاتر، وأنشطة الدعم والمعالجة بدقة.
                   - إذا ورد جدول لشبكة تقويم أو تمارين جدولية، انقله بالكامل في table_data.
                5. اكتب المحتوى مباشرة بصيغة الجذاذة المعتمدة دون أي عبارات استدراكية.

                أخرج الناتج بصيغة JSON صارمة جداً بالهيكل التالي:
                {{
                    "lesson_title": "[العنوان الحقيقي للدرس فقط]",
                    "learning_goals": [
                        "الهدف الأول كاملاً بالحرف",
                        "الهدف الثاني كاملاً بالحرف"
                    ],
                    "session_1": [
                        {{
                            "step": "اسم المرحلة أو النشاط",
                            "activities": "كل النصوص والأسئلة والسيناريو الديداكتيكي كاملاً دون اختصار سطر بسطر",
                            "table_data": null
                        }}
                    ],
                    "session_2": [
                        {{
                            "step": "اسم مرحلة التقويم أو الدعم",
                            "activities": "الأسئلة والأنشطة والمهام كاملة بالتفصيل",
                            "table_data": null
                        }}
                    ]
                }}
                ملاحظة: إذا لم يتضمن النشاط جدولاً، ضع في حقل "table_data" القيمة null. أما إذا كان يحتوي على جدول فانقله كمصفوفة مصفوفات نصوص.
                """

                # فرض إخراج JSON نقي ومنع أي خطأ في التحليل
                config = types.GenerateContentConfig(
                    temperature=0.2,
                    max_output_tokens=8192,
                    response_mime_type="application/json"
                )

                candidate_models = [
                    "gemini-3.8-flash",
                    "gemini-3.7-flash",
                    "gemini-3.6-flash",
                    "gemini-2.5-flash"
                ]

                response = None
                last_err = None

                for m in candidate_models:
                    try:
                        response = client.models.generate_content(
                            model=m,
                            contents=prompt,
                            config=config
                        )
                        if response and response.text:
                            break
                    except Exception as e:
                        last_err = e
                        continue

                if response is None or not response.text:
                    raise Exception(f"فشل الاتصال بالنماذج: {last_err}")

                raw_text = response.text.strip()
                
                # استخراج JSON بأمان حتى لو أحاطه بوسوم كود
                if "```" in raw_text:
                    match = re.search(r'```(?:json)?\s*([\s\S]*?)\s*```', raw_text)
                    if match:
                        raw_text = match.group(1).strip()

                res_data = json.loads(raw_text)

                s1_data = res_data.get("session_1", []) if "الأولى" in session_choice or "معاً" in session_choice else []
                s2_data = res_data.get("session_2", []) if "الثانية" in session_choice or "معاً" in session_choice else []

                clean_title = res_data.get("lesson_title", "")
                for prefix in ["الدرس :", "الدرس", f"{lesson_order}", ":"]:
                    if clean_title.startswith(prefix):
                        clean_title = clean_title.replace(prefix, "").strip()

                if exact_title:
                    clean_title = exact_title

                jodada_full = {
                    "level": level,
                    "subject": subject,
                    "teacher": st.session_state.teacher_name,
                    "school": st.session_state.school_name,
                    "unit": (lesson_order - 1) // 4 + 1,
                    "reference": reference,
                    "component": component,
                    "lesson_title": clean_title,
                    "jodada_num": lesson_order,
                    "week": week_num,
                    "session": session_choice,
                    "duration": "45 دقيقة" if "فقط" in session_choice or "الأولى" in session_choice or "الثانية" in session_choice else "45 دقيقة × 2",
                    "learning_goals": res_data.get("learning_goals", []),
                    "session_1": s1_data,
                    "session_2": s2_data
                }

                word_buffer = create_word_jodada(jodada_full)

                st.success(f"تم بنجاح إعداد الجذاذة الكاملة: الدرس {lesson_order} : {jodada_full['lesson_title']}")

                st.download_button(
                    label="تحميل الجذاذة الرسمية المسطرة الشاملة بصيغة Word (.docx)",
                    data=word_buffer,
                    file_name=f"جذاذة_كاملة_{subject}_{component}_{level}_الدرس_{lesson_order}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True
                )

            except Exception as e:
                st.error(f"حدث خطأ أثناء التوليد: {str(e)}")
