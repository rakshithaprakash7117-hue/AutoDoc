import streamlit as st
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from io import BytesIO
from PIL import Image

def improve_description(notes, department, event_title, participants):
    notes = notes.strip()

    if not notes:
        return ""

    improved = (
        f"The Department of {department} conducted "
        f"'{event_title}' as an academic activity for students. "
    )

    if participants > 0:
        improved += (
            f"A total of {participants} participants attended the programme. "
        )

    improved += (
        f"The programme focused on {notes}. "
        f"The session provided participants with practical exposure, "
        f"enhanced their understanding of the topic, and encouraged "
        f"active participation and interaction."
    )

    return improved

def create_report(
        faculty_name,
        department,
        activity_type,
        event_title,
        event_date,
        venue,
        participants,
        description,
        photos=None
):

    doc = Document()

    # ---------------- PAGE SETTINGS ----------------

    section = doc.sections[0]

    section.top_margin = Inches(0.5)
    section.bottom_margin = Inches(0.5)
    section.left_margin = Inches(0.65)
    section.right_margin = Inches(0.65)

    # ---------------- DEFAULT FONT ----------------

    style = doc.styles["Normal"]
    style.font.name = "Times New Roman"
    style.font.size = Pt(11)

    # ---------------- CHRIST LOGO ----------------

    logo = doc.add_paragraph()
    logo.alignment = WD_ALIGN_PARAGRAPH.CENTER

    logo_run = logo.add_run()
    logo_run.add_picture("logo.png", width=Inches(4.3))

    # ---------------- DIVIDER ----------------

    divider = doc.add_paragraph()
    divider.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # ---------------- EVENT REPORT ----------------

    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER

    run = title.add_run("EVENT REPORT")
    run.bold = True
    run.font.name = "Times New Roman"
    run.font.size = Pt(18)

    # ---------------- DEPARTMENT ----------------

    dept = doc.add_paragraph()
    dept.alignment = WD_ALIGN_PARAGRAPH.CENTER

    run = dept.add_run(f"Department of {department}")
    run.bold = True
    run.font.size = Pt(13)

    # ---------------- DETAILS TABLE ----------------

    table = doc.add_table(rows=0, cols=2)

    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER

    details = [
        ("Title of Activity", event_title),
        ("Activity Type", activity_type),
        ("Date", event_date.strftime("%d %B %Y")),
        ("Venue", venue),
        ("Number of Participants", str(participants)),
        ("Faculty Coordinator", faculty_name)
    ]

    for label, value in details:

        cells = table.add_row().cells

        cells[0].text = label
        cells[1].text = value

        cells[0].vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
        cells[1].vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER

        # Bold first column
        for run in cells[0].paragraphs[0].runs:
            run.bold = True
            run.font.name = "Times New Roman"
            run.font.size = Pt(10)

        for run in cells[1].paragraphs[0].runs:
            run.font.name = "Times New Roman"
            run.font.size = Pt(10)

    # ---------------- DESCRIPTION ----------------

    doc.add_paragraph()

    heading = doc.add_paragraph()

    heading_run = heading.add_run("Description")
    heading_run.bold = True
    heading_run.font.name = "Times New Roman"
    heading_run.font.size = Pt(12)

    desc = doc.add_paragraph(description)

    desc.alignment = WD_ALIGN_PARAGRAPH.LEFT
    desc.paragraph_format.line_spacing = 1.15
    desc.paragraph_format.space_after = Pt(8)

    for desc_run in desc.runs:
        desc_run.font.name = "Times New Roman"
        desc_run.font.size = Pt(11)

    # ---------------- PHOTOGRAPHS ----------------

    if photos:

        photo_heading = doc.add_paragraph()

        photo_heading_run = photo_heading.add_run("Event Photographs")
        photo_heading_run.bold = True
        photo_heading_run.font.name = "Times New Roman"
        photo_heading_run.font.size = Pt(12)

        photo_table = doc.add_table(rows=1, cols=2)

        for index, photo in enumerate(photos[:2]):

            # Convert uploaded image into a clean JPEG
            photo.seek(0)

            image = Image.open(photo)

            # Convert images such as PNG/RGBA to RGB
            if image.mode != "RGB":
                image = image.convert("RGB")

            image_buffer = BytesIO()

            image.save(
                image_buffer,
                format="JPEG",
                quality=90
            )

            image_buffer.seek(0)

            # Insert into Word
            cell = photo_table.rows[0].cells[index]

            paragraph = cell.paragraphs[0]
            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER

            picture_run = paragraph.add_run()

            picture_run.add_picture(
                image_buffer,
                width=Inches(3.0)
            )

            # Caption
            caption = cell.add_paragraph()
            caption.alignment = WD_ALIGN_PARAGRAPH.CENTER

            caption_run = caption.add_run(
                f"Photograph {index + 1}"
            )

            caption_run.italic = True
            caption_run.font.name = "Times New Roman"
            caption_run.font.size = Pt(9)

    # ---------------- FOOTER ----------------

    footer = section.footer

    footer_para = footer.paragraphs[0]
    footer_para.alignment = WD_ALIGN_PARAGRAPH.CENTER

    footer_run = footer_para.add_run(
        "Excellence and Service"
    )

    footer_run.italic = True
    footer_run.font.name = "Times New Roman"
    footer_run.font.size = Pt(10)

    # ---------------- SAVE ----------------

    buffer = BytesIO()

    doc.save(buffer)

    buffer.seek(0)

    return buffer

st.title("AutoDoc")
st.subheader("Faculty Documentation Assistant")

st.write("Enter the activity details below.")

faculty_name = st.text_input("Faculty Name")

department = st.selectbox(
    "Department",
    [
        "Computer Science",
        "Commerce and Management",
        "Food Science and Technology",
        "Other"
    ]
)

activity_type = st.selectbox(
    "Activity Type",
    [
        "Workshop",
        "FDP",
        "Conference",
        "Seminar",
        "Research Publication",
        "Student Activity",
        "Other"
    ]
)

event_title = st.text_input("Title of Activity")

event_date = st.date_input("Date")

venue = st.text_input("Venue")

participants = st.number_input(
    "Number of Participants",
    min_value=0
)

# ---------- DESCRIPTION / AI IMPROVEMENT ----------

if "description" not in st.session_state:
    st.session_state.description = ""

def improve_text():
    notes = st.session_state.description

    if notes.strip():
        st.session_state.description = improve_description(
            notes,
            department,
            event_title,
            participants
        )

description = st.text_area(
    "Description / Rough Notes",
    key="description",
    height=150,
    placeholder="Example: hands-on prompt engineering, large language models, students tried different prompts"
)

st.button(
    "✨ Improve Description",
    on_click=improve_text
)

# Always take the latest value
description = st.session_state.description


photos = st.file_uploader(
    "Upload Event Photograph(s)",
    type=["jpg", "jpeg", "png"],
    accept_multiple_files=True
)

if st.button("Generate Documentation"):
    st.success("Activity submitted successfully!")

    st.write("### Activity Preview")

    st.write("**Faculty:**", faculty_name)
    st.write("**Department:**", department)
    st.write("**Activity:**", activity_type)
    st.write("**Title:**", event_title)
    st.write("**Date:**", event_date)
    st.write("**Venue:**", venue)
    st.write("**Participants:**", participants)
    st.write("**Description:**", description)

    report = create_report(
        faculty_name,
        department,
        activity_type,
        event_title,
        event_date,
        venue,
        participants,
        description,
        photos
    )

    safe_title = event_title.replace(" ", "_")

    filename = (
        f"{safe_title}_"
        f"{event_date.strftime('%d-%m-%Y')}.docx"
    )

    st.success("Report generated successfully!")

    st.download_button(
        label="Download Report",
        data=report,
        file_name=filename,
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )

    if photos:
        st.write("### Event Photographs")

        for photo in photos:
            st.image(photo, width=300)