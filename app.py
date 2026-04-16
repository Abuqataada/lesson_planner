import os
import io
import json
from typing import Optional
from fastapi import FastAPI, HTTPException, Request
from fastapi.responses import HTMLResponse, Response
from pydantic import BaseModel
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from datetime import datetime

# Your AI service (keep as is)
from ai_analysis_service import AIAnalysisService

app = FastAPI()
ai_service = AIAnalysisService()

class LessonPlanRequest(BaseModel):
    class_name: str
    subject: str
    topic: str

# ------------------------------------------------------------
# Word document generator (unchanged)
# ------------------------------------------------------------
def set_landscape(doc):
    section = doc.sections[0]
    section.page_width = Inches(11)
    section.page_height = Inches(8.5)
    section.top_margin = Inches(0.5)
    section.bottom_margin = Inches(0.5)
    section.left_margin = Inches(0.5)
    section.right_margin = Inches(0.5)

def add_logo(doc, logo_paths):
    for logo_file in logo_paths:
        if os.path.exists(logo_file):
            para = doc.add_paragraph()
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = para.add_run()
            run.add_picture(logo_file, width=Inches(1.5))
            return True
    para = doc.add_paragraph()
    para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = para.add_run("[SCHOOL LOGO]")
    run.font.size = Pt(10)
    run.italic = True
    return False

def create_lesson_plan_doc(plan_data: dict, teacher_name: str = "ISAH YUSUF") -> bytes:
    doc = Document()
    set_landscape(doc)

    possible_logos = ["logo.png", "logo.jpg", "arndale_logo.png", "school_logo.jpg"]
    add_logo(doc, possible_logos)

    heading = doc.add_heading("Lesson Plan", level=1)
    heading.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # 3‑column header table
    table_main = doc.add_table(rows=9, cols=3)
    table_main.style = 'Table Grid'
    table_main.columns[0].width = Inches(0.6)

    rows_data = [
        ("Class:", plan_data.get("class", "Year 7")),
        ("Subject:", plan_data.get("subject", "Physics")),
        ("Topic:", plan_data.get("topic", "Introduction to Physics")),
        ("Subtopic", plan_data.get("subtopic", plan_data.get("topic", ""))),
        ("Date", plan_data.get("date", datetime.now().strftime("%d – %B, %Y"))),
        ("Week", plan_data.get("week", "Two")),
        ("Duration", plan_data.get("duration", "Forty Minutes")),
        ("Student Age Group", plan_data.get("age_group", "11 – 12 Years")),
        ("INSTRUCTIONAL RESOURCES:", "")
    ]

    for i, (col1, col2) in enumerate(rows_data):
        table_main.cell(i, 0).text = col1
        if i == 8:
            cell2 = table_main.cell(i, 1)
            cell2.text = ''
            resources = plan_data.get("instructional_resources", [])
            if isinstance(resources, list):
                for res in resources:
                    p = cell2.add_paragraph(f"• {res}")
                    p.runs[0].font.size = Pt(11)
            else:
                p = cell2.add_paragraph(f"• {resources}")
                p.runs[0].font.size = Pt(11)
        else:
            table_main.cell(i, 1).text = str(col2)

    # Merge column 3 for Learning Objectives
    col3_cell = table_main.cell(0, 2)
    for r in range(1, 9):
        col3_cell.merge(table_main.cell(r, 2))
    col3_cell.text = ''

    p_title = col3_cell.add_paragraph("Learning Objectives (Differentiated)")
    p_title.runs[0].bold = True
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER

    lo = plan_data.get("learning_objectives", {})
    col3_cell.add_paragraph("Basic Objective (for struggling learners):").runs[0].bold = True
    col3_cell.add_paragraph("By the end of the lesson, students will be able to:")
    p_basic = col3_cell.add_paragraph(f"• {lo.get('basic', 'Define Physics as the study of matter, energy, and their interactions.')}")
    p_basic.paragraph_format.left_indent = Pt(36)

    col3_cell.add_paragraph("Intermediate Objective (for most students):").runs[0].bold = True
    col3_cell.add_paragraph("By the end of the lesson, students will be able to:")
    p_inter = col3_cell.add_paragraph(f"• {lo.get('intermediate', 'Explain the scope of Physics and its relationship with other sciences and technology.')}")
    p_inter.paragraph_format.left_indent = Pt(36)

    col3_cell.add_paragraph("Advanced Objective (for high-achieving students):").runs[0].bold = True
    col3_cell.add_paragraph("By the end of the lesson, students will be able to:")
    p_adv = col3_cell.add_paragraph(f"• {lo.get('advanced', 'Analyse how a modern technology (e.g., a smartphone, GPS) is a practical application of multiple principles from different branches of Physics.')}")
    p_adv.paragraph_format.left_indent = Pt(36)

    doc.add_paragraph()

    # Main development table (6 columns)
    table_dev = doc.add_table(rows=3, cols=6)
    table_dev.style = 'Table Grid'

    prior_cell = table_dev.cell(0, 0)
    prior_cell.merge(table_dev.cell(1, 0))
    prior_cell.text = "Prior Knowledge"
    prior_cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

    warmup_cell = table_dev.cell(0, 1)
    warmup_cell.merge(table_dev.cell(1, 1))
    warmup_cell.text = "Warm-up Activity"
    warmup_cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

    note_cell = table_dev.cell(0, 2)
    note_cell.merge(table_dev.cell(1, 2))
    note_cell.text = "Summarised Learning Note"
    note_cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

    assess_cell = table_dev.cell(0, 5)
    assess_cell.merge(table_dev.cell(1, 5))
    assess_cell.text = "Assessment/Evaluation"
    assess_cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

    learning_cell = table_dev.cell(0, 3)
    learning_cell.merge(table_dev.cell(0, 4))
    learning_cell.text = "Learning Activities"
    learning_cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

    table_dev.cell(1, 3).text = "TEACHER'S ACTIVITIES"
    table_dev.cell(1, 3).paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    table_dev.cell(1, 4).text = "STUDENTS' ACTIVITIES"
    table_dev.cell(1, 4).paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

    table_dev.cell(2, 0).text = plan_data.get("prior_knowledge", "General knowledge from previous lessons.")
    table_dev.cell(2, 1).text = plan_data.get("warmup_activity", "Engaging starter to capture interest.")
    table_dev.cell(2, 2).text = plan_data.get("learning_note", "Core content with definitions and examples.")
    table_dev.cell(2, 3).text = plan_data.get("teacher_activities", "Teacher's actions during the lesson.")
    table_dev.cell(2, 4).text = plan_data.get("student_activities", "Group work, individual tasks, discussions.")
    table_dev.cell(2, 5).text = plan_data.get("assessment", "Formative assessment methods.")

    doc.add_paragraph()

    # Plenary / Homework / Flip Ticket
    table_plenary = doc.add_table(rows=3, cols=2)
    table_plenary.style = 'Table Grid'
    plenary_data = [
        ("Plenary", plan_data.get("plenary", "Summarise key points.")),
        ("Home-Work", plan_data.get("homework", "Reinforcement task.")),
        ("Flip Ticket (next Topic)", plan_data.get("flip_ticket", "Preview of next lesson."))
    ]
    for i, (label, text) in enumerate(plenary_data):
        table_plenary.cell(i, 0).text = label
        table_plenary.cell(i, 1).text = text

    doc.add_paragraph()

    # Signature boxes (side by side)
    sig_row = doc.add_table(rows=1, cols=2)
    sig_row.style = 'Table Grid'
    left_cell = sig_row.cell(0, 0)
    p_teacher = left_cell.paragraphs[0]
    p_teacher.text = "Teacher's Name: "
    run = p_teacher.add_run(teacher_name)
    run.italic = True
    left_cell.add_paragraph("Supervising Officer's Signature: ____________________")
    right_cell = sig_row.cell(0, 1)
    right_cell.text = (
        "Supervising officer's Comment:\n"
        "…………………………………………………………\n"
        "…………………………………………………………\n"
        "…………………………………………………………\n"
        "…………………………………………………………"
    )

    byte_io = io.BytesIO()
    doc.save(byte_io)
    byte_io.seek(0)
    return byte_io.getvalue()

# ------------------------------------------------------------
# Routes
# ------------------------------------------------------------
@app.get("/", response_class=HTMLResponse)
async def form():
    # Direct HTML string – no templates, no Jinja2
    return HTMLResponse("""
    <!DOCTYPE html>
    <html lang="en">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>AI Lesson Plan Generator</title>
        <style>
            body {
                font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
                background: #f0f2f5;
                margin: 0;
                padding: 2rem;
                display: flex;
                justify-content: center;
                align-items: center;
                min-height: 100vh;
            }
            .container {
                background: white;
                padding: 2rem;
                border-radius: 16px;
                box-shadow: 0 10px 25px rgba(0,0,0,0.1);
                max-width: 500px;
                width: 100%;
            }
            h1 {
                text-align: center;
                color: #1e3c72;
                margin-bottom: 1.5rem;
            }
            label {
                display: block;
                margin-top: 1rem;
                font-weight: 600;
                color: #333;
            }
            input {
                width: 100%;
                padding: 0.75rem;
                margin-top: 0.25rem;
                border: 1px solid #ccc;
                border-radius: 8px;
                font-size: 1rem;
                box-sizing: border-box;
            }
            button {
                width: 100%;
                padding: 0.75rem;
                margin-top: 1.5rem;
                background-color: #1e3c72;
                color: white;
                border: none;
                border-radius: 8px;
                font-size: 1rem;
                font-weight: bold;
                cursor: pointer;
                transition: background 0.3s;
            }
            button:hover {
                background-color: #0f2b4f;
            }
            .loading {
                display: none;
                text-align: center;
                margin-top: 1rem;
                color: #1e3c72;
            }
            .error {
                color: red;
                margin-top: 1rem;
                text-align: center;
            }
        </style>
    </head>
    <body>
    <div class="container">
        <h1>📘 AI Lesson Plan Generator</h1>
        <form id="lessonForm">
            <label for="class_name">Class Level</label>
            <input type="text" id="class_name" name="class_name" placeholder="e.g., Year 7, Grade 5, JSS 1" required>

            <label for="subject">Subject</label>
            <input type="text" id="subject" name="subject" placeholder="e.g., Physics, Mathematics, English" required>

            <label for="topic">Topic</label>
            <input type="text" id="topic" name="topic" placeholder="e.g., Introduction to Physics, Forces, Fractions" required>

            <button type="submit">✨ Generate Lesson Plan</button>
        </form>
        <div class="loading" id="loading">⏳ Generating lesson plan... (may take 10-20 seconds)</div>
        <div class="error" id="error"></div>
    </div>

    <script>
        const form = document.getElementById('lessonForm');
        const loadingDiv = document.getElementById('loading');
        const errorDiv = document.getElementById('error');

        form.addEventListener('submit', async (e) => {
            e.preventDefault();
            loadingDiv.style.display = 'block';
            errorDiv.textContent = '';

            const formData = new FormData(form);
            const payload = {
                class_name: formData.get('class_name'),
                subject: formData.get('subject'),
                topic: formData.get('topic')
            };

            try {
                const response = await fetch('/generate', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify(payload)
                });

                if (!response.ok) {
                    const errorText = await response.text();
                    throw new Error(errorText || 'Failed to generate plan');
                }

                const blob = await response.blob();
                const url = window.URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.href = url;
                a.download = `Lesson_Plan_${payload.subject}_${payload.topic}.docx`;
                document.body.appendChild(a);
                a.click();
                a.remove();
                window.URL.revokeObjectURL(url);
            } catch (err) {
                errorDiv.textContent = `Error: ${err.message}`;
            } finally {
                loadingDiv.style.display = 'none';
            }
        });
    </script>
    </body>
    </html>
    """)

@app.post("/generate")
async def generate_plan(request: LessonPlanRequest):
    try:
        # Generate lesson plan content using AI
        plan_data = ai_service.generate_lesson_plan(
            subject=request.subject,
            class_level=request.class_name,
            topic=request.topic
        )
        # Ensure required fields exist (fallback to AI or dummy)
        plan_data.setdefault("class", request.class_name)
        plan_data.setdefault("subject", request.subject)
        plan_data.setdefault("topic", request.topic)
        plan_data.setdefault("date", datetime.now().strftime("%d – %B, %Y"))
        plan_data.setdefault("week", "Two")
        plan_data.setdefault("duration", "Forty Minutes")
        plan_data.setdefault("age_group", f"{request.class_name} students")

        # Generate Word document
        doc_bytes = create_lesson_plan_doc(plan_data, teacher_name="ISAH YUSUF")

        filename = f"Lesson_Plan_{request.subject}_{request.topic}.docx"
        return Response(
            content=doc_bytes,
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            headers={"Content-Disposition": f"attachment; filename={filename}"}
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))