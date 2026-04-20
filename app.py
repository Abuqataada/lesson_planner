import io
import os
import unicodedata
import uuid
from datetime import datetime
from pathlib import Path
from typing import Optional

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, Pt
from fastapi import FastAPI, File, Form, Header, HTTPException, Query, Request, UploadFile
from fastapi.responses import HTMLResponse, Response
from fastapi.templating import Jinja2Templates
import psycopg
from psycopg.rows import dict_row
from pydantic import BaseModel

try:
    from pypdf import PdfReader
except ImportError:
    PdfReader = None

from ai_analysis_service import AIAnalysisService

app = FastAPI()
ai_service = AIAnalysisService()
BASE_DIR = Path(__file__).resolve().parent
templates = Jinja2Templates(directory=str(BASE_DIR / "templates"))

DATABASE_URL = os.getenv(
    "DATABASE_URL",
    "postgresql://postgres:postgres@localhost:5432/omtech_ei_db",
)
ADMIN_PASSWORD = os.getenv("ADMIN_PASSWORD", "omtech-admin")
PAYMENT_NAME = "OMTECH EI LIMITED"
PAYMENT_ACCOUNT = "7065894127"
PAYMENT_BANK = "MONIEPOINT"

SUBSCRIPTION_PLANS = [
    {
        "id": "termly",
        "name": "Termly Access",
        "price": "NGN 2,000",
        "billing": "per term",
        "badge": "Recommended",
        "description": "Best for schools and teachers who want access for a single term.",
    },
    {
        "id": "annual",
        "name": "Annual Access",
        "price": "NGN 20,000",
        "billing": "per year",
        "badge": "Best Value",
        "description": "Best for active users who want uninterrupted access.",
    },
]


class LessonPlanRequest(BaseModel):
    class_name: str
    subject: str
    topic: str
    subscriber_key: Optional[str] = None


class SubscriptionRequest(BaseModel):
    full_name: str
    email: str
    phone: str
    plan_id: str = "termly"
    payment_reference: str


def normalize_key(value: str) -> str:
    return (value or "").strip().lower()


def extract_text_from_docx(file_bytes: bytes) -> str:
    buffer = io.BytesIO(file_bytes)
    document = Document(buffer)
    parts: list[str] = []

    for paragraph in document.paragraphs:
        text = paragraph.text.strip()
        if text:
            parts.append(text)

    for table in document.tables:
        for row in table.rows:
            row_text = " | ".join(cell.text.strip() for cell in row.cells if cell.text.strip())
            if row_text:
                parts.append(row_text)

    return "\n".join(parts)


def extract_text_from_pdf(file_bytes: bytes) -> str:
    reader = PdfReader(io.BytesIO(file_bytes))
    parts: list[str] = []
    for page in reader.pages:
        text = page.extract_text() or ""
        text = text.strip()
        if text:
            parts.append(text)
    return "\n".join(parts)


def extract_template_outline(upload: UploadFile) -> tuple[str, str]:
    file_bytes = upload.file.read()
    filename = (upload.filename or "").lower()

    if filename.endswith(".docx"):
        return extract_text_from_docx(file_bytes), "docx"
    if filename.endswith(".pdf"):
        return extract_text_from_pdf(file_bytes), "pdf"

    raise HTTPException(status_code=400, detail="Template must be a PDF or DOCX file.")


def derive_template_labels(template_text: str) -> dict[str, str]:
    text = normalize_key(template_text)
    labels = {}

    mappings = {
        "class": "Class",
        "subject": "Subject",
        "topic": "Topic",
        "subtopic": "Subtopic",
        "date": "Date",
        "week": "Week",
        "duration": "Duration",
        "resources": "Instructional Resources",
        "learning objectives": "Learning Objectives",
        "prior knowledge": "Prior Knowledge",
        "warm-up": "Warm-up Activity",
        "warm up": "Warm-up Activity",
        "teacher activities": "Teacher's Activities",
        "students activities": "Students' Activities",
        "student activities": "Students' Activities",
        "assessment": "Assessment/Evaluation",
        "plenary": "Plenary",
        "homework": "Home-Work",
        "flip ticket": "Flip Ticket",
    }

    for needle, label in mappings.items():
        if needle in text:
            labels[needle] = label

    return labels


def get_db_connection():
    conn = psycopg.connect(DATABASE_URL, row_factory=dict_row)
    return conn


def init_db() -> None:
    with get_db_connection() as conn:
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS subscriptions (
                id TEXT PRIMARY KEY,
                subscriber_key TEXT NOT NULL UNIQUE,
                full_name TEXT NOT NULL,
                email TEXT NOT NULL,
                phone TEXT NOT NULL,
                plan_id TEXT NOT NULL,
                payment_reference TEXT NOT NULL,
                status TEXT NOT NULL,
                created_at TEXT NOT NULL,
                updated_at TEXT NOT NULL,
                approved_at TEXT,
                rejected_at TEXT
            )
            """
        )
        conn.execute("CREATE INDEX IF NOT EXISTS idx_subscriptions_status ON subscriptions(status)")
        conn.commit()


def load_subscriptions() -> list[dict]:
    with get_db_connection() as conn:
        rows = conn.execute(
            "SELECT id, subscriber_key, full_name, email, phone, plan_id, payment_reference, status, created_at, updated_at, approved_at, rejected_at FROM subscriptions ORDER BY created_at DESC"
        ).fetchall()
    return [dict(row) for row in rows]


def subscription_key(record: dict) -> str:
    return normalize_key(record.get("email") or record.get("phone") or record.get("subscriber_key") or "")


def find_subscription(subscriber_key: str) -> Optional[dict]:
    key = normalize_key(subscriber_key)
    if not key:
        return None

    with get_db_connection() as conn:
        row = conn.execute(
            """
            SELECT id, subscriber_key, full_name, email, phone, plan_id, payment_reference, status, created_at, updated_at, approved_at, rejected_at
            FROM subscriptions
            WHERE subscriber_key = %s OR lower(email) = %s OR lower(phone) = %s
            LIMIT 1
            """,
            (key, key, key),
        ).fetchone()
    if row:
        return dict(row)
    return None


def upsert_subscription(payload: SubscriptionRequest) -> dict:
    key = normalize_key(payload.email or payload.phone)
    now = datetime.now().isoformat()
    existing = find_subscription(key)

    if existing:
        status = "approved" if existing.get("status") == "approved" else "pending"
        approved_at = existing.get("approved_at") if status == "approved" else None
        rejected_at = None
        record_id = existing.get("id", str(uuid.uuid4()))
        created_at = existing.get("created_at", now)
    else:
        status = "pending"
        approved_at = None
        rejected_at = None
        record_id = str(uuid.uuid4())
        created_at = now

    record = {
        "id": record_id,
        "subscriber_key": key,
        "full_name": payload.full_name.strip(),
        "email": payload.email.strip(),
        "phone": payload.phone.strip(),
        "plan_id": payload.plan_id,
        "payment_reference": payload.payment_reference.strip(),
        "status": status,
        "created_at": created_at,
        "updated_at": now,
        "approved_at": approved_at,
        "rejected_at": rejected_at,
    }

    with get_db_connection() as conn:
        conn.execute(
            """
            INSERT INTO subscriptions (
                id, subscriber_key, full_name, email, phone, plan_id,
                payment_reference, status, created_at, updated_at, approved_at, rejected_at
            ) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
            ON CONFLICT(subscriber_key) DO UPDATE SET
                full_name=excluded.full_name,
                email=excluded.email,
                phone=excluded.phone,
                plan_id=excluded.plan_id,
                payment_reference=excluded.payment_reference,
                status=excluded.status,
                updated_at=excluded.updated_at,
                approved_at=excluded.approved_at,
                rejected_at=excluded.rejected_at
            """,
            (
                record["id"],
                record["subscriber_key"],
                record["full_name"],
                record["email"],
                record["phone"],
                record["plan_id"],
                record["payment_reference"],
                record["status"],
                record["created_at"],
                record["updated_at"],
                record["approved_at"],
                record["rejected_at"],
            ),
        )
        conn.commit()

    return record


def admin_required(admin_password: str) -> None:
    if admin_password != ADMIN_PASSWORD:
        raise HTTPException(status_code=401, detail="Invalid admin password")


def set_subscription_status(subscription_id: str, status: str) -> dict:
    now = datetime.now().isoformat()
    with get_db_connection() as conn:
        row = conn.execute(
            "SELECT * FROM subscriptions WHERE id = %s",
            (subscription_id,),
        ).fetchone()
        if not row:
            raise HTTPException(status_code=404, detail="Subscription not found")

        item = dict(row)
        item["status"] = status
        item["updated_at"] = now
        if status == "approved":
            item["approved_at"] = now
        if status == "rejected":
            item["rejected_at"] = now

        conn.execute(
            """
            UPDATE subscriptions
            SET status = %s, updated_at = %s, approved_at = %s, rejected_at = %s
            WHERE id = %s
            """,
            (
                item["status"],
                item["updated_at"],
                item.get("approved_at"),
                item.get("rejected_at"),
                subscription_id,
            ),
        )
        conn.commit()
        return item


init_db()


def set_landscape(doc: Document) -> None:
    section = doc.sections[0]
    section.page_width = Inches(11)
    section.page_height = Inches(8.5)
    section.top_margin = Inches(0.5)
    section.bottom_margin = Inches(0.5)
    section.left_margin = Inches(0.5)
    section.right_margin = Inches(0.5)


def add_logo(doc: Document, logo_paths: list[str]) -> bool:
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


def sanitize_filename(text: str, max_len: int = 40) -> str:
    safe = "".join(c for c in text if c.isalnum() or c in (" ", "-", "_"))
    safe = safe.replace(" ", "_")
    return safe[:max_len]


def clean_text(text: str) -> str:
    if not isinstance(text, str):
        text = str(text)

    replacements = {
        "\u2018": "'",
        "\u2019": "'",
        "\u201c": '"',
        "\u201d": '"',
        "\u2013": "-",
        "\u2014": "-",
        "\u2026": "...",
    }
    for bad, good in replacements.items():
        text = text.replace(bad, good)

    return unicodedata.normalize("NFKD", text).encode("ascii", "ignore").decode("ascii")


def create_lesson_plan_doc(
    plan_data: dict,
    teacher_name: str = "ISAH YUSUF",
    template_labels: Optional[dict[str, str]] = None,
    template_name: Optional[str] = None,
) -> bytes:
    doc = Document()
    set_landscape(doc)
    template_labels = template_labels or {}

    plan_data = {k: clean_text(v) if isinstance(v, str) else v for k, v in plan_data.items()}
    if "learning_objectives" in plan_data and isinstance(plan_data["learning_objectives"], dict):
        for key, value in list(plan_data["learning_objectives"].items()):
            if isinstance(value, str):
                plan_data["learning_objectives"][key] = clean_text(value)

    add_logo(doc, ["logo.png", "logo.jpg", "arndale_logo.png", "school_logo.jpg"])

    heading = doc.add_heading("Lesson Plan", level=1)
    heading.alignment = WD_ALIGN_PARAGRAPH.CENTER

    if template_name:
        note = doc.add_paragraph(f"Template: {clean_text(template_name)}")
        note.alignment = WD_ALIGN_PARAGRAPH.CENTER
        note.runs[0].italic = True
        doc.add_paragraph()

    table_main = doc.add_table(rows=9, cols=3)
    table_main.style = "Table Grid"
    table_main.columns[0].width = Inches(0.6)

    rows_data = [
        (template_labels.get("class", "Class:"), plan_data.get("class", "Year 7")),
        (template_labels.get("subject", "Subject:"), plan_data.get("subject", "Physics")),
        (template_labels.get("topic", "Topic:"), plan_data.get("topic", "Introduction to Physics")),
        (template_labels.get("subtopic", "Subtopic"), plan_data.get("subtopic", plan_data.get("topic", ""))),
        (template_labels.get("date", "Date"), plan_data.get("date", datetime.now().strftime("%d %B, %Y"))),
        (template_labels.get("week", "Week"), plan_data.get("week", "Two")),
        (template_labels.get("duration", "Duration"), plan_data.get("duration", "Forty Minutes")),
        (template_labels.get("age_group", "Student Age Group"), plan_data.get("age_group", "11 - 12 Years")),
        (template_labels.get("resources", "INSTRUCTIONAL RESOURCES:"), ""),
    ]

    for i, (col1, col2) in enumerate(rows_data):
        table_main.cell(i, 0).text = clean_text(col1)
        if i == 8:
            cell2 = table_main.cell(i, 1)
            cell2.text = ""
            resources = plan_data.get("instructional_resources", [])
            if isinstance(resources, list):
                for res in resources:
                    p = cell2.add_paragraph(f"* {clean_text(res)}")
                    p.runs[0].font.size = Pt(11)
            else:
                p = cell2.add_paragraph(f"* {clean_text(resources)}")
                p.runs[0].font.size = Pt(11)
        else:
            table_main.cell(i, 1).text = clean_text(str(col2))

    col3_cell = table_main.cell(0, 2)
    for row in range(1, 9):
        col3_cell.merge(table_main.cell(row, 2))
    col3_cell.text = ""

    title = col3_cell.add_paragraph(template_labels.get("learning objectives", "Learning Objectives (Differentiated)"))
    title.runs[0].bold = True
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER

    learning_objectives = plan_data.get("learning_objectives", {})
    col3_cell.add_paragraph(template_labels.get("basic objective", "Basic Objective (for struggling learners):")).runs[0].bold = True
    col3_cell.add_paragraph("By the end of the lesson, students will be able to:")
    p_basic = col3_cell.add_paragraph(f"* {clean_text(learning_objectives.get('basic', 'Define the topic clearly.'))}")
    p_basic.paragraph_format.left_indent = Pt(36)

    col3_cell.add_paragraph(template_labels.get("intermediate objective", "Intermediate Objective (for most students):")).runs[0].bold = True
    col3_cell.add_paragraph("By the end of the lesson, students will be able to:")
    p_inter = col3_cell.add_paragraph(f"* {clean_text(learning_objectives.get('intermediate', 'Explain the topic with examples.'))}")
    p_inter.paragraph_format.left_indent = Pt(36)

    col3_cell.add_paragraph(template_labels.get("advanced objective", "Advanced Objective (for high-achieving students):")).runs[0].bold = True
    col3_cell.add_paragraph("By the end of the lesson, students will be able to:")
    p_adv = col3_cell.add_paragraph(f"* {clean_text(learning_objectives.get('advanced', 'Analyse the topic in real-world situations.'))}")
    p_adv.paragraph_format.left_indent = Pt(36)

    doc.add_paragraph()

    table_dev = doc.add_table(rows=3, cols=6)
    table_dev.style = "Table Grid"

    prior_cell = table_dev.cell(0, 0)
    prior_cell.merge(table_dev.cell(1, 0))
    prior_cell.text = template_labels.get("prior knowledge", "Prior Knowledge")
    prior_cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

    warmup_cell = table_dev.cell(0, 1)
    warmup_cell.merge(table_dev.cell(1, 1))
    warmup_cell.text = template_labels.get("warm-up", "Warm-up Activity")
    warmup_cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

    note_cell = table_dev.cell(0, 2)
    note_cell.merge(table_dev.cell(1, 2))
    note_cell.text = template_labels.get("summarised learning note", "Summarised Learning Note")
    note_cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

    assess_cell = table_dev.cell(0, 5)
    assess_cell.merge(table_dev.cell(1, 5))
    assess_cell.text = template_labels.get("assessment", "Assessment/Evaluation")
    assess_cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

    learning_cell = table_dev.cell(0, 3)
    learning_cell.merge(table_dev.cell(0, 4))
    learning_cell.text = template_labels.get("learning activities", "Learning Activities")
    learning_cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

    table_dev.cell(1, 3).text = template_labels.get("teacher activities", "TEACHER'S ACTIVITIES")
    table_dev.cell(1, 3).paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    table_dev.cell(1, 4).text = template_labels.get("students activities", "STUDENTS' ACTIVITIES")
    table_dev.cell(1, 4).paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

    table_dev.cell(2, 0).text = clean_text(plan_data.get("prior_knowledge", "General knowledge from previous lessons."))
    table_dev.cell(2, 1).text = clean_text(plan_data.get("warmup_activity", "Engaging starter to capture interest."))
    table_dev.cell(2, 2).text = clean_text(plan_data.get("learning_note", "Core content with definitions and examples."))
    table_dev.cell(2, 3).text = clean_text(plan_data.get("teacher_activities", "Teacher's actions during the lesson."))
    table_dev.cell(2, 4).text = clean_text(plan_data.get("student_activities", "Group work, individual tasks, discussions."))
    table_dev.cell(2, 5).text = clean_text(plan_data.get("assessment", "Formative assessment methods."))

    doc.add_paragraph()

    table_plenary = doc.add_table(rows=3, cols=2)
    table_plenary.style = "Table Grid"
    plenary_data = [
        (template_labels.get("plenary", "Plenary"), plan_data.get("plenary", "Summarise key points.")),
        (template_labels.get("homework", "Home-Work"), plan_data.get("homework", "Reinforcement task.")),
        (template_labels.get("flip ticket", "Flip Ticket (next Topic)"), plan_data.get("flip_ticket", "Preview of next lesson.")),
    ]
    for index, (label, text) in enumerate(plenary_data):
        table_plenary.cell(index, 0).text = clean_text(label)
        table_plenary.cell(index, 1).text = clean_text(text)

    doc.add_paragraph()

    sig_row = doc.add_table(rows=1, cols=2)
    sig_row.style = "Table Grid"
    left_cell = sig_row.cell(0, 0)
    p_teacher = left_cell.paragraphs[0]
    p_teacher.text = "Teacher's Name: "
    run = p_teacher.add_run(clean_text(teacher_name))
    run.italic = True
    left_cell.add_paragraph("Supervising Officer's Signature: ____________________")

    right_cell = sig_row.cell(0, 1)
    right_cell.text = (
        "Supervising officer's Comment:\n"
        "........................................................\n"
        "........................................................\n"
        "........................................................\n"
        "........................................................"
    )

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()


@app.get("/", response_class=HTMLResponse)
async def home(request: Request):
    return templates.TemplateResponse(
        request,
        "form.html",
        {
            "subscription_plans": SUBSCRIPTION_PLANS,
            "payment_name": PAYMENT_NAME,
            "payment_account": PAYMENT_ACCOUNT,
            "payment_bank": PAYMENT_BANK,
            "admin_path": "/admin",
        },
    )


@app.get("/admin", response_class=HTMLResponse)
async def admin_dashboard(request: Request):
    return templates.TemplateResponse(
        request,
        "admin.html",
        {
            "subscription_plans": SUBSCRIPTION_PLANS,
            "admin_password_hint": "Set ADMIN_PASSWORD in your environment to protect approvals.",
        },
    )


@app.get("/api/plans")
async def get_plans():
    return {
        "plans": SUBSCRIPTION_PLANS,
        "payment": {
            "name": PAYMENT_NAME,
            "account": PAYMENT_ACCOUNT,
            "bank": PAYMENT_BANK,
        },
    }


@app.get("/api/subscription-status")
async def subscription_status(subscriber_key: str = Query(default="")):
    record = find_subscription(subscriber_key)
    return {
        "status": record.get("status", "pending") if record else "unregistered",
        "subscription": record,
    }


@app.post("/api/subscribe")
async def subscribe(payload: SubscriptionRequest):
    record = upsert_subscription(payload)
    return {
        "message": "Subscription request saved. Awaiting admin approval.",
        "subscription": record,
    }


@app.get("/api/admin/subscriptions")
async def list_subscriptions(x_admin_password: str = Header(default="", alias="X-Admin-Password")):
    admin_required(x_admin_password)
    items = load_subscriptions()
    items.sort(key=lambda item: item.get("created_at", ""), reverse=True)
    return {"items": items}


@app.post("/api/admin/subscriptions/{subscription_id}/approve")
async def approve_subscription(subscription_id: str, x_admin_password: str = Header(default="", alias="X-Admin-Password")):
    admin_required(x_admin_password)
    return {"subscription": set_subscription_status(subscription_id, "approved")}


@app.post("/api/admin/subscriptions/{subscription_id}/reject")
async def reject_subscription(subscription_id: str, x_admin_password: str = Header(default="", alias="X-Admin-Password")):
    admin_required(x_admin_password)
    return {"subscription": set_subscription_status(subscription_id, "rejected")}


@app.post("/generate")
async def generate_plan(
    class_name: str = Form(...),
    subject: str = Form(...),
    topic: str = Form(...),
    subscriber_key: str = Form(...),
    lesson_template: UploadFile | None = File(None),
):
    if not subscriber_key:
        raise HTTPException(status_code=403, detail="Subscription required. Please subscribe before using the app.")

    subscription = find_subscription(subscriber_key)
    if not subscription or subscription.get("status") != "approved":
        raise HTTPException(status_code=403, detail="Subscription pending. Access is unlocked after admin approval.")

    try:
        template_outline = ""
        template_name = None
        template_labels = {}
        if lesson_template and lesson_template.filename:
            template_outline, template_ext = extract_template_outline(lesson_template)
            template_name = lesson_template.filename
            if template_outline.strip():
                template_labels = derive_template_labels(template_outline)
        else:
            template_ext = None

        plan_data = ai_service.generate_lesson_plan(
            subject=subject,
            class_level=class_name,
            topic=topic,
            template_outline=template_outline,
        )
        plan_data.setdefault("class", class_name)
        plan_data.setdefault("subject", subject)
        plan_data.setdefault("topic", topic)
        plan_data.setdefault("date", datetime.now().strftime("%d %B, %Y"))
        plan_data.setdefault("week", "Two")
        plan_data.setdefault("duration", "Forty Minutes")
        plan_data.setdefault("age_group", f"{class_name} students")

        doc_bytes = create_lesson_plan_doc(
            plan_data,
            teacher_name="ISAH YUSUF",
            template_labels=template_labels,
            template_name=template_name,
        )

        safe_subject = sanitize_filename(subject, max_len=30)
        safe_topic = sanitize_filename(topic, max_len=40)
        filename = f"Lesson_Plan_{safe_subject}_{safe_topic}.docx"

        return Response(
            content=doc_bytes,
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            headers={"Content-Disposition": f"attachment; filename={filename}"},
        )
    except HTTPException:
        raise
    except Exception as exc:
        raise HTTPException(status_code=500, detail=str(exc))