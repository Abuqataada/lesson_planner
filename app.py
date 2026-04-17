import io
import json
import os
import unicodedata
import uuid
from datetime import datetime
from pathlib import Path
from typing import Optional

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, Pt
from fastapi import FastAPI, Header, HTTPException, Query, Request
from fastapi.responses import HTMLResponse, Response
from fastapi.templating import Jinja2Templates
from pydantic import BaseModel

from ai_analysis_service import AIAnalysisService

app = FastAPI()
ai_service = AIAnalysisService()
BASE_DIR = Path(__file__).resolve().parent
templates = Jinja2Templates(directory=str(BASE_DIR / "templates"))

SUBSCRIPTIONS_FILE = BASE_DIR / "subscriptions.json"
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
    plan_id: str = "monthly"
    payment_reference: str


def normalize_key(value: str) -> str:
    return (value or "").strip().lower()


def load_subscriptions() -> list[dict]:
    if not SUBSCRIPTIONS_FILE.exists():
        return []
    try:
        with SUBSCRIPTIONS_FILE.open("r", encoding="utf-8") as handle:
            data = json.load(handle)
        return data if isinstance(data, list) else []
    except Exception:
        return []


def save_subscriptions(items: list[dict]) -> None:
    with SUBSCRIPTIONS_FILE.open("w", encoding="utf-8") as handle:
        json.dump(items, handle, indent=2)


def subscription_key(record: dict) -> str:
    return normalize_key(record.get("email") or record.get("phone") or record.get("subscriber_key") or "")


def find_subscription(subscriber_key: str) -> Optional[dict]:
    key = normalize_key(subscriber_key)
    if not key:
        return None

    for item in load_subscriptions():
        if normalize_key(item.get("subscriber_key", "")) == key:
            return item
        if subscription_key(item) == key:
            return item
    return None


def upsert_subscription(payload: SubscriptionRequest) -> dict:
    items = load_subscriptions()
    key = normalize_key(payload.email or payload.phone)
    now = datetime.now().isoformat()

    existing_index = None
    for index, item in enumerate(items):
        if normalize_key(item.get("subscriber_key", "")) == key or subscription_key(item) == key:
            existing_index = index
            break

    if existing_index is not None:
        existing = items[existing_index]
        record_id = existing.get("id", str(uuid.uuid4()))
        created_at = existing.get("created_at", now)
        status = "approved" if existing.get("status") == "approved" else "pending"
        approved_at = existing.get("approved_at") if status == "approved" else None
        rejected_at = None if status == "pending" else existing.get("rejected_at")
    else:
        record_id = str(uuid.uuid4())
        created_at = now
        status = "pending"
        approved_at = None
        rejected_at = None

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

    if existing_index is None:
        items.append(record)
    else:
        items[existing_index] = record

    save_subscriptions(items)
    return record


def admin_required(admin_password: str) -> None:
    if admin_password != ADMIN_PASSWORD:
        raise HTTPException(status_code=401, detail="Invalid admin password")


def set_subscription_status(subscription_id: str, status: str) -> dict:
    items = load_subscriptions()
    now = datetime.now().isoformat()

    for item in items:
        if item.get("id") == subscription_id:
            item["status"] = status
            item["updated_at"] = now
            if status == "approved":
                item["approved_at"] = now
            if status == "rejected":
                item["rejected_at"] = now
            save_subscriptions(items)
            return item

    raise HTTPException(status_code=404, detail="Subscription not found")


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


def create_lesson_plan_doc(plan_data: dict, teacher_name: str = "ISAH YUSUF") -> bytes:
    doc = Document()
    set_landscape(doc)

    plan_data = {k: clean_text(v) if isinstance(v, str) else v for k, v in plan_data.items()}
    if "learning_objectives" in plan_data and isinstance(plan_data["learning_objectives"], dict):
        for key, value in list(plan_data["learning_objectives"].items()):
            if isinstance(value, str):
                plan_data["learning_objectives"][key] = clean_text(value)

    add_logo(doc, ["logo.png", "logo.jpg", "arndale_logo.png", "school_logo.jpg"])

    heading = doc.add_heading("Lesson Plan", level=1)
    heading.alignment = WD_ALIGN_PARAGRAPH.CENTER

    table_main = doc.add_table(rows=9, cols=3)
    table_main.style = "Table Grid"
    table_main.columns[0].width = Inches(0.6)

    rows_data = [
        ("Class:", plan_data.get("class", "Year 7")),
        ("Subject:", plan_data.get("subject", "Physics")),
        ("Topic:", plan_data.get("topic", "Introduction to Physics")),
        ("Subtopic", plan_data.get("subtopic", plan_data.get("topic", ""))),
        ("Date", plan_data.get("date", datetime.now().strftime("%d %B, %Y"))),
        ("Week", plan_data.get("week", "Two")),
        ("Duration", plan_data.get("duration", "Forty Minutes")),
        ("Student Age Group", plan_data.get("age_group", "11 - 12 Years")),
        ("INSTRUCTIONAL RESOURCES:", ""),
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

    title = col3_cell.add_paragraph("Learning Objectives (Differentiated)")
    title.runs[0].bold = True
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER

    learning_objectives = plan_data.get("learning_objectives", {})
    col3_cell.add_paragraph("Basic Objective (for struggling learners):").runs[0].bold = True
    col3_cell.add_paragraph("By the end of the lesson, students will be able to:")
    p_basic = col3_cell.add_paragraph(f"* {clean_text(learning_objectives.get('basic', 'Define the topic clearly.'))}")
    p_basic.paragraph_format.left_indent = Pt(36)

    col3_cell.add_paragraph("Intermediate Objective (for most students):").runs[0].bold = True
    col3_cell.add_paragraph("By the end of the lesson, students will be able to:")
    p_inter = col3_cell.add_paragraph(f"* {clean_text(learning_objectives.get('intermediate', 'Explain the topic with examples.'))}")
    p_inter.paragraph_format.left_indent = Pt(36)

    col3_cell.add_paragraph("Advanced Objective (for high-achieving students):").runs[0].bold = True
    col3_cell.add_paragraph("By the end of the lesson, students will be able to:")
    p_adv = col3_cell.add_paragraph(f"* {clean_text(learning_objectives.get('advanced', 'Analyse the topic in real-world situations.'))}")
    p_adv.paragraph_format.left_indent = Pt(36)

    doc.add_paragraph()

    table_dev = doc.add_table(rows=3, cols=6)
    table_dev.style = "Table Grid"

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
        ("Plenary", plan_data.get("plenary", "Summarise key points.")),
        ("Home-Work", plan_data.get("homework", "Reinforcement task.")),
        ("Flip Ticket (next Topic)", plan_data.get("flip_ticket", "Preview of next lesson.")),
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
async def generate_plan(request: LessonPlanRequest):
    if not request.subscriber_key:
        raise HTTPException(status_code=403, detail="Subscription required. Please subscribe before using the app.")

    subscription = find_subscription(request.subscriber_key)
    if not subscription or subscription.get("status") != "approved":
        raise HTTPException(status_code=403, detail="Subscription pending. Access is unlocked after admin approval.")

    try:
        plan_data = ai_service.generate_lesson_plan(
            subject=request.subject,
            class_level=request.class_name,
            topic=request.topic,
        )
        plan_data.setdefault("class", request.class_name)
        plan_data.setdefault("subject", request.subject)
        plan_data.setdefault("topic", request.topic)
        plan_data.setdefault("date", datetime.now().strftime("%d %B, %Y"))
        plan_data.setdefault("week", "Two")
        plan_data.setdefault("duration", "Forty Minutes")
        plan_data.setdefault("age_group", f"{request.class_name} students")

        doc_bytes = create_lesson_plan_doc(plan_data, teacher_name="ISAH YUSUF")

        safe_subject = sanitize_filename(request.subject, max_len=30)
        safe_topic = sanitize_filename(request.topic, max_len=40)
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
