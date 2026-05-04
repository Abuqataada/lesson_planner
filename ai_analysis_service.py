import json
import os
from datetime import datetime
from typing import List

from dotenv import load_dotenv
from openai import OpenAI

load_dotenv()


class AIAnalysisService:
    def __init__(self):
        self.huggingface_api_key = os.getenv("HUGGINGFACE_API_KEY")
        self.openai_client = None
        self.models = [
            "deepseek-ai/DeepSeek-V3.2-Exp:novita",
            "meta-llama/Llama-3.3-70B-Instruct",
            "google/gemma-2-9b-it",
            "Qwen/Qwen2.5-7B-Instruct",
            "microsoft/Phi-3.5-mini-instruct",
        ]

        if self.huggingface_api_key:
            try:
                self.openai_client = OpenAI(
                    base_url="https://router.huggingface.co/v1",
                    api_key=self.huggingface_api_key,
                )
                self._test_connection()
            except Exception:
                self.openai_client = None

    def _test_connection(self):
        try:
            for model in self.models[:2]:
                try:
                    self.openai_client.chat.completions.create(
                        model=model,
                        messages=[{"role": "user", "content": "Say Connected if you can read this."}],
                        max_tokens=10,
                        timeout=10,
                    )
                    self.preferred_model = model
                    return
                except Exception:
                    continue
        except Exception:
            self.openai_client = None

    def get_available_models(self) -> List[str]:
        if self.openai_client:
            try:
                models = self.openai_client.models.list()
                return [model.id for model in models.data[:10]]
            except Exception:
                return self.models
        return self.models

    def generate_lesson_plan(self, subject: str, class_level: str, topic: str, template_outline: str = "") -> dict:
        if not self.openai_client:
            return self._generate_dummy_lesson_plan(subject, class_level, topic)

        template_hint = ""
        if template_outline:
            template_hint = f"""
Uploaded lesson format outline:
{template_outline}

Follow the uploaded format as closely as possible:
- Preserve the section order implied by the outline.
- Use the same heading style and field names where possible.
- If the uploaded template has extra sections, include them in the generated plan.
- If the uploaded template omits a section, keep the default pedagogical structure.
"""
        prompt = f"""Generate a detailed lesson plan for a {class_level} class studying {subject}. The topic is "{topic}".
{template_hint}
Use valid JSON only with these keys:
- class, subject, topic, subtopic, date, week, duration, age_group
- instructional_resources (list)
- learning_objectives (object with basic, intermediate, advanced)
- prior_knowledge
- warmup_activity
- learning_note
- teacher_activities
- student_activities
- assessment
- plenary
- homework
- flip_ticket
"""

        models_to_try = [getattr(self, "preferred_model", self.models[0])] + self.models[:3]
        for model in models_to_try:
            try:
                response = self.openai_client.chat.completions.create(
                    model=model,
                    messages=[
                        {"role": "system", "content": "You generate lesson plans in valid JSON."},
                        {"role": "user", "content": prompt},
                    ],
                    temperature=0.7,
                    max_tokens=2000,
                    response_format={"type": "json_object"},
                )
                plan = json.loads(response.choices[0].message.content)
                required = [
                    "class",
                    "subject",
                    "topic",
                    "learning_objectives",
                    "prior_knowledge",
                    "warmup_activity",
                    "learning_note",
                    "teacher_activities",
                    "student_activities",
                    "assessment",
                    "plenary",
                    "homework",
                ]
                for key in required:
                    if key not in plan:
                        plan[key] = f"Auto-generated {key.replace('_', ' ')}"
                return plan
            except Exception:
                continue

        return self._generate_dummy_lesson_plan(subject, class_level, topic)

    def _generate_dummy_lesson_plan(self, subject, class_level, topic):
        return {
            "class": class_level,
            "subject": subject,
            "topic": topic,
            "subtopic": topic,
            "date": datetime.now().strftime("%d %B, %Y"),
            "week": "1",
            "duration": "Forty Minutes",
            "age_group": f"{class_level} students",
            "instructional_resources": ["Textbook", "Whiteboard", "Markers"],
            "learning_objectives": {
                "basic": f"Define {topic} in simple terms.",
                "intermediate": f"Explain the key concepts of {topic} with examples.",
                "advanced": f"Analyse how {topic} applies to real-world situations.",
            },
            "prior_knowledge": "Students have basic understanding of the topic area.",
            "warmup_activity": f"Ask students what they already know about {topic}.",
            "learning_note": f"Definition of {topic}.\n• Key principles.\n• Examples.\n• Importance in {subject}.",
            "teacher_activities": "Present the key points and guide discussion.",
            "student_activities": "Discuss, ask questions, and complete class exercises.",
            "assessment": "Observe participation and review classwork.",
            "plenary": f"Summarise the key points about {topic}.",
            "homework": f"Write one paragraph about {topic}.",
            "flip_ticket": f"Next topic: Applications of {topic}",
        }
