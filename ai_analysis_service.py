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

    def generate_presentation_content(self, subject: str, class_level: str, topic: str, template_outline: str = "") -> dict:
        if not self.openai_client:
            return self._generate_dummy_presentation_content(subject, class_level, topic)

        template_hint = ""
        if template_outline:
            template_hint = f"""
Uploaded lesson format outline:
{template_outline}

Use the uploaded format only as a tone and content guide. Do not copy the slide structure.
"""
        prompt = f"""Create classroom presentation content for a {class_level} class studying {subject}. The topic is "{topic}".
{template_hint}
Return valid JSON only with these keys:
- cover_subtitle
- overview_line
- meaning_heading
- meaning_text
- examples_heading
- examples (list)
- key_terms_heading
- key_terms (list)
- worked_examples_heading
- worked_examples (list)
- classwork_heading
- classwork (list)
- weekend_assignment_heading
- weekend_assignment (list)
- closing_line

Write the content in a teaching style that fits a classroom presentation. Use short clear sentences, topic-focused examples, and assessment items that match the topic. If the subject is scientific or mathematical, include formulas or calculations where relevant. If it is humanities or language-based, use discussion and application-oriented examples.
"""

        models_to_try = [getattr(self, "preferred_model", self.models[0])] + self.models[:3]
        for model in models_to_try:
            try:
                response = self.openai_client.chat.completions.create(
                    model=model,
                    messages=[
                        {"role": "system", "content": "You generate classroom presentation content in valid JSON."},
                        {"role": "user", "content": prompt},
                    ],
                    temperature=0.7,
                    max_tokens=2200,
                    response_format={"type": "json_object"},
                )
                content = json.loads(response.choices[0].message.content)
                return self._normalize_presentation_content(content, subject, class_level, topic)
            except Exception:
                continue

        return self._generate_dummy_presentation_content(subject, class_level, topic)

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

    def _normalize_presentation_content(self, content: dict, subject: str, class_level: str, topic: str) -> dict:
        content = content or {}
        return {
            "cover_subtitle": content.get("cover_subtitle") or f"Definition, examples, and applications of {topic}",
            "overview_line": content.get("overview_line") or f"{topic}: Definition, examples, and applications.",
            "meaning_heading": content.get("meaning_heading") or f"MEANING OF {str(topic).upper()}",
            "meaning_text": content.get("meaning_text") or f"{topic} is an important concept in {subject} for {class_level} learners.",
            "examples_heading": content.get("examples_heading") or "Examples",
            "examples": content.get("examples") or [f"An everyday example related to {topic}."],
            "key_terms_heading": content.get("key_terms_heading") or f"TERMS ASSOCIATED WITH {str(topic).upper()}",
            "key_terms": content.get("key_terms") or [f"Key term one in {topic}", f"Key term two in {topic}"],
            "worked_examples_heading": content.get("worked_examples_heading") or "Examples",
            "worked_examples": content.get("worked_examples") or [f"Worked example on {topic}."],
            "classwork_heading": content.get("classwork_heading") or "CLASSWORK",
            "classwork": content.get("classwork") or [f"Define {topic} and mention two applications."],
            "weekend_assignment_heading": content.get("weekend_assignment_heading") or "WEEKEND ASSIGNMENT",
            "weekend_assignment": content.get("weekend_assignment") or [f"Answer questions on {topic} at home."],
            "closing_line": content.get("closing_line") or "THANK YOU",
        }

    def _generate_dummy_presentation_content(self, subject, class_level, topic):
        topic_lower = topic.lower()
        examples = [
            f"A real-life example of {topic}.",
            f"An observed classroom example of {topic}.",
        ]
        key_terms = [
            f"Main idea related to {topic}",
            f"Application of {topic}",
        ]
        worked_examples = [
            f"Explain {topic} using a simple practical example.",
            f"Solve one short question based on {topic}.",
        ]
        classwork = [
            f"1. Define {topic}.",
            f"2. Mention two examples of {topic}.",
            f"3. State one application of {topic}.",
        ]
        weekend_assignment = [
            f"1. Write a short note on {topic}.",
            f"2. Answer four questions on {topic}.",
        ]

        if "projectile" in topic_lower:
            examples = [
                "A thrown rubber ball re-bouncing from a wall.",
                "An athlete doing the high jump.",
                "A stone released from a catapult.",
                "A bullet fired from a gun.",
                "A cricket ball thrown against a vertical wall.",
            ]
            key_terms = [
                "Time of flight - time required to return to the same level.",
                "Maximum height - highest vertical distance reached.",
                "Range - horizontal distance from projection to landing point.",
            ]
            worked_examples = [
                "A stone is shot out from a catapult with an initial velocity of 30m/s at an elevation of 60°. Find the time of flight, maximum height, and range.",
                "A body is projected horizontally with a velocity of 60m/s from the top of a building 120m above the ground. Calculate the time of flight and range.",
                "A projectile is fired at 60° with an initial velocity of 80m/s. Calculate the time of flight, maximum height, and velocity after 2 seconds.",
                "A stone is projected horizontally with a speed of 10m/s from the top of a tower 50m high. Find the speed with which it strikes the ground.",
            ]
            classwork = [
                "1. (a) Define the term projectile. (b) Mention two applications of projectiles.",
                "2. A ball is projected horizontally from the top of a hill with a velocity of 30m/s. If it reaches the ground 5 seconds later, find the height of the hill.",
                "3. A stone propelled from a catapult with a speed of 50m/s attains a height of 100m. Calculate the time of flight, the angle of projection, and the range attained.",
            ]
            weekend_assignment = [
                "1. A stone is projected at an angle of 60° and an initial velocity of 20m/s. Determine the time of flight.",
                "2. For a projectile, the maximum range is obtained when the angle of projection is which of the following?",
                "3. A gun fires a shell at an angle of elevation of 30° with a velocity of 20m/s. Find the horizontal and vertical components of the velocity, the range, and the maximum height.",
                "4. Explain the range of a projectile and calculate the maximum height attained by a body projected at 30° with speed 50m/s.",
            ]

        return self._normalize_presentation_content(
            {
                "cover_subtitle": f"Definition, derivation, examples, and applications of {topic}.",
                "overview_line": f"{topic}: Definition, derivation of equations, and applications.",
                "meaning_heading": f"MEANING OF {topic.upper()}",
                "meaning_text": f"{topic} is a concept studied in {subject} for {class_level} learners. It should be explained clearly with examples and applications.",
                "examples_heading": "Examples",
                "examples": examples,
                "key_terms_heading": f"TERMS ASSOCIATED WITH {topic.upper()}",
                "key_terms": key_terms,
                "worked_examples_heading": "Examples",
                "worked_examples": worked_examples,
                "classwork_heading": "CLASSWORK",
                "classwork": classwork,
                "weekend_assignment_heading": "WEEKEND ASSIGNMENT",
                "weekend_assignment": weekend_assignment,
                "closing_line": "THANK YOU",
            },
            subject,
            class_level,
            topic,
        )
