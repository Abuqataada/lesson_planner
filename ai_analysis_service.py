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

    def generate_lesson_plan(
        self,
        subject: str,
        class_level: str,
        topic: str,
        template_outline: str = "",
        class_notes: str = "",
    ) -> dict:
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
        notes_hint = ""
        if class_notes.strip():
            notes_hint = f"""
Teacher-provided class notes:
--- BEGIN CLASS NOTES ---
{class_notes.strip()}
--- END CLASS NOTES ---

Treat these notes as the primary source for lesson content. Extract and organize their definitions, facts, examples, activities, and exercises. Correct only clear errors, omit irrelevant material, and do not claim that the notes contain information they do not contain. Supplement them only when needed to make the lesson complete and pedagogically sound.
"""
        prompt = f"""Generate a complete, classroom-ready lesson plan for a {class_level} class studying {subject}. The topic is "{topic}".
{template_hint}
{notes_hint}
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

Quality requirements based on the supplied lesson-plan model:
- Make every field specific to {topic}; never use generic filler.
- Write three measurable, differentiated objectives: basic recall/definition, intermediate explanation/application, and advanced analysis/evaluation/creation.
- Give 3-5 realistic instructional resources suitable for the class and subject.
- State 2-3 concrete prior concepts or experiences learners already have.
- Make the warm-up an engaging 3-5 minute activity with a teacher prompt and expected learner response.
- Make learning_note the most detailed field, with accurate definitions, key ideas, examples, and real-life relevance in a clear sequence. Use newline-separated points, not markdown.
- In teacher_activities, include introduction, explanation or modelling, guided practice, differentiation, checks for understanding, and feedback.
- In student_activities, include collaborative work, individual practice, and a short discussion or report-back where appropriate.
- In assessment, state observable evidence and at least 3 topic-specific questions or tasks spanning the objectives.
- Make the plenary check essential learning, the homework a precise task, and flip_ticket a brief preview of a logical next topic.
- Keep the plan age-appropriate, factually correct, inclusive, and feasible in the stated duration.
- Use plain text only inside JSON strings. Do not use markdown headings or tables.
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
                    temperature=0.45,
                    max_tokens=3500,
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

    def generate_presentation_content(
        self,
        subject: str,
        class_level: str,
        topic: str,
        template_outline: str = "",
        class_notes: str = "",
    ) -> dict:
        if not self.openai_client:
            return self._generate_dummy_presentation_content(subject, class_level, topic)

        template_hint = ""
        if template_outline:
            template_hint = f"""
Uploaded lesson format outline:
{template_outline}

Use the uploaded format only as a tone and content guide. Do not copy the slide structure.
"""
        notes_hint = ""
        if class_notes.strip():
            notes_hint = f"""
Teacher-provided class notes:
--- BEGIN CLASS NOTES ---
{class_notes.strip()}
--- END CLASS NOTES ---

Use these notes as the primary source for definitions, explanations, examples, worked exercises, classwork, and assignments. Organize them into concise slides, omit irrelevant material, and supplement only where necessary for clarity and teaching flow.
"""
        prompt = f"""Create classroom presentation content for a {class_level} class studying {subject}. The topic is "{topic}".
{template_hint}
{notes_hint}
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
            "instructional_resources": [
                f"A class-level {subject} textbook",
                f"Pictures, objects, or a short video illustrating {topic}",
                "Whiteboard and markers",
                "Learner activity sheet",
            ],
            "learning_objectives": {
                "basic": f"Define {topic} in simple terms.",
                "intermediate": f"Explain the key concepts of {topic} with examples.",
                "advanced": f"Analyse how {topic} applies to real-world situations.",
            },
            "prior_knowledge": f"Learners can recall related ideas previously studied in {subject}, describe familiar experiences connected to {topic}, and follow simple observation and discussion routines.",
            "warmup_activity": f"Display a familiar example connected to {topic}. Ask learners to observe it, share what they notice with a partner, and suggest how it relates to {subject}. Record predictions to revisit.",
            "learning_note": f"Definition of {topic}.\n• Key principles.\n• Examples.\n• Importance in {subject}.",
            "teacher_activities": f"Introduce {topic} with a familiar example and elicit prior ideas. Explain and model the key meaning and features step by step. Ask checking questions, guide paired practice, support learners who need prompts or visuals, challenge faster learners with a new application, and correct misconceptions.",
            "student_activities": f"Observe the starter and share prior ideas. Record the main points about {topic}. Work in pairs to explain examples, complete an individual application task, report answers, and improve them after feedback.",
            "assessment": f"Use observation, oral questioning, and the individual task as evidence. Ask learners to: 1) define {topic}; 2) explain two key features with an example; and 3) apply the idea to a new real-life situation.",
            "plenary": f"Learners state the meaning of {topic}, one key idea, and one real-life application. Revisit the warm-up predictions and use a final exit question to correct misconceptions.",
            "homework": f"Write a short explanation of {topic}, give two relevant examples, and describe one everyday application. Include a labelled diagram or worked example where appropriate.",
            "flip_ticket": f"Next topic: Applications of {topic}. Find one everyday example to share at the start of the next lesson.",
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
