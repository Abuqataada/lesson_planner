import json
import logging
import os
import re
from datetime import datetime
from typing import Any, Dict, List, Optional

from dotenv import load_dotenv
from openai import OpenAI

load_dotenv()


logger = logging.getLogger(__name__)


class AIAnalysisService:
    """
    AI service for generating:
    - professional lesson plans
    - classroom presentation content

    Uses OpenRouter's free model router by default.
    """

    OPENROUTER_BASE_URL = "https://openrouter.ai/api/v1"

    LESSON_MODEL = "openrouter/free"
    PRESENTATION_MODEL = "openrouter/free"

    LESSON_MAX_TOKENS = 4000
    PRESENTATION_MAX_TOKENS = 3000

    def __init__(self) -> None:
        self.openrouter_api_key = os.getenv("OPENROUTER_API_KEY")

        self.app_url = os.getenv(
            "OPENROUTER_APP_URL",
            "http://localhost:8000"
        )

        self.app_name = os.getenv(
            "OPENROUTER_APP_NAME",
            "AI Lesson Planner"
        )

        self.openai_client: Optional[OpenAI] = None

        # Keep this list limited to models you actually intend to use.
        #
        # openrouter/free automatically selects an available FREE model.
        self.models: List[str] = [
            "openrouter/free"
        ]

        if not self.openrouter_api_key:
            logger.warning(
                "OPENROUTER_API_KEY is not configured. "
                "AI generation will use local fallback content."
            )
            return

        try:
            self.openai_client = OpenAI(
                base_url=self.OPENROUTER_BASE_URL,
                api_key=self.openrouter_api_key,
                default_headers={
                    "HTTP-Referer": self.app_url,
                    "X-Title": self.app_name,
                },
            )

            logger.info("OpenRouter client initialized successfully.")

        except Exception as exc:
            logger.exception(
                "Failed to initialize OpenRouter client: %s",
                exc
            )
            self.openai_client = None

    # ================================================================
    # GENERAL HELPERS
    # ================================================================

    def is_available(self) -> bool:
        """
        Returns whether an OpenRouter client has been initialized.

        This deliberately does NOT send a test request because free
        inference requests should only be consumed when the user
        actually asks for AI-generated content.
        """
        return self.openai_client is not None

    def get_available_models(self) -> List[str]:
        """
        Return models that this application is configured to use.

        We deliberately do not expose every model returned by OpenRouter
        because many of those models may be paid.
        """
        return self.models.copy()

    @staticmethod
    def _safe_text(value: Any) -> str:
        if value is None:
            return ""

        return str(value).strip()

    @staticmethod
    def _limit_text(
        value: Any,
        max_words: int
    ) -> str:
        text = str(value or "").strip()

        if not text:
            return ""

        words = text.split()

        if len(words) <= max_words:
            return text

        shortened = " ".join(words[:max_words]).rstrip(
            " ,;:"
        )

        if shortened.endswith((".", "!", "?")):
            return shortened

        return shortened + "."

    @staticmethod
    def _limit_learning_note(
        value: object,
        max_words: int = 140,
        max_points: int = 6
    ) -> str:
        """
        Keep the learning note concise even when the selected
        free model ignores the requested limits.
        """

        text = str(value or "").strip()

        if not text:
            return ""

        lines = [
            line.strip(" \t-*•")
            for line in text.splitlines()
            if line.strip()
        ]

        lines = lines[:max_points]

        compact = "\n".join(lines) if lines else text

        words = compact.split()

        if len(words) <= max_words:
            return compact

        shortened = " ".join(
            words[:max_words]
        ).rstrip(" ,;:")

        if shortened.endswith((".", "!", "?")):
            return shortened

        return shortened + "."

    @staticmethod
    def _extract_json(content: Any) -> Dict[str, Any]:
        """
        Robustly extract JSON from model output.

        Some free models may still wrap JSON in Markdown despite being
        told not to. This helper attempts to recover valid JSON instead
        of immediately failing.
        """

        if content is None:
            raise ValueError("The AI returned an empty response.")

        if isinstance(content, dict):
            return content

        text = str(content).strip()

        if not text:
            raise ValueError("The AI returned an empty response.")

        # Remove common markdown code fences.
        text = re.sub(
            r"^```(?:json)?\s*",
            "",
            text,
            flags=re.IGNORECASE
        )

        text = re.sub(
            r"\s*```$",
            "",
            text
        )

        text = text.strip()

        # First try the entire response.
        try:
            data = json.loads(text)

            if isinstance(data, dict):
                return data

            raise ValueError(
                "Expected a JSON object, but received another JSON type."
            )

        except json.JSONDecodeError:
            pass

        # Attempt to locate the first complete-looking JSON object.
        first_brace = text.find("{")
        last_brace = text.rfind("}")

        if first_brace != -1 and last_brace > first_brace:
            candidate = text[
                first_brace:last_brace + 1
            ]

            try:
                data = json.loads(candidate)

                if isinstance(data, dict):
                    return data

            except json.JSONDecodeError:
                pass

        raise ValueError(
            "The AI response did not contain valid JSON."
        )

    # ================================================================
    # OPENROUTER REQUEST HELPER
    # ================================================================

    def _generate_json(
        self,
        *,
        system_prompt: str,
        user_prompt: str,
        model: str,
        temperature: float,
        max_tokens: int
    ) -> Dict[str, Any]:
        """
        Send one structured-output request to OpenRouter.

        openrouter/free will route to a currently available free model.
        """

        if not self.openai_client:
            raise RuntimeError(
                "OpenRouter is not configured."
            )

        try:
            response = (
                self.openai_client.chat.completions.create(
                    model=model,
                    messages=[
                        {
                            "role": "system",
                            "content": system_prompt,
                        },
                        {
                            "role": "user",
                            "content": user_prompt,
                        },
                    ],
                    temperature=temperature,
                    max_tokens=max_tokens,

                    # This tells OpenRouter that JSON output is required.
                    # openrouter/free can route toward a compatible model.
                    response_format={
                        "type": "json_object"
                    },

                    timeout=90,
                )
            )

        except Exception as exc:
            logger.exception(
                "OpenRouter generation failed: %s",
                exc
            )

            raise

        if not response.choices:
            raise ValueError(
                "OpenRouter returned no completion choices."
            )

        message = response.choices[0].message

        if not message:
            raise ValueError(
                "OpenRouter returned an empty message."
            )

        return self._extract_json(
            message.content
        )

    # ================================================================
    # LESSON PLAN GENERATION
    # ================================================================

    def generate_lesson_plan(
        self,
        subject: str,
        class_level: str,
        topic: str,
        template_outline: str = "",
        class_notes: str = "",
    ) -> dict:

        subject = self._safe_text(subject)
        class_level = self._safe_text(class_level)
        topic = self._safe_text(topic)
        template_outline = self._safe_text(
            template_outline
        )
        class_notes = self._safe_text(
            class_notes
        )

        if not subject:
            raise ValueError(
                "Subject is required."
            )

        if not class_level:
            raise ValueError(
                "Class level is required."
            )

        if not topic:
            raise ValueError(
                "Topic is required."
            )

        if not self.openai_client:
            logger.warning(
                "AI unavailable; using lesson-plan fallback."
            )

            return self._generate_dummy_lesson_plan(
                subject,
                class_level,
                topic
            )

        template_hint = ""

        if template_outline:
            template_hint = f"""
SCHOOL / TEACHER LESSON-PLAN TEMPLATE

--- BEGIN TEMPLATE ---
{template_outline}
--- END TEMPLATE ---

Use this template as an important formatting and structural guide.

Requirements:
- Preserve the logical order represented in the template.
- Preserve meaningful field names wherever possible.
- Respect additional educational sections appearing in the template.
- Do not remove essential pedagogical sections merely because they
  are absent from the uploaded template.
"""

        notes_hint = ""

        if class_notes:
            notes_hint = f"""
TEACHER-PROVIDED CLASS NOTES

--- BEGIN CLASS NOTES ---
{class_notes}
--- END CLASS NOTES ---

The teacher-provided notes are the PRIMARY source for lesson content.

Use them to identify:
- definitions
- concepts
- formulas
- explanations
- examples
- exercises
- activities
- terminology
- curriculum emphasis

Do not blindly copy the notes.

Organize and improve them into professional classroom material.

Correct only obvious factual, grammatical, or mathematical errors.

Supplement them only when necessary to make the lesson pedagogically
complete and academically correct.

Do not claim that the notes contain information they do not contain.
"""

        system_prompt = """
You are an expert classroom teacher, curriculum specialist,
instructional designer, educational assessment specialist and
professional lesson-plan writer.

You prepare high-quality lesson plans that an experienced teacher
could confidently use in a real classroom.

Your priorities, in order, are:

1. factual accuracy
2. age appropriateness
3. curriculum relevance
4. measurable learning objectives
5. logical instructional progression
6. alignment between objectives, activities and assessment
7. learner participation
8. practical classroom feasibility
9. professional educational language
10. clarity and usefulness to the classroom teacher

Never use generic filler when specific educational content can be
provided.

When asked for JSON:
- return one valid JSON object only
- do not use Markdown
- do not add explanations before or after the JSON
- obey the requested schema
"""

        prompt = f"""
CREATE A PROFESSIONAL CLASSROOM-READY LESSON PLAN.

LESSON DETAILS

Class Level: {class_level}
Subject: {subject}
Topic: {topic}

{template_hint}

{notes_hint}

============================================================
GENERAL QUALITY STANDARD
============================================================

Write as an experienced teacher preparing a lesson that will actually
be delivered to learners.

Every section must be:
- specific to the topic
- academically accurate
- age appropriate
- practical
- clearly written
- professionally structured
- learner centred
- meaningful rather than generic

Avoid vague statements such as:

"Teacher explains the topic."
"Students listen."
"Teacher asks questions."
"Students answer questions."
"Teacher discusses the lesson."

Instead state:
- WHAT is explained
- HOW it is demonstrated
- WHAT example is used
- WHAT question is asked
- WHAT learners actually do
- WHAT evidence demonstrates learning

============================================================
PEDAGOGICAL SEQUENCE
============================================================

Use a logical progression:

1. Activate prior knowledge.
2. Introduce the topic through an engaging stimulus.
3. Explain or model the new concept.
4. Conduct guided learner participation.
5. Provide collaborative or practical activity where appropriate.
6. Give independent application.
7. Check understanding.
8. Identify and correct misconceptions.
9. Assess learning.
10. Consolidate the major lesson points.
11. Extend learning through homework.

============================================================
LEARNING OBJECTIVES
============================================================

Generate EXACTLY THREE differentiated learning objectives.

BASIC OBJECTIVE:
Use an observable action verb appropriate to remembering or basic
understanding.

Examples of useful verbs include:
- define
- identify
- list
- state
- name
- describe
- recognise

INTERMEDIATE OBJECTIVE:
Require learners to explain, classify, compare, calculate,
demonstrate, illustrate, interpret or apply knowledge.

ADVANCED OBJECTIVE:
Require analysis, evaluation, problem solving, investigation,
prediction, justification, design or creation.

Objectives must be measurable.

Avoid vague verbs such as:
- know
- learn
- appreciate
- understand

unless the expected observable behaviour is explicitly stated.

============================================================
INSTRUCTIONAL RESOURCES
============================================================

Provide between 3 and 6 realistic resources suitable for:
- {class_level}
- {subject}
- {topic}

Prefer resources that a real teacher can reasonably obtain or use,
such as:

- real objects
- charts
- diagrams
- flashcards
- models
- worksheets
- textbooks
- local materials
- simulations
- multimedia
- laboratory materials
- manipulatives
- maps
- photographs

Do not list technology simply for the sake of using technology.

============================================================
PRIOR KNOWLEDGE
============================================================

State the SPECIFIC knowledge, experience or skills learners should
already possess before this lesson.

Do not write:

"Learners already have knowledge of the topic."

Mention actual concepts or experiences.

============================================================
WARM-UP ACTIVITY
============================================================

Design an engaging starter lasting approximately 3-5 minutes.

It should contain:

1. what the teacher presents or does
2. the teacher's actual question or prompt
3. what learners do
4. the expected connection to today's lesson

It must activate prior knowledge and naturally lead into {topic}.

============================================================
LEARNING NOTE
============================================================

Write a concise academic summary suitable for inclusion in a
professional lesson-plan table.

Use approximately 4-6 short newline-separated points.

Target length:
approximately 80-140 words.

Include where relevant:

- meaning or definition
- major concepts
- important facts
- formula, rule or principle
- one short example
- classroom or real-life relevance

For Mathematics:
- use correct notation
- show formulas accurately
- include units
- avoid unexplained jumps in calculations

For Physics and other sciences:
- use precise scientific terminology
- state principles and relationships accurately
- use correct units and symbols
- distinguish evidence, observations, laws and explanations properly

For Languages and Humanities:
- include meaningful context
- definitions
- examples
- interpretation
- vocabulary
- applications where appropriate

Do NOT place:
- teacher instructions
- assessment questions
- homework
- warm-up procedures

inside learning_note.

============================================================
TEACHER ACTIVITIES
============================================================

Describe an actual sequence of instructional actions.

Include where appropriate:

- linkage with prior knowledge
- introduction of the topic
- explanation of important concepts
- modelling or demonstration
- use of instructional materials
- worked examples
- purposeful questioning
- guided practice
- checking for understanding
- correction of misconceptions
- differentiation
- feedback
- preparation for independent work

The teacher activities must contain meaningful content related directly
to {topic}.

============================================================
STUDENT ACTIVITIES
============================================================

Learner activities must correspond with what the teacher is doing.

Include meaningful combinations of:

- observation
- discussion
- questioning
- paired work
- group work
- practical activity
- manipulation
- calculation
- explanation
- recording
- problem solving
- individual practice
- peer checking
- report-back

Learners must be active participants.

============================================================
ASSESSMENT
============================================================

Assessment must directly measure the three learning objectives.

Include at least:

1. one BASIC question/task
2. one INTERMEDIATE question/task
3. one ADVANCED question/task

The questions must contain ACTUAL {topic} content.

Also state the evidence the teacher should look for.

Do not write generic statements such as:

"Ask learners questions about the topic."

============================================================
PLENARY
============================================================

Design a meaningful lesson conclusion lasting approximately 3-5
minutes.

It should:

- revisit the lesson objectives
- consolidate the key concept
- check learner understanding
- reveal remaining misconceptions
- require learners to explain or demonstrate something they learned

============================================================
HOMEWORK
============================================================

Give a precise, age-appropriate task that extends today's learning.

Homework must contain actual questions, exercises, investigation or
application related to {topic}.

Do not write:

"Read more about the topic."

============================================================
FLIP TICKET
============================================================

Provide ONE short preparation task that creates curiosity or prepares
learners for the logical next lesson.

============================================================
LOCAL RELEVANCE
============================================================

Where genuinely useful, relate examples to:

- everyday Nigerian life
- home
- school
- transportation
- markets
- community
- environment
- local technology
- familiar African contexts

Do not force local examples when they reduce academic clarity.

============================================================
CONTENT ACCURACY
============================================================

Never invent facts.

If calculations are needed:
- calculate carefully
- use correct formulas
- show appropriate units
- ensure numerical answers are consistent

If {topic} could mean different things, interpret it according to:
Subject = {subject}
Class = {class_level}

============================================================
OUTPUT
============================================================

Return VALID JSON ONLY.

Use this exact top-level schema:

{{
    "class": "{class_level}",
    "subject": "{subject}",
    "topic": "{topic}",
    "subtopic": "",
    "date": "",
    "week": "",
    "duration": "",
    "age_group": "",
    "instructional_resources": [],
    "learning_objectives": {{
        "basic": "",
        "intermediate": "",
        "advanced": ""
    }},
    "prior_knowledge": "",
    "warmup_activity": "",
    "learning_note": "",
    "teacher_activities": "",
    "student_activities": "",
    "assessment": "",
    "plenary": "",
    "homework": "",
    "flip_ticket": ""
}}

IMPORTANT:

- Return JSON only.
- Do not use Markdown.
- Do not use ```json fences.
- Every required field must contain meaningful content.
- Do not use placeholder text.
- Do not write "auto-generated".
- Do not leave instructional fields blank.
"""

        try:
            plan = self._generate_json(
                system_prompt=system_prompt,
                user_prompt=prompt,
                model=self.LESSON_MODEL,
                temperature=0.30,
                max_tokens=self.LESSON_MAX_TOKENS,
            )

            # First normalize model output
            plan = self._repair_lesson_plan(
                plan=plan,
                subject=subject,
                class_level=class_level,
                topic=topic,
            )

            # Keep learning note under control
            plan["learning_note"] = self._limit_learning_note(
                plan.get("learning_note")
            )

            # Validate repaired output
            is_valid, validation_errors = (
                self._validate_lesson_plan(plan)
            )

            if not is_valid:

                logger.warning(
                    "Lesson plan validation failed after repair.\n"
                    "Errors: %s\n"
                    "Generated plan:\n%s",
                    validation_errors,
                    json.dumps(
                        plan,
                        indent=2,
                        ensure_ascii=False
                    )
                )

                raise ValueError(
                    "Lesson plan validation failed after repair: "
                    + "; ".join(validation_errors)
                )

            return plan

        except Exception as exc:
            logger.exception(
                "Lesson-plan generation failed: %s",
                exc
            )

            return self._generate_dummy_lesson_plan(
                subject,
                class_level,
                topic
            )

    def _normalize_lesson_plan(
        self,
        *,
        plan: Dict[str, Any],
        subject: str,
        class_level: str,
        topic: str
    ) -> None:

        # These values should come from the application/user,
        # not be hallucinated by a free model.
        plan["class"] = class_level
        plan["subject"] = subject
        plan["topic"] = topic

        subtopic = self._safe_text(
            plan.get("subtopic")
        )

        if not subtopic:
            plan["subtopic"] = topic

        plan.setdefault("date", "")
        plan.setdefault("week", "")
        plan.setdefault(
            "duration",
            "40 minutes"
        )

        age_group = self._safe_text(
            plan.get("age_group")
        )

        if not age_group:
            plan["age_group"] = (
                f"{class_level} learners"
            )

    def _validate_lesson_plan(self, plan: dict) -> tuple[bool, list[str]]:
        errors = []

        if not isinstance(plan, dict):
            return False, ["Response is not a JSON object."]

        required_text_fields = [
            "class",
            "subject",
            "topic",
            "prior_knowledge",
            "warmup_activity",
            "learning_note",
            "teacher_activities",
            "student_activities",
            "assessment",
            "plenary",
            "homework",
            "flip_ticket",
        ]

        for key in required_text_fields:
            value = plan.get(key)

            if value is None:
                errors.append(f"Missing field: {key}")
                continue

            if not isinstance(value, str):
                errors.append(
                    f"{key} should be a string, got {type(value).__name__}"
                )
                continue

            if len(value.strip()) < 5:
                errors.append(
                    f"{key} is too short: {repr(value)}"
                )

        # ---------------------------------------------------------
        # Instructional resources
        # ---------------------------------------------------------

        resources = plan.get("instructional_resources")

        if not isinstance(resources, list):
            errors.append(
                "instructional_resources must be a list."
            )
        else:
            cleaned_resources = [
                str(item).strip()
                for item in resources
                if str(item).strip()
            ]

            if len(cleaned_resources) < 2:
                errors.append(
                    "instructional_resources must contain at least 2 items."
                )
            else:
                plan["instructional_resources"] = cleaned_resources[:6]

        # ---------------------------------------------------------
        # Learning objectives
        # ---------------------------------------------------------

        objectives = plan.get("learning_objectives")

        if not isinstance(objectives, dict):
            errors.append(
                "learning_objectives must be an object."
            )
        else:
            for level in (
                "basic",
                "intermediate",
                "advanced"
            ):
                objective = objectives.get(level)

                if objective is None:
                    errors.append(
                        f"Missing learning objective: {level}"
                    )
                    continue

                if not isinstance(objective, str):
                    errors.append(
                        f"Objective '{level}' must be a string."
                    )
                    continue

                if len(objective.strip()) < 5:
                    errors.append(
                        f"Objective '{level}' is too short."
                    )

        return len(errors) == 0, errors
    
    def _repair_lesson_plan(
        self,
        plan: dict,
        subject: str,
        class_level: str,
        topic: str
    ) -> dict:

        if not isinstance(plan, dict):
            plan = {}

        # ---------------------------------------------------------
        # Identity fields
        # ---------------------------------------------------------

        plan["class"] = class_level
        plan["subject"] = subject
        plan["topic"] = topic

        if not self._safe_text(plan.get("subtopic")):
            plan["subtopic"] = topic

        if "date" not in plan:
            plan["date"] = ""

        if "week" not in plan:
            plan["week"] = ""

        if not self._safe_text(plan.get("duration")):
            plan["duration"] = "40 minutes"

        if not self._safe_text(plan.get("age_group")):
            plan["age_group"] = f"{class_level} learners"

        # ---------------------------------------------------------
        # Resources
        # ---------------------------------------------------------

        resources = plan.get("instructional_resources")

        if not isinstance(resources, list):
            resources = []

        resources = [
            str(item).strip()
            for item in resources
            if str(item).strip()
        ]

        if len(resources) < 2:
            resources.extend([
                f"Relevant {subject} textbook",
                f"Visual aid or teaching material related to {topic}",
                "Whiteboard and markers",
                "Learner activity sheet",
            ])

        # Remove duplicates
        unique_resources = []

        for resource in resources:
            if resource not in unique_resources:
                unique_resources.append(resource)

        plan["instructional_resources"] = (
            unique_resources[:6]
        )

        # ---------------------------------------------------------
        # Objectives
        # ---------------------------------------------------------

        objectives = plan.get("learning_objectives")

        if not isinstance(objectives, dict):
            objectives = {}

        if not self._safe_text(objectives.get("basic")):
            objectives["basic"] = (
                f"Define or identify the main concept of {topic}."
            )

        if not self._safe_text(
            objectives.get("intermediate")
        ):
            objectives["intermediate"] = (
                f"Explain and apply the important concepts of "
                f"{topic} using appropriate examples."
            )

        if not self._safe_text(
            objectives.get("advanced")
        ):
            objectives["advanced"] = (
                f"Analyse or solve an appropriate problem "
                f"involving {topic}."
            )

        plan["learning_objectives"] = objectives

        # ---------------------------------------------------------
        # Non-critical text fields
        # ---------------------------------------------------------

        defaults = {
            "prior_knowledge": (
                f"Learners recall previously studied concepts "
                f"in {subject} that relate to {topic}."
            ),

            "warmup_activity": (
                f"The teacher presents a familiar example related "
                f"to {topic}. Learners discuss what they observe "
                f"and connect it to previous learning."
            ),

            "learning_note": (
                f"{topic} is an important concept in {subject}. "
                f"Learners identify its essential meaning, key "
                f"features and appropriate applications."
            ),

            "teacher_activities": (
                f"The teacher introduces {topic}, explains its "
                f"essential concepts using appropriate examples, "
                f"models the required skill, checks understanding "
                f"and guides learners through practice."
            ),

            "student_activities": (
                f"Learners observe examples of {topic}, answer "
                f"questions, participate in guided practice and "
                f"complete an individual application task."
            ),

            "assessment": (
                f"1. Define or identify {topic}. "
                f"2. Explain an important concept related to {topic}. "
                f"3. Apply the concept to an appropriate problem "
                f"or situation."
            ),

            "plenary": (
                f"Learners summarise the key idea of {topic} and "
                f"give one example or application."
            ),

            "homework": (
                f"Write a short explanation of {topic} and complete "
                f"two relevant application questions."
            ),

            "flip_ticket": (
                f"Find one real-life example related to {topic} "
                f"and prepare to explain it in the next lesson."
            ),
        }

        for key, fallback in defaults.items():

            value = plan.get(key)

            # Handle models returning lists instead of strings
            if isinstance(value, list):
                value = "\n".join(
                    str(item).strip()
                    for item in value
                    if str(item).strip()
                )

            # Handle dicts
            elif isinstance(value, dict):
                value = "\n".join(
                    f"{key_name}: {value_text}"
                    for key_name, value_text in value.items()
                    if str(value_text).strip()
                )

            value = self._safe_text(value)

            if not value:
                value = fallback

            plan[key] = value

        return plan
    # ================================================================
    # PRESENTATION GENERATION
    # ================================================================

    def generate_presentation_content(
        self,
        subject: str,
        class_level: str,
        topic: str,
        template_outline: str = "",
        class_notes: str = "",
    ) -> dict:

        subject = self._safe_text(subject)
        class_level = self._safe_text(class_level)
        topic = self._safe_text(topic)
        template_outline = self._safe_text(
            template_outline
        )
        class_notes = self._safe_text(
            class_notes
        )

        if not subject:
            raise ValueError(
                "Subject is required."
            )

        if not class_level:
            raise ValueError(
                "Class level is required."
            )

        if not topic:
            raise ValueError(
                "Topic is required."
            )

        if not self.openai_client:
            return (
                self._generate_dummy_presentation_content(
                    subject,
                    class_level,
                    topic
                )
            )

        template_hint = ""

        if template_outline:
            template_hint = f"""
REFERENCE LESSON TEMPLATE

--- BEGIN TEMPLATE ---
{template_outline}
--- END TEMPLATE ---

Use this only as a content and tone reference.
Do NOT reproduce the lesson-plan table as presentation slides.
"""

        notes_hint = ""

        if class_notes:
            notes_hint = f"""
TEACHER-PROVIDED CLASS NOTES

--- BEGIN CLASS NOTES ---
{class_notes}
--- END CLASS NOTES ---

These notes are the PRIMARY source for the presentation.

Preserve important:
- definitions
- facts
- formulas
- explanations
- examples
- exercises
- terminology

Condense long explanations into classroom-friendly presentation content.
"""

        system_prompt = """
You are an expert teacher, instructional designer and educational
presentation writer.

You transform academic lesson content into clear classroom presentation
material.

Your presentations must:
- be academically accurate
- suit the learners' class level
- avoid overcrowded slides
- use short, meaningful statements
- contain useful examples
- support classroom explanation
- include meaningful assessment
- follow a logical teaching sequence

Return one valid JSON object only.
Do not use Markdown.
"""

        prompt = f"""
Create professional classroom PowerPoint content.

Class: {class_level}
Subject: {subject}
Topic: {topic}

{template_hint}

{notes_hint}

The content must be suitable for actual classroom projection.

Use:
- concise sentences
- clear definitions
- key concepts
- topic-specific examples
- useful terminology
- worked examples where relevant
- classwork
- homework

Avoid lengthy paragraphs.

For Mathematics and Physics:
- include correct equations where relevant
- use correct symbols
- include units
- make worked examples mathematically sound

For Science:
- use correct scientific terminology

For humanities/languages:
- favour examples, interpretation, discussion and application.

Return JSON only with this structure:

{{
    "cover_subtitle": "",
    "overview_line": "",
    "meaning_heading": "",
    "meaning_text": "",
    "examples_heading": "",
    "examples": [],
    "key_terms_heading": "",
    "key_terms": [],
    "worked_examples_heading": "",
    "worked_examples": [],
    "classwork_heading": "",
    "classwork": [],
    "weekend_assignment_heading": "",
    "weekend_assignment": [],
    "closing_line": ""
}}

Requirements:

- examples must contain 2-5 meaningful examples
- key_terms must contain 2-6 useful terms
- worked_examples must contain 1-4 appropriate examples/tasks
- classwork must contain at least 3 topic-specific questions/tasks
- weekend_assignment must contain at least 2 meaningful tasks
- content must suit {class_level}
- do not use generic placeholders
- return JSON only
"""

        try:
            content = self._generate_json(
                system_prompt=system_prompt,
                user_prompt=prompt,
                model=self.PRESENTATION_MODEL,
                temperature=0.50,
                max_tokens=self.PRESENTATION_MAX_TOKENS,
            )

            content = (
                self._normalize_presentation_content(
                    content,
                    subject,
                    class_level,
                    topic
                )
            )

            if not self._validate_presentation_content(
                content
            ):
                raise ValueError(
                    "Generated presentation content failed validation."
                )

            return content

        except Exception as exc:
            logger.exception(
                "Presentation generation failed: %s",
                exc
            )

            return (
                self._generate_dummy_presentation_content(
                    subject,
                    class_level,
                    topic
                )
            )

    def _validate_presentation_content(
        self,
        content: dict
    ) -> bool:

        if not isinstance(content, dict):
            return False

        text_fields = [
            "cover_subtitle",
            "overview_line",
            "meaning_heading",
            "meaning_text",
            "examples_heading",
            "key_terms_heading",
            "worked_examples_heading",
            "classwork_heading",
            "weekend_assignment_heading",
            "closing_line",
        ]

        for field in text_fields:
            value = content.get(field)

            if not isinstance(value, str):
                return False

            if not value.strip():
                return False

        list_fields = [
            "examples",
            "key_terms",
            "worked_examples",
            "classwork",
            "weekend_assignment",
        ]

        for field in list_fields:
            value = content.get(field)

            if not isinstance(value, list):
                return False

            if not any(
                str(item).strip()
                for item in value
            ):
                return False

        return True

    # ================================================================
    # FALLBACK LESSON PLAN
    # ================================================================

    def _generate_dummy_lesson_plan(
        self,
        subject: str,
        class_level: str,
        topic: str
    ) -> dict:

        return {
            "class": class_level,

            "subject": subject,

            "topic": topic,

            "subtopic": topic,

            "date": datetime.now().strftime(
                "%d %B, %Y"
            ),

            "week": "1",

            "duration": "40 minutes",

            "age_group": (
                f"{class_level} learners"
            ),

            "instructional_resources": [
                f"Relevant {subject} textbook",
                f"Visual aid illustrating {topic}",
                "Whiteboard and markers",
                "Learner activity sheet",
            ],

            "learning_objectives": {
                "basic": (
                    f"Define or identify the major idea associated "
                    f"with {topic}."
                ),

                "intermediate": (
                    f"Explain the important features or principles "
                    f"of {topic} using appropriate examples."
                ),

                "advanced": (
                    f"Apply or analyse the principles of {topic} "
                    f"in a relevant classroom or real-life situation."
                ),
            },

            "prior_knowledge": (
                f"Learners recall previously studied concepts in "
                f"{subject} that relate to {topic} and describe "
                f"familiar experiences that can be linked to the "
                f"new lesson."
            ),

            "warmup_activity": (
                f"The teacher presents a familiar example related to "
                f"{topic} and asks learners what they notice, what they "
                f"already know about it, and how it may relate to "
                f"{subject}. Learners discuss briefly with a partner "
                f"before sharing responses."
            ),

            "learning_note": (
                f"Meaning: {topic} is an important concept in {subject}.\n"
                f"Key idea: Learners identify the principal features, "
                f"relationships or rules associated with {topic}.\n"
                f"Application: The concept can be connected to familiar "
                f"classroom or real-life situations.\n"
                f"Example: A suitable example should demonstrate the "
                f"central idea of {topic} clearly."
            ),

            "teacher_activities": (
                f"The teacher links learners' previous knowledge to "
                f"{topic}, introduces its important concepts using "
                f"appropriate examples and instructional resources, "
                f"models the required skill or reasoning, asks checking "
                f"questions, guides paired or group practice, corrects "
                f"misconceptions and provides feedback before assigning "
                f"an individual application task."
            ),

            "student_activities": (
                f"Learners observe the introductory example, share prior "
                f"ideas, record important information about {topic}, "
                f"participate in guided discussion or practice, work "
                f"collaboratively on an application activity, complete "
                f"an individual task and improve their responses after "
                f"feedback."
            ),

            "assessment": (
                f"Evidence should include accurate oral responses, "
                f"completed activities and individual work. "
                f"1. State or define the central idea of {topic}. "
                f"2. Explain two important features or principles of "
                f"{topic}. "
                f"3. Apply the concept to a new relevant example or "
                f"problem."
            ),

            "plenary": (
                f"Learners summarise the meaning and major idea of "
                f"{topic}, give one relevant example or application "
                f"and answer a final check-for-understanding question."
            ),

            "homework": (
                f"Write a concise explanation of {topic}, provide two "
                f"appropriate examples and complete one application "
                f"question. Include a diagram, calculation or labelled "
                f"illustration where appropriate."
            ),

            "flip_ticket": (
                f"Find one example connected to {topic} that could "
                f"help introduce the next related lesson."
            ),
        }

    # ================================================================
    # PRESENTATION NORMALIZATION
    # ================================================================

    def _normalize_presentation_content(
        self,
        content: dict,
        subject: str,
        class_level: str,
        topic: str
    ) -> dict:

        content = content or {}

        def clean_list(
            value: Any,
            fallback: List[str],
            maximum: int = 6
        ) -> List[str]:

            if not isinstance(value, list):
                return fallback

            cleaned = [
                str(item).strip()
                for item in value
                if str(item).strip()
            ]

            return (
                cleaned[:maximum]
                if cleaned
                else fallback
            )

        return {
            "cover_subtitle":
                self._safe_text(
                    content.get("cover_subtitle")
                )
                or (
                    f"Understanding {topic}: concepts, "
                    f"examples and applications"
                ),

            "overview_line":
                self._safe_text(
                    content.get("overview_line")
                )
                or (
                    f"Explore the meaning, important ideas "
                    f"and applications of {topic}."
                ),

            "meaning_heading":
                self._safe_text(
                    content.get("meaning_heading")
                )
                or f"MEANING OF {topic.upper()}",

            "meaning_text":
                self._safe_text(
                    content.get("meaning_text")
                )
                or (
                    f"{topic} is an important concept studied "
                    f"in {subject} at {class_level} level."
                ),

            "examples_heading":
                self._safe_text(
                    content.get("examples_heading")
                )
                or "EXAMPLES",

            "examples":
                clean_list(
                    content.get("examples"),
                    [
                        (
                            f"A familiar classroom or real-life "
                            f"example related to {topic}."
                        ),
                        (
                            f"A second example demonstrating "
                            f"the main idea of {topic}."
                        ),
                    ],
                    maximum=5
                ),

            "key_terms_heading":
                self._safe_text(
                    content.get("key_terms_heading")
                )
                or (
                    f"KEY TERMS IN {topic.upper()}"
                ),

            "key_terms":
                clean_list(
                    content.get("key_terms"),
                    [
                        f"Meaning of {topic}",
                        f"Application of {topic}",
                    ],
                    maximum=6
                ),

            "worked_examples_heading":
                self._safe_text(
                    content.get(
                        "worked_examples_heading"
                    )
                )
                or "WORKED EXAMPLES",

            "worked_examples":
                clean_list(
                    content.get("worked_examples"),
                    [
                        (
                            f"Explain {topic} using an "
                            f"appropriate example."
                        )
                    ],
                    maximum=4
                ),

            "classwork_heading":
                self._safe_text(
                    content.get("classwork_heading")
                )
                or "CLASSWORK",

            "classwork":
                clean_list(
                    content.get("classwork"),
                    [
                        f"1. Define {topic}.",
                        (
                            f"2. State two important features "
                            f"of {topic}."
                        ),
                        (
                            f"3. Apply {topic} to one relevant "
                            f"example."
                        ),
                    ],
                    maximum=6
                ),

            "weekend_assignment_heading":
                self._safe_text(
                    content.get(
                        "weekend_assignment_heading"
                    )
                )
                or "HOMEWORK",

            "weekend_assignment":
                clean_list(
                    content.get(
                        "weekend_assignment"
                    ),
                    [
                        (
                            f"1. Write a short explanation "
                            f"of {topic}."
                        ),
                        (
                            f"2. Give two examples or "
                            f"applications of {topic}."
                        ),
                    ],
                    maximum=6
                ),

            "closing_line":
                self._safe_text(
                    content.get("closing_line")
                )
                or "REVIEW • PRACTISE • APPLY",
        }

    # ================================================================
    # FALLBACK PRESENTATION
    # ================================================================

    def _generate_dummy_presentation_content(
        self,
        subject: str,
        class_level: str,
        topic: str
    ) -> dict:

        return self._normalize_presentation_content(
            {
                "cover_subtitle": (
                    f"Understanding {topic}: meaning, "
                    f"examples and applications"
                ),

                "overview_line": (
                    f"A classroom introduction to {topic} "
                    f"for {class_level} {subject}."
                ),

                "meaning_heading":
                    f"MEANING OF {topic.upper()}",

                "meaning_text": (
                    f"{topic} is an important concept in "
                    f"{subject}. This lesson introduces its "
                    f"main meaning, important features and "
                    f"applications."
                ),

                "examples_heading":
                    "EXAMPLES",

                "examples": [
                    (
                        f"A familiar real-life example "
                        f"connected to {topic}."
                    ),
                    (
                        f"A classroom example demonstrating "
                        f"the main concept of {topic}."
                    ),
                ],

                "key_terms_heading":
                    f"KEY TERMS IN {topic.upper()}",

                "key_terms": [
                    f"Meaning of {topic}",
                    f"Major principle related to {topic}",
                    f"Application of {topic}",
                ],

                "worked_examples_heading":
                    "GUIDED EXAMPLES",

                "worked_examples": [
                    (
                        f"Explain the main idea of {topic} "
                        f"using one appropriate example."
                    ),
                    (
                        f"Apply the concept of {topic} to "
                        f"a simple classroom problem."
                    ),
                ],

                "classwork_heading":
                    "CLASSWORK",

                "classwork": [
                    f"1. Define {topic}.",
                    (
                        f"2. State two important ideas "
                        f"associated with {topic}."
                    ),
                    (
                        f"3. Give one appropriate "
                        f"application of {topic}."
                    ),
                ],

                "weekend_assignment_heading":
                    "HOMEWORK",

                "weekend_assignment": [
                    (
                        f"1. Write a short explanatory note "
                        f"on {topic}."
                    ),
                    (
                        f"2. Give two examples or applications "
                        f"of {topic}."
                    ),
                ],

                "closing_line":
                    "REVIEW • PRACTISE • APPLY",
            },
            subject,
            class_level,
            topic,
        )
        
        