import os
import json
from typing import Dict, List, Any, Optional
import numpy as np
from datetime import datetime, timezone
from openai import OpenAI  # Using OpenAI client with Hugging Face base_url
from dotenv import load_dotenv
from pydantic import BaseModel, Field
from collections import Counter

load_dotenv()

class StudentAnalysis(BaseModel):
    strengths: List[str] = Field(description="List of student's academic strengths")
    weaknesses: List[str] = Field(description="List of student's academic weaknesses")
    recommendations: List[str] = Field(description="Recommendations for improvement")
    comment: str = Field(description="AI-generated comment for the student")

class AIAnalysisService:
    def __init__(self):
        self.huggingface_api_key = os.getenv("HUGGINGFACE_API_KEY")
        self.openai_client = None
        
        # Default models to try (free and open source)
        self.models = [
            "deepseek-ai/DeepSeek-V3.2-Exp:novita",  # DeepSeek (powerful)
            "meta-llama/Llama-3.3-70B-Instruct",     # Llama 3.3
            "google/gemma-2-9b-it",                  # Gemma 2
            "Qwen/Qwen2.5-7B-Instruct",              # Qwen 2.5
            "microsoft/Phi-3.5-mini-instruct",       # Phi 3.5 Mini
        ]
        
        if self.huggingface_api_key:
            try:
                # Initialize OpenAI client with Hugging Face base URL
                self.openai_client = OpenAI(
                    base_url="https://router.huggingface.co/v1",
                    api_key=self.huggingface_api_key,
                )
                print("Hugging Face client initialized successfully")
                
                # Test connection with first available model
                self._test_connection()
                
            except Exception as e:
                print(f"Error initializing Hugging Face client: {e}")
                self.openai_client = None
        else:
            print("Warning: HUGGINGFACE_API_KEY not found. AI features will use rule-based analysis.")
    
    def _test_connection(self):
        """Test connection to Hugging Face API"""
        try:
            # Try a simple request with each model until one works
            for model in self.models[:2]:  # Try first two models
                try:
                    test_response = self.openai_client.chat.completions.create(
                        model=model,
                        messages=[
                            {"role": "user", "content": "Say 'Connected' if you can read this."}
                        ],
                        max_tokens=10,
                        timeout=10
                    )
                    print(f"[OK] Connected to model: {model}")
                    self.preferred_model = model
                    return
                except Exception as e:
                    print(f"  Model {model} not available: {e}")
                    continue
            
            # If no model worked, try to get available models
            try:
                models_list = self.openai_client.models.list()
                available_models = [model.id for model in models_list.data[:5]]
                print(f"Available models: {available_models}")
                if available_models:
                    self.preferred_model = available_models[0]
                    print(f"[OK] Using model: {self.preferred_model}")
            except:
                print("Could not fetch available models")
                self.preferred_model = self.models[0]
                
        except Exception as e:
            print(f"Connection test failed: {e}")
            self.openai_client = None
 
    def get_available_models(self) -> List[str]:
        """Get list of available Hugging Face models"""
        if self.openai_client:
            try:
                models = self.openai_client.models.list()
                return [model.id for model in models.data[:10]]  # First 10 models
            except Exception as e:
                print(f"Could not fetch available models: {e}")
                return self.models
        return self.models
    
        
    def generate_lesson_plan(self, subject: str, class_level: str, topic: str, template_outline: str = "") -> dict:
        """Generate a complete lesson plan using HuggingFace AI"""
        
        if not self.openai_client:
            # Fallback to a basic template
            return self._generate_dummy_lesson_plan(subject, class_level, topic)
        
        template_hint = f"\nUploaded template outline:\n{template_outline}\n" if template_outline else ""

        prompt = f"""Generate a detailed, professional lesson plan for a {class_level} class studying {subject}. The topic is "{topic}".
    {template_hint}
    
    Use the exact structure below and return ONLY valid JSON with no extra text. The JSON must have these keys: 
    - class, subject, topic, subtopic (same as topic), date (today's date), week (1), duration ("Forty Minutes"), age_group ("appropriate for {class_level}")
    - instructional_resources (list of resources like textbooks, websites, videos)
    - learning_objectives: object with keys "basic", "intermediate", "advanced" (each a string describing what students will be able to do)
    - prior_knowledge (string, what students already know)
    - warmup_activity (string, a 2‑3 sentence engaging starter activity)
    - learning_note (string, the main content – definition, branches, examples, etc. Use bullet points)
    - teacher_activities (string, what teacher does during the lesson)
    - student_activities (string, what students do – group work, individual practice)
    - assessment (string, how teacher evaluates understanding)
    - plenary (string, summary of key points)
    - homework (string, take‑home task)
    - flip_ticket (string, next topic teaser)
    
    Make the plan pedagogically sound, differentiated, and aligned with modern teaching standards. Use bullet points (•) in strings where appropriate.
    """
        
        models_to_try = [getattr(self, 'preferred_model', self.models[0])] + self.models[:3]
        for model in models_to_try:
            try:
                response = self.openai_client.chat.completions.create(
                    model=model,
                    messages=[
                        {"role": "system", "content": "You are an expert curriculum designer. Generate high‑quality lesson plans in valid JSON."},
                        {"role": "user", "content": prompt}
                    ],
                    temperature=0.7,
                    max_tokens=2000,
                    response_format={"type": "json_object"}
                )
                content = response.choices[0].message.content
                plan = json.loads(content)
                # Validate required keys
                required = ['class', 'subject', 'topic', 'learning_objectives', 'prior_knowledge', 
                           'warmup_activity', 'learning_note', 'teacher_activities', 'student_activities',
                           'assessment', 'plenary', 'homework']
                for key in required:
                    if key not in plan:
                        plan[key] = f"Auto‑generated {key.replace('_',' ')}"
                return plan
            except Exception as e:
                print(f"Model {model} failed for lesson plan: {e}")
                continue
        
        return self._generate_dummy_lesson_plan(subject, class_level, topic)
    
    def _generate_dummy_lesson_plan(self, subject, class_level, topic):
        """Fallback when AI fails"""
        return {
            "class": class_level,
            "subject": subject,
            "topic": topic,
            "subtopic": topic,
            "date": datetime.now().strftime("%d %B, %Y"),
            "week": "1",
            "duration": "Forty Minutes",
            "age_group": f"{class_level} students",
            "instructional_resources": ["Textbook", "Whiteboard", "Markers", "Online videos (if available)"],
            "learning_objectives": {
                "basic": f"Define {topic} in simple terms.",
                "intermediate": f"Explain the key concepts of {topic} with examples.",
                "advanced": f"Analyse how {topic} applies to real‑world situations."
            },
            "prior_knowledge": "Students have basic understanding of scientific inquiry and everyday observations.",
            "warmup_activity": f"Ask students: 'What comes to mind when you hear the word {topic}?' Write their responses on the board.",
            "learning_note": f"Definition of {topic}.\n• Key principles.\n• Examples from daily life.\n• Importance in {subject}.",
            "teacher_activities": f"Present the definition and key points using visuals. Explain with examples. Guide discussion.",
            "student_activities": f"Group discussion: In small groups, list applications of {topic}. Share with class. Complete worksheet.",
            "assessment": "Observe group work. Review worksheets. Ask oral questions.",
            "plenary": f"Summarise main points: what {topic} is, why it matters, and where it is seen.",
            "homework": f"Find one real‑world example of {topic} and write a short paragraph explaining it.",
            "flip_ticket": "Next topic: Applications of " + topic
        }
        


# Example usage
if __name__ == "__main__":
    # Create .env file with: HUGGINGFACE_API_KEY=your_key_here
    service = AIAnalysisService()
    
    # Example student data
    student_data = {
        "first_name": "John",
        "last_name": "Doe",
        "class_name": "10A",
        "score": 85,
        "percentage": 85.0
    }
    
    exam_data = {
        "title": "Mathematics Midterm",
        "subject_name": "Mathematics",
        "total_marks": 100
    }
    
    question_responses = [
        {"question_id": 1, "question_text": "Solve 2x + 5 = 15", "question_type": "algebra", "is_correct": True},
        {"question_id": 2, "question_text": "Calculate area of circle", "question_type": "geometry", "is_correct": True},
        {"question_id": 3, "question_text": "Differentiate x^2", "question_type": "calculus", "is_correct": False},
        {"question_id": 4, "question_text": "Probability question", "question_type": "statistics", "is_correct": True},
        {"question_id": 5, "question_text": "Trigonometry problem", "question_type": "trigonometry", "is_correct": True},
    ]
    
    # Analyze student performance
    result = service.analyze_student_performance(student_data, exam_data, question_responses)
    print("\nStudent Analysis Result:")
    print(json.dumps(result, indent=2))
    
    # Get available models
    models = service.get_available_models()
    print(f"\nAvailable models: {models}")
