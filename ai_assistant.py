"""
AI Assistant module for Project Update Studio.
Handles all DeepSeek API calls for interviews and content generation.
"""

import requests
import json
import re
from typing import Optional, List, Dict, Any
from utils import get_string

DEEPSEEK_API_URL = "https://api.deepseek.com/v1/chat/completions"
DEEPSEEK_MODEL = "deepseek-chat"


class AIAssistant:
    """Interface to DeepSeek API."""

    def __init__(self, api_key: str, language: str = "en"):
        """Initialize AI assistant with API key and language."""
        self.api_key = api_key
        self.language = language
        self.conversation_history = []

    def test_connection(self) -> bool:
        """Test if the API key is valid."""
        if not self.api_key:
            return False

        try:
            response = requests.post(
                DEEPSEEK_API_URL,
                headers={
                    "Authorization": f"Bearer {self.api_key}",
                    "Content-Type": "application/json",
                },
                json={
                    "model": DEEPSEEK_MODEL,
                    "messages": [
                        {
                            "role": "user",
                            "content": "Hello",
                        }
                    ],
                    "max_tokens": 10,
                },
                timeout=10,
            )
            return response.status_code == 200
        except Exception as e:
            print(f"Connection test failed: {e}")
            return False

    def _build_system_prompt(
        self, project_brief: str = "", previous_updates: str = "", context: str = ""
    ) -> str:
        """Build the system prompt with context."""
        language_name = "German" if self.language == "de" else "English"

        system_prompt = f"""You are a professional project management assistant helping to create weekly status updates.

**Language:** {language_name}

**Your role:**
- Ask ONE clear, focused question at a time during interviews
- Be conversational and professional
- Respect the user's project context
- Extract structured information from answers
- Generate concise executive summaries
- Be helpful and efficient

**Project Context:**
{project_brief if project_brief else "No project brief available yet."}

**Previous Updates Summary:**
{previous_updates if previous_updates else "This is the first weekly update for this project."}

**Additional Context:**
{context if context else ""}

**Important instructions:**
- Use {language_name} for all responses
- Ask questions in a natural, conversational way
- When extracting data, be thorough but concise
- Respond to answers with encouragement and the next question
- At the end of the interview, provide structured JSON with all collected data
"""
        return system_prompt

    def ask_question(
        self,
        question: str,
        system_prompt: str,
        conversation_history: List[Dict] = None,
    ) -> str:
        """Send a question and get a response."""
        if conversation_history is None:
            conversation_history = []

        # Build messages
        messages = [{"role": "user", "content": question}]

        # Add conversation history if provided
        if conversation_history:
            messages = conversation_history + messages

        try:
            response = requests.post(
                DEEPSEEK_API_URL,
                headers={
                    "Authorization": f"Bearer {self.api_key}",
                    "Content-Type": "application/json",
                },
                json={
                    "model": DEEPSEEK_MODEL,
                    "messages": [{"role": "system", "content": system_prompt}]
                    + messages,
                    "temperature": 0.7,
                    "max_tokens": 1000,
                },
                timeout=30,
            )

            if response.status_code != 200:
                error_msg = response.json().get("error", {}).get("message", str(response.status_code))
                raise Exception(f"API error: {error_msg}")

            data = response.json()
            content = data["choices"][0]["message"]["content"]
            return content

        except requests.exceptions.Timeout:
            raise Exception(
                "Request timed out. The DeepSeek API is taking too long to respond."
            )
        except requests.exceptions.ConnectionError:
            raise Exception(
                "Connection error. Please check your internet connection and API key."
            )
        except Exception as e:
            raise Exception(f"AI Assistant error: {str(e)}")

    def _extract_json_from_response(self, content: str) -> Any:
        """Extract and parse JSON payload from an AI response."""
        if not content or not content.strip():
            raise Exception("AI returned empty response. Please try again.")

        try:
            return json.loads(content)
        except json.JSONDecodeError:
            # Fallback for responses that wrap JSON with extra prose/markdown.
            json_match = re.search(r"\{.*\}", content, re.DOTALL)
            if json_match:
                return json.loads(json_match.group())
            raise

    def generate_project_onboarding_questions(self) -> List[str]:
        """Generate the list of onboarding interview questions."""
        if self.language == "de":
            return [
                "Welches Problem oder welche Ineffizienz behandelt dieses Projekt?",
                "Wie sieht der gewünschte Endzustand oder das Ziel aus?",
                "Wie ist die aktuelle Situation / die Basiskennzahlen vor dem Projekt?",
                "Wie ist der Zeitplan und welche sind die Hauptmeilensteine?",
                "Welche sind die wichtigsten Risiken oder Abhängigkeiten?",
                "Wer sind die wichtigsten Stakeholder außerhalb des Teams?",
                "Wie sieht der Erfolg aus — wie werden Sie ihn messen?",
                "Ist an diesem Projekt ein Budget beteiligt?",
            ]
        else:
            return [
                "What problem or inefficiency does this project address?",
                "What is the desired end state or goal?",
                "What is the baseline (current situation / metrics before the project)?",
                "What is the timeline / key milestones?",
                "What are the main risks or dependencies?",
                "Who are the key stakeholders (outside the team)?",
                "What does success look like — how will you measure it?",
                "Is there a budget involved?",
            ]

    def generate_weekly_update_questions(self) -> List[str]:
        """Generate the list of weekly update interview questions."""
        if self.language == "de":
            return [
                "Welche Aufgaben wurden diese Woche abgeschlossen? Was waren die Ergebnisse?",
                "Wurden Meilensteine erreicht?",
                "Wie würden Sie den Fortschritt diese Woche beschreiben (grün/am Plan, gelb/gefährdet, rot/blockiert)?",
                "Gibt es KPI oder Leistungskennzahlen zu berichten?",
                "Gibt es Blockierungen, Verzögerungen oder Risiken?",
                "Welcher Mitigation ist geplant?",
                "Liegt das Projekt im Budget? Gibt es Budget-Updates?",
                "Welche Aufgaben sind für nächste Woche geplant? Wer ist verantwortlich?",
                "Gibt es Managemententscheidungen, die getroffen werden müssen?",
                "Gibt es Stakeholder-Updates oder Maßnahmen?",
            ]
        else:
            return [
                "What tasks were completed this week? What were the results?",
                "Were any milestones reached?",
                "How would you describe progress this week (green/on track, amber/at risk, red/blocked)?",
                "Are there any KPIs or metrics to report?",
                "Are there any blockers, delays, or risks encountered?",
                "What is the mitigation plan?",
                "Is the project on budget? Any budget updates?",
                "What are the planned tasks for next week? Who is assigned to each?",
                "Are there any management decisions that need to be made?",
                "Any stakeholder updates or actions needed?",
            ]

    def extract_weekly_update_data(
        self,
        conversation_summary: str,
        system_prompt: str,
    ) -> Dict:
        """Extract structured data from conversation summary."""
        extraction_prompt = f"""TASK: Extract project update data from the following conversation and return ONLY valid JSON.

CONVERSATION:
{conversation_summary}

INSTRUCTIONS:
1. Extract all relevant information from the conversation above
2. Return ONLY a valid JSON object (no markdown, no explanation, no extra text)
3. If information is missing from the conversation, use empty arrays [] for lists and empty strings "" for fields
4. Use these exact field names and structure

JSON OUTPUT REQUIRED:
{{"tasks_completed": [{{"task": "string", "result": "string", "owner": "string"}}], "next_tasks": [{{"task": "string", "owner": "string", "due_date": "string"}}], "rag_status": "green|amber|red", "rag_reason": "string", "management_decisions": [{{"decision": "string", "urgency": "urgent|this_week|when_possible", "context": "string"}}], "risks_blockers": [{{"issue": "string", "impact": "string", "mitigation": "string"}}], "budget_status": "on_track|over|under|not_applicable", "budget_notes": "string", "kpi_updates": [{{"metric": "string", "value": "string", "trend": "string"}}], "milestone_hit": ["string"], "ai_summary": "string"}}

Return ONLY the JSON object, nothing else."""

        try:
            response = requests.post(
                DEEPSEEK_API_URL,
                headers={
                    "Authorization": f"Bearer {self.api_key}",
                    "Content-Type": "application/json",
                },
                json={
                    "model": DEEPSEEK_MODEL,
                    "messages": [
                        {"role": "system", "content": system_prompt},
                        {"role": "user", "content": extraction_prompt},
                    ],
                    "temperature": 0.3,
                    "max_tokens": 2000,
                },
                timeout=30,
            )

            if response.status_code != 200:
                error_msg = response.json().get("error", {}).get("message", str(response.status_code))
                raise Exception(f"API error: {error_msg}")

            data = response.json()
            content = data["choices"][0]["message"]["content"]
            extracted_data = self._extract_json_from_response(content)
            if not isinstance(extracted_data, dict):
                raise Exception("AI response format was invalid (expected JSON object).")
            return extracted_data

        except json.JSONDecodeError as e:
            raise Exception(f"Failed to parse AI response as JSON: {e}")
        except requests.exceptions.Timeout:
            raise Exception("Request timed out. Please try again.")
        except Exception as e:
            raise Exception(f"Data extraction error: {str(e)}")

    def polish_document_inputs(
        self,
        project: Dict,
        updates: List[Dict],
        system_prompt: str,
    ) -> Dict:
        """Rewrite project and update inputs into clean, publication-ready text."""
        language_name = "German" if self.language == "de" else "English"

        payload = {
            "project": {
                "name": project.get("name", ""),
                "description": project.get("description", ""),
                "goal": project.get("goal", ""),
                "background": project.get("background", ""),
            },
            "updates": updates,
        }

        polishing_prompt = f"""TASK: Rewrite and normalize project documentation input.

LANGUAGE: {language_name}

INPUT JSON:
{json.dumps(payload, ensure_ascii=False)}

INSTRUCTIONS:
1. Keep the same number of updates and preserve update order.
2. Rewrite all free-text fields into clear executive-ready language.
3. Clean inconsistent formatting and remove raw JSON-looking text fragments.
4. Keep factual meaning; do not invent data.
5. Return ONLY valid JSON matching this schema:

{{
  "project": {{
    "description": "string",
    "goal": "string",
    "background": "string"
  }},
  "updates": [
    {{
      "ai_summary": "string",
      "tasks_completed": [{{"task": "string", "result": "string", "owner": "string"}}],
      "next_tasks": [{{"task": "string", "owner": "string", "due_date": "string"}}],
      "management_decisions": [{{"decision": "string", "urgency": "urgent|this_week|when_possible", "context": "string"}}],
      "risks_blockers": [{{"issue": "string", "impact": "string", "mitigation": "string"}}],
      "kpi_updates": [{{"metric": "string", "value": "string", "trend": "string"}}],
      "milestone_hit": ["string"],
      "budget_notes": "string"
    }}
  ]
}}

Return ONLY the JSON object, no explanation."""

        try:
            response = requests.post(
                DEEPSEEK_API_URL,
                headers={
                    "Authorization": f"Bearer {self.api_key}",
                    "Content-Type": "application/json",
                },
                json={
                    "model": DEEPSEEK_MODEL,
                    "messages": [
                        {"role": "system", "content": system_prompt},
                        {"role": "user", "content": polishing_prompt},
                    ],
                    "temperature": 0.3,
                    "max_tokens": 3000,
                },
                timeout=45,
            )

            if response.status_code != 200:
                error_msg = response.json().get("error", {}).get("message", str(response.status_code))
                raise Exception(f"API error: {error_msg}")

            data = response.json()
            content = data["choices"][0]["message"]["content"]
            polished = self._extract_json_from_response(content)

            if not isinstance(polished, dict):
                raise Exception("AI returned invalid format for document polishing.")

            return polished

        except Exception as e:
            raise Exception(f"Document input polishing error: {str(e)}")

    def rewrite_text_fields(
        self,
        fields: Dict[str, str],
        context: str = "",
        system_prompt: str = "",
    ) -> Dict[str, str]:
        """Rewrite plain text fields in concise professional language."""
        if not isinstance(fields, dict):
            raise Exception("Fields payload must be a dictionary.")

        language_name = "German" if self.language == "de" else "English"
        normalized_fields = {
            key: value.strip() if isinstance(value, str) else ""
            for key, value in fields.items()
        }

        if not any(normalized_fields.values()):
            return normalized_fields

        rewrite_prompt = f"""TASK: Rewrite the following user-entered fields.

LANGUAGE: {language_name}
STYLE: Professional, short, and precise.
CONTEXT: {context or 'Project management updates and documentation'}

INPUT JSON:
{json.dumps(normalized_fields, ensure_ascii=False)}

INSTRUCTIONS:
1. Keep the same keys and return ONLY a valid JSON object.
2. Preserve factual meaning. Do not invent data.
3. Improve grammar, clarity, and concise executive tone.
4. Keep person names unchanged where possible.

Return ONLY JSON."""

        effective_system_prompt = system_prompt or self._build_system_prompt(
            context="You rewrite user-entered project content into concise professional language."
        )

        try:
            response = requests.post(
                DEEPSEEK_API_URL,
                headers={
                    "Authorization": f"Bearer {self.api_key}",
                    "Content-Type": "application/json",
                },
                json={
                    "model": DEEPSEEK_MODEL,
                    "messages": [
                        {"role": "system", "content": effective_system_prompt},
                        {"role": "user", "content": rewrite_prompt},
                    ],
                    "temperature": 0.3,
                    "max_tokens": 1500,
                },
                timeout=30,
            )

            if response.status_code != 200:
                error_msg = response.json().get("error", {}).get("message", str(response.status_code))
                raise Exception(f"API error: {error_msg}")

            content = response.json()["choices"][0]["message"]["content"]
            rewritten = self._extract_json_from_response(content)
            if not isinstance(rewritten, dict):
                raise Exception("AI returned invalid rewrite format.")

            merged = dict(normalized_fields)
            for key in normalized_fields:
                value = rewritten.get(key)
                if isinstance(value, str) and value.strip():
                    merged[key] = value.strip()
            return merged

        except Exception as e:
            raise Exception(f"Text rewrite error: {str(e)}")

    def polish_single_weekly_update(
        self,
        update_payload: Dict,
        system_prompt: str,
    ) -> Dict:
        """Polish one weekly update payload using the document polishing schema."""
        polished_payload = self.polish_document_inputs(
            {"name": "", "description": "", "goal": "", "background": ""},
            [update_payload],
            system_prompt,
        )

        updates = polished_payload.get("updates", []) if isinstance(polished_payload, dict) else []
        if not updates or not isinstance(updates[0], dict):
            raise Exception("AI returned invalid weekly update polishing payload.")
        return updates[0]

    def generate_project_closure_summary(
        self,
        project_brief: str,
        all_updates_summary: str,
        system_prompt: str,
    ) -> str:
        """Generate a final project closure summary."""
        closure_prompt = f"""You are writing the final closure summary for a completed project.

Project Background:
{project_brief}

All Updates Summary:
{all_updates_summary}

Please write a professional project closure summary (approximately 300 words) that includes:
1. Project overview and objectives
2. Key achievements and outcomes
3. Final status and deliverables
4. Key learnings and recommendations
5. Closing statement

Write in {"German" if self.language == "de" else "English"}.
Be professional, concise, and suitable for executive stakeholders."""

        try:
            response = requests.post(
                DEEPSEEK_API_URL,
                headers={
                    "Authorization": f"Bearer {self.api_key}",
                    "Content-Type": "application/json",
                },
                json={
                    "model": DEEPSEEK_MODEL,
                    "messages": [
                        {"role": "system", "content": system_prompt},
                        {"role": "user", "content": closure_prompt},
                    ],
                    "temperature": 0.7,
                    "max_tokens": 1500,
                },
                timeout=30,
            )

            if response.status_code != 200:
                error_msg = response.json().get("error", {}).get("message", str(response.status_code))
                raise Exception(f"API error: {error_msg}")

            data = response.json()
            return data["choices"][0]["message"]["content"]

        except Exception as e:
            raise Exception(f"Closure summary generation error: {str(e)}")
