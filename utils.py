"""
Utility module for Project Update Studio.
Shared helpers, language strings, and formatting functions.
"""

from datetime import datetime
from dateutil import parser as date_parser
import json

# Color palette (Philips Ocean Blue)
COLORS = {
    "primary_blue": "#0B5ED7",
    "dark_blue": "#003087",
    "light_blue_accent": "#E8F0FE",
    "white": "#FFFFFF",
    "text_dark": "#1A1A2E",
    "success_green": "#28A745",
    "warning_amber": "#FFC107",
    "danger_red": "#DC3545",
}

# Language strings
STRINGS = {
    "en": {
        "app_title": "Project Update Studio",
        "app_subtitle": "Weekly Status Updates & Documentation Generator",
        # Sidebar
        "nav_dashboard": "Dashboard",
        "nav_projects": "Projects",
        "nav_weekly_update": "Weekly Update",
        "nav_slide_generator": "Slide Generator",
        "nav_gantt_chart": "Gantt Chart",
        "nav_final_docs": "Final Documentation",
        "nav_settings": "Settings",
        # Dashboard
        "dashboard_title": "Dashboard",
        "total_projects": "Total Projects",
        "on_track": "On Track",
        "at_risk": "At Risk",
        "blocked": "Blocked",
        "weeks_of_updates": "weeks of updates",
        "team_members": "Team members",
        "last_updated": "Last updated",
        "new_update": "New Update",
        "generate_slide": "Generate Slide",
        # Projects
        "projects_title": "Projects",
        "new_project": "New Project",
        "project_name": "Project Name",
        "project_description": "Description",
        "project_goal": "Goal",
        "project_background": "Background",
        "select_language": "Language",
        "team": "Team",
        "team_member_name": "Team Member Name",
        "role": "Role",
        "add_team_member": "Add Team Member",
        "create_project": "Create Project",
        "edit_project": "Edit Project",
        "delete_project": "Delete Project",
        "archive_project": "Archive Project",
        "project_brief": "Project Brief",
        "ai_interview": "AI Project Onboarding Interview",
        "next_question": "Next Question",
        "skip_question": "Skip",
        "complete_interview": "Complete Interview",
        "project_brief_confirmation": "Please review the project brief below and confirm to save:",
        "confirm_save": "Confirm & Save",
        "save_changes": "Save Changes",
        "cancel_edit": "Cancel Edit",
        "editing_project": "Editing Project",
        # Weekly Update
        "weekly_update_title": "Weekly Update",
        "select_project": "Select Project",
        "project_summary": "Project Summary",
        "previous_update": "Previous Update",
        "ai_interview_start": "AI Interview - Answer the questions to generate this week's update",
        "your_answer": "Your answer",
        "submit_answer": "Submit Answer",
        "interview_complete": "Interview Complete",
        "review_and_edit": "Review and edit the extracted information below:",
        "save_update": "Save Update",
        "update_saved": "Weekly update saved successfully!",
        "duplicate_week_warning": "An update for this week already exists. Click below to edit it instead.",
        "edit_existing": "Edit Existing Update",
        "cancel_week_edit": "Cancel Week Edit",
        "editing_existing_week": "You are editing this week's existing update.",
        # Slide Generator
        "slide_generator_title": "Slide Generator",
        "select_week": "Select Week",
        "generate_slide": "Generate Slide",
        "slide_generated": "Slide generated successfully!",
        "download_slide": "Download Slide",
        # Gantt Chart
        "gantt_chart_title": "Gantt Chart Generator",
        "generate_gantt": "Generate Gantt Chart",
        "gantt_generated": "Gantt chart generated successfully!",
        "download_gantt": "Download Gantt Chart",
        "gantt_tasks": "Gantt Chart Tasks",
        "add_task": "Add Task",
        "task_name": "Task Name",
        "start_date": "Start Date",
        "end_date": "End Date",
        "milestone_date": "Milestone Date",
        "task_type": "Type",
        "regular_task": "Regular Task",
        "milestone": "Milestone",
        "auto_generate_gantt": "Auto-Generate from Latest Update",
        "gantt_no_tasks": "No gantt chart tasks defined. Auto-generate from the latest weekly update or add tasks manually.",
        "edit_gantt_task": "Edit Task",
        "delete_gantt_task": "Delete Task",
        # Final Documentation
        "final_docs_title": "Final Documentation",
        "generate_documentation": "Generate Documentation",
        "doc_generated": "Documentation generated successfully!",
        "download_documentation": "Download Documentation",
        # Settings
        "settings_title": "Settings",
        "api_key": "DeepSeek API Key",
        "api_key_placeholder": "Enter your DeepSeek API key",
        "test_connection": "Test Connection",
        "connection_success": "✅ Connection successful!",
        "connection_failed": "❌ Connection failed. Check your API key.",
        "your_name": "Your Name",
        "language": "Language",
        "export_folder": "Export Folder",
        "reset_data": "Reset All Data",
        "reset_confirm": "Are you sure? This will delete all projects and updates.",
        "setting_saved": "Setting saved!",
        # Messages
        "missing_api_key": "⚠️ DeepSeek API key is not set. Please go to Settings and add your API key to use AI features.",
        "api_key_required": "API Key Required",
        "go_to_settings": "Go to Settings",
        "no_data": "No data available",
        "no_projects_hint": "Create a project first to continue.",
        "no_updates_hint": "Create at least one weekly update first.",
        "missing_node": "Node.js is required for slide and gantt generation.",
        "missing_pptxgen": "pptxgenjs dependency is missing. Run npm install.",
        "loading": "Loading...",
        "error": "Error",
        "success": "Success",
        "rewrite_with_ai": "Rewrite With AI",
        "ai_suggestions_ready": "AI suggestions are ready. Review and apply if desired.",
        "apply_ai_suggestions": "Apply AI Suggestions",
        "discard_ai_suggestions": "Discard Suggestions",
        "rewrite_failed": "AI rewrite failed. You can continue with your current text.",
        # Data fields (for forms)
        "tasks_completed": "Tasks Completed",
        "task": "Task",
        "result": "Result",
        "owner": "Owner",
        "next_steps": "Next Steps",
        "management_decisions": "Management Decisions",
        "decision": "Decision",
        "urgency": "Urgency",
        "context": "Context",
        "risks_blockers": "Risks & Blockers",
        "issue": "Issue",
        "impact": "Impact",
        "mitigation": "Mitigation",
        "budget_status": "Budget Status",
        "budget_notes": "Budget Notes",
        "on_track": "On Track",
        "over_budget": "Over Budget",
        "under_budget": "Under Budget",
        "n_a": "Not Applicable",
        "kpi_updates": "KPI Updates",
        "metric": "Metric",
        "value": "Value",
        "trend": "Trend",
        "stakeholder_notes": "Stakeholder Communications",
        "weekly_summary": "Weekly Summary",
        "executive_summary": "Executive Summary",
        # RAG Status
        "rag_green": "On Track",
        "rag_amber": "At Risk",
        "rag_red": "Blocked",
    },
    "de": {
        "app_title": "Project Update Studio",
        "app_subtitle": "Wöchentliche Statusberichte & Dokumentationsgenerator",
        # Sidebar
        "nav_dashboard": "Dashboard",
        "nav_projects": "Projekte",
        "nav_weekly_update": "Wöchentliche Aktualisierung",
        "nav_slide_generator": "Foliengenerator",
        "nav_gantt_chart": "Gantt-Diagramm",
        "nav_final_docs": "Abschlussdokumentation",
        "nav_settings": "Einstellungen",
        # Dashboard
        "dashboard_title": "Dashboard",
        "total_projects": "Gesamtprojekte",
        "on_track": "Im Plan",
        "at_risk": "Gefährdet",
        "blocked": "Blockiert",
        "weeks_of_updates": "Wochen mit Aktualisierungen",
        "team_members": "Teamangehörige",
        "last_updated": "Zuletzt aktualisiert",
        "new_update": "Neue Aktualisierung",
        "generate_slide": "Folie erstellen",
        # Projects
        "projects_title": "Projekte",
        "new_project": "Neues Projekt",
        "project_name": "Projektname",
        "project_description": "Beschreibung",
        "project_goal": "Ziel",
        "project_background": "Hintergrund",
        "select_language": "Sprache",
        "team": "Team",
        "team_member_name": "Name des Teamangehörigen",
        "role": "Rolle",
        "add_team_member": "Teamangehörigen hinzufügen",
        "create_project": "Projekt erstellen",
        "edit_project": "Projekt bearbeiten",
        "delete_project": "Projekt löschen",
        "archive_project": "Projekt archivieren",
        "project_brief": "Projektbrief",
        "ai_interview": "KI-Projekt-Onboarding-Interview",
        "next_question": "Nächste Frage",
        "skip_question": "Überspringen",
        "complete_interview": "Interview abschließen",
        "project_brief_confirmation": "Bitte überprüfen Sie den Projektbrief unten und bestätigen Sie das Speichern:",
        "confirm_save": "Bestätigen & Speichern",
        "save_changes": "Änderungen speichern",
        "cancel_edit": "Bearbeitung abbrechen",
        "editing_project": "Projekt bearbeiten",
        # Weekly Update
        "weekly_update_title": "Wöchentliche Aktualisierung",
        "select_project": "Projekt auswählen",
        "project_summary": "Projektzusammenfassung",
        "previous_update": "Vorherige Aktualisierung",
        "ai_interview_start": "KI-Interview - Beantworten Sie die Fragen, um die Aktualisierung dieser Woche zu generieren",
        "your_answer": "Ihre Antwort",
        "submit_answer": "Antwort absenden",
        "interview_complete": "Interview abgeschlossen",
        "review_and_edit": "Überprüfen und bearbeiten Sie die extrahierten Informationen unten:",
        "save_update": "Aktualisierung speichern",
        "update_saved": "Wöchentliche Aktualisierung erfolgreich gespeichert!",
        "duplicate_week_warning": "Für diese Woche existiert bereits eine Aktualisierung. Klicken Sie unten, um sie stattdessen zu bearbeiten.",
        "edit_existing": "Vorhandene Aktualisierung bearbeiten",
        "cancel_week_edit": "Wochenbearbeitung abbrechen",
        "editing_existing_week": "Sie bearbeiten die vorhandene Aktualisierung dieser Woche.",
        # Slide Generator
        "slide_generator_title": "Foliengenerator",
        "select_week": "Woche auswählen",
        "generate_slide": "Folie erstellen",
        "slide_generated": "Folie erfolgreich erstellt!",
        "download_slide": "Folie herunterladen",
        # Gantt Chart
        "gantt_chart_title": "Gantt-Diagramm-Generator",
        "generate_gantt": "Gantt-Diagramm erstellen",
        "gantt_generated": "Gantt-Diagramm erfolgreich erstellt!",
        "download_gantt": "Gantt-Diagramm herunterladen",
        "gantt_tasks": "Gantt-Diagramm-Aufgaben",
        "add_task": "Aufgabe hinzufügen",
        "task_name": "Aufgabennamen",
        "start_date": "Startdatum",
        "end_date": "Enddatum",
        "milestone_date": "Meilenstein-Datum",
        "task_type": "Typ",
        "regular_task": "Regelmäßige Aufgabe",
        "milestone": "Meilenstein",
        "auto_generate_gantt": "Automatisch aus neuester Aktualisierung generieren",
        "gantt_no_tasks": "Keine Gantt-Diagramm-Aufgaben definiert. Automatisch aus der neuesten wöchentlichen Aktualisierung generieren oder Aufgaben manuell hinzufügen.",
        "edit_gantt_task": "Aufgabe bearbeiten",
        "delete_gantt_task": "Aufgabe löschen",
        # Final Documentation
        "final_docs_title": "Abschlussdokumentation",
        "generate_documentation": "Dokumentation erstellen",
        "doc_generated": "Dokumentation erfolgreich erstellt!",
        "download_documentation": "Dokumentation herunterladen",
        # Settings
        "settings_title": "Einstellungen",
        "api_key": "DeepSeek-API-Schlüssel",
        "api_key_placeholder": "Geben Sie Ihren DeepSeek-API-Schlüssel ein",
        "test_connection": "Verbindung testen",
        "connection_success": "✅ Verbindung erfolgreich!",
        "connection_failed": "❌ Verbindung fehlgeschlagen. Überprüfen Sie Ihren API-Schlüssel.",
        "your_name": "Dein Name",
        "language": "Sprache",
        "export_folder": "Exportordner",
        "reset_data": "Alle Daten zurücksetzen",
        "reset_confirm": "Sind Sie sicher? Dies löscht alle Projekte und Aktualisierungen.",
        "setting_saved": "Einstellung gespeichert!",
        # Messages
        "missing_api_key": "⚠️ DeepSeek-API-Schlüssel ist nicht festgelegt. Bitte gehen Sie zu Einstellungen und fügen Sie Ihren API-Schlüssel hinzu, um KI-Funktionen zu nutzen.",
        "api_key_required": "API-Schlüssel erforderlich",
        "go_to_settings": "Zu den Einstellungen",
        "no_data": "Keine Daten verfügbar",
        "no_projects_hint": "Erstellen Sie zuerst ein Projekt, um fortzufahren.",
        "no_updates_hint": "Erstellen Sie zuerst mindestens eine wöchentliche Aktualisierung.",
        "missing_node": "Node.js wird für Folien- und Gantt-Erstellung benötigt.",
        "missing_pptxgen": "Die Abhängigkeit pptxgenjs fehlt. Bitte npm install ausführen.",
        "loading": "Lädt...",
        "error": "Fehler",
        "success": "Erfolg",
        "rewrite_with_ai": "Mit KI umschreiben",
        "ai_suggestions_ready": "KI-Vorschläge sind bereit. Bitte prüfen und bei Bedarf anwenden.",
        "apply_ai_suggestions": "KI-Vorschläge anwenden",
        "discard_ai_suggestions": "Vorschläge verwerfen",
        "rewrite_failed": "KI-Umschreiben fehlgeschlagen. Sie können mit Ihrem aktuellen Text fortfahren.",
        # Data fields (for forms)
        "tasks_completed": "Abgeschlossene Aufgaben",
        "task": "Aufgabe",
        "result": "Ergebnis",
        "owner": "Verantwortlich",
        "next_steps": "Nächste Schritte",
        "management_decisions": "Managemententscheidungen",
        "decision": "Entscheidung",
        "urgency": "Dringlichkeit",
        "context": "Kontext",
        "risks_blockers": "Risiken & Blockierungen",
        "issue": "Problem",
        "impact": "Auswirkung",
        "mitigation": "Minderung",
        "budget_status": "Budgetstatus",
        "budget_notes": "Bemerkungen zum Budget",
        "on_track": "Im Plan",
        "over_budget": "Über Budget",
        "under_budget": "Unter Budget",
        "n_a": "Nicht zutreffend",
        "kpi_updates": "KPI-Aktualisierungen",
        "metric": "Metrik",
        "value": "Wert",
        "trend": "Trend",
        "stakeholder_notes": "Stakeholder-Kommunikation",
        "weekly_summary": "Wöchentliche Zusammenfassung",
        "executive_summary": "Zusammenfassung für die Geschäftsführung",
        # RAG Status
        "rag_green": "Im Plan",
        "rag_amber": "Gefährdet",
        "rag_red": "Blockiert",
    },
}


def get_string(key: str, language: str = "en") -> str:
    """Get a translated string."""
    if language not in STRINGS:
        language = "en"
    return STRINGS[language].get(key, key)


def format_date(date_str: str, language: str = "en") -> str:
    """Format an ISO date string to display format."""
    if not date_str:
        return ""

    try:
        date_obj = date_parser.isoparse(date_str)
    except (ValueError, TypeError):
        return date_str

    if language == "de":
        return date_obj.strftime("%d. %B %Y")  # "3. April 2025"
    else:
        return date_obj.strftime("%B %d, %Y")  # "April 3, 2025"


def format_datetime(dt_str: str, language: str = "en") -> str:
    """Format an ISO datetime string to display format."""
    if not dt_str:
        return ""

    try:
        dt_obj = date_parser.isoparse(dt_str)
    except (ValueError, TypeError):
        return dt_str

    if language == "de":
        return dt_obj.strftime("%d. %B %Y, %H:%M")
    else:
        return dt_obj.strftime("%B %d, %Y, %I:%M %p")


def get_week_number() -> int:
    """Get current ISO week number."""
    return datetime.now().isocalendar()[1]


def get_week_label(week_number: int, language: str = "en", year: int = None) -> str:
    """Generate a week label."""
    if year is None:
        year = datetime.now().year

    if language == "de":
        return f"KW {week_number}, {year}"
    else:
        return f"Week {week_number}, {year}"


def get_rag_color(status: str) -> str:
    """Get the color for a RAG status."""
    status_lower = status.lower()
    if status_lower == "red":
        return COLORS["danger_red"]
    elif status_lower == "amber":
        return COLORS["warning_amber"]
    else:  # green
        return COLORS["success_green"]


def get_rag_emoji(status: str) -> str:
    """Get the emoji for a RAG status."""
    status_lower = status.lower()
    if status_lower == "red":
        return "🔴"
    elif status_lower == "amber":
        return "🟡"
    else:  # green
        return "🟢"


def safe_json_parse(json_input, default=None):
    """Safely parse JSON input and pass through already-parsed values."""
    if default is None:
        default = []

    if json_input is None:
        return default

    if isinstance(json_input, (list, dict)):
        return json_input

    if isinstance(json_input, str):
        stripped = json_input.strip()
        if not stripped:
            return default
        try:
            return json.loads(stripped)
        except (json.JSONDecodeError, TypeError):
            return default

    return default


def get_initials(name: str) -> str:
    """Get initials from a name."""
    parts = name.strip().split()
    if len(parts) >= 2:
        return (parts[0][0] + parts[-1][0]).upper()
    elif parts:
        return parts[0][:2].upper()
    return "?"


def apply_custom_css(language: str = "en"):
    """Apply custom Streamlit CSS styling."""
    css = f"""
    <style>
        /* Global styling */
        :root {{
            --primary-blue: {COLORS["primary_blue"]};
            --dark-blue: {COLORS["dark_blue"]};
            --light-blue-accent: {COLORS["light_blue_accent"]};
            --text-dark: {COLORS["text_dark"]};
            --success-green: {COLORS["success_green"]};
            --warning-amber: {COLORS["warning_amber"]};
            --danger-red: {COLORS["danger_red"]};
        }}

        /* Streamlit main container */
        .main {{
            background-color: {COLORS["white"]};
        }}

        /* Sidebar */
        [data-testid="stSidebar"] {{
            background-color: {COLORS["light_blue_accent"]};
            border-right: 2px solid {COLORS["primary_blue"]};
        }}

        /* Buttons */
        .stButton > button {{
            background-color: {COLORS["primary_blue"]};
            color: {COLORS["white"]};
            border: none;
            border-radius: 4px;
            font-weight: 600;
        }}

        .stButton > button:hover {{
            background-color: {COLORS["dark_blue"]};
        }}

        /* Metrics */
        .stMetric {{
            background-color: {COLORS["light_blue_accent"]};
            padding: 16px;
            border-radius: 8px;
            border-left: 4px solid {COLORS["primary_blue"]};
        }}

        /* Cards */
        .card {{
            background-color: {COLORS["white"]};
            border: 1px solid {COLORS["light_blue_accent"]};
            border-radius: 8px;
            padding: 16px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.05);
        }}

        /* Headers */
        h1, h2, h3 {{
            color: {COLORS["dark_blue"]};
        }}

        /* Text */
        body {{
            color: {COLORS["text_dark"]};
            font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif;
        }}

        /* Forms */
        .stTextInput > div > div > input,
        .stTextArea > div > div > textarea {{
            border: 1px solid {COLORS["primary_blue"]};
            border-radius: 4px;
        }}

        .stTextInput > div > div > input:focus,
        .stTextArea > div > div > textarea:focus {{
            border: 2px solid {COLORS["dark_blue"]};
            box-shadow: 0 0 0 0.2rem {COLORS["light_blue_accent"]};
        }}

        /* Select boxes */
        .stSelectbox > div > div > div {{
            border: 1px solid {COLORS["primary_blue"]};
            border-radius: 4px;
        }}

        /* Expanders */
        .streamlit-expanderHeader {{
            background-color: {COLORS["light_blue_accent"]};
            border-left: 4px solid {COLORS["primary_blue"]};
        }}

        /* RAG Status indicators */
        .rag-green {{
            color: {COLORS["success_green"]};
            font-weight: 600;
        }}

        .rag-amber {{
            color: {COLORS["warning_amber"]};
            font-weight: 600;
        }}

        .rag-red {{
            color: {COLORS["danger_red"]};
            font-weight: 600;
        }}

        /* Status pills */
        .status-pill {{
            display: inline-block;
            padding: 6px 12px;
            border-radius: 20px;
            font-size: 12px;
            font-weight: 600;
            text-align: center;
        }}

        .status-pill.green {{
            background-color: {COLORS["success_green"]};
            color: {COLORS["white"]};
        }}

        .status-pill.amber {{
            background-color: {COLORS["warning_amber"]};
            color: {COLORS["text_dark"]};
        }}

        .status-pill.red {{
            background-color: {COLORS["danger_red"]};
            color: {COLORS["white"]};
        }}
    </style>
    """
    return css
