"""
Project Update Studio - Main Streamlit Application
A tool for project managers to generate weekly status updates and project documentation.
"""

import streamlit as st
import database as db
import utils
from ai_assistant import AIAssistant
from doc_generator import generate_project_documentation
from gantt_generator import generate_gantt_chart_from_tasks, auto_generate_gantt_from_tasks
from datetime import datetime, timedelta
import json
import subprocess
import os
import shutil
from pathlib import Path
import tempfile

# ============= Page Configuration =============
st.set_page_config(
    page_title="Project Update Studio",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# Initialize database
db.init_db()

# ============= Session State Initialization =============
def init_session_state():
    """Initialize session state variables."""
    if "language" not in st.session_state:
        st.session_state.language = db.get_setting("language", "en")
    if "current_page" not in st.session_state:
        st.session_state.current_page = "Dashboard"
    if "nav_page" not in st.session_state:
        st.session_state.nav_page = st.session_state.current_page
    if "pending_nav_page" not in st.session_state:
        st.session_state.pending_nav_page = None
    if "interview_history" not in st.session_state:
        st.session_state.interview_history = []
    if "interview_answers" not in st.session_state:
        st.session_state.interview_answers = {}
    if "interview_q_idx" not in st.session_state:
        st.session_state.interview_q_idx = 0
    if "current_question_idx" not in st.session_state:
        st.session_state.current_question_idx = 0
    if "onboarding_answers" not in st.session_state:
        st.session_state.onboarding_answers = {}
    if "onboarding_q_idx" not in st.session_state:
        st.session_state.onboarding_q_idx = 0
    if "extracted_update" not in st.session_state:
        st.session_state.extracted_update = None
    if "generated_brief" not in st.session_state:
        st.session_state.generated_brief = None
    if "project_created_name" not in st.session_state:
        st.session_state.project_created_name = None
    if "selected_project_id" not in st.session_state:
        st.session_state.selected_project_id = None


init_session_state()

# Apply custom CSS
st.markdown(utils.apply_custom_css(st.session_state.language), unsafe_allow_html=True)

# ============= Utility Functions =============

def get_string(key: str) -> str:
    """Get translated string."""
    return utils.get_string(key, st.session_state.language)


def check_api_key() -> bool:
    """Check if DeepSeek API key is set."""
    api_key = db.get_setting("deepseek_api_key")
    return bool(api_key)


def get_api_key() -> str:
    """Get the stored API key."""
    return db.get_setting("deepseek_api_key", "")


def show_api_key_warning():
    """Show warning if API key is missing."""
    if not check_api_key():
        st.warning(get_string("missing_api_key"))
        col1, col2 = st.columns([1, 1])
        with col2:
            if st.button(f"→ {get_string('go_to_settings')}"):
                navigate_to("Settings")
                st.rerun()


def navigate_to(page: str, project_id: int = None):
    """Navigate to a page and optionally set current project context."""
    st.session_state.current_page = page
    st.session_state.pending_nav_page = page
    if project_id is not None:
        st.session_state.selected_project_id = project_id


def get_project_select_index(projects: list) -> int:
    """Get selectbox index from stored selected project context."""
    if not projects:
        return 0
    project_ids = [p["id"] for p in projects]
    selected_project_id = st.session_state.get("selected_project_id")
    if selected_project_id in project_ids:
        return project_ids.index(selected_project_id)
    return 0


def check_node_dependencies() -> tuple[bool, str]:
    """Check Node.js and pptxgenjs dependency for slide/gantt generation."""
    if shutil.which("node") is None:
        return False, get_string("missing_node")

    project_root = os.path.dirname(os.path.abspath(__file__))
    pptxgen_path = os.path.join(project_root, "node_modules", "pptxgenjs")
    if not os.path.exists(pptxgen_path):
        return False, get_string("missing_pptxgen")

    return True, ""


def _normalize_update_for_document_ai(update: dict) -> dict:
    """Normalize a weekly update into AI-ready JSON-safe fields."""
    return {
        "week_number": update.get("week_number"),
        "week_label": update.get("week_label", ""),
        "created_at": update.get("created_at", ""),
        "rag_status": update.get("rag_status", "green"),
        "budget_status": update.get("budget_status", "not_applicable"),
        "budget_notes": update.get("budget_notes", ""),
        "ai_summary": update.get("ai_summary", ""),
        "tasks_completed": utils.safe_json_parse(update.get("tasks_completed", []), []),
        "next_tasks": utils.safe_json_parse(update.get("next_tasks", []), []),
        "management_decisions": utils.safe_json_parse(update.get("management_decisions", []), []),
        "risks_blockers": utils.safe_json_parse(update.get("risks_blockers", []), []),
        "milestone_hit": utils.safe_json_parse(update.get("milestone_hit", []), []),
        "kpi_updates": utils.safe_json_parse(update.get("kpi_updates", []), []),
    }


def _build_document_ai_payload(project: dict, updates: list) -> dict:
    """Build payload that will be rewritten by AI before doc export."""
    return {
        "project": {
            "name": project.get("name", ""),
            "description": project.get("description", ""),
            "goal": project.get("goal", ""),
            "background": project.get("background", ""),
        },
        "updates": [_normalize_update_for_document_ai(update) for update in updates],
    }


def _merge_ai_polished_document_content(project: dict, updates: list, polished_payload: dict):
    """Merge AI-polished text back into project and updates for document generation."""
    if not isinstance(polished_payload, dict):
        raise Exception("AI returned invalid payload for document polishing.")

    polished_project = polished_payload.get("project", {})
    polished_updates = polished_payload.get("updates", [])

    if not isinstance(polished_project, dict):
        raise Exception("AI returned invalid project section for document polishing.")
    if not isinstance(polished_updates, list):
        raise Exception("AI returned invalid updates section for document polishing.")
    if len(polished_updates) != len(updates):
        raise Exception("AI returned incomplete updates for document polishing.")

    merged_project = dict(project)
    for field in ["description", "goal", "background"]:
        value = polished_project.get(field)
        if isinstance(value, str) and value.strip():
            merged_project[field] = value.strip()

    merged_updates = []
    list_fields = [
        "tasks_completed",
        "next_tasks",
        "management_decisions",
        "risks_blockers",
        "milestone_hit",
        "kpi_updates",
    ]
    text_fields = ["ai_summary", "budget_notes"]

    for idx, update in enumerate(updates):
        merged_update = dict(update)
        polished_update = polished_updates[idx] if isinstance(polished_updates[idx], dict) else {}

        for field in list_fields:
            if field in polished_update:
                merged_update[field] = polished_update[field]

        for field in text_fields:
            value = polished_update.get(field)
            if isinstance(value, str):
                merged_update[field] = value

        merged_updates.append(merged_update)

    return merged_project, merged_updates


# ============= PAGE: DASHBOARD =============

def page_dashboard():
    """Dashboard page - overview of all projects."""
    st.title(f"🏠 {get_string('dashboard_title')}")

    # Get all projects and stats
    projects = db.get_all_projects()
    dashboard_stats = db.get_dashboard_stats()

    # Summary bar
    col1, col2, col3, col4, col5 = st.columns(5)
    with col1:
        st.metric(get_string("total_projects"), dashboard_stats["total_active"])
    with col2:
        st.metric("🟢 " + get_string("on_track"), dashboard_stats["green"])
    with col3:
        st.metric("🟡 " + get_string("at_risk"), dashboard_stats["amber"])
    with col4:
        st.metric("🔴 " + get_string("blocked"), dashboard_stats["red"])
    with col5:
        new_proj = st.button(f"✨ {get_string('new_project')}")
        if new_proj:
            navigate_to("Projects")
            st.rerun()

    st.divider()

    # Project cards
    if projects:
        for i in range(0, len(projects), 2):
            col1, col2 = st.columns(2)

            for col, project in [(col1, projects[i]), (col2, projects[i + 1] if i + 1 < len(projects) else None)]:
                if project:
                    with col:
                        with st.container(border=True):
                            st.subheader(project["name"])

                            # RAG status
                            rag_status = project.get("rag_status", "green")
                            rag_emoji = utils.get_rag_emoji(rag_status)
                            rag_label = get_string(f"rag_{rag_status}")
                            st.markdown(f"**{rag_emoji} {rag_label}**")

                            # Stats
                            stats = db.get_project_stats(project["id"])
                            col_a, col_b, col_c = st.columns(3)
                            with col_a:
                                st.caption(f"📝 {stats['weeks_count']} {get_string('weeks_of_updates')}")
                            with col_b:
                                st.caption(f"👥 {stats['team_count']} {get_string('team_members')}")
                            with col_c:
                                if stats["latest_update"]:
                                    last_date = utils.format_date(stats["latest_update"], st.session_state.language)
                                    st.caption(f"⏱️ {last_date}")

                            # Action buttons
                            action_col1, action_col2 = st.columns(2)
                            with action_col1:
                                if st.button(f"📝 {get_string('new_update')}", key=f"new_update_{project['id']}"):
                                    navigate_to("Weekly Update", project["id"])
                                    st.rerun()
                            with action_col2:
                                if st.button(f"📊 {get_string('generate_slide')}", key=f"gen_slide_{project['id']}"):
                                    navigate_to("Slide Generator", project["id"])
                                    st.rerun()
    else:
        st.info(get_string("no_data"))


# ============= PAGE: PROJECTS =============

def page_projects():
    """Projects page - create and manage projects."""
    st.title(f"📁 {get_string('projects_title')}")

    tab1, tab2 = st.tabs([get_string("new_project"), get_string("projects_title")])

    with tab1:
        st.subheader(get_string("new_project"))

        col1, col2 = st.columns(2)
        with col1:
            project_name = st.text_input(
                get_string("project_name"),
                placeholder="e.g., Process Optimization Q2",
                key="project_name_create",
            )
        with col2:
            project_lang = st.selectbox(get_string("select_language"), ["English", "Deutsch"])
            project_lang_code = "en" if project_lang == "English" else "de"

        project_desc = st.text_area(get_string("project_description"), height=80, key="project_desc_create")
        project_goal = st.text_input(get_string("project_goal"), key="project_goal_create")
        project_bg = st.text_area(get_string("project_background"), height=60, key="project_bg_create")

        rewrite_col1, rewrite_col2 = st.columns([2, 1])
        with rewrite_col2:
            if st.button(get_string("rewrite_with_ai"), key="rewrite_project_create", use_container_width=True):
                if not check_api_key():
                    show_api_key_warning()
                else:
                    try:
                        ai = AIAssistant(get_api_key(), project_lang_code)
                        st.session_state.project_create_suggestions = ai.rewrite_text_fields(
                            {
                                "project_name": project_name,
                                "project_description": project_desc,
                                "project_goal": project_goal,
                                "project_background": project_bg,
                            },
                            context="Project creation form. Rewrite to professional, short, and precise style.",
                        )
                    except Exception:
                        st.warning(get_string("rewrite_failed"))

        if st.session_state.get("project_create_suggestions"):
            st.info(get_string("ai_suggestions_ready"))
            apply_col1, apply_col2 = st.columns(2)
            with apply_col1:
                if st.button(get_string("apply_ai_suggestions"), key="apply_project_create_suggestions", use_container_width=True):
                    suggestions = st.session_state.get("project_create_suggestions", {})
                    st.session_state.project_name_create = suggestions.get("project_name", st.session_state.project_name_create)
                    st.session_state.project_desc_create = suggestions.get("project_description", st.session_state.project_desc_create)
                    st.session_state.project_goal_create = suggestions.get("project_goal", st.session_state.project_goal_create)
                    st.session_state.project_bg_create = suggestions.get("project_background", st.session_state.project_bg_create)
                    st.session_state.project_create_suggestions = None
                    st.rerun()
            with apply_col2:
                if st.button(get_string("discard_ai_suggestions"), key="discard_project_create_suggestions", use_container_width=True):
                    st.session_state.project_create_suggestions = None
                    st.rerun()

        # Team members
        st.subheader(get_string("team"))
        team_members = []
        num_members = st.number_input(
            f"{get_string('team_members')} {get_string('n_a')}",
            min_value=0,
            max_value=20,
            value=1,
            key="project_create_member_count",
        )

        for i in range(num_members):
            cols = st.columns([2, 2])
            with cols[0]:
                name = st.text_input(get_string("team_member_name"), key=f"member_name_{i}")
            with cols[1]:
                role = st.text_input(get_string("role"), key=f"member_role_{i}")
            if name:
                team_members.append({"name": name, "role": role})

        if st.button(f"✅ {get_string('create_project')}", use_container_width=True):
            if project_name:
                project_id = db.create_project(
                    name=project_name,
                    language=project_lang_code,
                    description=project_desc,
                    goal=project_goal,
                    background=project_bg,
                    team_members=team_members,
                )
                # Save project name to session for confirmation message
                st.session_state.project_created_name = project_name

                # Start AI onboarding interview
                st.session_state.current_page = "Projects"
                st.session_state.onboarding_project_id = project_id
                st.session_state.onboarding_language = project_lang_code
                st.session_state.project_create_suggestions = None
                st.rerun()
            else:
                st.error(f"❌ {get_string('project_name')} {get_string('error')}")

    with tab2:
        st.subheader(get_string("projects_title"))
        projects = db.get_all_projects()

        if projects:
            for project in projects:
                with st.expander(f"📁 {project['name']} ({project.get('status', 'active')})"):
                    col1, col2, col3 = st.columns(3)

                    stats = db.get_project_stats(project["id"])
                    with col1:
                        st.metric(get_string("weeks_of_updates"), stats["weeks_count"])
                    with col2:
                        st.metric(get_string("team_members"), stats["team_count"])
                    with col3:
                        rag = project.get("rag_status", "green")
                        st.markdown(f"**{utils.get_rag_emoji(rag)} {get_string(f'rag_{rag}')}**")

                    # Details
                    st.markdown(f"**{get_string('project_description')}:** {project.get('description', '-')}")
                    st.markdown(f"**{get_string('project_goal')}:** {project.get('goal', '-')}")

                    # Team
                    team = db.get_project_team(project["id"])
                    if team:
                        st.markdown(f"**{get_string('team')}:**")
                        for member in team:
                            role_text = f" ({member['role']})" if member.get("role") else ""
                            st.caption(f"• {member['name']}{role_text}")

                    # Actions
                    action_col1, action_col2, action_col3 = st.columns(3)
                    with action_col1:
                        if st.button(f"📝 {get_string('edit_project')}", key=f"edit_{project['id']}"):
                            st.session_state.editing_project_id = project["id"]
                            st.rerun()
                    with action_col2:
                        if st.button(f"🔖 {get_string('archive_project')}", key=f"archive_{project['id']}"):
                            db.update_project(project["id"], status="archived")
                            st.success(f"{get_string('success')}")
                            st.rerun()
                    with action_col3:
                        if st.button(f"🗑️ {get_string('delete_project')}", key=f"delete_{project['id']}"):
                            db.delete_project(project["id"])
                            st.success(f"{get_string('success')}")
                            st.rerun()
        else:
            st.info(get_string("no_data"))

    editing_project_id = st.session_state.get("editing_project_id")
    if editing_project_id:
        project_to_edit = db.get_project(editing_project_id)
        if not project_to_edit:
            st.warning(get_string("error"))
            st.session_state.editing_project_id = None
            st.rerun()

        edit_team = db.get_project_team(editing_project_id)
        if st.session_state.get("editing_project_loaded_id") != editing_project_id:
            st.session_state.edit_project_name = project_to_edit.get("name", "")
            st.session_state.edit_project_desc = project_to_edit.get("description", "")
            st.session_state.edit_project_goal = project_to_edit.get("goal", "")
            st.session_state.edit_project_bg = project_to_edit.get("background", "")
            st.session_state.edit_project_lang = "English" if project_to_edit.get("language", "en") == "en" else "Deutsch"
            st.session_state.edit_project_member_count = max(1, len(edit_team))
            for idx, member in enumerate(edit_team):
                st.session_state[f"edit_member_name_{idx}"] = member.get("name", "")
                st.session_state[f"edit_member_role_{idx}"] = member.get("role", "")
            st.session_state.editing_project_loaded_id = editing_project_id

        st.divider()
        st.subheader(f"✏️ {get_string('editing_project')}")

        col1, col2 = st.columns(2)
        with col1:
            edit_name = st.text_input(get_string("project_name"), key="edit_project_name")
        with col2:
            edit_lang_name = st.selectbox(
                get_string("select_language"),
                ["English", "Deutsch"],
                key="edit_project_lang",
            )
        edit_lang_code = "en" if edit_lang_name == "English" else "de"

        edit_desc = st.text_area(get_string("project_description"), height=80, key="edit_project_desc")
        edit_goal = st.text_input(get_string("project_goal"), key="edit_project_goal")
        edit_bg = st.text_area(get_string("project_background"), height=60, key="edit_project_bg")

        rewrite_col1, rewrite_col2 = st.columns([2, 1])
        with rewrite_col2:
            if st.button(get_string("rewrite_with_ai"), key="rewrite_project_edit", use_container_width=True):
                if not check_api_key():
                    show_api_key_warning()
                else:
                    try:
                        ai = AIAssistant(get_api_key(), edit_lang_code)
                        st.session_state.project_edit_suggestions = ai.rewrite_text_fields(
                            {
                                "project_name": edit_name,
                                "project_description": edit_desc,
                                "project_goal": edit_goal,
                                "project_background": edit_bg,
                            },
                            context="Project edit form. Rewrite to professional, short, and precise style.",
                        )
                    except Exception:
                        st.warning(get_string("rewrite_failed"))

        if st.session_state.get("project_edit_suggestions"):
            st.info(get_string("ai_suggestions_ready"))
            apply_col1, apply_col2 = st.columns(2)
            with apply_col1:
                if st.button(get_string("apply_ai_suggestions"), key="apply_project_edit_suggestions", use_container_width=True):
                    suggestions = st.session_state.get("project_edit_suggestions", {})
                    st.session_state.edit_project_name = suggestions.get("project_name", st.session_state.edit_project_name)
                    st.session_state.edit_project_desc = suggestions.get("project_description", st.session_state.edit_project_desc)
                    st.session_state.edit_project_goal = suggestions.get("project_goal", st.session_state.edit_project_goal)
                    st.session_state.edit_project_bg = suggestions.get("project_background", st.session_state.edit_project_bg)
                    st.session_state.project_edit_suggestions = None
                    st.rerun()
            with apply_col2:
                if st.button(get_string("discard_ai_suggestions"), key="discard_project_edit_suggestions", use_container_width=True):
                    st.session_state.project_edit_suggestions = None
                    st.rerun()

        st.markdown(f"**{get_string('team')}**")
        edit_num_members = st.number_input(
            f"{get_string('team_members')} {get_string('n_a')}",
            min_value=0,
            max_value=20,
            key="edit_project_member_count",
        )

        edited_team_members = []
        for i in range(edit_num_members):
            cols = st.columns([2, 2])
            with cols[0]:
                member_name = st.text_input(get_string("team_member_name"), key=f"edit_member_name_{i}")
            with cols[1]:
                member_role = st.text_input(get_string("role"), key=f"edit_member_role_{i}")
            if member_name:
                edited_team_members.append({"name": member_name, "role": member_role})

        action_col1, action_col2 = st.columns(2)
        with action_col1:
            if st.button(f"✅ {get_string('save_changes')}", key="save_project_edit", use_container_width=True):
                if not edit_name.strip():
                    st.error(f"❌ {get_string('project_name')} {get_string('error')}")
                else:
                    db.update_project(
                        editing_project_id,
                        name=edit_name.strip(),
                        language=edit_lang_code,
                        description=edit_desc,
                        goal=edit_goal,
                        background=edit_bg,
                    )
                    db.replace_project_team(editing_project_id, edited_team_members)
                    st.session_state.editing_project_id = None
                    st.session_state.editing_project_loaded_id = None
                    st.session_state.project_edit_suggestions = None
                    st.success(get_string("success"))
                    st.rerun()

        with action_col2:
            if st.button(f"❌ {get_string('cancel_edit')}", key="cancel_project_edit", use_container_width=True):
                st.session_state.editing_project_id = None
                st.session_state.editing_project_loaded_id = None
                st.session_state.project_edit_suggestions = None
                st.rerun()

    # AI Onboarding Interview (if active)
    if "onboarding_project_id" in st.session_state:
        st.divider()
        # Show confirmation of project creation
        if st.session_state.project_created_name:
            st.success(f"✅ {get_string('success')}: {st.session_state.project_created_name}")
            st.session_state.project_created_name = None
        
        st.subheader(get_string("ai_interview"))

        project_id = st.session_state.onboarding_project_id
        language = st.session_state.onboarding_language

        if not check_api_key():
            show_api_key_warning()
        else:
            run_onboarding_interview(project_id, language)


def run_onboarding_interview(project_id: int, language: str):
    """Run the AI project onboarding interview."""
    api_key = get_api_key()
    ai = AIAssistant(api_key, language)

    questions = ai.generate_project_onboarding_questions()

    q_idx = st.session_state.onboarding_q_idx

    if q_idx < len(questions):
        st.info(f"Question {q_idx + 1}/{len(questions)}")
        st.progress((q_idx + 1) / len(questions))
        st.write(questions[q_idx])

        answer = st.text_area(get_string("your_answer"), height=100, key=f"q_{q_idx}")

        col1, col2 = st.columns(2)
        with col1:
            if st.button(get_string("submit_answer"), key=f"onboarding_submit_{q_idx}"):
                st.session_state.onboarding_answers[q_idx] = answer
                st.session_state.onboarding_q_idx += 1
                st.rerun()
        with col2:
            if st.button(get_string("skip_question"), key=f"onboarding_skip_{q_idx}"):
                st.session_state.onboarding_q_idx += 1
                st.rerun()

    else:
        st.success(get_string("interview_complete"))

        # Generate project brief using AI
        with st.spinner(get_string("loading")):
            try:
                # Build conversation summary
                conversation = "\n".join(
                    [f"Q: {questions[i]}\nA: {st.session_state.onboarding_answers.get(i, '')}"
                     for i in range(len(questions))]
                )

                # Generate brief
                system_prompt = ai._build_system_prompt(context="You are generating a project brief from interview answers. Create a concise, structured brief.")
                project_brief = ai.ask_question(
                    f"Based on the following interview answers, generate a professional project brief (3-5 sentences) summarizing the project goals, scope, and success criteria.\n\n{conversation}",
                    system_prompt,
                )

                st.session_state.generated_brief = project_brief

            except Exception as e:
                st.error(f"❌ {e}")

        # Show brief for confirmation
        if "generated_brief" in st.session_state:
            st.markdown(f"**{get_string('project_brief')}:**")
            st.info(st.session_state.generated_brief)

            if st.button(f"✅ {get_string('confirm_save')}", key="onboarding_confirm_save"):
                db.update_project(
                    project_id,
                    background=st.session_state.generated_brief,
                    description="\n".join([st.session_state.onboarding_answers.get(i, "") for i in range(2)]),
                )
                st.session_state.onboarding_project_id = None
                st.session_state.onboarding_language = None
                st.session_state.onboarding_answers = {}
                st.session_state.onboarding_q_idx = 0
                st.session_state.generated_brief = None
                st.success(f"✅ {get_string('update_saved')}")
                st.rerun()


# ============= PAGE: WEEKLY UPDATE =============

def page_weekly_update():
    """Weekly Update page - AI-guided interview."""
    st.title(f"📝 {get_string('weekly_update_title')}")

    if not check_api_key():
        show_api_key_warning()
        return

    # Select project
    projects = db.get_all_projects()
    if not projects:
        st.warning(get_string("no_data"))
        st.info(get_string("no_projects_hint"))
        return

    selected_project_id = st.selectbox(
        get_string("select_project"),
        options=[p["id"] for p in projects],
        index=get_project_select_index(projects),
        format_func=lambda x: next((p["name"] for p in projects if p["id"] == x), "Unknown"),
    )
    st.session_state.selected_project_id = selected_project_id

    if st.session_state.get("weekly_project_context") != selected_project_id:
        st.session_state.weekly_project_context = selected_project_id
        st.session_state.interview_answers = {}
        st.session_state.interview_q_idx = 0
        st.session_state.extracted_update = None
        st.session_state.weekly_update_suggestion = None
        st.session_state.editing_update_id = None

    project = db.get_project(selected_project_id)
    st.subheader(project["name"])

    # Show previous update summary
    updates = db.get_project_updates(selected_project_id)
    if updates:
        st.info(f"**{get_string('previous_update')}:** {updates[0]['ai_summary'][:200]}...")

    # Run weekly update interview
    api_key = get_api_key()
    ai = AIAssistant(api_key, project.get("language", "en"))

    questions = ai.generate_weekly_update_questions()

    # Check for duplicate week
    current_week = utils.get_week_number()
    existing_update = db.get_weekly_update_by_week(selected_project_id, current_week)

    editing_update_id = st.session_state.get("editing_update_id")
    is_editing_existing_week = bool(existing_update and editing_update_id == existing_update.get("id"))

    if existing_update and not is_editing_existing_week:
        st.warning(get_string("duplicate_week_warning"))
        if st.button(get_string("edit_existing")):
            st.session_state.editing_update_id = existing_update["id"]
            st.session_state.interview_q_idx = len(questions)
            st.session_state.interview_answers = {}
            st.session_state.extracted_update = db.get_weekly_update(existing_update["id"])
            st.rerun()
        st.stop()

    if is_editing_existing_week:
        st.info(get_string("editing_existing_week"))
        if st.button(get_string("cancel_week_edit"), key="cancel_week_edit"):
            st.session_state.editing_update_id = None
            st.session_state.interview_answers = {}
            st.session_state.interview_q_idx = 0
            st.session_state.extracted_update = None
            st.rerun()

    q_idx = st.session_state.interview_q_idx

    if q_idx < len(questions):
        st.info(f"Question {q_idx + 1}/{len(questions)}")
        st.progress((q_idx + 1) / len(questions))
        st.write(questions[q_idx])

        answer = st.text_area(get_string("your_answer"), height=100, key=f"weekly_q_{q_idx}")

        col1, col2 = st.columns(2)
        with col1:
            if st.button(get_string("submit_answer"), key=f"weekly_submit_{q_idx}"):
                st.session_state.interview_answers[q_idx] = answer
                st.session_state.interview_q_idx += 1
                st.rerun()
        with col2:
            if st.button(get_string("skip_question"), key=f"weekly_skip_{q_idx}"):
                st.session_state.interview_q_idx += 1
                st.rerun()

    else:
        st.success(get_string("interview_complete"))

        if not is_editing_existing_week and st.session_state.extracted_update is None:
            # Extract data using AI
            with st.spinner(get_string("loading")):
                try:
                    conversation = "\n".join(
                        [f"Q: {questions[i]}\nA: {st.session_state.interview_answers.get(i, '')}"
                         for i in range(len(questions))]
                    )

                    # Validate conversation has actual content
                    answered_count = sum(1 for i in range(len(questions)) if st.session_state.interview_answers.get(i, "").strip())
                    if answered_count == 0:
                        st.warning("⚠️ No answers provided. Please answer at least one question before extracting.")
                        st.stop()

                    # Build system prompt with context
                    previous_updates_summary = "\n".join(
                        [f"Week {u['week_number']}: {u['ai_summary']}" for u in updates[:3]]
                    ) if updates else "No previous updates"

                    system_prompt = ai._build_system_prompt(
                        project_brief=project.get("background", ""),
                        previous_updates=previous_updates_summary,
                    )

                    # Extract structured data
                    structured_data = ai.extract_weekly_update_data(conversation, system_prompt)
                    st.session_state.extracted_update = structured_data

                except Exception as e:
                    st.error(f"❌ {get_string('error')}: {e}")
                    # Set default empty structure to allow manual entry
                    st.session_state.extracted_update = {
                        "tasks_completed": [],
                        "next_tasks": [],
                        "rag_status": "green",
                        "rag_reason": "",
                        "management_decisions": [],
                        "risks_blockers": [],
                        "budget_status": "not_applicable",
                        "budget_notes": "",
                        "kpi_updates": [],
                        "milestone_hit": [],
                        "ai_summary": "",
                    }
                    st.warning("✏️ " + get_string("review_and_edit"))

        # Show review and edit interface
        if "extracted_update" in st.session_state and st.session_state.extracted_update is not None:
            st.markdown(f"**{get_string('review_and_edit')}**")

            data = st.session_state.extracted_update

            # Editable fields
            with st.expander(get_string("tasks_completed"), expanded=True):
                tasks = data.get("tasks_completed", [])
                edited_tasks = []
                for i, task in enumerate(tasks):
                    col1, col2, col3 = st.columns(3)
                    with col1:
                        t = st.text_input("Task", value=task.get("task", "") if isinstance(task, dict) else task, key=f"task_{i}")
                    with col2:
                        r = st.text_input("Result", value=task.get("result", "") if isinstance(task, dict) else "", key=f"result_{i}")
                    with col3:
                        o = st.text_input("Owner", value=task.get("owner", "") if isinstance(task, dict) else "", key=f"owner_{i}")
                    edited_tasks.append({"task": t, "result": r, "owner": o})
                data["tasks_completed"] = edited_tasks

            with st.expander(get_string("next_steps")):
                next_tasks = data.get("next_tasks", [])
                edited_next = []
                for i, task in enumerate(next_tasks):
                    col1, col2, col3 = st.columns([2, 1, 1])
                    with col1:
                        t = st.text_input("Task", value=task.get("task", "") if isinstance(task, dict) else task, key=f"next_task_{i}")
                    with col2:
                        o = st.text_input("Owner", value=task.get("owner", "") if isinstance(task, dict) else "", key=f"next_owner_{i}")
                    with col3:
                        d = st.text_input("Due", value=task.get("due_date", "") if isinstance(task, dict) else "", key=f"due_{i}")
                    edited_next.append({"task": t, "owner": o, "due_date": d})
                data["next_tasks"] = edited_next

            col1, col2 = st.columns(2)
            with col1:
                st.markdown(f"**RAG Status:** {data.get('rag_status', 'green').upper()}")
                data["rag_status"] = st.select_slider(
                    "Override RAG",
                    options=["red", "amber", "green"],
                    value=data.get("rag_status", "green"),
                    key="rag_override",
                )

            with col2:
                st.markdown(f"**Budget:** {data.get('budget_status', 'n/a')}")
                budget_options = ["not_applicable", "on_track", "over", "under"]
                budget_value = data.get("budget_status", "not_applicable")
                if budget_value not in budget_options:
                    budget_value = "not_applicable"
                data["budget_status"] = st.selectbox(
                    "Budget Status",
                    options=budget_options,
                    index=budget_options.index(budget_value),
                    key="budget_override",
                )

            data["ai_summary"] = st.text_area("Executive Summary", value=data.get("ai_summary", ""), key="summary_edit")
            data["budget_notes"] = st.text_area("Budget Notes", value=data.get("budget_notes", ""), key="budget_notes_edit")

            if st.button(get_string("rewrite_with_ai"), key="rewrite_weekly_review", use_container_width=True):
                try:
                    polish_system_prompt = ai._build_system_prompt(
                        project_brief=project.get("background", ""),
                        previous_updates="\n".join(
                            [f"Week {u.get('week_number')}: {u.get('ai_summary', '')}" for u in updates[:5]]
                        ),
                        context="Polish one weekly update for concise, professional executive communication.",
                    )
                    suggestion = ai.polish_single_weekly_update(
                        {
                            "ai_summary": data.get("ai_summary", ""),
                            "tasks_completed": data.get("tasks_completed", []),
                            "next_tasks": data.get("next_tasks", []),
                            "management_decisions": data.get("management_decisions", []),
                            "risks_blockers": data.get("risks_blockers", []),
                            "kpi_updates": data.get("kpi_updates", []),
                            "milestone_hit": data.get("milestone_hit", []),
                            "budget_notes": data.get("budget_notes", ""),
                        },
                        polish_system_prompt,
                    )
                    st.session_state.weekly_update_suggestion = suggestion
                except Exception:
                    st.warning(get_string("rewrite_failed"))

            if st.session_state.get("weekly_update_suggestion"):
                st.info(get_string("ai_suggestions_ready"))
                apply_col1, apply_col2 = st.columns(2)
                with apply_col1:
                    if st.button(get_string("apply_ai_suggestions"), key="apply_weekly_suggestions", use_container_width=True):
                        suggestion = st.session_state.get("weekly_update_suggestion", {})
                        for key in [
                            "ai_summary",
                            "tasks_completed",
                            "next_tasks",
                            "management_decisions",
                            "risks_blockers",
                            "kpi_updates",
                            "milestone_hit",
                            "budget_notes",
                        ]:
                            if key in suggestion:
                                data[key] = suggestion[key]
                        st.session_state.extracted_update = data
                        st.session_state.weekly_update_suggestion = None
                        st.rerun()
                with apply_col2:
                    if st.button(get_string("discard_ai_suggestions"), key="discard_weekly_suggestions", use_container_width=True):
                        st.session_state.weekly_update_suggestion = None
                        st.rerun()

            # Save button
            if st.button(f"✅ {get_string('save_update')}", use_container_width=True):
                current_week = utils.get_week_number()
                year = datetime.now().year
                week_label = utils.get_week_label(current_week, st.session_state.language, year)

                if is_editing_existing_week and editing_update_id:
                    db.update_weekly_update(
                        editing_update_id,
                        week_label=week_label,
                        rag_status=data.get("rag_status", "green"),
                        tasks_completed=data.get("tasks_completed", []),
                        next_tasks=data.get("next_tasks", []),
                        management_decisions=data.get("management_decisions", []),
                        risks_blockers=data.get("risks_blockers", []),
                        budget_status=data.get("budget_status", "not_applicable"),
                        budget_notes=data.get("budget_notes", ""),
                        milestone_hit=data.get("milestone_hit", []),
                        kpi_updates=data.get("kpi_updates", []),
                        ai_summary=data.get("ai_summary", ""),
                    )
                else:
                    db.create_weekly_update(
                        project_id=selected_project_id,
                        week_number=current_week,
                        week_label=week_label,
                        rag_status=data.get("rag_status", "green"),
                        tasks_completed=json.dumps(data.get("tasks_completed", [])),
                        next_tasks=json.dumps(data.get("next_tasks", [])),
                        management_decisions=json.dumps(data.get("management_decisions", [])),
                        risks_blockers=json.dumps(data.get("risks_blockers", [])),
                        budget_status=data.get("budget_status", "not_applicable"),
                        budget_notes=data.get("budget_notes", ""),
                        milestone_hit=json.dumps(data.get("milestone_hit", [])),
                        kpi_updates=json.dumps(data.get("kpi_updates", [])),
                        ai_summary=data.get("ai_summary", ""),
                    )

                # Update project RAG status
                db.update_project(selected_project_id, rag_status=data.get("rag_status", "green"))

                st.success(get_string("update_saved"))
                
                # Clear interview state
                st.session_state.interview_answers = {}
                st.session_state.interview_q_idx = 0
                st.session_state.extracted_update = None
                st.session_state.weekly_update_suggestion = None
                st.session_state.editing_update_id = None

                st.rerun()


# ============= PAGE: SLIDE GENERATOR =============

def page_slide_generator():
    """Slide Generator page."""
    st.title(f"📊 {get_string('slide_generator_title')}")

    # Select project and week
    projects = db.get_all_projects()
    if not projects:
        st.warning(get_string("no_data"))
        st.info(get_string("no_projects_hint"))
        return

    selected_project_id = st.selectbox(
        get_string("select_project"),
        options=[p["id"] for p in projects],
        index=get_project_select_index(projects),
        format_func=lambda x: next((p["name"] for p in projects if p["id"] == x), "Unknown"),
    )
    st.session_state.selected_project_id = selected_project_id

    project = db.get_project(selected_project_id)
    updates = db.get_project_updates(selected_project_id)

    if not updates:
        st.warning(f"No updates for {project['name']}")
        st.info(get_string("no_updates_hint"))
        return

    node_ready, node_message = check_node_dependencies()
    if not node_ready:
        st.warning(node_message)

    selected_week_idx = st.selectbox(
        get_string("select_week"),
        options=range(len(updates)),
        format_func=lambda x: f"{updates[x]['week_label']} ({utils.format_date(updates[x]['created_at'], st.session_state.language)})",
    )

    selected_update = updates[selected_week_idx]

    if st.button(f"✨ {get_string('generate_slide')}", use_container_width=True, disabled=not node_ready):
        with st.spinner(get_string("loading")):
            try:
                # Prepare slide data
                team = db.get_project_team(selected_project_id)
                
                # Create absolute path for output file
                output_filename = f"{project['name'].replace(' ', '_')}_W{selected_update['week_number']}.pptx"
                output_path = os.path.join(
                    os.path.dirname(os.path.abspath(__file__)),
                    "exports",
                    output_filename
                )
                
                # Ensure exports directory exists
                exports_dir = os.path.dirname(output_path)
                os.makedirs(exports_dir, exist_ok=True)
                
                slide_data = {
                    "project_name": project["name"],
                    "week_number": selected_update["week_number"],
                    "week_label": selected_update["week_label"],
                    "rag_status": selected_update["rag_status"],
                    "language": project.get("language", "en"),
                    "tasks_completed": selected_update.get("tasks_completed", "[]"),
                    "next_tasks": selected_update.get("next_tasks", "[]"),
                    "management_decisions": selected_update.get("management_decisions", "[]"),
                    "kpi_updates": selected_update.get("kpi_updates", "[]"),
                    "ai_summary": selected_update.get("ai_summary", ""),
                    "team_members": [{"name": m["name"], "role": m.get("role", "")} for m in team],
                    "output_path": output_path,
                }

                # Write temp JSON file
                with tempfile.NamedTemporaryFile(mode='w', suffix='.json', delete=False) as f:
                    json.dump(slide_data, f)
                    temp_json = f.name

                # Run slide generator from project_update_studio directory
                result = subprocess.run(
                    ["node", "slide_generator.js", temp_json],
                    capture_output=True,
                    text=True,
                    cwd=os.path.dirname(os.path.abspath(__file__)),
                )

                os.unlink(temp_json)

                if result.returncode != 0:
                    st.error(f"❌ Slide generation failed with exit code {result.returncode}")
                    st.error("**Errors:**")
                    st.code(result.stderr)
                    st.error("**Output:**")
                    st.code(result.stdout)
                else:
                    # Show script output for debugging
                    if result.stdout:
                        st.info("**Script output:**")
                        st.code(result.stdout)
                    
                    if os.path.exists(output_path):
                        st.success(get_string("slide_generated"))

                        with open(output_path, "rb") as f:
                            st.download_button(
                                label=get_string("download_slide"),
                                data=f.read(),
                                file_name=os.path.basename(output_path),
                                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                                use_container_width=True,
                            )
                    else:
                        st.error(f"❌ Output file not found at: {output_path}")
                        if result.stdout:
                            st.error("**Debug output from script:**")
                            st.code(result.stdout)
                        if result.stderr:
                            st.error("**Errors from script:**")
                            st.code(result.stderr)

            except Exception as e:
                st.error(f"❌ {e}")


# ============= PAGE: FINAL DOCUMENTATION =============

def page_final_documentation():
    """Final Documentation page."""
    st.title(f"📄 {get_string('final_docs_title')}")

    projects = db.get_all_projects()
    if not projects:
        st.warning(get_string("no_data"))
        st.info(get_string("no_projects_hint"))
        return

    selected_project_id = st.selectbox(
        get_string("select_project"),
        options=[p["id"] for p in projects],
        index=get_project_select_index(projects),
        format_func=lambda x: next((p["name"] for p in projects if p["id"] == x), "Unknown"),
    )
    st.session_state.selected_project_id = selected_project_id

    project = db.get_project(selected_project_id)
    team = db.get_project_team(selected_project_id)
    updates = db.get_project_updates(selected_project_id)

    if not updates:
        st.warning(f"No updates for {project['name']}")
        st.info(get_string("no_updates_hint"))
        return

    if not check_api_key():
        st.warning("AI formatting is required before generating final documentation.")
        show_api_key_warning()
        return

    if st.button(f"✨ {get_string('generate_documentation')}", use_container_width=True):
        with st.spinner(get_string("loading")):
            try:
                # Ensure exports directory exists
                exports_dir = os.path.join(
                    os.path.dirname(os.path.abspath(__file__)),
                    "exports"
                )
                os.makedirs(exports_dir, exist_ok=True)

                # Build AI and rewrite all user inputs before writing to Word.
                api_key = get_api_key()
                ai = AIAssistant(api_key, project.get("language", "en"))

                ai_payload = _build_document_ai_payload(project, updates)
                polish_system_prompt = ai._build_system_prompt(
                    project_brief=project.get("background", ""),
                    previous_updates="\n".join([
                        f"Week {u.get('week_number')}: {u.get('ai_summary', '')}" for u in updates[:5]
                    ]),
                    context="You are polishing project documentation content for executive readability.",
                )
                polished_payload = ai.polish_document_inputs(
                    ai_payload["project"],
                    ai_payload["updates"],
                    polish_system_prompt,
                )

                project_for_doc, updates_for_doc = _merge_ai_polished_document_content(
                    project,
                    updates,
                    polished_payload,
                )

                all_updates_summary = "\n".join(
                    [f"{u.get('week_label', 'Week')}: {u.get('ai_summary', '')}" for u in updates_for_doc]
                )
                ai_summary = ai.generate_project_closure_summary(
                    project_for_doc.get("background", ""),
                    all_updates_summary,
                    ai._build_system_prompt(
                        project_for_doc.get("background", ""),
                        all_updates_summary,
                        context="You are writing a final closure summary for a polished project record.",
                    ),
                )

                # Generate document
                filepath = generate_project_documentation(
                    project=project_for_doc,
                    team_members=team,
                    all_updates=updates_for_doc,
                    ai_closure_summary=ai_summary,
                    language=project.get("language", "en"),
                )

                st.success(get_string("doc_generated"))

                with open(filepath, "rb") as f:
                    st.download_button(
                        label=get_string("download_documentation"),
                        data=f.read(),
                        file_name=os.path.basename(filepath),
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        use_container_width=True,
                    )

            except Exception as e:
                st.error(f"❌ {e}")


# ============= PAGE: GANTT CHART =============

def page_gantt_chart():
    """Gantt Chart Generator page."""
    st.title(f"📊 {get_string('gantt_chart_title')}")

    projects = db.get_all_projects()
    if not projects:
        st.warning(get_string("no_data"))
        st.info(get_string("no_projects_hint"))
        return

    selected_project_id = st.selectbox(
        get_string("select_project"),
        options=[p["id"] for p in projects],
        index=get_project_select_index(projects),
        format_func=lambda x: next((p["name"] for p in projects if p["id"] == x), "Unknown"),
    )
    st.session_state.selected_project_id = selected_project_id

    project = db.get_project(selected_project_id)
    team = db.get_project_team(selected_project_id)
    gantt_tasks = db.get_gantt_tasks(selected_project_id)
    project_updates = db.get_project_updates(selected_project_id)

    # Tabs for managing and generating gantt chart
    tab1, tab2 = st.tabs([get_string("gantt_tasks"), get_string("generate_gantt")])

    with tab1:
        st.subheader(get_string("gantt_tasks"))

        # Auto-generate button
        col1, col2 = st.columns([2, 1])
        with col1:
            st.write(get_string("gantt_no_tasks") if not gantt_tasks else f"✅ {len(gantt_tasks)} tasks defined")
            if not project_updates:
                st.caption(get_string("no_updates_hint"))
        with col2:
            if st.button(
                f"🔄 {get_string('auto_generate_gantt')}",
                use_container_width=True,
                disabled=not bool(project_updates),
            ):
                with st.spinner(get_string("loading")):
                    result = auto_generate_gantt_from_tasks(selected_project_id)
                    if result["success"]:
                        st.success(result["message"])
                        st.rerun()
                    else:
                        st.warning(result["message"])

        st.divider()

        # Display existing tasks
        if gantt_tasks:
            st.subheader("📋 Tasks Overview")
            for task in gantt_tasks:
                with st.container(border=True):
                    col1, col2, col3 = st.columns([2, 1, 1])
                    with col1:
                        st.write(f"**{task['name']}**")
                        if task["task_type"] == "milestone":
                            st.caption(f"🎯 Milestone: {task['milestone_date']}")
                        else:
                            st.caption(f"📅 {task['start_date']} → {task['end_date']}")
                            if task.get("team_members"):
                                st.caption(f"👥 {', '.join(task['team_members'])}")
                    with col2:
                        if st.button("✏️", key=f"edit_task_{task['id']}", help="Edit"):
                            st.session_state.editing_task_id = task["id"]
                            st.rerun()
                    with col3:
                        if st.button("🗑️", key=f"delete_task_{task['id']}", help="Delete"):
                            db.delete_gantt_task(task["id"])
                            st.success("Task deleted")
                            st.rerun()

            st.divider()

        # Add/Edit task form
        st.subheader(f"➕ {get_string('add_task')}")

        editing_task_id = st.session_state.get("editing_task_id")
        editing_task = None
        if editing_task_id:
            for task in gantt_tasks:
                if task["id"] == editing_task_id:
                    editing_task = task
                    break

        col1, col2 = st.columns(2)
        with col1:
            task_name = st.text_input(
                get_string("task_name"),
                value=editing_task["name"] if editing_task else "",
                key="task_name_input"
            )
            task_type = st.selectbox(
                get_string("task_type"),
                [get_string("regular_task"), get_string("milestone")],
                index=1 if editing_task and editing_task["task_type"] == "milestone" else 0,
                key="task_type_select"
            )

        with col2:
            if task_type == get_string("regular_task"):
                start_date = st.date_input(
                    get_string("start_date"),
                    value=datetime.strptime(editing_task["start_date"], "%Y-%m-%d").date() if editing_task and editing_task["start_date"] else datetime.now().date(),
                    key="start_date_input"
                )
                end_date = st.date_input(
                    get_string("end_date"),
                    value=datetime.strptime(editing_task["end_date"], "%Y-%m-%d").date() if editing_task and editing_task["end_date"] else (datetime.now() + timedelta(days=14)).date(),
                    key="end_date_input"
                )
            else:
                milestone_date = st.date_input(
                    get_string("milestone_date"),
                    value=datetime.strptime(editing_task["milestone_date"], "%Y-%m-%d").date() if editing_task and editing_task["milestone_date"] else datetime.now().date(),
                    key="milestone_date_input"
                )

        if st.button(get_string("rewrite_with_ai"), key="rewrite_gantt_task", use_container_width=True):
            if not check_api_key():
                show_api_key_warning()
            else:
                try:
                    ai = AIAssistant(get_api_key(), project.get("language", "en"))
                    st.session_state.gantt_task_suggestion = ai.rewrite_text_fields(
                        {"task_name": task_name},
                        context="Gantt chart task naming. Keep concise and professional.",
                    )
                except Exception:
                    st.warning(get_string("rewrite_failed"))

        if st.session_state.get("gantt_task_suggestion"):
            st.info(get_string("ai_suggestions_ready"))
            apply_col1, apply_col2 = st.columns(2)
            with apply_col1:
                if st.button(get_string("apply_ai_suggestions"), key="apply_gantt_suggestion", use_container_width=True):
                    st.session_state.task_name_input = st.session_state.gantt_task_suggestion.get("task_name", task_name)
                    st.session_state.gantt_task_suggestion = None
                    st.rerun()
            with apply_col2:
                if st.button(get_string("discard_ai_suggestions"), key="discard_gantt_suggestion", use_container_width=True):
                    st.session_state.gantt_task_suggestion = None
                    st.rerun()

        # Team members for task
        team_names = [t["name"] for t in team]
        task_team = st.multiselect(
            get_string("team"),
            options=team_names,
            default=editing_task["team_members"] if editing_task else [],
            key="task_team_select"
        )

        col1, col2 = st.columns(2)
        with col1:
            if st.button(f"✅ {get_string('add_task') if not editing_task else get_string('edit_gantt_task')}", use_container_width=True):
                if task_name:
                    if editing_task:
                        # Update task
                        db.update_gantt_task(
                            editing_task["id"],
                            name=task_name,
                            start_date=start_date.isoformat() if task_type == get_string("regular_task") else None,
                            end_date=end_date.isoformat() if task_type == get_string("regular_task") else None,
                            milestone_date=milestone_date.isoformat() if task_type == get_string("milestone") else None,
                            task_type="milestone" if task_type == get_string("milestone") else "task",
                            team_members=task_team
                        )
                        st.success("Task updated!")
                        st.session_state.editing_task_id = None
                    else:
                        # Create new task
                        db.create_gantt_task(
                            project_id=selected_project_id,
                            name=task_name,
                            start_date=start_date.isoformat() if task_type == get_string("regular_task") else None,
                            end_date=end_date.isoformat() if task_type == get_string("regular_task") else None,
                            milestone_date=milestone_date.isoformat() if task_type == get_string("milestone") else None,
                            task_type="milestone" if task_type == get_string("milestone") else "task",
                            team_members=task_team
                        )
                        st.success("Task added!")
                    st.rerun()
                else:
                    st.error(f"❌ {get_string('task_name')} {get_string('error')}")

        with col2:
            if editing_task and st.button("❌ Cancel", use_container_width=True):
                st.session_state.editing_task_id = None
                st.rerun()

    with tab2:
        st.subheader(get_string("generate_gantt"))

        node_ready, node_message = check_node_dependencies()
        if not node_ready:
            st.warning(node_message)

        if not gantt_tasks:
            st.warning(get_string("gantt_no_tasks"))
            st.info("💡 Use the 'Auto-Generate' button to create tasks from your latest weekly update.")
        else:
            st.write(f"📊 Ready to generate Gantt chart with {len(gantt_tasks)} tasks")

            if st.button(f"✨ {get_string('generate_gantt')}", use_container_width=True, disabled=not node_ready):
                with st.spinner(get_string("loading")):
                    result = generate_gantt_chart_from_tasks(selected_project_id, project["name"])

                    if result["success"]:
                        st.success(get_string("gantt_generated"))

                        # Provide download button
                        with open(result["file_path"], "rb") as f:
                            st.download_button(
                                label=get_string("download_gantt"),
                                data=f.read(),
                                file_name=os.path.basename(result["file_path"]),
                                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                                use_container_width=True,
                            )
                    else:
                        st.error(f"❌ {result['message']}")


# ============= PAGE: SETTINGS =============

def page_settings():
    """Settings page."""
    st.title(f"⚙️ {get_string('settings_title')}")

    # API Key
    st.subheader(get_string("api_key"))
    api_key = db.get_setting("deepseek_api_key", "")
    new_api_key = st.text_input(
        get_string("api_key"),
        value=api_key,
        type="password",
        placeholder=get_string("api_key_placeholder"),
    )

    if st.button(f"💾 {get_string('setting_saved')} API"):
        db.set_setting("deepseek_api_key", new_api_key)
        st.success(get_string("setting_saved"))

    if new_api_key:
        if st.button(get_string("test_connection")):
            with st.spinner(get_string("loading")):
                ai = AIAssistant(new_api_key)
                if ai.test_connection():
                    st.success(get_string("connection_success"))
                else:
                    st.error(get_string("connection_failed"))

    st.divider()

    # Your Name
    st.subheader(get_string("your_name"))
    user_name = db.get_setting("user_name", "")
    new_name = st.text_input(get_string("your_name"), value=user_name)

    if st.button(f"💾 {get_string('setting_saved')} Name"):
        db.set_setting("user_name", new_name)
        st.success(get_string("setting_saved"))

    st.divider()

    # Language
    st.subheader(get_string("language"))
    current_lang = st.session_state.language
    lang_options = {"English": "en", "Deutsch": "de"}
    selected_lang_name = [k for k, v in lang_options.items() if v == current_lang][0]

    new_lang_name = st.radio(
        get_string("language"),
        options=list(lang_options.keys()),
        index=list(lang_options.keys()).index(selected_lang_name),
    )

    if new_lang_name:
        new_lang_code = lang_options[new_lang_name]
        if new_lang_code != current_lang:
            db.set_setting("language", new_lang_code)
            st.session_state.language = new_lang_code
            st.success(get_string("setting_saved"))
            st.rerun()

    st.divider()

    # Export Folder
    st.subheader(get_string("export_folder"))
    export_path = db.get_setting("export_folder", "./exports/")
    new_export_path = st.text_input(get_string("export_folder"), value=export_path)

    if st.button(f"💾 {get_string('setting_saved')} Export"):
        db.set_setting("export_folder", new_export_path)
        st.success(get_string("setting_saved"))

    st.divider()

    # Reset Data
    st.subheader(f"🔴 {get_string('reset_data')}")
    if st.checkbox(get_string("reset_confirm"), key="reset_confirm_checkbox"):
        confirm_word = st.text_input("Type RESET to confirm", key="reset_confirm_word")
        if st.button(f"⚠️ {get_string('reset_data')}", key="reset_btn"):
            if confirm_word.strip().upper() != "RESET":
                st.error("Please type RESET to confirm.")
                return
            # Delete database
            if os.path.exists("project_updates.db"):
                os.remove("project_updates.db")
            st.success(f"✅ {get_string('success')}")
            st.rerun()


# ============= MAIN NAVIGATION =============

def main():
    """Main app."""
    pending_nav_page = st.session_state.get("pending_nav_page")
    if pending_nav_page:
        st.session_state.current_page = pending_nav_page
        st.session_state.nav_page = pending_nav_page
        st.session_state.pending_nav_page = None

    # Sidebar navigation
    with st.sidebar:
        st.title(f"🎯 {get_string('app_title')}")
        st.caption(get_string("app_subtitle"))

        st.divider()

        page = st.radio(
            "Navigate",
            [
                "Dashboard",
                "Projects",
                "Weekly Update",
                "Slide Generator",
                "Gantt Chart",
                "Final Documentation",
                "Settings",
            ],
            format_func=lambda x: {
                "Dashboard": f"🏠 {get_string('nav_dashboard')}",
                "Projects": f"📁 {get_string('nav_projects')}",
                "Weekly Update": f"📝 {get_string('nav_weekly_update')}",
                "Slide Generator": f"📊 {get_string('nav_slide_generator')}",
                "Gantt Chart": f"📈 {get_string('gantt_chart_title')}",
                "Final Documentation": f"📄 {get_string('nav_final_docs')}",
                "Settings": f"⚙️ {get_string('nav_settings')}",
            }.get(x, x),
            key="nav_page",
        )

        st.session_state.current_page = page

        st.divider()

        # Footer
        st.caption(f"v1.0.0 • {get_string('language')}: {st.session_state.language.upper()}")

    # Render page
    if st.session_state.current_page == "Dashboard":
        page_dashboard()
    elif st.session_state.current_page == "Projects":
        page_projects()
    elif st.session_state.current_page == "Weekly Update":
        page_weekly_update()
    elif st.session_state.current_page == "Slide Generator":
        page_slide_generator()
    elif st.session_state.current_page == "Gantt Chart":
        page_gantt_chart()
    elif st.session_state.current_page == "Final Documentation":
        page_final_documentation()
    elif st.session_state.current_page == "Settings":
        page_settings()


if __name__ == "__main__":
    main()
