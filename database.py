"""
Database module for Project Update Studio.
Handles all SQLite database operations.
"""

import sqlite3
import json
from datetime import datetime
from contextlib import contextmanager
from pathlib import Path

DB_FILE = "project_updates.db"


@contextmanager
def get_db_connection():
    """Context manager for database connections."""
    conn = sqlite3.connect(DB_FILE)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    try:
        yield conn
    finally:
        conn.close()


def init_db():
    """Initialize the database with all required tables."""
    with get_db_connection() as conn:
        cursor = conn.cursor()

        # Projects table
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS projects (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL,
                description TEXT,
                goal TEXT,
                background TEXT,
                status TEXT DEFAULT 'active',
                rag_status TEXT DEFAULT 'green',
                language TEXT DEFAULT 'en',
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        """)

        # Team members table
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS team_members (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                project_id INTEGER NOT NULL,
                name TEXT NOT NULL,
                role TEXT,
                FOREIGN KEY (project_id) REFERENCES projects(id) ON DELETE CASCADE
            )
        """)

        # Weekly updates table
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS weekly_updates (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                project_id INTEGER NOT NULL,
                week_number INTEGER,
                week_label TEXT,
                rag_status TEXT,
                tasks_completed TEXT,
                next_tasks TEXT,
                management_decisions TEXT,
                risks_blockers TEXT,
                budget_status TEXT,
                budget_notes TEXT,
                milestone_hit TEXT,
                kpi_updates TEXT,
                stakeholder_notes TEXT,
                ai_summary TEXT,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                FOREIGN KEY (project_id) REFERENCES projects(id) ON DELETE CASCADE,
                UNIQUE(project_id, week_number)
            )
        """)

        # Gantt chart tasks table
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS gantt_tasks (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                project_id INTEGER NOT NULL,
                name TEXT NOT NULL,
                start_date TEXT,
                end_date TEXT,
                team_members TEXT,
                task_type TEXT DEFAULT 'task',
                milestone_date TEXT,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                FOREIGN KEY (project_id) REFERENCES projects(id) ON DELETE CASCADE
            )
        """)

        # Settings table
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS settings (
                key TEXT PRIMARY KEY,
                value TEXT
            )
        """)

        conn.commit()


# Projects CRUD


def create_project(
    name: str,
    language: str = "en",
    description: str = "",
    goal: str = "",
    background: str = "",
    team_members: list = None,
) -> int:
    """Create a new project. Returns project ID."""
    with get_db_connection() as conn:
        cursor = conn.cursor()
        now = datetime.now().isoformat()
        cursor.execute(
            """
            INSERT INTO projects
            (name, description, goal, background, language, created_at, updated_at)
            VALUES (?, ?, ?, ?, ?, ?, ?)
            """,
            (name, description, goal, background, language, now, now),
        )
        project_id = cursor.lastrowid

        # Add team members
        if team_members:
            for member in team_members:
                cursor.execute(
                    """
                    INSERT INTO team_members (project_id, name, role)
                    VALUES (?, ?, ?)
                    """,
                    (project_id, member.get("name"), member.get("role", "")),
                )

        conn.commit()
        return project_id


def get_all_projects() -> list:
    """Get all projects."""
    with get_db_connection() as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM projects ORDER BY created_at DESC")
        rows = cursor.fetchall()
        return [dict(row) for row in rows]


def get_project(project_id: int) -> dict:
    """Get a single project by ID."""
    with get_db_connection() as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM projects WHERE id = ?", (project_id,))
        row = cursor.fetchone()
        return dict(row) if row else None


def update_project(project_id: int, **kwargs) -> bool:
    """Update project fields. Returns True if successful."""
    allowed_fields = [
        "name",
        "description",
        "goal",
        "background",
        "status",
        "rag_status",
        "language",
    ]
    updates = {k: v for k, v in kwargs.items() if k in allowed_fields}

    if not updates:
        return False

    updates["updated_at"] = datetime.now().isoformat()

    with get_db_connection() as conn:
        cursor = conn.cursor()
        set_clause = ", ".join([f"{k} = ?" for k in updates.keys()])
        values = list(updates.values())
        values.append(project_id)

        cursor.execute(
            f"UPDATE projects SET {set_clause} WHERE id = ?", values
        )
        conn.commit()
        return cursor.rowcount > 0


def delete_project(project_id: int) -> bool:
    """Delete a project (cascades to team members and updates)."""
    with get_db_connection() as conn:
        cursor = conn.cursor()
        cursor.execute("DELETE FROM projects WHERE id = ?", (project_id,))
        conn.commit()
        return cursor.rowcount > 0


# Team members CRUD


def get_project_team(project_id: int) -> list:
    """Get all team members for a project."""
    with get_db_connection() as conn:
        cursor = conn.cursor()
        cursor.execute(
            "SELECT * FROM team_members WHERE project_id = ? ORDER BY name",
            (project_id,),
        )
        rows = cursor.fetchall()
        return [dict(row) for row in rows]


def add_team_member(project_id: int, name: str, role: str = "") -> int:
    """Add a team member to a project."""
    with get_db_connection() as conn:
        cursor = conn.cursor()
        cursor.execute(
            "INSERT INTO team_members (project_id, name, role) VALUES (?, ?, ?)",
            (project_id, name, role),
        )
        conn.commit()
        return cursor.lastrowid


def delete_team_member(member_id: int) -> bool:
    """Delete a team member."""
    with get_db_connection() as conn:
        cursor = conn.cursor()
        cursor.execute("DELETE FROM team_members WHERE id = ?", (member_id,))
        conn.commit()
        return cursor.rowcount > 0


def replace_project_team(project_id: int, team_members: list) -> bool:
    """Replace all team members for a project in one transaction."""
    with get_db_connection() as conn:
        cursor = conn.cursor()
        cursor.execute("DELETE FROM team_members WHERE project_id = ?", (project_id,))

        for member in team_members or []:
            name = (member.get("name") or "").strip()
            if not name:
                continue
            cursor.execute(
                """
                INSERT INTO team_members (project_id, name, role)
                VALUES (?, ?, ?)
                """,
                (project_id, name, (member.get("role") or "").strip()),
            )

        conn.commit()
        return True


# Weekly updates CRUD


def create_weekly_update(project_id: int, week_number: int, **kwargs) -> int:
    """Create a new weekly update. Returns update ID."""
    with get_db_connection() as conn:
        cursor = conn.cursor()
        now = datetime.now().isoformat()

        # Prepare default values
        data = {
            "project_id": project_id,
            "week_number": week_number,
            "week_label": kwargs.get("week_label", f"Week {week_number}"),
            "rag_status": kwargs.get("rag_status", "green"),
            "tasks_completed": kwargs.get("tasks_completed", "[]"),
            "next_tasks": kwargs.get("next_tasks", "[]"),
            "management_decisions": kwargs.get("management_decisions", "[]"),
            "risks_blockers": kwargs.get("risks_blockers", "[]"),
            "budget_status": kwargs.get("budget_status", "not_applicable"),
            "budget_notes": kwargs.get("budget_notes", ""),
            "milestone_hit": kwargs.get("milestone_hit", "[]"),
            "kpi_updates": kwargs.get("kpi_updates", "[]"),
            "stakeholder_notes": kwargs.get("stakeholder_notes", ""),
            "ai_summary": kwargs.get("ai_summary", ""),
            "created_at": now,
        }

        # Convert lists to JSON strings
        for key in [
            "tasks_completed",
            "next_tasks",
            "management_decisions",
            "risks_blockers",
            "milestone_hit",
            "kpi_updates",
        ]:
            if isinstance(data[key], list):
                data[key] = json.dumps(data[key])

        placeholders = ", ".join(["?"] * len(data))
        columns = ", ".join(data.keys())

        cursor.execute(
            f"INSERT INTO weekly_updates ({columns}) VALUES ({placeholders})",
            tuple(data.values()),
        )
        conn.commit()
        return cursor.lastrowid


def get_weekly_update(update_id: int) -> dict:
    """Get a single weekly update by ID."""
    with get_db_connection() as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM weekly_updates WHERE id = ?", (update_id,))
        row = cursor.fetchone()
        if not row:
            return None

        update = dict(row)
        # Parse JSON fields
        for key in [
            "tasks_completed",
            "next_tasks",
            "management_decisions",
            "risks_blockers",
            "milestone_hit",
            "kpi_updates",
        ]:
            try:
                update[key] = json.loads(update[key]) if update[key] else []
            except (json.JSONDecodeError, TypeError):
                update[key] = []

        return update


def get_project_updates(project_id: int) -> list:
    """Get all weekly updates for a project, ordered by week."""
    with get_db_connection() as conn:
        cursor = conn.cursor()
        cursor.execute(
            """
            SELECT * FROM weekly_updates
            WHERE project_id = ?
            ORDER BY week_number DESC
            """,
            (project_id,),
        )
        rows = cursor.fetchall()
        updates = []
        for row in rows:
            update = dict(row)
            # Parse JSON fields
            for key in [
                "tasks_completed",
                "next_tasks",
                "management_decisions",
                "risks_blockers",
                "milestone_hit",
                "kpi_updates",
            ]:
                try:
                    update[key] = json.loads(update[key]) if update[key] else []
                except (json.JSONDecodeError, TypeError):
                    update[key] = []
            updates.append(update)
        return updates


def get_weekly_update_by_week(
    project_id: int, week_number: int
) -> dict:
    """Get a weekly update for a specific project and week."""
    with get_db_connection() as conn:
        cursor = conn.cursor()
        cursor.execute(
            """
            SELECT * FROM weekly_updates
            WHERE project_id = ? AND week_number = ?
            """,
            (project_id, week_number),
        )
        row = cursor.fetchone()
        if not row:
            return None

        update = dict(row)
        # Parse JSON fields
        for key in [
            "tasks_completed",
            "next_tasks",
            "management_decisions",
            "risks_blockers",
            "milestone_hit",
            "kpi_updates",
        ]:
            try:
                update[key] = json.loads(update[key]) if update[key] else []
            except (json.JSONDecodeError, TypeError):
                update[key] = []

        return update


def update_weekly_update(update_id: int, **kwargs) -> bool:
    """Update a weekly update."""
    allowed_fields = [
        "week_label",
        "rag_status",
        "tasks_completed",
        "next_tasks",
        "management_decisions",
        "risks_blockers",
        "budget_status",
        "budget_notes",
        "milestone_hit",
        "kpi_updates",
        "stakeholder_notes",
        "ai_summary",
    ]

    updates = {}
    for key, value in kwargs.items():
        if key in allowed_fields:
            if isinstance(value, list):
                updates[key] = json.dumps(value)
            else:
                updates[key] = value

    if not updates:
        return False

    with get_db_connection() as conn:
        cursor = conn.cursor()
        set_clause = ", ".join([f"{k} = ?" for k in updates.keys()])
        values = list(updates.values())
        values.append(update_id)

        cursor.execute(
            f"UPDATE weekly_updates SET {set_clause} WHERE id = ?", values
        )
        conn.commit()
        return cursor.rowcount > 0


def delete_weekly_update(update_id: int) -> bool:
    """Delete a weekly update."""
    with get_db_connection() as conn:
        cursor = conn.cursor()
        cursor.execute("DELETE FROM weekly_updates WHERE id = ?", (update_id,))
        conn.commit()
        return cursor.rowcount > 0


# Gantt Tasks CRUD


def create_gantt_task(
    project_id: int,
    name: str,
    start_date: str = None,
    end_date: str = None,
    team_members: list = None,
    task_type: str = "task",
    milestone_date: str = None,
) -> int:
    """Create a new gantt task. Returns task ID."""
    with get_db_connection() as conn:
        cursor = conn.cursor()
        now = datetime.now().isoformat()
        
        team_members_json = json.dumps(team_members) if team_members else "[]"
        
        cursor.execute(
            """
            INSERT INTO gantt_tasks
            (project_id, name, start_date, end_date, team_members, task_type, milestone_date, created_at, updated_at)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (project_id, name, start_date, end_date, team_members_json, task_type, milestone_date, now, now),
        )
        conn.commit()
        return cursor.lastrowid


def get_gantt_tasks(project_id: int) -> list:
    """Get all gantt tasks for a project."""
    with get_db_connection() as conn:
        cursor = conn.cursor()
        cursor.execute(
            """
            SELECT * FROM gantt_tasks
            WHERE project_id = ?
            ORDER BY CASE WHEN task_type = 'milestone' THEN 1 ELSE 0 END,
                     start_date, milestone_date
            """,
            (project_id,),
        )
        rows = cursor.fetchall()
        tasks = []
        for row in rows:
            task = dict(row)
            try:
                task["team_members"] = json.loads(task["team_members"]) if task["team_members"] else []
            except (json.JSONDecodeError, TypeError):
                task["team_members"] = []
            tasks.append(task)
        return tasks


def update_gantt_task(task_id: int, **kwargs) -> bool:
    """Update a gantt task."""
    allowed_fields = [
        "name",
        "start_date",
        "end_date",
        "team_members",
        "task_type",
        "milestone_date",
    ]
    
    updates = {}
    for key, value in kwargs.items():
        if key in allowed_fields:
            if key == "team_members" and isinstance(value, list):
                updates[key] = json.dumps(value)
            else:
                updates[key] = value
    
    if not updates:
        return False
    
    updates["updated_at"] = datetime.now().isoformat()
    
    with get_db_connection() as conn:
        cursor = conn.cursor()
        set_clause = ", ".join([f"{k} = ?" for k in updates.keys()])
        values = list(updates.values())
        values.append(task_id)
        
        cursor.execute(
            f"UPDATE gantt_tasks SET {set_clause} WHERE id = ?", values
        )
        conn.commit()
        return cursor.rowcount > 0


def delete_gantt_task(task_id: int) -> bool:
    """Delete a gantt task."""
    with get_db_connection() as conn:
        cursor = conn.cursor()
        cursor.execute("DELETE FROM gantt_tasks WHERE id = ?", (task_id,))
        conn.commit()
        return cursor.rowcount > 0


def delete_project_gantt_tasks(project_id: int) -> bool:
    """Delete all gantt tasks for a project."""
    with get_db_connection() as conn:
        cursor = conn.cursor()
        cursor.execute("DELETE FROM gantt_tasks WHERE project_id = ?", (project_id,))
        conn.commit()
        return cursor.rowcount > 0


# Settings CRUD


def set_setting(key: str, value: str) -> None:
    """Set a setting value (creates or updates)."""
    with get_db_connection() as conn:
        cursor = conn.cursor()
        cursor.execute(
            "INSERT OR REPLACE INTO settings (key, value) VALUES (?, ?)",
            (key, value),
        )
        conn.commit()


def get_setting(key: str, default: str = None) -> str:
    """Get a setting value."""
    with get_db_connection() as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT value FROM settings WHERE key = ?", (key,))
        row = cursor.fetchone()
        return row[0] if row else default


def get_all_settings() -> dict:
    """Get all settings as a dictionary."""
    with get_db_connection() as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT key, value FROM settings")
        rows = cursor.fetchall()
        return {row[0]: row[1] for row in rows}


def delete_setting(key: str) -> bool:
    """Delete a setting."""
    with get_db_connection() as conn:
        cursor = conn.cursor()
        cursor.execute("DELETE FROM settings WHERE key = ?", (key,))
        conn.commit()
        return cursor.rowcount > 0


# Statistics


def get_project_stats(project_id: int) -> dict:
    """Get statistics for a project."""
    with get_db_connection() as conn:
        cursor = conn.cursor()

        # Weeks count
        cursor.execute(
            "SELECT COUNT(*) as count FROM weekly_updates WHERE project_id = ?",
            (project_id,),
        )
        weeks_count = cursor.fetchone()[0]

        # Team count
        cursor.execute(
            "SELECT COUNT(*) as count FROM team_members WHERE project_id = ?",
            (project_id,),
        )
        team_count = cursor.fetchone()[0]

        # Latest RAG status
        cursor.execute(
            "SELECT rag_status FROM weekly_updates WHERE project_id = ? ORDER BY week_number DESC LIMIT 1",
            (project_id,),
        )
        row = cursor.fetchone()
        latest_rag = row[0] if row else "green"

        # Latest update date
        cursor.execute(
            "SELECT MAX(created_at) FROM weekly_updates WHERE project_id = ?",
            (project_id,),
        )
        row = cursor.fetchone()
        latest_update = row[0] if row[0] else None

        return {
            "weeks_count": weeks_count,
            "team_count": team_count,
            "latest_rag": latest_rag,
            "latest_update": latest_update,
        }


def get_dashboard_stats() -> dict:
    """Get overall dashboard statistics."""
    with get_db_connection() as conn:
        cursor = conn.cursor()

        # Total projects
        cursor.execute("SELECT COUNT(*) as count FROM projects WHERE status = 'active'")
        total_active = cursor.fetchone()[0]

        # RAG status breakdown
        cursor.execute(
            """
            SELECT rag_status, COUNT(*) as count FROM projects WHERE status = 'active'
            GROUP BY rag_status
            """
        )
        rag_counts = {row[0]: row[1] for row in cursor.fetchall()}

        return {
            "total_active": total_active,
            "green": rag_counts.get("green", 0),
            "amber": rag_counts.get("amber", 0),
            "red": rag_counts.get("red", 0),
        }
