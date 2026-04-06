"""
Gantt Chart Generator - Convert project tasks to gantt chart PPTX.
Handles the conversion of gantt tasks to JSON format and calls the Node.js generator.
"""

import json
import subprocess
import os
from datetime import datetime, timedelta
import database as db
from pathlib import Path


def generate_gantt_chart_from_tasks(project_id: int, project_name: str) -> dict:
    """
    Generate a gantt chart from stored gantt tasks.
    Returns: {"success": bool, "file_path": str, "message": str}
    """
    try:
        # Get all gantt tasks for the project
        tasks = db.get_gantt_tasks(project_id)
        
        if not tasks:
            return {
                "success": False,
                "file_path": None,
                "message": "No gantt chart tasks defined for this project."
            }
        
        # Build the gantt chart JSON format
        gantt_data = {
            "projectName": project_name,
            "tasks": [],
            "gates": []
        }
        
        for task in tasks:
            if task["task_type"] == "milestone":
                # Add as milestone
                gantt_data["tasks"].append({
                    "name": task["name"],
                    "date": task["milestone_date"],
                    "type": "milestone"
                })
            else:
                # Add as regular task
                team_members = task.get("team_members", [])
                gantt_data["tasks"].append({
                    "name": task["name"],
                    "start": task["start_date"],
                    "end": task["end_date"],
                    "team": team_members if isinstance(team_members, list) else []
                })
        
        # Generate the gantt chart
        file_path = _call_gantt_script(gantt_data)
        
        return {
            "success": True,
            "file_path": file_path,
            "message": f"Gantt chart generated successfully: {file_path}"
        }
    
    except Exception as e:
        return {
            "success": False,
            "file_path": None,
            "message": f"Error generating gantt chart: {str(e)}"
        }


def _call_gantt_script(gantt_data: dict) -> str:
    """
    Call the Node.js gantt_chart.js script with the gantt data.
    Returns the path to the generated PPTX file.
    """
    # Convert gantt data to JSON string (escaped for shell)
    json_input = json.dumps(gantt_data)
    
    # Create exports directory if it doesn't exist
    exports_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "exports")
    os.makedirs(exports_dir, exist_ok=True)
    
    # Get the script path
    script_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "gantt_chart.js")
    
    # Call Node.js script
    try:
        result = subprocess.run(
            ["node", script_path, json_input],
            cwd=exports_dir,
            capture_output=True,
            text=True,
            timeout=30
        )
        
        if result.returncode != 0:
            raise Exception(f"Node.js error: {result.stderr}")
        
        # The script saves to gantt_chart.pptx in the current working directory
        output_file = os.path.join(exports_dir, "gantt_chart.pptx")
        
        if not os.path.exists(output_file):
            raise Exception("Generated PPTX file not found")
        
        return output_file
    
    except subprocess.TimeoutExpired:
        raise Exception("Gantt chart generation timed out")
    except FileNotFoundError:
        raise Exception("Node.js not found. Please ensure Node.js is installed.")


def auto_generate_gantt_from_tasks(project_id: int) -> dict:
    """
    Auto-generate gantt tasks from the latest weekly update's task data.
    This allows users to create a gantt chart without explicitly defining dates.
    Returns the number of tasks created.
    """
    try:
        # Get the latest weekly update
        updates = db.get_project_updates(project_id)
        if not updates:
            return {
                "success": False,
                "count": 0,
                "message": "No weekly updates found for this project."
            }
        
        latest_update = updates[0]  # Most recent update
        
        # Clear existing gantt tasks for this project
        db.delete_project_gantt_tasks(project_id)
        
        task_count = 0
        project = db.get_project(project_id)
        team = db.get_project_team(project_id)
        team_names = [t["name"] for t in team]
        
        # Process next tasks and create gantt tasks from them
        next_tasks = latest_update.get("next_tasks", [])
        for task in next_tasks:
            task_name = task.get("task", "") if isinstance(task, dict) else str(task)
            if not task_name:
                continue
            
            # Extract task details
            owner = task.get("owner", "") if isinstance(task, dict) else ""
            due_date = task.get("due_date", "") if isinstance(task, dict) else ""
            
            # Create gantt task with estimated start/end dates
            start_date = datetime.now().date().isoformat()
            
            # If due date exists, use it; otherwise estimate 2 weeks from now
            if due_date:
                try:
                    end_date = due_date
                except:
                    end_date = (datetime.now() + timedelta(days=14)).date().isoformat()
            else:
                end_date = (datetime.now() + timedelta(days=14)).date().isoformat()
            
            # Determine team members (owner + default team if owner not in team)
            team_for_task = [owner] if owner else team_names[:1] if team_names else []
            
            db.create_gantt_task(
                project_id=project_id,
                name=task_name,
                start_date=start_date,
                end_date=end_date,
                team_members=team_for_task,
                task_type="task"
            )
            task_count += 1
        
        # Process milestones from the latest update
        milestones = latest_update.get("milestone_hit", [])
        for milestone in milestones:
            milestone_name = milestone.get("name", "") if isinstance(milestone, dict) else str(milestone)
            if not milestone_name:
                continue
            
            milestone_date = milestone.get("date", "") if isinstance(milestone, dict) else ""
            if not milestone_date:
                milestone_date = datetime.now().date().isoformat()
            
            db.create_gantt_task(
                project_id=project_id,
                name=milestone_name,
                task_type="milestone",
                milestone_date=milestone_date
            )
            task_count += 1
        
        return {
            "success": True,
            "count": task_count,
            "message": f"Auto-generated {task_count} gantt tasks from latest weekly update."
        }
    
    except Exception as e:
        return {
            "success": False,
            "count": 0,
            "message": f"Error auto-generating gantt tasks: {str(e)}"
        }
