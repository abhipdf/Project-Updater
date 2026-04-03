# 📊 Project Update Studio

**A Streamlit web application for project managers to generate weekly status update PowerPoint slides and comprehensive project documentation Word files — powered by AI.**

## Overview

Project Update Studio is designed for operations and process improvement project managers who need to manage 4–7 simultaneous projects and generate professional weekly status updates and final project documentation without manually writing presentations.

### Key Features

✅ **AI-Powered Weekly Interviews** — Let an AI assistant guide you through structured interviews to capture all project information  
✅ **Automatic PowerPoint Generation** — Generate professional single-slide weekly updates with all key information  
✅ **Comprehensive Documentation** — Create multi-section Word documents summarizing the entire project lifecycle  
✅ **Multi-Language Support** — Full English and German language support throughout the app and generated documents  
✅ **Project Management Dashboard** — Visual overview of all active projects with RAG (Red/Amber/Green) status tracking  
✅ **DeepSeek AI Integration** — Uses DeepSeek API for intelligent interview guidance and content generation  
✅ **SQLite Database** — All data stored locally in a single file for data privacy and easy backup  

---

## Prerequisites

Before you begin, ensure you have the following installed:

- **Python** 3.11 or higher  
- **Node.js** 18 or higher (for PowerPoint generation)  
- **pip** (Python package manager)  
- **npm** (Node.js package manager)  
- A **DeepSeek API key** (get one at https://platform.deepseek.com/)

### Verify Installation

```bash
python --version          # Should be 3.11+
node --version           # Should be 18+
npm --version            # Should be installed
```

---

## Installation & Setup

### Step 1: Clone or Navigate to Project

```bash
cd project_update_studio
```

### Step 2: Install Python Dependencies

```bash
pip install -r requirements.txt
```

### Step 3: Install Node.js Dependencies

```bash
npm install
```

### Step 4: Run the Application

```bash
streamlit run app.py
```

The app will open in your default browser at `http://localhost:8501`

### Step 5: Configure Settings

1. **Go to Settings** (⚙️ icon in sidebar)
2. **Enter your DeepSeek API Key** — Get one free at https://platform.deepseek.com/
3. **Click "Test Connection"** to verify the key works
4. **Set your preferred Language** — English or German
5. **Enter your name** — Used as default assignee in tasks

---

## Quick Start Guide

### 1. Create Your First Project

1. Click **Projects** (📁) in the sidebar
2. Fill in:
   - **Project Name** — e.g., "Process Optimization Q2 2025"
   - **Language** — English or German
   - **Description** — What the project is about
   - **Goal** — Main objective
   - **Background** — Additional context
   - **Team Members** — Add 1-5 team members with roles
3. Click **Create Project**
4. Answer the **AI Onboarding Interview** questions (8 questions about your project)
5. Review and confirm the generated project brief
6. Click **Confirm & Save**

### 2. Create Your First Weekly Update

1. Click **Weekly Update** (📝) in the sidebar
2. Select your project from the dropdown
3. Answer the **AI Interview** questions (10 questions about this week's progress)
4. Review and edit the extracted data if needed
5. Click **Save Update**

### 3. Generate Your First Slide

1. Click **Slide Generator** (📊) in the sidebar
2. Select your project and week
3. Click **Generate Slide**
4. Download the PowerPoint file (`.pptx`)

The slide includes:
- Project name and week label
- This week's completed tasks
- Team members
- Next steps
- KPIs/Metrics (if available)
- Management decisions required
- AI-generated executive summary
- Professional footer with timestamp and confidentiality notice

### 4. Generate Final Documentation

1. Click **Final Documentation** (📄) in the sidebar
2. Select your project
3. Click **Generate Documentation**
4. Download the Word document (`.docx`)

The document includes:
- Cover page with project info
- Project overview section
- Weekly updates log (all weeks)
- Key milestones & achievements
- Decisions log
- Risks & issues log
- AI-generated project closure summary

---

## File Structure

```
project_update_studio/
├── app.py                    # Main Streamlit application with all pages
├── database.py               # SQLite database operations and schema
├── ai_assistant.py           # DeepSeek API integration for AI features
├── doc_generator.py          # Word document generation (python-docx)
├── slide_generator.js        # PowerPoint slide generation (Node.js + pptxgenjs)
├── utils.py                  # Shared utilities, language strings, colors
├── requirements.txt          # Python dependencies
├── package.json              # Node.js dependencies
├── README.md                 # This file
├── project_updates.db        # SQLite database (created automatically)
└── exports/                  # Output folder for generated .pptx and .docx files
```

---

## Database Schema

### Projects Table
Stores all project information:
```
id, name, description, goal, background, status, rag_status, language, created_at, updated_at
```

### Team Members Table
Stores team member information:
```
id, project_id, name, role
```

### Weekly Updates Table
Stores all weekly status update data:
```
id, project_id, week_number, week_label, rag_status, tasks_completed, next_tasks,
management_decisions, risks_blockers, budget_status, budget_notes, milestone_hit,
kpi_updates, stakeholder_notes, ai_summary, created_at
```

### Settings Table
Stores user preferences:
```
key (primary), value
```

Example settings keys:
- `deepseek_api_key` — Your DeepSeek API key
- `user_name` — Your name
- `language` — App language (en/de)
- `export_folder` — Export folder path

---

## Features Explained

### 🏠 Dashboard
Overview of all active projects with:
- Total project count and RAG status breakdown
- Project cards showing name, RAG status, weeks of updates, team size, last update date
- Quick action buttons to create new updates or generate slides

### 📁 Projects
Manage your projects:
- **Create New Project** — Start a new project with AI onboarding interview
- **Project List** — View all projects, edit, archive, or delete
- Each project shows: description, goal, team members, status, RAG

### 📝 Weekly Update
AI-guided interview to create weekly updates:
- Conversational questions covering tasks, blockers, KPIs, decisions, budget, next steps
- AI automatically extracts structured data
- Edit interface to refine extracted information
- Automatic duplicate week detection with option to edit existing

### 📊 Slide Generator
Generate professional PowerPoint slides:
- Single-slide format (16:9 widescreen)
- Automatic layout with Philips Ocean Blue color scheme
- Shows: tasks, team, next steps, KPIs, management decisions
- AI-generated executive summary in footer
- Downloads as `.pptx` file

### 📄 Final Documentation
Generate comprehensive Word documents:
- Multi-section format with professional styling
- Sections: Overview, Weekly Log, Milestones, Decisions, Risks, Closure Summary
- AI-generated project closure summary
- All updates and decisions compiled with week references
- Professional styling with ocean blue headings, page numbers, headers

### ⚙️ Settings
Configure user preferences:
- **DeepSeek API Key** — Required for all AI features; includes test button
- **Your Name** — Used as default assignee
- **Language** — Global app language (English/Deutsch)
- **Export Folder** — Where to save generated files
- **Reset Data** — Wipe all data (use with caution!)

---

## Color Palette

**Philips Ocean Blue theme applied throughout:**

| Color | Hex | Use |
|-------|-----|-----|
| Primary Blue | `#0B5ED7` | Buttons, links, primary accents |
| Dark Blue | `#003087` | Headers, dark backgrounds |
| Light Blue Accent | `#E8F0FE` | Background panels, light accents |
| Success Green | `#28A745` | Green RAG status, success messages |
| Warning Amber | `#FFC107` | Amber RAG status, warnings |
| Danger Red | `#DC3545` | Red RAG status, errors |
| Text Dark | `#1A1A2E` | Body text |
| White | `#FFFFFF` | Background |

---

## AI Features

### Project Onboarding Interview
When you create a new project, the AI asks 8 key questions:
1. What problem or inefficiency does this project address?
2. What is the desired end state or goal?
3. What is the baseline (current metrics)?
4. What is the timeline and key milestones?
5. What are the main risks or dependencies?
6. Who are the key stakeholders?
7. What does success look like?
8. Is there a budget involved?

The AI then generates a project brief from your answers for confirmation.

### Weekly Update Interview
For each week's update, the AI asks 10 conversational questions covering:
- This week's completed tasks and results
- Milestones reached
- RAG status assessment
- KPI/metric updates
- Blockers and risks encountered
- Mitigation plans
- Budget status
- Next week's planned tasks
- Management decisions required
- Stakeholder communications

The AI extracts structured data and generates an executive summary in 3–5 sentences suitable for executives.

### Project Closure Summary
When generating final documentation, the AI reads all updates and generates a ~300-word professional closure summary covering:
- Project overview and objectives
- Key achievements
- Final status and deliverables
- Lessons learned
- Closing remarks

---

## Language Support

The entire app supports **English and German**:

- ✅ All UI labels and navigation
- ✅ All form fields and prompts
- ✅ AI interview questions
- ✅ All generated content (slides, documents)
- ✅ Error messages and confirmations

**Set your language in Settings (⚙️)** — changes apply globally.

---

## How the AI Works

### DeepSeek Integration

All AI features use the **DeepSeek API** (model: `deepseek-chat`):

- **Endpoint:** `https://api.deepseek.com/v1/chat/completions`
- **Authentication:** Bearer token (your API key)
- **Communication:** Direct HTTP requests (no SDK)

### Interview Flow

1. **System Prompt** includes:
   - Your preferred language
   - Full project context (goal, description, background)
   - Summary of all previous updates (for context continuity)
   - Instructions to ask one question at a time

2. **Question Asked** — AI asks one question at a time, contextually

3. **Answer Captured** — User provides answer, question index increments

4. **Data Extraction** — After all questions, AI analyzes conversation and extracts structured JSON with:
   - Completed tasks (with results and owners)
   - Next tasks (with assignments and due dates)
   - RAG status (with reasoning)
   - Management decisions (with urgency levels)
   - Risks/blockers (with impact and mitigation)
   - KPI updates
   - Budget status
   - Executive summary

5. **User Review** — All extracted data shown for editing before saving

---

## Data Privacy & Security

✅ **All data stored locally** in SQLite database (`project_updates.db`)  
✅ **No cloud storage** — only DeepSeek API calls leave your machine  
✅ **API key stored locally** in encrypted settings (accessible only within the app)  
✅ **Easy backup** — Simply copy the `project_updates.db` file  
✅ **Easy deletion** — Use Settings → Reset Data to wipe all data  

---

## Troubleshooting

### "⚠️ DeepSeek API key is not set"
→ Go to **Settings** (⚙️), enter your API key, and click the test button.

### "Connection failed" when testing API
→ Check:
- API key is correct (copy from https://platform.deepseek.com/)
- Internet connection is active
- DeepSeek service is not down

### Slide generation fails
→ Check:
- Node.js is installed (`node --version`)
- npm dependencies installed (`npm install`)
- Exports folder exists

### Word document generation fails
→ Check:
- python-docx is installed (`pip list | grep docx`)
- Exports folder has write permissions

### "File not found: project_updates.db"
→ This is normal! The database is created automatically on first run.

### App is slow
→ This can happen with large projects:
- Reduce the number of weekly updates shown
- Archive old projects
- Clear old data in Settings

---

## Tips & Best Practices

### 🎯 Project Setup
- **Clear project goal** — The more specific, the better the AI guidance
- **Accurate background** — Include relevant context about the problem being solved
- **Complete team** — Add all active team members even if not always present

### 📝 Weekly Updates
- **Answer honestly** — AI quality depends on truthful answers
- **Use consistent terminology** — Refer to same metrics, risks by same names
- **Check the AI summary** — Edit if it doesn't capture nuances
- **Review extracted data** — Correct any misinterpretations before saving

### 📊 Slide Generation
- **Check for overflows** — Long task descriptions may truncate
- **Review before sharing** — Always check generated slide for accuracy
- **Update archived projects** — Final slide often your best/final snapshot

### 📄 Documentation
- **Include all weeks** — Complete project reviews are valuable
- **Review before distribution** — Check cover page, milestones, closure summary
- **Archive project when done** — Keeps dashboard clean

---

## Requirements

### Python Packages
- `streamlit>=1.32.0` — Web app framework
- `requests>=2.31.0` — HTTP requests for DeepSeek API
- `python-docx>=1.1.0` — Word document generation
- `python-dateutil>=2.8.2` — Date parsing and formatting

### Node.js Packages
- `pptxgenjs@3.12.0` — PowerPoint generation

---

## Roadmap

Potential future enhancements:
- Timeline/Gantt chart visualizations
- Budget tracking and charts
- Risk register visualization
- Export to other formats (PDF, Excel)
- Team member dashboard view
- Email notifications for weekly reminders
- Slack integration
- Custom slide templates

---

## Support & Contributing

For issues, feature requests, or contributions:

1. Check this README for troubleshooting
2. Verify all prerequisites are installed
3. Check the database isn't corrupted: `sqlite3 project_updates.db ".tables"`
4. Review logs in your Streamlit Console

---

## License

MIT License — Feel free to use, modify, and distribute.

---

## Credits

Built with:
- **Streamlit** — Web app framework
- **SQLite** — Database
- **pptxgenjs** — PowerPoint generation
- **python-docx** — Word document generation
- **DeepSeek API** — AI intelligence
- **Philips Design System** — Ocean Blue color palette

---

## Version

**v1.0.0** — Initial Release  
Last Updated: April 2025

---

## FAQ

**Q: Can I use this offline?**  
A: Almost! All features work offline except for AI-powered interviews and content generation, which require the DeepSeek API.

**Q: How much does DeepSeek cost?**  
A: See https://platform.deepseek.com/pricing. It's very affordable with per-token pricing.

**Q: Can I export data from the database?**  
A: Yes! You can use any SQLite client to query the `project_updates.db` file directly.

**Q: Is my data encrypted?**  
A: Data is stored locally in plain SQLite. For sensitive data, ensure physical machine security.

**Q: Can I run this on the cloud (AWS, Azure, etc.)?**  
A: Yes! Deploy the Streamlit app to any cloud provider. Database stays local to the instance.

**Q: What if I lose my database file?**  
A: All data is lost. Keep regular backups of `project_updates.db`.

**Q: Can I edit a project after creation?**  
A: Not yet — feature coming soon. For now, archive and create new projects.

**Q: Can multiple users work on the same project?**  
A: If you run the app on a shared server, yes. Each project tracks who made updates via timestamps.

---

## Getting Help

- 📖 **Read this README** — Most common questions are answered here
- ⚙️ **Check Settings** — Test your API key connection
- 🔄 **Restart the app** — Run `streamlit run app.py` again
- 💾 **Backup your data** — Copy `project_updates.db` to a safe location

---

Happy project managing! 🚀
