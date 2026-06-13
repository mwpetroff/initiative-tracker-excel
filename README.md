# Initiative Tracker — Excel Export

  A styled Excel workbook with VBA macros for managing organizational initiatives directly from Excel. Connects live to the Initiative Tracker REST API — create, edit, delete initiatives, and post progress updates without leaving Excel.

  ## What's in the zip

  | File | Description |
  |------|-------------|
  | `initiative_tracker.xlsx` | Styled workbook — Summary, Initiatives, Updates Log, and two hidden form sheets |
  | `initiative_tracker_macros.bas` | VBA module to import into the workbook |

  ## Setup (one-time)

  ### 1 — Import the VBA module
  1. Open `initiative_tracker.xlsx`
  2. Open the VBA editor: **Windows** → `Alt+F11` · **Mac** → `Option+F11`
  3. **File → Import File** → select `initiative_tracker_macros.bas`
  4. Close the editor and **Save As `.xlsm`** (macro-enabled format)

  ### 2 — Set the API URL
  Open the **Config** sheet and edit cell **B4**:
  - Local (app running on your machine): `http://localhost/api`
  - Deployed: `https://your-app.replit.app/api`

  ### 3 — Run `CreateButtons` once
  Press `Alt+F8` (Windows) or `Tools → Macro` (Mac), pick **CreateButtons**, and run it.  
  This builds the sidebar buttons on the Initiatives sheet and creates the two form sheets.

  ### 4 — Test the connection
  Run **TestConnection** — a "Connected!" popup confirms everything is working.

  ---

  ## Daily Use

  All actions are triggered by the **sidebar buttons** on the Initiatives sheet:

  | Button | What it does |
  |--------|--------------|
  | **Test Connection** | Verify the API URL is reachable |
  | **Create Initiative** | Opens a styled form sheet to fill in all fields |
  | **Edit Selected Row** | Click a data row first, then click — opens the form pre-filled |
  | **Delete Selected** | Click a row, confirm the prompt, row is removed from the sheet and the database |
  | **Add Update Note** | Click a row, opens a form to write a progress note |

  ### Form sheet UI

  Create and Edit open a dedicated **Initiative Form** sheet instead of a chain of popups:

  - Status and Priority are **dropdown menus** (data validation)
  - Progress is a validated 0–100 number field
  - All optional fields are present: Owner, Department, Description, Start/End Date
  - **Save** → calls the API *and* updates the spreadsheet row immediately
  - **Cancel** → closes the form and returns to the Initiatives sheet

  Add Update Note opens an **Update Form** sheet with a large note field (pre-filled author name) and a single **Post Update** button.

  ---

  ## Workbook sheets

  | Sheet | Purpose |
  |-------|---------|
  | **Config** | API URL + full setup instructions |
  | **Summary** | Executive KPI cards, status pie chart, department breakdown |
  | **Initiatives** | Full table with progress bars, status colours, sidebar buttons |
  | **Updates Log** | Chronological log of all progress notes |
  | **Initiative Form** | Hidden — VBA shows this for Create / Edit |
  | **Update Form** | Hidden — VBA shows this for Add Update |

  ---

  ## Compatibility

  | Platform | Requirement |
  |----------|-------------|
  | Windows | Excel 2016+ (uses `MSXML2.XMLHTTP60`, built-in) |
  | Mac | Excel 2016+ (uses `curl` via AppleScript, pre-installed on macOS) |

  No external add-ins, no COM dependencies. Pure VBA — works offline for reading; requires network access for API calls.

  ---

  ## Tech stack

  | Layer | Technology |
  |-------|-----------|
  | Workbook generation | [xlsxwriter](https://xlsxwriter.readthedocs.io/) (Python) |
  | REST API | Express 5 + Drizzle ORM (PostgreSQL) |
  | VBA macros | Pure VBA — Windows + Mac compatible |
  | Auth | None required (internal tool) |

  ---

  ## Re-exporting

  The workbook is generated from live database data. To refresh the snapshot:

  ```bash
  python scripts/generate_excel.py
  ```

  The output zip lands at the project root as `initiative_tracker_export.zip`.
  