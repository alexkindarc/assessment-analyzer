# Excel Online Structure

The application automatically creates and manages three worksheets in your Excel Online workbook.

---

## Worksheet 1: Results_Data

Stores metadata from **Results Reports**.

| Column | Description | Example |
|--------|-------------|---------|
| unit_id | Unique identifier for the unit | BSAD-ABC123 |
| unit_type | Academic or Administrative | Academic |
| unit_name | Canonical unit name | Business Administration Core BBA |
| college_division | College or division code | BSAD |
| degree_level | UG/GR/Doctoral/Certificate/NA | UG |
| modality | Delivery mode | On-campus, FWC |
| academic_year | Assessment year | 2024-2025 |
| report_type | Always "Results" | Results |
| outcome_id | Short outcome identifier | Legal-Social-Environment |
| outcome_text | Full outcome statement | Students will demonstrate... |
| outcome_label | Auto-set label | Student Learning Outcome |
| related_competency_or_function | Linked competency/function | Critical thinking skills |
| strategic_plan_theme | Mapped theme | Student Success |
| core_objective | TX core objective if applicable | Critical Thinking |
| assessment_method | Full method description | Cross-discipline exam in MANA 4322 |
| assessment_method_normalized | Simplified for comparison | Exam in capstone course |
| sample_size | Number assessed | 912 |
| benchmark | Criteria for success | 80% score 2/3 correct |
| result_value | Actual results | 10% |
| achievement_level | Status | Not Achieved |
| gap_from_benchmark | Difference | -70 |
| proposed_improvement | Planned action | Instructors will prioritize... |
| responsible_party | Who implements | BLAW 3310 instructors |
| improvement_timeline | When | Fall 2025 |
| upload_timestamp | When first uploaded | 2025-12-11T14:30:00 |
| last_updated | Last modification | 2025-12-11T14:30:00 |

---

## Worksheet 2: Improvement_Data

Stores metadata from **Improvement Reports**.

| Column | Description | Example |
|--------|-------------|---------|
| unit_id | Unique identifier | AA-FAC-456DEF |
| unit_type | Academic or Administrative | Administrative |
| unit_name | Canonical unit name | Faculty Success |
| college_division | Division | AA |
| academic_year | Year implemented | 2024-2025 |
| report_type | Always "Improvement" | Improvement |
| outcome_id | What was addressed | New Faculty Development Series |
| improvement_action_taken | What was done | Renamed program, combined cohorts... |
| connection_to_previous | Links to original finding? | Yes/No/Partial |
| previous_proposal_text | What was proposed before | (auto-populated) |
| upload_timestamp | When uploaded | 2025-12-11T14:30:00 |
| last_updated | Last modification | 2025-12-11T14:30:00 |

---

## Worksheet 3: Plan_Data

Stores metadata from **Next Cycle Plans**.

| Column | Description | Example |
|--------|-------------|---------|
| unit_id | Unique identifier | BSAD-ABC123 |
| unit_type | Academic or Administrative | Academic |
| unit_name | Canonical unit name | Business Administration Core BBA |
| college_division | College/division | BSAD |
| degree_level | Degree level | UG |
| academic_year | Year being planned | 2025-2026 |
| report_type | Always "Plan" | Plan |
| outcome_id | Outcome identifier | SLO-1 |
| outcome_text | Planned outcome | Students will demonstrate... |
| outcome_label | Auto-set | Student Learning Outcome |
| related_competency_or_function | Linked competency | Critical thinking |
| strategic_plan_theme | Mapped theme | Student Success |
| core_objective | Core objective if applicable | Communication |
| planned_method | How will assess | Portfolio review |
| planned_benchmark | Target criterion | 75% score proficient |
| action_steps | How students prepared | Course X will cover... |
| responsible_party | Who responsible | Program coordinator |
| upload_timestamp | When uploaded | 2025-12-11T14:30:00 |
| last_updated | Last modification | 2025-12-11T14:30:00 |

---

## Version Management

When you upload a report for the same unit + year + outcome that already exists:

1. The system finds the existing row using the unique key:
   - `unit_id` + `academic_year` + `outcome_id`

2. The existing row is **updated** (not duplicated)

3. The `last_updated` timestamp is refreshed

4. If the new upload has fewer outcomes than before, orphaned outcomes are **cleared**

---

## Using Excel Online Features

### Creating Filter Views

1. Select all data (Ctrl+A)
2. Go to **Data** tab → **Filter**
3. Click dropdown arrows to filter by any column

### Useful Filters to Create

**Needs Attention:**
- Filter `achievement_level` = "Not Achieved"
- Sort by `academic_year` descending

**By College:**
- Filter `college_division` = [select college]

**Current Year:**
- Filter `academic_year` = "2024-2025"

### Creating PivotTables

1. Select your data range
2. Go to **Insert** → **PivotTable**
3. Useful summaries:
   - Achievement level by college
   - Outcomes by strategic theme
   - Trends by academic year

### Conditional Formatting

Highlight "Not Achieved" outcomes:
1. Select the `achievement_level` column
2. **Home** → **Conditional Formatting** → **Highlight Cells Rules**
3. Select "Text that Contains" → "Not Achieved" → Red fill

---

## Sharing the Workbook

### Within UTA

1. Open the Excel file in SharePoint
2. Click **Share** in the top right
3. Enter UTA email addresses
4. Choose permission level:
   - **Can edit**: For IE staff
   - **Can view**: For deans, department chairs

### Important Notes

- The service account needs **Edit** access (configured during setup)
- Don't change the worksheet names (Results_Data, etc.)
- Don't delete or rearrange columns
- You CAN add additional columns to the right
- You CAN add additional worksheets for your own analysis

---

## Backup Strategy

### Automatic Versioning
- SharePoint automatically keeps version history
- View: Click file → **Version history**
- Restore: Click on any previous version

### Manual Backup
- Periodically download: **File** → **Save As** → **Download a Copy**
- Store in a separate location

---

## Troubleshooting Excel Issues

### Data Not Appearing

1. Refresh the browser
2. Check if you're looking at the correct worksheet tab
3. Verify the app shows "Connected to Microsoft 365" in sidebar

### Columns Out of Order

The app expects specific column positions. If columns were moved:
1. Download a backup copy first
2. Delete all worksheets in the file
3. Re-run a report analysis - app will recreate worksheets with correct structure

### "Used Range" Issues

If Excel reports strange used ranges:
1. Select all cells beyond your data
2. Right-click → **Delete** → **Entire Row**
3. Save the file

---

## Data Dictionary Quick Reference

| Worksheet | Purpose | Key Fields |
|-----------|---------|------------|
| Results_Data | Assessment results with outcomes | achievement_level, proposed_improvement |
| Improvement_Data | Actions taken on previous findings | improvement_action_taken, connection_to_previous |
| Plan_Data | Future assessment plans | planned_method, planned_benchmark |

### Achievement Level Values
- **Fully Achieved** - All criteria met
- **Partially Achieved** - Some criteria met
- **Not Achieved** - Criteria not met
- **Inconclusive** - Insufficient data

### Report Types
- **Results** - End-of-cycle with data
- **Improvement** - Actions taken
- **Plan** - Next cycle planning
