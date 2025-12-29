# Weekly Marketing Report Generation Instructions

## Overview

This document provides instructions for generating formatted weekly marketing reports from raw lead data for Mayersohn Law Group. The output is an Excel file with multiple analysis sections.

---

## Input Data Structure

### Required CSV Columns

| Column Name | Description |
|-------------|-------------|
| `Creation Date` | Date the lead was created (MM/DD/YYYY) |
| `Month` | Month name (e.g., "December") |
| `Intake Agent Name` | Name of the intake agent who handled the lead |
| `Lead Name` | Name of the lead/prospect |
| `Contact Info (Phone)` | Phone number |
| `Type of Law` | Practice area (Family, Civil, Criminal, Real Estate, Estate Planning, Probate, Other) |
| `Consult Scheduled?` | "Yes" or "No" - whether a consultation was scheduled |
| `Reason Not Scheduled` | Reason code if not scheduled (e.g., "Scheduled", "In-Progress", "Outside Jurisdiction") |
| `No Show (Yes/No)` | "Yes" if they no-showed, "No" if they showed up |
| `Scheduled With` | Name of person the consult was scheduled with |
| `Lead Source` | Marketing channel (Google, Google LSA, Denise Isaacs, etc.) |
| `Zip Code` | Lead's zip code |
| `Marketing Qualified Lead (MQL)?` | "Yes" or "No" |
| `Sales Qualified Lead (SQL)?` | "Yes" or "No" |
| `Annual Income > $100k?` | "Yes" or "No" |
| `Case Urgency` | "Low", "Medium", or "High" |
| `Sales Outcome` | "Hired", "Did Not Hire", or "Follow-up" |
| `Marketing Contacts` | Marketing contact info if applicable |
| `Clio Profile` | Link to Clio profile |
| `Additional Notes` | Free text notes - **check for "Strategy Session" keyword** |
| `Case Revenue ($)` | Revenue amount if hired (e.g., "$10,000") |

---

## Key Business Logic

### Determining Showups vs No-Shows

```
IF "Consult Scheduled?" = "Yes" AND "No Show (Yes/No)" = "No" → SHOWUP
IF "Consult Scheduled?" = "Yes" AND "No Show (Yes/No)" = "Yes" → NO SHOW
```

**Important:** A lead SHOWED UP if "No Show (Yes/No)" = "No" (counterintuitive but correct)

### Identifying Strategy Sessions

```
IF "Additional Notes" CONTAINS "Strategy Session" (case-insensitive) → Strategy Session = True
```

Strategy sessions are a subset of hired leads where the hire was specifically for a strategy session rather than a full matter.

### Hired Categories

- **Hired (All):** All leads where `Sales Outcome` = "Hired"
- **Hired (Excl. Strategy):** Hired leads that are NOT strategy sessions
- **Strategy Sessions:** Hired leads where `Additional Notes` contains "Strategy Session"
- **Converted Sessions:** Same as Strategy Sessions (sessions that resulted in hire)
- **Unconverted Sessions:** Strategy sessions that did not convert (typically 0 in weekly reports)

### Revenue Calculation

- Clean the `Case Revenue ($)` column by removing `$` and `,` characters
- Convert to numeric, treating blanks/errors as 0
- Sum for totals

---

## Output Report Structure

### Sheet Layout

The report uses a single sheet with multiple sections arranged in columns:

```
Columns A-C: Summary Table, Reason Analysis
Columns E-Q: Intake Performance, Practice Area Breakdown, Source Analysis
```

### Section 1: Summary Table (Columns A-B, Rows 1-17)

| Metric | Formula |
|--------|---------|
| Total Leads | COUNT(all rows) |
| Total Hired (All) | COUNT WHERE Sales Outcome = "Hired" |
| Total Hired (Excluding Strategy Sessions) | Hired - Strategy Sessions |
| Total SQLs | COUNT WHERE SQL = "Yes" |
| Total MQLs | COUNT WHERE MQL = "Yes" |
| Total Scheduled | COUNT WHERE Consult Scheduled = "Yes" |
| Total Showups | COUNT WHERE Scheduled = "Yes" AND No Show = "No" |
| Total No Shows | COUNT WHERE Scheduled = "Yes" AND No Show = "Yes" |
| Total Marketing Contacts | COUNT WHERE Marketing Contacts is not empty |
| Average Daily Leads | Total Leads / 7 |
| Consultation Rate (%) | Total Scheduled / Total Leads |
| Conversion Rate (%) | Total Hired / Total Leads |
| MQL Rate (%) | Total MQLs / Total Leads |
| SQL Rate (%) | Total SQLs / Total Leads |
| Total Revenue | SUM of Case Revenue ($) |

### Section 2: Intake Performance (Columns E-Q, Rows 1-10+)

**Headers (Row 2):**
```
Agent | Total Leads | MQLs | SQLs | Scheduled | Showups | No Shows | Hired (All) | Hired (Excl. Strategy) | Strategy Sessions | Converted Sessions | Unconverted Sessions | Total Value
```

**Data:** Group by `Intake Agent Name` and calculate each metric per agent. Sort by Total Leads descending. Include a TOTAL row at the bottom.

### Section 3: Reason Analysis (Columns A-C, Rows 21+)

**Headers:**
```
Reason | Count | Percentage
```

**Data:** Value counts of `Reason Not Scheduled` column, sorted by count descending. Percentage = Count / Total * 100

### Section 4: Practice Area Breakdown (Columns E-N, Rows 23+)

For each practice area (Family, Civil, Criminal, Real Estate, Estate Planning, Probate, Other):

**Headers:**
```
Leads | MQLs | SQLs | Scheduled | Showups | New Matters | Marketing Contacts | Total Value | Strategy Sessions | Converted Sessions
```

**Notes:**
- "New Matters" = Hired (Excl. Strategy)
- "Other" category includes any Type of Law not in the main categories
- Leave 5 rows between each practice area section

### Section 5: Source Analysis (Columns E-O, after Practice Areas)

**Headers:**
```
Channel | Total Leads | MQLs | SQLs | Scheduled | Hired | Total Value | Conversion Rate (%) | Schedule Rate (%) | MQL Rate (%) | SQL Rate (%)
```

**Data:** Group by `Lead Source`, calculate metrics, sort by Total Leads descending. Include TOTAL row.

**Rate Calculations:**
- Conversion Rate = Hired / Total Leads * 100
- Schedule Rate = Scheduled / Total Leads * 100
- MQL Rate = MQLs / Total Leads * 100
- SQL Rate = SQLs / Total Leads * 100

---

## Python Implementation

### Required Libraries

```python
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
```

### Data Cleaning Steps

```python
# Read CSV
df = pd.read_csv('raw_data.csv')

# Clean revenue column
df['Case Revenue ($)'] = df['Case Revenue ($)'].replace(r'[\$,]', '', regex=True)
df['Case Revenue ($)'] = pd.to_numeric(df['Case Revenue ($)'], errors='coerce').fillna(0)

# Identify strategy sessions
df['Is_Strategy_Session'] = df['Additional Notes'].str.contains('Strategy Session', case=False, na=False)
```

### Key Calculations Template

```python
# Core metrics
total_leads = len(df)
total_hired = len(df[df['Sales Outcome'] == 'Hired'])
total_hired_excl_strategy = len(df[(df['Sales Outcome'] == 'Hired') & (~df['Is_Strategy_Session'])])
total_sqls = len(df[df['Sales Qualified Lead (SQL)?'] == 'Yes'])
total_mqls = len(df[df['Marketing Qualified Lead (MQL)?'] == 'Yes'])
total_scheduled = len(df[df['Consult Scheduled?'] == 'Yes'])
total_showups = len(df[(df['Consult Scheduled?'] == 'Yes') & (df['No Show (Yes/No)'] == 'No')])
total_no_shows = len(df[(df['Consult Scheduled?'] == 'Yes') & (df['No Show (Yes/No)'] == 'Yes')])
total_revenue = df['Case Revenue ($)'].sum()
```

### Grouping by Agent

```python
for agent_name, group in df.groupby('Intake Agent Name'):
    scheduled = group[group['Consult Scheduled?'] == 'Yes']
    showups = scheduled[scheduled['No Show (Yes/No)'] == 'No']
    hired_all = group[group['Sales Outcome'] == 'Hired']
    strategy_sessions = hired_all[hired_all['Is_Strategy_Session']]
    hired_excl_strategy = hired_all[~hired_all['Is_Strategy_Session']]
    revenue = group['Case Revenue ($)'].sum()
```

### Grouping by Practice Area

```python
practice_areas = ['Family', 'Civil', 'Criminal', 'Real Estate', 'Estate Planning', 'Probate']

for pa in practice_areas:
    pa_df = df[df['Type of Law'] == pa]
    # Calculate metrics...

# Handle "Other" category
other_df = df[~df['Type of Law'].isin(practice_areas)]
```

### Grouping by Lead Source

```python
for source_name, group in df.groupby('Lead Source'):
    leads = len(group)
    hired = len(group[group['Sales Outcome'] == 'Hired'])
    conversion_rate = hired / leads * 100 if leads > 0 else 0
```

---

## Formatting Guidelines

### Excel Styling

```python
# Header style
header_font = Font(bold=True)
header_fill = PatternFill('solid', fgColor='D9E1F2')

# Apply to header cells
cell.font = header_font
cell.fill = header_fill
```

### Column Widths

```python
ws.column_dimensions['A'].width = 38  # Metric names
ws.column_dimensions['B'].width = 15  # Values
# Data columns: 16 width each
```

### Number Formatting

- Currency: `"$"#,##0` or display as `$X,XXX`
- Percentages: Display as `XX%` or `XX.X%`
- Counts: Plain integers

---

## File Naming Convention

```
Weekly_[START_DATE]-[END_DATE]_Report.xlsx

Example: Weekly_12_13-12_19_Report.xlsx
```

---

## Validation Checklist

Before finalizing the report, verify:

- [ ] Total Leads matches row count in raw data
- [ ] Sum of agent leads equals Total Leads
- [ ] Sum of source leads equals Total Leads
- [ ] Revenue total matches sum of all Case Revenue values
- [ ] Showups + No Shows = Total Scheduled
- [ ] Hired (All) = Hired (Excl. Strategy) + Strategy Sessions
- [ ] All percentages are calculated correctly (not exceeding 100% unexpectedly)

---

## Common Issues & Solutions

### Issue: Revenue shows as 0 when there are hired leads
**Solution:** Check that revenue column is being cleaned properly - remove $ and , before converting to numeric

### Issue: Strategy sessions not being identified
**Solution:** Ensure case-insensitive search: `str.contains('Strategy Session', case=False, na=False)`

### Issue: Showups count seems wrong
**Solution:** Remember the counterintuitive logic - "No" in the No Show column means they DID show up

### Issue: Missing practice areas in breakdown
**Solution:** Check for exact spelling match in Type of Law column; handle variations

---

## Example Raw Data Row

```csv
12/19/2025,December,Lu,Elizabeth Reus,786-200-8585,Family,Yes,Scheduled,No,Lu,Denise Isaacs,33138,Yes,Yes,No,Medium,Hired,,https://grow.clio.com/matters/...,Hired a Full family mater,"$11,000"
```

**This row would be counted as:**
- 1 Lead (Claud)
- 1 MQL (MQL = Yes)
- 1 SQL (SQL = Yes)
- 1 Scheduled (Consult Scheduled = Yes)
- 1 Showup (No Show = No)
- 1 Hired (Sales Outcome = Hired)
- NOT a Strategy Session (Additional Notes doesn't contain "Strategy Session")
- $11,000 Revenue
- Family practice area
- Denise Isaacs source
