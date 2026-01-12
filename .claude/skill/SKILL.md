# Linear to Excel Planning Tool

CLI tool that exports Linear issues to Excel planning spreadsheets.

## Structure

```
linear_to_excel.py      # CLI entry point
src/
├── __init__.py
├── linear_api.py       # Linear GraphQL API client
└── excel_generator.py  # Excel generation logic
```

## Usage

```bash
# 1. List initiatives to get their IDs
python linear_to_excel.py --list-initiatives

# 2. Generate new planning spreadsheet
python linear_to_excel.py -t APP1 -i "37e115a3f23c,75024d4f765d" -s 2025-10-06 -e 2026-01-30 -o ~/Downloads/APP1_Q4_2025_planning.xlsx

# 3. Refresh existing file with latest Linear data (preserves manual edits)
python linear_to_excel.py -t APP1 -i "37e115a3f23c,75024d4f765d" -s 2025-10-06 -e 2026-01-30 -f ~/Downloads/APP1_Q4_2025_planning.xlsx
```

## Modules

### src/linear_api.py
- `get_api_key()` - Get LINEAR_API_KEY from env
- `linear_request(query, variables)` - GraphQL request helper
- `fetch_teams()` - Get all teams
- `fetch_all_initiatives(include_archived=False)` - Get initiatives (excludes [Archive])
- `fetch_issues_for_team(team_id, initiative_slugs)` - Paginated issue fetch with filter
- `get_team_by_key(team_key)` - Find team by key
- `fetch_issue_by_identifier(identifier)` - Fetch single issue by ID
- `fetch_issue_history(issue_id)` - Fetch history for an issue

### src/excel_generator.py
- `generate_week_dates(start_date, num_weeks)` - Generate week dates
- `extract_unique_assignees(issues)` - Get unique assignees from issues
- `create_excel(...)` - Generate new Excel with capacity section, SUMIF formulas, styling
- `refresh_excel(...)` - Refresh existing Excel, preserving manual edits when Linear data is missing
- `read_existing_capacity_data(wb, start_date)` - Read existing capacity/assignee data from Excel

### linear_to_excel.py
- CLI using Click
- Options: -t/--team, -o/--output, -s/--start-date, -e/--end-date, -i/--initiatives, -f/--file, --list-teams, --list-initiatives, --issue-history

## Refresh Logic

When refreshing an existing Excel file (`-f` option):
- **Linear has assignee AND cycle**: Uses Linear data for estimate placement
- **Linear missing assignee OR cycle**: Preserves existing Excel estimate placement and assignee
- Column A shows "Linear" or "Estimated" to indicate data source

## Colors

- Yellow `#FFF2CC`: Headers, unassigned estimates
- Green `#B7E1CD`: Estimates with assignee
- Gray `#D9D9D9`: Initiative separators
