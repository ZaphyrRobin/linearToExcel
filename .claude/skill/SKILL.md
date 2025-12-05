# Linear to Excel Planning Tool

CLI tool that exports Linear issues to Excel planning spreadsheets.

## Structure

```
linear_to_excel.py      # Entry point
src/
├── __init__.py
├── linear_api.py       # Linear GraphQL API client
├── excel_generator.py  # Excel generation logic
└── main.py             # CLI commands
```

## Modules

### src/linear_api.py
- `get_api_key()` - Get LINEAR_API_KEY from env
- `linear_request(query, variables)` - GraphQL request helper
- `fetch_teams()` - Get all teams
- `fetch_all_initiatives(include_archived=False)` - Get initiatives (excludes [Archive])
- `fetch_issues_for_team(team_id, initiative_slugs)` - Paginated issue fetch with filter
- `get_team_by_key(team_key)` - Find team by key

### src/excel_generator.py
- `generate_week_dates(start_date, num_weeks)` - Generate week dates
- `extract_unique_assignees(issues)` - Get unique assignees from issues
- `create_excel(team_name, quarter, issues, output_file, start_date, num_weeks)` - Generate Excel

### src/main.py
- CLI entry point using Click
- Options: --team, --quarter, --output, --start-date, --weeks, --initiatives, --list-teams, --list-initiatives

## Colors

- Yellow `#FFF2CC`: Headers
- Green `#B7E1CD`: Estimates
