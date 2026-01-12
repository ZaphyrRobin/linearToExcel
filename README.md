# Linear to Excel Planning Tool

A Python CLI that exports Linear issues to quarterly planning Excel spreadsheets.

## Features

- Export issues from a Linear team to Excel
- Filter by specific initiatives
- Auto-generate capacity section with assignees from issues
- Color-coded cells (yellow headers, green estimates, gray initiative separators)
- SUMIF formulas for per-assignee capacity calculation
- Refresh existing Excel files with latest Linear data (preserves manual edits when Linear data is missing)
- "Linear vs Estimated" column to indicate data source for each issue

## Project Structure

```
├── linear_to_excel.py      # CLI entry point
├── src/
│   ├── __init__.py
│   ├── linear_api.py       # Linear GraphQL API client
│   └── excel_generator.py  # Excel generation logic
├── requirements.txt
└── .env                    # API key (gitignored)
```

## Setup

1. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

2. Create `.env` with your Linear API key:
   ```
   LINEAR_API_KEY=lin_api_xxxxxxxxxxxxx
   ```
   Generate at: Linear Settings → API → Personal API keys

## Usage

```bash
# 1. List initiatives to get their IDs
python linear_to_excel.py --list-initiatives

# 2. Generate new planning spreadsheet
python linear_to_excel.py -t APP1 -i "37e115a3f23c,75024d4f765d" -s 2025-10-06 -e 2026-01-30 -o ~/Downloads/APP1_Q4_2025_planning.xlsx

# 3. Refresh existing file with latest Linear data (preserves manual edits)
python linear_to_excel.py -t APP1 -i "37e115a3f23c,75024d4f765d" -s 2025-10-06 -e 2026-01-30 -f ~/Downloads/APP1_Q4_2025_planning.xlsx
```

## Options

| Option | Description |
|--------|-------------|
| `-t, --team` | Linear team key (required) |
| `-o, --output` | Output filename |
| `-s, --start-date` | Start date (YYYY-MM-DD) |
| `-e, --end-date` | End date (YYYY-MM-DD) - calculates number of weeks |
| `-i, --initiatives` | Comma-separated initiative slugs |
| `-f, --file` | Existing xlsx file to refresh with latest data |
| `--issue-history` | Show history of a specific issue (e.g., 'APP1-123') |
| `--list-teams` | List available teams |
| `--list-initiatives` | List available initiatives |

## Output Format

| Column | Content |
|--------|---------|
| A | Linear vs Estimated (data source indicator) |
| B | Initiative |
| C | Project |
| D | Issue title |
| E | Estimate (days) |
| F | Description |
| G | Linear URL |
| H | Assigned to (first name) |
| I+ | Weekly allocation |

## Requirements

- Python 3.8+
- Linear API key
