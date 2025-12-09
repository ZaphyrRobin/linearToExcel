# Linear to Excel Planning Tool

A Python CLI that exports Linear issues to quarterly planning Excel spreadsheets.

## Features

- Export issues from a Linear team to Excel
- Filter by specific initiatives
- Auto-generate capacity section with assignees from issues
- Color-coded cells (yellow headers, green estimates, gray initiative separators)
- SUMIF formulas for per-assignee capacity calculation
- Overwrite existing Excel files with latest data
- Append new tabs to existing Excel files (named by cycle start date)
- Generate multiple tabs by Linear cycle (each tab shows all issues, capacity filled per cycle)

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
# Generate new planning spreadsheet
python linear_to_excel.py -t APP1 -q "Q4 2025"

# List initiatives hash numbers
python linear_to_excel.py --list-initiatives

# Filter by initiatives' hash numbers
python linear_to_excel.py -t APP1 -i "37e115a3f23c,75024d4f765d" -o ~/Downloads/APP1_Q4_2025_planning.xlsx

# Overwrite existing file with latest Linear data
python linear_to_excel.py -t APP1 -i "37e115a3f23c,75024d4f765d" --input ~/Downloads/APP1_Q4_2025_planning.xlsx

# Append new tab to existing file (tab named by cycle start date)
python linear_to_excel.py -t APP1 -i "37e115a3f23c,75024d4f765d" --append ~/Downloads/APP1_Q4_2025_planning.xlsx

# Generate with separate tabs per Linear cycle
python linear_to_excel.py -t APP1 -i "37e115a3f23c,75024d4f765d" --by-cycles -o ~/Downloads/APP1_Q4_2025_planning.xlsx

# Generate with separate tabs per week (accumulated capacity)
python linear_to_excel.py -t APP1 -i "75024d4f765d" --by-weeks -o ~/Downloads/APP1_Q4_2025_planning.xlsx

python linear_to_excel.py --issue-history APP1-923

# List teams
python linear_to_excel.py --list-teams
```

## Options

| Option | Description |
|--------|-------------|
| `-t, --team` | Linear team key (required) |
| `-q, --quarter` | Quarter label (default: current) |
| `-o, --output` | Output filename |
| `-s, --start-date` | Start date (YYYY-MM-DD) |
| `-w, --weeks` | Number of weeks (default: 13) |
| `-i, --initiatives` | Comma-separated initiative slugs |
| `-f, --input` | Existing xlsx file to overwrite with latest data |
| `-a, --append` | Existing xlsx file to append a new tab to |
| `--by-cycles` | Create separate tabs for each Linear cycle |
| `--by-weeks` | Create separate tabs for each week with accumulated capacity |
| `--list-teams` | List available teams |
| `--list-initiatives` | List available initiatives |

## Output Format

| Column | Content |
|--------|---------|
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
