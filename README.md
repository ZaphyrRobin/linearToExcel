# Linear to Excel Planning Tool

A Python CLI that exports Linear issues to quarterly planning Excel spreadsheets.

## Features

- Export issues from a Linear team to Excel
- Filter by specific initiatives
- Auto-generate capacity section with assignees from issues
- Color-coded cells (yellow headers, green estimates)
- SUMIF formulas for per-assignee capacity calculation

## Project Structure

```
├── linear_to_excel.py      # Entry point
├── src/
│   ├── __init__.py
│   ├── linear_api.py       # Linear GraphQL API client
│   ├── excel_generator.py  # Excel generation logic
│   └── main.py             # CLI commands
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
# Generate planning spreadsheet
python linear_to_excel.py -t APP1 -q "Q4 2025"

# List initiatives hash numbers
python linear_to_excel.py --list-initiatives

# Filter by initiatives' hash numbers
python linear_to_excel.py -t APP1 -i "37e115a3f23c,75024d4f765d"

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
| H | Dependency |
| I | Assigned to |
| J+ | Weekly allocation |

## Requirements

- Python 3.8+
- Linear API key
