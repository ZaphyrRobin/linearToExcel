"""Excel spreadsheet generation for planning documents."""

from datetime import datetime, timedelta

import click
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, Border, Side, PatternFill
from openpyxl.utils import get_column_letter

# Colors matching the original spreadsheet
YELLOW_FILL = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
GREEN_FILL = PatternFill(start_color="B7E1CD", end_color="B7E1CD", fill_type="solid")
GRAY_FILL = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")


def generate_week_dates(start_date: datetime, num_weeks: int) -> list:
    """Generate a list of week start dates."""
    return [start_date + timedelta(weeks=i) for i in range(num_weeks)]


def format_name(name: str, first_only: bool = True) -> str:
    """Convert email-style names to proper names.

    Examples (first_only=True):
        'john@company.com' -> 'John'
        'john.doe@company.com' -> 'John'
        'John Doe' -> 'John'
    Examples (first_only=False):
        'john.doe@company.com' -> 'John Doe'
    """
    if not name:
        return name

    # If it looks like an email, extract the name part
    if "@" in name:
        name = name.split("@")[0]
        # Convert dots to spaces and capitalize each word
        name = " ".join(word.capitalize() for word in name.split("."))

    if first_only:
        return name.split()[0] if name else name

    return name


def extract_unique_assignees(issues: list) -> list:
    """Extract unique assignee names from issues, formatted as proper names."""
    raw_names = {
        issue["assignee"]["name"]
        for issue in issues
        if issue.get("assignee") and issue["assignee"].get("name")
    }
    return sorted(format_name(name) for name in raw_names)


def create_excel(
    team_name: str,
    quarter: str,
    issues: list,
    output_file: str,
    start_date: datetime,
    num_weeks: int = 13,
):
    """Create the Excel planning spreadsheet."""
    wb = Workbook()
    ws = wb.active
    ws.title = datetime.now().strftime("%d%m%Y")

    # Styles
    title_font = Font(bold=True, size=14)
    header_font = Font(bold=True)
    thin_border = Border(
        left=Side(style="thin"),
        right=Side(style="thin"),
        top=Side(style="thin"),
        bottom=Side(style="thin"),
    )

    # Row 2: Title
    ws.merge_cells("B2:C2")
    ws["B2"] = f"{team_name} - {quarter} Planning"
    ws["B2"].font = title_font

    # Get unique assignees from issues
    assignees = extract_unique_assignees(issues)

    # Row 4: Capacity header and week dates
    ws["G4"].fill = YELLOW_FILL
    ws["H4"] = "Capacity"
    ws["H4"].font = header_font
    ws["H4"].fill = YELLOW_FILL

    week_dates = generate_week_dates(start_date, num_weeks)
    for i, date in enumerate(week_dates):
        cell = ws.cell(row=4, column=9 + i, value=date)
        cell.number_format = "M/D"
        cell.font = header_font
        cell.fill = YELLOW_FILL

    # Add "Capacity/week" column after week dates
    capacity_week_col = 9 + num_weeks
    cell = ws.cell(row=4, column=capacity_week_col, value="Capacity/week")
    cell.font = header_font
    cell.fill = YELLOW_FILL

    # Calculate header row position (after capacity section with spacing)
    header_row = 5 + len(assignees) + 4

    # Rows 5+: Engineer capacity rows (SUMIF formulas)
    data_start_row = header_row + 1
    estimated_last_row = data_start_row + len(issues)

    for idx, assignee_name in enumerate(assignees):
        row = 5 + idx
        ws.cell(row=row, column=8, value=assignee_name)

        for i in range(num_weeks):
            col = 9 + i
            col_letter = get_column_letter(col)
            formula = f'=SUMIF($H${data_start_row}:$H${estimated_last_row},$H{row},{col_letter}${data_start_row}:{col_letter}${estimated_last_row})'
            ws.cell(row=row, column=col, value=formula)

    # Header row (Dependency column removed)
    headers = [
        ("B", "Initiative"),
        ("C", "Projects"),
        ("D", "Issue"),
        ("E", "Estimate (days)"),
        ("F", "Description"),
        ("G", "Linear Ticket"),
        ("H", "Assigned to"),
    ]

    for col_letter, header_text in headers:
        cell = ws[f"{col_letter}{header_row}"]
        cell.value = header_text
        cell.font = header_font
        cell.border = thin_border
        cell.fill = YELLOW_FILL

    # Week date headers in header row
    for i in range(num_weeks):
        col = 9 + i
        cell = ws.cell(row=header_row, column=col)
        cell.value = f"={get_column_letter(col)}4"
        cell.font = header_font
        cell.border = thin_border
        cell.fill = YELLOW_FILL

    # Group issues by initiative and project
    grouped_issues = {}
    for issue in issues:
        project = issue.get("project") or {}
        project_name = project.get("name", "No Project")
        initiatives = project.get("initiatives", {}).get("nodes", [])
        initiative_name = initiatives[0].get("name") if initiatives else "No Initiative"

        key = (initiative_name, project_name)
        grouped_issues.setdefault(key, []).append(issue)

    # Write issues with gray separator rows between initiatives
    current_row = data_start_row
    last_initiative = None
    last_col = 8 + num_weeks  # Last column for gray fill

    for initiative_name, project_name in sorted(grouped_issues.keys()):
        # Add gray separator row when initiative changes (except for first)
        if last_initiative is not None and initiative_name != last_initiative:
            for col in range(2, last_col + 1):
                ws.cell(row=current_row, column=col).fill = GRAY_FILL
            current_row += 1

        last_initiative = initiative_name

        for issue in grouped_issues[(initiative_name, project_name)]:
            ws.cell(row=current_row, column=2, value=initiative_name)
            ws.cell(row=current_row, column=3, value=project_name)
            ws.cell(row=current_row, column=4, value=issue.get("title", ""))

            estimate = issue.get("estimate")
            if estimate is not None:
                cell = ws.cell(row=current_row, column=5, value=float(estimate))
                cell.fill = GREEN_FILL

            description = issue.get("description") or ""
            ws.cell(row=current_row, column=6, value=description[:500] if len(description) > 500 else description)
            cell = ws.cell(row=current_row, column=7, value=issue.get("url", ""))
            cell.alignment = Alignment(wrap_text=True)

            assignee = issue.get("assignee") or {}
            ws.cell(row=current_row, column=8, value=format_name(assignee.get("name", "")))

            current_row += 1

    # Update SUMIF formulas with actual row range
    actual_last_row = current_row - 1
    for idx in range(len(assignees)):
        row = 5 + idx
        for i in range(num_weeks):
            col = 9 + i
            col_letter = get_column_letter(col)
            formula = f'=SUMIF($H${data_start_row}:$H${actual_last_row},$H{row},{col_letter}${data_start_row}:{col_letter}${actual_last_row})'
            ws.cell(row=row, column=col, value=formula)

    # Column widths (G=Linear Ticket fixed width, H=Assigned to)
    widths = {"B": 30, "C": 35, "D": 50, "E": 15, "F": 50, "G": 40, "H": 15}
    for col_letter, width in widths.items():
        ws.column_dimensions[col_letter].width = width

    for i in range(num_weeks + 1):
        ws.column_dimensions[get_column_letter(9 + i)].width = 8

    wb.save(output_file)
    click.echo(f"Excel file saved to: {output_file}")
