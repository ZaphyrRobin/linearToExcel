"""Excel spreadsheet generation for planning documents."""

from datetime import datetime, timedelta

import click
from openpyxl import Workbook
from openpyxl.styles import Font, Border, Side, PatternFill
from openpyxl.utils import get_column_letter

# Colors matching the original spreadsheet
YELLOW_FILL = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
GREEN_FILL = PatternFill(start_color="B7E1CD", end_color="B7E1CD", fill_type="solid")


def generate_week_dates(start_date: datetime, num_weeks: int) -> list:
    """Generate a list of week start dates."""
    return [start_date + timedelta(weeks=i) for i in range(num_weeks)]


def extract_unique_assignees(issues: list) -> list:
    """Extract unique assignee names from issues."""
    return sorted({
        issue["assignee"]["name"]
        for issue in issues
        if issue.get("assignee") and issue["assignee"].get("name")
    })


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
    ws["H4"].fill = YELLOW_FILL
    ws["I4"] = "Capacity"
    ws["I4"].font = header_font
    ws["I4"].fill = YELLOW_FILL

    week_dates = generate_week_dates(start_date, num_weeks)
    for i, date in enumerate(week_dates):
        cell = ws.cell(row=4, column=10 + i, value=date)
        cell.number_format = "M/D"
        cell.font = header_font
        cell.fill = YELLOW_FILL

    # Add "Capacity/week" column after week dates
    capacity_week_col = 10 + num_weeks
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
        ws.cell(row=row, column=9, value=assignee_name)

        for i in range(num_weeks):
            col = 10 + i
            col_letter = get_column_letter(col)
            formula = f'=SUMIF($I${data_start_row}:$I${estimated_last_row},$I{row},{col_letter}${data_start_row}:{col_letter}${estimated_last_row})'
            ws.cell(row=row, column=col, value=formula)

    # Header row
    headers = [
        ("B", "Initiative"),
        ("C", "Projects"),
        ("D", "Issue"),
        ("E", "Estimate (days)"),
        ("F", "Description"),
        ("G", "Linear Ticket"),
        ("H", "Dependency"),
        ("I", "Assigned to"),
    ]

    for col_letter, header_text in headers:
        cell = ws[f"{col_letter}{header_row}"]
        cell.value = header_text
        cell.font = header_font
        cell.border = thin_border
        cell.fill = YELLOW_FILL

    # Week date headers in header row
    for i in range(num_weeks):
        col = 10 + i
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

    # Write issues
    current_row = data_start_row
    for initiative_name, project_name in sorted(grouped_issues.keys()):
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
            ws.cell(row=current_row, column=7, value=issue.get("url", ""))
            ws.cell(row=current_row, column=8, value="")

            assignee = issue.get("assignee") or {}
            ws.cell(row=current_row, column=9, value=assignee.get("name", ""))

            current_row += 1

    # Update SUMIF formulas with actual row range
    actual_last_row = current_row - 1
    for idx in range(len(assignees)):
        row = 5 + idx
        for i in range(num_weeks):
            col = 10 + i
            col_letter = get_column_letter(col)
            formula = f'=SUMIF($I${data_start_row}:$I${actual_last_row},$I{row},{col_letter}${data_start_row}:{col_letter}${actual_last_row})'
            ws.cell(row=row, column=col, value=formula)

    # Column widths
    widths = {"B": 30, "C": 35, "D": 50, "E": 15, "F": 50, "G": 50, "H": 15, "I": 20}
    for col_letter, width in widths.items():
        ws.column_dimensions[col_letter].width = width

    for i in range(num_weeks + 1):
        ws.column_dimensions[get_column_letter(10 + i)].width = 8

    wb.save(output_file)
    click.echo(f"Excel file saved to: {output_file}")
