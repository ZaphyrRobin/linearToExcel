"""Excel spreadsheet generation for planning documents."""

from datetime import datetime, timedelta

import click
from openpyxl import Workbook, load_workbook
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


def extract_cycles(issues: list) -> list:
    """Extract unique cycles from issues, sorted by start date."""
    cycles = {}
    for issue in issues:
        cycle = issue.get("cycle")
        if cycle and cycle.get("id"):
            cycles[cycle["id"]] = cycle
    return sorted(cycles.values(), key=lambda c: c.get("startsAt", ""))


def get_week_index(cycle_start: str, start_date: datetime, num_weeks: int) -> int:
    """Calculate which week column index (0-based) a cycle falls into.

    Returns -1 if the cycle is outside the date range or invalid.
    """
    if not cycle_start:
        return -1
    try:
        cycle_date = datetime.fromisoformat(cycle_start.replace("Z", "+00:00")).replace(tzinfo=None)
        days_diff = (cycle_date - start_date).days
        week_index = days_diff // 7
        if 0 <= week_index < num_weeks:
            return week_index
        return -1
    except (ValueError, TypeError):
        return -1


def _populate_sheet(
    ws,
    team_name: str,
    quarter: str,
    issues: list,
    start_date: datetime,
    num_weeks: int = 13,
    cycle_id: str = None,
    cycle_start: str = None,
):
    """Populate a worksheet with planning data.

    If cycle_id is provided, only issues in that cycle will have estimates
    placed in the weekly columns. All issues are still shown.
    cycle_start is used to determine which week column to place the estimates.
    """

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
        cell.number_format = "m/d"
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

    # Week date headers in header row (use actual dates, not formulas)
    for i, date in enumerate(week_dates):
        col = 9 + i
        cell = ws.cell(row=header_row, column=col)
        cell.value = date
        cell.number_format = "m/d"
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

            # Only show estimate if no cycle filter, or if issue belongs to this cycle
            issue_cycle = issue.get("cycle") or {}
            issue_in_cycle = (cycle_id is None) or (issue_cycle.get("id") == cycle_id)

            estimate = issue.get("estimate")
            if estimate is not None:
                cell = ws.cell(row=current_row, column=5, value=float(estimate))
                cell.fill = GREEN_FILL

            description = issue.get("description") or ""
            ws.cell(row=current_row, column=6, value=description[:500] if len(description) > 500 else description)
            cell = ws.cell(row=current_row, column=7, value=issue.get("url", ""))
            cell.alignment = Alignment(wrap_text=True)

            assignee = issue.get("assignee") or {}
            assignee_name = format_name(assignee.get("name", ""))
            ws.cell(row=current_row, column=8, value=assignee_name)

            # Only fill weekly capacity for issues in the current cycle
            if issue_in_cycle and estimate is not None and assignee_name:
                # Use the passed cycle_start (from tab's cycle), not the issue's cycle
                # This ensures all estimates for this tab's cycle go to the correct week
                effective_cycle_start = cycle_start if cycle_start else issue_cycle.get("startsAt", "")
                week_idx = get_week_index(effective_cycle_start, start_date, num_weeks)
                if week_idx >= 0:
                    cell = ws.cell(row=current_row, column=9 + week_idx, value=float(estimate))
                    cell.fill = GREEN_FILL

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


def create_excel(
    team_name: str,
    quarter: str,
    issues: list,
    output_file: str,
    start_date: datetime,
    num_weeks: int = 13,
):
    """Create a new Excel planning spreadsheet."""
    wb = Workbook()
    ws = wb.active
    ws.title = start_date.strftime("%m-%d")

    _populate_sheet(ws, team_name, quarter, issues, start_date, num_weeks)

    wb.save(output_file)
    click.echo(f"Excel file saved to: {output_file}")


def overwrite_excel(
    team_name: str,
    quarter: str,
    issues: list,
    input_file: str,
    start_date: datetime,
    num_weeks: int = 13,
):
    """Overwrite the first sheet of an existing Excel file with latest data."""
    wb = load_workbook(input_file)
    old_sheet = wb.active
    old_title = old_sheet.title

    # Remove old sheet and create fresh one
    wb.remove(old_sheet)
    ws = wb.create_sheet(title=old_title, index=0)

    _populate_sheet(ws, team_name, quarter, issues, start_date, num_weeks)

    wb.save(input_file)
    click.echo(f"Excel file overwritten: {input_file}")


def append_to_excel(
    team_name: str,
    quarter: str,
    issues: list,
    input_file: str,
    start_date: datetime,
    num_weeks: int = 13,
):
    """Append a new tab to an existing Excel file with the cycle start date as tab name."""
    wb = load_workbook(input_file)

    # Use cycle start date as tab name
    tab_name = start_date.strftime("%m-%d")

    # Ensure unique tab name
    existing_names = [sheet.title for sheet in wb.worksheets]
    if tab_name in existing_names:
        counter = 1
        while f"{tab_name} ({counter})" in existing_names:
            counter += 1
        tab_name = f"{tab_name} ({counter})"

    ws = wb.create_sheet(title=tab_name)

    _populate_sheet(ws, team_name, quarter, issues, start_date, num_weeks)

    wb.save(input_file)
    click.echo(f"New tab '{tab_name}' added to: {input_file}")


def create_excel_by_cycles(
    team_name: str,
    quarter: str,
    issues: list,
    output_file: str,
    start_date: datetime,
    num_weeks: int = 13,
):
    """Create Excel with multiple tabs - one per cycle.

    Each tab shows ALL issues, but only issues in that cycle have
    their estimates placed in the weekly capacity columns.
    The estimate is placed in the week column matching the cycle's start date.
    """
    cycles = extract_cycles(issues)

    if not cycles:
        click.echo("No cycles found in issues. Creating single sheet.")
        create_excel(team_name, quarter, issues, output_file, start_date, num_weeks)
        return

    # Use the earliest cycle's start date as the spreadsheet start date
    # This ensures all cycles fall within the visible week columns
    first_cycle_start = cycles[0].get("startsAt", "")
    if first_cycle_start:
        effective_start = datetime.fromisoformat(first_cycle_start.replace("Z", "+00:00")).replace(tzinfo=None)
        # Adjust to Monday of that week
        effective_start -= timedelta(days=effective_start.weekday())
    else:
        effective_start = start_date

    wb = Workbook()
    # Remove default sheet
    wb.remove(wb.active)

    for cycle in cycles:
        cycle_id = cycle["id"]
        cycle_start = cycle.get("startsAt", "")

        # Parse cycle start date for tab name
        if cycle_start:
            cycle_date = datetime.fromisoformat(cycle_start.replace("Z", "+00:00"))
            tab_name = cycle_date.strftime("%m-%d")
        else:
            tab_name = cycle.get("name", f"Cycle {cycle.get('number', '?')}")

        # Ensure tab name is valid (max 31 chars, no invalid chars)
        tab_name = tab_name[:31].replace("/", "-").replace("\\", "-")

        ws = wb.create_sheet(title=tab_name)
        _populate_sheet(ws, team_name, quarter, issues, effective_start, num_weeks, cycle_id=cycle_id, cycle_start=cycle_start)
        click.echo(f"Created tab: {tab_name}")

    wb.save(output_file)
    click.echo(f"Excel file saved to: {output_file} ({len(cycles)} cycle tabs)")
