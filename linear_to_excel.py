#!/usr/bin/env python3
"""CLI entry point for Linear to Excel Planning Tool."""

import sys
from datetime import datetime, timedelta

import click

from src.linear_api import fetch_teams, fetch_all_initiatives, fetch_issues_for_team, get_team_by_key
from src.excel_generator import create_excel, append_to_excel, overwrite_excel, create_excel_by_cycles


@click.command()
@click.option("--team", "-t", default=None, help="Linear team key (e.g., 'APP1')")
@click.option("--quarter", "-q", default=None, help="Quarter label (e.g., 'Q4 2025')")
@click.option("--output", "-o", default=None, help="Output Excel filename")
@click.option("--start-date", "-s", default=None, help="Start date (YYYY-MM-DD)")
@click.option("--weeks", "-w", default=13, help="Number of weeks (default: 13)")
@click.option("--initiatives", "-i", default=None, help="Comma-separated initiative slugs")
@click.option("--list-teams", is_flag=True, help="List available teams")
@click.option("--list-initiatives", is_flag=True, help="List available initiatives")
@click.option("--input", "-f", "input_file", default=None, help="Existing xlsx file to overwrite with latest data")
@click.option("--append", "-a", default=None, help="Existing xlsx file to append a new tab to")
@click.option("--by-cycles", is_flag=True, help="Create separate tabs for each Linear cycle")
def main(team, quarter, output, start_date, weeks, initiatives, list_teams, list_initiatives, input_file, append, by_cycles):
    """Generate a quarterly planning Excel spreadsheet from Linear."""
    if list_teams:
        teams = fetch_teams()
        if not teams:
            click.echo("No teams found.")
            return
        click.echo("Available teams:")
        for t in teams:
            click.echo(f"  {t.get('key', 'N/A'):10} - {t.get('name', 'Unknown')}")
        return

    if list_initiatives:
        all_initiatives = fetch_all_initiatives()
        if not all_initiatives:
            click.echo("No initiatives found.")
            return
        click.echo("Available initiatives:")
        for init in all_initiatives:
            click.echo(f"  {init.get('slugId', 'N/A'):20} - {init.get('name', 'Unknown')}")
        return

    if not team:
        click.echo("Error: --team / -t is required.", err=True)
        sys.exit(1)

    team_data = get_team_by_key(team)
    if not team_data:
        click.echo(f"Error: Team '{team}' not found. Use --list-teams to see available teams.", err=True)
        sys.exit(1)

    team_id = team_data["id"]
    team_name = team_data["name"]
    click.echo(f"Fetching data for team: {team_name} ({team})")

    initiative_slugs = [s.strip() for s in initiatives.split(",")] if initiatives else None
    if initiative_slugs:
        click.echo(f"Filtering by initiatives: {initiative_slugs}")

    # Determine quarter and start date
    now = datetime.now()
    if not quarter:
        quarter = f"Q{(now.month - 1) // 3 + 1} {now.year}"

    if start_date:
        start = datetime.strptime(start_date, "%Y-%m-%d")
    else:
        quarter_start_month = ((now.month - 1) // 3) * 3 + 1
        start = datetime(now.year, quarter_start_month, 1)
        start -= timedelta(days=start.weekday())  # Adjust to Monday

    click.echo("Fetching issues from Linear...")
    issues = fetch_issues_for_team(team_id, initiative_slugs)
    click.echo(f"Found {len(issues)} issues")

    if not issues:
        click.echo("Warning: No issues found.", err=True)

    # Determine which mode to use
    if input_file:
        # Overwrite existing file with latest data
        click.echo(f"Overwriting existing file: {input_file}")
        overwrite_excel(team_name, quarter, issues, input_file, start, weeks)
    elif append:
        # Append new tab to existing file
        click.echo(f"Appending new tab to: {append}")
        append_to_excel(team_name, quarter, issues, append, start, weeks)
    elif by_cycles:
        # Create file with separate tabs per cycle
        if not output:
            output = f"{team_name} - {quarter} Planning.xlsx"
        click.echo("Generating Excel file with cycle tabs...")
        create_excel_by_cycles(team_name, quarter, issues, output, start, weeks)
    else:
        # Create new file
        if not output:
            output = f"{team_name} - {quarter} Planning.xlsx"
        click.echo("Generating Excel file...")
        create_excel(team_name, quarter, issues, output, start, weeks)

    click.echo("Done!")


if __name__ == "__main__":
    main()
