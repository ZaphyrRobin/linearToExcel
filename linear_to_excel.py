#!/usr/bin/env python3
"""CLI entry point for Linear to Excel Planning Tool."""

import sys
from datetime import datetime, timedelta

import click

from src.linear_api import fetch_teams, fetch_all_initiatives, fetch_issues_for_team, get_team_by_key, fetch_issue_by_identifier, fetch_issue_history
from src.excel_generator import create_excel, refresh_excel


PRIORITY_LABELS = {0: "No priority", 1: "Urgent", 2: "High", 3: "Medium", 4: "Low"}


def format_history_entry(entry: dict) -> str:
    """Format a single history entry for display."""
    changes = []

    # Actor info
    actor_name = "System"
    if entry.get("actor"):
        actor_name = entry["actor"].get("name", "Unknown")
    elif entry.get("botActor"):
        actor_name = f"Bot: {entry['botActor'].get('name', 'Unknown')}"

    # Timestamp
    created_at = entry.get("createdAt", "")
    if created_at:
        # Parse and format the timestamp
        try:
            dt = datetime.fromisoformat(created_at.replace("Z", "+00:00"))
            created_at = dt.strftime("%Y-%m-%d %H:%M:%S")
        except (ValueError, AttributeError):
            pass

    # State change
    if entry.get("fromState") or entry.get("toState"):
        from_state = entry.get("fromState", {}).get("name", "None") if entry.get("fromState") else "None"
        to_state = entry.get("toState", {}).get("name", "None") if entry.get("toState") else "None"
        if from_state != to_state:
            changes.append(f"Status: {from_state} → {to_state}")

    # Assignee change
    if entry.get("fromAssignee") or entry.get("toAssignee"):
        from_assignee = entry.get("fromAssignee", {}).get("name", "Unassigned") if entry.get("fromAssignee") else "Unassigned"
        to_assignee = entry.get("toAssignee", {}).get("name", "Unassigned") if entry.get("toAssignee") else "Unassigned"
        if from_assignee != to_assignee:
            changes.append(f"Assignee: {from_assignee} → {to_assignee}")

    # Title change
    if entry.get("fromTitle") or entry.get("toTitle"):
        from_title = entry.get("fromTitle", "None") or "None"
        to_title = entry.get("toTitle", "None") or "None"
        if from_title != to_title:
            # Truncate long titles
            from_title = (from_title[:30] + "...") if len(from_title) > 33 else from_title
            to_title = (to_title[:30] + "...") if len(to_title) > 33 else to_title
            changes.append(f"Title: \"{from_title}\" → \"{to_title}\"")

    # Priority change
    if entry.get("fromPriority") is not None or entry.get("toPriority") is not None:
        from_priority = PRIORITY_LABELS.get(entry.get("fromPriority"), "None")
        to_priority = PRIORITY_LABELS.get(entry.get("toPriority"), "None")
        if entry.get("fromPriority") != entry.get("toPriority"):
            changes.append(f"Priority: {from_priority} → {to_priority}")

    # Estimate change
    if entry.get("fromEstimate") is not None or entry.get("toEstimate") is not None:
        from_est = entry.get("fromEstimate") if entry.get("fromEstimate") is not None else "None"
        to_est = entry.get("toEstimate") if entry.get("toEstimate") is not None else "None"
        if from_est != to_est:
            changes.append(f"Estimate: {from_est} → {to_est}")

    # Due date change
    if entry.get("fromDueDate") or entry.get("toDueDate"):
        from_due = entry.get("fromDueDate", "None") or "None"
        to_due = entry.get("toDueDate", "None") or "None"
        if from_due != to_due:
            changes.append(f"Due date: {from_due} → {to_due}")

    # Cycle change
    if entry.get("fromCycle") or entry.get("toCycle"):
        from_cycle = entry.get("fromCycle", {}).get("name", "None") if entry.get("fromCycle") else "None"
        to_cycle = entry.get("toCycle", {}).get("name", "None") if entry.get("toCycle") else "None"
        if from_cycle != to_cycle:
            changes.append(f"Cycle: {from_cycle} → {to_cycle}")

    # Project change
    if entry.get("fromProject") or entry.get("toProject"):
        from_project = entry.get("fromProject", {}).get("name", "None") if entry.get("fromProject") else "None"
        to_project = entry.get("toProject", {}).get("name", "None") if entry.get("toProject") else "None"
        if from_project != to_project:
            changes.append(f"Project: {from_project} → {to_project}")

    # Parent change
    if entry.get("fromParent") or entry.get("toParent"):
        from_parent = entry.get("fromParent", {}).get("identifier", "None") if entry.get("fromParent") else "None"
        to_parent = entry.get("toParent", {}).get("identifier", "None") if entry.get("toParent") else "None"
        if from_parent != to_parent:
            changes.append(f"Parent: {from_parent} → {to_parent}")

    # Team change
    if entry.get("fromTeam") or entry.get("toTeam"):
        from_team = entry.get("fromTeam", {}).get("key", "None") if entry.get("fromTeam") else "None"
        to_team = entry.get("toTeam", {}).get("key", "None") if entry.get("toTeam") else "None"
        if from_team != to_team:
            changes.append(f"Team: {from_team} → {to_team}")

    # Labels added/removed
    added_labels = entry.get("addedLabels") or []
    removed_labels = entry.get("removedLabels") or []
    if added_labels:
        label_names = ", ".join(l.get("name", "?") for l in added_labels)
        changes.append(f"Labels added: {label_names}")
    if removed_labels:
        label_names = ", ".join(l.get("name", "?") for l in removed_labels)
        changes.append(f"Labels removed: {label_names}")

    # Description updated
    if entry.get("updatedDescription"):
        changes.append("Description updated")

    # Trashed
    if entry.get("trashed") is True:
        changes.append("Issue trashed")
    elif entry.get("trashed") is False:
        changes.append("Issue restored from trash")

    # Auto-archived/closed
    if entry.get("autoArchived"):
        changes.append("Auto-archived")
    if entry.get("autoClosed"):
        changes.append("Auto-closed")

    if not changes:
        return ""

    return f"[{created_at}] {actor_name}:\n  " + "\n  ".join(changes)


def display_issue_history(identifier: str) -> None:
    """Fetch and display the history of a Linear issue."""
    click.echo(f"Fetching issue: {identifier.upper()}...")

    issue = fetch_issue_by_identifier(identifier)
    if not issue:
        click.echo(f"Error: Issue '{identifier}' not found.", err=True)
        sys.exit(1)

    click.echo(f"\nIssue: {issue['identifier']} - {issue['title']}")
    click.echo(f"Status: {issue.get('state', {}).get('name', 'Unknown')}")
    click.echo(f"Assignee: {issue.get('assignee', {}).get('name', 'Unassigned') if issue.get('assignee') else 'Unassigned'}")
    click.echo(f"URL: {issue.get('url', 'N/A')}")
    click.echo("\n" + "=" * 60)
    click.echo("HISTORY")
    click.echo("=" * 60 + "\n")

    history = fetch_issue_history(issue["id"])
    if not history:
        click.echo("No history found for this issue.")
        return

    # Sort by createdAt descending (most recent first)
    history.sort(key=lambda x: x.get("createdAt", ""), reverse=True)

    displayed_count = 0
    for entry in history:
        formatted = format_history_entry(entry)
        if formatted:
            click.echo(formatted)
            click.echo("")
            displayed_count += 1

    if displayed_count == 0:
        click.echo("No significant changes found in history.")
    else:
        click.echo(f"Total: {displayed_count} history entries")


@click.command()
@click.option("--team", "-t", default=None, help="Linear team key (e.g., 'APP1')")
@click.option("--output", "-o", default=None, help="Output Excel filename")
@click.option("--start-date", "-s", default=None, help="Start date (YYYY-MM-DD)")
@click.option("--end-date", "-e", default=None, help="End date (YYYY-MM-DD) - used with --file to set week range")
@click.option("--initiatives", "-i", default=None, help="Comma-separated initiative slugs")
@click.option("--list-teams", is_flag=True, help="List available teams")
@click.option("--list-initiatives", is_flag=True, help="List available initiatives")
@click.option("--file", "-f", "existing_file", default=None, help="Existing xlsx file to refresh with latest Linear data")
@click.option("--issue-history", "issue_id", default=None, help="Show history of a specific issue (e.g., 'APP1-123')")
@click.option("--exclude-completed", is_flag=True, help="Exclude issues that are already completed")
def main(team, output, start_date, end_date, initiatives, list_teams, list_initiatives, existing_file, issue_id, exclude_completed):
    """Generate a quarterly planning Excel spreadsheet from Linear."""
    if issue_id:
        display_issue_history(issue_id)
        return

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

    # Determine quarter label and start date
    now = datetime.now()
    quarter = f"Q{(now.month - 1) // 3 + 1} {now.year}"

    if start_date:
        start = datetime.strptime(start_date, "%Y-%m-%d")
    else:
        quarter_start_month = ((now.month - 1) // 3) * 3 + 1
        start = datetime(now.year, quarter_start_month, 1)
        start -= timedelta(days=start.weekday())  # Adjust to Monday

    click.echo("Fetching issues from Linear...")
    if exclude_completed:
        click.echo("Excluding completed issues...")
    issues = fetch_issues_for_team(team_id, initiative_slugs, exclude_completed=exclude_completed)
    click.echo(f"Found {len(issues)} issues")

    if not issues:
        click.echo("Warning: No issues found.", err=True)

    # Calculate num_weeks from end_date if provided (default: 13 weeks)
    num_weeks = 13
    if end_date:
        try:
            end = datetime.strptime(end_date, "%Y-%m-%d")
            days_diff = (end - start).days
            num_weeks = (days_diff // 7) + 1  # +1 to include the end week
            click.echo(f"Date range: {start.strftime('%Y-%m-%d')} to {end.strftime('%Y-%m-%d')} ({num_weeks} weeks)")
        except ValueError:
            click.echo("Error: Invalid end-date format. Use YYYY-MM-DD", err=True)
            sys.exit(1)

    # Determine which mode to use
    if existing_file:
        # Refresh existing file with latest Linear data
        click.echo(f"Refreshing existing file: {existing_file}")
        refresh_excel(team_name, quarter, issues, existing_file, start, num_weeks)
    else:
        # Create new file
        if not output:
            output = f"{team_name} - {quarter} Planning.xlsx"
        click.echo("Generating Excel file...")
        create_excel(team_name, quarter, issues, output, start, num_weeks)

    click.echo("Done!")


if __name__ == "__main__":
    main()
