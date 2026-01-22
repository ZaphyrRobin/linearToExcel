"""Linear API client for fetching teams, initiatives, and issues."""

import os
import sys
from typing import Optional

import click
import requests
from dotenv import load_dotenv

load_dotenv()

LINEAR_API_URL = "https://api.linear.app/graphql"


def get_api_key() -> str:
    """Get Linear API key from environment."""
    api_key = os.getenv("LINEAR_API_KEY")
    if not api_key:
        click.echo("Error: LINEAR_API_KEY not found in environment or .env file", err=True)
        sys.exit(1)
    return api_key


def linear_request(query: str, variables: Optional[dict] = None) -> dict:
    """Make a GraphQL request to Linear API."""
    headers = {
        "Authorization": get_api_key(),
        "Content-Type": "application/json",
    }
    payload = {"query": query}
    if variables:
        payload["variables"] = variables

    response = requests.post(LINEAR_API_URL, json=payload, headers=headers)
    response.raise_for_status()

    result = response.json()
    if "errors" in result:
        click.echo(f"GraphQL errors: {result['errors']}", err=True)
        sys.exit(1)

    return result.get("data", {})


def fetch_teams() -> list:
    """Fetch all teams from Linear."""
    query = """
    query {
        teams {
            nodes { id, key, name }
        }
    }
    """
    return linear_request(query).get("teams", {}).get("nodes", [])


def fetch_all_initiatives(include_archived: bool = False) -> list:
    """Fetch all initiatives from Linear with their slugIds."""
    query = """
    query {
        initiatives(first: 100) {
            nodes {
                id
                name
                slugId
            }
        }
    }
    """
    initiatives = linear_request(query).get("initiatives", {}).get("nodes", [])

    if not include_archived:
        initiatives = [i for i in initiatives if "[Archive]" not in i.get("name", "")]

    return initiatives


def fetch_issues_for_team(team_id: str, initiative_slugs: Optional[list] = None, exclude_completed: bool = False) -> list:
    """Fetch all issues for a specific team with pagination, optionally filtered by initiatives.

    Args:
        team_id: The Linear team ID
        initiative_slugs: Optional list of initiative slugs to filter by
        exclude_completed: If True, exclude issues with completed status
    """
    all_issues = []
    end_cursor = None

    # Build initiative filter set if specified
    initiative_ids = None
    if initiative_slugs:
        initiatives = fetch_all_initiatives()
        initiative_ids = {i["id"] for i in initiatives if i.get("slugId") in initiative_slugs}
        if not initiative_ids:
            click.echo(f"Warning: No matching initiatives found for slugs: {initiative_slugs}", err=True)
            return []

    query = """
    query($teamId: ID!, $after: String) {
        issues(filter: { team: { id: { eq: $teamId } } }, first: 100, after: $after) {
            nodes {
                id
                identifier
                title
                description
                url
                estimate
                updatedAt
                assignee { name }
                state {
                    name
                    type
                }
                completedAt
                cycle {
                    id
                    name
                    number
                    startsAt
                    endsAt
                }
                project {
                    id
                    name
                    initiatives { nodes { id, name, slugId } }
                }
            }
            pageInfo { hasNextPage, endCursor }
        }
    }
    """

    has_next_page = True
    while has_next_page:
        variables = {"teamId": team_id, "after": end_cursor}
        data = linear_request(query, variables)
        issues_data = data.get("issues", {})
        issues = issues_data.get("nodes", [])

        # Filter out cancelled/archived issues (keep: triage, backlog, unstarted, started, completed)
        excluded_types = {"canceled", "cancelled"}
        issues = [
            issue for issue in issues
            if (issue.get("state") or {}).get("type", "").lower() not in excluded_types
        ]

        # Filter out completed issues if requested
        if exclude_completed:
            issues = [
                issue for issue in issues
                if (issue.get("state") or {}).get("type", "").lower() != "completed"
            ]

        # Filter by initiative if specified
        if initiative_ids:
            issues = [
                issue for issue in issues
                if any(
                    init.get("id") in initiative_ids
                    for init in (issue.get("project") or {}).get("initiatives", {}).get("nodes", [])
                )
            ]

        all_issues.extend(issues)

        page_info = issues_data.get("pageInfo", {})
        has_next_page = page_info.get("hasNextPage", False)
        end_cursor = page_info.get("endCursor")

    return all_issues


def get_team_by_key(team_key: str) -> Optional[dict]:
    """Find team by its key."""
    for team in fetch_teams():
        if team.get("key", "").upper() == team_key.upper():
            return team
    return None


def fetch_issue_by_identifier(identifier: str) -> Optional[dict]:
    """Fetch a single issue by its identifier (e.g., 'APP1-123')."""
    query = """
    query($term: String!) {
        searchIssues(term: $term, first: 1) {
            nodes {
                id
                identifier
                title
                url
                state { name }
                assignee { name }
            }
        }
    }
    """
    data = linear_request(query, {"term": identifier.upper()})
    issues = data.get("searchIssues", {}).get("nodes", [])

    # Verify we got the exact match
    for issue in issues:
        if issue.get("identifier", "").upper() == identifier.upper():
            return issue

    return None


def fetch_issue_history(issue_id: str) -> list:
    """Fetch the history of changes for an issue."""
    all_history = []
    end_cursor = None

    query = """
    query($issueId: String!, $after: String) {
        issue(id: $issueId) {
            history(first: 100, after: $after) {
                nodes {
                    id
                    createdAt
                    updatedDescription
                    fromTitle
                    toTitle
                    fromPriority
                    toPriority
                    fromEstimate
                    toEstimate
                    fromDueDate
                    toDueDate
                    trashed
                    autoArchived
                    autoClosed
                    addedLabelIds
                    removedLabelIds
                    actor {
                        id
                        name
                    }
                    botActor {
                        id
                        name
                    }
                    fromState {
                        id
                        name
                        type
                    }
                    toState {
                        id
                        name
                        type
                    }
                    fromAssignee {
                        id
                        name
                    }
                    toAssignee {
                        id
                        name
                    }
                    fromCycle {
                        id
                        name
                        number
                    }
                    toCycle {
                        id
                        name
                        number
                    }
                    fromProject {
                        id
                        name
                    }
                    toProject {
                        id
                        name
                    }
                    fromParent {
                        id
                        identifier
                        title
                    }
                    toParent {
                        id
                        identifier
                        title
                    }
                    fromTeam {
                        id
                        key
                        name
                    }
                    toTeam {
                        id
                        key
                        name
                    }
                    addedLabels {
                        id
                        name
                    }
                    removedLabels {
                        id
                        name
                    }
                }
                pageInfo {
                    hasNextPage
                    endCursor
                }
            }
        }
    }
    """

    has_next_page = True
    while has_next_page:
        variables = {"issueId": issue_id, "after": end_cursor}
        data = linear_request(query, variables)
        issue_data = data.get("issue", {})
        if not issue_data:
            break
        history_data = issue_data.get("history", {})
        history = history_data.get("nodes", [])
        all_history.extend(history)

        page_info = history_data.get("pageInfo", {})
        has_next_page = page_info.get("hasNextPage", False)
        end_cursor = page_info.get("endCursor")

    return all_history


def fetch_issues_history_bulk(issue_ids: list, max_workers: int = 10) -> dict:
    """Fetch history for multiple issues in parallel.

    Args:
        issue_ids: List of issue IDs to fetch history for
        max_workers: Maximum number of parallel requests (default: 10)

    Returns:
        Dict mapping issue_id -> list of history entries
    """
    from concurrent.futures import ThreadPoolExecutor, as_completed

    result = {}

    if not issue_ids:
        return result

    def fetch_single(issue_id: str) -> tuple:
        """Fetch history for a single issue and return (issue_id, history)."""
        return issue_id, fetch_issue_history(issue_id)

    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        futures = {executor.submit(fetch_single, issue_id): issue_id for issue_id in issue_ids}

        for future in as_completed(futures):
            issue_id, history = future.result()
            result[issue_id] = history

    return result
