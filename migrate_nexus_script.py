import csv
import json
import os
import sys
import time
from datetime import datetime

import requests


# ============================================================
# NEXUS CONSULTING GROUP - monday.com DATA MIGRATION SCRIPT
# ============================================================
#
# CODING LANGUAGE:
# This script is written in Python.
#
# IMPORTANT NOTE:
# You will see GraphQL queries inside triple quotes.
# Those GraphQL queries are text instructions that Python sends to the monday.com API.
#
# You will also see JSON-like structures such as:
#   {"label": "Active"}
#
# In Python, these are called dictionaries.
# Before sending them to monday.com, the script converts them into JSON.
#
# PURPOSE OF THIS SCRIPT:
# This script migrates raw Smartsheet CSV data into monday.com.
#
# WHAT THIS SCRIPT DOES:
# 1. Reads the raw CSV export from Smartsheet
# 2. Clears existing rows from both monday.com boards
#    - Clears Nexus Deliverables first
#    - Then clears Nexus Engagements
# 3. Deduplicates engagement records
# 4. Normalizes inconsistent statuses
# 5. Creates engagement records in monday.com
# 6. Creates deliverable records in monday.com
# 7. Links each deliverable to its parent engagement
# 8. Prints progress in the terminal so you can demo the migration live
#
#
# HOW TO RUN:
#   export MONDAY_API_TOKEN="your_api_token_here"
#   python3 migrate_nexus_annotated.py
#
# ============================================================


# =========================
# CONFIGURATION SECTION
# =========================
#
# This section stores the settings the script needs before it can run.


# This is monday.com's GraphQL API endpoint.
MONDAY_API_URL = "https://api.monday.com/v2"


# The API token is loaded from your terminal environment variable.
# I did NOT hardcode the token here because that would expose credentials.
MONDAY_API_TOKEN = os.getenv("MONDAY_API_TOKEN")


# This tells monday.com which API version to use.
# If the terminal variable is not set, it defaults to 2025-10.
MONDAY_API_VERSION = os.getenv("MONDAY_API_VERSION", "2025-10")


# This is the raw source file exported from Smartsheet.
SOURCE_CSV = "nexus_smartsheet_export.csv"


# Monday.com board IDs.
# Nexus Engagements stores project-level information.
# Nexus Deliverables stores task/deliverable-level information.
ENGAGEMENTS_BOARD_ID = "18411674428"
DELIVERABLES_BOARD_ID = "18411677653"


# These are optional group IDs.
# They are set to None because the script can create items in the default group.
ENGAGEMENTS_GROUP_ID = None
DELIVERABLES_GROUP_ID = None


# Monday.com column IDs from the Nexus Engagements board.
# The left side is the friendly name used in this Python script.
# The right side is the real monday.com column ID.
ENGAGEMENT_COLUMNS = {
    "engagement_id": "text_mm32y44j",
    "client": "text_mm32f83a",
    "engagement_lead": "text_mm32gqne",
    "start_date": "date",
    "end_date": "date_mm32kasr",
    "budget": "numeric_mm32q776",
    "status": "color_mm32y5vm",
}


# Monday.com column IDs from the Nexus Deliverables board.
# The engagement_link column is the connect-board relationship column.
DELIVERABLE_COLUMNS = {
    "deliverable_id": "text_mm322s08",
    "assignee": "text_mm32nd0j",
    "due_date": "date4",
    "priority": "color_mm32xg2s",
    "status": "color_mm32aqx",
    "hours_estimated": "numeric_mm32335k",
    "engagement_link": "board_relation_mm32d8pz",
}


# =========================
# STATUS TRANSFORMATION SECTION
# =========================
#
# The discovery call explained that the old Smartsheet data had inconsistent
# statuses. These dictionaries clean up those messy values before loading them
# into monday.com.

# Engagement-level statuses (overall project lifecycle):
#
# The left side represents the original status values coming from the
# raw Smartsheet export.
#
# The right side represents the cleaned and standardized status values
# used in monday.com after normalization.
ENGAGEMENT_STATUS_MAP = {
    "Active": "Active",
    "In Progress": "Active",
    "Complete": "Complete",
    "Done": "Complete",
    "On Hold": "On Hold",
    "Not Started": "Not Started",
}


# Deliverable-level statuses (where a task is in the workflow):
#
# The left side represents the original status values coming from the
# raw Smartsheet export.
#
# The right side represents the cleaned and standardized workflow statuses
# used in monday.com after normalization.
DELIVERABLE_STATUS_MAP = {
    "To Do": "To Do",
    "Not Started": "To Do",
    "In Progress": "In Progress",
    "Working on it": "In Progress",
    "In Review": "In Review",
    "Done": "Done",
}


# =========================
# BASIC HELPER FUNCTIONS
# =========================
#
# These functions handle small repeatable jobs.
# Keeping this logic in functions makes the main workflow easier to read.


def require_env_vars():
    """
    Confirms that the monday.com API token exists.

    I keep the token outside the script as an environment variable so I do
    not expose credentials in the code.
    """
    if not MONDAY_API_TOKEN:
        print("ERROR: Missing MONDAY_API_TOKEN environment variable.")
        print("Run this first:")
        print('export MONDAY_API_TOKEN="your_api_token_here"')
        sys.exit(1)


def monday_request(query, variables=None):
    """
    Sends a GraphQL request to monday.com.

    This is the central API helper. Every time the script needs to read,
    create, or delete data in monday.com, it goes through this function.

    What happens here:
    1. Build the API headers
    2. Add the authorization token
    3. Send the GraphQL query or mutation
    4. Handle errors
    5. Retry if monday.com says the complexity budget was exhausted
    """
    headers = {
        "Authorization": MONDAY_API_TOKEN,
        "API-Version": MONDAY_API_VERSION,
        "Content-Type": "application/json",
    }

    payload = {
        "query": query,
        "variables": variables or {},
    }

    max_attempts = 5

    for attempt in range(max_attempts):
        response = requests.post(
            MONDAY_API_URL,
            headers=headers,
            json=payload,
            timeout=30,
        )

        try:
            data = response.json()
        except ValueError:
            raise Exception(f"Non-JSON response from monday.com: {response.text}")

        if "errors" in data:
            first_error = data["errors"][0]
            error_code = first_error.get("extensions", {}).get("code")

            # monday.com uses a complexity budget.
            # If the script sends too many heavy API operations too quickly,
            # monday.com asks the script to wait before trying again.
            if error_code == "COMPLEXITY_BUDGET_EXHAUSTED":
                retry_seconds = first_error.get("extensions", {}).get("retry_in_seconds", 20)
                print(f"Complexity limit hit. Waiting {retry_seconds} seconds before retrying...")
                time.sleep(retry_seconds + 2)
                continue

            raise Exception(json.dumps(data, indent=2))

        if response.status_code != 200:
            raise Exception(json.dumps(data, indent=2))

        return data["data"]

    raise Exception("monday.com API retry limit reached.")


def load_csv(path):
    """
    This reads the raw Smartsheet export and turns each row into a Python
    dictionary, so the script can access fields like engagement_id and
    deliverable_id.
    """
    with open(path, newline="", encoding="utf-8-sig") as file:
        return list(csv.DictReader(file))


def parse_mmddyyyy(date_value):
    """
    Converts dates from the CSV format into the format monday.com expects.

    Source CSV date format:
      MM/DD/YYYY

    monday.com date format:
      YYYY-MM-DD
    """
    if not date_value:
        return None

    return datetime.strptime(date_value, "%m/%d/%Y").strftime("%Y-%m-%d")


def to_number(value):
    """
    Converts budget and hours values into numeric values.

    This helps clean values like:
      "$150000"
      "150,000"
      "40"

    into numbers that monday.com can store in number columns.
    """
    if value in (None, ""):
        return None

    return float(str(value).replace(",", "").replace("$", ""))


def normalize_engagement_status(status):
    """
    Standardizes engagement statuses.

    If the script finds a status that is not in the map, it stops and raises
    an error. This prevents unknown or messy values from being loaded.
    """
    if status not in ENGAGEMENT_STATUS_MAP:
        raise ValueError(f"Unknown engagement status: {status}")

    return ENGAGEMENT_STATUS_MAP[status]


def normalize_deliverable_status(status):
    """
    Standardizes deliverable statuses.

    This keeps the deliverable board clean and prevents duplicate labels like
    "Working on it" and "In Progress" from both appearing in monday.com.
    """
    if status not in DELIVERABLE_STATUS_MAP:
        raise ValueError(f"Unknown deliverable status: {status}")

    return DELIVERABLE_STATUS_MAP[status]


# =========================
# BOARD CLEANUP FUNCTIONS
# =========================
#
# These functions clear old data from the boards before rerunning migration.
# This makes the live demo repeatable and prevents duplicate records.


def get_board_items(board_id):
    """
    Solve for duplication here:
    
    Before loading fresh data, the script checks what rows already exist
    on each board so it can delete them and avoid duplicates."

    This GraphQL query asks monday.com for:
    - item ID
    - item name

    Pagination is included in case the board grows later.
    """
    query = """
    query GetBoardItems($board_id: ID!, $cursor: String) {
      boards(ids: [$board_id]) {
        items_page(limit: 500, cursor: $cursor) {
          cursor
          items {
            id
            name
          }
        }
      }
    }
    """

    all_items = []
    cursor = None

    while True:
        variables = {
            "board_id": str(board_id),
            "cursor": cursor,
        }

        data = monday_request(query, variables)
        page = data["boards"][0]["items_page"]

        all_items.extend(page["items"])
        cursor = page["cursor"]

        if not cursor:
            break

    return all_items


def delete_item(item_id):
    """
    This function deletes an existing row by its monday.com item ID.
    """
    mutation = """
    mutation DeleteItem($item_id: ID!) {
      delete_item(item_id: $item_id) {
        id
      }
    }
    """

    monday_request(mutation, {"item_id": str(item_id)})


def clear_board(board_id, board_name):
    """
    Deletes all existing rows from one board.

    This allows me to rerun the migration from a clean state at anytime.
    I clear deliverables, then engagements.
    """
    print(f"\nClearing existing rows from {board_name}...")

    items = get_board_items(board_id)

    if not items:
        print(f"No existing rows found on {board_name}.")
        return

    for item in items:
        print(f"Deleting {board_name} item: {item['name']} | monday item ID: {item['id']}")
        delete_item(item["id"])

        # Small pause to avoid overwhelming the monday.com API.
        time.sleep(1)

    print(f"Deleted {len(items)} rows from {board_name}.")


# =========================
# ITEM CREATION FUNCTIONS
# =========================
#
# These functions create new monday.com rows.


def create_item(board_id, item_name, column_values, group_id=None):
    """
    This function creates a new row in a monday.com board and fills in
    the column values at the same time.

    In monday.com's GraphQL API, creating data is done with a mutation.
    """
    mutation = """
    mutation CreateItem(
      $board_id: ID!,
      $item_name: String!,
      $column_values: JSON!,
      $group_id: String
    ) {
      create_item(
        board_id: $board_id,
        item_name: $item_name,
        column_values: $column_values,
        group_id: $group_id
      ) {
        id
        name
      }
    }
    """

    variables = {
        "board_id": str(board_id),
        "item_name": item_name,
        "column_values": json.dumps(column_values),
        "group_id": group_id,
    }

    data = monday_request(mutation, variables)
    return data["create_item"]


def deduplicate_engagements(rows):
    """
    The source file repeats engagement information on every deliverable row.
    I use engagement_id to keep one copy of each engagement.

    Example:
    If ENG-001 appears on five deliverable rows, this function keeps one
    ENG-001 engagement record.
    """
    engagements = {}

    for row in rows:
        engagement_id = row["engagement_id"]

        if engagement_id not in engagements:
            engagements[engagement_id] = row

    return engagements


# =========================
# MIGRATION LOGIC
# =========================
#
# This is where the actual migration happens.


def migrate_engagements(engagements):
    """
    Creates engagement items first.

    """
    engagement_item_ids = {}

    print(f"\nCreating {len(engagements)} engagement records...")

    for engagement_id, row in engagements.items():
        # Build the monday.com column values for one engagement.
        column_values = {
            ENGAGEMENT_COLUMNS["engagement_id"]: row["engagement_id"],
            ENGAGEMENT_COLUMNS["client"]: row["client"],
            ENGAGEMENT_COLUMNS["engagement_lead"]: row["engagement_lead"],
            ENGAGEMENT_COLUMNS["start_date"]: {
                "date": parse_mmddyyyy(row["engagement_start"])
            },
            ENGAGEMENT_COLUMNS["end_date"]: {
                "date": parse_mmddyyyy(row["engagement_end"])
            },
            ENGAGEMENT_COLUMNS["budget"]: to_number(row["budget"]),
            ENGAGEMENT_COLUMNS["status"]: {
                "label": normalize_engagement_status(row["engagement_status"])
            },
        }

        # Create the engagement row in monday.com.
        item = create_item(
            board_id=ENGAGEMENTS_BOARD_ID,
            item_name=row["engagement_name"],
            column_values=column_values,
            group_id=ENGAGEMENTS_GROUP_ID,
        )

        # Store the mapping between source engagement_id and new monday item ID.
        # Example:
        #   ENG-001 -> 11923585409
        engagement_item_ids[engagement_id] = item["id"]

        print(
            f"Created engagement {engagement_id}: "
            f"{item['name']} | monday item ID: {item['id']}"
        )

        # Small pause to avoid hitting monday.com complexity limits.
        time.sleep(1)

    return engagement_item_ids


def migrate_deliverables(rows, engagement_item_ids):
    """
    Creates deliverable items and links them to engagements.

    """
    print(f"\nCreating {len(rows)} deliverable records...")

    created_count = 0

    for row in rows:
        source_engagement_id = row["engagement_id"]

        # Safety check:
        # A deliverable should never be created if its parent engagement
        # was not created first.
        if source_engagement_id not in engagement_item_ids:
            raise Exception(
                f"Cannot create deliverable {row['deliverable_id']} because "
                f"engagement {source_engagement_id} was not created."
            )

        linked_engagement_item_id = int(engagement_item_ids[source_engagement_id])

        # Build the monday.com column values for one deliverable.
        column_values = {
            DELIVERABLE_COLUMNS["deliverable_id"]: row["deliverable_id"],
            DELIVERABLE_COLUMNS["assignee"]: row["assignee"],
            DELIVERABLE_COLUMNS["due_date"]: {
                "date": parse_mmddyyyy(row["due_date"])
            },
            DELIVERABLE_COLUMNS["priority"]: {
                "label": row["priority"]
            },
            DELIVERABLE_COLUMNS["status"]: {
                "label": normalize_deliverable_status(row["deliverable_status"])
            },
            DELIVERABLE_COLUMNS["hours_estimated"]: to_number(row["hours_estimated"]),

            # Connect-board column:
            # This creates the relationship between the deliverable and
            # its parent engagement.
            DELIVERABLE_COLUMNS["engagement_link"]: {
                "item_ids": [linked_engagement_item_id]
            },
        }

        # Create the deliverable row in monday.com.
        item = create_item(
            board_id=DELIVERABLES_BOARD_ID,
            item_name=row["deliverable_name"],
            column_values=column_values,
            group_id=DELIVERABLES_GROUP_ID,
        )

        created_count += 1

        print(
            f"Created deliverable {row['deliverable_id']}: "
            f"{item['name']} | monday item ID: {item['id']} | "
            f"linked to engagement {source_engagement_id}"
        )

        # Small pause to avoid hitting monday.com complexity limits.
        time.sleep(1)

    return created_count


# =========================
# MAIN WORKFLOW
# =========================
#
# This is the main storyline of the script.

def main():
    """
    Runs the complete migration process from start to finish.

    """
    require_env_vars()

    print("Starting Nexus migration...")
    print("=" * 50)

    # Step 1: Read the source CSV.
    rows = load_csv(SOURCE_CSV)

    if not rows:
        print("No rows found in source CSV.")
        return

    print(f"Loaded {len(rows)} rows from {SOURCE_CSV}")

    # Step 2: Clear existing data from monday.com.
    #
    # Important:
    # Deliverables are cleared first because they are linked to engagements.
    clear_board(DELIVERABLES_BOARD_ID, "Nexus Deliverables")
    clear_board(ENGAGEMENTS_BOARD_ID, "Nexus Engagements")

    # Step 3: Deduplicate the flat source file into unique engagements.
    engagements = deduplicate_engagements(rows)

    print("\nSource data summary:")
    print(f"Unique engagements found: {len(engagements)}")
    print(f"Deliverables found: {len(rows)}")

    # Step 4: Create parent engagement records.
    engagement_item_ids = migrate_engagements(engagements)

    # Step 5: Create child deliverable records and link them to engagements.
    deliverable_count = migrate_deliverables(rows, engagement_item_ids)

    # Step 6: Print final summary.
    print("\nMigration complete.")
    print("=" * 50)
    print(f"Engagements created: {len(engagement_item_ids)}")
    print(f"Deliverables created: {deliverable_count}")
    print("=" * 50)


# This tells Python to run main() when this file is executed directly.
if __name__ == "__main__":
    main()
