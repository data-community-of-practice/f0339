#!/usr/bin/env python3
"""
Extract Grants with Publications
===================================
Reads a grants Excel file and a publications Excel file.
Produces a JSON of grants that have associated publications,
including grant number, participant list, start year, and end year.

Output: Grants.json

Usage:
  python extract_grants.py grants.xlsx publications.xlsx [--output Grants.json]
"""

import sys
import json
import argparse
from pathlib import Path

try:
    import openpyxl
except ImportError:
    print("ERROR: pip install openpyxl")
    sys.exit(1)

SCRIPT_DIR = Path(__file__).resolve().parent


def read_header(ws):
    """Read header row and return {column_name: index} mapping."""
    header = list(next(ws.iter_rows(max_row=1, values_only=True)))
    return {val: idx for idx, val in enumerate(header) if val is not None}


def main():
    parser = argparse.ArgumentParser(
        description="Extract grants that have associated publications"
    )
    parser.add_argument("grants_xlsx",
                        help="Grants Excel file (with Project Code, Start Date, etc.)")
    parser.add_argument("publications_xlsx",
                        help="Publications Excel file (with Grant_ID column)")
    parser.add_argument("--output", "-o", default=None,
                        help="Output JSON (default: Grants.json in same dir as grants file)")
    args = parser.parse_args()

    grants_path = Path(args.grants_xlsx).resolve()
    pubs_path = Path(args.publications_xlsx).resolve()

    if not grants_path.exists():
        print(f"ERROR: {grants_path} not found")
        sys.exit(1)
    if not pubs_path.exists():
        print(f"ERROR: {pubs_path} not found")
        sys.exit(1)

    output_path = Path(args.output).resolve() if args.output else grants_path.parent / "Grants.json"

    print(f"Grants:       {grants_path}")
    print(f"Publications: {pubs_path}")
    print(f"Output:       {output_path}")
    print()

    # Read Grant_IDs from publications file
    wb_pubs = openpyxl.load_workbook(pubs_path, read_only=True)
    ws_pubs = wb_pubs.active
    pub_headers = read_header(ws_pubs)

    grant_id_idx = pub_headers.get("Grant_ID")
    if grant_id_idx is None:
        print("ERROR: No 'Grant_ID' column in publications file")
        print(f"  Found: {list(pub_headers.keys())}")
        sys.exit(1)

    pub_grant_ids = set()
    for row in ws_pubs.iter_rows(min_row=2, values_only=True):
        if grant_id_idx < len(row) and row[grant_id_idx]:
            pub_grant_ids.add(str(row[grant_id_idx]).strip())
    wb_pubs.close()

    print(f"Grant IDs in publications file: {len(pub_grant_ids)}")

    # Read grants file
    wb_grants = openpyxl.load_workbook(grants_path, read_only=True)
    ws_grants = wb_grants.active
    grant_headers = read_header(ws_grants)

    # Map expected columns
    col_project_code = grant_headers.get("Project Code")
    col_title = grant_headers.get("TITLE")
    col_start = grant_headers.get("Start Date")
    col_end = grant_headers.get("End Date")
    col_first_name = grant_headers.get("Preferred First Name")
    col_full_name = grant_headers.get("Preferred Full Name")
    col_last_name = grant_headers.get("Preferred Last Name")
    col_other_investigators = grant_headers.get("Other Investigators")

    if col_project_code is None:
        print("ERROR: No 'Project Code' column in grants file")
        print(f"  Found: {list(grant_headers.keys())}")
        sys.exit(1)

    # Build grant records
    grants = {}  # project_code -> grant dict

    for row in ws_grants.iter_rows(min_row=2, values_only=True):
        if col_project_code >= len(row):
            continue

        project_code = row[col_project_code]
        if not project_code:
            continue
        project_code = str(project_code).strip()

        # Only include grants that have publications
        if project_code not in pub_grant_ids:
            continue

        # Extract fields safely
        def get(idx):
            if idx is not None and idx < len(row):
                val = row[idx]
                return str(val).strip() if val is not None else ""
            return ""

        title = get(col_title)

        # Extract years from dates
        start_year = None
        if col_start is not None and col_start < len(row) and row[col_start]:
            start_val = row[col_start]
            if hasattr(start_val, 'year'):
                start_year = start_val.year
            else:
                # Try parsing string
                import re
                year_match = re.search(r'(\d{4})', str(start_val))
                if year_match:
                    start_year = int(year_match.group(1))

        end_year = None
        if col_end is not None and col_end < len(row) and row[col_end]:
            end_val = row[col_end]
            if hasattr(end_val, 'year'):
                end_year = end_val.year
            else:
                import re
                year_match = re.search(r'(\d{4})', str(end_val))
                if year_match:
                    end_year = int(year_match.group(1))

        # Build participant list
        participants = []

        # Primary investigator
        pi_name = get(col_full_name)
        if pi_name:
            participants.append(pi_name)

        # Other investigators
        others = get(col_other_investigators)
        if others:
            for name in others.split(","):
                name = name.strip()
                if name and name not in participants:
                    participants.append(name)

        # Deduplicate grant (same project code might appear multiple times)
        if project_code not in grants:
            grants[project_code] = {
                "grant_id": project_code,
                "title": title,
                "start_year": start_year,
                "end_year": end_year,
                "participant_list": participants,
            }
        else:
            # Merge participants if same grant appears again
            existing = grants[project_code]
            for p in participants:
                if p not in existing["participant_list"]:
                    existing["participant_list"].append(p)

    wb_grants.close()

    grant_list = list(grants.values())

    # Save
    with open(output_path, "w", encoding="utf-8") as f:
        json.dump(grant_list, f, ensure_ascii=False, indent=2)

    # Summary
    total_participants = sum(len(g["participant_list"]) for g in grant_list)

    def safe(s):
        return s.encode("ascii", errors="replace").decode("ascii")

    print(f"\n{'='*55}")
    print(f"GRANT EXTRACTION SUMMARY")
    print(f"{'='*55}")
    print(f"Grants in file:            {len(pub_grant_ids) + len(set()) }")
    print(f"Grants with publications:  {len(grant_list)}")
    print(f"Total participants:        {total_participants}")

    for g in grant_list[:5]:
        print(f"\n  {g['grant_id']}: {safe(g['title'][:50])}")
        print(f"    Years: {g['start_year']} - {g['end_year']}")
        print(f"    Participants: {', '.join(g['participant_list'])}")

    if len(grant_list) > 5:
        print(f"\n  ... and {len(grant_list) - 5} more")

    print(f"\nSaved: {output_path}")


if __name__ == "__main__":
    main()