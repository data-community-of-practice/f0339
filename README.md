# f0339
# Extract Grants with Publications

A Python script that reads a grants Excel file and a publications Excel file, filters the grants to only those that have at least one associated publication, and produces a clean `Grants.json` node file for the pipeline graph.

No API calls required — pure Excel parsing.

## Pipeline context

This script works alongside the author/publication pipeline rather than sitting in sequence after it. The publications Excel file is the same source used by [f0334](https://github.com/data-community-of-practice/f0334); f0339 uses it to identify which grants are in scope.

```
grants.xlsx  ─────────────────────────────┐
                                          ├──  f0339  →  Grants.json
publications.xlsx  (same file as f0334) ──┘

publications.xlsx  →  f0334  →  ...  →  f0338  →  Researchers.json
                                                    Researcher_Publication.json
                                        f0337  →  Organisations.json
                                                   Researcher_Organisation.json
```

Together `Grants.json`, `Researchers.json`, `Researcher_Publication.json`, `Organisations.json`, and `Researcher_Organisation.json` form the complete grant-researcher-publication graph.

## How it works

1. Reads the `Grant_ID` column from the publications file to build the set of grant IDs that have at least one associated publication.
2. Reads the grants file row by row, keeping only grants whose `Project Code` appears in that set.
3. For each matching grant, extracts the title, start year, end year, and participant list (primary investigator + other investigators).
4. If the same project code appears on multiple rows, participant lists are merged.
5. Writes one JSON record per unique grant to `Grants.json`.

## Output

One file: **`Grants.json`** — a JSON array of grant records.

### Grant record

```json
{
  "grant_id": "12345678",
  "title": "Understanding Neural Mechanisms of Anxiety",
  "start_year": 2018,
  "end_year": 2022,
  "participant_list": [
    "Jane Louise Doe",
    "John M. Smith",
    "Wei Zhang"
  ]
}
```

| Field | Description |
|-------|-------------|
| `grant_id` | The project code as read from the grants file. |
| `title` | Grant title from the `TITLE` column. Empty string if not present. |
| `start_year` | Four-digit year extracted from the `Start Date` column. `null` if not parseable. |
| `end_year` | Four-digit year extracted from the `End Date` column. `null` if not parseable. |
| `participant_list` | Ordered list of participant names. The primary investigator (`Preferred Full Name`) appears first, followed by names from `Other Investigators`. |

## Input file formats

### Grants file

Must contain a `Project Code` column. All other columns are optional but used when present:

| Column | Description |
|--------|-------------|
| `Project Code` | **Required.** Grant identifier — matched against `Grant_ID` in the publications file. |
| `TITLE` | Grant title. |
| `Start Date` | Grant start date. Accepts Excel date values or any string containing a four-digit year. |
| `End Date` | Grant end date. Same format as `Start Date`. |
| `Preferred Full Name` | Full name of the primary investigator. Added as the first participant. |
| `Preferred First Name` | First name (read but not used separately in output). |
| `Preferred Last Name` | Last name (read but not used separately in output). |
| `Other Investigators` | Comma-separated list of additional investigator names. |

### Publications file

Must contain a `Grant_ID` column. This is the same file used as input to f0334 — no preparation needed beyond what f0334 already requires.

| Column | Description |
|--------|-------------|
| `Grant_ID` | **Required.** Used to identify which grants have publications. All other columns are ignored. |

## Requirements

- Python 3.7+
- Library: `openpyxl`

```bash
pip install openpyxl
```

## Usage

Both input files are required positional arguments:

```bash
python f0339.py grants.xlsx publications.xlsx
```

Specify a custom output path:

```bash
python f0339.py grants.xlsx publications.xlsx --output path/to/Grants.json
```

By default the output is written to `Grants.json` in the same directory as the grants file.

### From Spyder or Jupyter

```python
!python "E:\your\folder\f0339.py" "E:\your\folder\grants.xlsx" "E:\your\folder\publications.xlsx"
```

## Console output

```
Grants:       /path/to/grants.xlsx
Publications: /path/to/publications.xlsx
Output:       /path/to/Grants.json

Grant IDs in publications file: 387

=======================================================
GRANT EXTRACTION SUMMARY
=======================================================
Grants with publications:  387
Total participants:        1243

  12345678: Understanding Neural Mechanisms of Anxiety
    Years: 2018 - 2022
    Participants: Jane Louise Doe, John M. Smith, Wei Zhang

  87654321: Depression in Adolescent Populations
    Years: 2019 - 2023
    Participants: Alex Brown, Sarah Connor

  ... and 382 more

Saved: /path/to/Grants.json
```

## Troubleshooting

| Problem | Solution |
|---------|----------|
| `ERROR: No 'Grant_ID' column in publications file` | The publications file must have a column named exactly `Grant_ID`. This is the same column required by f0334. |
| `ERROR: No 'Project Code' column in grants file` | The grants file must have a column named exactly `Project Code`. The error message lists what columns were found. |
| `start_year` or `end_year` is `null` | The date column either had no value or a format the script could not find a four-digit year in. Check the raw cell value in Excel. |
| Participants missing from a grant | Only `Preferred Full Name` and `Other Investigators` are used. Ensure those column names match exactly, including capitalisation. |
| Fewer grants than expected | Only grants whose `Project Code` appears at least once in the `Grant_ID` column of the publications file are included. Grants with no associated publications are filtered out by design. |

## Limitations

- **Column name sensitivity**: all column name matching is exact (case and spacing must match the values listed above).
- **Other Investigators parsing**: the `Other Investigators` column is split on commas. Names that contain commas (rare but possible) will be split incorrectly.
- **Year extraction only**: full dates are not preserved — only the four-digit year is extracted from `Start Date` and `End Date`.
- **First active sheet only**: the script reads the first (active) worksheet in each Excel file. Data on other sheets is ignored.

## License

This script is provided as-is for research and data management purposes.
