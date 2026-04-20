# f0339
BDBSF Organisation Geo Tagging
# Organisation Geotagger

A Python script that adds geographic coordinates and location data to classified organisations, and generates an interactive HTML map colour-coded by classification. Uses an LLM to intelligently parse location information from messy affiliation strings.

## The problem

Many organisations in the pipeline have no ROR ID and no structured location data. Their only location clue is buried inside raw affiliation strings like:

- `"Black Dog Institute  Randwick New South Wales Australia"`
- `"Monash Alfred Psychiatry Research Centre (MAPrc) The Alfred Hospital and Monash University Central Clinical School Melbourne Vic. Australia"`

These strings have no consistent delimiters — some use commas, some use double-spaces, some run everything together. A hardcoded list of cities and countries can't cover every case. Instead, this script uses Claude to parse them.

## Location strategy

The script tries four tiers in order, stopping when coordinates are found:

### Tier 1 — ROR API (most precise)
For organisations with a ROR ID, fetches latitude, longitude, city, and country directly from the ROR record's GeoNames data.

### Tier 2 — Nominatim by organisation name
Geocodes the full organisation name (plus any known city/country) using the OpenStreetMap Nominatim API.

### Tier 3 — LLM extraction from raw affiliations
Sends each raw affiliation string to Claude (`temperature=0`) with a prompt that encodes affiliation-parsing knowledge: scan from the end, country is usually last, handle state abbreviations (VIC → Victoria, NSW → New South Wales, MD → Maryland), recognise that double-spaces act as delimiters, and that "England" means United Kingdom. Claude responds with structured JSON: `{"city": "Randwick", "state": "New South Wales", "country": "Australia"}`.

All LLM results are cached so re-runs are free and deterministic.

### Tier 4 — Nominatim by LLM-extracted location
Geocodes whatever city/state/country the LLM extracted. Tries the org name + parsed location first, then just the location for at minimum a city-level pin.

**City and country are always populated from the best available source**, even when exact coordinates can't be determined.

## Requirements

- Python 3.7+
- Libraries: `openpyxl`, `requests`
- An Anthropic API key (for Tier 3 LLM extraction — optional but recommended)

```bash
pip install openpyxl requests
```

## Setup

Uses the same `config.ini` as the other scripts:

```ini
[crossref]
email = yourname@example.com
delay = 1
save_every = 50
max_retries = 3

[anthropic]
api_key = sk-ant-your-key-here
```

Without an API key, Tiers 1 and 2 still run but Tier 3 (LLM parsing) is skipped, which means organisations without ROR IDs or Nominatim matches will likely have no location data.

## Input file format

The script expects the output of `classify_orgs.py` with at minimum:

| Column | Required | Description |
|---|---|---|
| `Organisation_Name` | Yes | Canonical institution name |
| `ROR_ID` | No | Used for Tier 1 geocoding |
| `Country`, `City` | No | Used as context for Nominatim |
| `Raw_Variants` | No | Raw affiliation strings — parsed by LLM for location |
| `Classification` | No | Used for map colour coding |

## Usage

### From a terminal

```bash
python geotag_orgs.py classified_orgs.xlsx
```

### From Spyder

```python
!python "E:\your\folder\geotag_orgs.py" "E:\your\folder\classified_orgs.xlsx"
```

## Output

### 1. Excel file (`<input>_geotagged.xlsx`)

| Column | Description |
|---|---|
| `Latitude` | Geographic latitude |
| `Longitude` | Geographic longitude |
| `City` | City name — enriched from ROR, Nominatim, or LLM parsing |
| `Country` | Country name — enriched from best available source |
| `Geo_Source` | Where coordinates came from: `ROR`, `Nominatim`, `Nominatim (LLM-parsed)`, `Nominatim (city)`, or error |

Plus all columns from the classified organisations file.

### 2. Interactive HTML map (`<input>_geotagged_map.html`)

A standalone Leaflet.js map with colour-coded pins:

| Colour | Classification |
|---|---|
| Blue | Research |
| Purple | Health |
| Green | Government |
| Red | Industry |
| Orange | Unclassified |

Click any pin for details. A "Not mapped" panel lists organisations that have city/country but couldn't be geocoded.

## Cache files

| File | Purpose |
|---|---|
| `<input>_geo_cache.json` | Coordinates and location for each organisation |
| `<input>_llm_location_cache.json` | LLM-extracted city/state/country from raw affiliation strings |

Re-running skips cached results. Delete a cache file to force re-processing for that tier.

## Full pipeline

This is step 6:

```
1. python crossref_author_fetch.py <input>.xlsx
     → adds Crossref_Authors column

2. python extract_unique_authors.py <step1_output>.xlsx
     → one row per unique author with DOIs

3. python fetch_affiliations.py <step2_output>.xlsx
     → adds Affiliations column (Crossref + OpenAlex)

4. python extract_unique_orgs.py <step3_output>.xlsx
     → one row per unique institution with ROR ID

5. python classify_orgs.py <step4_output>.xlsx
     → adds Classification (Research/Health/Government/Industry)

6. python geotag_orgs.py <step5_output>.xlsx
     → adds coordinates + city/country + generates interactive map
```

## Troubleshooting

| Problem | Solution |
|---|---|
| LLM tier skipped | Add your Anthropic API key to `config.ini` under `[anthropic]`. |
| Orgs still missing location | The raw affiliation strings may contain no location info at all (e.g., just a department name). These are listed in the map's "Not mapped" panel. |
| `403` from Nominatim | Set a real email in config.ini. Nominatim blocks requests without a proper User-Agent. |
| Wrong coordinates | Check `Geo_Source`. `Nominatim (city)` means it geocoded the city, not the institution. ROR is most precise. |
| Old cache causing issues | Delete `_geo_cache.json` and/or `_llm_location_cache.json` and re-run. |

## Limitations

- **Nominatim rate limit**: Max 1 request/second on the public API. Large datasets may take time.
- **LLM parsing accuracy**: Claude handles the vast majority of affiliation formats correctly, but unusual or non-English affiliations may occasionally be misparsed. Results are cached so you can inspect `_llm_location_cache.json` to audit.
- **Offline map viewing**: The HTML map requires an internet connection to load map tiles. The marker data is embedded in the file.

## License

This script is provided as-is for research and data management purposes. Map data from OpenStreetMap (ODbL license).
