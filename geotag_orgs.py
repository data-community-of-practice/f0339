"""
Organisation Geotagger (v4 — LLM-powered)
============================================
Adds latitude/longitude + city/country to classified organisations and
generates an interactive HTML map colour-coded by classification.

Location strategy (tries each tier in order):
  Tier 1 - ROR API: fetch coords + city/country from ROR record.
  Tier 2 - Nominatim (org name): geocode org name + known location.
  Tier 3 - LLM extraction: use Claude to parse city/state/country from
           raw affiliation strings (cached, temperature=0).
  Tier 4 - Nominatim (LLM-extracted location): geocode the parsed location.

City and country are ALWAYS populated from the best available source.

Setup:
  config.ini with [crossref] email and [anthropic] api_key.

Usage:
  python geotag_orgs.py <classified_orgs.xlsx> [output.xlsx] [config.ini]
"""

import sys
import re
import json
import time
import configparser
from pathlib import Path
from collections import Counter
import requests
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill

ROR_API = "https://api.ror.org/v2/organizations"
NOMINATIM_API = "https://nominatim.openstreetmap.org/search"
ANTHROPIC_API = "https://api.anthropic.com/v1/messages"
SCRIPT_DIR = Path(__file__).resolve().parent


# ── Config ──────────────────────────────────────────────────────────────

def load_config(config_path=None):
    if config_path is None:
        config_path = SCRIPT_DIR / "config.ini"
    else:
        config_path = Path(config_path).resolve()

    if not config_path.exists():
        print(f"ERROR: Config file not found at {config_path}")
        sys.exit(1)

    config = configparser.ConfigParser()
    config.read(config_path)
    email = config.get("crossref", "email", fallback=None)
    if not email or email.strip() == "your_email@example.com":
        print(f"ERROR: Please set your real email in {config_path}")
        sys.exit(1)

    api_key = config.get("anthropic", "api_key", fallback="")

    return {
        "email": email.strip(),
        "delay": config.getfloat("crossref", "delay", fallback=1),
        "save_every": config.getint("crossref", "save_every", fallback=50),
        "max_retries": config.getint("crossref", "max_retries", fallback=3),
        "api_key": api_key.strip(),
    }


# ── Cache ───────────────────────────────────────────────────────────────

def load_cache(cache_path):
    if cache_path.exists():
        with open(cache_path, "r", encoding="utf-8") as f:
            return json.load(f)
    return {}


def save_cache(cache, cache_path):
    with open(cache_path, "w", encoding="utf-8") as f:
        json.dump(cache, f, ensure_ascii=False, indent=2)


# ── ROR geocoding ──────────────────────────────────────────────────────

def geocode_by_ror(ror_id, session, max_retries=3):
    if "/" in ror_id:
        ror_path = ror_id.rstrip("/").split("/")[-1]
    else:
        ror_path = ror_id

    url = f"{ROR_API}/{ror_path}"

    for attempt in range(1, max_retries + 1):
        try:
            resp = session.get(url, timeout=30)
            if resp.status_code == 404:
                return {"lat": None, "lng": None, "city": "", "country": "", "source": "ROR ID not found"}
            if resp.status_code == 429:
                time.sleep(2 ** attempt * 5)
                continue
            resp.raise_for_status()
            data = resp.json()

            locations = data.get("locations", [])
            if locations:
                geo = locations[0].get("geonames_details", {})
                return {
                    "lat": float(geo["lat"]) if geo.get("lat") is not None else None,
                    "lng": float(geo["lng"]) if geo.get("lng") is not None else None,
                    "city": geo.get("name", ""),
                    "country": geo.get("country_name", ""),
                    "source": "ROR",
                }

            return {"lat": None, "lng": None, "city": "", "country": "", "source": "No location in ROR"}

        except requests.exceptions.RequestException as e:
            if attempt < max_retries:
                time.sleep(2 ** attempt)
            else:
                return {"lat": None, "lng": None, "city": "", "country": "", "source": f"ROR error: {e}"}

    return {"lat": None, "lng": None, "city": "", "country": "", "source": "ROR failed"}


# ── Nominatim geocoding ────────────────────────────────────────────────

def nominatim_search(query, session, max_retries=3):
    params = {"q": query, "format": "json", "limit": 1, "addressdetails": 1}

    for attempt in range(1, max_retries + 1):
        try:
            resp = session.get(NOMINATIM_API, params=params, timeout=30)
            if resp.status_code == 429:
                time.sleep(2 ** attempt * 5)
                continue
            resp.raise_for_status()
            results = resp.json()
            if results:
                r = results[0]
                addr = r.get("address", {})
                return {
                    "lat": float(r["lat"]),
                    "lng": float(r["lon"]),
                    "city": (addr.get("city") or addr.get("town") or addr.get("village")
                             or addr.get("municipality") or addr.get("suburb") or ""),
                    "country": addr.get("country", ""),
                    "source": "Nominatim",
                }
            return None
        except requests.exceptions.RequestException:
            if attempt < max_retries:
                time.sleep(2 ** attempt)
            else:
                return None
    return None


def geocode_by_nominatim(org_name, city, country, session, delay=1):
    queries = []
    if city and country:
        queries.append(f"{org_name}, {city}, {country}")
    if country:
        queries.append(f"{org_name}, {country}")
    queries.append(org_name)

    for query in queries:
        result = nominatim_search(query, session)
        if result:
            return result
        time.sleep(delay)

    return None


def geocode_location_only(city, state, country, session, delay=1):
    queries = []
    if city and state and country:
        queries.append(f"{city}, {state}, {country}")
    if city and country:
        queries.append(f"{city}, {country}")
    if state and country:
        queries.append(f"{state}, {country}")
    if country:
        queries.append(country)

    for query in queries:
        result = nominatim_search(query, session)
        if result:
            return result
        time.sleep(delay)

    return None


# ── LLM location extraction ────────────────────────────────────────────

LLM_LOCATION_PROMPT = """You are extracting location information from raw academic affiliation strings.

Given an affiliation string (which may contain an organisation name, department, faculty, city, state/province, and country mixed together), extract:
- city: the city or suburb name
- state: the state, province, or region (use the full name, not abbreviation)
- country: the full country name

Rules:
- Affiliations usually end with location, reading right to left: country, then state, then city.
- "Vic", "Vic.", "VIC" = Victoria, Australia
- "NSW" = New South Wales, Australia
- "QLD" = Queensland, Australia
- "SA" = South Australia, Australia (in Australian context)
- "WA" = Western Australia, Australia (in Australian context)
- "ACT" = Australian Capital Territory, Australia
- "England", "Scotland", "Wales" = United Kingdom
- "MD", "CA", "NY", "MA", etc. in US context = US state abbreviations
- If text has no location info at all, return empty values.
- Do NOT guess. Only extract what is clearly present in the text.
- Affiliations separated by double spaces use spaces as delimiters.

Respond with ONLY a JSON object on a single line, no markdown, no backticks:
{"city": "...", "state": "...", "country": "..."}

Use empty strings for missing values. Examples:

Input: "Black Dog Institute  Randwick New South Wales Australia"
{"city": "Randwick", "state": "New South Wales", "country": "Australia"}

Input: "Department of Psychiatry, St Vincent's Hospital, Melbourne, VIC, Australia"
{"city": "Melbourne", "state": "Victoria", "country": "Australia"}

Input: "Harvard Medical School, Boston, MA, United States"
{"city": "Boston", "state": "Massachusetts", "country": "United States"}

Input: "Institute of Psychiatry, King's College London, London, England"
{"city": "London", "state": "", "country": "United Kingdom"}

Input: "Centre for Genomics and Personalised Health"
{"city": "", "state": "", "country": ""}"""


def extract_location_llm(texts, api_key, session, cache, max_retries=3):
    """
    Use Claude to extract city/state/country from affiliation texts.
    Tries each text, returns first result with a country match.
    All results are cached.
    """
    best = {"city": "", "state": "", "country": ""}

    for text in texts:
        if not text:
            continue

        cache_key = text.lower().strip()[:200]
        if cache_key in cache:
            cached = cache[cache_key]
            score = bool(cached.get("country")) + bool(cached.get("state")) + bool(cached.get("city"))
            best_score = bool(best["country"]) + bool(best["state"]) + bool(best["city"])
            if score > best_score:
                best = cached
            if score >= 2:
                return best
            continue

        if not api_key:
            cache[cache_key] = {"city": "", "state": "", "country": ""}
            continue

        payload = {
            "model": "claude-sonnet-4-20250514",
            "max_tokens": 100,
            "temperature": 0,
            "system": LLM_LOCATION_PROMPT,
            "messages": [{"role": "user", "content": text}],
        }

        headers = {
            "Content-Type": "application/json",
            "x-api-key": api_key,
            "anthropic-version": "2023-06-01",
        }

        result = {"city": "", "state": "", "country": ""}

        for attempt in range(1, max_retries + 1):
            try:
                resp = session.post(ANTHROPIC_API, json=payload, headers=headers, timeout=60)
                if resp.status_code == 429:
                    time.sleep(min(2 ** attempt * 5, 60))
                    continue
                resp.raise_for_status()
                data = resp.json()

                response_text = ""
                for block in data.get("content", []):
                    if block.get("type") == "text":
                        response_text += block.get("text", "")

                response_text = response_text.strip()
                response_text = response_text.replace("```json", "").replace("```", "").strip()

                parsed = json.loads(response_text)
                result = {
                    "city": parsed.get("city", "").strip(),
                    "state": parsed.get("state", "").strip(),
                    "country": parsed.get("country", "").strip(),
                }
                break

            except (requests.exceptions.RequestException, json.JSONDecodeError, KeyError) as e:
                if attempt < max_retries:
                    time.sleep(2 ** attempt)
                else:
                    break

        cache[cache_key] = result

        score = bool(result["country"]) + bool(result["state"]) + bool(result["city"])
        best_score = bool(best["country"]) + bool(best["state"]) + bool(best["city"])
        if score > best_score:
            best = result
        if score >= 2:
            return best

        time.sleep(0.5)

    return best


# ── HTML Map ────────────────────────────────────────────────────────────

def generate_map_html(orgs, output_path):
    colours = {
        "Research": "#3498db", "Health": "#9b59b6",
        "Government": "#27ae60", "Industry": "#e74c3c",
        "Unclassified": "#f39c12",
    }

    markers = []
    for org in orgs:
        if org.get("lat") is not None and org.get("lng") is not None:
            cls = org.get("classification", "Unclassified")
            colour = colours.get(cls, "#999999")
            name = org["name"].replace("'", "\\'").replace('"', '\\"')
            city = org.get("city", "").replace("'", "\\'")
            country = org.get("country", "").replace("'", "\\'")
            markers.append({
                "lat": org["lat"], "lng": org["lng"],
                "name": name, "cls": cls, "colour": colour,
                "city": city, "country": country,
                "authors": str(org.get("author_count", 0)),
                "dois": str(org.get("doi_count", 0)),
            })

    markers_json = json.dumps(markers, ensure_ascii=False)

    cls_counts = Counter(m["cls"] for m in markers)
    legend_items = ""
    for cls in ["Research", "Health", "Government", "Industry", "Unclassified"]:
        count = cls_counts.get(cls, 0)
        if count > 0:
            colour = colours[cls]
            legend_items += f'<div style="margin:4px 0;"><span style="display:inline-block;width:14px;height:14px;border-radius:50%;background:{colour};margin-right:8px;vertical-align:middle;"></span>{cls} ({count})</div>\n'

    no_coords = [o for o in orgs if o.get("lat") is None]
    no_coords_section = ""
    if no_coords:
        items = ""
        for o in no_coords:
            cls = o.get("classification", "Unclassified")
            colour = colours.get(cls, "#999999")
            loc = ", ".join(filter(None, [o.get("city", ""), o.get("country", "")]))
            items += f'<div style="margin:3px 0;font-size:12px;"><span style="display:inline-block;width:10px;height:10px;border-radius:50%;background:{colour};margin-right:6px;vertical-align:middle;"></span>{o["name"]}<br><span style="color:#888;margin-left:16px;">{loc if loc else "No location data"}</span></div>\n'
        no_coords_section = f"""
<div class="no-coords">
  <h4>Not mapped ({len(no_coords)})</h4>
  <div style="max-height:200px;overflow-y:auto;">{items}</div>
</div>"""

    html_content = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Organisation Map</title>
<link rel="stylesheet" href="https://unpkg.com/leaflet@1.9.4/dist/leaflet.css" />
<script src="https://unpkg.com/leaflet@1.9.4/dist/leaflet.js"></script>
<style>
  body {{ margin: 0; padding: 0; font-family: Arial, sans-serif; }}
  #map {{ width: 100%; height: 100vh; }}
  .legend {{ position: absolute; bottom: 30px; left: 10px; z-index: 1000; background: white; padding: 12px 16px; border-radius: 8px; box-shadow: 0 2px 8px rgba(0,0,0,0.2); font-size: 13px; }}
  .legend h4 {{ margin: 0 0 8px 0; font-size: 14px; }}
  .info-box {{ position: absolute; top: 10px; right: 10px; z-index: 1000; background: white; padding: 10px 14px; border-radius: 8px; box-shadow: 0 2px 8px rgba(0,0,0,0.2); font-size: 13px; }}
  .no-coords {{ position: absolute; bottom: 30px; right: 10px; z-index: 1000; background: white; padding: 12px 16px; border-radius: 8px; box-shadow: 0 2px 8px rgba(0,0,0,0.2); font-size: 13px; max-width: 350px; }}
  .no-coords h4 {{ margin: 0 0 8px 0; font-size: 14px; color: #888; }}
</style>
</head>
<body>
<div id="map"></div>
<div class="legend"><h4>Classification</h4>{legend_items}</div>
<div class="info-box">{len(markers)} of {len(orgs)} organisations mapped</div>
{no_coords_section}
<script>
var map = L.map('map').setView([0, 0], 2);
L.tileLayer('https://{{s}}.tile.openstreetmap.org/{{z}}/{{x}}/{{y}}.png', {{
  attribution: '&copy; OpenStreetMap contributors', maxZoom: 18,
}}).addTo(map);
var markers = {markers_json};
var markerGroup = L.featureGroup();
markers.forEach(function(m) {{
  var circle = L.circleMarker([m.lat, m.lng], {{
    radius: 7, fillColor: m.colour, color: '#fff',
    weight: 1.5, opacity: 1, fillOpacity: 0.85,
  }});
  circle.bindPopup(
    '<b>' + m.name + '</b><br>' +
    '<span style="color:' + m.colour + ';font-weight:bold;">' + m.cls + '</span><br>' +
    (m.city ? m.city + ', ' : '') + m.country + '<br>' +
    'Authors: ' + m.authors + ' | DOIs: ' + m.dois
  );
  circle.addTo(markerGroup);
}});
markerGroup.addTo(map);
if (markers.length > 0) {{ map.fitBounds(markerGroup.getBounds().pad(0.1)); }}
</script>
</body>
</html>"""

    map_path = output_path.parent / f"{output_path.stem}_map.html"
    with open(map_path, "w", encoding="utf-8") as f:
        f.write(html_content)
    return map_path


# ── Main ────────────────────────────────────────────────────────────────

def main(input_file, output_file=None, config_path=None):
    cfg = load_config(config_path)
    has_api_key = bool(cfg["api_key"])
    print(f"Config: email={cfg['email']}, LLM: {'enabled' if has_api_key else 'disabled'}")

    input_path = Path(input_file)
    if not input_path.is_absolute():
        if not input_path.exists():
            fallback = SCRIPT_DIR / input_path
            if fallback.exists():
                input_path = fallback
    input_path = input_path.resolve()

    if not input_path.exists():
        print(f"ERROR: Input file not found: {input_path}")
        sys.exit(1)

    if output_file is None:
        output_path = input_path.parent / f"{input_path.stem}_geotagged.xlsx"
    else:
        output_path = Path(output_file).resolve()

    geo_cache_path = output_path.parent / f"{input_path.stem}_geo_cache.json"
    llm_loc_cache_path = output_path.parent / f"{input_path.stem}_llm_location_cache.json"

    print(f"Input:  {input_path}")
    print(f"Output: {output_path}")

    geo_cache = load_cache(geo_cache_path)
    llm_loc_cache = load_cache(llm_loc_cache_path)
    if geo_cache:
        print(f"  Loaded geo cache: {len(geo_cache)} entries")
    if llm_loc_cache:
        print(f"  Loaded LLM location cache: {len(llm_loc_cache)} entries")

    # --- Read input ---
    wb = openpyxl.load_workbook(input_path)
    ws = wb.active

    headers = {cell.value: cell.column for cell in ws[1]}
    name_col = headers.get("Organisation_Name")
    cls_col = headers.get("Classification")
    cls_src_col = headers.get("Classification_Source")
    just_col = headers.get("Justification")
    ror_id_col = headers.get("ROR_ID")
    country_col = headers.get("Country")
    city_col = headers.get("City")
    type_col = headers.get("ROR_Type")
    variants_col = headers.get("Raw_Variants")
    author_count_col = headers.get("Author_Count")
    authors_col = headers.get("Authors")
    doi_count_col = headers.get("DOI_Count")
    dois_col = headers.get("DOIs")

    if name_col is None:
        print("ERROR: Could not find 'Organisation_Name' column.")
        sys.exit(1)

    orgs = []
    for row in range(2, ws.max_row + 1):
        org = {
            "name": str(ws.cell(row=row, column=name_col).value or "").strip(),
            "classification": str(ws.cell(row=row, column=cls_col).value or "").strip() if cls_col else "",
            "cls_source": str(ws.cell(row=row, column=cls_src_col).value or "").strip() if cls_src_col else "",
            "justification": str(ws.cell(row=row, column=just_col).value or "").strip() if just_col else "",
            "ror_id": str(ws.cell(row=row, column=ror_id_col).value or "").strip() if ror_id_col else "",
            "country": str(ws.cell(row=row, column=country_col).value or "").strip() if country_col else "",
            "city": str(ws.cell(row=row, column=city_col).value or "").strip() if city_col else "",
            "ror_type": str(ws.cell(row=row, column=type_col).value or "").strip() if type_col else "",
            "variants": str(ws.cell(row=row, column=variants_col).value or "").strip() if variants_col else "",
            "author_count": ws.cell(row=row, column=author_count_col).value if author_count_col else 0,
            "authors": str(ws.cell(row=row, column=authors_col).value or "").strip() if authors_col else "",
            "doi_count": ws.cell(row=row, column=doi_count_col).value if doi_count_col else 0,
            "dois": str(ws.cell(row=row, column=dois_col).value or "").strip() if dois_col else "",
        }
        if org["name"]:
            orgs.append(org)

    print(f"\nFound {len(orgs)} organisations to geotag.\n")

    # --- Geocode ---
    session = requests.Session()
    session.headers.update({
        "User-Agent": f"OrgGeotagger/1.0 ({cfg['email']})",
        "Accept": "application/json",
    })

    stats = {"ror": 0, "nom_org": 0, "llm_nom": 0, "nom_city": 0, "cached": 0, "no_coords": 0}

    try:
        for i, org in enumerate(orgs):
            cache_key = org["name"].lower().strip()
            display = org["name"][:60].encode("ascii", errors="replace").decode("ascii")

            # Check cache
            if cache_key in geo_cache:
                cached = geo_cache[cache_key]
                org["lat"] = cached.get("lat")
                org["lng"] = cached.get("lng")
                org["geo_source"] = cached.get("source", "Cached")
                if not org["city"] and cached.get("city"):
                    org["city"] = cached["city"]
                if not org["country"] and cached.get("country"):
                    org["country"] = cached["country"]
                stats["cached"] += 1
                continue

            print(f"[{i+1}/{len(orgs)}] {display}")

            result = None

            # ── Tier 1: ROR ──
            if org["ror_id"]:
                result = geocode_by_ror(org["ror_id"], session, cfg["max_retries"])
                if result.get("city"):
                    org["city"] = result["city"]
                if result.get("country"):
                    org["country"] = result["country"]
                if result["lat"] is not None:
                    stats["ror"] += 1
                time.sleep(cfg["delay"])

            # ── Tier 2: Nominatim by org name ──
            if result is None or result["lat"] is None:
                nom_result = geocode_by_nominatim(
                    org["name"], org["city"], org["country"], session, cfg["delay"]
                )
                if nom_result:
                    result = nom_result
                    stats["nom_org"] += 1
                    if nom_result.get("city") and not org["city"]:
                        org["city"] = nom_result["city"]
                    if nom_result.get("country") and not org["country"]:
                        org["country"] = nom_result["country"]

            # ── Tier 3: LLM extraction from raw affiliations ──
            if (result is None or result["lat"] is None) or (not org["city"] and not org["country"]):
                # Build list of texts to try: raw variants, then org name
                texts_to_parse = []
                if org["variants"]:
                    for v in org["variants"].split(';'):
                        v = v.strip()
                        if v:
                            texts_to_parse.append(v)
                texts_to_parse.append(org["name"])

                llm_loc = extract_location_llm(
                    texts_to_parse, cfg["api_key"], session, llm_loc_cache, cfg["max_retries"]
                )

                if llm_loc.get("country") and not org["country"]:
                    org["country"] = llm_loc["country"]
                if llm_loc.get("city") and not org["city"]:
                    org["city"] = llm_loc["city"]

                # ── Tier 4: Nominatim by LLM-extracted location ──
                if result is None or result["lat"] is None:
                    search_city = org["city"] or llm_loc.get("city", "")
                    search_state = llm_loc.get("state", "")
                    search_country = org["country"] or llm_loc.get("country", "")

                    if search_city or search_country:
                        # Try org name + parsed location first
                        if search_city and search_country:
                            nom_result = nominatim_search(
                                f"{org['name']}, {search_city}, {search_country}", session
                            )
                            if nom_result:
                                result = nom_result
                                result["source"] = "Nominatim (LLM-parsed)"
                                stats["llm_nom"] += 1
                            time.sleep(cfg["delay"])

                        # Then just the location
                        if result is None or result["lat"] is None:
                            loc_result = geocode_location_only(
                                search_city, search_state, search_country,
                                session, cfg["delay"]
                            )
                            if loc_result:
                                result = loc_result
                                result["source"] = "Nominatim (city)"
                                stats["nom_city"] += 1
                                if loc_result.get("city") and not org["city"]:
                                    org["city"] = loc_result["city"]
                                if loc_result.get("country") and not org["country"]:
                                    org["country"] = loc_result["country"]

            # ── Final assignment ──
            if result and result["lat"] is not None:
                org["lat"] = result["lat"]
                org["lng"] = result["lng"]
                org["geo_source"] = result["source"]
            else:
                org["lat"] = None
                org["lng"] = None
                org["geo_source"] = "No coordinates found"
                stats["no_coords"] += 1

            geo_cache[cache_key] = {
                "lat": org["lat"], "lng": org["lng"],
                "city": org["city"], "country": org["country"],
                "source": org["geo_source"],
            }

            if (i + 1) % cfg["save_every"] == 0:
                save_cache(geo_cache, geo_cache_path)
                save_cache(llm_loc_cache, llm_loc_cache_path)
                print(f"  >> Caches saved ({i+1}/{len(orgs)})")

    except KeyboardInterrupt:
        print(f"\n>> Interrupted! Saving caches...")
        save_cache(geo_cache, geo_cache_path)
        save_cache(llm_loc_cache, llm_loc_cache_path)
        print("Re-run to continue.")
        sys.exit(0)

    save_cache(geo_cache, geo_cache_path)
    save_cache(llm_loc_cache, llm_loc_cache_path)

    # --- Stats ---
    has_coords = sum(1 for o in orgs if o.get("lat") is not None)
    has_city = sum(1 for o in orgs if o.get("city"))
    has_country = sum(1 for o in orgs if o.get("country"))
    no_location = sum(1 for o in orgs if not o.get("city") and not o.get("country") and o.get("lat") is None)

    print(f"\n--- Summary ---")
    print(f"  Coordinates from:")
    print(f"    ROR:                     {stats['ror']}")
    print(f"    Nominatim (org name):    {stats['nom_org']}")
    print(f"    Nominatim (LLM-parsed):  {stats['llm_nom']}")
    print(f"    Nominatim (city only):   {stats['nom_city']}")
    print(f"    Cached:                  {stats['cached']}")
    print(f"  Results:")
    print(f"    With coordinates:        {has_coords}/{len(orgs)}")
    print(f"    With city:               {has_city}/{len(orgs)}")
    print(f"    With country:            {has_country}/{len(orgs)}")
    print(f"    No location at all:      {no_location}/{len(orgs)}")

    # --- Write Excel ---
    out_wb = openpyxl.Workbook()
    out_ws = out_wb.active
    out_ws.title = "Geotagged Organisations"

    header_font = Font(name="Arial", bold=True, color="FFFFFF")
    header_fill = PatternFill("solid", fgColor="4472C4")
    header_align = Alignment(horizontal="left", vertical="center")

    col_headers = [
        "Organisation_Name", "Classification", "Latitude", "Longitude",
        "City", "Country", "Geo_Source", "ROR_ID", "ROR_Type",
        "Classification_Source", "Justification",
        "Raw_Variants", "Author_Count", "Authors", "DOI_Count", "DOIs"
    ]
    for col_idx, hdr in enumerate(col_headers, 1):
        cell = out_ws.cell(row=1, column=col_idx, value=hdr)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_align

    class_fills = {
        "Research": PatternFill("solid", fgColor="D6EAF8"),
        "Health": PatternFill("solid", fgColor="E8DAEF"),
        "Government": PatternFill("solid", fgColor="D5F5E3"),
        "Industry": PatternFill("solid", fgColor="FADBD8"),
        "Unclassified": PatternFill("solid", fgColor="F9E79F"),
    }

    data_font = Font(name="Arial")
    wrap_align = Alignment(wrap_text=True, vertical="top")

    for row_idx, org in enumerate(orgs, 2):
        out_ws.cell(row=row_idx, column=1, value=org["name"]).font = data_font
        cls = org.get("classification", "")
        cell = out_ws.cell(row=row_idx, column=2, value=cls)
        cell.font = Font(name="Arial", bold=True)
        cell.fill = class_fills.get(cls, PatternFill())
        out_ws.cell(row=row_idx, column=3, value=org.get("lat")).font = data_font
        out_ws.cell(row=row_idx, column=4, value=org.get("lng")).font = data_font
        out_ws.cell(row=row_idx, column=5, value=org.get("city", "")).font = data_font
        out_ws.cell(row=row_idx, column=6, value=org.get("country", "")).font = data_font
        out_ws.cell(row=row_idx, column=7, value=org.get("geo_source", "")).font = data_font
        out_ws.cell(row=row_idx, column=8, value=org["ror_id"]).font = data_font
        out_ws.cell(row=row_idx, column=9, value=org["ror_type"]).font = data_font
        out_ws.cell(row=row_idx, column=10, value=org.get("cls_source", "")).font = data_font
        cell = out_ws.cell(row=row_idx, column=11, value=org.get("justification", ""))
        cell.font = data_font
        cell.alignment = wrap_align
        cell = out_ws.cell(row=row_idx, column=12, value=org["variants"])
        cell.font = data_font
        cell.alignment = wrap_align
        out_ws.cell(row=row_idx, column=13, value=org["author_count"]).font = data_font
        cell = out_ws.cell(row=row_idx, column=14, value=org["authors"])
        cell.font = data_font
        cell.alignment = wrap_align
        out_ws.cell(row=row_idx, column=15, value=org["doi_count"]).font = data_font
        cell = out_ws.cell(row=row_idx, column=16, value=org["dois"])
        cell.font = data_font
        cell.alignment = wrap_align

    out_ws.column_dimensions['A'].width = 40
    out_ws.column_dimensions['B'].width = 16
    out_ws.column_dimensions['C'].width = 14
    out_ws.column_dimensions['D'].width = 14
    out_ws.column_dimensions['E'].width = 20
    out_ws.column_dimensions['F'].width = 20
    out_ws.column_dimensions['G'].width = 22
    out_ws.column_dimensions['H'].width = 35
    out_ws.column_dimensions['I'].width = 15
    out_ws.column_dimensions['J'].width = 20
    out_ws.column_dimensions['K'].width = 40
    out_ws.column_dimensions['L'].width = 50
    out_ws.column_dimensions['M'].width = 14
    out_ws.column_dimensions['N'].width = 50
    out_ws.column_dimensions['O'].width = 12
    out_ws.column_dimensions['P'].width = 60
    out_ws.freeze_panes = "A2"
    out_ws.auto_filter.ref = f"A1:P{out_ws.max_row}"

    out_wb.save(output_path)
    print(f"\nExcel saved to: {output_path}")

    # --- Generate HTML map ---
    mapped = [o for o in orgs if o.get("lat") is not None]
    map_path = generate_map_html(orgs, output_path)
    print(f"Map saved to:   {map_path}")
    print(f"  {len(mapped)} mapped, {len(orgs) - len(mapped)} not mapped")

    print("\nDone.")


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python geotag_orgs.py <classified_orgs.xlsx> [output.xlsx] [config.ini]")
        sys.exit(1)
    in_arg = sys.argv[1]
    out_arg = sys.argv[2] if len(sys.argv) > 2 else None
    cfg_arg = sys.argv[3] if len(sys.argv) > 3 else None
    main(in_arg, out_arg, cfg_arg)