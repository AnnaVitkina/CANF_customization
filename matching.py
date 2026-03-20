"""
Compare and Find (CANF) - Match Shipments with Rate Card (JSON-only)

Inputs (JSON only):
- vocabulary_mapping.json (etof_data: list of shipments)
- Filtered_Rate_Card_with_Conditions.json (rate_card_data, conditions, business_rules)

For each shipment, compare to every lane using Has conditional Rule / Has Business Rule;
pick lane(s) with least differences. Output: JSON file + Excel file created from that JSON.
"""

import json
import os
import re
import sys
from datetime import datetime

import pandas as pd

# Columns used only for date-range filtering; excluded from value comparison
VALIDITY_DATE_COLUMNS = ("Valid to", "Valid from")


def _get_lane_value_columns(lane_dict):
    """Return list of column names to compare: exclude 'Lane #' and ' - Has Business Rule' / ' - Has conditional Rule'."""
    return [
        k for k in lane_dict
        if k != "Lane #"
        and not k.endswith(" - Has Business Rule")
        and not k.endswith(" - Has conditional Rule")
    ]


def _rate_card_columns_contains(rate_card_columns_str, target_col):
    """True if target_col is one of the columns in the comma-separated rate_card_columns_str."""
    if not rate_card_columns_str or not target_col:
        return False
    parts = [p.strip() for p in str(rate_card_columns_str).split(",") if p.strip()]
    return target_col in parts


def _get_lane_origin_destination_countries(lane, business_rules_list):
    """
    Derive lane's origin and destination countries from business_rules.
    Rate Card Columns can be "Origin City, Destination" (comma-separated); we match by column and rule name vs lane value.
    Origin: rule applies to Origin City and rule name matches lane's "Origin City".
    Destination: rule applies to "Destination" (or "Destination City") and rule name matches lane's "Destination" or "Destination City".
    Returns: (origin_countries: set of normalized codes, dest_countries: set of normalized codes).
    """
    origin_countries = set()
    dest_countries = set()
    for r in business_rules_list or []:
        name = r.get("Rule Name")
        cols_str = r.get("Rate Card Columns")
        country = r.get("Country")
        if not name or not country:
            continue
        raw = str(country).strip()
        if not raw:
            continue
        codes = {c.strip().lower() for c in raw.split(",") if c.strip()}
        if _rate_card_columns_contains(cols_str, "Origin City") and _normalize_for_compare(name) == _normalize_for_compare(lane.get("Origin City")):
            origin_countries = codes
            break
    for r in business_rules_list or []:
        name = r.get("Rule Name")
        cols_str = r.get("Rate Card Columns")
        country = r.get("Country")
        if not name or not country:
            continue
        raw = str(country).strip()
        if not raw:
            continue
        codes = {c.strip().lower() for c in raw.split(",") if c.strip()}
        # Rate card uses "Destination" (not "Destination City"); lane has "Destination"
        lane_dest_val = lane.get("Destination") or lane.get("Destination City")
        if (_rate_card_columns_contains(cols_str, "Destination") or _rate_card_columns_contains(cols_str, "Destination City")) and _normalize_for_compare(name) == _normalize_for_compare(lane_dest_val):
            dest_countries = codes
            break
    return origin_countries, dest_countries


def _lane_matches_shipment_countries(shipment, lane, business_rules_list):
    """
    True if lane's origin/destination countries (from business rules) match shipment's Origin Country / Destination Country.
    If lane has no rule for origin or dest, that side is considered a match (don't exclude).
    """
    ship_orig = _normalize_for_compare(shipment.get("Origin Country") or shipment.get("SHIP_COUNTRY"))
    ship_dest = _normalize_for_compare(shipment.get("Destination Country") or shipment.get("CUST_COUNTRY"))
    lane_orig, lane_dest = _get_lane_origin_destination_countries(lane, business_rules_list)
    if lane_orig and ship_orig and ship_orig not in lane_orig:
        return False
    if lane_dest and ship_dest and ship_dest not in lane_dest:
        return False
    return True


def _normalize_for_compare(val):
    """Normalize value for comparison (lowercase, strip, treat None/empty)."""
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return None
    s = str(val).strip()
    if s.lower() in ("", "nan", "none"):
        return None
    return s.lower()


def _parse_date_for_validity(s):
    """
    Parse a date string to a date object for range comparison.
    Supports: YYYY-MM-DD, DD.MM.YYYY, DD/MM/YYYY, and similar.
    Returns None if unparseable or empty.
    """
    if s is None or (isinstance(s, float) and pd.isna(s)):
        return None
    s = str(s).strip()
    if not s or s.lower() in ("nan", "none", ""):
        return None
    for fmt in ("%Y-%m-%d", "%d.%m.%Y", "%d/%m/%Y", "%d-%m-%Y", "%Y/%m/%d"):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    return None


def _lane_valid_for_shipment_date(lane, ship_date_str):
    """
    True if the lane should be considered for this shipment based on Valid from / Valid to.
    Rule: SHIP_DATE must be on or after Valid from (if present) and on or before Valid to (if present).
    If both Valid from and Valid to are missing/empty, the lane is valid.
    """
    ship_d = _parse_date_for_validity(ship_date_str)
    valid_to_str = lane.get("Valid to")
    valid_from_str = lane.get("Valid from")
    valid_to = _parse_date_for_validity(valid_to_str)
    valid_from = _parse_date_for_validity(valid_from_str)

    # No validity dates on lane -> lane can be checked
    if valid_to is None and valid_from is None:
        return True

    # No shipment date -> cannot enforce range; allow lane to be checked (or could return False)
    if ship_d is None:
        return True

    if valid_to is not None and ship_d > valid_to:
        return False
    if valid_from is not None and ship_d < valid_from:
        return False
    return True


def _parse_condition_rule_for_rate_card_value(condition_rule_text, rate_card_value):
    """
    Find the line in condition_rule_text that applies to rate_card_value (e.g. 'Shanghai', 'Reefer').
    Only a line whose label exactly matches the rate card value is used; otherwise we fall back to plain compare.
    Returns: (codes_list, mode) or (None, None) if not found.
    mode: True = equals (value must be in list), False = does not equal, 'contains' = value must contain one of codes,
    'not_contains' = value must NOT contain any of codes.
    """
    if not condition_rule_text or not rate_card_value:
        return None, None
    rc_val_norm = _normalize_for_compare(rate_card_value)
    if not rc_val_norm:
        return None, None
    rc_norm = rc_val_norm.replace(" ", "") if rc_val_norm else ""
    lines = [ln.strip() for ln in str(condition_rule_text).split("\n") if ln.strip()]
    for line in lines:
        if ":" not in line:
            continue
        rest = re.sub(r"^\d+\.\s*", "", line).strip()
        if ":" not in rest:
            continue
        label, rule_part = rest.split(":", 1)
        label_norm = label.strip().lower().replace(" ", "")
        # Exact match only: e.g. rate card "STD_HUB_ATD" must not match label "STD_HUB_ATD/STD_DIR_ATD"
        if label_norm != rc_norm:
            continue
        rule_part = rule_part.strip().lower()
        if "does not equal" in rule_part or "does not equal to" in rule_part:
            mode = False
            part = rule_part.replace("does not equal to", "").replace("does not equal", "").strip()
        elif "does not contain" in rule_part:
            # Must check before "contains" so "does not contain" is not treated as "contains"
            mode = "not_contains"
            part = rule_part.split("does not contain", 1)[1].strip()
        elif "contains" in rule_part:
            mode = "contains"
            part = rule_part.split("contains", 1)[1].strip()
        elif "equals" in rule_part:
            mode = True
            part = rule_part.split("equals", 1)[1].strip()
        else:
            continue
        part = part.split(" in ")[0].strip() if " in " in part else part
        codes = [c.strip().lower() for c in part.split(",") if c.strip()]
        if codes:
            return codes, mode
    return None, None


def _check_conditional_rule(shipment_value, rate_card_value, column_name, conditions_list):
    """Returns: (matches: bool, differ_message: str or None)."""
    cond = None
    for c in conditions_list or []:
        if c.get("Column") == column_name and (c.get("Has Condition") == "Yes" or c.get("Condition Rule")):
            cond = c
            break
    if not cond or not cond.get("Condition Rule"):
        return True, None
    allowed, mode = _parse_condition_rule_for_rate_card_value(cond["Condition Rule"], rate_card_value)
    rule_name = rate_card_value  # condition label / rate card value is the "rule name" for display
    if allowed is None:
        # Rate card value not in conditional rules / unparsed -> plain compare; do not label as (Conditional)
        match = _normalize_for_compare(shipment_value) == _normalize_for_compare(rate_card_value)
        if not match:
            return False, f"{column_name}: rate card '{rate_card_value}' -> shipment '{shipment_value}' differs"
        return True, None
    ship_norm = _normalize_for_compare(shipment_value)
    if mode is True:  # equals
        match = ship_norm and ship_norm in allowed
        if not match:
            return False, f"{column_name} (Conditional) \"{rule_name}\" -> shipment value '{shipment_value}' differs"
        return True, None
    if mode == "contains":
        # Shipment value must contain at least one of the codes (e.g. "FCL - 40HC" contains "HC")
        match = ship_norm and any(code in ship_norm for code in allowed)
        if not match:
            return False, f"{column_name} (Conditional) \"{rule_name}\" -> shipment value '{shipment_value}' differs"
        return True, None
    if mode == "not_contains":
        # Shipment value must NOT contain any of the codes (e.g. "not ODH" -> value must not contain "ODH")
        match = not ship_norm or not any(code in ship_norm for code in allowed)
        if not match:
            return False, f"{column_name} (Conditional) \"{rule_name}\" -> shipment value '{shipment_value}' differs"
        return True, None
    # mode is False: does not equal
    match = not ship_norm or ship_norm not in allowed
    if not match:
        return False, f"{column_name} (Conditional) \"{rule_name}\" -> shipment value '{shipment_value}' differs"
    return True, None


def _get_origin_destination_from_shipment(shipment, is_origin):
    """Get (country, postal) from shipment for origin (SHIP) or destination (CUST)."""
    if is_origin:
        country = shipment.get("Origin Country") or shipment.get("SHIP_COUNTRY")
        postal = shipment.get("Origin Postal Code") or shipment.get("SHIP_POST")
    else:
        country = shipment.get("Destination Country") or shipment.get("CUST_COUNTRY")
        postal = shipment.get("Destination Postal Code") or shipment.get("CUST_POST")
    return _normalize_for_compare(country), _normalize_for_compare(postal)


def _check_business_rule(shipment, rate_card_value, column_name, business_rules_list):
    """Returns: (diff_count, list of differ messages)."""
    rule = None
    for r in business_rules_list or []:
        if not _rate_card_columns_contains(r.get("Rate Card Columns"), column_name):
            continue
        name = r.get("Rule Name")
        if name and _normalize_for_compare(name) == _normalize_for_compare(rate_card_value):
            rule = r
            break
    if not rule:
        return 0, []
    col_lower = column_name.lower()
    is_origin = "origin" in col_lower or "loading" in col_lower or "ship" in col_lower
    ship_country, ship_postal = _get_origin_destination_from_shipment(shipment, is_origin=is_origin)
    rule_country = rule.get("Country")
    rule_country = [c.strip().lower() for c in str(rule_country).split(",") if c.strip()] if rule_country else []
    rule_postal_raw = rule.get("Postal Codes")
    rule_postal = [x.strip().lower() for x in str(rule_postal_raw).split(",") if x.strip()] if rule_postal_raw else []
    exclude_raw = rule.get("Exclude")
    exclude_vals = None
    if exclude_raw and str(exclude_raw).strip().lower() not in ("no", ""):
        exclude_vals = [x.strip().lower() for x in str(exclude_raw).split(",") if x.strip()]
    # Display values: show origin vs destination based on column (so Destination rule shows Destination Country/Postal, not Origin)
    if is_origin:
        display_country = shipment.get("Origin Country") or shipment.get("SHIP_COUNTRY")
        display_postal = shipment.get("Origin Postal Code") or shipment.get("SHIP_POST")
    else:
        display_country = shipment.get("Destination Country") or shipment.get("CUST_COUNTRY")
        display_postal = shipment.get("Destination Postal Code") or shipment.get("CUST_POST")
    rule_name = rule.get("Rule Name") or rate_card_value
    dest_orig = "origin" if is_origin else "destination"
    diffs = []
    if rule_country:
        if not ship_country or ship_country not in rule_country:
            diffs.append(f"{column_name} (Business Rule) \"{rule_name}\" -> shipment {dest_orig} country '{display_country}' differs")
    if rule_postal:
        if not ship_postal:
            diffs.append(f"{column_name} (Business Rule) \"{rule_name}\" -> shipment {dest_orig} postal missing")
        else:
            if not any(ship_postal.startswith(p) for p in rule_postal):
                diffs.append(f"{column_name} (Business Rule) \"{rule_name}\" -> shipment {dest_orig} postal '{display_postal}' differs")
    if exclude_vals and ship_postal:
        for ex in exclude_vals:
            if ship_postal.startswith(ex) or ex in ship_postal:
                diffs.append(f"{column_name} (Business Rule) \"{rule_name}\" (Exclude) -> shipment {dest_orig} postal '{display_postal}' is excluded")
                break
    return len(diffs), diffs


def _column_from_diff(diff_str):
    """Extract column name from a difference string (e.g. 'Service Type: ...' -> 'Service Type', 'Service (Conditional) \"X\" -> ...' -> 'Service')."""
    if not diff_str:
        return ""
    s = diff_str.strip()
    if " (Conditional)" in s:
        return s.split(" (Conditional)")[0].strip()
    if " (Business Rule)" in s:
        return s.split(" (Business Rule)")[0].strip()
    if ":" in s:
        return s.split(":", 1)[0].strip()
    return ""


def _is_geo_column(col):
    """True if column is city/airport/port/postal (not country)."""
    if not col:
        return False
    c = col.lower()
    return (
        "postal" in c or "post " in c or c.endswith(" post")
        or "city" in c
        or "airport" in c
        or "port" in c
        or "seaport" in c
    ) and "country" not in c


def _is_country_column(col):
    """True if column is origin/destination country."""
    if not col:
        return False
    return "country" in col.lower()


def _display_priority(diff_str):
    """
    Return sort key for display order (only the order matters; section numbers are assigned dynamically):
    0 = Service, 1 = Airport/Seaport, 2 = City/Postal code, 3 = Country, 4 = Other, 5 = Service SPECIAL/EXP_DUTY (always last).
    Used to order differs when showing per-lane results.
    """
    col = _column_from_diff(diff_str)
    d = (diff_str or "").lower()
    # Service differ with rate card 'SPECIAL' or 'EXP_DUTY' is always last priority
    if col and "service" in col.lower():
        if ("'special'" in d or "rate card 'special'" in d or "rate card \"special\"" in d or
            "'exp_duty'" in d or "rate card 'exp_duty'" in d or "rate card \"exp_duty\"" in d):
            return 5
        return 0
    # Airport/Seaport (second priority when present)
    if col and ("airport" in col.lower() or "seaport" in col.lower()):
        return 1
    # City/Postal: column is city/postal, or message mentions postal (e.g. "destination postal", "origin postal")
    if (
        "postal" in (col or "").lower()
        or "city" in (col or "").lower()
        or "postal" in d
        or "postal codes" in d
        or "postal missing" in d
        or ("destination" in (col or "").lower() and "postal" in d)
        or ("origin" in (col or "").lower() and "postal" in d)
        or _is_geo_column(col)
    ):
        return 2
    # Country (fourth when present)
    if (
        "country" in (col or "").lower()
        or "country rule" in d
        or "destination country" in d
        or "origin country" in d
        or _is_country_column(col or "")
    ):
        return 3
    return 4


def _service_parts(val):
    """Split service value into parts (by space or underscore). Returns list of up to 3 parts."""
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return []
    s = str(val).strip().replace("_", " ")
    parts = [p for p in s.split() if p]
    return parts[:3] if len(parts) >= 3 else parts


def _service_first_two_match(shipment_val, lane_val):
    """True if both values have at least 2 parts and first two parts match (case-insensitive)."""
    sp = _service_parts(shipment_val)
    lp = _service_parts(lane_val)
    if len(sp) < 2 or len(lp) < 2:
        return False
    return sp[0].lower() == lp[0].lower() and sp[1].lower() == lp[1].lower()


def _priority_key(shipment, lane, diffs, rate_card_to_isd_key):
    """
    Return a sort key (tuple) for prioritization: lower = higher priority.
    P1: has ISD match (rate card value == shipment[ISD] for a differing column)
    P2: service column differ with first two parts matching
    P3: other service differ
    P4: geo (postal/city/port/airport) differ
    P5: country differ
    """
    has_isd = False
    has_service_2part = False
    has_service_other = False
    has_geo = False
    has_country = False
    for d in diffs:
        col = _column_from_diff(d)
        pri = _display_priority(d)  # 0=service, 1=airport, 2=city/postal, 3=country, 4=other
        if pri == 3:
            has_country = True
        elif pri in (1, 2):
            has_geo = True
        elif pri == 0 and col:
            rc_val = lane.get(col)
            ship_val = shipment.get(col)
            if _service_first_two_match(ship_val, rc_val):
                has_service_2part = True
            else:
                has_service_other = True
        # ISD: for this column, rate card value equals shipment's ISD value
        if col:
            isd_key = rate_card_to_isd_key.get(col)
            if isd_key and shipment.get(isd_key) is not None and str(shipment.get(isd_key)).strip():
                if _normalize_for_compare(lane.get(col)) == _normalize_for_compare(shipment.get(isd_key)):
                    has_isd = True
    # Sort: ISD first, then service 2-part, then service other, then geo, then country last
    return (
        0 if has_isd else 1,
        0 if has_service_2part else 1,
        0 if has_service_other else 1,
        0 if has_geo else 1,
        1 if has_country else 0,
    )


def _isd_match_columns(shipment, lane, diffs, rate_card_to_isd_key):
    """Return set of column names (from diffs) where lane value matches shipment's ISD value."""
    out = set()
    for d in diffs:
        col = _column_from_diff(d)
        if not col:
            continue
        isd_key = rate_card_to_isd_key.get(col)
        if isd_key and shipment.get(isd_key) is not None and str(shipment.get(isd_key)).strip():
            if _normalize_for_compare(lane.get(col)) == _normalize_for_compare(shipment.get(isd_key)):
                out.add(col)
    return out


def compare_shipment_to_lane(shipment, lane, conditions_list, business_rules_list, value_columns):
    """Compare one shipment to one lane. Returns: (diff_count, list of difference strings)."""
    differences = []
    for col in value_columns:
        rc_val = lane.get(col)
        # Rate card null/empty means "ANY value from etof_data is acceptable" -> no difference for this column
        if _normalize_for_compare(rc_val) is None:
            continue
        has_business = lane.get(col + " - Has Business Rule") == "Yes"
        has_conditional = lane.get(col + " - Has conditional Rule") == "Yes"
        ship_val = shipment.get(col)

        if has_conditional:
            match, msg = _check_conditional_rule(ship_val, rc_val, col, conditions_list)
            if not match and msg:
                differences.append(msg)
            continue
        if has_business and rc_val:
            n, msgs = _check_business_rule(shipment, rc_val, col, business_rules_list)
            differences.extend(msgs)
            continue
        rn = _normalize_for_compare(rc_val)
        sn = _normalize_for_compare(ship_val)
        if rn is None:
            continue
        if sn != rn:
            differences.append(f"{col}: rate card '{rc_val}' -> shipment '{ship_val}' differs")
    return len(differences), differences


def run_matching_json_only(
    vocabulary_json_path=None,
    rate_card_json_path=None,
    output_dir=None,
):
    """
    Load vocabulary_mapping.json and Filtered_Rate_Card_with_Conditions.json;
    compare each shipment to every lane; pick best lane(s); write JSON then Excel from that JSON.
    Returns: (path_to_xlsx, path_to_json) or (None, None).
    """
    try:
        script_dir = os.path.dirname(os.path.abspath(__file__))
    except NameError:
        script_dir = os.getcwd()
    partly_df = os.path.join(script_dir, "partly_df")
    vocabulary_json_path = vocabulary_json_path or os.path.join(partly_df, "vocabulary_mapping.json")
    rate_card_json_path = rate_card_json_path or os.path.join(partly_df, "Filtered_Rate_Card_with_Conditions.json")
    output_dir = output_dir or partly_df

    if not os.path.exists(vocabulary_json_path):
        print(f"[ERROR] Vocabulary JSON not found: {vocabulary_json_path}")
        return None, None
    if not os.path.exists(rate_card_json_path):
        print(f"[ERROR] Rate card JSON not found: {rate_card_json_path}")
        return None, None

    with open(vocabulary_json_path, "r", encoding="utf-8") as f:
        vocab = json.load(f)
    etof_data = vocab.get("etof_data", [])
    if not etof_data:
        print("[ERROR] No etof_data in vocabulary_mapping.json")
        return None, None

    with open(rate_card_json_path, "r", encoding="utf-8") as f:
        rate_card = json.load(f)
    rate_card_data = rate_card.get("rate_card_data", [])
    conditions_list = rate_card.get("conditions", [])
    business_rules_list = rate_card.get("business_rules", [])

    if not rate_card_data:
        print("[ERROR] No rate_card_data in Filtered_Rate_Card_with_Conditions.json")
        return None, None

    all_value_columns = _get_lane_value_columns(rate_card_data[0])
    value_columns = [c for c in all_value_columns if c not in VALIDITY_DATE_COLUMNS]
    print(f"Comparing on columns: {value_columns} (Valid to/Valid from used for date filter only)")

    # Build rate_card_column -> ISD key (e.g. "Service Type" -> "SERVICE_ISD") from vocabulary mapping
    etof_mappings = vocab.get("etof_mappings", {})
    rate_card_to_isd_key = {rc: etof + "_ISD" for rc, etof in etof_mappings.items()}

    results = []
    for idx, shipment in enumerate(etof_data):
        # AUSID: if ORIG_FILE_NAME contains AUSID, only consider lanes whose origin/destination countries match shipment
        orig_file = str(shipment.get("ORIG_FILE_NAME") or "").upper()
        use_country_filter = "AUSID" in orig_file
        candidates = rate_card_data
        if use_country_filter:
            candidates = [lane for lane in rate_card_data if _lane_matches_shipment_countries(shipment, lane, business_rules_list)]
            if len(candidates) < len(rate_card_data):
                print(f"Shipment {idx + 1} (ETOF {shipment.get('ETOF')}): AUSID -> filtered to {len(candidates)} lanes (origin/destination country match)")
            if not candidates:
                # Fallback: no lane matched country filter -> use standard lookup over all lanes
                candidates = rate_card_data
                print(f"Shipment {idx + 1} (ETOF {shipment.get('ETOF')}): AUSID -> no lane matched origin/destination country; fallback to standard lookup (all lanes)")

        # Valid to / Valid from: only consider lanes where SHIP_DATE is on or after Valid from and on or before Valid to
        ship_date = shipment.get("SHIP_DATE")
        before_date_filter = len(candidates)
        candidates = [lane for lane in candidates if _lane_valid_for_shipment_date(lane, ship_date)]
        if len(candidates) < before_date_filter:
            print(f"Shipment {idx + 1} (ETOF {shipment.get('ETOF')}): Valid to/Valid from -> {len(candidates)} lanes (SHIP_DATE {ship_date} in range)")
        if not candidates:
            print(f"Shipment {idx + 1} (ETOF {shipment.get('ETOF')}): No lanes in Valid to/Valid from range for SHIP_DATE {ship_date}")

        lane_scores = []
        for lane in candidates:
            lane_num = lane.get("Lane #", "")
            diff_count, diffs = compare_shipment_to_lane(
                shipment, lane, conditions_list, business_rules_list, value_columns
            )
            lane_scores.append((lane_num, diff_count, diffs, lane))
        min_diff = min(s[1] for s in lane_scores) if lane_scores else None
        best_lanes = [s for s in lane_scores if s[1] == min_diff] if lane_scores else []
        # Sort by priority: ISD match > service 2-part match > service other > geo > country
        sorted_best = sorted(
            best_lanes,
            key=lambda s: _priority_key(shipment, s[3], s[2], rate_card_to_isd_key),
        )
        best_lane_nums = [s[0] for s in sorted_best]
        # Build differences list: show ALL differs per lane; order LANES by priority (service first, then postal, then country)
        if not best_lane_nums:
            lane_label = "No lane (none in Valid to/Valid from range or no match)"
            best_diffs = sorted_best[0][2] if sorted_best else []
            if not best_diffs:
                diffs_with_lane = [f"{lane_label}: SHIP_DATE not within any lane's Valid from–Valid to range"]
            else:
                diffs_with_lane = [f"{lane_label}: {d}" for d in best_diffs]
        else:
            diffs_with_lane = []
            if len(best_lane_nums) > 1:
                diffs_with_lane.append(f"Best lanes (tied): {', '.join(str(b) for b in best_lane_nums)} ({len(best_lane_nums)} lanes with {min_diff} differ(s))")
            # Category labels (numbers are assigned dynamically by order of appearance)
            _priority_categories = {
                0: "Service differs",
                1: "Airport/Seaport differs",
                2: "City/Postal code differs",
                3: "Country differs",
                4: "Other differs",
                5: "Service (SPECIAL/EXP_DUTY) differs",  # always last
            }
            # Sort lanes by display priority of their differ(s): service first, then airport/seaport, then city/postal, then country, then other
            def _lane_display_priority(item):
                lane_num, _dc, diffs, _lane = item
                if not diffs:
                    return 99
                return min(_display_priority(d) for d in diffs)
            lanes_for_display = sorted(sorted_best, key=_lane_display_priority)
            # Collect (priority, lane_num, "Lane N: message") for all differs, then group by priority with section headers
            all_items = []
            for lane_num, _diff_count, diffs, lane in lanes_for_display:
                isd_match_cols = _isd_match_columns(shipment, lane, diffs, rate_card_to_isd_key)
                def add_isd_comment(d):
                    col = _column_from_diff(d)
                    if col in isd_match_cols:
                        return d + " (matches the Carrier data (ISD))"
                    return d
                ordered_diffs = sorted(diffs, key=_display_priority)
                annotated = [add_isd_comment(d) for d in ordered_diffs]
                for d in annotated:
                    pri = _display_priority(d)
                    all_items.append((pri, str(lane_num), f"Lane {lane_num}: {d}"))
            # Sort by priority, then by lane number; emit section header when priority changes (Priority 1, 2, 3... by order of appearance)
            all_items.sort(key=lambda x: (x[0], x[1]))
            last_pri = None
            section_number = 0
            for pri, _ln, line in all_items:
                if last_pri is None or last_pri != pri:
                    section_number += 1
                    category = _priority_categories.get(pri, "Other differs")
                    diffs_with_lane.append(f"Priority {section_number}: {category}")
                    last_pri = pri
                diffs_with_lane.append(line)
        row = dict(shipment)
        row["best_lane(s)"] = ", ".join(str(b) for b in best_lane_nums)
        row["diff_count"] = min_diff
        row["differences"] = "\n".join(diffs_with_lane) if diffs_with_lane else ""
        row["differences_list"] = diffs_with_lane
        results.append(row)
        if len(best_lane_nums) > 1:
            print(f"Shipment {idx + 1} (ETOF {row.get('ETOF')}): {len(best_lane_nums)} lanes tied with {min_diff} differ(s): {best_lane_nums}")
        else:
            print(f"Shipment {idx + 1} (ETOF {row.get('ETOF')}): best lane(s) {best_lane_nums}, {min_diff} differ(s)")

    os.makedirs(output_dir, exist_ok=True)
    out_json = os.path.join(output_dir, "Matched_Shipments_with.json")
    out_xlsx = os.path.join(output_dir, "Matched_Shipments_with.xlsx")

    # 1) Write JSON (canonical output)
    json_payload = {"matched_shipments": []}
    for r in results:
        j = {k: v for k, v in r.items() if k != "differences_list"}
        j["differences_list"] = r.get("differences_list", [])
        json_payload["matched_shipments"].append(j)
    with open(out_json, "w", encoding="utf-8") as f:
        json.dump(json_payload, f, indent=2, ensure_ascii=False)
    print(f"Saved: {out_json}")

    # 2) Create Excel from the JSON file (converted from JSON)
    with open(out_json, "r", encoding="utf-8") as f:
        data = json.load(f)
    rows = data.get("matched_shipments", [])
    excel_rows = []
    for r in rows:
        row = {k: v for k, v in r.items() if k != "differences_list"}
        excel_rows.append(row)
    df = pd.DataFrame(excel_rows)
    with pd.ExcelWriter(out_xlsx, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Matched Shipments", index=False)
    print(f"Saved: {out_xlsx} (from JSON)")

    return out_xlsx, out_json


def run_matching_from_json(rate_card_json_path=None, vocabulary_json_path=None, output_dir=None):
    """Convenience wrapper: run JSON-only matching and return (xlsx_path, json_path)."""
    return run_matching_json_only(
        vocabulary_json_path=vocabulary_json_path,
        rate_card_json_path=rate_card_json_path,
        output_dir=output_dir,
    )


if __name__ == "__main__":
    run_matching_from_json()
