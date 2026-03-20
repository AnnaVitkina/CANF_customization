"""
Format the matched shipments JSON: add "Possible Best Match" column by picking
the single best lane among tied lanes (Service exact match first, then City/Postal rule name match).
"""

import json
import os
import re


def _normalize(s):
    """Normalize for comparison: lower, strip, replace spaces with underscore."""
    if s is None:
        return ""
    return str(s).strip().lower().replace(" ", "_")


def _normalize_alnum(s):
    """Normalize to alphanumeric only (for containment: 'Elk Grove' in 'US_ELKGROVE')."""
    if s is None:
        return ""
    return re.sub(r"[^a-z0-9]", "", str(s).strip().lower())


def _parse_differences_list(differences_list):
    """
    Parse differences_list into sections and typed entries.
    Returns: {
        "service": [(lane_num, rate_card_val, shipment_val), ...],
        "geo": [(lane_num, column, rule_name, side, field, shipment_val), ...],
          side in origin|destination, field in postal|country (Priority 2 / CityPostal business rules).
    }
    Explicit country-only section (Priority 3) is ignored for best-match logic.
    """
    out = {"service": [], "geo": []}
    if not differences_list:
        return out

    current_section = None  # "service" | "postal" | "country" (we skip country)
    for line in differences_list:
        line = (line or "").strip()
        if not line:
            continue
        # Check section by content first (more reliable than priority number)
        if "Service differs" in line or "Service (SPECIAL" in line or "Service (EXP_DUTY" in line:
            current_section = "service"
            continue
        if "City/Postal" in line:
            current_section = "postal"
            continue
        if "Country differs" in line:
            current_section = "country"
            continue
        if line.startswith("Priority "):
            # Fallback: if no content match, use priority number (but less reliable)
            if "Service" in line:
                current_section = "service"
            elif "City/Postal" in line or "Airport" in line:
                current_section = "postal"
            elif "Country" in line:
                current_section = "country"
            else:
                current_section = None
            continue
        if line.startswith("Best lanes") or "Other differs" in line:
            continue

        # "Lane N: ..."
        lane_m = re.match(r"Lane\s+(\d+)\s*:\s*(.+)", line, re.IGNORECASE)
        if not lane_m:
            continue
        lane_num = lane_m.group(1).strip()
        rest = lane_m.group(2).strip()

        # Service: "Service: rate card 'X' -> shipment 'Y' differs" or "Service (Conditional) \"X\" -> shipment value 'Y' differs"
        # OR reformatted: "Shipment service value 'Y' should be changed to 'X'"
        service_m = re.search(r"rate card\s+'([^']*)'\s*->\s*shipment\s+'([^']*)'", rest)
        if service_m and current_section == "service":
            out["service"].append((lane_num, service_m.group(1), service_m.group(2)))
            continue
        service_cond_m = re.search(
            r'Service\s*\(Conditional\)\s*"([^"]+)"\s*->\s*shipment value\s+\'([^\']*)\'',
            rest,
        )
        if service_cond_m and current_section == "service":
            out["service"].append(
                (lane_num, service_cond_m.group(1), service_cond_m.group(2))
            )
            continue
        # Reformatted service: "Shipment service value 'Y' should be changed to 'X'"
        service_reformatted_m = re.search(
            r"Shipment service value\s+'([^']*)'\s+should be changed to\s+'([^']*)'",
            rest,
            re.IGNORECASE,
        )
        if service_reformatted_m and current_section == "service":
            # Note: group(1) is shipment_val, group(2) is rate_card_val (reversed from original format)
            out["service"].append(
                (lane_num, service_reformatted_m.group(2), service_reformatted_m.group(1))
            )
            continue

        # Any column (Business Rule) -> shipment origin|destination postal|country 'val' differs
        # Covers Destination, Origin City, Origin Airport, etc.
        br_geo_m = re.search(
            r"^(.+?)\s*\(Business Rule\)\s*\"([^\"]+)\"\s*->\s*shipment\s+(origin|destination)\s+(postal|country)\s+\'([^\']*)\'\s*differs",
            rest,
            re.IGNORECASE,
        )
        if br_geo_m and (current_section == "postal" or current_section == "country"):
            out["geo"].append(
                (
                    lane_num,
                    br_geo_m.group(1).strip(),
                    br_geo_m.group(2),
                    br_geo_m.group(3).lower(),
                    br_geo_m.group(4).lower(),
                    br_geo_m.group(5),
                )
            )
            continue
        
        # Reformatted geo: "Shipment Origin Country value 'CN' does not match rate card 'AT'"
        # or "Shipment origin country 'CN' does not match the origin \"AT\""
        geo_reformatted_rate_card = re.search(
            r"Shipment\s+(Origin|Destination)\s+(Country|Postal|postal)\s+(?:value\s+)?'([^']*)'\s+does not match\s+rate card\s+'([^']*)'",
            rest,
            re.IGNORECASE,
        )
        if geo_reformatted_rate_card and (current_section == "postal" or current_section == "country"):
            side = geo_reformatted_rate_card.group(1).lower()
            field_raw = geo_reformatted_rate_card.group(2).lower()
            shipment_val = geo_reformatted_rate_card.group(3)
            rule_name = geo_reformatted_rate_card.group(4)
            # Use a generic column name for reformatted format
            col_name = "Origin" if side == "origin" else "Destination"
            # Normalize field: "country" or "postal"
            field = "country" if "country" in field_raw else "postal"
            if field == "postal" and current_section == "postal":
                out["geo"].append((lane_num, col_name, rule_name, side, "postal", shipment_val))
            elif field == "country" and current_section == "country":
                out["geo"].append((lane_num, col_name, rule_name, side, "country", shipment_val))
            continue
        
        geo_reformatted_rule = re.search(
            r"Shipment\s+(Origin|Destination)\s+(Country|postal)\s+(?:value\s+)?'([^']*)'\s+does not match\s+(?:the\s+(?:origin|destination|Origin|Destination)\s+)?\"([^\"]+)\"",
            rest,
            re.IGNORECASE,
        )
        if geo_reformatted_rule and (current_section == "postal" or current_section == "country"):
            side = geo_reformatted_rule.group(1).lower()
            field = geo_reformatted_rule.group(2).lower()
            shipment_val = geo_reformatted_rule.group(3)
            rule_name = geo_reformatted_rule.group(4)
            col_name = "Origin" if side == "origin" else "Destination"
            if field == "postal" and current_section == "postal":
                out["geo"].append((lane_num, col_name, rule_name, side, "postal", shipment_val))
            elif field == "country" and current_section == "country":
                out["geo"].append((lane_num, col_name, rule_name, side, "country", shipment_val))
            continue

    return out


def _best_match_service(service_entries):
    """
    Among service entries (lane_num, rate_card_val, shipment_val), return the lane number(s)
    whose rate_card_val best matches shipment_val. Best = exact match (normalized).
    If no exact match, prefer same first two parts (e.g. STD_DIR_ATA vs STD_DIR_DTA).
    Returns: single lane_num (str) for exact match, or list of lane_nums (str) for first-2-parts matches.
    """
    if not service_entries:
        return None
    shipment_val = None
    for _ln, _rc, sv in service_entries:
        shipment_val = sv
        break
    if not shipment_val:
        return None
    ship_norm = _normalize(shipment_val)

    # Exact match
    for lane_num, rc_val, sv in service_entries:
        if _normalize(rc_val) == ship_norm:
            return lane_num

    # Same first two parts (e.g. STD_DIR_ATA vs STD_DIR_DTA)
    def first_two_parts(v):
        parts = re.split(r"[\s_]+", _normalize(v))
        return (parts[0], parts[1]) if len(parts) >= 2 else (None, None)

    ship_parts = first_two_parts(shipment_val)
    matching_lanes = []
    for lane_num, rc_val, _sv in service_entries:
        if first_two_parts(rc_val) == ship_parts:
            matching_lanes.append(lane_num)
    
    if matching_lanes:
        return matching_lanes if len(matching_lanes) > 1 else matching_lanes[0]

    return None


# Order to try geo buckets when picking Possible Best Match (dest postal first, etc.)
_GEO_TRY_ORDER = [
    ("destination", "postal"),
    ("destination", "country"),
    ("origin", "postal"),
    ("origin", "country"),
]


def _best_match_tied_geo(entries):
    """
    entries: list of (lane_num, rule_name, shipment_value).
    If rule name matches shipment value (alnum), return best length matches.
    If none match but all same shipment_value, return all lanes for combined display.
    Returns: list of (lane_num, rule_name) or None.
    """
    if not entries:
        return None
    shipment_value = None
    for _ln, _rn, sv in entries:
        shipment_value = sv
        break
    if not shipment_value:
        return None
    ship_alnum = _normalize_alnum(shipment_value)
    all_same = all(sv == shipment_value for _ln, _rn, sv in entries)

    candidates = []
    for lane_num, rule_name, sv in entries:
        if sv != shipment_value:
            continue
        rn_alnum = _normalize_alnum(rule_name)
        if not rn_alnum:
            continue
        if rn_alnum in ship_alnum:
            candidates.append((len(rn_alnum), lane_num, rule_name))
        elif ship_alnum.startswith(rn_alnum) or rn_alnum.startswith(ship_alnum):
            candidates.append((len(rn_alnum), lane_num, rule_name))

    if candidates:
        candidates.sort(key=lambda x: (-x[0], x[1]))
        max_len = candidates[0][0]
        result = [(ln, rn) for _len, ln, rn in candidates if _len == max_len]
        return result if len(result) > 1 else [(candidates[0][1], candidates[0][2])]
    if all_same:
        return [(ln, rn) for ln, rn, _sv in entries]
    return None


# Service type that should never be chosen as Possible Best Match (e.g. catch-all "SPECIAL", "EXP_DUTY")
EXCLUDED_SERVICE_FOR_BEST_MATCH = {"SPECIAL", "EXP_DUTY"}


def compute_possible_best_match(differences_list):
    """
    From parsed differences_list, compute the best lane(s) and which section chose it.
    Returns (lane_info, section) where lane_info is:
    - For service: single lane_num (str) for exact match, or list of lane_nums (str) for first-2-parts matches
    - For geo: list of (lane_num, rule_name) tuples
    section is "service" or ("geo", side, field) with side in origin|destination, field in postal|country.
    Priority: 1) Service (excluding SPECIAL/EXP_DUTY), 2) Priority 2 geo business rules in order:
       destination postal, destination country, origin postal, origin country.
    """
    parsed = _parse_differences_list(differences_list)

    # Exclude SPECIAL (and any other excluded types) from service candidates entirely
    excluded_norm = {_normalize(s) for s in EXCLUDED_SERVICE_FOR_BEST_MATCH}
    service_candidates = [
        (lane_num, rc_val, sv)
        for lane_num, rc_val, sv in parsed["service"]
        if _normalize(rc_val) not in excluded_norm
    ]

    best = _best_match_service(service_candidates)
    if best is not None:
        return (best, "service")

    for side, field in _GEO_TRY_ORDER:
        bucket = [
            (ln, rn, val)
            for ln, _col, rn, s, f, val in parsed["geo"]
            if s == side and f == field
        ]
        best = _best_match_tied_geo(bucket)
        if best is not None:
            return (best, ("geo", side, field))

    # Fallback: if only one lane appears in the differences, it is the best match
    lane_lines = _get_lane_differ_lines(differences_list)
    if len(lane_lines) == 1:
        lane_num, full_line = lane_lines[0]
        return (lane_num, "single")
    # Fallback: multiple tied lanes (e.g. all conditional differs) — show first as example
    if len(lane_lines) > 1:
        lane_num, _ = lane_lines[0]
        return (lane_num, "tied_first")

    return (None, None)


def _get_lane_differ_lines(differences_list):
    """Return list of (lane_num, full_line) for each 'Lane N: ...' line in differences_list."""
    result = []
    for line in differences_list or []:
        line = (line or "").strip()
        m = re.match(r"Lane\s+(\d+)\s*:\s*(.+)", line, re.IGNORECASE)
        if m:
            result.append((m.group(1).strip(), line))
    return result


# Pattern: "Column (Conditional) \"RuleVal\" -> shipment value 'ShipVal' differs"
_CONDITIONAL_DIFFER = re.compile(
    r"^.+\s*\(Conditional\)\s*\"([^\"]+)\"\s*->\s*shipment value\s+'([^']*)'\s*differs",
    re.IGNORECASE,
)


def _parse_conditional_differ(line):
    """
    If line is 'Lane N: Column (Conditional) "X" -> shipment value 'Y' differs',
    return (shipment_val, rule_val). Otherwise return None.
    """
    if not line or not isinstance(line, str):
        return None
    s = line.strip()
    m = re.match(r"^Lane\s+\d+\s*:\s*(.+)$", s, re.IGNORECASE)
    rest = m.group(1).strip() if m else s
    cond = _CONDITIONAL_DIFFER.search(rest)
    if cond:
        return (cond.group(2), cond.group(1))  # shipment_val, rule_val
    return None


def _build_combined_conditional_message(lane_lines):
    """
    If all lines are same conditional type with same shipment value, return
    Shipment Service Type value 'SHK' does not match rate card "Shenzhen", "Haiphong", "Shanghai"
    Otherwise return None.
    """
    if not lane_lines or len(lane_lines) < 2:
        return None
    parsed = []
    for _ln, full_line in lane_lines:
        p = _parse_conditional_differ(full_line)
        if p is None:
            return None
        parsed.append(p)
    shipment_vals = {p[0] for p in parsed}
    if len(shipment_vals) != 1:
        return None
    shipment_val = next(iter(shipment_vals))
    rule_vals = [p[1] for p in parsed]
    rules_str = ", ".join(f'"{r}"' for r in rule_vals)
    return f"Shipment Service Type value '{shipment_val}' does not match rate card {rules_str}"


def _find_lane_differ_line(differences_list, lane_num, section):
    """
    Return the first line in differences_list that is "Lane {lane_num}: ..." and matches the section.
    section: "service" | tuple ("geo", side, field)
    """
    prefix = f"Lane {lane_num}:"
    if section == "service":
        for line in differences_list or []:
            line = (line or "").strip()
            if not line.startswith(prefix):
                continue
            if "Service" in line and ("rate card" in line or "shipment value" in line):
                return line
    elif isinstance(section, tuple) and section[0] == "geo":
        _, side, field = section
        needle = f"shipment {side} {field}"
        for line in differences_list or []:
            line = (line or "").strip()
            if line.startswith(prefix) and needle in line.lower():
                return line
    # Fallback: first line for this lane
    for line in differences_list or []:
        line = (line or "").strip()
        if line.startswith(prefix):
            return line
    return None


def _build_combined_service_message(differences_list, lane_nums):
    """
    Build a combined message for multiple service lanes with same first-2-parts match.
    Example: "Lane 8492, 9184: Shipment service value 'STD_DIR_DTA' should be changed to 'STD_DIR_ATA'/'STD_DIR_ATD'"
    lane_nums: list of lane number strings
    Handles both original format (rate card 'X' -> shipment 'Y') and reformatted format.
    """
    if not lane_nums or len(lane_nums) < 2:
        return None
    
    # Patterns for both original and reformatted formats
    original_pattern = re.compile(
        r"rate card\s+'([^']*)'\s*->\s*shipment\s+'([^']*)'",
        re.IGNORECASE,
    )
    reformatted_pattern = re.compile(
        r"Shipment service value\s+'([^']*)'\s+should be changed to\s+'([^']*)'",
        re.IGNORECASE,
    )
    lane_data = []
    shipment_val = None
    
    for lane_num in lane_nums:
        full_line = _find_lane_differ_line(differences_list, lane_num, "service")
        if not full_line:
            continue
        
        # Try reformatted format first (if already reformatted)
        m = reformatted_pattern.search(full_line)
        if m:
            sv, rc = m.group(1), m.group(2)
        else:
            # Try original format
            m = original_pattern.search(full_line)
            if m:
                rc, sv = m.group(1), m.group(2)  # Note: original format has rc first, sv second
            else:
                continue
        
        if shipment_val is None:
            shipment_val = sv
        elif shipment_val != sv:
            # Different shipment values - can't combine
            return None
        lane_data.append((lane_num, rc))
    
    if not lane_data or shipment_val is None:
        return None
    
    # All lanes have same shipment value; combine rate card values
    rate_card_vals = [rc for _ln, rc in lane_data]
    lanes_str = ", ".join(ln for ln, _rc in lane_data)
    rc_vals_str = "/".join(f"'{rc}'" for rc in rate_card_vals)
    return f"Lane {lanes_str}: Shipment service value '{shipment_val}' should be changed to {rc_vals_str}"


def _find_all_geo_lanes_same_type(differences_list, geo_section):
    """
    Find all lanes in differences_list that have the same geo differ type (same side, same field).
    Returns list of (lane_num, rule_name) tuples, or None if not found.
    """
    if not differences_list:
        return None
    _, side, field = geo_section
    
    # Pattern to match reformatted geo: "Shipment Origin Country value 'CN' does not match rate card 'AT'"
    side_pattern = "Origin" if side == "origin" else "Destination"
    field_pattern = "Country" if field == "country" else "(?:Postal|postal)"
    pattern = re.compile(
        rf"Lane\s+(\d+)\s*:\s*Shipment\s+{side_pattern}\s+{field_pattern}\s+(?:value\s+)?'([^']*)'\s+does not match\s+rate card\s+'([^']*)'",
        re.IGNORECASE,
    )
    
    lanes = []
    shipment_val = None
    
    for line in differences_list:
        line = (line or "").strip()
        m = pattern.search(line)
        if m:
            lane_num = m.group(1)
            sv = m.group(2)
            rc = m.group(3)
            if shipment_val is None:
                shipment_val = sv
            elif shipment_val != sv:
                # Different shipment values - can't combine
                return None
            lanes.append((lane_num, rc))
    
    return lanes if len(lanes) > 1 else None


def _build_combined_geo_message(differences_list, lane_rule_pairs, geo_section):
    """
    Build a combined message for multiple tied geo (postal or country) lanes.
    lane_rule_pairs: list of (lane_num, rule_name) tuples
    geo_section: ("geo", side, field) with field in postal|country
    Handles both original and reformatted formats.
    """
    if not lane_rule_pairs:
        return None
    _, side, field = geo_section
    
    # Patterns for reformatted format: "Shipment origin country 'CN' does not match rate card 'AT'"
    # or "Shipment origin country 'CN' does not match the origin \"AT\""
    reformatted_pattern_rate_card = re.compile(
        r"Shipment\s+(?:origin|destination)\s+(?:country|postal)\s+(?:value\s+)?'([^']*)'\s+does not match\s+rate card\s+'([^']*)'",
        re.IGNORECASE,
    )
    reformatted_pattern_rule = re.compile(
        r"Shipment\s+(?:origin|destination)\s+(?:country|postal)\s+(?:value\s+)?'([^']*)'\s+does not match\s+(?:the\s+(?:origin|destination|Origin|Destination)\s+)?\"([^\"]+)\"",
        re.IGNORECASE,
    )
    
    # Original format: "-> shipment origin country 'X' differs" with rule name from lane_rule_pairs
    original_pattern = re.compile(
        r"->\s*shipment\s+(?:origin|destination)\s+(?:country|postal)\s+'([^']*)'",
        re.IGNORECASE,
    )
    
    lane_data = []
    shipment_val = None
    
    for lane_num, rule_name in lane_rule_pairs:
        full_line = _find_lane_differ_line(differences_list, lane_num, geo_section)
        if not full_line:
            continue
        
        # Try reformatted format with "rate card"
        m = reformatted_pattern_rate_card.search(full_line)
        if m:
            sv, rc = m.group(1), m.group(2)
        else:
            # Try reformatted format with rule name in quotes
            m = reformatted_pattern_rule.search(full_line)
            if m:
                sv, rc = m.group(1), m.group(2)
            else:
                # Try original format - use rule_name from lane_rule_pairs
                m = original_pattern.search(full_line)
                if m:
                    sv = m.group(1)
                    rc = rule_name
                else:
                    continue
        
        if shipment_val is None:
            shipment_val = sv
        elif shipment_val != sv:
            # Different shipment values - can't combine
            return None
        lane_data.append((lane_num, rc))
    
    if not lane_data or shipment_val is None:
        return None
    
    # All lanes have same shipment value; combine rate card values
    rate_card_vals = [rc for _ln, rc in lane_data]
    if field == "postal":
        tag = "Shipment Origin postal" if side == "origin" else "Shipment Destination postal"
    else:
        tag = "Shipment Origin Country" if side == "origin" else "Shipment Destination Country"
    
    if len(rate_card_vals) == 1:
        return f"{tag} value '{shipment_val}' does not match rate card '{rate_card_vals[0]}'"
    
    rc_vals_str = "/".join(f"'{rc}'" for rc in rate_card_vals)
    return f"{tag} value '{shipment_val}' does not match rate card {rc_vals_str}"


def add_possible_best_match_column(rows):
    """Add 'Possible Best Match' to each row with the full differ comment. Modifies rows in place."""
    for row in rows:
        diff_list = row.get("differences_list") or []
        lane_info, section = compute_possible_best_match(diff_list)
        if lane_info is None:
            row["Possible Best Match"] = ""
        elif section == "single":
            # Only one lane in differences -> it is the best match; use its full line
            single_lanes = _get_lane_differ_lines(diff_list)
            row["Possible Best Match"] = single_lanes[0][1] if single_lanes else f"Lane {lane_info}"
        elif section == "tied_first":
            # Multiple tied lanes with same conditional + same shipment value -> combined message
            lane_lines = _get_lane_differ_lines(diff_list)
            combined = _build_combined_conditional_message(lane_lines) if lane_lines else None
            row["Possible Best Match"] = (
                combined if combined else (lane_lines[0][1] if lane_lines else f"Lane {lane_info}")
            )
        elif section == "service":
            if isinstance(lane_info, list) and len(lane_info) > 1:
                # Multiple lanes with same first-2-parts match
                combined = _build_combined_service_message(diff_list, lane_info)
                row["Possible Best Match"] = combined if combined else ""
            elif isinstance(lane_info, list) and len(lane_info) == 1:
                lane_num = lane_info[0]
                full_comment = _find_lane_differ_line(diff_list, lane_num, section)
                row["Possible Best Match"] = full_comment if full_comment else f"Lane {lane_num}"
            else:
                # Single lane (exact match or single first-2-parts match)
                lane_num = lane_info
                full_comment = _find_lane_differ_line(diff_list, lane_num, section)
                row["Possible Best Match"] = full_comment if full_comment else f"Lane {lane_num}"
        elif isinstance(section, tuple) and section[0] == "geo":
            if isinstance(lane_info, list) and len(lane_info) > 1:
                combined = _build_combined_geo_message(diff_list, lane_info, section)
                row["Possible Best Match"] = combined if combined else ""
            elif isinstance(lane_info, list) and len(lane_info) == 1:
                # Even if only one lane returned, check if there are more lanes with same geo differ
                lane_num = lane_info[0][0]
                # Try to find all lanes with same geo pattern
                all_geo_lanes = _find_all_geo_lanes_same_type(diff_list, section)
                if all_geo_lanes and len(all_geo_lanes) > 1:
                    combined = _build_combined_geo_message(diff_list, all_geo_lanes, section)
                    row["Possible Best Match"] = combined if combined else ""
                else:
                    full_comment = _find_lane_differ_line(diff_list, lane_num, section)
                    row["Possible Best Match"] = full_comment if full_comment else f"Lane {lane_num}"
            else:
                row["Possible Best Match"] = ""
        else:
            row["Possible Best Match"] = ""
    return rows


def _reformat_differ_line(line):
    """
    Transform a differ line to the new outlook.
    - Service: "Lane N: Service: rate card 'X' -> shipment 'Y' differs" -> "Lane N: Shipment service value 'Y' should be changed to 'X'"
    - Destination postal: "Lane N: Destination (Business Rule) \"R\" -> shipment destination postal 'Z' differs" -> "Lane N: Shipment destination postal 'Z' does not match the Destination \"R\""
    - Destination country: "Lane N: Destination (Business Rule) \"R\" -> shipment destination country 'C' differs" -> "Lane N: Shipment destination country 'C' does not match the destination \"R\""
    - Origin postal/country: analogous.
    - Other flat: "Lane N: Col: rate card 'X' -> shipment 'Y' differs" -> "Lane N: Shipment {col} value 'Y' does not match rate card 'X'"
    """
    if not line or not isinstance(line, str):
        return line
    s = line.strip()
    lane_prefix = ""
    m = re.match(r"^(Lane\s+\d+\s*:\s*)(.+)$", s, re.IGNORECASE)
    if m:
        lane_prefix = m.group(1)
        rest = m.group(2).strip()
    else:
        rest = s

    # Service: rate card 'X' -> shipment 'Y' differs
    service_m = re.search(r"Service\s*:\s*rate card\s+'([^']*)'\s*->\s*shipment\s+'([^']*)'\s*differs", rest, re.IGNORECASE)
    if service_m:
        return f"{lane_prefix}Shipment service value '{service_m.group(2)}' should be changed to '{service_m.group(1)}'"

    # Service (Conditional) "X" -> shipment value 'Y' differs
    service_cond_m = re.search(
        r'Service\s*\(Conditional\)\s*"([^"]+)"\s*->\s*shipment value\s+\'([^\']*)\'\s*differs',
        rest,
        re.IGNORECASE,
    )
    if service_cond_m:
        return f"{lane_prefix}Shipment service value '{service_cond_m.group(2)}' should be changed to '{service_cond_m.group(1)}'"

    # Any column (Business Rule) "R" -> shipment origin|destination postal|country 'V' differs
    # (Destination, Origin City, Origin Airport, etc.)
    gen_br_m = re.search(
        r"^(.+?)\s*\(Business Rule\)\s*\"([^\"]+)\"\s*->\s*shipment\s+(origin|destination)\s+(postal|country)\s+\'([^\']*)\'\s*differs",
        rest,
        re.IGNORECASE,
    )
    if gen_br_m:
        _col, rule_name, side, fld, val = gen_br_m.groups()
        side = side.lower()
        fld = fld.lower()
        if fld == "postal":
            loc = "Destination" if side == "destination" else "Origin"
            return f'{lane_prefix}Shipment {side} postal \'{val}\' does not match the {loc} "{rule_name}"'
        return f'{lane_prefix}Shipment {side} country \'{val}\' does not match the {side} "{rule_name}"'

    # Generic: "Column: rate card 'X' -> shipment 'Y' differs"
    generic_m = re.search(
        r"([^:]+)\s*:\s*rate card\s+'([^']*)'\s*->\s*shipment\s+'([^']*)'\s*differs",
        rest,
        re.IGNORECASE,
    )
    if generic_m:
        col = generic_m.group(1).strip()
        return f"{lane_prefix}Shipment {col} value '{generic_m.group(3)}' does not match rate card '{generic_m.group(2)}'"

    return line


def reformat_comments(rows):
    """
    Apply new comment outlook to differences_list, differences, and Possible Best Match.
    Call this AFTER add_possible_best_match_column.
    """
    for row in rows:
        diff_list = row.get("differences_list")
        if isinstance(diff_list, list):
            row["differences_list"] = [_reformat_differ_line(ln) for ln in diff_list]
            row["differences"] = "\n".join(row["differences_list"])
        elif row.get("differences") and isinstance(row["differences"], str):
            row["differences"] = "\n".join(
                _reformat_differ_line(ln) for ln in row["differences"].split("\n")
            )
        if row.get("Possible Best Match"):
            row["Possible Best Match"] = _reformat_differ_line(row["Possible Best Match"])
    return rows


def run_formatting(
    input_json_path=None,
    output_json_path=None,
    output_xlsx_path=None,
):
    """
    Read Matched_Shipments_with.json, add Possible Best Match column, write JSON and optional Excel.
    """
    try:
        script_dir = os.path.dirname(os.path.abspath(__file__))
    except NameError:
        script_dir = os.getcwd()
    partly_df = os.path.join(script_dir, "partly_df")

    input_json_path = input_json_path or os.path.join(
        partly_df, "Matched_Shipments_with.json"
    )
    if not os.path.exists(input_json_path):
        print(f"[ERROR] Input JSON not found: {input_json_path}")
        return None, None

    with open(input_json_path, "r", encoding="utf-8") as f:
        data = json.load(f)

    rows = data.get("matched_shipments", [])
    if not rows:
        print("[WARNING] No matched_shipments in JSON.")
        return input_json_path, None

    add_possible_best_match_column(rows)
    reformat_comments(rows)

    output_json_path = output_json_path or os.path.join(
        partly_df, "Matched_Shipments_formatted.json"
    )
    output_data = {
        "matched_shipments": rows,
    }
    with open(output_json_path, "w", encoding="utf-8") as f:
        json.dump(output_data, f, indent=2, ensure_ascii=False)
    print(f"Saved: {output_json_path}")

    if output_xlsx_path is not False:
        try:
            import pandas as pd
            output_xlsx_path = output_xlsx_path or os.path.join(
                partly_df, "Matched_Shipments_formatted.xlsx"
            )
            # Build Excel from list of dicts; differences_list may be a list - keep as string for Excel
            excel_rows = []
            for r in rows:
                row_copy = dict(r)
                if "differences_list" in row_copy and isinstance(
                    row_copy["differences_list"], list
                ):
                    row_copy["differences_list"] = "\n".join(
                        row_copy["differences_list"]
                    )
                excel_rows.append(row_copy)
            df = pd.DataFrame(excel_rows)
            with pd.ExcelWriter(output_xlsx_path, engine="openpyxl") as writer:
                df.to_excel(writer, sheet_name="Matched Shipments", index=False)
            print(f"Saved: {output_xlsx_path}")
        except Exception as e:
            print(f"[WARNING] Could not write Excel: {e}")

    return output_json_path, output_xlsx_path if output_xlsx_path else None


if __name__ == "__main__":
    run_formatting()
