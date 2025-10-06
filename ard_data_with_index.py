import csv
import re
from math import inf

def open_input_with_fallback(path):
    """Open text file trying several encodings; force a read to trigger decode."""
    for enc in ("utf-8-sig", "cp1252", "latin-1"):
        try:
            f = open(path, newline="", encoding=enc)
            _ = f.readline()   # force decode
            f.seek(0)
            return f, enc
        except UnicodeDecodeError:
            try:
                f.close()
            except Exception:
                pass
    raise UnicodeDecodeError("all-encodings", b"", 0, 1, "Could not decode with utf-8-sig/cp1252/latin-1")

def extract_from_physical_availability(text):
    """
    From 'Physical Availability' keep:
      - before 1st ';' as the cleaned Physical Availability (location)
      - between 2nd and 3rd ';' as CNfromPA
    Drops between 1st–2nd and after 3rd.
    """
    if not text:
        return "", ""
    s = text.strip()
    m = re.match(r'^([^;]*);[^;]*;([^;]*);.*$', s)
    if m:
        location = m.group(1).strip()
        cn_from_pa = m.group(2).strip()
        return location, cn_from_pa
    # Fallback if pattern not met
    return s, ""

def clean_call_number_keep_space_before_year(s):
    """
    Remove all spaces except the one immediately before a trailing 4-digit year.
    """
    if not s:
        return ""
    s = s.strip()
    s = re.sub(r"\s+", " ", s)          # normalize runs of whitespace
    s = re.sub(r" (?!\d{4}$)", "", s)   # remove spaces not before final 4 digits
    return s

# ---------- LC call number parsing to sortable key ----------

_cutter_re = re.compile(r"\.?\s*([A-Z])\s*([0-9]+)")
_year_re = re.compile(r"(?:^|[^0-9])([12][0-9]{3})(?![0-9])")  # 1000–2999, take last occurrence

def _parse_class_and_number(s):
    """
    Extract class letters (A–Z, 1–3 chars) and the main class number (int/float).
    Returns (letters, number_float, rest_string_start_index)
    """
    m = re.match(r"\s*([A-Z]{1,3})\s*([0-9]+(?:\.[0-9]+)?)?", s)
    if not m:
        return ("", inf, 0)
    letters = m.group(1) or ""
    num = float(m.group(2)) if m.group(2) else inf
    return (letters, num, m.end())

def _parse_cutters_and_extras(s):
    """
    Parse up to several cutters like .A123 .B45 and any trailing year/volume info.
    Return list of standardized cutter tuples and year (int or +inf).
    """
    cutters = []
    for m in _cutter_re.finditer(s):
        cutters.append( (m.group(1), int(m.group(2))) )
    # Year: take the last 4-digit year if present
    year = inf
    last = None
    for y in _year_re.finditer(s):
        last = y
    if last:
        year = int(last.group(1))
    return cutters, year

def lc_sort_key(raw):
    """
    Create a sortable tuple for an LC call number.
    Structure:
      (class_letters, class_number, cutter1_letter, cutter1_num, cutter2_letter, cutter2_num, cutter3_letter, cutter3_num, year)
    Missing parts are padded to keep tuple length consistent. Empty/invalid => sorts last.
    """
    if not raw:
        return ("ZZZ", inf, "", inf, "", inf, "", inf, inf)

    s = raw.strip().upper()
    # Normalize spaces/punctuation a bit
    s = re.sub(r"\s+", " ", s)
    s = s.replace("..", ".").replace(" .", ".").strip()

    letters, main_num, idx = _parse_class_and_number(s)
    rest = s[idx:]
    cutters, year = _parse_cutters_and_extras(rest)

    # Pad/trim cutters to 3 for tuple stability
    cutters = (cutters + [("", inf)] * 3)[:3]
    (c1l, c1n), (c2l, c2n), (c3l, c3n) = cutters

    return (letters or "ZZZ",
            main_num if main_num == main_num else inf,  # keep inf if missing
            c1l, c1n,
            c2l, c2n,
            c3l, c3n,
            year)

def best_call_number_for_sort(row):
    """
    Prefer LC Call Number; if empty, fall back to Local Call Number; then CNfromPA.
    """
    for k in ("LC Call Number", "Local Call Number", "CNfromPA"):
        v = (row.get(k) or "").strip()
        if v:
            return v
    return ""

# ------------------- Main script -------------------

# Prompt user for file paths
input_file = input("Enter input CSV file path: ").strip()
output_file = input("Enter output CSV file path: ").strip()

# Open input with fallback encoding
infile, used_encoding = open_input_with_fallback(input_file)

with infile, open(output_file, "w", newline="", encoding="utf-8-sig") as outfile:
    reader = csv.DictReader(infile)

    # Ensure expected columns exist
    required = ["Title", "Local Call Number", "LC Call Number", "Physical Availability"]
    for col in required:
        if col not in reader.fieldnames:
            raise ValueError(f"Missing expected column: {col}")

    # New column: LC_SortIndex (Excel-friendly)
    fieldnames = reader.fieldnames + ["CNfromPA", "LC_SortIndex"]
    writer = csv.DictWriter(outfile, fieldnames=fieldnames)
    writer.writeheader()

    # First pass: clean rows and keep them in memory
    cleaned_rows = []
    row_id = 0
    for row in reader:
        row_id += 1
        row["_row_id"] = row_id  # internal ID to map back indices

        # Split/clean Physical Availability → location + CNfromPA
        location, cn_from_pa = extract_from_physical_availability(row.get("Physical Availability", ""))
        row["Physical Availability"] = location
        row["CNfromPA"] = clean_call_number_keep_space_before_year(cn_from_pa)

        # Clean Local Call Number
        row["Local Call Number"] = clean_call_number_keep_space_before_year(row.get("Local Call Number", ""))

        # LC Call Number: remove "and others" first, then clean
        lc_raw = row.get("LC Call Number", "")
        lc_no_others = re.sub(r"\band others\b", "", lc_raw, flags=re.IGNORECASE).strip()
        row["LC Call Number"] = clean_call_number_keep_space_before_year(lc_no_others)

        # compute sort key on the *cleaned* values
        call_for_sort = best_call_number_for_sort(row)
        row["_lc_sort_key"] = lc_sort_key(call_for_sort)

        cleaned_rows.append(row)

    # Second pass: compute stable LC_SortIndex by sorting a copy with the key
    sorted_rows = sorted(cleaned_rows, key=lambda r: r["_lc_sort_key"])

    # Assign 1-based index following true LC sort order
    id_to_index = {r["_row_id"]: i + 1 for i, r in enumerate(sorted_rows)}

    # Final write: preserve original order but include LC_SortIndex for Excel sorting
    for row in cleaned_rows:
        row["LC_SortIndex"] = id_to_index[row["_row_id"]]
        # remove internals
        row.pop("_row_id", None)
        row.pop("_lc_sort_key", None)
        writer.writerow(row)

print(f"Read using '{used_encoding}'. Cleaned file written to {output_file}")
print("Added LC_SortIndex: sort by this column in Excel to get proper LC order.")
