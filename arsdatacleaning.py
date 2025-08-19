import csv
import re

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

    fieldnames = reader.fieldnames + ["CNfromPA"]
    writer = csv.DictWriter(outfile, fieldnames=fieldnames)
    writer.writeheader()

    for row in reader:
        # Split/clean Physical Availability → location + CNfromPA
        location, cn_from_pa = extract_from_physical_availability(row.get("Physical Availability", ""))
        row["Physical Availability"] = location
        row["CNfromPA"] = clean_call_number_keep_space_before_year(cn_from_pa)

        # Clean Local Call Number (2nd column)
        row["Local Call Number"] = clean_call_number_keep_space_before_year(row.get("Local Call Number", ""))

        # LC Call Number (3rd column): remove "and others" first, then clean
        lc_raw = row.get("LC Call Number", "")
        lc_no_others = re.sub(r"\band others\b", "", lc_raw, flags=re.IGNORECASE).strip()
        row["LC Call Number"] = clean_call_number_keep_space_before_year(lc_no_others)

        writer.writerow(row)

print(f"Read using '{used_encoding}'. Cleaned file written to {output_file}")
