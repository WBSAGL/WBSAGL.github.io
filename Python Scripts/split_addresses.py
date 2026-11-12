"""
Split addresses into Local Address 1-4, Local Postcode and Local Country (or
Address 1-4, Postcode and Country - see --target) for database upload, and
write a SQL UPDATE statement per row.

Input
-----
Either of these layouts is detected automatically:
* One long address string in an 'Address' column: it is split up using the rules below.
* Separate address columns ('Address1', 'Address2', ... or 'Local Address 1', ...):
  rows whose lines already follow the rules are kept as entered (commas and
  apostrophes removed, town moved to line 4). A row is reworked with the rules
  below if a line is over 30 characters, there are more than 4 lines, a line
  contains a postcode or country, a line is repeated, or something is split
  across lines (a house number on its own line, "Flat" | "3", "Dushi" | "Road").
  Split Notes says "Kept as entered" or why the row was reworked.

Rules
-----
(These describe the local address fields; the address fields work the same way.)
* Local Address 1-4 are VARCHAR(30); Local Postcode is VARCHAR(10).
* Local Address 4 holds the City/Town where available: from the City/Town
  column if the file has one, otherwise the last piece of the address when it
  looks like a town (no digits, 30 characters or fewer). For an overseas
  address, Local Address 4 is "Town Region" when it fits, e.g. "Nanjing Jiangsu".
* Local Country comes from the Country column if the file has one, otherwise
  from a recognised country at the end of the address ("China", "United Kingdom"),
  which is then taken out of the address lines.
* Repeated pieces ("Shanghai, Shanghai, China") are only kept once.
* Every word of each address line (including the town) is given a capital
  first letter; the rest of each word is left as typed, and words starting
  with a number are unchanged ("12 high street" -> "12 High Street").
* City/Town, Postcode and Country columns are recognised from many headings
  ("City", "Town", "Post Code", "Postal Code", "Postcode:", "Zip", ...), ignoring
  case, spaces and punctuation. The Postcode column is absorbed into the target
  postcode field (Local Postcode by default, Postcode with --target address) and
  is not copied to the output separately.
* The output contains no commas or apostrophes/quotes of any kind (they cause
  problems in SQL) - including Chinese/Japanese commas (， 、), the Arabic comma (،)
  and curly or accented apostrophes. Every kind of comma also separates address
  pieces. Pieces sharing a field are separated by a space, e.g. "Nanjing Jiangsu",
  and "O'Brien" becomes "OBrien".
* The City/Town, Postcode and Country columns are optional. Only Address is required.
* The address is split on commas and the pieces are spread as evenly as possible
  across Local Address 1-3 (or 1-4 when there is no City/Town), never cutting a
  piece in half unless that piece alone is longer than 30 characters.
* No Local Address field ever exceeds 30 characters. When the address doesn't
  fit that way, it is packed in these steps, stopping at the first that fits
  (see pack_lines): spread over all four fields with the City/Town at the end
  of Local Address 4; then also splitting pieces between words; then every
  field filled to 30 characters. Only text that still doesn't fit goes into
  the 'Overspill' column, and only then is the Overspill cell flagged amber.
  Split Notes says which step was used.
* The Postcode column is copied (tidied) into Local Postcode. If there is no
  Postcode column, or it is blank, a postcode found in the Address is used:
  a UK postcode, or a 5-6 digit overseas one ("52200", "China(310000)",
  "Pin-248001"). A postcode longer than 10 characters is left blank in Local
  Postcode and written to Overspill (Overspill cell flagged amber).
* Postcodes typed inside the Address text are removed from the address lines
  (use --keep-postcodes to leave them in). A UK postcode only counts if it
  starts with a real UK postcode area and ends with valid letters, so room
  codes like "D306 IQ" are left in the address.

Output
------
* An Excel file with the original columns, except Address, City/Town, Postcode
  and Country, which are replaced by Local Address 1-4, Local Postcode and
  Local Country. An 'SQL Update' column holds the UPDATE statement for each row.
* The output is the Excel file only - no separate .sql file is written.
* The UPDATE matches on Regnumber (or a 'Student Number...' column) and, if the
  file has an Acronym column, Acronym too. Blank values are written as NULL.
* If the file has an 'Email', 'MobilePhone' and/or 'Usedname' column, the UPDATE
  also sets c.email / c.mobile_phone / c.usedname. A field whose column isn't in
  the file is not updated. A blank Usedname cell leaves c.usedname unchanged
  (it is left out of that row's statement rather than set to NULL). A row with no address but with one of these values
  only updates those fields (and is flagged for review).
* The Regnumber is used exactly as in the file (never trimmed). A Regnumber
  longer than 7 characters - including any spaces - is highlighted red (that
  cell only) and Split Notes says so.
* Overspill highlights only the Overspill cell amber, not the whole row.
* A mobile number may only contain digits (with an optional + at the start);
  spaces, dashes, dots and brackets are removed. Anything else ("Not yet",
  "N/A") is set to NULL.

Usage
-----
    python split_addresses.py                        # uses the defaults below
    python split_addresses.py input.xlsx output.xlsx                     # updates localaddress fields
    python split_addresses.py input.xlsx output.xlsx --target address    # updates address fields
    python split_addresses.py --keep-postcodes       # leave postcodes in the address lines
"""

# ---------------------------------------------------------------------------
# Libraries
# ---------------------------------------------------------------------------
import argparse            # reads the input/output file names typed after the script name
import re                  # "regular expressions": pattern matching, used to spot postcodes
from pathlib import Path   # handles file paths (e.g. adding .xlsx to the output name)

import openpyxl            # reads and writes Excel .xlsx files
from openpyxl.styles import Font, PatternFill

# ---------------------------------------------------------------------------
# Settings - change these if the database or input files change
# ---------------------------------------------------------------------------
MAX_LEN = 30            # each address line: VARCHAR(30)
POSTCODE_MAX_LEN = 10   # postcode: VARCHAR(10)

# Which set of database fields to update, chosen with --target when running.
# "columns" are the output column headings (address lines 1-4, postcode, country);
# "sql" are the matching database field names used in the UPDATE statement.
TARGETS = {
    "localaddress": {
        "columns": ["Local Address 1", "Local Address 2", "Local Address 3", "Local Address 4",
                    "Local Postcode", "Local Country"],
        "sql": ["localaddress1", "localaddress2", "localaddress3", "localaddress4",
                "localpostcode", "localcountry"],
    },
    "address": {
        "columns": ["Address 1", "Address 2", "Address 3", "Address 4", "Postcode", "Country"],
        "sql": ["address1", "address2", "address3", "address4", "postcode", "country"],
    },
}
DEFAULT_TARGET = "localaddress"
# Columns added after the address/postcode/country columns, in this order
EXTRA_COLUMNS = ["Overspill", "Split Notes", "SQL Update"]

# Used when the script is run without file names
DEFAULT_INPUT = "DSSSPG-6744 - Anoymised Data.xlsx"
DEFAULT_OUTPUT = "DSSSPG-6744 - Split Addresses.xlsx"

# Column headings looked for in the input file (not case-sensitive).
# "PREFIX" means the heading only has to start with this text, e.g.
# "City/Town (e.g. Coventry)" matches "City/Town".
ADDRESS_HEADER = "Address"            # one long address string...
# ...or numbered address columns: "Address1", "Address 2", "Address Line 3",
# "Local Address 4", "LocalAddress1" etc. If the file has any of these, they are used instead.
ADDRESS_LINE_HEADER_RE = re.compile(r"^(local\s*)?address\s*(line\s*)?(\d+)$", re.IGNORECASE)
# City/Town, Postcode and Country headings are matched loosely: case, spaces and
# punctuation are ignored, and anything after the heading is allowed, so e.g.
# "Post Code", "Postcode:", "Postal code (e.g. CV4 7AL)" all count as Postcode.
# Each list is tried in order; "Local ..." headings are preferred when updating
# localaddress fields, the plain ones when updating address fields.
CITY_HEADINGS = ["City/Town", "Town/City", "City", "Town"]
POSTCODE_HEADINGS = ["Postcode", "Post Code", "Postal Code", "Zip Code", "Zip"]
COUNTRY_HEADINGS = ["Country"]
REGNUMBER_HEADER_PREFIXES = ["Regnumber", "Student Number"]  # first one found is used
ACRONYM_HEADER = "Acronym"

# Optional contact columns: (input column heading, database field). If a column
# is in the input file, the UPDATE statement also sets that field. Headings are
# matched ignoring case, spaces and underscores ("MobilePhone" = "Mobile Phone").
CONTACT_COLUMNS = [("Email", "email"), ("MobilePhone", "mobile_phone"), ("Usedname", "usedname")]
MOBILE_HEADING = "MobilePhone"   # this column gets the mobile number checks (clean_mobile)
# Fields left out of a row's statement when that row's cell is blank (so the
# database value is left unchanged instead of being set to NULL)
SKIP_IF_BLANK = {"usedname"}

REGNUMBER_MAX_LEN = 7   # a Regnumber longer than this (including any spaces) is highlighted red

AMBER_FILL = PatternFill("solid", fgColor="FFC000")  # flags a row (or the Overspill cell) for review
RED_FILL = PatternFill("solid", fgColor="FF0000")    # flags a Regnumber over 7 characters

# ---------------------------------------------------------------------------
# Postcode patterns
# ---------------------------------------------------------------------------
# UK postcode shape, tolerant of missing/extra spaces (e.g. "CV4 8GL", "CV48GL", "CV1 1 AQ").
# The shape alone isn't enough - room codes like "D306 IQ" have it too - so
# is_uk_postcode also checks the area letters and the final two letters.
POSTCODE_RE = re.compile(r"^([A-Z]{1,2})\d[A-Z\d]?\s*\d\s*([A-Z]{2})$", re.IGNORECASE)
# Postcode at the end of a piece, e.g. "Coventry CV4 8GL"
TRAILING_POSTCODE_RE = re.compile(r"^(.*?)\s+([A-Z]{1,2}\d[A-Z\d]?\s*\d\s*[A-Z]{2})$", re.IGNORECASE)
# The real UK postcode areas (the letters at the start, e.g. CV = Coventry, B = Birmingham)
UK_POSTCODE_AREAS = set("""
AB AL B BA BB BD BF BH BL BN BR BS BT CA CB CF CH CM CO CR CT CV CW DA DD DE DG DH DL DN DT DY
E EC EH EN EX FK FY G GL GU GY HA HD HG HP HR HS HU HX IG IM IP IV JE KA KT KW KY L LA LD LE LL
LN LS LU M ME MK ML N NE NG NN NP NR NW OL OX PA PE PH PL PO PR RG RH RM S SA SE SG SK SL SM SN
SO SP SR SS ST SW SY TA TD TF TN TQ TR TS TW UB W WA WC WD WF WN WR WS WV YO ZE
""".split())
# The last two letters of a UK postcode never use C, I, K, M, O or V
UK_INWARD_LETTERS = set("ABDEFGHJLNPQRSTUWXYZ")
# Numeric (overseas) postcode as a piece of its own, e.g. "52200", "Pin-248001"
NUMERIC_POSTCODE_RE = re.compile(r"^(?:pin\s*-?\s*)?(\d{5,6})$", re.IGNORECASE)
# Numeric postcode after a place name, e.g. "China(310000)", "Vietnam - 71107", "Shah Alam 40170"
TRAILING_NUMERIC_POSTCODE_RE = re.compile(r"^([^\d]+?)(?:\s+|\s*-\s*|\s*\(\s*)(\d{5,6})\)?$")
# ...but not when the "place name" is really a building/room label, e.g. "Building 12211"
UNIT_WORDS = {"building", "bldg", "block", "blk", "room", "rm", "unit", "flat", "apartment", "apt",
              "no", "no.", "house", "tower", "floor", "phase", "plot", "sector", "zone", "lane", "road", "street"}
# Postcode at the start of a place name, e.g. "81100 Johor Bahru"
LEADING_NUMERIC_POSTCODE_RE = re.compile(r"^(\d{5,6})\s+([^\d]+)$")

# ---------------------------------------------------------------------------
# Words that should stay with their neighbours when a line has to be split
# ---------------------------------------------------------------------------
# Street types that come AFTER the name: never start a new line with one,
# e.g. never "Dushi" | "Road". Includes Chinese (lu, jie, dao), Indian (marg,
# nagar, salai) and Korean (ro, gil, gu, dong) words.
ROAD_SUFFIXES = {"road", "rd", "street", "st", "lane", "ln", "avenue", "ave", "way", "drive", "dr",
                 "court", "ct", "close", "crescent", "place", "terrace", "boulevard", "blvd", "highway",
                 "lu", "jie", "dao", "marg", "nagar", "salai", "path", "ro", "gil", "gu", "dong"}
# Street types that come BEFORE the name: never end a line with one,
# e.g. never "Calle" | "Mayor", "Jalan" | "Sudirman", "Rue" | "de Rivoli"
STREET_PREFIXES = {"calle", "via", "viale", "rue", "avenida", "av", "av.", "rua", "jalan", "jln", "jl", "jl.",
                   "strada", "ulica", "ul", "ul.", "ulitsa", "soi", "thanon", "carrer", "camino", "paseo",
                   "rambla", "piazza", "largo", "travessa", "avenue", "boulevard", "chaussee", "chemin"}
# Labels followed by a number: never end a line with one, e.g. never "Lane" | "5058"
NUMBER_LABELS = {"room", "rm", "no", "no.", "nr", "nr.", "unit", "flat", "building", "bldg", "block", "blk",
                 "lane", "floor", "tower", "phase", "plot", "apartment", "apt", "apto", "piso", "house",
                 "sector", "suite", "lot", "villa", "kav", "kav."}
# A house number on its own, e.g. "12" or "49a"
HOUSE_NUMBER_RE = re.compile(r"^\d+[A-Za-z]?$")
# A piece starting with one of these is the flat/room itself. If it has been
# written at the end of the address it is moved to the front (Local Address 1)
ROOM_WORDS = {"room", "rm", "flat", "apartment", "apt", "suite", "unit"}

# ---------------------------------------------------------------------------
# Country names recognised at the end of an address
# ---------------------------------------------------------------------------
# UK names (a UK address keeps just the town in Local Address 4, not "Town Region")
UK_COUNTRIES = {"united kingdom", "uk", "u.k.", "great britain", "gb", "england", "scotland", "wales",
                "northern ireland"}
# Overseas countries and territories (compared in lower case without any "(...)"
# suffix, so "Vietnam (viet Nam)" matches). Apostrophes are already removed by
# sql_safe, so "Cote d'Ivoire" is listed as "cote divoire".
WORLD_COUNTRIES = {
    "afghanistan", "albania", "algeria", "andorra", "angola", "antigua and barbuda", "argentina", "armenia",
    "australia", "austria", "azerbaijan", "bahamas", "bahrain", "bangladesh", "barbados", "belarus",
    "belgium", "belize", "benin", "bermuda", "bhutan", "bolivia", "bosnia and herzegovina", "botswana",
    "brazil", "brunei", "brunei darussalam", "bulgaria", "burkina faso", "burundi", "cambodia", "cameroon",
    "canada", "cape verde", "cayman islands", "central african republic", "chad", "chile", "china",
    "colombia", "comoros", "congo", "costa rica", "cote divoire", "ivory coast", "croatia", "cuba",
    "curacao", "cyprus", "czech republic", "czechia", "democratic republic of the congo", "denmark",
    "djibouti", "dominica", "dominican republic", "ecuador", "egypt", "el salvador", "equatorial guinea",
    "eritrea", "estonia", "eswatini", "swaziland", "ethiopia", "fiji", "finland", "france", "gabon",
    "gambia", "georgia", "germany", "ghana", "gibraltar", "greece", "greenland", "grenada", "guatemala",
    "guinea", "guinea-bissau", "guyana", "haiti", "honduras", "hong kong", "hungary", "iceland", "india",
    "indonesia", "iran", "iraq", "ireland", "republic of ireland", "eire", "israel", "italy", "jamaica",
    "japan", "jordan", "kazakhstan", "kenya", "kiribati", "korea", "south korea", "republic of korea",
    "north korea", "kosovo", "kuwait", "kyrgyzstan", "laos", "latvia", "lebanon", "lesotho", "liberia",
    "libya", "liechtenstein", "lithuania", "luxembourg", "macao", "macau", "madagascar", "malawi",
    "malaysia", "maldives", "mali", "malta", "marshall islands", "mauritania", "mauritius", "mexico",
    "micronesia", "moldova", "monaco", "mongolia", "montenegro", "morocco", "mozambique", "myanmar", "burma",
    "namibia", "nauru", "nepal", "netherlands", "the netherlands", "holland", "new zealand", "nicaragua",
    "niger", "nigeria", "north macedonia", "macedonia", "norway", "oman", "pakistan", "palau", "palestine",
    "panama", "papua new guinea", "paraguay", "peru", "philippines", "poland", "portugal", "puerto rico",
    "qatar", "romania", "russia", "russian federation", "rwanda", "saint lucia", "samoa", "san marino",
    "saudi arabia", "senegal", "serbia", "seychelles", "sierra leone", "singapore", "slovakia",
    "slovak republic", "slovenia", "solomon islands", "somalia", "south africa", "south sudan", "spain",
    "sri lanka", "sudan", "suriname", "sweden", "switzerland", "syria", "taiwan", "tajikistan", "tanzania",
    "thailand", "timor-leste", "east timor", "togo", "tonga", "trinidad and tobago", "tunisia", "turkey",
    "turkiye", "turkmenistan", "tuvalu", "uganda", "ukraine", "united arab emirates", "uae",
    "united states", "united states of america", "usa", "u.s.a.", "u.s.", "us", "uruguay", "uzbekistan",
    "vanuatu", "vatican city", "venezuela", "vietnam", "viet nam", "yemen", "zambia", "zimbabwe",
    "peoples republic of china", "pr china", "p.r. china", "prc", "mainland china",
}
# Other names for the same country, so the right postcode format is used
COUNTRY_ALIASES = {
    "usa": "united states", "u.s.a.": "united states", "u.s.": "united states", "us": "united states",
    "united states of america": "united states", "the netherlands": "netherlands", "holland": "netherlands",
    "republic of ireland": "ireland", "eire": "ireland", "czechia": "czech republic",
    "slovak republic": "slovakia", "republic of korea": "south korea", "korea": "south korea",
    "russian federation": "russia", "turkiye": "turkey", "viet nam": "vietnam", "macedonia": "north macedonia",
    "peoples republic of china": "china", "pr china": "china", "p.r. china": "china", "prc": "china",
    "mainland china": "china",
}

# ---------------------------------------------------------------------------
# Postcode formats by country
# ---------------------------------------------------------------------------
# Used once the address's country is known (from the Country column or the
# address). Many countries use plain 4-digit postcodes, which look just like
# house or room numbers - knowing the country is what makes them safe to spot.
# (5-6 digit postcodes are spotted for every country - see NUMERIC_POSTCODE_RE.)
FOUR_DIGITS = r"\d{4}"
COUNTRY_POSTCODES = {
    **{country: FOUR_DIGITS for country in (
        "australia", "new zealand", "switzerland", "austria", "belgium", "denmark", "norway", "hungary",
        "philippines", "south africa", "bangladesh", "luxembourg", "cyprus", "georgia", "armenia", "tunisia",
        "bulgaria", "slovenia", "north macedonia", "liechtenstein", "afghanistan", "paraguay", "venezuela",
        "albania", "mozambique", "ethiopia", "latvia", "argentina", "greenland", "eswatini", "swaziland")},
    "netherlands": r"\d{4}\s?[A-Z]{2}",                                   # 1015 CJ
    "canada": r"[A-Z]\d[A-Z]\s?\d[A-Z]\d",                                 # K1N 5W8
    "portugal": r"\d{4}-\d{3}",                                            # 1100-053
    "poland": r"\d{2}-\d{3}",                                              # 00-590
    "japan": r"\d{3}-\d{4}",                                               # 131-0045
    "brazil": r"\d{5}-\d{3}",                                              # 01310-100
    "united states": r"\d{5}(?:-\d{4})?",                                  # 20500 or 10118-0110
    "sweden": r"\d{3}\s?\d{2}",                                            # 111 51
    "czech republic": r"\d{3}\s?\d{2}",
    "slovakia": r"\d{3}\s?\d{2}",
    "greece": r"\d{3}\s?\d{2}",
    "ireland": r"(?:[AC-FHKNPRTV-Y]\d{2}|D6W)\s?[0-9AC-FHKNPRTV-Y]{4}",   # Eircode, e.g. D02 X285
    "malta": r"[A-Z]{3}\s?\d{4}",                                          # VLT 1117
    "iran": r"\d{5}-?\d{5}",
}


# ===========================================================================
# STEP 1 - Clean the address and pull out the postcode
# ===========================================================================

# Every kind of comma: normal (,), full-width Chinese/Japanese (，), ideographic (、),
# small (﹐ ﹑), Arabic (،) and the low quote that looks like one (‚)
COMMA_CHARS = ",，、﹐﹑،‚"
# Every kind of apostrophe or quote: ' ‘ ’ ` " “ ” ´ ′ ″ ʼ ʻ ＇ ＂ ‛ „ (and the low quote above)
QUOTE_CHARS = "'‘’`\"“”´′″ʼʻ＇＂‛„"
COMMA_RE = re.compile("[" + re.escape(COMMA_CHARS) + "]")
QUOTE_RE = re.compile("[" + re.escape(QUOTE_CHARS) + "]")


def capitalise_words(line):
    """
    Give every word of an address line (or town) a capital first letter. Only
    the first letter of each word changes - the rest is left as typed, and
    words starting with a number are left alone:
        "12 high street" -> "12 High Street"     "leamington spa" -> "Leamington Spa"
        "CODE Student"   -> "CODE Student"       "2nd floor"      -> "2nd Floor"
    """
    return " ".join(word[0].upper() + word[1:] if word[:1].isalpha() else word for word in line.split(" "))


def normalise_commas(text):
    """Turn every kind of comma into a normal comma, e.g. 'Room 5，Block 2' -> 'Room 5,Block 2'."""
    return COMMA_RE.sub(",", str(text or ""))


def sql_safe(text):
    """
    Remove every kind of comma and apostrophe/quote, which cause problems in SQL:
      * a comma becomes a space, so words either side stay apart: "Room 5，Block 2" -> "Room 5 Block 2"
      * an apostrophe is simply removed: "St Mary's" -> "St Marys"
    Then runs of spaces are collapsed into one.
    """
    text = QUOTE_RE.sub("", COMMA_RE.sub(" ", str(text or "")))
    return " ".join(text.split())


def is_uk_postcode(text):
    """
    True if the text is a real UK postcode: the right shape, a real postcode
    area at the start, and valid final letters.
        "CV4 8GL" -> True      "D306 IQ" -> False (no "D" area; "I" never used at the end)
    """
    match = POSTCODE_RE.match(text.strip())
    if not match:
        return False
    area, last_two = match.group(1).upper(), match.group(2).upper()
    return area in UK_POSTCODE_AREAS and set(last_two) <= UK_INWARD_LETTERS


def trailing_uk_postcode(part):
    """For a piece ending in a real UK postcode, e.g. "Coventry CV4 8GL", return the match, else None."""
    match = TRAILING_POSTCODE_RE.match(part)
    return match if match and is_uk_postcode(match.group(2)) else None


def split_numeric_postcode(part):
    """Return (rest_of_piece, postcode) for an overseas numeric postcode, or (part, '') if none."""
    # Whole piece is a postcode, e.g. "52200" -> nothing left, postcode "52200"
    match = NUMERIC_POSTCODE_RE.match(part)
    if match:
        return "", match.group(1)
    # Postcode on the end of a place name, e.g. "China(310000)" -> "China", "310000"
    match = TRAILING_NUMERIC_POSTCODE_RE.match(part)
    if match and match.group(1).strip().lower() not in UNIT_WORDS:
        return match.group(1).strip(" -("), match.group(2)
    # Postcode at the start of a place name, e.g. "81100 Johor Bahru" -> "Johor Bahru", "81100".
    # Not when the rest is a street ("12345 Long Road" is a house number, not a postcode).
    match = LEADING_NUMERIC_POSTCODE_RE.match(part)
    if match:
        words = match.group(2).lower().split()
        if words[0] not in UNIT_WORDS and words[-1] not in ROAD_SUFFIXES:
            return match.group(2), match.group(1)
    return part, ""


def clean_parts(address, city, strip_postcodes):
    """
    Split an address on commas, tidy whitespace and drop empty/duplicate pieces.
    Returns (parts, postcode) where postcode is the last postcode found in the text.

    e.g. "Flat 3, 12 High Street, , Kenilworth, CV8 1LZ"
         -> (["Flat 3", "12 High Street", "Kenilworth"], "CV8 1LZ")
    """
    parts, found_postcode = [], ""

    # Go through the address one comma-separated piece at a time
    for part in normalise_commas(address).replace("\n", ",").split(","):
        part = sql_safe(part)
        if not part:
            continue  # skip empty pieces such as the gap in ", ,"

        # Is this piece (or the end of it) a postcode? If so, remember it and
        # take it out of the address so it isn't repeated in the address lines.
        if is_uk_postcode(part):                               # e.g. "CV8 1LZ"
            found_postcode = part
            if strip_postcodes:
                continue
        else:
            trailing = trailing_uk_postcode(part)              # e.g. "Leamington Spa CV32 4RR"
            if trailing:
                found_postcode = trailing.group(2)
                if strip_postcodes:
                    part = trailing.group(1)
            else:
                rest, numeric = split_numeric_postcode(part)   # e.g. "52200", "China(310000)"
                if numeric:
                    found_postcode = numeric
                    if strip_postcodes:
                        part = rest
                        if not part:
                            continue

        # Shorten country names with a bracketed extra, e.g. "Vietnam (viet Nam)" -> "Vietnam"
        if country_name(part):
            part = re.sub(r"\s*\(.*\)$", "", part)

        # If the City/Town column was given and the address repeats it, drop it
        # here - the City/Town is written to Local Address 4 separately
        if city and part.lower() == city.lower():
            continue
        parts.append(part)

    # Drop repeated pieces ("Shanghai, Shanghai, China"), keeping the last one so a
    # trailing town/country stays at the end. "Guangdong Province" counts as a
    # repeat of "Guangdong" (see place_key).
    seen, unique = set(), []
    for part in reversed(parts):
        if place_key(part) not in seen:
            seen.add(place_key(part))
            unique.append(part)
    unique.reverse()
    return unique, found_postcode


def join_fragments(parts):
    """
    Join fragments that can't stand on their own to their neighbour, so they
    can't end up on separate lines:
      * a house number or label on its own joins the next piece:
            "12, High Street" -> "12 High Street",   "Flat, 3 Station Road" -> "Flat 3 Station Road"
        ...except before the last piece (usually the town), where it joins the
        previous piece instead: "Calle Mayor 10, 3B, Madrid" -> "Calle Mayor 10 3B", "Madrid"
      * a street word on its own joins its name:
            "Dushi, Road" -> "Dushi Road",   "Jalan, Sudirman 5" -> "Jalan Sudirman 5"
    Runs after the country and postcode have been taken out, so they can't be joined on.
    """
    def word(part):
        return part.lower().strip(".")

    merged = []
    for n, part in enumerate(parts):
        prev = merged[-1] if merged else None
        lone = prev is not None and (HOUSE_NUMBER_RE.match(prev) or word(prev) in NUMBER_LABELS
                                     or word(prev) in STREET_PREFIXES)
        if lone and not (n == len(parts) - 1 and len(merged) >= 2 and HOUSE_NUMBER_RE.match(prev)):
            merged[-1] = f"{prev} {part}"
        elif lone:
            # house number just before the town: it belongs with the piece before it
            merged.pop()
            merged[-1] = f"{merged[-1]} {prev}"
            merged.append(part)
        elif merged and word(part) in ROAD_SUFFIXES:
            merged[-1] = f"{merged[-1]} {part}"
        else:
            merged.append(part)
    return merged


def place_key(part):
    """Name used to spot repeats: lower case, without a trailing 'Province'/'City', e.g. 'Guangdong Province' -> 'guangdong'."""
    return re.sub(r"\s+(province|city|shi|sheng)$", "", part.lower())


def clean_postcode(postcode):
    """Uppercase and normalise spacing; UK postcodes get a single space before the inward code."""
    text = sql_safe(postcode).upper()
    if is_uk_postcode(text):
        compact = text.replace(" ", "")           # "CV48GL" / "CV4 8GL" / "CV4  8 GL"
        text = f"{compact[:-3]} {compact[-3:]}"   # -> "CV4 8GL"
    return text


# ===========================================================================
# STEP 2 - Work out the Local Country and the City/Town
# ===========================================================================

def country_name(part):
    """Lower-case country name if this piece is a recognised country, e.g. 'Vietnam (viet Nam)' -> 'vietnam'."""
    name = re.sub(r"\s*\(.*\)$", "", part).strip().lower()
    return name if name in WORLD_COUNTRIES or name in UK_COUNTRIES else ""


def split_country_postcode(part, country):
    """
    Return (rest_of_piece, postcode) if the piece is, starts with, or ends with
    a postcode in the country's own format (see COUNTRY_POSTCODES), else (part, '').
        Australia:   "Sydney NSW 2000"    -> ("Sydney NSW", "2000")
        Netherlands: "1015 CJ Amsterdam"  -> ("Amsterdam", "1015 CJ")
        Ireland:     "D02 X285"           -> ("", "D02 X285")
    The rest must look like a place name (no digits, not "Building ..." or "... Street"),
    so "Building 1234" or "Oxford Street 4000" are left alone.
    """
    name = country_name(country)
    pattern = COUNTRY_POSTCODES.get(COUNTRY_ALIASES.get(name, name))
    if not pattern:
        return part, ""
    code = rf"(?:{pattern})"

    def place_like(text, before_postcode):
        words = text.lower().split()
        if not words or any(ch.isdigit() for ch in text) or words[0].strip(".") in NUMBER_LABELS | UNIT_WORDS:
            return False
        last = words[-1].strip(".")
        if before_postcode:
            return last not in NUMBER_LABELS | UNIT_WORDS | ROAD_SUFFIXES   # not "Building 1234"
        return last not in ROAD_SUFFIXES                                   # "Roma RM" is fine

    if re.fullmatch(code, part, re.IGNORECASE):                       # the whole piece
        return "", part.upper()
    match = re.fullmatch(rf"(.*?)\s*[\s(-]\s*({code})\)?", part, re.IGNORECASE)   # place, then postcode
    if match and place_like(match.group(1), before_postcode=True):
        return match.group(1).strip(" -("), match.group(2).upper()
    match = re.fullmatch(rf"({code})\s+(.*)", part, re.IGNORECASE)    # postcode, then place
    if match and place_like(match.group(2), before_postcode=False):
        return match.group(2), match.group(1).upper()
    return part, ""


def take_country_postcode(parts, country):
    """
    Find and remove a postcode in the country's own format (see split_country_postcode),
    searching from the end. The first piece is never checked - it's usually a
    house or flat number. Returns the postcode, or '' if none found.
    """
    for idx in range(len(parts) - 1, 0, -1):
        rest, postcode = split_country_postcode(parts[idx], country)
        if postcode:
            if rest:
                parts[idx] = rest
            else:
                parts.pop(idx)
            return postcode
    return ""


def take_country(parts, country, want_town=True):
    """
    Return (Local Country, town): the Country column value if given (removing it
    from the address if it is repeated there), otherwise a recognised country
    found in the address, which is taken out of the address pieces.

    The country is usually the last piece, but is sometimes typed part-way
    through, e.g. "..., Wuhan, China 430050, Park Avenue, Central District".
    In that case the piece just before the country is usually the town, so it
    is returned as the town for Local Address 4 (e.g. "Wuhan") - unless
    want_town is False because the file has a City/Town column.
    """
    # The input file has a Country column with a value: use that
    if country:
        if parts and parts[-1].lower() == country.lower():
            parts.pop()
        return country, ""
    # Otherwise, look for a country in the address, starting from the end.
    # (The address must keep at least one other piece, so it's never reduced to nothing.)
    if len(parts) < 2:
        return "", ""
    for idx in range(len(parts) - 1, -1, -1):
        if country_name(parts[idx]):
            found = parts.pop(idx)
            town = ""
            if want_town and idx < len(parts) and idx >= 2 and looks_like_town(parts[idx - 1]):
                # country was part-way through: the piece before it is the town
                town = parts.pop(idx - 1)
            return found, town
    return "", ""


def looks_like_town(part):
    """
    A town name has no digits, fits in one field and isn't a street
    (so '7 Park Road' and 'Qiujiang Road' are not towns).
    """
    return (len(part) <= MAX_LEN and not any(ch.isdigit() for ch in part)
            and part.split()[-1].lower().strip(".") not in ROAD_SUFFIXES)


def move_rooms_to_front(parts):
    """
    A flat/room written at the end of the address ("..., Jingan District, Room 2002")
    is moved to the front so it ends up in Local Address 1, where it belongs.
    """
    # count the room pieces at the end; only move them if something else comes before them
    n = 0
    while n < len(parts) and is_room(parts[-1 - n]):
        n += 1
    if 0 < n < len(parts):
        parts[:] = parts[-n:] + parts[:-n]


def is_room(part):
    words = part.lower().split()
    return words[0].strip(".") in ROOM_WORDS and any(ch.isdigit() for ch in part)


def town_for_line_4(parts, overseas):
    """
    Take the City/Town for Local Address 4 off the end of the address (when the
    file has no City/Town column). For an overseas address the region is kept
    with the town when it fits, e.g. 'Nanjing Jiangsu'.
    """
    # Need at least one other piece left for Local Address 1, and the last piece
    # must look like a town
    if len(parts) < 2 or not looks_like_town(parts[-1]):
        return ""
    taken = [parts.pop()]
    if overseas:
        # Overseas addresses usually end "..., Town, Region": take the piece in
        # front too, while it fits in 30 characters and still leaves a line 1
        while (len(taken) < 2 and len(parts) >= 2 and looks_like_town(parts[-1])
               and len(" ".join([parts[-1]] + taken)) <= MAX_LEN):
            taken.insert(0, parts.pop())
    return " ".join(taken)


# ===========================================================================
# STEP 3 - Fit the address into Local Address 1-4
# ===========================================================================
# The address is handled as a list of "units". Each unit is (separator, text):
#   separator ", " -> this unit starts a new comma piece of the address
#   separator " "  -> this unit continues the previous piece (it was split between words)
# The separator is only used to know where pieces start; in the output, units
# are always joined with a plain space.

def piece_words(part):
    """The words of one piece; a single word longer than MAX_LEN is hard-cut into 30-character chunks."""
    words = []
    for word in part.split(" "):
        while len(word) > MAX_LEN:
            words.append(word[:MAX_LEN])
            word = word[MAX_LEN:]
        if word:
            words.append(word)
    return words


def wrap_piece(part):
    """
    Split a piece longer than MAX_LEN into chunks of up to 30 characters. Each
    chunk is filled with as many words as fit, but if that would separate words
    that belong together (see is_bad_break), words are moved down to the next
    chunk until it wouldn't:
        "Room 501 No 2 Lane 5058 Dushi Road"
        -> "Room 501 No 2 Lane 5058" + "Dushi Road"   (not "...Dushi" + "Road")
    """
    words = piece_words(part)
    chunks, i = [], 0
    while i < len(words):
        # fill the chunk with as many words as fit
        j = i + 1
        while j < len(words) and len(" ".join(words[i:j + 1])) <= MAX_LEN:
            j += 1
        # if the break is a bad one, move words down to the next chunk (but keep at least one)
        if j < len(words):
            k = j
            while k > i + 1 and is_bad_break(words[k - 1], words[k]):
                k -= 1
            if not is_bad_break(words[k - 1], words[k]):
                j = k
        chunks.append(" ".join(words[i:j]))
        i = j
    return chunks


def to_units(parts):
    """
    Turn comma pieces into (separator, text) units, one unit per piece. A piece
    longer than MAX_LEN is wrapped into 30-character chunks (see wrap_piece),
    each chunk after the first joined on with " " rather than ", ".
    """
    units = []
    for part in parts:
        chunks = wrap_piece(part)
        if chunks:
            units.append((", ", chunks[0]))              # start of a new piece
            units.extend((" ", c) for c in chunks[1:])   # rest of a long piece
    return units


def to_word_units(parts):
    """Like to_units, but every word is its own unit so lines can break between words."""
    units = []
    for part in parts:
        words = piece_words(part)
        if words:
            units.append((", ", words[0]))
            units.extend((" ", w) for w in words[1:])
    return units


def is_bad_break(word_before, word_after):
    """
    True if a line break between these two words splits words that belong together:
      * a street type would start the next line ("Dushi" | "Road")
      * a label would end the line ("Lane" | "5058", "Unit" | "2906")
      * a number would start the next line, cut off from its name ("Weiluyayuan" | "6")
    """
    before = word_before.lower().strip(".")
    after = word_after.lower().strip(".")
    street_after = after in ROAD_SUFFIXES and after not in STREET_PREFIXES   # "Avenue des..." may start a line
    return street_after or before in NUMBER_LABELS or before in STREET_PREFIXES or after.isdigit()


def bad_break(units, s):
    """True if starting a new line at unit s, part-way through a piece, is a bad break (see is_bad_break)."""
    if s == 0 or units[s][0] != " ":
        return False
    return is_bad_break(units[s - 1][1].split()[-1], units[s][1].split()[0])


def join_units(units):
    """Turn a list of units back into text, e.g. [(", ", "Hefei"), (", ", "Anhui")] -> "Hefei Anhui"."""
    return " ".join(piece for _, piece in units)


def balanced_split(units, n_lines):
    """
    Split units into at most n_lines contiguous groups. In order of priority:
      1. the fewest pieces are cut between words (lines break at commas where possible)
      2. the fewest "bad" cuts that separate words belonging together (see bad_break)
      3. the longest line is as short as possible (lines are evenly balanced)
    Returns the list of line strings, or None if no split keeps every line
    within MAX_LEN.

    This tries every possible place to break the lines (a technique called
    "dynamic programming") and keeps the best split it finds.
    """
    n = len(units)
    if n == 0:
        return []
    k = min(n, n_lines)
    # best[j][i] = ((pieces cut, bad cuts, max line length), lines) for the first i units in j lines
    best = [[None] * (n + 1) for _ in range(k + 1)]
    best[0][0] = ((0, 0, 0), [])
    for j in range(1, k + 1):              # j = number of lines used so far
        for i in range(j, n + 1):          # i = number of units placed so far
            for s in range(j - 1, i):      # s = where line j starts
                prev = best[j - 1][s]
                if prev is None:
                    continue
                line = join_units(units[s:i])
                if len(line) > MAX_LEN:
                    continue               # this line would be too long - not allowed
                cut = 1 if s > 0 and units[s][0] == " " else 0  # line starts mid-piece
                bad = 1 if bad_break(units, s) else 0
                score = (prev[0][0] + cut, prev[0][1] + bad, max(prev[0][2], len(line)))
                if best[j][i] is None or score < best[j][i][0]:
                    best[j][i] = (score, prev[1] + [line])
    result = best[k][n]
    return result[1] if result else None


def greedy_fill(units, n_lines):
    """Fill the first n_lines lines up to MAX_LEN; return (lines, leftover_units)."""
    lines, i = [], 0
    while len(lines) < n_lines and i < len(units):
        line_units = [units[i]]
        i += 1
        # keep adding units to this line while it stays within 30 characters
        while i < len(units) and len(join_units(line_units + [units[i]])) <= MAX_LEN:
            line_units.append(units[i])
            i += 1
        lines.append(join_units(line_units))
    return lines, units[i:]


def pad_lines(lines, city=""):
    """Make sure there are exactly 4 lines (blank where unused), with the City/Town in line 4 if given."""
    lines = (lines + [""] * 4)[:4]
    if city:
        lines[3] = city
    return lines


def pack_lines(parts, city):
    """
    Fit the address pieces (and City/Town) into Local Address 1-4, trying each
    step in turn and stopping at the first that fits. The order of the address
    is always kept and the City/Town always ends Local Address 4.

      1. Pieces spread evenly over Local Address 1-3, City/Town alone in 4.
      2. Pieces spread evenly over Local Address 1-4, ending with the City/Town
         (Local Address 4 becomes e.g. "Central District Nanjing").
      3. As step 2, but pieces may also be split between words to fill the lines.
      4. Every field filled to 30 characters; anything left goes to Overspill,
         taken from the middle of the address so the City/Town stays in 4.

    Returns (lines, overspill_text, note).
    """
    piece_units = to_units(parts)
    city_units = to_units([city]) if city else []

    # 1. City/Town on its own in Local Address 4 (or all four lines when there's none)
    lines = balanced_split(piece_units, 3 if city else 4)
    if lines is not None:
        return pad_lines(lines, city), "", ""

    # 2. and 3. City/Town shares Local Address 4 with the end of the address
    word_units = to_word_units(parts)
    for units, note in [(piece_units, "Address spread across all four fields"),
                        (word_units, "Address split between words to fit all four fields")]:
        lines = balanced_split(units + city_units, 4)
        if lines is not None:
            return pad_lines(lines), "", note

    # 4. Doesn't fit: fill every field, overspill the middle of the address
    lines, leftover = greedy_fill(word_units, 3)
    if city:
        # Local Address 4 = as many of the next words as fit, then the City/Town
        last_units = []
        while leftover and len(join_units(last_units + [leftover[0]] + city_units)) <= MAX_LEN:
            last_units.append(leftover.pop(0))
        lines = pad_lines(lines)[:3] + [join_units(last_units + city_units)]
    else:
        last, leftover = greedy_fill(leftover, 1)
        lines = pad_lines(lines + last)
    return lines, join_units(leftover), ""


# ===========================================================================
# STEP 4 - Put it all together for one row
# ===========================================================================

def split_address(address, city, postcode=None, strip_postcodes=True, country=None):
    """
    Return (lines [Local Address 1..4], local_postcode, local_country,
            overspill, notes, review) for one row.
    review is True when the row should be flagged amber. (Overspill alone doesn't
    flag the row - only the Overspill cell is highlighted.)
    """
    # Step 1: clean the address into pieces and find any postcode in it
    city = sql_safe(city)
    parts, found_postcode = clean_parts(address, city, strip_postcodes)

    # Step 2: take the country off the end, then the City/Town (if no City/Town column)
    local_country, town_before_country = take_country(parts, sql_safe(country), want_town=not city)
    overseas = country_name(local_country) not in UK_COUNTRIES if local_country else False

    # Now the country is known, look for a postcode in that country's format
    # (e.g. Australia "2000", Canada "K1N 5W8") if one wasn't found already
    if not found_postcode and strip_postcodes:
        found_postcode = take_country_postcode(parts, local_country)

    # Join fragments like a house number on its own to their neighbour
    parts[:] = join_fragments(parts)
    if not city:
        city = town_before_country or town_for_line_4(parts, overseas)
    # A flat/room written at the end of the address goes to the front instead
    move_rooms_to_front(parts)
    notes, overspill, review = [], [], False

    # A City/Town too long for one field is treated as part of the address instead
    if len(city) > MAX_LEN:
        notes.append(f"City/Town is {len(city)} characters (over {MAX_LEN}) - treated as part of the address")
        review = True
        parts.append(city)
        city = ""

    # Step 3: fit everything into Local Address 1-4
    lines, leftover, note = pack_lines(parts, city)
    if note:
        notes.append(note)
    if leftover:
        # (the Overspill cell is highlighted amber - see main)
        overspill.append(leftover)
        notes.insert(0, "Address does not fit in Local Address fields even with every field filled")

    # Local Postcode: the Postcode column if it has a value, otherwise one found in the address
    local_postcode = clean_postcode(postcode) or clean_postcode(found_postcode)
    if len(local_postcode) > POSTCODE_MAX_LEN:
        notes.append(f"Postcode is {len(local_postcode)} characters (over {POSTCODE_MAX_LEN})")
        overspill.append(local_postcode)
        local_postcode = ""

    return lines, local_postcode, local_country, "; ".join(overspill), "; ".join(notes), review


# ===========================================================================
# STEP 4b - For files with separate address columns (Address1, Address2, ...)
# ===========================================================================
# The address has already been split into lines by whoever entered it. If those
# lines already follow the rules they are kept as entered (only commas and
# apostrophes are removed, and the town is moved to line 4). If not, the lines
# are joined back together and reworked with the same rules as a single
# address string (STEP 1-4 above).

def has_postcode(line, country=""):
    """
    True if the line is, or ends with, a postcode: UK, a 5-6 digit overseas one,
    or one in the country's own format if the country is known.
    """
    return bool(is_uk_postcode(line) or trailing_uk_postcode(line) or split_numeric_postcode(line)[1]
                or (country and split_country_postcode(line, country)[1]))


def has_country(line):
    """True if the line is a country, or ends with one (e.g. 'Wuhan China')."""
    low = re.sub(r"\s*\(.*\)$", "", line).lower()
    for name in WORLD_COUNTRIES | UK_COUNTRIES:
        if low == name or (len(name) > 3 and low.endswith(" " + name)):
            return True
    return False


def split_across_lines(line_before, line_after):
    """
    True if something that belongs together has been split across two lines:
      * a house number on its own line ("12" | "High Street")
      * a label at the end of a line ("Flat" | "3 Station Road")
      * a street type at the start of a line ("Dushi" | "Road"). "St" and "Lane"
        are allowed to start a line, as they often begin a name ("St Marys Court",
        "Lane 1488").
    """
    last = line_before.split()[-1].lower().strip(".")
    words = line_after.lower().split()
    first = words[0].strip(".")
    street_cut = (first in ROAD_SUFFIXES and first not in STREET_PREFIXES
                  and (len(words) == 1 or first not in {"st", "lane"}))
    return (bool(HOUSE_NUMBER_RE.match(line_before)) or last in NUMBER_LABELS
            or last in STREET_PREFIXES or street_cut)


def line_problem(lines, city, country=""):
    """
    Check address lines that were already split. Returns '' if they can be kept
    as entered, otherwise the reason they need reworking.
    lines: the non-blank lines, already cleaned with sql_safe.
    """
    if len(lines) > 4 or (city and len(lines) > 3):
        return "too many address lines"
    for n, line in enumerate(lines, 1):
        if len(line) > MAX_LEN:
            return f"line {n} is over {MAX_LEN} characters"
        if has_postcode(line, country):
            return f"postcode in line {n}"
        if has_country(line):
            return f"country in line {n}"
    keys = [place_key(line) for line in lines]
    if len(set(keys)) < len(keys):
        return "repeated line"
    for n in range(1, len(lines)):
        if split_across_lines(lines[n - 1], lines[n]):
            return f"name or number split across lines {n} and {n + 1}"
    return ""


def split_address_lines(values, city, postcode=None, strip_postcodes=True, country=None):
    """
    Like split_address, but for an address already split across several columns.
    values: the address column values in order (Address1, Address2, ...).
    Returns the same (lines, postcode, country, overspill, notes, review).
    """
    raw = [str(v) for v in values if v is not None and str(v).strip()]
    city = sql_safe(city)
    lines = [sql_safe(v) for v in raw]
    lines = [line for line in lines if line and not (city and line.lower() == city.lower())]

    problem = line_problem(lines, city, sql_safe(country))
    if problem:
        # Rework: join the lines back into one address and apply the normal rules
        result = list(split_address(", ".join(raw), city, postcode, strip_postcodes, country))
        result[4] = "; ".join(filter(None, [f"Reworked: {problem}", result[4]]))
        return tuple(result)

    # Keep as entered. The town goes in line 4: the City/Town column if given,
    # otherwise the last line when it looks like a town.
    if city:
        lines = lines + [""] * (3 - len(lines)) + [city]
    elif 2 <= len(lines) < 4 and looks_like_town(lines[-1]):
        lines = lines[:-1] + [""] * (4 - len(lines)) + [lines[-1]]
    lines = pad_lines(lines)

    notes, overspill, review = ["Kept as entered"], [], False
    local_postcode = clean_postcode(postcode)
    if len(local_postcode) > POSTCODE_MAX_LEN:
        notes.append(f"Postcode is {len(local_postcode)} characters (over {POSTCODE_MAX_LEN})")
        overspill.append(local_postcode)
        local_postcode = ""
    return lines, local_postcode, sql_safe(country), "; ".join(overspill), "; ".join(notes), review


# ===========================================================================
# STEP 5 - Build the SQL UPDATE statement for one row
# ===========================================================================

def sql_value(value):
    """
    Turn a value into SQL: blank -> NULL, anything else -> 'text' in single quotes.
    Whole numbers stored by Excel as e.g. 8000002.0 lose the '.0'. Any single
    quote in the value is doubled ('') so it can't break the statement.
    """
    if value is None:
        return "NULL"
    if isinstance(value, float) and value.is_integer():
        value = int(value)
    text = " ".join(str(value).split())
    return "'" + text.replace("'", "''") + "'" if text else "NULL"


def regnumber_text(value):
    """
    The Regnumber exactly as it is in the file - NOT trimmed, so trailing spaces
    are kept (and counted by the 7-character check). Whole numbers stored by
    Excel as e.g. 8000002.0 lose the '.0'.
    """
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        value = int(value)
    return str(value)


def build_update(fields, values, regnumber, acronym=None, match_acronym=False):
    """
    e.g. update core c set c.localaddress1 = 'Flat 3', ..., c.localcountry = NULL
         where c.regnumber = '8000002' and c.acronym = 'EXTMOR';
    fields: the database field names (from TARGETS, e.g. localaddress1 ... localcountry).
    values: [address lines 1-4, postcode, country] in the same order.
    Returns '' when there is no Regnumber to match on.
    """
    reg = regnumber_text(regnumber)
    if not reg.strip():
        return ""
    assignments = ", ".join(f"c.{field} = {sql_value(v or None)}" for field, v in zip(fields, values))
    # the Regnumber goes into the SQL exactly as in the file - no trimming
    where = "c.regnumber = '" + reg.replace("'", "''") + "'"
    if match_acronym:
        where += f" and c.acronym = {sql_value(acronym)}"
    return f"update core c set {assignments} where {where};"


# ===========================================================================
# Reading and writing the files
# ===========================================================================

def find_column_loose(headers, name):
    """Position of the column called `name`, ignoring case, spaces and underscores, or None."""
    want = re.sub(r"[\s_]", "", name).lower()
    for idx, header in enumerate(headers):
        if re.sub(r"[\s_]", "", str(header or "")).lower() == want:
            return idx
    return None


def contact_value(value):
    """Email/mobile value as text (None if blank). A number stored by Excel as 7700900123.0 loses the '.0'."""
    if value is None:
        return None
    if isinstance(value, float) and value.is_integer():
        value = int(value)
    return str(value).strip() or None


def clean_mobile(value):
    """
    Return (mobile, note). A mobile number may only contain digits, with an
    optional + at the start; spaces, dashes, dots and brackets are removed:
        "07700 900123" -> "07700900123"     "+44 (0)7700-900123" -> "+4407700900123"
    Anything else (e.g. "Not yet", "N/A") isn't a phone number, so it becomes
    None (NULL in the SQL) and a note says so.
    """
    text = contact_value(value)
    if text is None:
        return None, ""
    compact = re.sub(r"[\s\-.()]", "", text)
    if re.fullmatch(r"\+?\d+", compact):
        return compact, ""
    return None, f"Mobile phone '{text}' is not a number, so set to NULL"


def heading_key(text):
    """A heading reduced to lower-case letters only, e.g. 'Post Code:' -> 'postcode'."""
    return re.sub(r"[^a-z]", "", str(text or "").lower())


def find_heading(headers, names, local_first):
    """
    Position of the first column whose heading starts with one of `names`,
    ignoring case, spaces and punctuation ("Post Code:" matches "Postcode").
    A "Local ..." version of each name is also accepted ("Local Postcode"):
    tried first when local_first is True (updating localaddress fields), last otherwise.
    Returns None if there is no such column.
    """
    keys = [heading_key(n) for n in names]
    local = ["local" + k for k in keys]
    for key in (local + keys if local_first else keys + local):
        for idx, header in enumerate(headers):
            if heading_key(header).startswith(key):
                return idx
    return None


def find_column(headers, name, prefix=False, required=True):
    """Return the position of the column called `name` (or starting with it if prefix=True)."""
    for idx, header in enumerate(headers):
        h = str(header or "").strip().lower()
        if (h.startswith(name.lower()) if prefix else h == name.lower()):
            return idx
    if required:
        raise SystemExit(f"Could not find a '{name}' column in the header row: {headers}")
    return None


def main():
    # --- Read the file names typed after the script name -------------------
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("input", nargs="?", default=DEFAULT_INPUT)
    parser.add_argument("output", nargs="?", default=DEFAULT_OUTPUT)
    parser.add_argument("--keep-postcodes", action="store_true",
                        help="Leave postcodes found inside the Address text in the address lines")
    parser.add_argument("--target", choices=list(TARGETS), default=DEFAULT_TARGET,
                        help="Which database fields to update: localaddress (localaddress1-4, "
                             "localpostcode, localcountry - the default) or address (address1-4, "
                             "postcode, country)")
    args = parser.parse_args()
    target = TARGETS[args.target]
    print(f"Updating the {args.target} fields: {', '.join(target['sql'])}")

    # --- Open the input file (first sheet) and read the header row ---------
    # data_only: copy formula columns (e.g. CONCAT) as their values, not the formula text
    src = openpyxl.load_workbook(args.input, read_only=True, data_only=True).worksheets[0]
    rows = src.iter_rows(values_only=True)
    headers = list(next(rows))

    # --- Find the address column(s) -----------------------------------------
    # Numbered address columns (Address1, Address2, ...) if the file has them,
    # sorted by their number; otherwise the single Address column
    numbered = sorted((int(m.group(3)), idx) for idx, h in enumerate(headers)
                      if (m := ADDRESS_LINE_HEADER_RE.match(str(h or "").strip())))
    line_idxs = [idx for _, idx in numbered]
    if line_idxs:
        addr_idx = None
        print(f"Address columns: {', '.join(str(headers[i]) for i in line_idxs)} "
              "(rows already within the rules are kept as entered)")
    else:
        addr_idx = find_column(headers, ADDRESS_HEADER)
        print(f"Address column: {headers[addr_idx]} (one address string per row)")

    # --- Find the other columns (all optional) ------------------------------
    local_first = args.target == "localaddress"
    city_idx = find_heading(headers, CITY_HEADINGS, local_first)
    postcode_idx = find_heading(headers, POSTCODE_HEADINGS, local_first)
    country_idx = find_heading(headers, COUNTRY_HEADINGS, local_first)
    for name, idx in [("City/Town", city_idx), ("Postcode", postcode_idx), ("Country", country_idx)]:
        if idx is not None:
            print(f"{name} column: {headers[idx]}")
    for name, idx in [("City/Town", city_idx), ("Postcode", postcode_idx), ("Country", country_idx)]:
        if idx is None:
            print(f"No {name} column found - taking it from the Address where possible")

    regnumber_idx = next((idx for idx in (find_column(headers, p, prefix=True, required=False)
                                          for p in REGNUMBER_HEADER_PREFIXES) if idx is not None), None)
    acronym_idx = find_column(headers, ACRONYM_HEADER, required=False)
    if regnumber_idx is None:
        print("No Regnumber column found - SQL UPDATE statements will not be created")
    elif acronym_idx is None:
        print("No Acronym column found - SQL UPDATE statements will match on regnumber only")

    # Optional Email / MobilePhone / Usedname columns: [(column position, database field), ...]
    # A field is only updated if its column is in the input file.
    contacts = [(idx, field) for heading, field in CONTACT_COLUMNS
                if (idx := find_column_loose(headers, heading)) is not None]
    mobile_idx = find_column_loose(headers, MOBILE_HEADING)
    if contacts:
        print(f"Also updating: {', '.join(field for _, field in contacts)}")

    # --- Set up the output file --------------------------------------------
    # Keep every original column except the address/city/postcode/country
    # columns, which are replaced by the new ones
    # (with numbered columns, a plain 'Address' column is dropped too if there is one)
    plain_address_idx = find_column(headers, ADDRESS_HEADER, required=False)
    replaced = {addr_idx, plain_address_idx, city_idx, postcode_idx, country_idx, *line_idxs}
    keep_idx = [i for i in range(len(headers)) if i not in replaced]

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Split Addresses"
    ws.append([headers[i] for i in keep_idx] + target["columns"] + EXTRA_COLUMNS)
    for cell in ws[1]:
        cell.font = Font(bold=True)

    # Output column numbers (1 = column A) of the cells that can be highlighted on their own
    regnumber_col = keep_idx.index(regnumber_idx) + 1 if regnumber_idx in keep_idx else None
    overspill_col = len(keep_idx) + len(target["columns"]) + 1

    # --- Process each row ---------------------------------------------------
    processed = flagged = statements = overspill_count = long_regnumbers = 0
    for row in rows:
        if all(v is None for v in row):
            continue  # skip completely empty rows

        # Read this row's values (None when the file has no such column)
        city = row[city_idx] if city_idx is not None else None
        postcode = row[postcode_idx] if postcode_idx is not None else None
        country = row[country_idx] if country_idx is not None else None

        # Split the address (steps 1-4 above), or check/rework the existing lines (step 4b)
        if line_idxs:
            lines, local_postcode, local_country, overspill, note, review = split_address_lines(
                [row[i] for i in line_idxs], city, postcode,
                strip_postcodes=not args.keep_postcodes, country=country)
        else:
            lines, local_postcode, local_country, overspill, note, review = split_address(
                row[addr_idx], city, postcode, strip_postcodes=not args.keep_postcodes, country=country)

        # Safety net: make certain no comma or apostrophe of any kind reaches the
        # output or the SQL, whichever route the row took
        # ...and give every word of each address line / town a capital first letter
        lines = [capitalise_words(sql_safe(line)) for line in lines]
        local_postcode, local_country, overspill = sql_safe(local_postcode), sql_safe(local_country), sql_safe(overspill)

        # Email / mobile values for this row (only if those columns exist)
        contact_fields = [field for _, field in contacts]
        contact_values = []
        for idx, field in contacts:
            if idx == mobile_idx:
                mobile, mobile_note = clean_mobile(row[idx])   # non-numbers like "Not yet" -> NULL
                contact_values.append(mobile)
                note = "; ".join(filter(None, [note, mobile_note]))
                if isinstance(row[idx], (int, float)):
                    # Excel stores numbers without a leading 0, e.g. 07700900123 -> 7700900123
                    note = "; ".join(filter(None, [note, "Mobile phone is stored as a number in Excel - check the leading 0"]))
                    review = True
            else:
                contact_values.append(contact_value(row[idx]))

        # A blank Usedname leaves c.usedname unchanged: it's left out of this row's statement
        keep = [not (field in SKIP_IF_BLANK and value is None) for field, value in zip(contact_fields, contact_values)]
        contact_fields = [f for f, k in zip(contact_fields, keep) if k]
        contact_values = [v for v, k in zip(contact_values, keep) if k]

        # Build the SQL UPDATE (step 5). A row with no address doesn't update the
        # address fields - otherwise it would set them to NULL in the database.
        fields = target["sql"] + contact_fields
        values = lines + [local_postcode, local_country] + contact_values
        if not any(lines + [local_postcode, local_country]):
            review = True
            if any(contact_values):
                fields, values = contact_fields, contact_values
                no_address = "No address in this row, so only " + "/".join(contact_fields) + " updated"
            else:
                fields = []
                no_address = "No address in this row, so no SQL update"
            note = "; ".join(filter(None, [no_address, note]))

        update = ""
        if fields and regnumber_idx is not None:
            update = build_update(fields, values, row[regnumber_idx],
                                  row[acronym_idx] if acronym_idx is not None else None,
                                  match_acronym=acronym_idx is not None)
            if not update:
                note = "; ".join(filter(None, [note, "No Regnumber, so no SQL update"]))
                review = True

        # Regnumber longer than 7 characters, counting any spaces (it is not trimmed)
        long_regnumber = False
        if regnumber_idx is not None:
            reg = regnumber_text(row[regnumber_idx])
            if len(reg) > REGNUMBER_MAX_LEN:
                long_regnumber = True
                spaces = " (includes spaces at the start or end)" if reg != reg.strip() else ""
                note = "; ".join(filter(None, [note, f"Regnumber is {len(reg)} characters, more than "
                                                      f"{REGNUMBER_MAX_LEN}{spaces}"]))

        # Write the row to the output file
        ws.append([row[i] for i in keep_idx]
                  + [v or None for v in lines + [local_postcode, local_country, overspill, note, update]])
        out_row = ws.max_row

        # Highlighting, in this order so a later colour wins on the same cell:
        #   amber row        - the row needs review (see Split Notes)
        #   amber Overspill  - text that wouldn't fit is in the Overspill cell (only that cell)
        #   red Regnumber    - the Regnumber is more than 7 characters (only that cell)
        if review:
            for cell in ws[out_row]:
                cell.fill = AMBER_FILL
            flagged += 1
        if overspill:
            ws.cell(row=out_row, column=overspill_col).fill = AMBER_FILL
            overspill_count += 1
        if long_regnumber and regnumber_col:
            ws.cell(row=out_row, column=regnumber_col).fill = RED_FILL
            long_regnumbers += 1

        if update:
            statements += 1
        processed += 1

    # --- Save the output file -----------------------------------------------
    # The SQL UPDATE statements are only in the 'SQL Update' column - no .sql file is written
    for idx in range(1, ws.max_column + 1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(idx)].width = 32

    output = Path(args.output)
    if output.suffix.lower() != ".xlsx":
        output = output.with_name(output.name + ".xlsx")   # add .xlsx if it was left off
    wb.save(output)
    print(f"Processed {processed} rows -> {output.resolve()}")
    print(f"SQL update statements (in the 'SQL Update' column): {statements}")
    print(f"Rows flagged amber for review: {flagged}")
    print(f"Overspill cells highlighted amber: {overspill_count}")
    print(f"Regnumbers over {REGNUMBER_MAX_LEN} characters highlighted red: {long_regnumbers}")


if __name__ == "__main__":
    main()
