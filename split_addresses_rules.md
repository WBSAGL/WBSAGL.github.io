# split_addresses.py – Rules

Splits student addresses into four 30‑character lines, a postcode and a country, cleans contact details, and writes a new Excel file with one SQL `UPDATE` per row.

## Preparing the input file

Save the student data as an Excel file (`.xlsx`) with **these columns**, headings in row 1 of the first sheet:

| Regnumber | Acronym | Firstname | Usedname | Lastname | Email | mobile_phone | Address1 | Address2 | Address3 | Address4 | City/Town | Postcode |
|---|---|---|---|---|---|---|---|---|---|---|---|---|

- Paste the data in **as received** – the script does the cleaning. Leave a cell blank if there's no value.
- **Address in one long string?** Put the whole thing in **Address1** and leave Address2–4 blank.
- **Address over several fields?** Put them in Address1, Address2, … in order.
- **City/Town filled?** Leave **Address4** blank – the town goes in line 4.
- Firstname and Lastname are for reference only – the script doesn't update them.
- A **Country** column can be added if the data has one (otherwise the country is taken from the address).

## How to run it (step by step)

1. **Close the input file in Excel** – the script can't read a file that's open.
2. **Open a terminal.** In VS Code: *Terminal → New Terminal*. Check the line starts with `PS` (not `>>>` – if so, type `exit()` and press Enter).
3. **Go to the folder with your Excel file** (put the path in quotes):
   ```
   cd "C:\path\to\folder with the Excel file"
   ```
4. **Type the run command** – `python`, then the script, then the input file, then a name for the new output file, each in quotes:
   ```
   python "C:\path\to\split_addresses.py" "My input.xlsx" "My output.xlsx"
   ```
   Add `--target address` at the end to update the **address** fields instead of the **localaddress** fields.
5. **Press Enter.** The script prints which columns it found and a summary (rows processed, SQL statements, rows highlighted).
6. **Open the output file** and check the highlighted rows/cells and the **Split Notes** column, then copy the statements from the **SQL Update** column.

**If something goes wrong**

| Message | What it means |
|---|---|
| `Permission denied` | The Excel file is still open – close it and try again |
| `can't find '__main__' module` | The script's path was left out of the command |
| `Could not find a 'Address' column` | Check the column headings in the input file |
| `SyntaxError … unicodeescape` / `>>>` on screen | The command was typed into Python, not the terminal – type `exit()` first |

## Options

```
python split_addresses.py "input.xlsx" "output.xlsx" [--target localaddress | address] [--keep-postcodes]
```

| Option | Effect |
|---|---|
| `--target localaddress` (default) | Updates `localaddress1–4`, `localpostcode`, `localcountry` |
| `--target address` | Updates `address1–4`, `postcode`, `country` |
| `--keep-postcodes` | Leaves postcodes found in the address text in the address lines |

- The input file is never changed. Output is **one Excel file only** (no `.sql` file).
- Field limits: address lines **30**, postcode **10**, Regnumber **7** characters.

## Columns the script recognises

Headings ignore case, spaces and punctuation, so close variations also work.

| Column | Required? | Headings recognised |
|---|---|---|
| Address | **Yes** | `Address1`, `Address 2`, `Address Line 3`… (used in order) **or** one `Address` column |
| City/Town | No | City/Town, Town/City, City, Town |
| Postcode | No | Postcode, Post Code, Postal Code, Zip |
| Country | No | Country… |
| Regnumber | For SQL | Regnumber…, Student Number… |
| Acronym | No | Acronym |
| Email / MobilePhone / Usedname | No | Email, MobilePhone, Usedname |

- "Local …" versions (e.g. Local Postcode) are preferred with `--target localaddress`.
- **When City/Town is filled, leave Address4 blank** – the town goes in line 4.
- Other columns (e.g. Firstname, Lastname) are copied to the output unchanged and never used.

## Address rules

**Separate address columns:** a row that already follows the rules is **kept as entered** (cleaned, capitalised, town moved to line 4). It is **reworked** if: more than 4 lines (3 with City/Town), a line over 30 characters, a postcode or country in a line, a repeated line, or something split across lines (house number alone, "Flat" | "3", "Calle" | "Mayor", "Dushi" | "Road").

**Cleaning**
1. The address is split at **every kind of comma** (`, ， 、 ﹐ ﹑ ، ‚`) and line break.
2. Commas become spaces; **all apostrophes/quotes are removed**. No commas or apostrophes reach any address, postcode or country output.
3. Empty and repeated pieces are dropped ("Guangdong Province" = "Guangdong").
4. Loose fragments are joined: a house number or label joins the next piece; a street type joins the previous one.
5. **Every word is capitalised** (first letter only; "12 high street" → "12 High Street").

**Postcode**
1. The Postcode column is used if filled, otherwise one found in the address. It goes into the target postcode field and is not output separately.
2. Recognised in the address: real **UK postcodes** (valid area and letters – "D306 IQ" is not one), **5–6 digit** codes, and the **country's own format** once the country is known (4 digits for Australia, Switzerland etc.; Canada `K1N 5W8`; Netherlands `1015 CJ`; Japan `131-0045`; Ireland `D02 X285`; and others).
3. Upper‑cased. Over 10 characters → moved to Overspill.

**Country**
1. The Country column if filled, otherwise a recognised country name (~200, plus USA, PR China etc.) found anywhere in the address, which is removed from the lines.
2. Kept as typed (bracketed extras dropped: "Vietnam (viet Nam)" → "Vietnam").

**Town (line 4)**
1. The City/Town column if filled, otherwise the last piece if it has no digits, isn't a street and fits. Overseas addresses add the region if it fits ("Nanjing Jiangsu").
2. A room/flat written at the end of an address is moved to line 1.

**Fitting into four lines** (address order always kept, town always ends line 4) – first step that fits wins:
1. Spread evenly over lines 1–3, town alone in line 4.
2. Spread over all four lines, town at the end of line 4.
3. As 2, splitting pieces between words – never "Dushi | Road", "Lane | 5058", "Calle | Mayor".
4. Fill every line to 30 characters; the rest goes to **Overspill**.

## Contact details

| Field | Rule | Blank cell |
|---|---|---|
| Email | Used as entered (trimmed) | **Not updated** |
| Mobile | Digits and optional leading `+` only (spaces, dashes, brackets removed). Anything else ("Not yet") → NULL | **Not updated** |
| Usedname | Used as entered (trimmed). **Only from the Usedname column – never from Firstname** | **Not updated** |

- Each field is only updated if its column exists in the input.
- **Not updated** = left out of that row's statement, so the database value is unchanged.

| In the input | In the SQL |
|---|---|
| Email blank | no `c.email` |
| Mobile blank (or only spaces) | no `c.mobile_phone` |
| Mobile "Not yet" / "N/A" | `c.mobile_phone = NULL` (not blank – just not a number) |
| Usedname blank (even if Firstname is filled) | no `c.usedname` |
| No address, email only | `update core c set c.email = '…' where …` |
| No address and no contact details | no statement |

## Regnumber

- Used **exactly as in the file** – never trimmed.
- More than 7 characters (spaces included) → Regnumber cell **red**.

## SQL

```sql
update core c set c.localaddress1 = '…', …, c.localcountry = '…'[, c.email = '…'][, c.mobile_phone = '…'][, c.usedname = '…']
where c.regnumber = '…'[ and c.acronym = '…'];
```

- In the **SQL Update** column only. Values in single quotes, `'` doubled.
- Blank address, postcode or country → `NULL`. Blank email, mobile or usedname → left out of the statement.
- No address in the row → only email/mobile/usedname updated (or no statement).
- No Regnumber → no statement.

## Output

- All input columns kept **except** the address, City/Town, Postcode and Country columns, then: Address 1–4, Postcode, Country, **Overspill**, **Split Notes** (what happened to the row), **SQL Update**.

| Highlight | When |
|---|---|
| Amber – whole row | No address, no Regnumber, mobile stored as a number (leading 0 lost), City/Town over 30 |
| Amber – Overspill cell | Text or postcode didn't fit |
| Red – Regnumber cell | Regnumber over 7 characters |

## Known limitations

- Postcodes in unlisted formats stay in an address line (nothing is lost).
- Town detection is pattern‑based; an area name can occasionally land in line 4.
- Country names aren't standardised; "…, Georgia" without "USA" is read as the country Georgia.
- Removing apostrophes changes names in address lines ("St Mary's" → "St Marys").
- Blank address, postcode or country values clear the existing database value.
