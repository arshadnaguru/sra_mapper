import pandas as pd
import os
import sys
from glob import glob
from datetime import datetime

# ── CONFIG ────────────────────────────────────────────────────────────────────
SRA_FOLDER  = "sra_files"          # folder with all SRA Excel files
RAPID_FILE  = "rapid_export.xlsx"  # full RAPID export
OUTPUT_FILE = f"rapid_with_sra_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
SRS_COLUMN  = "SRS #"              # common SRS column name in both files

# ── HELPERS ───────────────────────────────────────────────────────────────────
def clear(): os.system('cls' if os.name == 'nt' else 'clear')

def print_header():
    print("=" * 65)
    print("   SRA MAPPER — RIT Office of the VP for Research")
    print("=" * 65)
    print()

def smart_read(path):
    """Read correct sheet from RAPID-style Excel — skips empty sheets."""
    xl = pd.ExcelFile(path)
    if 'Worksheet' in xl.sheet_names:
        return pd.read_excel(path, sheet_name='Worksheet')
    # Pick sheet with most rows
    best, best_rows = xl.sheet_names[0], 0
    for sheet in xl.sheet_names:
        df = pd.read_excel(path, sheet_name=sheet)
        if len(df) > best_rows:
            best, best_rows = sheet, len(df)
    return pd.read_excel(path, sheet_name=best)

def find_srs_column(df, filename=""):
    """Find SRS column regardless of exact name."""
    for col in df.columns:
        if 'srs' in col.lower():
            return col
    print(f"\n  [!] Could not find SRS column in {filename}")
    print(f"      Available columns: {df.columns.tolist()}")
    col = input("      Enter the exact SRS column name: ").strip()
    return col if col in df.columns else None

def extract_sra_name(filename):
    """Extract clean SRA name from filename."""
    name = os.path.basename(filename)
    name = name.replace('.xlsx', '').replace('.xls', '')
    name = name.replace('_', ' ').replace('-', ' ')
    return name.strip()

# ── MAIN ──────────────────────────────────────────────────────────────────────
clear()
print_header()

# ── STEP 1: Read RAPID export ─────────────────────────────────────────────────
print("[1] Reading RAPID full export...")

if not os.path.exists(RAPID_FILE):
    print(f"\n  [ERROR] '{RAPID_FILE}' not found in this folder.")
    print("  Make sure rapid_export.xlsx is in the same folder as this script.")
    input("\n  Press Enter to exit.")
    sys.exit()

rapid = smart_read(RAPID_FILE)
rapid_srs_col = find_srs_column(rapid, RAPID_FILE)

if not rapid_srs_col:
    print("  [ERROR] Cannot find SRS column in RAPID file. Exiting.")
    input("\n  Press Enter to exit.")
    sys.exit()

rapid[rapid_srs_col] = rapid[rapid_srs_col].astype(str).str.strip()
rapid_srs_all = set(rapid[rapid_srs_col].dropna())

print(f"  OK — {len(rapid)} rows, {rapid[rapid_srs_col].nunique()} unique SRS numbers")
print(f"  SRS column: '{rapid_srs_col}'")

# ── STEP 2: Read all SRA files ────────────────────────────────────────────────
print(f"\n[2] Reading SRA files from '{SRA_FOLDER}/' folder...")

sra_files = glob(os.path.join(SRA_FOLDER, "*.xlsx")) + \
            glob(os.path.join(SRA_FOLDER, "*.xls"))

if not sra_files:
    print(f"\n  [ERROR] No Excel files found in '{SRA_FOLDER}/' folder.")
    print("  Create the folder and drop your SRA files in it.")
    input("\n  Press Enter to exit.")
    sys.exit()

# Build SRS → SRA name mapping
srs_to_sra = {}   # SRS number → SRA name
sra_srs_map = {}  # SRA name → set of SRS numbers

print()
for path in sorted(sra_files):
    sra_name = extract_sra_name(path)
    try:
        df = smart_read(path)
        srs_col = find_srs_column(df, path)
        if not srs_col:
            print(f"  [SKIP] {os.path.basename(path)} — no SRS column found")
            continue

        df[srs_col] = df[srs_col].astype(str).str.strip()
        srs_set = set(df[srs_col].dropna())
        sra_srs_map[sra_name] = srs_set

        for srs in srs_set:
            if srs in srs_to_sra:
                # Already assigned to another SRA — handle overlap
                existing = srs_to_sra[srs]
                if isinstance(existing, list):
                    existing.append(sra_name)
                else:
                    srs_to_sra[srs] = [existing, sra_name]
            else:
                srs_to_sra[srs] = sra_name

        print(f"  OK  {sra_name:<30} — {len(srs_set):>3} SRS numbers")

    except Exception as e:
        print(f"  [ERROR] {os.path.basename(path)} — {e}")

# ── STEP 3: Match and create SRA column ──────────────────────────────────────
print(f"\n[3] Matching SRS numbers...")

assigned   = 0
unassigned = []
overlapping = []

sra_labels = []

for _, row in rapid.iterrows():
    srs = str(row[rapid_srs_col]).strip()
    match = srs_to_sra.get(srs)

    if match is None:
        sra_labels.append("Unassigned")
        unassigned.append(srs)
    elif isinstance(match, list):
        # In multiple SRA files
        sra_labels.append(" | ".join(match))
        overlapping.append((srs, match))
        assigned += 1
    else:
        sra_labels.append(match)
        assigned += 1

rapid['SRA Name'] = sra_labels

print(f"  Assigned:      {assigned}")
print(f"  Unassigned:    {len(set(unassigned))}")
print(f"  Overlapping:   {len(overlapping)} (in 2+ SRA files)")

# ── STEP 4: Show unassigned and ask for input ─────────────────────────────────
unassigned_unique = list(set(unassigned))

if unassigned_unique:
    print()
    print("=" * 65)
    print(f"  {len(unassigned_unique)} SRS NUMBERS NOT FOUND IN ANY SRA FILE")
    print("=" * 65)
    print()

    # Get details from RAPID for each unassigned
    unassigned_details = rapid[
        rapid[rapid_srs_col].isin(unassigned_unique)
    ][[rapid_srs_col] + [c for c in ['PI', 'Department', 'College', 'Title']
                          if c in rapid.columns]].drop_duplicates(rapid_srs_col)

    for _, row in unassigned_details.iterrows():
        srs  = row[rapid_srs_col]
        pi   = row.get('PI', 'N/A')
        dept = row.get('Department', 'N/A')
        print(f"  SRS: {srs}")
        print(f"       PI: {pi}  |  Dept: {dept}")
        print()

    print("-" * 65)
    print("  Available SRAs loaded:")
    for i, name in enumerate(sorted(sra_srs_map.keys()), 1):
        print(f"    {i}. {name}")
    print()

    print("  For each unassigned SRS, enter the SRA name.")
    print("  Press Enter to leave as 'Unassigned'.")
    print("  Type 'all' + SRA name to assign ALL unassigned to one SRA.")
    print("-" * 65)
    print()

    # Check if user wants to assign all to one SRA
    bulk = input("  Assign ALL unassigned to one SRA? (type name or press Enter to assign individually): ").strip()

    if bulk:
        for srs in unassigned_unique:
            mask = rapid[rapid_srs_col] == srs
            rapid.loc[mask, 'SRA Name'] = bulk
        print(f"\n  All {len(unassigned_unique)} unassigned SRS assigned to: {bulk}")
    else:
        # Assign individually
        for srs in unassigned_unique:
            row_info = unassigned_details[unassigned_details[rapid_srs_col] == srs]
            if not row_info.empty:
                pi   = row_info.iloc[0].get('PI', '')
                dept = row_info.iloc[0].get('Department', '')
                print(f"  {srs}  |  PI: {pi}  |  Dept: {dept}")

            answer = input(f"  Assign to SRA (Enter = Unassigned): ").strip()
            if answer:
                mask = rapid[rapid_srs_col] == srs
                rapid.loc[mask, 'SRA Name'] = answer
                print(f"  → Assigned to: {answer}")
            else:
                print(f"  → Left as: Unassigned")
            print()

else:
    print("\n  All SRS numbers are assigned to an SRA.")

# ── STEP 5: Show overlapping ───────────────────────────────────────────────────
if overlapping:
    print()
    print("=" * 65)
    print(f"  {len(overlapping)} SRS NUMBERS APPEAR IN MULTIPLE SRA FILES")
    print("=" * 65)
    for srs, sras in overlapping:
        print(f"  {srs} → {' & '.join(sras)}")
    print()
    print("  These are kept as 'SRA1 | SRA2' in the output.")
    print("  Edit manually in Excel if needed.")

# ── STEP 5b: Show SRS in SRA files but NOT in RAPID ──────────────────────────
all_sra_srs = set()
for srs_set in sra_srs_map.values():
    all_sra_srs |= srs_set

in_sra_not_rapid = all_sra_srs - rapid_srs_all

if in_sra_not_rapid:
    print()
    print("=" * 65)
    print(f"  {len(in_sra_not_rapid)} SRS NUMBERS IN SRA FILES BUT NOT IN RAPID EXPORT")
    print("=" * 65)
    print()
    print("  These proposals exist in an SRA's file but are missing")
    print("  from the full RAPID export. Possible reasons:")
    print("  — RAPID export was filtered (date range, status, etc.)")
    print("  — Proposal was deleted or merged in RAPID")
    print("  — Wrong RAPID export file uploaded")
    print()

    # Show which SRA file each orphan came from
    for srs in sorted(in_sra_not_rapid):
        owners = [name for name, srs_set in sra_srs_map.items() if srs in srs_set]
        print(f"  {srs:<20} → Found in: {', '.join(owners)}")

    print()
    print("  These SRS numbers cannot be matched and will NOT appear")
    print("  in the output file. Check your RAPID export filters.")
    print("=" * 65)
else:
    print()
    print("  All SRS numbers in SRA files are present in the RAPID export.")

# ── STEP 6: Export ────────────────────────────────────────────────────────────
print()
print("[4] Exporting clean file...")

# Move SRA Name column to front (after SRS column)
cols = rapid.columns.tolist()
cols.remove('SRA Name')
srs_idx = cols.index(rapid_srs_col)
cols.insert(srs_idx + 1, 'SRA Name')
rapid = rapid[cols]

# Write output
with pd.ExcelWriter(OUTPUT_FILE, engine='openpyxl') as writer:
    rapid.to_excel(writer, sheet_name='RAPID with SRA', index=False)

    # Summary sheet
    summary_rows = []
    for sra, srs_set in sorted(sra_srs_map.items()):
        summary_rows.append({
            'SRA Name': sra,
            'Proposals in SRA file': len(srs_set),
            'Matched in RAPID': len(srs_set & rapid_srs_all)
        })

    # Add unassigned row
    final_unassigned = (rapid['SRA Name'] == 'Unassigned').sum()
    summary_rows.append({
        'SRA Name': 'UNASSIGNED',
        'Proposals in SRA file': 0,
        'Matched in RAPID': final_unassigned
    })

    pd.DataFrame(summary_rows).to_excel(writer, sheet_name='SRA Summary', index=False)

    # Orphan SRS sheet — in SRA files but not in RAPID
    if in_sra_not_rapid:
        orphan_rows = []
        for srs in sorted(in_sra_not_rapid):
            owners = [name for name, srs_set in sra_srs_map.items() if srs in srs_set]
            orphan_rows.append({
                'SRS #': srs,
                'Found in SRA File': ', '.join(owners),
                'In RAPID Export': 'NO — Missing from RAPID'
            })
        pd.DataFrame(orphan_rows).to_excel(
            writer, sheet_name='In SRA Not in RAPID', index=False)

# ── STEP 7: Final report ──────────────────────────────────────────────────────
print()
print("=" * 65)
print("  DONE")
print("=" * 65)
print()
print(f"  Output file:    {OUTPUT_FILE}")
print(f"  Total rows:     {len(rapid)}")
print(f"  Unique SRS:     {rapid[rapid_srs_col].nunique()}")
print()
print("  SRA breakdown:")
sra_counts = rapid['SRA Name'].value_counts()
for sra, count in sra_counts.items():
    flag = " ⚠ UNASSIGNED" if sra == "Unassigned" else ""
    print(f"    {sra:<35} {count:>4} rows{flag}")

print()
print("  Load this file into Power BI and refresh.")
print("=" * 65)
input("\n  Press Enter to exit.")
