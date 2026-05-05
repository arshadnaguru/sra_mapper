# SRA Mapper

Automated proposal assignment tool for university research administration offices.

Maps SRS numbers across multiple administrator portfolios against a full institutional 
research export, flags unassigned and orphan records, and outputs a Power BI-ready 
Excel with a clean SRA assignment column.

---

## Folder Setup
sra_mapper/
├── sra_mapper.py           ← main script
├── RUN_SRA_MAPPER.bat      ← double click to run
├── rapid_export.xlsx       ← drop your institutional export here
└── sra_files/              ← drop all administrator Excel files here
## Requirements

- Python 3.9+
- pandas
- openpyxl

Install dependencies:
```bash
pip install pandas openpyxl
```

## How to Use

1. Drop your full institutional export as `rapid_export.xlsx`
2. Drop all administrator Excel files into `sra_files/` folder
   — filename of each Excel file is used as the administrator name
3. Double click `RUN_SRA_MAPPER.bat`
4. Review flagged records in the terminal
5. Assign unassigned records when prompted
6. Load the output Excel into Power BI

## What It Does

- Reads all administrator files from the folder automatically
- Extracts administrator name directly from the filename
- Matches proposal IDs between administrator files and the full export
- Creates a new assignment column right next to the proposal ID column
- Flags proposals in the full export with no administrator assigned
- Flags proposals in administrator files missing from the full export
- Detects proposals assigned to multiple administrators
- Prompts user to assign unassigned proposals individually or in bulk
- Outputs a timestamped Excel file ready for Power BI

## Output

| Sheet | Contents |
|---|---|
| RAPID with SRA | Full export with administrator assignment column added |
| SRA Summary | Proposal count per administrator |
| In SRA Not in RAPID | Orphan records requiring investigation |

Output filename: `rapid_with_sra_YYYYMMDD_HHMM.xlsx`
