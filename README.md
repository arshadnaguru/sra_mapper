<<<<<<< HEAD
# SRA Mapper — RIT OVPR

Maps SRS numbers from SRA files to the full RAPID export and creates a new SRA Name column.

## Folder setup

```
sra_mapper/
├── sra_mapper.py           ← main script
├── RUN_SRA_MAPPER.bat      ← double click to run
├── rapid_export.xlsx       ← drop your full RAPID export here
└── sra_files/
    ├── Maria_Cortes.xlsx
    ├── Stacey_Fisher.xlsx
    ├── April_Burns.xlsx
    ├── KwokKeung_Koo.xlsx
    ├── Dawid_Grames.xlsx
    ├── Vandezande_Sharon.xlsx
    └── Brittany_Neyland.xlsx
```

## How to use

1. Drop full RAPID export as `rapid_export.xlsx`
2. Drop all SRA Excel files into `sra_files/` folder
3. Double click `RUN_SRA_MAPPER.bat`
4. For any unassigned SRS — type the SRA name when prompted
5. Load the output Excel into Power BI

## What it does

- Reads all SRA files from the folder automatically
- Extracts SRA name from the filename
- Matches SRS numbers between SRA files and RAPID export
- Creates a new `SRA Name` column in the RAPID export
- Pops up unassigned SRS with PI and Department details
- Lets you assign unassigned SRS individually or all at once
- Flags overlapping SRS (same proposal in 2+ SRA files)
- Outputs one clean Excel ready for Power BI

## Output

- `rapid_with_sra_YYYYMMDD_HHMM.xlsx` — RAPID export with SRA Name column added
- Sheet 1: Full data with SRA Name column right after SRS column
- Sheet 2: SRA summary — proposal count per SRA
=======
# sra_mapper
Automated SRA-to-proposal mapping tool for research administration - matches SRS numbers across multiple administrator portfolios against a full RAPID export, flags unassigned and orphan records, and outputs a Power BI-ready Excel with SRA assignments.
>>>>>>> 260887aa3008d5ace9de523bf8522acf19c5c355
