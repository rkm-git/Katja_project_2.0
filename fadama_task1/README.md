# Fadama III Task 1 — Location reliability cross-check

Cross-references GODAD Fadama III geocodes (`P096572` / `P130788`) with AidData Nigeria `level_1a` project `NGA/000245`, and checks World Bank PAD/ICR coverage claims.

## Verdict

**Neither GODAD nor AidData is a reliable complete site list.** Both are sparse LGA-scale samples (~220 points) versus ~685 LGAs in the ICR. Where both list an LGA, names and coordinates usually agree; uneven state counts (e.g. Plateau = 2, Nasarawa = 0 in GODAD) are sampling artefacts relative to documented national rollout.

## Outputs

- `output/FadamaIII_master_locations.csv` — **master location set** (deduplicated; confidence + notes)
- `output/FadamaIII_master_locations.xlsx` — same data with README sheet
- `output/MASTER_LOCATIONS_README.md` — documentation for the master set
- `output/FadamaIII_Task1_location_crosscheck.xlsx` — Task 1 cross-check workbook
- `output/state_comparison.csv`, `output/stats.json` — machine-readable summaries
- `compare_godad_aiddata.py` — reproducible comparison script

## Reproduce

```bash
python3 fadama_task1/compare_godad_aiddata.py
```

Input CSVs are read from the Cursor uploads path used in this run (or set `GODAD_CSV` / `AIDDATA_CSV`).
