# Fadama III Master Location Set

**File:** `FadamaIII_master_locations.csv` (also `FadamaIII_master_locations.xlsx` with this README as sheet 1)

**Version:** 1.0 · **Date:** 2026-11-19

## Purpose

Deduplicated master list of Fadama III geocoded locations from GODAD and AidData, with a confidence rating and notes from the Task 1 cross-check. Matched places appear once.

## Foundational sources

| Source | Path | Role |
|--------|------|------|
| GODAD | `fadama_task1/data/GODAD_Fadama_sites.csv` | World Bank `P096572` (218 points) + `P130788` AF (8 state-level) |
| AidData | `fadama_task1/data/level_1a.csv` | Project `NGA/000245` only (223 of ~1,400 Nigeria rows) |
| Task brief | `fadama_task1/data/fadama_tasks.pdf` | ClimAid Paper 2 tasks |
| Prior cross-check | `fadama_task1/output/FadamaIII_Task1_location_crosscheck.xlsx` | Full state/LGA comparison |

## Important context

The World Bank ICR reports Fadama III in **all 36 states + FCT** and about **685 LGAs**. This master set has **248 rows**. It is a **geocode sample**, not a complete site census. Use confidence flags accordingly.

## How rows were built

| `match_category` | Rule | `confidence` | `notes` |
|------------------|------|--------------|---------|
| `matched_both` | LGA name matched in both sources after normalization; one row; coordinates averaged | **high** | Not applicable. |
| `same_location_name_mismatch` | Mutual nearest unmatched LGAs in the same state within **5 km**; one row; both names retained | **moderate-high** | Same location by coordinates; names do not match. |
| `godad_only_aiddata_state_centroid` | GODAD LGA point where AidData has only a state centroid | **moderate** | One set does not fully confirm (no AidData LGA). |
| `aiddata_state_only_no_godad` | AidData state centroid; GODAD absent (Nasarawa) | **moderate** | One set only; not LGA-confirmed. |
| `godad_af_p130788_state` | Additional Financing state-level GODAD points | **moderate** | AF state marker; not AidData-cross-checked here. |
| `godad_only_unverified` / `aiddata_only_unverified` | No reasonable same-place match | **low** | Lower confidence due to no reasonable match. |

### Confidence scale

- **high** — both sources agree on LGA identity  
- **moderate-high** — coordinates strongly agree (&lt;5 km) but names differ  
- **moderate** — single-source LGA or state-level only; other source does not fully confirm  
- **low** — one source only, no reasonable coordinate match  

## Abbreviated analysis (Task 1)

- **159** LGAs matched by name in both sources (median coordinate gap **1.4 km**; **95.6%** within 25 km).
- **7** same-place name mismatches (&lt;5 km): e.g. Ife North↔Ife Central, Atisbo↔Tede, Illela↔Gada, Orlu↔Nkwerre, Isoko North↔Okpe, Mbaitoli↔Ihitte Ogwa, Igalamela-Odolu↔Idah.
- **28** GODAD LGA points in states where AidData provides only a state centroid (granularity difference, not necessarily a different site).
- **10** GODAD-only and **35** AidData-only with no reasonable match → real sample gaps / complementary coverage.
- GODAD missing **Nasarawa**; AidData missing **Borno**. WB documents indicate national participation, so these are geocode gaps relative to the ICR footprint.

## Column guide

| Column | Description |
|--------|-------------|
| `master_id` | Stable id (`F3-####`) |
| `match_category` | Classification above |
| `confidence` | high / moderate-high / moderate / low |
| `state`, `lga`, `lga_alternate` | Primary LGA (GODAD when available); alternate = AidData name when different |
| `place_name_aiddata` | AidData registry place name(s) |
| `latitude`, `longitude` | Preferred analysis coordinates (average when both sources present) |
| `latitude_*`, `longitude_*` | Source-specific coordinates |
| `coord_distance_km` | Distance between source points when both present |
| `sources` | Contributing foundational set(s) |
| `godad_location_ids`, `aiddata_location_ids` | Traceability to source row ids |
| `notes` | Match rationale, or `Not applicable.` for clean matches |
