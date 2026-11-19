# Task 1 findings: Are Fadama III site locations reliable?

**Verdict: No — not as a complete site list.** GODAD and AidData are sparse LGA-scale geocode samples that agree with each other where they overlap, but both badly understate the project’s documented geography.

## Counts

| Source | Project | Locations | States/FCT |
|--------|---------|-----------|------------|
| GODAD | P096572 (Fadama III) | 218 | 36 (missing Nasarawa) |
| GODAD | P130788 (AF) | 8 (all state-level) | 8 |
| AidData | NGA/000245 | 223 | 36 (missing Borno) |
| World Bank ICR | — | ~685 LGAs | all 36 + FCT |

LGA comparison after name normalization: **159 in both**, **45 GODAD-only** (28 of those in states where AidData only has a state centroid), **42 AidData-only**.

## What the coordinates represent

- **GODAD:** `place_name` is blank for every row. Points are attributed to GADM LGAs from IATI, with precision codes 1 (97) and 2 (121). Treat as LGA-area points, not community GPS.
- **AidData:** Mostly precision 3 / ADM2 LGA names (176); some town seats (25); 21 state centroids.
- **Agreement:** For matched LGAs, median coordinate distance is **1.4 km**; **95.6%** within 25 km.

## Document cross-check

- PAD (2008): national design; up to ~560 LGAs; Nasarawa among the 19 new Fadama III states.
- ICR (2020): all 36 states + FCT; ~685 LGAs (~80% of LGAs); effectiveness 23 Mar 2009; closed 31 Dec 2019 (with AF).
- Therefore GODAD’s Plateau = 2 and Nasarawa = 0 (and AidData’s missing Borno) are **sample gaps**, not evidence the project skipped those places.

## Start dates

GODAD `startyear=2008` / `closingyear=2019` are project-level. No per-state or per-LGA start dates appear in these CSVs. WB effectiveness was March 2009.

## Implication for the drought–conflict analysis

Do not use GODAD points as the Fadama III treatment map or as intensity. Options:

1. State-level treatment after national rollout (all states + FCT), or
2. Obtain true LGA/community lists from M&E / ICR annexes / state reports, or
3. Keep these geocodes only as an incomplete observed subset with explicit missingness.

Full tables: `output/FadamaIII_Task1_location_crosscheck.xlsx`.
