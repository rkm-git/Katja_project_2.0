#!/usr/bin/env python3
"""Cross-reference GODAD Fadama III sites with AidData NGA/000245."""

from __future__ import annotations

import csv
import json
import math
import os
import re
import unicodedata
from collections import Counter, defaultdict
from difflib import SequenceMatcher
from io import StringIO
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

ROOT = Path(__file__).resolve().parent
OUT = ROOT / "output"
DEFAULT_GODAD = ROOT / "data" / "GODAD_Fadama_sites.csv"
DEFAULT_AID = ROOT / "data" / "level_1a.csv"

ALIASES = {
    ("abia", "osisioma"): "osisioma ngwa",
    ("abia", "umuahia"): "umuahia north",
    ("akwa ibom", "utu etim ekpo"): "etim ekpo",
    ("akwa ibom", "eman uruan"): "uruan",
    ("akwa ibom", "oruko"): "urue offong oruko",
    ("anambra", "nnewi"): "nnewi south",
    ("kano", "makuda"): "makoda",
    ("rivers", "ahoada"): "ahoada east",
    ("rivers", "emuoha"): "emohua",
    ("benue", "markurdi"): "makurdi",
    ("benue", "oturkpo"): "otukpo",
    ("benue", "vandeikya"): "vande ikya",
    ("benue", "bukuru"): "buruku",
    ("katsina", "dutsin ma"): "dutsinma",
    ("imo", "mbatoli"): "mbaitoli",
}

TRUNC = {
    "ogba egbe": "ogba egbema ndoni",
    "isokonor": "isoko north",
    "isoko nor": "isoko north",
    "nnewisou": "nnewi south",
    "nnewi sou": "nnewi south",
    "abakalik": "abakaliki",
    "konshish": "konshisha",
    "tambawal": "tambuwal",
    "quaan pan": "qua an pan",
    "urue offong": "urue offong oruko",
    "mbatoli": "mbaitoli",
}

WB_STATES = {
    "abia",
    "adamawa",
    "akwa ibom",
    "anambra",
    "bauchi",
    "bayelsa",
    "benue",
    "borno",
    "cross river",
    "delta",
    "ebonyi",
    "edo",
    "ekiti",
    "enugu",
    "gombe",
    "imo",
    "jigawa",
    "kaduna",
    "kano",
    "katsina",
    "kebbi",
    "kogi",
    "kwara",
    "lagos",
    "nasarawa",
    "niger",
    "ogun",
    "ondo",
    "osun",
    "oyo",
    "plateau",
    "rivers",
    "sokoto",
    "taraba",
    "yobe",
    "zamfara",
    "federal capital territory",
}


def unwrap_line(line: str) -> str | None:
    line = line.strip()
    if not line:
        return None
    if line.endswith(";;"):
        line = line[:-2]
    if line.startswith('"') and line.endswith('"'):
        return line[1:-1].replace('""', '"')
    return line


def norm(s: str | None) -> str:
    if s is None:
        return ""
    s = str(s).strip().lower()
    s = unicodedata.normalize("NFKD", s)
    s = s.replace("state", "").replace("lga", "")
    s = s.replace("/", " ").replace("-", " ").replace("_", " ")
    s = re.sub(r"[^a-z0-9\s]", "", s)
    s = re.sub(r"\s+", " ", s).strip()
    if s in {"fct", "abuja"}:
        return "federal capital territory"
    s2 = s.replace(" ", "")
    if s2 == "akwaibom":
        return "akwa ibom"
    if s2 == "crossriver":
        return "cross river"
    return s


def norm_lga(s: str | None, state: str | None = None) -> str:
    s = norm(s)
    if s in TRUNC:
        s = TRUNC[s]
    if state and (state, s) in ALIASES:
        s = ALIASES[(state, s)]
    return s


def parse_eu_float(x) -> float | None:
    if x is None or x == "":
        return None
    x = str(x).strip().replace(" ", "")
    if "," in x and "." in x:
        x = x.replace(".", "").replace(",", ".")
    elif "," in x:
        x = x.replace(",", ".")
    try:
        return float(x)
    except ValueError:
        return None


def haversine_km(lat1, lon1, lat2, lon2) -> float | None:
    if None in (lat1, lon1, lat2, lon2):
        return None
    r = 6371.0
    p1, p2 = math.radians(lat1), math.radians(lat2)
    dp = math.radians(lat2 - lat1)
    dl = math.radians(lon2 - lon1)
    a = math.sin(dp / 2) ** 2 + math.cos(p1) * math.cos(p2) * math.sin(dl / 2) ** 2
    return 2 * r * math.asin(math.sqrt(a))


def parse_adm(gaz: str) -> tuple[str, str]:
    parts = [p.strip() for p in (gaz or "").split("|")]
    state = lga = ""
    if "Nigeria" in parts:
        i = parts.index("Nigeria")
        if len(parts) > i + 1:
            state = parts[i + 1]
        if len(parts) > i + 2:
            lga = parts[i + 2]
    return state, lga


def load_godad(path: Path) -> list[dict]:
    with path.open(newline="", encoding="utf-8-sig") as f:
        rows = list(csv.DictReader(f, delimiter=";"))
    for r in rows:
        r["lat"] = parse_eu_float(r["latitude"])
        r["lon"] = parse_eu_float(r["longitude"])
        r["state_n"] = norm(r.get("state_gadm41") or r.get("state_godad(gadm36)"))
        r["lga_n"] = norm_lga(r.get("lga_gadm41") or r.get("lga_godad(gadm36)"), r["state_n"])
        if r["state_n"] == "benue" and norm(r.get("lga_godad(gadm36)")) == "buruku":
            r["lga_n"] = "buruku"
    return rows


def load_aiddata(path: Path) -> list[dict]:
    raw = path.read_text(encoding="utf-8-sig").splitlines()
    recon = [unwrap_line(l) for l in raw]
    recon = [l for l in recon if l]
    rows = [
        r
        for r in csv.DictReader(StringIO("\n".join(recon)))
        if r.get("project_id") == "NGA/000245"
    ]
    for r in rows:
        r["lat"] = parse_eu_float(r["latitude"])
        r["lon"] = parse_eu_float(r["longitude"])
        st, lg = parse_adm(r.get("gazetteer_adm_name", ""))
        pcode = r.get("precision_code")
        ltype = r.get("location_type_code")
        if ltype == "ADM1" or pcode == "4":
            r["state_n"] = norm(r.get("place_name") or st)
            r["lga_n"] = ""
            r["level"] = "state"
        else:
            r["state_n"] = norm(st)
            raw_lga = (r.get("place_name") or lg) if ltype == "ADM2" else (lg or r.get("place_name"))
            r["lga_n"] = norm_lga(raw_lga, r["state_n"])
            r["level"] = "site"
    return rows


def fuzzy_match(only_g, only_a, thr=0.84):
    matches = []
    used = set()
    for gk in sorted(only_g):
        gs, gl = gk.split("|", 1)
        best = None
        best_sc = 0.0
        for ak in only_a:
            if ak in used:
                continue
            as_, al = ak.split("|", 1)
            if gs != as_:
                continue
            sc = max(
                SequenceMatcher(None, gl, al).ratio(),
                SequenceMatcher(None, gl.replace(" ", ""), al.replace(" ", "")).ratio(),
            )
            if gl in al or al in gl:
                sc = max(sc, 0.9)
            if sc > best_sc:
                best_sc = sc
                best = ak
        if best and best_sc >= thr:
            matches.append((gk, best, best_sc))
            used.add(best)
    return matches


def key(state: str, lga: str) -> str:
    return f"{state}|{lga}"


def write_excel(godad, aid, matched_pairs, only_g2, only_a2, godad_lga, aid_lga, dists, states_aid_has_lga, stats):
    g_keys = set(godad_lga)
    a_keys = set(aid_lga)
    g_states = set(r["state_n"] for r in godad if r["project_id"] == "P096572")
    a_states = set(r["state_n"] for r in aid)
    median_d = stats["median_dist_km"]

    wb = Workbook()
    ws = wb.active
    ws.title = "README"
    for row in [
        ["Task 1: Fadama III site location reliability — GODAD vs AidData cross-reference"],
        ["Date", "2026-11-18"],
        ["Sources"],
        ["GODAD", "GODAD_Fadama_sites.csv — P096572 (218) + P130788 AF (8 state-level)"],
        ["AidData", "level_1a.csv NGA/000245 (223)"],
        ["WB docs", "PAD 2008; ICR 2020; ISR 2015; WB 2016 case study"],
        [],
        ["Verdict"],
        ["Both sources are sparse LGA-scale geocode samples (~220 points), NOT a complete Fadama III site census."],
        ["WB ICR: all 36 states + FCT; ~685 LGAs (~80% of LGAs). Samples undercover by an order of magnitude."],
        ["LGA names agree where both list them; coordinates usually match within ~25 km."],
        ["GODAD missing Nasarawa; AidData missing Borno — sample gaps, not documented non-participation."],
        ["Do not interpret state location counts (e.g. Plateau=2) as project intensity."],
    ]:
        ws.append(row)
    ws["A1"].font = Font(bold=True, size=14)

    ws = wb.create_sheet("State_summary")
    header = [
        "state",
        "godad_locations_P096572",
        "godad_unique_lgas",
        "aiddata_locations",
        "aiddata_unique_lgas",
        "aiddata_has_lga_level",
        "lgas_in_both",
        "lgas_godad_only",
        "lgas_aiddata_only",
        "agreement_note",
    ]
    ws.append(header)
    state_rows = []
    for s in sorted(WB_STATES):
        gl = {k.split("|", 1)[1] for k in g_keys if k.startswith(s + "|")}
        al = {k.split("|", 1)[1] for k in a_keys if k.startswith(s + "|")}
        both_s = set()
        go = set(gl)
        ao = set(al)
        for g, a, _ in matched_pairs:
            gs, glga = g.split("|", 1)
            _, alga = a.split("|", 1)
            if gs != s:
                continue
            both_s.add(glga)
            go.discard(glga)
            ao.discard(alga)
        g_loc = sum(1 for r in godad if r["project_id"] == "P096572" and r["state_n"] == s)
        a_loc = sum(1 for r in aid if r["state_n"] == s)
        has_lga = s in states_aid_has_lga
        if s == "nasarawa":
            note = "GODAD ZERO sites; AidData state-only. PAD: Nasarawa in 19 new states; ICR national. Sample gap."
        elif s == "borno":
            note = "GODAD has Maiduguri; AidData has NO Borno row. ICR includes Borno. Sample gap in AidData."
        elif s == "plateau":
            note = "GODAD 2 LGAs; AidData state-only. Sparse sample, not true site count."
        elif not has_lga and g_loc > 0:
            note = "GODAD has LGA points; AidData only state-level."
        elif has_lga and len(both_s) >= max(len(gl), 1) * 0.7:
            note = "Strong LGA overlap."
        elif has_lga and len(both_s) > 0:
            note = "Partial LGA overlap; residuals mostly naming/town-vs-LGA."
        else:
            note = "Little/no LGA overlap or state-only samples."
        row = [
            s,
            g_loc,
            len(gl),
            a_loc,
            len(al),
            "yes" if has_lga else ("no (state only)" if a_loc else "none"),
            len(both_s),
            len(go),
            len(ao),
            note,
        ]
        ws.append(row)
        state_rows.append(row)

    ws = wb.create_sheet("LGA_comparison")
    ws.append(
        [
            "match_status",
            "match_how",
            "state",
            "lga_godad",
            "lga_aiddata",
            "godad_n",
            "aiddata_n",
            "aiddata_place_names",
            "godad_precision",
            "aiddata_precision",
            "coord_distance_km",
            "note",
        ]
    )
    dist_map = {d[1]: d[0] for d in dists}
    for g, a, how in sorted(matched_pairs, key=lambda x: x[0]):
        gs, gl = g.split("|", 1)
        _, al = a.split("|", 1)
        d = dist_map.get(g if g == a else g + "~" + a)
        ws.append(
            [
                "both",
                how,
                gs,
                gl,
                al,
                len(godad_lga[g]),
                len(aid_lga[a]),
                "; ".join(sorted(set(x.get("place_name") for x in aid_lga[a]))),
                ";".join(sorted(set(x["precision_code"] for x in godad_lga[g]))),
                ";".join(sorted(set(x.get("precision_code") for x in aid_lga[a]))),
                round(d, 1) if d is not None else "",
                "Large coordinate gap" if d and d > 50 else "",
            ]
        )
    for g in sorted(only_g2):
        gs, gl = g.split("|", 1)
        ws.append(
            [
                "GODAD_only",
                "",
                gs,
                gl,
                "",
                len(godad_lga[g]),
                0,
                "",
                ";".join(sorted(set(x["precision_code"] for x in godad_lga[g]))),
                "",
                "",
                "AidData state-only for this state"
                if gs not in states_aid_has_lga
                else "Not in AidData LGA list",
            ]
        )
    for a in sorted(only_a2):
        as_, al = a.split("|", 1)
        ws.append(
            [
                "AidData_only",
                "",
                as_,
                "",
                al,
                0,
                len(aid_lga[a]),
                "; ".join(sorted(set(x.get("place_name") for x in aid_lga[a]))),
                "",
                ";".join(sorted(set(x.get("precision_code") for x in aid_lga[a]))),
                "",
                "Town/community or residual name mismatch",
            ]
        )

    ws = wb.create_sheet("Precision_and_coords")
    ws.append(["Question", "Finding"])
    ws.append(["GODAD place names?", "Blank for all GODAD rows — no community names."])
    ws.append(
        [
            "GODAD precision P096572",
            f"{stats['precision_godad'].get('1', 0)} code 1; {stats['precision_godad'].get('2', 0)} code 2. Treat as LGA-area IATI points.",
        ]
    )
    ws.append(
        [
            "AidData precision",
            f"{stats['precision_aid'].get('3', 0)} code 3 (LGA); {stats['precision_aid'].get('1', 0)} code 1; {stats['precision_aid'].get('4', 0)} code 4 (state).",
        ]
    )
    ws.append(
        [
            "Matched distances",
            f"n={len(dists)}; median {median_d} km; {stats['pct_within_25km']}% within 25 km.",
        ]
    )
    ws.append(
        [
            "P130788 AF",
            "8 state-level (code 4), startyear 2013: Anambra, Enugu, Lagos, Ebonyi, Kogi, Niger, Kano, Ogun.",
        ]
    )
    ws.append(
        [
            "Start dates",
            "GODAD startyear=2008 / closing=2019 project-level. WB effectiveness 23 Mar 2009; actual close 31 Dec 2019.",
        ]
    )

    ws = wb.create_sheet("Documents_vs_samples")
    ws.append(["Document claim", "Implication"])
    ws.append(["PAD 2008: ~560 LGAs design; Nasarawa in 19 new states", "Geocode samples << design footprint"])
    ws.append(
        [
            "ICR: 36 states+FCT; ~685 LGAs; ~80% LGAs; eff. 2009; close 2019",
            "National coverage; sample gaps ≠ non-participation",
        ]
    )
    ws.append(["Uneven sample counts", "Not usable as treatment intensity"])

    ws = wb.create_sheet("Task1_findings")
    for row in [
        ["Task 1 findings summary"],
        [],
        ["Reliability verdict: NOT reliable as a complete site list. Partial LGA-scale sample with known gaps."],
        [
            f"Overlap: {stats['both_lgas']} LGAs in both; {stats['godad_only']} GODAD-only "
            f"({stats['godad_only_statelevel_artifact']} because AidData is state-only there); "
            f"{stats['aiddata_only']} AidData-only."
        ],
        ["GODAD missing Nasarawa; AidData missing Borno."],
        ["Coordinates: LGA-scale agreement; not community GPS."],
        ["Documents overturn “Plateau only 2 / Nasarawa none” as substantive geography."],
    ]:
        ws.append(row)
    ws["A1"].font = Font(bold=True, size=14)

    ws = wb.create_sheet("Documents_to_find")
    ws.append(["doc_id", "project", "title", "url_or_path", "status", "short_filename", "notes"])
    for d in [
        [
            "PAD_P096572",
            "P096572",
            "PAD",
            "https://documents1.worldbank.org/curated/en/268821468110347872/pdf/394890PAD0P0961nly10IDAR20081022111.pdf",
            "reviewed",
            "WB_P096572_PAD.pdf",
            "Nasarawa in 19 new states",
        ],
        [
            "ICR_P096572",
            "P096572",
            "ICR",
            "https://documents1.worldbank.org/curated/en/540261604338534589/pdf/Nigeria-Third-National-Fadama-Development-Project.pdf",
            "reviewed",
            "WB_P096572_ICR.pdf",
            "36 states+FCT; ~685 LGAs",
        ],
        [
            "ISR_2015",
            "P096572",
            "ISR 2015",
            "https://documents1.worldbank.org/curated/en/171611468290157111/pdf/ISR-Disclosable-P096572-02-20-2015-1424436406638.pdf",
            "reviewed",
            "WB_P096572_ISR_2015.pdf",
            "",
        ],
        [
            "WB_2016",
            "series",
            "Case study",
            "https://documents1.worldbank.org/curated/en/276351467990996530/pdf/The-Nigeria-Fadama-National-Development-Series-How-to-Build-a-Pilot-into-a-National-Program-Through-Learning-and-Adaptation-Delivery-Case-Study.pdf",
            "reviewed",
            "WB_2016.pdf",
            "",
        ],
        [
            "AF_P130788",
            "P130788",
            "AF docs",
            "https://projects.worldbank.org/en/projects-operations/project-detail/P130788",
            "to collect",
            "WB_P130788_*.pdf",
            "",
        ],
        [
            "M&E_site_lists",
            "P096572",
            "M&E/beneficiary/ESIA site lists",
            "WB docs page",
            "to collect",
            "",
            "Needed for true LGA/community lists",
        ],
    ]:
        ws.append(d)

    header_fill = PatternFill("solid", fgColor="1F4E79")
    header_font = Font(bold=True, color="FFFFFF")
    for sheet_name in ["State_summary", "LGA_comparison", "Documents_to_find"]:
        ws = wb[sheet_name]
        for cell in ws[1]:
            cell.fill = header_fill
            cell.font = header_font
        ws.freeze_panes = "A2"
        for col in ws.columns:
            maxlen = min(60, max((len(str(c.value)) if c.value else 0) for c in col))
            ws.column_dimensions[get_column_letter(col[0].column)].width = max(12, maxlen + 2)

    fills = {
        "both": PatternFill("solid", fgColor="C6EFCE"),
        "GODAD_only": PatternFill("solid", fgColor="FFEB9C"),
        "AidData_only": PatternFill("solid", fgColor="FFC7CE"),
    }
    ws = wb["LGA_comparison"]
    for row in ws.iter_rows(min_row=2, max_col=1):
        if row[0].value in fills:
            for c in ws[row[0].row]:
                c.fill = fills[row[0].value]

    OUT.mkdir(parents=True, exist_ok=True)
    path = OUT / "FadamaIII_Task1_location_crosscheck.xlsx"
    wb.save(path)
    with (OUT / "state_comparison.csv").open("w", newline="") as f:
        w = csv.writer(f)
        w.writerow(header)
        w.writerows(state_rows)
    return state_rows, path


def main():
    godad_path = Path(os.environ.get("GODAD_CSV", DEFAULT_GODAD))
    aid_path = Path(os.environ.get("AIDDATA_CSV", DEFAULT_AID))
    godad = load_godad(godad_path)
    aid = load_aiddata(aid_path)

    godad_lga = defaultdict(list)
    for r in godad:
        if r["project_id"] == "P096572":
            godad_lga[key(r["state_n"], r["lga_n"])].append(r)
    aid_lga = defaultdict(list)
    for r in aid:
        if r["lga_n"]:
            aid_lga[key(r["state_n"], r["lga_n"])].append(r)

    g_keys = set(godad_lga)
    a_keys = set(aid_lga)
    both = g_keys & a_keys
    only_g = g_keys - a_keys
    only_a = a_keys - g_keys
    fm = fuzzy_match(only_g, only_a)
    both2 = both | {m[0] for m in fm}
    only_g2 = only_g - {m[0] for m in fm}
    only_a2 = only_a - {m[1] for m in fm}
    states_aid_has_lga = {k.split("|")[0] for k in a_keys}
    godad_only_state = [k for k in only_g2 if k.split("|")[0] not in states_aid_has_lga]
    godad_only_true = [k for k in only_g2 if k.split("|")[0] in states_aid_has_lga]
    matched_pairs = [(k, k, "exact") for k in both] + [
        (g, a, f"fuzzy:{sc:.2f}") for g, a, sc in fm
    ]

    dists = []
    for g, a, _ in matched_pairs:
        gpts = godad_lga[g]
        apts = aid_lga[a]
        glat = sum(p["lat"] for p in gpts) / len(gpts)
        glon = sum(p["lon"] for p in gpts) / len(gpts)
        alat = sum(p["lat"] for p in apts) / len(apts)
        alon = sum(p["lon"] for p in apts) / len(apts)
        d = haversine_km(glat, glon, alat, alon)
        if d is not None:
            dists.append((d, g if g == a else g + "~" + a))
    median_d = sorted(x[0] for x in dists)[len(dists) // 2]

    g_states = {r["state_n"] for r in godad if r["project_id"] == "P096572"}
    a_states = {r["state_n"] for r in aid}
    stats = {
        "godad_p096572": sum(1 for r in godad if r["project_id"] == "P096572"),
        "godad_p130788": sum(1 for r in godad if r["project_id"] == "P130788"),
        "aiddata": len(aid),
        "both_lgas": len(both2),
        "godad_only": len(only_g2),
        "aiddata_only": len(only_a2),
        "godad_only_statelevel_artifact": len(godad_only_state),
        "godad_only_true": len(godad_only_true),
        "median_dist_km": round(median_d, 1),
        "pct_within_25km": round(100 * sum(1 for d in dists if d[0] < 25) / len(dists), 1),
        "godad_states": len(g_states),
        "aid_states": len(a_states),
        "godad_missing": sorted(WB_STATES - g_states),
        "aid_missing": sorted(WB_STATES - a_states),
        "precision_godad": dict(
            Counter(r["precision_code"] for r in godad if r["project_id"] == "P096572")
        ),
        "precision_aid": dict(Counter(r.get("precision_code") for r in aid)),
        "godad_only_true_list": sorted(godad_only_true),
        "aiddata_only_list": sorted(only_a2),
    }

    state_rows, path = write_excel(
        godad,
        aid,
        matched_pairs,
        only_g2,
        only_a2,
        godad_lga,
        aid_lga,
        dists,
        states_aid_has_lga,
        stats,
    )
    stats["state_rows"] = [
        {
            "state": r[0],
            "g_loc": r[1],
            "g_lga": r[2],
            "a_loc": r[3],
            "a_lga": r[4],
            "both": r[6],
            "g_only": r[7],
            "a_only": r[8],
            "note": r[9],
        }
        for r in state_rows
    ]
    (OUT / "stats.json").write_text(json.dumps(stats, indent=2))
    print(f"Wrote {path}")
    print(
        json.dumps(
            {k: v for k, v in stats.items() if k not in ("state_rows", "godad_only_true_list", "aiddata_only_list")},
            indent=2,
        )
    )


if __name__ == "__main__":
    main()
