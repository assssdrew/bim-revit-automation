# Navisworks clash tolerances (draft matrix)

Reference draft for Clash Detective: what to test, which tolerances to use at **stage P** vs **stage R** (Russian PD / WD), and what to **exclude**.

> Discussion source: an internal “collision tolerances” table (not yet split by P/R).  
> Numbers below are a **starting point for agreement**, not a company standard. Replace with values from the project DOCX after review.

Related portfolio rule: [SAFETY.md](SAFETY.md) — without tolerances, reports become noise.

Russian version (primary for the team): [navisworks-clash-tolerances.ru.md](navisworks-clash-tolerances.ru.md)

---

## How to read the matrix

| Column | Meaning |
|--------|---------|
| **Pair** | Selection Set A ↔ Selection Set B (never “whole model ↔ whole model”) |
| **Type** | Hard / Hard (Conservative) / Clearance / Duplicates |
| **P** | Design documentation stage: catch systemic clashes; tolerate modelling noise |
| **R** | Working documentation: tighter hard + meaningful clearance for install/access |
| **Exclude** | Ignore rules / category filters — otherwise false New floods |

Operator rule: **Hard** on critical pairs first → group → coordinate → then **Clearance**. Do not dump everything into one test.

---

## Hard clashes

Hard tolerance = ignore intersections smaller than N mm (rounding / touching noise).

| Pair (A ↔ B) | Type | Tol. P | Tol. R | Note |
|--------------|------|--------|--------|------|
| Structure ↔ HVAC ducts | Hard | 25 mm | 10 mm | Structure wins; tighten on R |
| Structure ↔ Plumbing/Piping | Hard | 25 mm | 10 mm | |
| Structure ↔ Cable trays / busbars | Hard | 25 mm | 10 mm | |
| Structure ↔ Fire protection piping | Hard | 25 mm | 10 mm | |
| HVAC ↔ Plumbing | Hard | 20 mm | 10 mm | Cross-discipline |
| HVAC ↔ Electrical | Hard | 20 mm | 10 mm | |
| Plumbing ↔ Electrical | Hard | 20 mm | 10 mm | |
| Architecture walls/slabs ↔ MEP mains | Hard | 25 mm | 10 mm | Do not confuse with intentional penetrations |
| Architecture ↔ Structure | Hard | 15 mm | 5–10 mm | Often LOD noise |
| Equipment ↔ Structure/Architecture | Hard | 15 mm | 5–10 mm | |
| Per-discipline duplicates | Duplicates | 10 mm | 5 mm | Copied elements |
| Per-discipline self Hard | Hard | — | per BEP | Only if explicit QA; noisy otherwise |

---

## Clearance

Clearance fires when distance is **below** the set gap.

| Pair | Gap P | Gap R | Why |
|------|-------|-------|-----|
| Insulated duct ↔ structure | off* | 50 mm | Install / insulation |
| Insulated pipe ↔ structure | off* | 50 mm | |
| Tray ↔ duct | off* | 50–100 mm | Access / hangers |
| Sprinkler ↔ beam soffit | code / off | 100–150 mm | Confirm against project codes |
| Equipment ↔ service zone | schematic | 600–1000 mm† | Prefer explicit access sets |
| Tray ↔ hot piping | off | 100+ mm | Heat / access |

\* On stage P, clearance is often **off** or limited to shafts/plant rooms.  
† Service clearances come from equipment data sheets / client brief.

---

## Must-exclude list

1. Rebar ↔ concrete (when both modelled)  
2. Fireproofing ↔ structure  
3. Pipe/duct insulation ↔ host (if clearance already accounts for it)  
4. Bolts / welds / connection plates ↔ parent members  
5. Curtain-wall mullions ↔ panels within the same system  
6. Floor/ceiling finishes ↔ host slab (separate build-up)  
7. FF&E ↔ architecture in MEP coordination hard tests  
8. Rooms/Spaces, areas, grids, levels, 2D annotation  
9. Temporary / to-be-demolished phase elements (per project phasing)  
10. Double-loaded links in the NWF (fix the federation)  
11. Intentional sleeves / openings / niches ↔ services (tracked status, not New hard)  
12. Hangers/clamps modelled into structure  
13. Embedded plates under equipment  
14. Intra-system fittings at junctions (self-noise)  
15. Previously **Approved/Reviewed** clashes without geometry change  
16. Intra-set clashes when the test is cross-discipline  
17. Below-significance sizes on stage P if the BEP says so  

---

## Suggested additions if missing from the DOCX

| Topic | Why |
|-------|-----|
| Explicit **P / R** columns or sheets | Stated gap in the team chat |
| Clash status workflow (New → … → Resolved) | Do not confuse tolerance with approval |
| Grouping rules (level / system / grid) | One duct × ten beams = one issue |
| Separate tests for shafts / plant / roof / parking | Different criticality |
| Clearance **R-only** + pair list | Keep P reports readable |
| Exclusions: rebar / fireproofing / insulation / sleeves | Top false-positive sources |
| Fire protection ↔ all; ELV ↔ power | Often missing in v1 matrices |
| When to use Hard (Conservative) | Near-miss parallel elements |
| Owner discipline per pair | Assignment in Clash Detective |
| Named Search/Selection Sets | Matrix must be reproducible |
| Matrix version + date + project | Avoid divergent NWF setups |

---

## Minimal starter test pack

1. Structure ↔ HVAC (Hard)  
2. Structure ↔ Plumbing (Hard)  
3. Structure ↔ Electrical (Hard)  
4. HVAC ↔ Plumbing (Hard)  
5. HVAC ↔ Electrical (Hard)  
6. Architecture ↔ MEP mains (Hard)  
7. Duplicates per discipline (spot)  
8. On R: 2–3 Clearance tests per BEP  

---

## Sanitisation

No corporate UNC paths, IPs, or personal folders in this public repo. Keep the live DOCX on the internal share; keep only the anonymised matrix here.

After the actual DOCX/screenshots are provided, replace draft millimetres with project values and mark “already present / proposed”.
