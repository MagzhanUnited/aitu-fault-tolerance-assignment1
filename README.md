# Assignment 1 — Reliability Engineering and Failure Analysis

**Astana IT University · Fault Tolerance · 2026–2027**
Student: Zhumabayev Magzhan (ID 255408, group CSE-2506) · Instructor: Serek Azamat

Synthetic AITU student-portal reliability analysis: availability metrics, Reliability Block Diagram, FMEA and Fault Tree Analysis.

## Files
| File | Purpose |
|---|---|
| `reliability_analysis.py` | Single script that reproduces every number in the report (Parts A–D), exports the workbook and draws all figures |
| `reliability_calculations.xlsx` | Excel workbook with live formulas (sheets PartA_Metrics, PartB_RBD, PartC_FMEA, PartD_FTA) |
| `run_log.txt` | Console output of the script (intermediate calculations) |
| `results.json` | Machine-readable results used to build the report |
| `rbd_diagram.png`, `fta_diagram.png` | RBD and FTA diagrams |
| `downtime_by_component.png`, `failure_timeline.png` | Supporting figures for Part A |
| `.github/workflows/build.yml` | GitHub Actions workflow that re-runs the script on every push and commits the regenerated artifacts (reproducibility evidence — see the Actions tab) |

## Run
```bash
pip install matplotlib openpyxl
python reliability_analysis.py      # -> xlsx, png, run_log.txt, results.json
```

## Key results
- Downtime 27 h / uptime 693 h of 720 h → availability **96.25 %**
- MTTF 57.75 h · MTTR 2.25 h · MTBF 60 h
- RBD: R_sys = 0.995 × 0.9991 × 0.9998 × 0.995 = **0.9889**
- FMEA top RPN: Primary DB node failure (252), DB data corruption/replication lag (216), standby not promotable (216)
- FTA: q_TOP = **0.0111** (consistent with 1 − R_sys)

_All data are synthetic educational data and do not describe real AITU infrastructure._
