# Free Classroom Finder — 2026-27 ODD

A standalone dashboard to find **empty classrooms by day and time**, grouped by building.

## Use it
Double-click **`dashboard.html`** — it opens in any browser, works **offline**, no server needed.

1. Pick a **day** (MON–SAT).
2. Pick **when**:
   - **Single slot** — one period (e.g. 10:00–11:00).
   - **Time range** — free across *all* selected consecutive periods (e.g. 10:00→13:00).
   - **Whole day** — no class at all that day.
3. Optional **filters**: limit to certain **buildings** or a **room type** (Theory Room-CST, -PSY, lab, etc.).
4. Hit **Find free rooms** — results are grouped by building, showing each free room number and type.

A room counts as **free only if the timetable cell is truly blank**. Anything written in
(course code, `Special Class`, `TBC`, a teacher name) is treated as occupied.

## Load a different Excel (no Python needed)
Click **📂 Load a different Excel** (top-right) and pick any `.xlsx` in the same room-allotment
format. The file is parsed **in your browser** — nothing is uploaded anywhere — and the whole
dashboard instantly switches to that file's rooms/buildings/timetable. Great for a new semester.

The in-browser parser mirrors `parse_rooms.py` exactly, so it expects the same layout:
each room is a block with a `Room No.…` header in column A, a `Days` row holding the time
labels, and `MON`–`SAT` rows below it.

## Refresh the baked-in data when the Excel changes
The dashboard's data is baked in from `Class room Allotment_2026-27_ODD.xlsx`. To rebuild after editing the Excel:

```bash
pip install openpyxl        # once
python parse_rooms.py       # regenerates rooms.json AND dashboard.html
```

## Files
- `dashboard.html` — the finished, shareable dashboard (data embedded).
- `template.html` — UI template (edit styling/layout here, then re-run the parser).
- `parse_rooms.py` — reads the Excel, writes `rooms.json`, injects it into the template.
- `rooms.json` — extracted timetable data (176 rooms, 6 buildings).

## Coverage
6 buildings parsed (the `INTERNSHIP` sheet is skipped — it isn't a room timetable):
UB-I Satyajit · UB-II Vidyasagar · UB-III Prafulla · UB-IV Jagadish · UB-V Rabindra · UB-VI Rammohan.
11 daily periods: 08:00–19:00.
