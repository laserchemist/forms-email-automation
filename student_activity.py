#!/usr/bin/env python3
"""
student_activity.py
===================
Joins three sources into one per-student picture of weekly activity:

  * JupyterHub hours   – Grafana/Prometheus (hub_weekly_activity.py), keyed by hub
                         username = Temple email
  * Office visits      – the Student Meetings form/Sheet (names only), fuzzy-matched
                         to the roster by name, with the class time as a cross-check
  * Master roster      – master_roster.csv from the Instructor Dashboard:
                         Student, SIS Login ID, Section, Class Time, Instructor

Used two ways:
  1. weekly_instructor_reports.py imports it to add a "JupyterHub activity" block
     to each instructor's email (their sections only; matched on Class Time =
     course_section in instructor_mappings.csv).
  2. Run directly to rewrite a "Weekly activity" tab in a Google Sheet:
       python student_activity.py --sheet-id <ACTIVITY_SHEET_ID>
       python student_activity.py --csv activity.csv --meetings-csv meetings.csv   # local test

Environment
  GRAFANA_TOKEN                 Grafana service-account token            (hub hours)
  HUB_NAMESPACE                 your hub, from hub_weekly_activity.py --list-namespaces
  MASTER_ROSTER_CSV             roster CSV *contents* (GitHub secret) — or
  ROSTER_PATH                   path to master_roster.csv (default: next to this file)
  GOOGLE_SHEET_ID               Student Meetings Sheet                    (visits)
  GOOGLE_CREDENTIALS_JSON       service-account JSON (needs Editor on the activity Sheet)
  ACTIVITY_SHEET_ID             Sheet to write (or --sheet-id)
  SEMESTER_START                YYYY-MM-DD (default 2026-08-24)

No student names are written to the log.
"""

import argparse
import difflib
import io
import json
import logging
import os
import re
import sys
import unicodedata
from datetime import datetime, timedelta
from functools import lru_cache
from pathlib import Path
from zoneinfo import ZoneInfo

import pandas as pd

import hub_weekly_activity as hw

log = logging.getLogger("activity")
HERE = Path(__file__).resolve().parent
ET = ZoneInfo("America/New_York")
EMAIL_DOMAIN = "@temple.edu"
MATCH_HIGH, MATCH_ACCEPT = 90, 75
SHEET_TAB = "Weekly activity"

# Canvas section -> class-time label exactly as on the meetings form (and as the
# course_section values in instructor_mappings.csv). Used when master_roster.csv
# has no Class Time. Sections not listed here are learned each run from visits
# whose student name matches the roster exactly (see fill_class_times). Override
# without editing code via env SECTION_CLASS_TIMES='{"001": "T/R 9:30 AM", ...}'.
SECTION_CLASS_TIMES = {
    "004": "T/R 5:30 PM",
    "005": "M/W 5:30 PM",
    "007": "MWF 12:00 PM",
    "SCTC-1913-001": "Honors T/R 9:30",
}
TEST_STUDENTS = {"student, test", "test student", "test, student"}

# Shown in the email and used for the attached CSV's column names.
HUB_DEFINITION = ("7-day hub hours = the total time a student's JupyterHub server was running in the "
                  "last 7 days, sampled every 5 minutes. A server starts when the student opens the hub "
                  "and is shut down automatically after a period of inactivity, so this measures time "
                  "with the hub open — engagement, not effort or progress. 0 h means the student did "
                  "not open the hub at all this week.")
HUB_BINS   = [0, 1e-9, 2, 4, 6, 8, 10, 15, 20, float("inf")]
HUB_LABELS = ["0", "0–2", "2–4", "4–6", "6–8", "8–10", "10–15", "15–20", "20+"]
SECTION_COLORS = ["#2f6db5", "#1aa6a6", "#8e5ac8", "#2e9d5b", "#d1495b", "#6b7280"]


def hub_histogram(studs, sections):
    """PNG: students per 7-day-hub-hours bin, stacked by section; 0 h bar in orange."""
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    from io import BytesIO

    bins = pd.cut(studs["hub_hours"], HUB_BINS, labels=HUB_LABELS, right=True, include_lowest=True)
    fig, ax = plt.subplots(figsize=(9, 3.9))
    x = range(len(HUB_LABELS))
    bottom = pd.Series(0, index=HUB_LABELS)
    present = [sec for sec in sections if (studs["class_time"] == sec).any()]
    for i, sec in enumerate(present):
        c = bins[studs["class_time"] == sec].value_counts().reindex(HUB_LABELS, fill_value=0)
        colors = ["#e08a1e" if lab == "0" else SECTION_COLORS[i % len(SECTION_COLORS)] for lab in HUB_LABELS]
        ax.bar(x, c.values, bottom=bottom.values, color=colors, edgecolor="white", linewidth=0.6)
        bottom = bottom + c
    for xi, v in zip(x, bottom.values):
        if v:
            ax.text(xi, v, int(v), ha="center", va="bottom", fontsize=9)
    active = studs.loc[studs["hub_hours"] > 0, "hub_hours"]
    title = f"7-day hub hours — {len(studs)} students"
    if len(active):
        title += f" (median {active.median():.1f} h for the {len(active)} who used it)"
    ax.set_title(title, fontsize=11, fontweight="bold", loc="left", pad=12)
    ax.set_xticks(list(x))
    ax.set_xticklabels([f"{lab} h" for lab in HUB_LABELS], fontsize=9)
    ax.set_ylabel("Students")
    ax.set_xlabel("Hours with a JupyterHub server running, last 7 days (orange = not on the hub)")
    from matplotlib.ticker import MaxNLocator
    ax.yaxis.set_major_locator(MaxNLocator(integer=True))
    ax.set_ylim(0, max(1, int(bottom.max())) * 1.22)
    ax.grid(True, axis="y", alpha=0.3)
    ax.set_axisbelow(True)
    for sp in ("top", "right"):
        ax.spines[sp].set_visible(False)
    from matplotlib.patches import Patch
    handles = [Patch(color="#e08a1e", label="0 h – not on the hub")]
    if len(present) > 1:
        handles += [Patch(color=SECTION_COLORS[i % len(SECTION_COLORS)], label=sec)
                    for i, sec in enumerate(present)]
    ax.legend(handles=handles, fontsize=8, frameon=False, loc="upper left", bbox_to_anchor=(1.01, 1))
    buf = BytesIO()
    fig.savefig(buf, format="png", dpi=130, bbox_inches="tight", facecolor="white")
    plt.close(fig)
    return buf.getvalue()

MEET = {"date": "date", "time": "time", "first": "Student First Name",
        "last": "Student Last Name", "section": "Course Section"}
TEST_NAME = re.compile(r"\btest\b", re.I)


# ── name matching (same rules as the Instructor Dashboard) ────────────────────
def _norm_name(s):
    """Lowercase ASCII; accents/apostrophes removed; '-' and '/' kept as
    separators. 'A’Lanna' -> 'alanna', 'Zoë' -> 'zoe'."""
    s = unicodedata.normalize("NFKD", str(s)).encode("ascii", "ignore").decode()
    s = re.sub(r"['`]", "", s.lower())
    s = re.sub(r"[^a-z/\- ]", " ", s)
    return re.sub(r"\s+", " ", s).strip()


def _name_tokens(s):
    return [t for t in re.split(r"[\s/\-]+", _norm_name(s)) if t]


def _split_roster_name(name):
    """Canvas exports 'Last, First Middle' (sortable name); fall back to
    'First ... Last' if there's no comma."""
    name = str(name).strip()
    if "," in name:
        last, first = name.split(",", 1)
    else:
        parts = name.split()
        first, last = " ".join(parts[:-1]), (parts[-1] if parts else "")
    return first.strip(), last.strip()


@lru_cache(maxsize=None)
def _part_score(sheet_part, roster_part):
    """0-1 similarity of one name part (first OR last). Handles alternates
    typed with '/' ('Zhiling / Lucas'), hyphenated/compound names (matches
    any piece), middle names on the roster, short forms ('Alex'/'Alexis'),
    and typos (difflib ratio)."""
    r_full = _norm_name(roster_part).replace("-", " ")
    r_toks = set(_name_tokens(roster_part))
    if not r_full:
        return 0.0
    best = 0.0
    for alt in _norm_name(sheet_part).split("/"):
        a_full = alt.replace("-", " ").strip()
        a_toks = set(_name_tokens(alt))
        if not a_full:
            continue
        if a_full.replace(" ", "") == r_full.replace(" ", ""):
            return 1.0
        best = max(best, difflib.SequenceMatcher(
            None, a_full.replace(" ", ""), r_full.replace(" ", "")).ratio())
        if a_toks & r_toks:
            best = max(best, 0.92)                 # shared piece / middle name
        for at in a_toks:
            for rt in r_toks:
                if min(len(at), len(rt)) >= 3 and (rt.startswith(at) or at.startswith(rt)):
                    best = max(best, 0.85)         # short form
                best = max(best, 0.95 * difflib.SequenceMatcher(None, at, rt).ratio())
    return best


def _name_score(first, last, r_first, r_last):
    """0-100. Last name weighted a bit more (fewer nicknames); also tries
    first/last swapped, at a small discount."""
    straight = 0.45 * _part_score(first, r_first) + 0.55 * _part_score(last, r_last)
    swapped  = 0.45 * _part_score(first, r_last) + 0.55 * _part_score(last, r_first)
    return round(100 * max(straight, 0.95 * swapped), 1)



# ── roster ───────────────────────────────────────────────────────────────────
ROSTER_DIAG = ""   # why the last load_roster() came back empty (no student data)


def _read_table(text):
    """CSV or tab-separated text (pasting from Excel/Sheets gives tabs),
    tolerating a BOM, Windows line endings, blank lines and surrounding whitespace."""
    text = text.replace("\ufeff", "").replace("\r\n", "\n").replace("\r", "\n").strip()
    lines = [ln.strip() for ln in text.split("\n") if ln.strip()]
    if not lines:
        return pd.DataFrame()
    sep = "\t" if lines[0].count("\t") > lines[0].count(",") else ","
    return pd.read_csv(io.StringIO("\n".join(lines)), dtype=str, sep=sep).fillna("")


def load_roster():
    """master_roster.csv -> DataFrame(user, name, section, class_time, _first, _last).
    Reads MASTER_ROSTER_CSV (contents, e.g. a GitHub secret) if set, else ROSTER_PATH.
    Accepts comma- or tab-separated text; header names are matched case-insensitively."""
    global ROSTER_DIAG
    empty = pd.DataFrame(columns=["user", "name", "section", "class_time", "_first", "_last"])
    text = os.getenv("MASTER_ROSTER_CSV", "")
    if text.strip():
        source = "MASTER_ROSTER_CSV"
    else:
        path = Path(os.getenv("ROSTER_PATH") or HERE / "master_roster.csv")
        if not path.exists():
            ROSTER_DIAG = "MASTER_ROSTER_CSV is empty and no master_roster.csv file was found"
            return empty
        source, text = str(path), path.read_text(encoding="utf-8-sig")
    try:
        raw = _read_table(text)
    except Exception as e:
        ROSTER_DIAG = f"{source} could not be parsed: {e}"
        return empty
    by_lower = {c.strip().lower(): c for c in raw.columns}

    def col(*names):
        for n in names:
            if n in by_lower:
                return raw[by_lower[n]].astype(str).str.strip()
        return pd.Series("", index=raw.index)
    login = col("sis login id", "login id", "username")
    if (login == "").all():
        email = col("email", "e-mail")
        login = email.str.split("@").str[0]
    login = login.str.lower()
    df = pd.DataFrame({
        "user":       login + EMAIL_DOMAIN,
        "name":       col("student", "name", "student name"),
        "section":    col("section").str.replace("Section: ", "", regex=False),
        "class_time": col("class time"),
    })[login != ""]
    df = df[~df["name"].str.lower().isin(TEST_STUDENTS)].drop_duplicates("user").reset_index(drop=True)
    if df.empty:
        expected = ["Student", "SIS Login ID", "Section", "Class Time"]
        present = [e for e in expected if e.lower() in by_lower]
        ROSTER_DIAG = (f"{source} parsed as {len(raw)} data row(s) x {len(raw.columns)} column(s); "
                       f"expected header columns found: {present or 'none'} "
                       f"(the first line must be the header: Student,SIS Login ID,Section,Class Time,Instructor)")
        return empty
    ROSTER_DIAG = ""
    split = df["name"].map(_split_roster_name)
    df["_first"], df["_last"] = split.str[0], split.str[1]
    return df


# ── meetings ─────────────────────────────────────────────────────────────────
def _frame(values):
    """Grid -> DataFrame, finding the header row wherever it is (sorting the
    Sheet with the header included moves it into the middle)."""
    def is_header(row):
        low = [str(c).strip().lower() for c in row]
        return "date" in low and any("first" in c and "name" in c for c in low)
    hdr = next((i for i, r in enumerate(values) if is_header(r)), None)
    if hdr is None:
        raise RuntimeError("Could not find the header row in the meetings data")
    cols = [re.sub(r"\s+", " ", str(c)).strip() for c in values[hdr]]
    rows = [list(r) + [""] * (len(cols) - len(r)) for i, r in enumerate(values)
            if i != hdr and any(str(c).strip() for c in r)]
    return pd.DataFrame([r[:len(cols)] for r in rows], columns=cols)


def _gspread_client(write=False):
    import gspread
    creds_json = os.getenv("GOOGLE_CREDENTIALS_JSON")
    scopes = (["https://www.googleapis.com/auth/spreadsheets",
               "https://www.googleapis.com/auth/drive"] if write else
              ["https://www.googleapis.com/auth/spreadsheets.readonly",
               "https://www.googleapis.com/auth/drive.readonly"])
    if creds_json:
        from google.oauth2.service_account import Credentials
        return gspread.authorize(Credentials.from_service_account_info(json.loads(creds_json),
                                                                      scopes=scopes))
    return gspread.service_account(filename=os.getenv("GOOGLE_CREDS", hw.GOOGLE_CREDS), scopes=scopes)


def load_meetings(csv_path=None, semester_start=None):
    """Meetings this semester: first, last, label (class time), datetime (naive ET)."""
    if csv_path:
        import csv
        with open(csv_path, newline="", encoding="utf-8-sig") as f:
            values = list(csv.reader(f))
    else:
        sheet_id = os.getenv("GOOGLE_SHEET_ID")
        if not sheet_id:
            raise RuntimeError("GOOGLE_SHEET_ID is not set (or pass --meetings-csv)")
        values = _gspread_client().open_by_key(sheet_id).sheet1.get_all_values()
    raw = _frame(values)
    for k in ("date", "time", "first", "last", "section"):
        if MEET[k] not in raw.columns:
            raise RuntimeError(f"Meetings column '{MEET[k]}' missing")
    stamp = raw[MEET["date"]].str.strip() + " " + raw[MEET["time"]].str.strip().replace("", "00:00")
    dt = pd.to_datetime(stamp, format="%m/%d/%Y %H:%M", errors="coerce")
    miss = dt.isna()
    if miss.any():
        dt[miss] = pd.to_datetime(stamp[miss], format="mixed", errors="coerce")
    df = pd.DataFrame({"first": raw[MEET["first"]].str.strip(), "last": raw[MEET["last"]].str.strip(),
                       "label": raw[MEET["section"]].str.strip(), "datetime": dt})
    df = df.dropna(subset=["datetime"])
    df = df[~(df["first"] + " " + df["last"]).str.contains(TEST_NAME)]
    start = pd.Timestamp(semester_start or os.getenv("SEMESTER_START") or "2026-08-24")
    return df[df["datetime"] >= start].reset_index(drop=True)


def match_visits(meet, roster):
    """Add user / roster_name / match_score / match_quality to each visit.
    Quality: exact | high | review | none (none -> user is None).
    The class time the student picked must agree with the roster's Class Time
    for 'exact'/'high'; a mismatch drops it to 'review'."""
    meet = meet.copy()
    cols = ["user", "roster_name", "match_score", "match_quality"]
    if meet.empty or roster.empty:
        for c in cols:
            meet[c] = None
        meet["match_quality"] = "none"
        return meet
    results = {}
    for first, last, label in meet[["first", "last", "label"]].drop_duplicates().itertuples(index=False):
        raw = pd.Series([_name_score(first, last, f, l) for f, l in zip(roster["_first"], roster["_last"])],
                        index=roster.index)
        adj = raw + roster["class_time"].map(lambda ct: 0 if not ct else (3 if ct == label else -8))
        order = adj.sort_values(ascending=False)
        best, score = order.index[0], float(raw[order.index[0]])
        tie = len(order) > 1 and order.iloc[0] - order.iloc[1] < 3
        ct = roster.loc[best, "class_time"]
        sect_ok = (not ct) or ct == label
        if score < MATCH_ACCEPT:
            results[(first, last, label)] = (None, None, score, "none")
            continue
        q = ("exact" if score == 100 and not tie and sect_ok else
             "high" if score >= MATCH_HIGH and not tie and sect_ok else "review")
        results[(first, last, label)] = (roster.loc[best, "user"], roster.loc[best, "name"], score, q)
    keys = list(zip(meet["first"], meet["last"], meet["label"]))
    meet[cols] = pd.DataFrame([results[k] for k in keys], index=meet.index, columns=cols)
    return meet


def fill_class_times(roster, meet):
    """
    Give every roster student a class time (the form's section label), in order:
      1. SECTION_CLASS_TIMES (code or env)                         source 'set'
      2. Class Time already in master_roster.csv                   source 'roster'
      3. learned: the label most students in that Canvas section   source 'learned'
         picked on the form (exact/high name matches, >= 2 students, >= 60%)
    Returns (roster, {section: (label, source, n_students_seen)}).
    """
    roster = roster.copy()
    explicit = dict(SECTION_CLASS_TIMES)
    try:
        explicit.update(json.loads(os.getenv("SECTION_CLASS_TIMES") or "{}"))
    except ValueError:
        log.warning("SECTION_CLASS_TIMES env is not valid JSON — ignored")

    learned = {}
    if len(meet):
        blank = roster.assign(class_time="")          # name-only match, no section bias
        m = match_visits(meet, blank)
        good = m[m["match_quality"].isin(["exact", "high"])].drop_duplicates(["user", "label"])
        good = good.assign(section=good["user"].map(dict(zip(roster["user"], roster["section"]))))
        for sec, labels in good.groupby("section")["label"]:
            top = labels.value_counts()
            if top.iloc[0] >= 2 and top.iloc[0] / top.sum() >= 0.6:
                learned[sec] = (top.index[0], int(top.sum()))

    info = {}
    for sec in sorted(roster["section"].unique()):
        have = roster.loc[roster["section"] == sec, "class_time"]
        have = have[have != ""]
        if sec in explicit:
            info[sec] = (explicit[sec], "set", None)
        elif len(have):
            info[sec] = (have.mode().iloc[0], "roster", None)
        elif sec in learned:
            info[sec] = (learned[sec][0], "learned", learned[sec][1])
        else:
            info[sec] = ("", "unknown", None)
        roster.loc[roster["section"] == sec, "class_time"] = info[sec][0]
    return roster, info


def _aware(ts):
    ts = pd.Timestamp(ts).to_pydatetime()
    return ts if ts.tzinfo else ts.replace(tzinfo=ET)


# ── weekly email block ───────────────────────────────────────────────────────
def load_week_activity(meetings, week_start, now):
    """
    For weekly_instructor_reports.py. `meetings` is that script's DataFrame
    (columns from its COL map + 'datetime'). Returns a dict with roster,
    matched visits for the window, hub hours {user: hours} (or None if
    Grafana is unavailable), and a list of problems for the admin email.
    """
    problems = []
    roster = load_roster()
    if roster.empty:
        return {"roster": roster, "visits": pd.DataFrame(), "hub": None, "class_times": {},
                "problems": [f"Hub activity skipped: no master roster ({ROSTER_DIAG})"]}
    allm = pd.DataFrame({"first": meetings[MEET["first"]], "last": meetings[MEET["last"]],
                         "label": meetings[MEET["section"]], "datetime": meetings["datetime"]})
    allm = allm[~(allm["first"] + " " + allm["last"]).str.contains(TEST_NAME)]
    roster, class_times = fill_class_times(roster, allm)
    unknown = [s for s, (lab, src, _) in class_times.items() if src == "unknown"]
    if unknown:
        problems.append(f"No class time known for section(s) {unknown} — add them to SECTION_CLASS_TIMES "
                        f"in student_activity.py; their students are left out of instructor emails")

    m = allm[(allm["datetime"] >= week_start) & (allm["datetime"] <= now)]
    visits = match_visits(m, roster)
    n_none = int((visits["match_quality"] == "none").sum()) if len(visits) else 0
    if n_none:
        problems.append(f"{n_none} meeting(s) this week could not be matched to a roster student "
                        f"(nickname, typo, or student missing from master_roster.csv)")

    hub = None
    try:
        hub = hw.hours_between(_aware(week_start), _aware(now), os.getenv("HUB_NAMESPACE") or None)
    except hw.HubQueryError as e:
        problems.append(f"Hub hours unavailable this week: {e}")
    return {"roster": roster, "visits": visits, "hub": hub, "problems": problems,
            "class_times": class_times}


def instructor_block(rep, act, sections, csv_name="the attached CSV"):
    """(html, csv_bytes|None, png_bytes|None) for one instructor's sections (form
    class-time labels). `rep` is the Reporter (for its _cards/_table helpers).
    The html references the histogram as cid:hubchart."""
    navy, green, orange, blue = "#1f3a5f", "#2e9d5b", "#e08a1e", "#2f6db5"
    if not act or act["roster"].empty:
        return "", None, None
    studs = act["roster"][act["roster"]["class_time"].isin(sections)].copy()
    head = f"<h3 style='color:{navy};'>JupyterHub activity – last 7 days</h3>"
    if studs.empty:
        return (head + "<p style='font-size:13px;color:#666;'>No roster students are linked to these "
                "sections yet (the master roster's Class Time must match the section names above).</p>"), None, None

    v = act["visits"]
    v = v[v["match_quality"] != "none"] if len(v) else v
    vcount = v.groupby("user").size() if len(v) else pd.Series(dtype=int)
    hub = act["hub"]
    studs["visits"] = studs["user"].map(vcount).fillna(0).astype(int)
    studs["hub_hours"] = studs["user"].map(hub).fillna(0.0) if hub is not None else float("nan")

    n = len(studs)
    if hub is not None:
        active = studs[studs["hub_hours"] > 0]
        quiet = studs[(studs["hub_hours"] == 0) & (studs["visits"] == 0)]
        med = f"{active['hub_hours'].median():.1f}" if len(active) else "–"
        cards = rep._cards([(f"{len(active)}/{n}", "Students on the hub", green),
                            (med, "Median hours (active)", blue),
                            (int(studs["visits"].sum()), "Office visits (matched)", navy),
                            (len(quiet), "No hub time & no visit", orange if len(quiet) else green)])
    else:
        quiet = studs[studs["visits"] == 0].iloc[0:0]
        cards = ("<p style='font-size:13px;background:#fff7e6;padding:8px 10px;'>Hub hours could not be "
                 "retrieved this week; office visits are shown below.</p>"
                 + rep._cards([(int(studs["visits"].sum()), "Office visits (matched)", navy),
                               (int((studs["visits"] > 0).sum()), "Students with a visit", green)]))

    rows = []
    for sec in sections:
        s = studs[studs["class_time"] == sec]
        if s.empty:
            continue
        if hub is not None:
            a = s[s["hub_hours"] > 0]
            rows.append([sec, len(s), len(a), f"{a['hub_hours'].mean():.1f}" if len(a) else "–",
                         int(s["visits"].sum())])
        else:
            rows.append([sec, len(s), "–", "–", int(s["visits"].sum())])
    table = rep._table(["Section", "Students", "On the hub", "Avg hours (active)", "Office visits"],
                       rows, highlight_cols=(2,))

    quiet_html = ""
    if len(quiet):
        names = sorted(quiet["name"])
        shown = ", ".join(names[:30]) + (f" … and {len(names) - 30} more (see attached CSV)" if len(names) > 30 else "")
        quiet_html = (f"<p style='font-size:13px;margin:8px 0 2px;'><b>No hub time and no office visit this "
                      f"week ({len(names)}):</b></p><p style='font-size:13px;margin:0 0 8px;'>{shown}</p>")

    png = hub_histogram(studs, sections) if hub is not None else None
    chart = ("<img src='cid:hubchart' alt='Histogram of 7-day hub hours' "
             "style='width:100%;max-width:676px;margin:6px 0;'>") if png else ""
    note = (f"<p style='font-size:12px;color:#555;background:#f6f8fb;border-left:4px solid {blue};"
            f"padding:8px 10px;margin:8px 0;'><b>What “hub hours” means.</b> {HUB_DEFINITION} "
            f"Office visits are matched to students by name and class time. Every student's 7-day hub "
            f"hours and office visits are in <b>{csv_name}</b>, sorted with the least active first.</p>")
    out = studs.sort_values(["hub_hours", "visits", "name"] if hub is not None else ["visits", "name"])
    csv = out[["name", "user", "section", "class_time", "hub_hours", "visits"]].rename(columns={
        "name": "Student", "user": "Email", "section": "Canvas section", "class_time": "Class time",
        "hub_hours": "Hub hours (7 days)", "visits": "Office visits (7 days)"}).to_csv(index=False).encode()
    return head + cards + chart + table + quiet_html + note, csv, png


# ── weekly Sheet ─────────────────────────────────────────────────────────────
def semester_mondays(start, now):
    first = (pd.Timestamp(start) - pd.Timedelta(days=pd.Timestamp(start).weekday())).normalize()
    this = (pd.Timestamp(now) - pd.Timedelta(days=pd.Timestamp(now).weekday())).normalize()
    return list(pd.date_range(first, this, freq="7D"))


def build_weekly_table(roster, visits, hub_long, mondays):
    """One row per roster student; totals, then newest-first week pairs."""
    df = roster[["name", "user", "section", "class_time"]].copy()
    acc = visits[visits["match_quality"] != "none"].copy() if len(visits) else visits
    if len(acc):
        acc["week"] = (acc["datetime"] - pd.to_timedelta(acc["datetime"].dt.weekday, unit="D")).dt.normalize()
    hub_long = hub_long.copy()
    if len(hub_long):
        hub_long["week"] = pd.to_datetime(hub_long["week_of"])
    hrs_cols, vis_cols = [], []
    for wk in sorted(mondays, reverse=True):
        label = wk.strftime("%b %d").replace(" 0", " ")
        h = hub_long[hub_long["week"] == wk].set_index("user")["hours"] if len(hub_long) else pd.Series(dtype=float)
        vc = acc[acc["week"] == wk].groupby("user").size() if len(acc) else pd.Series(dtype=int)
        df[f"{label} hrs"] = df["user"].map(h).fillna(0.0).round(1)
        df[f"{label} visits"] = df["user"].map(vc).fillna(0).astype(int)
        hrs_cols.append(f"{label} hrs")
        vis_cols.append(f"{label} visits")
    df.insert(4, "Hub hours (total)", df[hrs_cols].sum(axis=1).round(1) if hrs_cols else 0.0)
    df.insert(5, "Visits (total)", df[vis_cols].sum(axis=1) if vis_cols else 0)
    df.insert(6, "Weeks on hub", (df[hrs_cols] > 0).sum(axis=1) if hrs_cols else 0)
    df = df.rename(columns={"name": "Student", "user": "Email", "section": "Section",
                            "class_time": "Class time"})
    return df.sort_values(["Section", "Student"]).reset_index(drop=True)


def write_tabs(sheet_id, tabs):
    """tabs: {title: (note, DataFrame)} — each tab is cleared and rewritten."""
    import gspread
    sh = _gspread_client(write=True).open_by_key(sheet_id)
    for title, (note, df) in tabs.items():
        try:
            ws = sh.worksheet(title)
            ws.clear()
        except gspread.WorksheetNotFound:
            ws = sh.add_worksheet(title=title, rows=max(len(df) + 10, 20), cols=max(len(df.columns) + 2, 8))
        body = df.astype(object).where(df.notna(), "").values.tolist()
        ws.update([[note]] + [df.columns.tolist()] + body)
        try:
            ws.freeze(rows=2, cols=1)
        except Exception:
            pass
        log.info("Wrote tab '%s' (%d rows)", title, len(df))


def main():
    logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")
    ap = argparse.ArgumentParser(description="Weekly hub hours + office visits per student")
    ap.add_argument("--sheet-id", default=os.getenv("ACTIVITY_SHEET_ID"), help="Sheet to write")
    ap.add_argument("--csv", help="also save the weekly table to this CSV")
    ap.add_argument("--meetings-csv", help="read meetings from a CSV export instead of the Sheet")
    ap.add_argument("--no-hub", action="store_true", help="skip Grafana (visits only)")
    a = ap.parse_args()

    roster = load_roster()
    if roster.empty:
        sys.exit(f"No master roster: {ROSTER_DIAG}")
    meet = load_meetings(a.meetings_csv)
    roster, class_times = fill_class_times(roster, meet)
    for sec, (lab, src, n) in class_times.items():
        log.info("Section %s -> class time %r (%s%s)", sec, lab, src, f", {n} students" if n else "")
    visits = match_visits(meet, roster)
    q = visits["match_quality"].value_counts().to_dict()
    log.info("Roster %d students; %d meetings this semester; matches %s", len(roster), len(meet), q)

    now = datetime.now(ET)
    mondays = semester_mondays(os.getenv("SEMESTER_START") or "2026-08-24", now.replace(tzinfo=None))
    hub_long = pd.DataFrame(columns=["user", "week_of", "hours"])
    if not a.no_hub:
        try:
            hub_long = hw.weekly_hours(len(mondays), os.getenv("HUB_NAMESPACE") or None)
        except hw.HubQueryError as e:
            log.error("Hub hours unavailable: %s — writing visits only", e)
    table = build_weekly_table(roster, visits, hub_long, mondays)

    unmatched = visits[visits["match_quality"].isin(["none", "review"])]
    check = pd.DataFrame({"Typed name": unmatched["first"] + " " + unmatched["last"],
                          "Class time": unmatched["label"],
                          "When": unmatched["datetime"].dt.strftime("%Y-%m-%d %H:%M"),
                          "Quality": unmatched["match_quality"],
                          "Best guess": unmatched["roster_name"].fillna(""),
                          "Score": unmatched["match_score"].round(0)})
    if a.csv:
        table.to_csv(a.csv, index=False)
        log.info("Saved %d rows -> %s", len(table), a.csv)
    if a.sheet_id:
        stamp = now.strftime("%Y-%m-%d %H:%M ET")
        write_tabs(a.sheet_id, {
            SHEET_TAB: (f"Per student, per week (Mon–Sun): hours with a running JupyterHub server and "
                        f"matched office visits. Updated {stamp}", table),
            "Visits to check": (f"Meetings not matched to a roster student ('none') or matched with a "
                                f"caveat ('review'). Updated {stamp}", check),
        })
    if not (a.csv or a.sheet_id):
        log.info("Nothing written — pass --sheet-id or --csv")


if __name__ == "__main__":
    main()
