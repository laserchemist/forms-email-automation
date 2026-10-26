#!/usr/bin/env python3
"""
Weekly meeting reports for EDS (Fall 2026).

Reads the "Student Meetings" Google Sheet (filled by Microsoft Form -> Power Automate)
and sends three kinds of email:

  1. Instructors   - their own sections: semester + this week, who students met, charts,
                      and (with student_activity.py) JupyterHub hours + office visits per student
  2. Course Assistants (CAs) - class-wide activity, *their own* meetings (week + cumulative),
                      a nudge to record every meeting, and the form QR code
  3. Admin summary - one overview + data-quality warnings (unmapped sections, unknown names)

Configuration (environment variables)
  EMAIL_USER, EMAIL_PASSWORD     Gmail sender + app password          (required to send)
  GOOGLE_SHEET_ID                new Fall sheet id                    (required unless --csv)
  GOOGLE_CREDENTIALS_JSON        service-account JSON                 (required unless --csv)
  ADMIN_EMAIL                    gets the summary (default: EMAIL_USER)
  SEMESTER_START                 YYYY-MM-DD, earlier rows are ignored (default 2026-08-24)
  MEETING_FORM_URL               link shown with the QR code
  SEND_INSTRUCTORS / SEND_CAS    "false" to skip that audience        (default true)
  DRY_RUN                        "true" -> write HTML previews to ./preview, send nothing
  TEST_RECIPIENT                 send every email to this address only (subject shows the real recipient)
  HUB_ACTIVITY                   "false" to leave out the JupyterHub activity block (default true)
  GRAFANA_TOKEN, HUB_NAMESPACE,  used by student_activity.py for hub hours / the roster
  MASTER_ROSTER_CSV

CLI:  python weekly_instructor_reports.py [--dry-run] [--csv local_file.csv]

Config files next to this script:
  instructor_mappings.csv   course_section,instructor_name,email,course_name,active[,notes]
  ca_mappings.csv           form_name,ca_name,email,active[,notes]
  assets/meeting_form_qr.png
  student_activity.py, hub_weekly_activity.py   (optional: JupyterHub activity block)
"""

import argparse
import base64
import json
import logging
import os
import re
import smtplib
import sys
from datetime import datetime, timedelta
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from html import escape
from io import BytesIO
from pathlib import Path
from zoneinfo import ZoneInfo

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt  # noqa: E402
import pandas as pd  # noqa: E402

try:                                   # JupyterHub hours + matched visits (optional)
    import student_activity as sa      # noqa: E402
    SA_IMPORT_ERROR = None
except Exception as _e:                # module or one of its packages missing
    sa = None
    SA_IMPORT_ERROR = str(_e)

logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")
log = logging.getLogger("weekly")

HERE = Path(__file__).resolve().parent
ET = ZoneInfo("America/New_York")          # sheet times are Eastern; GitHub runners are UTC
DEFAULT_FORM_URL = "https://forms.cloud.microsoft/r/5zNSiU70si?origin=QRCode"
QR_PATH = HERE / "assets" / "meeting_form_qr.png"

COL = {
    "date": "date", "time": "time",
    "first": "Student First Name", "last": "Student Last Name",
    "section": "Course Section", "person": "Meeting person",
    "type": "Meeting Type", "topic": "Topic", "rating": "rating",
}

NAVY, BLUE, TEAL, ORANGE, GREEN, GRAY = "#1f3a5f", "#2f6db5", "#1aa6a6", "#e08a1e", "#2e9d5b", "#6b7280"


def truthy(v, default=True):
    if v is None or (isinstance(v, float) and pd.isna(v)) or str(v).strip() == "":
        return default
    return str(v).strip().lower() not in ("false", "0", "no", "n")


def env_flag(name, default=True):
    return truthy(os.getenv(name), default)


def now_et():
    return datetime.now(ET).replace(tzinfo=None)


def slug(s):
    return re.sub(r"[^A-Za-z0-9]+", "_", s).strip("_")


def norm(s):
    return re.sub(r"\s+", " ", str(s).replace("\xa0", " ")).strip()


# --------------------------------------------------------------------------- config / data
class Reporter:
    def __init__(self, dry_run=False, csv_path=None):
        self.email_user = os.getenv("EMAIL_USER")
        self.email_password = os.getenv("EMAIL_PASSWORD")
        self.sheet_id = os.getenv("GOOGLE_SHEET_ID")
        self.creds_json = os.getenv("GOOGLE_CREDENTIALS_JSON")
        self.admin_email = os.getenv("ADMIN_EMAIL") or self.email_user
        # TEST_RECIPIENT: every email goes here instead, subject tagged with the real recipient
        self.test_to = (os.getenv("TEST_RECIPIENT") or "").strip()
        self.form_url = os.getenv("MEETING_FORM_URL") or DEFAULT_FORM_URL
        self.semester_start = pd.Timestamp(os.getenv("SEMESTER_START") or "2026-08-24")
        self.dry_run = dry_run or env_flag("DRY_RUN", False)
        self.csv_path = csv_path
        self.now = now_et()
        self.week_start = self.now - timedelta(days=7)
        self.instructors = self.load_instructors()
        self.cas = self.load_cas()
        self.activity = None        # student_activity.load_week_activity() result
        self.hub_csv = None         # per-student activity CSV for the current instructor

    # ---- mappings
    def _read_cfg(self, name):
        p = HERE / name
        if not p.exists():
            log.error("Missing config file %s", name)
            return pd.DataFrame()
        df = pd.read_csv(p, dtype=str).fillna("")
        df.columns = [c.strip() for c in df.columns]
        return df

    def load_instructors(self):
        """section -> dict(instructor, email, course_name)"""
        df = self._read_cfg("instructor_mappings.csv")
        out = {}
        for _, r in df.iterrows():
            if not truthy(r.get("active"), True):
                continue
            sec = norm(r.get("course_section", ""))
            if sec:
                out[sec] = {"instructor": norm(r.get("instructor_name", "")),
                            "email": norm(r.get("email", "")),
                            "course_name": norm(r.get("course_name", "")) or "Elements of Data Science"}
        log.info("Loaded %d active section mappings", len(out))
        return out

    def load_cas(self):
        """form_name(lower) -> dict(name, email)"""
        df = self._read_cfg("ca_mappings.csv")
        out = {}
        for _, r in df.iterrows():
            if not truthy(r.get("active"), True):
                continue
            # form_name may list several spellings: "Soul|Soul Moreno"
            aliases = [norm(a) for a in str(r.get("form_name", "")).split("|") if norm(a)]
            if aliases:
                key = aliases[0].lower()
                out[key] = {"form_name": aliases[0],
                            "aliases": {a.lower() for a in aliases},
                            "name": norm(r.get("ca_name", "")) or aliases[0],
                            "email": norm(r.get("email", ""))}
        log.info("Loaded %d active CAs", len(out))
        return out

    # ---- data
    @staticmethod
    def _header_row(vals):
        """Index of the header row. Usually 0, but sorting the Sheet with the
        header included (e.g. by Meeting person) moves it into the middle."""
        for i, row in enumerate(vals):
            low = [norm(c).lower() for c in row]
            if "date" in low and any("first" in c and "name" in c for c in low):
                return i
        return 0

    def load_data(self):
        if self.csv_path:
            raw = pd.read_csv(self.csv_path, dtype=str).fillna("")
        else:
            import gspread
            from google.oauth2.service_account import Credentials
            if not (self.creds_json and self.sheet_id):
                raise RuntimeError("GOOGLE_SHEET_ID / GOOGLE_CREDENTIALS_JSON not set")
            creds = Credentials.from_service_account_info(
                json.loads(self.creds_json),
                scopes=["https://www.googleapis.com/auth/spreadsheets.readonly",
                        "https://www.googleapis.com/auth/drive.readonly"])
            ws = gspread.authorize(creds).open_by_key(self.sheet_id).sheet1
            vals = ws.get_all_values()
            if len(vals) < 2:
                return pd.DataFrame()
            h = self._header_row(vals)
            if h:
                log.warning("Header found on sheet row %d, not row 1 (sheet sorted with header?) — handled", h + 1)
            width = len(vals[h])
            rows = [list(r) + [""] * (width - len(r)) for i, r in enumerate(vals) if i != h]
            raw = pd.DataFrame([r[:width] for r in rows], columns=[norm(c) for c in vals[h]])
        raw.columns = [norm(c) for c in raw.columns]
        log.info("Loaded %d raw rows", len(raw))

        for k in ("date", "time", "first", "last", "section", "person"):
            if COL[k] not in raw.columns:
                raise RuntimeError(f"Column '{COL[k]}' missing. Found: {list(raw.columns)}")
        for k in ("type", "topic", "rating"):
            if COL[k] not in raw.columns:
                raw[COL[k]] = ""

        df = raw.copy()
        for c in df.columns:
            df[c] = df[c].map(norm)
        df = df[df[COL["date"]] != ""].copy()
        stamp = df[COL["date"]] + " " + df[COL["time"]].replace("", "00:00")
        dt = pd.to_datetime(stamp, format="%m/%d/%Y %H:%M", errors="coerce")
        miss = dt.isna()
        if miss.any():
            dt[miss] = pd.to_datetime(stamp[miss], format="mixed", errors="coerce")
        df["datetime"] = dt
        bad = int(df["datetime"].isna().sum())
        self.bad_dates = bad
        if bad:
            log.warning("%d rows have unparseable dates and were dropped", bad)
        df = df.dropna(subset=["datetime"])
        before = len(df)
        df = df[df["datetime"] >= self.semester_start]
        log.info("%d rows on/after %s (%d older rows ignored)",
                 len(df), self.semester_start.date(), before - len(df))

        df["section"] = df[COL["section"]]
        df["person"] = df[COL["person"]]
        df["student"] = (df[COL["first"]].str.lower() + "|" + df[COL["last"]].str.lower())
        df["rating_n"] = pd.to_numeric(df[COL["rating"]], errors="coerce")
        df["week"] = df["datetime"].dt.normalize() - pd.to_timedelta(df["datetime"].dt.weekday, unit="D")
        return df.sort_values("datetime").reset_index(drop=True)

    def in_week(self, df):
        return df[(df["datetime"] >= self.week_start) & (df["datetime"] <= self.now)]

    def all_weeks(self):
        start = self.semester_start - timedelta(days=self.semester_start.weekday())
        cur = pd.Timestamp(self.now.date()) - timedelta(days=self.now.weekday())
        return pd.date_range(start, cur, freq="7D")

    # ---- stats
    @staticmethod
    def summarize(all_df, wk_df):
        s = {
            "total": len(all_df), "students": all_df["student"].nunique(),
            "week_total": len(wk_df), "week_students": wk_df["student"].nunique(),
            "avg_rating": all_df["rating_n"].mean() if len(all_df) else float("nan"),
            "week_rating": wk_df["rating_n"].mean() if len(wk_df) else float("nan"),
            "first": all_df["datetime"].min() if len(all_df) else None,
        }
        s["by_section"] = all_df.groupby("section").size().sort_values(ascending=False)
        s["week_by_section"] = wk_df.groupby("section").size()
        s["by_person"] = all_df.groupby("person").size().sort_values(ascending=False)
        s["week_by_person"] = wk_df.groupby("person").size()
        s["week_by_type"] = wk_df.groupby(COL["type"]).size().sort_values(ascending=False)
        return s

    # ----------------------------------------------------------------------- charts
    @staticmethod
    def _style(ax, title, ylabel=None, xlabel=None):
        ax.set_title(title, fontsize=11, fontweight="bold", loc="left")
        if ylabel:
            ax.set_ylabel(ylabel)
        if xlabel:
            ax.set_xlabel(xlabel)
        ax.grid(True, axis="y", alpha=0.3)
        ax.set_axisbelow(True)
        ax.tick_params(axis="both", labelsize=9)
        for sp in ("top", "right"):
            ax.spines[sp].set_visible(False)

    @staticmethod
    def _int_axis(ax, axis="y"):
        from matplotlib.ticker import MaxNLocator
        (ax.yaxis if axis == "y" else ax.xaxis).set_major_locator(MaxNLocator(integer=True))

    def _png(self, fig):
        buf = BytesIO()
        fig.savefig(buf, format="png", dpi=130, bbox_inches="tight", facecolor="white")
        plt.close(fig)
        return buf.getvalue()

    def _weekly_counts(self, df, weeks):
        return df.groupby("week").size().reindex(weeks, fill_value=0)

    def _week_labels(self, weeks):
        return [w.strftime("%b %d").replace(" 0", " ") for w in weeks]

    def _cum_line(self, ax, series_list, title):
        """series_list: [(df, label, color)] cumulative meetings by date"""
        days = pd.date_range(self.semester_start.normalize(), pd.Timestamp(self.now.date()), freq="D")
        for df, label, color in series_list:
            daily = df.groupby(df["datetime"].dt.normalize()).size().reindex(days, fill_value=0).cumsum()
            ax.plot(days, daily.values, color=color, lw=2.2, label=f"{label} ({int(daily.iloc[-1])})")
        self._style(ax, title, "Meetings (cumulative)")
        self._int_axis(ax)
        ax.legend(fontsize=8, frameon=False, loc="upper left")
        import matplotlib.dates as mdates
        ax.xaxis.set_major_locator(mdates.WeekdayLocator(byweekday=0, interval=2))
        ax.xaxis.set_major_formatter(mdates.DateFormatter("%b %d"))
        ax.tick_params(axis="x", rotation=45)

    def _daily_week(self, ax, wk_df, title):
        days = pd.date_range((self.now - timedelta(days=6)).date(), self.now.date(), freq="D")
        c = wk_df.groupby(wk_df["datetime"].dt.normalize()).size().reindex(days, fill_value=0)
        bars = ax.bar(range(len(days)), c.values, color=BLUE)
        ax.set_xticks(range(len(days)))
        ax.set_xticklabels([d.strftime("%a\n%m/%d") for d in days])
        for b, v in zip(bars, c.values):
            if v:
                ax.text(b.get_x() + b.get_width() / 2, v, int(v), ha="center", va="bottom", fontsize=8)
        self._style(ax, title, "Meetings")
        self._int_axis(ax)

    def instructor_chart(self, name, all_df, wk_df, sections):
        weeks = self.all_weeks()
        fig, ax = plt.subplots(2, 2, figsize=(13, 9))
        fig.suptitle(f"{name} – meetings in your sections", fontsize=14, fontweight="bold")

        self._cum_line(ax[0, 0], [(all_df, "Your sections", NAVY)], "Cumulative meetings")

        # weekly stacked by section
        bottom = pd.Series(0, index=weeks)
        palette = [BLUE, TEAL, ORANGE, GREEN, "#8e5ac8", "#d1495b", GRAY]
        for i, sec in enumerate(sections):
            c = self._weekly_counts(all_df[all_df["section"] == sec], weeks)
            ax[0, 1].bar(range(len(weeks)), c.values, bottom=bottom.values,
                         color=palette[i % len(palette)], label=sec)
            bottom = bottom + c
        ax[0, 1].set_xticks(range(len(weeks)))
        ax[0, 1].set_xticklabels(self._week_labels(weeks), rotation=45)
        for x, v in enumerate(bottom.values):
            if v:
                ax[0, 1].text(x, v, int(v), ha="center", va="bottom", fontsize=8)
        self._style(ax[0, 1], "Meetings per week (week starting Monday)", "Meetings")
        ax[0, 1].margins(y=0.15)
        self._int_axis(ax[0, 1])
        ax[0, 1].legend(fontsize=8, frameon=False)

        # who students met
        bp = all_df.groupby("person").size().sort_values()
        wp = wk_df.groupby("person").size().reindex(bp.index, fill_value=0)
        y = range(len(bp))
        ax[1, 0].barh([i + 0.2 for i in y], bp.values, height=0.4, color=NAVY, label="Semester")
        ax[1, 0].barh([i - 0.2 for i in y], wp.values, height=0.4, color=TEAL, label="This week")
        ax[1, 0].set_yticks(list(y))
        ax[1, 0].set_yticklabels(bp.index, fontsize=9)
        self._style(ax[1, 0], "Who your students met with", xlabel="Meetings")
        ax[1, 0].grid(True, axis="x", alpha=0.3)
        ax[1, 0].grid(False, axis="y")
        self._int_axis(ax[1, 0], "x")
        ax[1, 0].legend(fontsize=8, frameon=False, loc="lower right")

        self._daily_week(ax[1, 1], wk_df, "Last 7 days")
        fig.tight_layout(rect=[0, 0, 1, 0.96])
        return self._png(fig)

    def ca_chart(self, ca, all_df, mine_df, mine_week):
        weeks = self.all_weeks()
        x = range(len(weeks))
        fig, ax = plt.subplots(2, 2, figsize=(13, 9))
        fig.suptitle(f"{ca['name']} – class activity and your meetings", fontsize=14, fontweight="bold")

        allw, minew = self._weekly_counts(all_df, weeks), self._weekly_counts(mine_df, weeks)
        ax[0, 0].bar(x, allw.values, color="#c7d6ea", label=f"All meetings ({int(allw.sum())})")
        ax[0, 0].bar(x, minew.values, color=ORANGE, width=0.5, label=f"Yours ({int(minew.sum())})")
        ax[0, 0].set_xticks(list(x))
        ax[0, 0].set_xticklabels(self._week_labels(weeks), rotation=45)
        self._style(ax[0, 0], "Meetings per week", "Meetings")
        self._int_axis(ax[0, 0])
        ax[0, 0].legend(fontsize=8, frameon=False)

        self._cum_line(ax[0, 1], [(all_df, "Whole class", NAVY), (mine_df, "You", ORANGE)],
                       "Cumulative meetings")

        # staff comparison: cumulative by person (top 8) with me highlighted
        bp = all_df.groupby("person").size().sort_values().tail(8)
        wp = self_wp = all_df[all_df["datetime"] >= self.week_start].groupby("person").size().reindex(bp.index, fill_value=0)
        y = list(range(len(bp)))
        colors = [ORANGE if p.lower() in ca["aliases"] else NAVY for p in bp.index]
        ax[1, 0].barh([i + 0.2 for i in y], bp.values, height=0.4, color=colors)
        ax[1, 0].barh([i - 0.2 for i in y], wp.values, height=0.4, color=TEAL)
        ax[1, 0].set_yticks(y)
        ax[1, 0].set_yticklabels(bp.index, fontsize=9)
        self._style(ax[1, 0], "Meetings by staff member (dark = semester, teal = this week)", xlabel="Meetings")
        ax[1, 0].grid(True, axis="x", alpha=0.3)
        ax[1, 0].grid(False, axis="y")
        self._int_axis(ax[1, 0], "x")

        sec = all_df.groupby("section").size().sort_values()
        mine_sec = mine_df.groupby("section").size().reindex(sec.index, fill_value=0)
        y = list(range(len(sec)))
        ax[1, 1].barh([i + 0.2 for i in y], sec.values, height=0.4, color="#c7d6ea", label="All")
        ax[1, 1].barh([i - 0.2 for i in y], mine_sec.values, height=0.4, color=ORANGE, label="You")
        ax[1, 1].set_yticks(y)
        ax[1, 1].set_yticklabels(sec.index, fontsize=9)
        self._style(ax[1, 1], "Meetings by section (semester)", xlabel="Meetings")
        ax[1, 1].grid(True, axis="x", alpha=0.3)
        ax[1, 1].grid(False, axis="y")
        self._int_axis(ax[1, 1], "x")
        ax[1, 1].legend(fontsize=8, frameon=False, loc="lower right")

        fig.tight_layout(rect=[0, 0, 1, 0.96])
        return self._png(fig)

    def wordcloud(self, df, title):
        try:
            from wordcloud import WordCloud
        except Exception:
            return None
        t = df[COL["topic"]].astype(str)
        t = t[(t != "") & (~t.str.lower().isin(["not specified", "nan", "none", "n/a"]))]
        text = re.sub(r"[^\w\s]", " ", " ".join(t)).strip()
        if len(text) < 3:
            return None
        wc = WordCloud(width=900, height=380, background_color="white", max_words=50,
                       colormap="viridis", min_font_size=12).generate(text)
        fig = plt.figure(figsize=(9, 3.8))
        plt.imshow(wc, interpolation="bilinear")
        plt.axis("off")
        plt.title(title, fontsize=12, fontweight="bold")
        return self._png(fig)

    # ----------------------------------------------------------------------- HTML helpers
    @staticmethod
    def _cards(items):
        """items: [(value, label, color)] -> table-based cards (Gmail-safe)"""
        w = int(100 / len(items))
        tds = "".join(
            f"<td width='{w}%' align='center' style='padding:14px 6px;background:#f6f8fb;"
            f"border:1px solid #e3e8ef;'><div style='font-size:28px;font-weight:bold;color:{c};'>{v}</div>"
            f"<div style='font-size:12px;color:#555;margin-top:4px;'>{escape(l)}</div></td>"
            for v, l, c in items)
        return f"<table width='100%' cellspacing='6' cellpadding='0' style='border-collapse:separate;'><tr>{tds}</tr></table>"

    @staticmethod
    def _table(headers, rows, highlight_cols=()):
        th = "".join(f"<th align='{'left' if i == 0 else 'center'}' style='padding:8px;border:1px solid #ddd;"
                     f"background:#eef2f7;font-size:12px;'>{escape(h)}</th>" for i, h in enumerate(headers))
        body = ""
        for r in rows:
            tds = ""
            for i, v in enumerate(r):
                bg = "background:#e8f5ec;" if i in highlight_cols else ""
                al = "left" if i == 0 else "center"
                tds += f"<td align='{al}' style='padding:7px;border:1px solid #ddd;font-size:13px;{bg}'>{escape(str(v))}</td>"
            body += f"<tr>{tds}</tr>"
        return f"<table width='100%' cellspacing='0' cellpadding='0' style='border-collapse:collapse;margin:10px 0;'><tr>{th}</tr>{body}</table>"

    @staticmethod
    def _wrap(title, subtitle, inner):
        return f"""<html><body style="margin:0;padding:0;background:#eef1f5;font-family:'Segoe UI',Arial,sans-serif;color:#333;">
<table width="100%" cellspacing="0" cellpadding="0"><tr><td align="center">
<table width="720" cellspacing="0" cellpadding="0" style="max-width:720px;background:#ffffff;">
<tr><td style="background:{NAVY};color:#ffffff;padding:26px 22px;text-align:center;">
<div style="font-size:22px;font-weight:bold;color:#ffffff;">{title}</div>
<div style="font-size:13px;color:#cfe0f5;margin-top:6px;">{subtitle}</div></td></tr>
<tr><td style="padding:24px 22px;">{inner}</td></tr>
<tr><td align="center" style="background:#f6f8fb;color:#777;font-size:11px;padding:14px;">
Automated report &middot; generated {datetime.now(ET).strftime('%b %d, %Y %I:%M %p')} ET</td></tr>
</table></td></tr></table></body></html>"""

    def _period(self):
        return f"{self.week_start.strftime('%b %d')} – {self.now.strftime('%b %d, %Y')}"

    @staticmethod
    def _fmt(x, nd=1):
        return "–" if x is None or pd.isna(x) else f"{x:.{nd}f}"

    # ----------------------------------------------------------------------- instructor email
    def build_instructor(self, name, email, course, sections, all_df):
        wk = self.in_week(all_df)
        s = self.summarize(all_df, wk)
        rows = []
        for sec in sections:
            sd = all_df[all_df["section"] == sec]
            sw = wk[wk["section"] == sec]
            rows.append([sec, len(sd), sd["student"].nunique(), len(sw), sw["student"].nunique()])
        rows.append(["Total", s["total"], s["students"], s["week_total"], s["week_students"]])
        by_person = "".join(
            f"<li>{escape(p)}: {n} total, {int(s['week_by_person'].get(p, 0))} this week</li>"
            for p, n in s["by_person"].items())
        types = ", ".join(f"{escape(str(k))} {v}" for k, v in s["week_by_type"].items()) or "none"

        # JupyterHub hours + matched office visits for these sections' roster students
        hub_html, self.hub_csv = "", None
        if sa is not None and self.activity is not None:
            try:
                hub_html, self.hub_csv = sa.instructor_block(self, self.activity, sections)
            except Exception as e:
                log.error("Hub activity block failed for %s: %s", name, e)

        inner = f"""
<p style="margin-top:0;">Hello {escape(name)}, here is the meeting activity for students in your section(s) of
<b>{escape(course)}</b> since the semester began, with the past seven days highlighted.</p>
{self._cards([(s['week_total'], 'Meetings this week', GREEN), (s['week_students'], 'Students this week', GREEN),
              (s['total'], 'Meetings this semester', NAVY), (s['students'], 'Students this semester', NAVY)])}
<h3 style="color:{NAVY};">Your sections</h3>
{self._table(['Section', 'Semester meetings', 'Semester students', 'This week', 'Students this week'], rows, highlight_cols=(3, 4))}
{hub_html}
<h3 style="color:{NAVY};">Who your students met with</h3>
<ul style="margin:4px 0 12px 18px;padding:0;font-size:13px;">{by_person}</ul>
<p style="font-size:13px;">This week by type: {types}. Average rating this week: <b>{self._fmt(s['week_rating'])}</b>
(semester {self._fmt(s['avg_rating'])}, on a 1–5 scale).</p>
<img src="cid:chart" alt="Charts" style="width:100%;max-width:676px;">
{"<img src='cid:wordcloud' alt='Topics' style='width:100%;max-width:676px;'>" if self.wordcloud_png else ""}
<p style="font-size:12px;color:#666;">Attached: this week's meeting list (CSV){" and each student's hub hours and office visits (CSV)" if self.hub_csv else ""}.
Student names are for instructional use only – please do not forward.</p>"""
        return (self._wrap("Weekly Meeting Report", f"{escape(course)} &middot; {self._period()}", inner), s, wk)

    # ----------------------------------------------------------------------- CA email
    def build_ca(self, ca, all_df, s_all):
        mine = all_df[all_df["person"].str.lower().isin(ca["aliases"])]
        mine_wk = self.in_week(mine)
        sm = self.summarize(mine, mine_wk)
        team_rank = s_all["by_person"].index.tolist()
        wk_all = s_all["week_total"]

        if len(mine_wk):
            show = mine_wk.sort_values("datetime", ascending=False).head(15)
            mrows = [[r["datetime"].strftime("%a %m/%d %I:%M %p").replace(" 0", " "),
                      f"{r[COL['first']]} {r[COL['last']][:1]}.", r["section"],
                      r[COL["type"]] or "–", (r[COL["topic"]] or "–")[:50]] for _, r in show.iterrows()]
            my_list = self._table(["When", "Student", "Section", "Type", "Topic"], mrows)
        else:
            my_list = ("<p style='background:#fff7e6;border-left:4px solid #e08a1e;padding:10px 12px;font-size:13px;'>"
                       "No meetings are recorded under your name in the last 7 days. If you met with students, "
                       "please have them (or you) submit the form – and it only counts if your name is selected.</p>")

        week_rows = [[p, int(s_all['week_by_person'].get(p, 0)), n] for p, n in s_all["by_person"].items()]
        week_tbl = self._table(["Staff member", "This week", "Semester"], week_rows, highlight_cols=(1,))

        reminders = f"""
<table width="100%" cellspacing="0" cellpadding="0" style="margin:18px 0;background:#fff7e6;border:1px solid #f1d9a8;">
<tr><td style="padding:16px;" valign="top">
<div style="font-size:16px;font-weight:bold;color:#8a5a00;">Please record every meeting</div>
<ul style="margin:8px 0 8px 18px;padding:0;font-size:13px;line-height:1.6;">
<li><b>Every</b> student meeting – in-person, virtual, or drop-in – should be logged with the form, even short ones.</li>
<li>Choose <b>your name</b> under “Who are you meeting with” so the meeting is credited to you.</li>
<li><b>Be available at the times you put in the Calendar</b> – or edit the Calendar in real time so it always reflects when you are actually available.</li>
<li>Have this QR code ready (print it or save it on your phone) so every meeting can be logged on the spot.</li>
<li>Ask each student to scan it before they leave – it takes about 30 seconds.</li></ul>
<a href="{escape(self.form_url)}" style="display:inline-block;background:{ORANGE};color:#ffffff;text-decoration:none;
 padding:9px 18px;border-radius:4px;font-size:13px;font-weight:bold;">Open the meeting form</a>
</td>
<td width="190" align="center" style="padding:14px;" valign="middle">
<img src="cid:qr" alt="Meeting form QR code" width="170" style="display:block;">
</td></tr></table>"""

        rank_txt = ""
        mine_names = [p for p in team_rank if p.lower() in ca["aliases"]]
        if mine_names:
            rank_txt = f" You are #{team_rank.index(mine_names[0]) + 1} of {len(team_rank)} people on the log."

        inner = f"""
<p style="margin-top:0;">Hi {escape(ca['name'])} – thank you for the time you put in with students. Here is how the
whole class looked this past week and how your own meetings are adding up.</p>
{reminders}
<h3 style="color:{NAVY};margin-bottom:4px;">Your meetings</h3>
{self._cards([(sm['week_total'], 'You – this week', ORANGE), (sm['week_students'], 'Students you met this week', ORANGE),
              (sm['total'], 'You – semester total', NAVY), (sm['students'], 'Different students, semester', NAVY)])}
<p style="font-size:12px;color:#666;margin:2px 0 8px;">Average rating on your meetings: this week {self._fmt(sm['week_rating'])},
semester {self._fmt(sm['avg_rating'])}.{rank_txt}</p>
{my_list}
<h3 style="color:{NAVY};margin-bottom:4px;">Overall class activity</h3>
{self._cards([(wk_all, 'All meetings this week', GREEN), (s_all['week_students'], 'Students this week', GREEN),
              (s_all['total'], 'All meetings, semester', NAVY), (s_all['students'], 'Students, semester', NAVY)])}
{week_tbl}
<img src="cid:chart" alt="Charts" style="width:100%;max-width:676px;">
<p style="font-size:12px;color:#666;">Attached: your meetings this week (CSV). Student information is confidential course
material – please do not share it outside the teaching team.</p>"""
        return (self._wrap("Weekly CA Report", f"Elements of Data Science &middot; {self._period()}", inner),
                mine, mine_wk, sm)

    # ----------------------------------------------------------------------- admin email
    def build_admin(self, all_df, s_all, sent, problems):
        sec_rows = []
        for sec in sorted(set(all_df["section"]) | set(self.instructors)):
            sd = all_df[all_df["section"] == sec]
            inst = self.instructors.get(sec)
            sec_rows.append([sec, inst["instructor"] if inst else "NOT MAPPED", len(sd),
                             int((sd["datetime"] >= self.week_start).sum()), sd["student"].nunique()])
        ca_rows = []
        for key, ca in self.cas.items():
            m = all_df[all_df["person"].str.lower().isin(ca["aliases"])]
            ca_rows.append([ca["name"], ca["email"] or "NO EMAIL", int((m["datetime"] >= self.week_start).sum()), len(m)])
        prob = "".join(f"<li>{escape(p)}</li>" for p in problems) or "<li>None 🎉</li>"
        sent_html = "".join(f"<li>{escape(x)}</li>" for x in sent) or "<li>nothing sent</li>"

        hub_html = ""
        act = self.activity
        if act is not None and not act["roster"].empty:
            roster = act["roster"]
            hub = act["hub"]
            on_hub = sum(1 for u in roster["user"] if hub and hub.get(u, 0) > 0) if hub is not None else "–"
            v = act["visits"]
            matched = int((v["match_quality"] != "none").sum()) if len(v) else 0
            hub_html = (f"<h3 style='color:{NAVY};'>JupyterHub activity (last 7 days)</h3>"
                        + self._cards([(len(roster), "Roster students", NAVY),
                                       (on_hub, "On the hub this week", GREEN),
                                       (f"{matched}/{len(v)}", "Visits matched to roster", BLUE)]))

        inner = f"""
{self._cards([(s_all['week_total'], 'Meetings this week', GREEN), (s_all['total'], 'Semester meetings', NAVY),
              (s_all['students'], 'Unique students', NAVY), (len(problems), 'Things to fix', ORANGE if problems else GREEN)])}
<h3 style="color:{NAVY};">Needs attention</h3><ul style="font-size:13px;">{prob}</ul>
{hub_html}
<h3 style="color:{NAVY};">Sections</h3>
{self._table(['Section', 'Instructor', 'Semester', 'This week', 'Students'], sec_rows)}
<h3 style="color:{NAVY};">CAs</h3>
{self._table(['CA', 'Email on file', 'This week', 'Semester'], ca_rows)}
<h3 style="color:{NAVY};">Sent this run</h3><ul style="font-size:13px;">{sent_html}</ul>"""
        return self._wrap("Weekly Report – Admin Summary", self._period(), inner)

    # ----------------------------------------------------------------------- sending
    def send(self, to, subject, html, images=None, attachments=None, tag="msg"):
        """images: {cid: png_bytes}; attachments: [(filename, bytes)]"""
        images, attachments = images or {}, attachments or []
        if self.test_to:
            subject = f"[TEST – would go to {to or '(no email on file)'}] {subject}"
            to = self.test_to
        if self.dry_run:
            out = HERE / "preview"
            out.mkdir(exist_ok=True)
            page = html
            for cid, data in images.items():
                page = page.replace(f"cid:{cid}", "data:image/png;base64," + base64.b64encode(data).decode())
            (out / f"{tag}.html").write_text(f"<!-- To: {to} | Subject: {subject} -->\n{page}", encoding="utf-8")
            log.info("DRY RUN – wrote preview/%s.html (would send to %s)", tag, to)
            return True
        if not (self.email_user and self.email_password):
            log.error("EMAIL_USER / EMAIL_PASSWORD not set")
            return False
        root = MIMEMultipart("mixed")
        root["From"], root["To"], root["Subject"] = self.email_user, to, subject
        rel = MIMEMultipart("related")
        alt = MIMEMultipart("alternative")
        alt.attach(MIMEText(re.sub(r"<[^>]+>", " ", html)[:1500] + "\n(Open in an HTML-capable mail client.)", "plain"))
        alt.attach(MIMEText(html, "html"))
        rel.attach(alt)
        for cid, data in images.items():
            img = MIMEImage(data, _subtype="png")
            img.add_header("Content-ID", f"<{cid}>")
            img.add_header("Content-Disposition", "inline", filename=f"{cid}.png")
            rel.attach(img)
        root.attach(rel)
        for fname, data in attachments:
            part = MIMEBase("application", "octet-stream")
            part.set_payload(data)
            encoders.encode_base64(part)
            part.add_header("Content-Disposition", "attachment", filename=fname)
            root.attach(part)
        try:
            with smtplib.SMTP("smtp.gmail.com", 587) as srv:
                srv.starttls()
                srv.login(self.email_user, self.email_password)
                srv.send_message(root)
            return True
        except Exception as e:
            log.error("Send failed for %s: %s", tag, e)
            return False

    @staticmethod
    def _csv_bytes(df):
        cols = [c for c in (COL["date"], COL["time"], COL["first"], COL["last"], COL["section"],
                            COL["person"], COL["type"], COL["topic"], COL["rating"]) if c in df.columns]
        return df[cols].to_csv(index=False).encode("utf-8")

    # ----------------------------------------------------------------------- orchestration
    def load_activity(self, df, problems):
        """JupyterHub hours + roster-matched visits for the last 7 days (optional)."""
        if not env_flag("HUB_ACTIVITY", True):
            return
        if sa is None:
            problems.append(f"Hub activity skipped: student_activity.py could not be imported ({SA_IMPORT_ERROR})")
            return
        try:
            self.activity = sa.load_week_activity(df, self.week_start, self.now)
            problems.extend(self.activity["problems"])
        except Exception as e:
            self.activity = None
            problems.append(f"Hub activity skipped: {e}")

    def run(self):
        df = self.load_data()
        if df is None or df.empty:
            log.error("No data for this semester – nothing to send")
            return 1
        problems, sent = [], []
        if getattr(self, "bad_dates", 0):
            problems.append(f"{self.bad_dates} sheet rows had unreadable dates and were skipped")

        s_all = self.summarize(df, self.in_week(df))
        stamp = self.now.strftime("%Y%m%d")
        self.load_activity(df, problems)

        # unmapped sections / unknown people
        for sec in sorted(set(df["section"]) - set(self.instructors)):
            if sec:
                problems.append(f"Section “{sec}” has no instructor in instructor_mappings.csv "
                                f"({int((df['section'] == sec).sum())} meetings)")
        known = set().union(*[c["aliases"] for c in self.cas.values()]) if self.cas else set()
        for p in sorted(set(df["person"])):
            if p and p.lower() not in known and not re.match(r"^(dr|prof)\.?\s", p, re.I):
                problems.append(f"Meeting person “{p}” is not in ca_mappings.csv (add them or fix the form)")

        # ---- instructors (grouped by email so one person with several sections gets one message)
        if env_flag("SEND_INSTRUCTORS", True):
            groups = {}
            for sec, info in self.instructors.items():
                if not info["email"] and not self.test_to:
                    problems.append(f"No email for instructor of section “{sec}”")
                    continue
                g = groups.setdefault(info["email"], {"name": info["instructor"], "course": info["course_name"], "sections": []})
                g["sections"].append(sec)
            for email, g in groups.items():
                sdf = df[df["section"].isin(g["sections"])]
                wk = self.in_week(sdf)
                self.wordcloud_png = self.wordcloud(wk, "Topics this week")
                chart = self.instructor_chart(g["name"], sdf, wk, g["sections"]) if len(sdf) else None
                html, s, wk = self.build_instructor(g["name"], email, g["course"], g["sections"], sdf)
                imgs = {"chart": chart} if chart else {}
                if self.wordcloud_png:
                    imgs["wordcloud"] = self.wordcloud_png
                attach = [(f"meetings_week_{stamp}.csv", self._csv_bytes(wk))] if len(wk) else []
                if self.hub_csv:
                    attach.append((f"student_activity_week_{stamp}.csv", self.hub_csv))
                subj = f"Weekly meeting report – {s['week_total']} this week, {s['total']} this semester"
                ok = self.send(email, subj, html, imgs, attach, tag=f"instructor_{slug(g['name'])}")
                sent.append(f"{g['name']} ({'ok' if ok else 'FAILED'})")
                if not ok:
                    problems.append(f"Email to instructor {g['name']} failed")

        # ---- CAs
        if env_flag("SEND_CAS", True):
            for key, ca in self.cas.items():
                if not ca["email"] and not self.test_to:
                    problems.append(f"CA “{ca['name']}” has no email in ca_mappings.csv – skipped")
                    continue
                html, mine, mine_wk, sm = self.build_ca(ca, df, s_all)
                chart = self.ca_chart(ca, df, mine, mine_wk)
                imgs = {"chart": chart}
                if QR_PATH.exists():
                    imgs["qr"] = QR_PATH.read_bytes()
                else:
                    problems.append("assets/meeting_form_qr.png is missing – CA emails sent without QR")
                subj = f"CA weekly report – you: {sm['week_total']} this week, {sm['total']} total · remember to log every meeting"
                ok = self.send(ca["email"], subj, html, imgs,
                               [(f"my_meetings_week_{stamp}.csv", self._csv_bytes(mine_wk))] if len(mine_wk) else [],
                               tag=f"ca_{slug(ca['name'])}")
                sent.append(f"CA {ca['name']} ({'ok' if ok else 'FAILED'})")
                if not ok:
                    problems.append(f"Email to CA {ca['name']} failed")

        # ---- admin (last, so it can report everything above)
        if self.admin_email:
            html = self.build_admin(df, s_all, sent, problems)
            self.send(self.admin_email, f"Weekly report summary – {len(sent)} sent, {len(problems)} issue(s)",
                      html, tag="admin_summary")
        log.info("Done. %d emails, %d issues", len(sent), len(problems))
        for p in problems:
            log.warning("Issue: %s", p)           # no student names are logged
        return 0


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--dry-run", action="store_true", help="write HTML previews to ./preview, send nothing")
    ap.add_argument("--csv", help="read meetings from a local CSV instead of Google Sheets")
    a = ap.parse_args()
    rep = Reporter(dry_run=a.dry_run, csv_path=a.csv)
    rep.wordcloud_png = None
    sys.exit(rep.run())


if __name__ == "__main__":
    main()
