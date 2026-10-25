"""
Shared loader for the Student Meetings table.

Sources (choose with MEETINGS_SOURCE = email | xlsx | sheet; default: xlsx if a URL is set, else sheet):
  email  - Power Automate emails each submission to the Gmail account in EMAIL_USER (subject MEETING-ROW);
           history/backfill files are emails with a CSV attached (subject MEETING-HISTORY).
           Read over IMAP with EMAIL_USER / EMAIL_PASSWORD (the same Gmail app password used to send).
  xlsx   - MEETINGS_XLSX_URL  - an OneDrive/SharePoint "Anyone with the link" share of the Excel workbook
                          that Power Automate appends to (no Google API, no Graph API).
  2. GOOGLE_SHEET_ID + GOOGLE_CREDENTIALS_JSON - the Google Sheet (legacy / fallback).

Returns a DataFrame of strings with columns like:
  date, time, Student First Name, Student Last Name, Course Section, Meeting person,
  Meeting Type, Topic, rating
`date` is normalised to MM/DD/YYYY and `time` to HH:MM even if Excel turned them into
real dates/times or serial numbers.

The URL is a secret (anyone holding it can read the file): it is never logged.
"""
import email as emaillib
import html as htmllib
import imaplib
import io
import json
import logging
import os
import re
import urllib.request
from datetime import datetime, timedelta

import pandas as pd

log = logging.getLogger("meetings_source")

_ONEDRIVE = re.compile(r"(sharepoint\.com|1drv\.ms|onedrive\.live\.com)", re.I)


def _download(url):
    if _ONEDRIVE.search(url) and "download=1" not in url:
        url += ("&" if "?" in url else "?") + "download=1"
    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0 (meetings-report)"})
    try:
        with urllib.request.urlopen(req, timeout=90) as r:
            data = r.read()
    except Exception as e:  # do not chain: the traceback must not expose the URL
        raise RuntimeError(f"Could not download the workbook ({type(e).__name__}: {e}). "
                           "Check the share link is 'Anyone with the link can view'.") from None
    if not data.startswith(b"PK"):
        raise RuntimeError("The link did not return an .xlsx file (probably a sign-in page). "
                           "The share must be 'Anyone with the link' and view-only is fine.")
    return data


def _norm_date(v):
    v = str(v).strip()
    m = re.match(r"^(\d{4})-(\d{2})-(\d{2})(?:[ T].*)?$", v)
    if m:
        return f"{m[2]}/{m[3]}/{m[1]}"
    if re.match(r"^\d+(\.\d+)?$", v) and 30000 < float(v) < 70000:   # Excel serial date
        return (datetime(1899, 12, 30) + timedelta(days=int(float(v)))).strftime("%m/%d/%Y")
    m = re.match(r"^(\d{1,2})/(\d{1,2})/(\d{4})$", v)                  # 8/5/2026 -> 08/05/2026
    if m:
        return f"{int(m[1]):02d}/{int(m[2]):02d}/{m[3]}"
    return v


def _norm_time(v):
    v = str(v).strip()
    m = re.match(r"^(?:\d{4}-\d{2}-\d{2}[ T])?(\d{1,2}):(\d{2})(?::\d{2}(?:\.\d+)?)?$", v)
    if m:
        return f"{int(m[1]):02d}:{m[2]}"
    if re.match(r"^\d*\.\d+$", v) and float(v) < 1:                    # Excel fraction of a day
        mins = int(round(float(v) * 1440))
        return f"{mins // 60 % 24:02d}:{mins % 60:02d}"
    return v


def _finish(df):
    df = df.copy()
    df.columns = [re.sub(r"\s+", " ", str(c).replace("\xa0", " ")).strip() for c in df.columns]
    df = df.fillna("").astype(str)
    if "date" in df.columns:
        df["date"] = df["date"].map(_norm_date)
    if "time" in df.columns:
        df["time"] = df["time"].map(_norm_time)
    return df


# ---------------------------------------------------------------------------- email source
_STD = ["date", "time", "Student First Name", "Student Last Name", "Course Section",
        "Meeting person", "Meeting Type", "Topic", "rating"]
_ROW_KEYS = {"date": "date", "time": "time", "first": "Student First Name", "last": "Student Last Name",
             "section": "Course Section", "person": "Meeting person", "type": "Meeting Type",
             "topic": "Topic", "rating": "rating"}


def _clean(x):
    return re.sub(r"\s+", " ", str(x).replace("\xa0", " ")).strip()


def _body_text(msg):
    """Text of an email (HTML tags stripped); prefers the part that contains the @@ markers."""
    texts = []
    for part in msg.walk():
        if part.get_content_type() in ("text/plain", "text/html") and not part.get_filename():
            payload = part.get_payload(decode=True) or b""
            txt = payload.decode(part.get_content_charset() or "utf-8", errors="replace")
            if part.get_content_type() == "text/html":
                txt = re.sub(r"(?is)<(script|style).*?</\1>", " ", txt)
                txt = re.sub(r"(?i)<br\s*/?>|</p>|</div>|</tr>", "\n", txt)
                txt = re.sub(r"<[^>]+>", " ", txt)
            texts.append(htmllib.unescape(txt))
    for t in texts:
        if "@@" in t:
            return t
    return "\n".join(texts)


def parse_row_email(msg):
    """One submission -> dict, from '@@key@@ value' markers. Returns None if not a row email."""
    text = _body_text(msg)
    found = dict((k.lower(), _clean(v)) for k, v in
                 re.findall(r"@@(\w+)@@(.*?)(?=@@\w+@@|\Z)", text, flags=re.S))
    if "date" not in found or "first" not in found:
        return None
    row = {col: found.get(k, "") for k, col in _ROW_KEYS.items()}
    row["_id"] = found.get("id", "")
    return row


def _forms_export_to_std(df):
    """Microsoft Forms 'Download results' export -> standard columns."""
    cols = list(df.columns)
    done = pd.to_datetime(df["Completion time"], format="mixed", errors="coerce")
    pick = lambda prefix: next(c for c in cols if c.lower().startswith(prefix.lower()))
    out = pd.DataFrame({
        "date": done.dt.strftime("%m/%d/%Y"), "time": done.dt.strftime("%H:%M"),
        "Student First Name": df[pick("Student First")].map(_clean),
        "Student Last Name": df[pick("Last name")].map(_clean),
        "Course Section": df[pick("Student course")].map(_clean).replace({"T/R 9:30": "T/R 9:30 AM"}),
        "Meeting person": df[pick("Who are you")].map(_clean),
        "Meeting Type": df[pick("Session Type")].map(_clean),
        "Topic": df[pick("Topic")].map(_clean),
        "rating": df[pick("Visit helped")].map(_clean)})
    return out[done.notna()]


def parse_history_email(msg):
    """CSV/XLSX attachments -> DataFrame in standard columns (resync CSV or raw Forms export)."""
    frames = []
    for part in msg.walk():
        name = (part.get_filename() or "").lower()
        if not name.endswith((".csv", ".xlsx")):
            continue
        data = part.get_payload(decode=True) or b""
        if name.endswith(".xlsx"):
            df = pd.read_excel(io.BytesIO(data), dtype=str, engine="openpyxl")
        else:
            try:
                txt = data.decode("utf-8-sig")
            except UnicodeDecodeError:
                txt = data.decode("cp1252")
            df = pd.read_csv(io.StringIO(txt), dtype=str)
        df = _finish(df)
        if "Completion time" in df.columns:
            df = _forms_export_to_std(df)
        keep = [c for c in _STD if c in df.columns]
        if "date" in keep:
            frames.append(df[keep])
    return pd.concat(frames, ignore_index=True) if frames else pd.DataFrame(columns=_STD)


def _from_email(user, password):
    if not (user and password):
        raise RuntimeError("EMAIL_USER / EMAIL_PASSWORD are needed to read meetings from email")
    M = imaplib.IMAP4_SSL("imap.gmail.com")
    try:
        M.login(user, password)
        if M.select('"[Gmail]/All Mail"', readonly=True)[0] != "OK":
            M.select("INBOX", readonly=True)
        rows, hist = [], []
        for tag, bucket in (("MEETING-HISTORY", hist), ("MEETING-ROW", rows)):
            typ, data = M.search(None, "SUBJECT", f'"{tag}"')
            for num in (data[0].split() if typ == "OK" and data and data[0] else []):
                typ, msgdata = M.fetch(num, "(RFC822)")
                if typ != "OK":
                    continue
                msg = emaillib.message_from_bytes(msgdata[0][1])
                if tag == "MEETING-ROW":
                    r = parse_row_email(msg)
                    if r:
                        bucket.append(r)
                else:
                    bucket.append(parse_history_email(msg))
    finally:
        try:
            M.logout()
        except Exception:
            pass
    return _combine(rows, hist)


def _combine(row_dicts, history_frames):
    frames = [h for h in history_frames if len(h)]
    if row_dicts:
        frames.append(_finish(pd.DataFrame(row_dicts)[_STD]))
    if not frames:
        return pd.DataFrame(columns=_STD)
    df = pd.concat(frames, ignore_index=True)
    df = _finish(df)
    key = (df["date"] + "|" + df["time"] + "|" + df["Student First Name"].str.lower().str.strip() + "|" +
           df["Student Last Name"].str.lower().str.strip() + "|" + df["Meeting person"].str.lower().str.strip())
    df = df[~key.duplicated(keep="last")]
    log.info("Loaded %d meetings from email (%d row emails, %d history files)",
             len(df), len(row_dicts), len(history_frames))
    return df.reset_index(drop=True)


def _exclude(df):
    """EXCLUDE_STUDENTS='first|last,first|last' - drop test rows."""
    pairs = {x.strip().lower() for x in (os.getenv("EXCLUDE_STUDENTS") or "").split(",") if x.strip()}
    if not pairs or df.empty or "Student First Name" not in df.columns:
        return df
    key = df["Student First Name"].str.strip().str.lower() + "|" + df["Student Last Name"].str.strip().str.lower()
    dropped = int(key.isin(pairs).sum())
    if dropped:
        log.info("Excluded %d test rows", dropped)
    return df[~key.isin(pairs)].reset_index(drop=True)


def fetch_raw(sheet_id=None, creds_json=None, xlsx_url=None):
    source = (os.getenv("MEETINGS_SOURCE") or "").strip().lower()
    if source == "email":
        df = _from_email(os.getenv("EMAIL_USER"), os.getenv("EMAIL_PASSWORD"))
    else:
        df = _fetch_file_or_sheet(sheet_id, creds_json, xlsx_url if source != "sheet" else None)
    return _exclude(df)


def _fetch_file_or_sheet(sheet_id, creds_json, xlsx_url):
    if xlsx_url:
        data = _download(xlsx_url.strip())
        df = pd.read_excel(io.BytesIO(data), sheet_name=0, dtype=str, engine="openpyxl")
        log.info("Loaded %d rows from the Excel workbook", len(df))
        return _finish(df)

    import gspread
    from google.oauth2.service_account import Credentials
    if not (sheet_id and creds_json):
        raise RuntimeError("Set MEETINGS_XLSX_URL, or GOOGLE_SHEET_ID + GOOGLE_CREDENTIALS_JSON")
    creds = Credentials.from_service_account_info(
        json.loads(creds_json),
        scopes=["https://www.googleapis.com/auth/spreadsheets.readonly",
                "https://www.googleapis.com/auth/drive.readonly"])
    vals = gspread.authorize(creds).open_by_key(sheet_id).sheet1.get_all_values()
    if len(vals) < 2:
        return pd.DataFrame()
    log.info("Loaded %d rows from the Google Sheet", len(vals) - 1)
    return _finish(pd.DataFrame(vals[1:], columns=vals[0]))
