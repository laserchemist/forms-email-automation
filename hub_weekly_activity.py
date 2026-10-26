"""
hub_weekly_activity.py
======================
Weekly JupyterHub activity per student -> Google Sheet (or CSV).

For each student, counts the hours their JupyterHub server was running in
each week (Mon-Sun, US Eastern), using the same Prometheus data as Grafana's
"User Sessions" panel. Note: running != actively working — the hub shuts
idle servers down after a while, so this is an upper bound on time on task.

Setup (once):
  pip install requests pandas gspread
  export GRAFANA_TOKEN=...            (read -s GRAFANA_TOKEN && export GRAFANA_TOKEN)
  Share the target Google Sheet (Editor) with the service account's
  client_email from the credentials JSON.

Usage:
  python hub_weekly_activity.py --weeks 8 --csv weekly.csv        # test, no Sheet
  python hub_weekly_activity.py --weeks 8 --sheet-id <SHEET_ID>   # write Sheet
  python hub_weekly_activity.py --list-namespaces                 # find your hub
"""

import argparse, os, sys
from datetime import datetime, timedelta
from zoneinfo import ZoneInfo

import pandas as pd
import requests

GRAFANA_URL = os.environ.get("GRAFANA_URL", "https://grafana.temple.2i2c.cloud")
PROM_UID    = os.environ.get("PROM_UID", "P1809F7CD0C75ACF3")
GOOGLE_CREDS = os.environ.get("GOOGLE_CREDS",
                              "/home/jovyan/shared/nordic-knowledge-6598a7fdb7c8.json")
TZ   = ZoneInfo("America/New_York")
STEP = 5          # minutes between samples when counting running time
TAB  = "Weekly hub hours"

# Running jupyter-* pods, labelled with the hub username (from Grafana's
# jupyterhub/grafana-dashboards "User Sessions" panel).
SESSIONS = (
    '(kube_pod_status_phase{{phase="Running", pod=~"jupyter-.*"{ns}}} == 1)'
    ' * on (namespace, pod) group_left(annotation_hub_jupyter_org_username)'
    ' group by (namespace, pod, annotation_hub_jupyter_org_username)'
    ' (kube_pod_annotations{{pod=~"jupyter-.*"{ns}}})'
)


class HubQueryError(RuntimeError):
    """Grafana/Prometheus unreachable, token missing/expired, or bad query."""


def prom_query(query, at=None):
    """Instant PromQL query through Grafana's datasource proxy."""
    token = os.environ.get("GRAFANA_TOKEN")
    if not token:
        raise HubQueryError("GRAFANA_TOKEN is not set — run: read -s GRAFANA_TOKEN && export GRAFANA_TOKEN")
    params = {"query": query}
    if at is not None:
        params["time"] = at.timestamp()
    try:
        r = requests.get(f"{GRAFANA_URL}/api/datasources/proxy/uid/{PROM_UID}/api/v1/query",
                         headers={"Authorization": f"Bearer {token}"}, params=params, timeout=120)
    except requests.RequestException as e:
        raise HubQueryError(f"Could not reach {GRAFANA_URL}: {e}")
    if r.status_code != 200:
        hint = " (token expired? make a new one in Grafana)" if r.status_code == 401 else ""
        raise HubQueryError(f"Prometheus query failed: HTTP {r.status_code}{hint}: {r.text[:200]}")
    body = r.json()
    if body.get("status") != "success":
        raise HubQueryError(f"Prometheus error: {body.get('error', body)}")
    return body["data"]["result"]


def list_namespaces():
    res = prom_query('count by (namespace) (kube_pod_status_phase{pod=~"jupyter-.*"})')
    for row in sorted(res, key=lambda x: x["metric"].get("namespace", "")):
        print(f'{row["metric"].get("namespace")}: {row["value"][1]} user pod(s) now')


def week_starts(n_weeks, now):
    """Mondays 00:00 Eastern, oldest first, ending with the current week."""
    this_monday = (now - timedelta(days=now.weekday())).replace(hour=0, minute=0,
                                                               second=0, microsecond=0)
    return [this_monday - timedelta(weeks=k) for k in range(n_weeks - 1, -1, -1)]


def weekly_hours(n_weeks, namespace=None):
    """Long table: user, week_of, hours."""
    ns = f', namespace="{namespace}"' if namespace else ""
    base = SESSIONS.format(ns=ns)
    now = datetime.now(TZ)
    rows = []
    for start in week_starts(n_weeks, now):
        end = min(start + timedelta(weeks=1), now)
        secs = int((end - start).total_seconds())
        if secs < STEP * 60:
            continue
        q = (f"sum by (annotation_hub_jupyter_org_username) "
             f"(sum_over_time(({base})[{secs}s:{STEP}m]))")
        for r in prom_query(q, at=end):
            user = r["metric"].get("annotation_hub_jupyter_org_username", "?")
            rows.append({"user": user, "week_of": start.date().isoformat(),
                         "hours": round(float(r["value"][1]) * STEP / 60, 1)})
        print(f"  week of {start:%b %d}: {sum(1 for x in rows if x['week_of'] == start.date().isoformat())} user(s)")
    return pd.DataFrame(rows, columns=["user", "week_of", "hours"])


def hours_between(start, end, namespace=None):
    """{username: hours with a running server} between two aware datetimes."""
    ns = f', namespace="{namespace}"' if namespace else ""
    secs = int((end - start).total_seconds())
    q = (f"sum by (annotation_hub_jupyter_org_username) "
         f"(sum_over_time(({SESSIONS.format(ns=ns)})[{secs}s:{STEP}m]))")
    return {r["metric"].get("annotation_hub_jupyter_org_username", "?"):
            round(float(r["value"][1]) * STEP / 60, 1) for r in prom_query(q, at=end)}


def to_wide(df_long):
    """One row per user, one column per week, plus totals."""
    if df_long.empty:
        return df_long
    wide = (df_long.pivot_table(index="user", columns="week_of", values="hours",
                                aggfunc="sum", fill_value=0)
                   .sort_index(axis=1))
    wide.columns = [f"wk {c}" for c in wide.columns]
    wide.insert(0, "weeks_active", (wide > 0).sum(axis=1))
    wide.insert(0, "total_hours", wide.drop(columns="weeks_active").sum(axis=1).round(1))
    return wide.sort_values("total_hours", ascending=False).reset_index()


def write_sheet(wide, sheet_id, tab=TAB):
    import gspread
    gc = gspread.service_account(filename=GOOGLE_CREDS)
    sh = gc.open_by_key(sheet_id)
    try:
        ws = sh.worksheet(tab)
        ws.clear()
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(title=tab, rows=len(wide) + 10, cols=len(wide.columns) + 2)
    stamp = datetime.now(TZ).strftime("%Y-%m-%d %H:%M ET")
    ws.update([[f"Hours with a running JupyterHub server, per week (Mon–Sun ET). Updated {stamp}"]]
              + [wide.columns.tolist()] + wide.astype(object).values.tolist())
    print(f"✅ Wrote {len(wide)} user(s) to sheet tab '{tab}'")


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--weeks", type=int, default=8, help="weeks back, including this one")
    ap.add_argument("--namespace", help="limit to one hub (see --list-namespaces)")
    ap.add_argument("--sheet-id", help="Google Sheet ID to write to")
    ap.add_argument("--csv", help="also/instead save to this CSV file")
    ap.add_argument("--list-namespaces", action="store_true")
    a = ap.parse_args()

    try:
        if a.list_namespaces:
            return list_namespaces()
        print(f"Querying {a.weeks} week(s) from {GRAFANA_URL} ...")
        wide = to_wide(weekly_hours(a.weeks, a.namespace))
    except HubQueryError as e:
        sys.exit(f"❌ {e}")
    if wide.empty:
        sys.exit("No running user servers found in that period — try --list-namespaces.")
    print(wide.head(10).to_string(index=False))
    if a.csv:
        wide.to_csv(a.csv, index=False)
        print(f"✅ Saved {len(wide)} user(s) -> {a.csv}")
    if a.sheet_id:
        write_sheet(wide, a.sheet_id)


if __name__ == "__main__":
    main()
