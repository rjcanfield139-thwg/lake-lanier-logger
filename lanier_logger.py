#!/usr/bin/env python3
"""
Lake Lanier inflow / outflow logger
-----------------------------------
Reads https://lakelanierwater.com/ every hour, on the hour, and appends a row to
a CSV table (opens directly in Excel / Numbers / Google Sheets):

  date, time, inflow_cfs, outflow_cfs, net_cfs, rain_24h_in, lake_level_ft,
  net_acre_ft, cumulative_net_acre_ft, inflow_as_of, outflow_as_of,
  rain_as_of, status

Only the Python 3 standard library is used, so there is nothing to install.

Usage
  python3 lanier_logger.py                 # 24 readings, each at the top of the hour, then stop
  python3 lanier_logger.py --hours 48      # run for 48 hours instead
  python3 lanier_logger.py --now           # take a reading immediately, then continue hourly
  python3 lanier_logger.py --once          # one reading and exit (for cron / Task Scheduler)
  python3 lanier_logger.py --file my.csv   # choose the output file (default lake_lanier_log.csv)

Cumulative difference
  Each reading is treated as representing one hour of flow.
  1 cubic foot per second for 1 hour = 3,600 ft^3 = 0.08264 acre-feet.
  net_acre_ft = (inflow - outflow) * 0.08264, and cumulative_net_acre_ft is the
  running total across the whole log file (it continues across runs).
  A negative number means more water left the lake than came in.

Daily summary
  After every reading, lake_lanier_daily.csv is rebuilt with one row per date:
  readings taken, average inflow/outflow/net CFS, the day's TOTAL net difference
  in acre-feet and million gallons (sum of that day's hourly rows; negative =
  lake lost water), the site's 24h rainfall at the day's last reading, and the
  lake level change from the first to the last reading of the day.
  Only rows with status "ok" are counted.

Estimated vs. actual level change (daily table)
  est_change_from_flow_ft  = day's net acre-ft / lake surface area (AREA_TABLE)
  est_change_from_rain_ft  = 24h rainfall (inches) / 12  (rain landing on the lake)
  est_change_ft            = the two added together
  actual_change_ft         = today's last lake level - yesterday's last level
  gap_ft                   = actual - estimated (what the gauges don't see:
                             ungauged creeks/runoff minus water-supply
                             withdrawals and evaporation)
  unmeasured_flow_cfs      = that gap expressed as a steady flow in CFS
                             (positive = extra water coming in)
  correction_cfs           = average unmeasured_flow_cfs of all EARLIER
                             near-complete days (>= 20 readings)
  corrected_est_change_ft  = estimate with that correction applied - this is
                             the number to watch as it gets calibrated over weeks
"""

import argparse
import csv
import html
import os
import re
import sys
import time
import urllib.request
from datetime import datetime, timedelta
from html.parser import HTMLParser

URL = "https://lakelanierwater.com/"
DEFAULT_FILE = "lake_lanier_log.csv"
ACRE_FT_PER_CFS_HOUR = 3600 / 43560  # 0.082645...
USER_AGENT = "LakeLanierHourlyLogger/1.0 (personal use)"

COLUMNS = [
    "date", "time", "inflow_cfs", "outflow_cfs", "net_cfs", "rain_24h_in",
    "lake_level_ft", "net_acre_ft", "cumulative_net_acre_ft",
    "inflow_as_of", "outflow_as_of", "rain_as_of", "status",
]


# ---------------------------------------------------------------- fetching ---

class _TextExtractor(HTMLParser):
    """Turn HTML into visible text, one space between every element."""

    def __init__(self):
        super().__init__()
        self.parts, self._skip = [], 0

    def handle_starttag(self, tag, attrs):
        if tag in ("script", "style", "noscript", "template"):
            self._skip += 1
        self.parts.append(" ")

    def handle_endtag(self, tag):
        if tag in ("script", "style", "noscript", "template") and self._skip:
            self._skip -= 1
        self.parts.append(" ")

    def handle_data(self, data):
        if not self._skip:
            self.parts.append(data)

    def text(self):
        return re.sub(r"\s+", " ", "".join(self.parts))


def fetch_page(url=URL, timeout=30):
    req = urllib.request.Request(url, headers={
        "User-Agent": USER_AGENT,
        "Cache-Control": "no-cache",
    })
    with urllib.request.urlopen(req, timeout=timeout) as resp:
        return resp.read().decode("utf-8", errors="replace")


# ----------------------------------------------------------------- parsing ---

NUM = r"(\d[\d,]*(?:\.\d+)?)"
TIME = r"(\d{1,2}:\d{2}\s*[AP]M)"


def _num(s):
    return float(s.replace(",", "")) if s is not None else None


def _search(pattern, text):
    return re.search(pattern, text, re.IGNORECASE | re.DOTALL)


def parse_readings(page_html):
    """Return a dict of readings parsed from the page (values may be None)."""
    visible = _TextExtractor()
    visible.feed(page_html)
    texts = [visible.text()]
    # Fallback: data embedded in scripts (Next.js payload), with tags removed.
    raw = html.unescape(re.sub(r"<[^>]+>", " ", page_html))
    texts.append(re.sub(r"\s+", " ", raw.replace("\\n", " ").replace('\\"', '"')))

    out = {}
    for text in texts:
        if out.get("inflow_cfs") is None:
            m = _search(r"Inflow\b.{0,200}?" + NUM + r"\s*CFS(.{0,120}?as of\s*" + TIME + ")?", text)
            if m:
                out["inflow_cfs"], out["inflow_as_of"] = _num(m.group(1)), m.group(3)
        if out.get("outflow_cfs") is None:
            m = _search(r"Outflow\b.{0,200}?" + NUM + r"\s*CFS(.{0,120}?as of\s*" + TIME + ")?", text)
            if m:
                out["outflow_cfs"], out["outflow_as_of"] = _num(m.group(1)), m.group(3)
        if out.get("rain_24h_in") is None:
            m = _search(r"Rainfall\b.{0,200}?" + NUM + r"\s*in\b(.{0,120}?as of\s*" + TIME + ")?", text)
            if m:
                out["rain_24h_in"], out["rain_as_of"] = _num(m.group(1)), m.group(3)
        if out.get("lake_level_ft") is None:
            # Current elevation: first 4-digit number with decimals followed by "ft"
            m = _search(r"\b(1,?0\d\d\.\d{1,2})\s*ft\b", text)
            if m:
                out["lake_level_ft"] = _num(m.group(1))
    return out


# ----------------------------------------------------------------- logging ---

def last_cumulative(path):
    """Cumulative total from the last row of an existing log (0 if none)."""
    if not os.path.exists(path):
        return 0.0
    total = 0.0
    with open(path, newline="", encoding="utf-8") as f:
        for row in csv.DictReader(f):
            v = row.get("cumulative_net_acre_ft")
            if v not in (None, ""):
                total = float(v)
    return total


def append_row(path, row):
    new_file = not os.path.exists(path) or os.path.getsize(path) == 0
    with open(path, "a", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=COLUMNS)
        if new_file:
            w.writeheader()
        w.writerow(row)


DAILY_COLUMNS = [
    "date", "readings", "avg_inflow_cfs", "avg_outflow_cfs", "avg_net_cfs",
    "total_net_acre_ft", "total_net_million_gallons", "rain_24h_in_last_reading",
    "lake_level_first_ft", "lake_level_last_ft", "lake_level_change_ft",
    # --- level estimate (water balance) ---
    "surface_area_acres", "est_change_from_flow_ft", "est_change_from_rain_ft",
    "est_change_ft", "actual_change_ft", "gap_ft", "unmeasured_flow_cfs",
    "correction_cfs", "corrected_est_change_ft",
]
MGAL_PER_ACRE_FT = 0.325851
CFS_PER_ACRE_FT_PER_DAY = 43560 / 86400  # 0.504

# ---- Lake surface area by elevation -----------------------------------------
# The level estimate divides the day's net volume by the lake's surface area.
# Published figure: ~38,000 acres at full pool (1,071 ft). The area shrinks as
# the lake drops; 34,000 acres near 1,062 ft is a ROUGH ASSUMPTION, not an
# official number. For best accuracy, replace/extend these points with values
# from the USACE Lake Lanier elevation-area-capacity table. Values in between
# are interpolated; values outside the table use the nearest point.
AREA_TABLE = [  # (elevation_ft, surface_area_acres)
    (1062.0, 34000),   # assumption - replace with USACE value
    (1071.0, 38000),   # full pool, published
]
MIN_READINGS_FOR_CALIBRATION = 20  # only near-complete days feed the correction


def surface_area(level_ft):
    pts = sorted(AREA_TABLE)
    if level_ft is None:
        return pts[-1][1]
    if level_ft <= pts[0][0]:
        return pts[0][1]
    if level_ft >= pts[-1][0]:
        return pts[-1][1]
    for (e1, a1), (e2, a2) in zip(pts, pts[1:]):
        if e1 <= level_ft <= e2:
            return a1 + (a2 - a1) * (level_ft - e1) / (e2 - e1)


def daily_path(hourly_path):
    base, ext = os.path.splitext(hourly_path)
    base = base[:-4] if base.endswith("_log") else base
    return f"{base}_daily{ext or '.csv'}"


def write_daily_summary(hourly_path):
    """Rebuild the one-row-per-day summary from the hourly log."""
    days = {}
    with open(hourly_path, newline="", encoding="utf-8") as f:
        for r in csv.DictReader(f):
            if r.get("status") != "ok" or not r.get("net_acre_ft"):
                continue
            days.setdefault(r["date"], []).append(r)

    rows = []
    prev_last_level = None   # last level of the previous logged day
    gap_history = []         # unmeasured_flow_cfs of earlier near-complete days
    for date in sorted(days):
        rs = days[date]
        n = len(rs)
        avg = lambda k: sum(float(r[k]) for r in rs) / n
        total_af = sum(float(r["net_acre_ft"]) for r in rs)
        levels = [float(r["lake_level_ft"]) for r in rs if r.get("lake_level_ft")]

        # ---- water-balance level estimate ----
        est = {k: "" for k in DAILY_COLUMNS[11:]}
        if levels:
            area = surface_area(sum(levels) / len(levels))
            flow_ft = total_af / area
            rain = rs[-1].get("rain_24h_in")
            rain_ft = float(rain) / 12 if rain else 0.0  # rain on the lake adds its depth
            est_ft = flow_ft + rain_ft
            # Actual change: previous day's last level -> today's last level
            # (falls back to first->last reading on the first logged day).
            start = prev_last_level if prev_last_level is not None else levels[0]
            actual_ft = levels[-1] - start
            gap_ft = actual_ft - est_ft
            gap_cfs = gap_ft * area * CFS_PER_ACRE_FT_PER_DAY * (24 / n)
            correction = sum(gap_history) / len(gap_history) if gap_history else None
            corrected = (est_ft + correction * n / 24 / CFS_PER_ACRE_FT_PER_DAY / area
                         if correction is not None else None)
            est.update({
                "surface_area_acres": fmt(area, 0),
                "est_change_from_flow_ft": fmt(flow_ft, 3),
                "est_change_from_rain_ft": fmt(rain_ft, 3),
                "est_change_ft": fmt(est_ft, 3),
                "actual_change_ft": fmt(actual_ft, 3),
                "gap_ft": fmt(gap_ft, 3),
                "unmeasured_flow_cfs": fmt(gap_cfs, 0),
                "correction_cfs": fmt(correction, 0),
                "corrected_est_change_ft": fmt(corrected, 3),
            })
            if n >= MIN_READINGS_FOR_CALIBRATION and prev_last_level is not None:
                gap_history.append(gap_cfs)
            prev_last_level = levels[-1]

        rows.append({
            **est,
            "date": date,
            "readings": n,
            "avg_inflow_cfs": fmt(avg("inflow_cfs"), 0),
            "avg_outflow_cfs": fmt(avg("outflow_cfs"), 0),
            "avg_net_cfs": fmt(avg("net_cfs"), 0),
            "total_net_acre_ft": fmt(total_af, 1),
            "total_net_million_gallons": fmt(total_af * MGAL_PER_ACRE_FT, 1),
            "rain_24h_in_last_reading": rs[-1].get("rain_24h_in", ""),
            "lake_level_first_ft": fmt(levels[0], 2) if levels else "",
            "lake_level_last_ft": fmt(levels[-1], 2) if levels else "",
            "lake_level_change_ft": fmt(levels[-1] - levels[0], 2) if levels else "",
        })

    with open(daily_path(hourly_path), "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=DAILY_COLUMNS)
        w.writeheader()
        w.writerows(rows)


def fmt(v, digits):
    return "" if v is None else f"{v:.{digits}f}"


TIMEZONE = "America/New_York"  # log times in Lake Lanier's local time, even on a UTC server


def local_now():
    try:
        from zoneinfo import ZoneInfo
        return datetime.now(ZoneInfo(TIMEZONE)).replace(tzinfo=None)
    except Exception:  # e.g. Windows without the tzdata package: use the computer's clock
        return datetime.now()


def take_reading(path, retries=3, retry_wait=60):
    stamp = local_now()
    readings, error = {}, None
    for attempt in range(1, retries + 1):
        try:
            readings = parse_readings(fetch_page())
            if readings.get("inflow_cfs") is not None and readings.get("outflow_cfs") is not None:
                error = None
                break
            error = "could not find inflow/outflow on page"
        except Exception as e:  # network error, timeout, etc.
            error = f"{type(e).__name__}: {e}"
        if attempt < retries:
            time.sleep(retry_wait)

    inflow, outflow = readings.get("inflow_cfs"), readings.get("outflow_cfs")
    cumulative = last_cumulative(path)
    net_cfs = net_af = None
    if inflow is not None and outflow is not None:
        net_cfs = inflow - outflow
        net_af = net_cfs * ACRE_FT_PER_CFS_HOUR
        cumulative += net_af

    row = {
        "date": stamp.strftime("%Y-%m-%d"),
        "time": stamp.strftime("%H:%M"),
        "inflow_cfs": fmt(inflow, 0),
        "outflow_cfs": fmt(outflow, 0),
        "net_cfs": fmt(net_cfs, 0),
        "rain_24h_in": fmt(readings.get("rain_24h_in"), 2),
        "lake_level_ft": fmt(readings.get("lake_level_ft"), 2),
        "net_acre_ft": fmt(net_af, 1),
        "cumulative_net_acre_ft": fmt(cumulative, 1),
        "inflow_as_of": readings.get("inflow_as_of") or "",
        "outflow_as_of": readings.get("outflow_as_of") or "",
        "rain_as_of": readings.get("rain_as_of") or "",
        "status": "ok" if error is None else f"ERROR - {error}",
    }
    append_row(path, row)
    try:
        write_daily_summary(path)
    except PermissionError:
        print("  (daily summary not updated - close it in Excel and it will update next hour)")
    print(f"[{row['date']} {row['time']}] in={row['inflow_cfs']} cfs  "
          f"out={row['outflow_cfs']} cfs  net={row['net_cfs']} cfs  "
          f"rain24h={row['rain_24h_in']} in  cumulative={row['cumulative_net_acre_ft']} ac-ft  "
          f"({row['status']})", flush=True)
    return row


def sleep_until_next_hour():
    now = datetime.now()
    nxt = (now + timedelta(hours=1)).replace(minute=0, second=0, microsecond=0)
    print(f"Next reading at {nxt:%Y-%m-%d %H:%M}", flush=True)
    while True:
        remaining = (nxt - datetime.now()).total_seconds()
        if remaining <= 0:
            return
        time.sleep(min(remaining, 60))  # short naps keep it on time after laptop sleep


def main():
    ap = argparse.ArgumentParser(description="Log Lake Lanier inflow/outflow hourly.")
    ap.add_argument("--file", default=DEFAULT_FILE, help="CSV file to append to")
    ap.add_argument("--hours", type=int, default=24, help="number of hourly readings (default 24)")
    ap.add_argument("--now", action="store_true", help="also take a reading immediately")
    ap.add_argument("--once", action="store_true", help="take one reading and exit")
    args = ap.parse_args()

    path = os.path.abspath(args.file)
    print(f"Logging to {path}", flush=True)

    if args.once:
        take_reading(path)
        return

    if args.now:
        take_reading(path)
    try:
        for _ in range(args.hours):
            sleep_until_next_hour()
            take_reading(path)
    except KeyboardInterrupt:
        print("\nStopped.")
    print(f"Done. Table saved in {path}")


if __name__ == "__main__":
    sys.exit(main())
